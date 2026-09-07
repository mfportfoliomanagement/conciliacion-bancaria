# -*- coding: utf-8 -*-
"""
============================================================================
 DRIVE_IO — Capa de acceso a Google Drive / Google Sheets
============================================================================
Es el ÚNICO módulo del proyecto que habla con Google. Ningún otro archivo
(`conversor_extractos.py`, `importador_xubio.py`, `app.py`) debe importar
`gspread` ni `googleapiclient` directamente: todos pasan por acá.

Responsabilidades:
  - Autenticarse con la cuenta de servicio (Service Account).
  - Entrar a la carpeta de Drive de un cliente.
  - Buscar, dentro de esa carpeta, un archivo por su "nombre base" (cada
    cliente le agrega su propio sufijo al nombre, ej. "... + Venezuela").
  - Leer una solapa de un Google Sheet como lista de diccionarios (una
    fila del Sheet = un dict, con el encabezado de cada columna como clave).
  - Bajar los bytes de un PDF guardado en Drive.
  - Subir un Excel (bytes) a la carpeta del cliente.

Este módulo NO clasifica movimientos ni conoce nombres de clientes,
proveedores, cuentas, etc. — es pura "cañería" de Google.

Credenciales
------------
La clave de la cuenta de servicio NUNCA se lee de un archivo del repo.
Sale de una de estas dos fuentes, en este orden:
  1. `st.secrets["gcp_service_account"]`  (modo normal: app corriendo en
     Streamlit Cloud, con la clave cargada en los Secrets de la app).
  2. Variable de entorno `GOOGLE_SERVICE_ACCOUNT_JSON`  (modo pruebas
     locales desde consola: se pega ahí el JSON completo de la clave,
     nunca en un archivo).
============================================================================
"""
from __future__ import annotations

import difflib
import functools
import io
import json
import os
import re
import time
import unicodedata
from typing import Dict, List, Optional

import gspread
from google.oauth2.service_account import Credentials
from google.oauth2.credentials import Credentials as CredencialesUsuario
from googleapiclient.discovery import build
from googleapiclient.errors import HttpError
from googleapiclient.http import MediaIoBaseDownload, MediaIoBaseUpload

_SCOPES = [
    "https://www.googleapis.com/auth/drive",
    "https://www.googleapis.com/auth/spreadsheets",
]

_MIME_FOLDER = "application/vnd.google-apps.folder"
_MIME_XLSX = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"

# ---------------------------------------------------------------------------
# Autenticación (se arma una sola vez y se reutiliza)
# ---------------------------------------------------------------------------
_creds = None
_gc = None
_drive = None


def _cargar_credenciales() -> Credentials:
    """Arma las credenciales de la Service Account sin leer ningún archivo
    del repo. Ver docstring del módulo para el orden de búsqueda."""
    info = None
    try:
        import streamlit as st  # import local: no todo el que use drive_io corre bajo Streamlit
        if "gcp_service_account" in st.secrets:
            info = dict(st.secrets["gcp_service_account"])
    except Exception:
        pass  # no hay Streamlit corriendo, o no está esa clave en sus secrets

    if info is None:
        raw = os.environ.get("GOOGLE_SERVICE_ACCOUNT_JSON")
        if raw:
            info = json.loads(raw)

    if info is None:
        raise RuntimeError(
            "No encontré credenciales de Google. Configurá "
            "st.secrets['gcp_service_account'] (Streamlit) o la variable de "
            "entorno GOOGLE_SERVICE_ACCOUNT_JSON con el JSON de la cuenta de "
            "servicio. La clave NUNCA va en un archivo del repo."
        )
    return Credentials.from_service_account_info(info, scopes=_SCOPES)


def _clientes():
    """Devuelve (gspread_client, drive_service), armándolos una sola vez."""
    global _creds, _gc, _drive
    if _creds is None:
        _creds = _cargar_credenciales()
        _gc = gspread.authorize(_creds)
        _drive = build("drive", "v3", credentials=_creds, cache_discovery=False)
    return _gc, _drive


# ---------------------------------------------------------------------------
# Reintentos ante errores TRANSITORIOS de Google (límite de lecturas por
# minuto, servidor ocupado) -- nunca ante errores reales (permiso denegado,
# archivo inexistente). Sin esto, un pico de uso (varios clientes seguidos
# en el modo lote, o probar y usar la app al mismo tiempo) puede tirar la
# app entera por un error que se hubiera resuelto solo unos segundos después.
# ---------------------------------------------------------------------------
def _codigo_http(e) -> Optional[int]:
    if isinstance(e, gspread.exceptions.APIError):
        try:
            return e.response.status_code
        except Exception:
            return None
    if isinstance(e, HttpError):
        return e.resp.status if getattr(e, "resp", None) else None
    return None


def _es_error_transitorio(e) -> bool:
    """True si vale la pena reintentar (límite de lecturas, servidor
    ocupado). False si es un error real -- ahí no hay que insistir."""
    return _codigo_http(e) in (429, 500, 502, 503, 504)


def _reintentable(func):
    """Decorador: si `func` falla con un error transitorio de Google,
    reintenta unas pocas veces con espera creciente antes de rendirse."""
    @functools.wraps(func)
    def envoltura(*args, intentos=4, espera_inicial=2, **kwargs):
        espera = espera_inicial
        for intento in range(1, intentos + 1):
            try:
                return func(*args, **kwargs)
            except Exception as e:
                if intento == intentos or not _es_error_transitorio(e):
                    raise
                print(f"[drive_io] Google devolvió un error transitorio "
                      f"({_codigo_http(e)}) en {func.__name__} -- "
                      f"reintento {intento}/{intentos} en {espera}s...")
                time.sleep(espera)
                espera *= 2
    return envoltura


# ---------------------------------------------------------------------------
# Credenciales de un USUARIO REAL (no la cuenta de servicio), solo para
# ESCRIBIR archivos en Drive.
#
# Las cuentas de servicio no tienen cuota de almacenamiento propia: Google
# rechaza que creen archivos nuevos (.xlsx, PDFs, etc.) dentro de una
# carpeta de Drive normal ("Mi unidad"), aunque tengan permiso de Editor
# sobre esa carpeta. Por eso, para subir el resultado a la carpeta del
# cliente, el bot actúa como un usuario real (autorizado una única vez —
# ver el script autorizar_google_drive.py), usando el espacio de ESE
# usuario. La lectura (Sheets, PDFs, listar archivos) sigue igual, con la
# cuenta de servicio.
#
# Se arma igual que las credenciales de servicio: st.secrets primero,
# variable de entorno después. Es OPCIONAL -- si no está configurado,
# escribir_excel_en_carpeta() tira un error claro en vez de fallar
# confuso contra Google.
# ---------------------------------------------------------------------------
_creds_usuario = None
_drive_usuario = None


def _cargar_credenciales_usuario():
    """Arma las credenciales OAuth de un usuario real a partir de un
    refresh_token ya autorizado. Devuelve None si no está configurado
    (no es un error: el resto del módulo sigue funcionando para lectura)."""
    info = None
    try:
        import streamlit as st
        if "google_oauth_usuario" in st.secrets:
            info = dict(st.secrets["google_oauth_usuario"])
    except Exception:
        pass

    if info is None:
        raw = os.environ.get("GOOGLE_OAUTH_USUARIO_JSON")
        if raw:
            info = json.loads(raw)

    if info is None:
        return None

    return CredencialesUsuario(
        token=None,
        refresh_token=info["refresh_token"],
        client_id=info["client_id"],
        client_secret=info["client_secret"],
        token_uri="https://oauth2.googleapis.com/token",
        scopes=["https://www.googleapis.com/auth/drive"],
    )


def _cliente_drive_usuario():
    """Servicio de Drive autenticado como el usuario real, o None si no
    está configurado (ver _cargar_credenciales_usuario)."""
    global _creds_usuario, _drive_usuario
    if _drive_usuario is None:
        _creds_usuario = _cargar_credenciales_usuario()
        if _creds_usuario is None:
            return None
        _drive_usuario = build("drive", "v3", credentials=_creds_usuario, cache_discovery=False)
    return _drive_usuario


# ---------------------------------------------------------------------------
# Comparación tolerante de NOMBRES DE ARCHIVO (mayúsculas/acentos/espacios).
# Ojo: esto es solo para encontrar archivos por nombre en Drive. La
# comparación tolerante de CONCEPTOS/CUENTAS de negocio vive en
# conversor_extractos.py (Fase 2), no acá.
# ---------------------------------------------------------------------------
def _normalizar(s: str) -> str:
    s = unicodedata.normalize("NFKD", s or "").encode("ascii", "ignore").decode("ascii")
    return re.sub(r"\s+", " ", s).strip().lower()


# ---------------------------------------------------------------------------
# Carpeta del cliente
# ---------------------------------------------------------------------------
def _extraer_id_de_link(link_o_id: str) -> str:
    """Acepta tanto una URL de Drive como el ID pelado, y devuelve el ID."""
    m = re.search(r"/folders/([a-zA-Z0-9_-]+)", link_o_id)
    if m:
        return m.group(1)
    m = re.search(r"[?&]id=([a-zA-Z0-9_-]+)", link_o_id)
    if m:
        return m.group(1)
    return link_o_id.strip()


@_reintentable
def abrir_carpeta_cliente(link_o_id: str) -> str:
    """Entra a la carpeta del cliente y devuelve su ID de Drive, validando
    que exista y que la cuenta de servicio tenga acceso."""
    _, drive = _clientes()
    folder_id = _extraer_id_de_link(link_o_id)
    meta = drive.files().get(
        fileId=folder_id, fields="id, name, mimeType", supportsAllDrives=True
    ).execute()
    if meta.get("mimeType") != _MIME_FOLDER:
        raise ValueError(f"El link/ID '{link_o_id}' no apunta a una carpeta de Drive.")
    return folder_id


@_reintentable
def abrir_archivo_por_id(archivo_id: str) -> dict:
    """Abre un archivo de Drive directamente por su ID (para archivos que
    NO viven dentro de la carpeta de un cliente, ej. el Maestro de
    Clientes). Devuelve {'id', 'name', 'mimeType'}."""
    _, drive = _clientes()
    return drive.files().get(
        fileId=archivo_id, fields="id, name, mimeType", supportsAllDrives=True
    ).execute()


@_reintentable
def buscar_subcarpeta(carpeta_id: str, nombre: str) -> Optional[str]:
    """Busca una subcarpeta por nombre (tolerante a mayúsculas/acentos/
    espacios) dentro de `carpeta_id`. Devuelve su ID de Drive, o None si
    no existe (ej. todavía no se cargó nada para ese año/mes)."""
    _, drive = _clientes()
    resp = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false and mimeType = '{_MIME_FOLDER}'",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    objetivo = _normalizar(nombre)
    for f in resp.get("files", []):
        if _normalizar(f["name"]) == objetivo:
            return f["id"]
    return None


# ---------------------------------------------------------------------------
# Buscar un archivo dentro de la carpeta por su nombre base
# ---------------------------------------------------------------------------
@_reintentable
def buscar_archivo_por_nombre_base(carpeta_id: str, nombre_base: str, umbral: float = 0.85) -> Optional[dict]:
    """Busca, dentro de la carpeta del cliente, el archivo cuyo nombre
    corresponde a `nombre_base` (el resto es el sufijo del cliente, ej.
    '+ Venezuela'). Devuelve {'id', 'name', 'mimeType'} o None si no
    aparece (el llamador decide qué hacer — nunca se inventa un archivo).

    Primero prueba "empieza con" (caso normal). Si nada matchea así, tolera
    pequeños errores de tipeo en el nombre del archivo (ej. una letra de
    menos) comparando por similitud de texto — los archivos los nombran
    personas y a veces se equivocan al escribir el nombre base."""
    _, drive = _clientes()
    resp = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false",
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    archivos = resp.get("files", [])
    objetivo = _normalizar(nombre_base)

    candidatos = [f for f in archivos if _normalizar(f["name"]).startswith(objetivo)]

    if not candidatos:
        # Nadie matcheo exacto -> probamos tolerando errores de tipeo chicos
        # en el nombre del archivo (comparamos el objetivo contra el
        # comienzo del nombre real, con la misma cantidad de caracteres).
        puntajes = []
        for f in archivos:
            nombre_norm = _normalizar(f["name"])
            recorte = nombre_norm[: len(objetivo) + 5]
            ratio = difflib.SequenceMatcher(None, objetivo, recorte).ratio()
            if ratio >= umbral:
                puntajes.append((ratio, f))
        if puntajes:
            puntajes.sort(key=lambda t: -t[0])
            candidatos = [f for _, f in puntajes]

    if not candidatos:
        return None
    if len(candidatos) > 1:
        print(f"[drive_io] Aviso: hay {len(candidatos)} archivos que coinciden "
              f"con '{nombre_base}' en la carpeta; uso '{candidatos[0]['name']}'.")
    return candidatos[0]


# ---------------------------------------------------------------------------
# Leer una solapa de un Google Sheet como lista de dicts
# ---------------------------------------------------------------------------
@_reintentable
def leer_sheet(archivo: dict, solapa: str) -> List[Dict[str, str]]:
    """Lee la solapa `solapa` del Google Sheet `archivo` (el dict que
    devuelve buscar_archivo_por_nombre_base). Devuelve una lista de dicts:
    una entrada por fila, con el encabezado de cada columna como clave.
    Si la solapa no existe, devuelve [] (el llamador decide si ese
    archivo/solapa es opcional para ese cliente)."""
    gc, _ = _clientes()
    sh = gc.open_by_key(archivo["id"])
    try:
        ws = sh.worksheet(solapa)
    except gspread.exceptions.WorksheetNotFound:
        return []
    return ws.get_all_records()  # usa la fila 1 como encabezado


@_reintentable
def leer_sheet_crudo(archivo: dict, solapa: str) -> List[List[str]]:
    """Como leer_sheet, pero devuelve las filas TAL CUAL (lista de listas),
    sin usar la fila 1 como encabezado. Sirve para solapas donde el
    encabezado no es único (ej. columnas repetidas, como BANCOS del
    Maestro de Clientes) y por eso leer_sheet() no se puede usar."""
    gc, _ = _clientes()
    sh = gc.open_by_key(archivo["id"])
    try:
        ws = sh.worksheet(solapa)
    except gspread.exceptions.WorksheetNotFound:
        return []
    return ws.get_all_values()


@_reintentable
def listar_solapas(archivo: dict) -> List[str]:
    """Nombres de todas las solapas de un Sheet (para debug / validar)."""
    gc, _ = _clientes()
    sh = gc.open_by_key(archivo["id"])
    return [ws.title for ws in sh.worksheets()]


# ---------------------------------------------------------------------------
# Bajar un PDF en bytes (para pasárselo directo a pdfplumber)
# ---------------------------------------------------------------------------
@_reintentable
def descargar_pdf(archivo: dict) -> bytes:
    """Baja los BYTES crudos del PDF (nunca base64, nunca texto). Se le
    puede pasar directo a pdfplumber:
        pdfplumber.open(io.BytesIO(descargar_pdf(archivo)))"""
    _, drive = _clientes()
    buf = io.BytesIO()
    req = drive.files().get_media(fileId=archivo["id"], supportsAllDrives=True)
    downloader = MediaIoBaseDownload(buf, req)
    done = False
    while not done:
        _, done = downloader.next_chunk()
    return buf.getvalue()


@_reintentable
def listar_pdfs(carpeta_id: str) -> List[dict]:
    """Lista los PDF sueltos dentro de la carpeta del cliente (útil para el
    modo 'lote' de la Fase 4)."""
    _, drive = _clientes()
    resp = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false and mimeType = 'application/pdf'",
        fields="files(id, name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    return resp.get("files", [])


@_reintentable
def existe_archivo(carpeta_id: str, nombre: str) -> bool:
    """True si ya hay un archivo con ese nombre (tolerante a mayúsculas/
    acentos/espacios) en la carpeta. Se usa para no reprocesar un cliente
    cuya salida de un período ya se generó (modo lote retomable)."""
    _, drive = _clientes()
    resp = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false",
        fields="files(name)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    objetivo = _normalizar(nombre)
    return any(_normalizar(f["name"]) == objetivo for f in resp.get("files", []))


# ---------------------------------------------------------------------------
# Subir el Excel de salida a la carpeta del cliente
# ---------------------------------------------------------------------------
@_reintentable
def escribir_excel_en_carpeta(carpeta_id: str, nombre: str, contenido: bytes) -> str:
    """Sube un .xlsx a la carpeta del cliente, actuando como un USUARIO
    REAL (no la cuenta de servicio) -- ver el bloque de credenciales de
    usuario más arriba en este archivo; las cuentas de servicio no tienen
    cuota de almacenamiento propia y Google rechaza que creen archivos
    nuevos en una carpeta de Drive normal. Si ya existe un archivo con ese
    nombre en esa carpeta, lo REEMPLAZA (mismo ID de Drive, no duplica).
    Devuelve el ID del archivo en Drive."""
    drive = _cliente_drive_usuario()
    if drive is None:
        raise RuntimeError(
            "Para subir archivos a Drive hace falta la autorización de un "
            "usuario real (la cuenta de servicio no tiene cuota de "
            "almacenamiento propia). Corré autorizar_google_drive.py una vez "
            "y configurá st.secrets['google_oauth_usuario'] (o la variable de "
            "entorno GOOGLE_OAUTH_USUARIO_JSON) con client_id/client_secret/"
            "refresh_token."
        )
    if not nombre.lower().endswith(".xlsx"):
        nombre += ".xlsx"
    media = MediaIoBaseUpload(io.BytesIO(contenido), mimetype=_MIME_XLSX, resumable=True)

    existentes = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false and name = '{nombre}'",
        fields="files(id)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute().get("files", [])

    if existentes:
        archivo_id = existentes[0]["id"]
        drive.files().update(fileId=archivo_id, media_body=media, supportsAllDrives=True).execute()
        return archivo_id

    meta = {"name": nombre, "parents": [carpeta_id]}
    creado = drive.files().create(
        body=meta, media_body=media, fields="id", supportsAllDrives=True
    ).execute()
    return creado["id"]
