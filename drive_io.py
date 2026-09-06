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

import io
import json
import os
import re
import unicodedata
from typing import Dict, List, Optional

import gspread
from google.oauth2.service_account import Credentials
from googleapiclient.discovery import build
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


# ---------------------------------------------------------------------------
# Buscar un archivo dentro de la carpeta por su nombre base
# ---------------------------------------------------------------------------
def buscar_archivo_por_nombre_base(carpeta_id: str, nombre_base: str) -> Optional[dict]:
    """Busca, dentro de la carpeta del cliente, el archivo cuyo nombre
    EMPIEZA con `nombre_base` (el resto es el sufijo del cliente, ej.
    '+ Venezuela'). Devuelve {'id', 'name', 'mimeType'} o None si no
    aparece (el llamador decide qué hacer — nunca se inventa un archivo)."""
    _, drive = _clientes()
    resp = drive.files().list(
        q=f"'{carpeta_id}' in parents and trashed = false",
        fields="files(id, name, mimeType)",
        supportsAllDrives=True,
        includeItemsFromAllDrives=True,
    ).execute()
    objetivo = _normalizar(nombre_base)
    candidatos = [
        f for f in resp.get("files", [])
        if _normalizar(f["name"]).startswith(objetivo)
    ]
    if not candidatos:
        return None
    if len(candidatos) > 1:
        print(f"[drive_io] Aviso: hay {len(candidatos)} archivos que empiezan "
              f"con '{nombre_base}' en la carpeta; uso '{candidatos[0]['name']}'.")
    return candidatos[0]


# ---------------------------------------------------------------------------
# Leer una solapa de un Google Sheet como lista de dicts
# ---------------------------------------------------------------------------
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


def listar_solapas(archivo: dict) -> List[str]:
    """Nombres de todas las solapas de un Sheet (para debug / validar)."""
    gc, _ = _clientes()
    sh = gc.open_by_key(archivo["id"])
    return [ws.title for ws in sh.worksheets()]


# ---------------------------------------------------------------------------
# Bajar un PDF en bytes (para pasárselo directo a pdfplumber)
# ---------------------------------------------------------------------------
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


# ---------------------------------------------------------------------------
# Subir el Excel de salida a la carpeta del cliente
# ---------------------------------------------------------------------------
def escribir_excel_en_carpeta(carpeta_id: str, nombre: str, contenido: bytes) -> str:
    """Sube un .xlsx a la carpeta del cliente. Si ya existe un archivo con
    ese nombre en esa carpeta, lo REEMPLAZA (mismo ID de Drive, no duplica).
    Devuelve el ID del archivo en Drive."""
    _, drive = _clientes()
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
