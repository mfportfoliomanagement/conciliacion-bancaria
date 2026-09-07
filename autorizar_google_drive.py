# -*- coding: utf-8 -*-
"""
Autorización ÚNICA para que el bot pueda ESCRIBIR archivos en Google Drive
actuando como un usuario real.

Por qué hace falta: la cuenta de servicio (la de credenciales.json / los
secrets de Streamlit) puede LEER todo perfecto, pero no tiene cuota de
almacenamiento propia -- Google no la deja crear archivos nuevos dentro de
una carpeta de Drive normal ("Mi unidad"), ni siquiera con permiso de
Editor. Por eso, para subir el papel de trabajo y el importador de Xubio a
la carpeta de cada cliente, el bot necesita actuar como un usuario real
(con su propio espacio de Drive).

Este script se corre UNA SOLA VEZ, a mano, en la PC de esa persona (no es
parte de la app). Abre un link de Google para que hagas el login y le des
permiso al bot -- Claude/el bot NUNCA ve tu contraseña. Al final imprime
3 valores que hay que guardar en los Secrets de Streamlit bajo la clave
'google_oauth_usuario' (nunca en un archivo del repo).

Requisito previo (una sola vez, lo hace quien tenga acceso al proyecto de
Google Cloud): crear un "ID de cliente de OAuth" de tipo "Aplicación de
escritorio" en
    https://console.cloud.google.com/apis/credentials?project=<TU_PROYECTO>
y descargar el JSON. Ver COMO_DESPLEGAR.md para el paso a paso completo.

Uso:
    python autorizar_google_drive.py ruta/al/client_secret_....json
"""
import sys

from google_auth_oauthlib.flow import InstalledAppFlow

SCOPES = ["https://www.googleapis.com/auth/drive"]


def main():
    if len(sys.argv) < 2:
        print("Uso: python autorizar_google_drive.py ruta/al/client_secret_....json")
        sys.exit(1)
    archivo_client_secret = sys.argv[1]

    flow = InstalledAppFlow.from_client_secrets_file(archivo_client_secret, SCOPES)

    # open_browser=False: en algunos entornos el navegador no se abre solo.
    # Imprime el link para pegarlo a mano si hace falta.
    creds = flow.run_local_server(port=0, prompt="consent", open_browser=False)

    print("\n" + "=" * 70)
    print("LISTO. Guardá estos 3 valores en los Secrets de Streamlit, bajo")
    print("la clave 'google_oauth_usuario' (nunca en un archivo del repo):")
    print("=" * 70)
    print(f"client_id     = {creds.client_id}")
    print(f"client_secret = {creds.client_secret}")
    print(f"refresh_token = {creds.refresh_token}")
    print("=" * 70)
    print("\nEjemplo de bloque para pegar en .streamlit/secrets.toml (local) o")
    print("en los Secrets de la app en Streamlit Cloud:")
    print("""
[google_oauth_usuario]
client_id = "..."
client_secret = "..."
refresh_token = "..."
""")


if __name__ == "__main__":
    main()
