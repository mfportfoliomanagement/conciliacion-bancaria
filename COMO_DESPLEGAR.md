# Cómo desplegar — Conciliación Bancaria (modelo con config externa)

Este documento reemplaza la versión vieja (la de "subir PDF → bajar Excel").
El sistema ahora lee TODO el criterio de clasificación de los archivos de
Google Drive/Sheets de cada cliente — el código no sabe nada de ningún
cliente en particular. Ver `PLAN_CLAUDE_CODE_conciliacion.md` para el
diseño completo; esto es solo el paso a paso para dejarlo andando.

## Arquitectura, en una frase

La app lee todo (Drive + Sheets) con una **cuenta de servicio** de Google,
y escribe los archivos de salida actuando como un **usuario real** (porque
las cuentas de servicio no tienen espacio propio en Drive y Google no las
deja crear archivos nuevos en una carpeta normal).

## Archivos del repo

| Archivo | Capa | Qué hace |
|---|---|---|
| `drive_io.py` | 1 | Único módulo que habla con Google Drive/Sheets. |
| `conversor_extractos.py` | 2 | Lee los PDF de cada banco, controla el saldo, clasifica leyendo la config del cliente. |
| `importador_xubio.py` | 3 | Arma los asientos agrupados por circuito, genera el importador de Xubio + el papel de trabajo. |
| `orquestador.py` | 4 | Procesa un cliente para un período: junta las 3 capas de arriba + sube el resultado a Drive. |
| `app.py` | UI | La pantalla de Streamlit (pestañas Lote / Individual). |
| `autorizar_google_drive.py` | Setup | Script de autorización única (no es parte de la app — se corre a mano, una sola vez). |

---

## 1. Google Cloud: cuenta de servicio (para LEER)

1. Crear (o reusar) un proyecto en https://console.cloud.google.com/.
2. Habilitar dos APIs en ese proyecto:
   - **Google Drive API**
   - **Google Sheets API**
3. Crear una **cuenta de servicio** (IAM y administración → Cuentas de
   servicio → Crear). Generarle una clave nueva tipo **JSON** y descargarla.
   **Esa clave nunca se sube al repo** — el `.gitignore` ya bloquea
   cualquier `.json` de credenciales por las dudas.
4. Copiar el email de la cuenta de servicio (termina en
   `...iam.gserviceaccount.com`).

## 2. Google Cloud: autorización de un usuario real (para ESCRIBIR)

1. En el mismo proyecto: **APIs y servicios → Pantalla de consentimiento
   OAuth**. Tipo "Externo", completar lo básico, agregarse a uno mismo
   como "usuario de prueba" (no hace falta publicar la app).
2. **APIs y servicios → Credenciales → + Crear credenciales → ID de
   cliente de OAuth**. Tipo de aplicación: **Aplicación de escritorio**.
   Descargar el JSON resultante (`client_secret_....json`) — tampoco se
   sube al repo.
3. Correr, una sola vez, en la PC de la persona que va a "prestarle" su
   espacio de Drive al bot (normalmente quien administra las carpetas de
   los clientes):
   ```bash
   pip install -r requirements.txt
   python autorizar_google_drive.py ruta/al/client_secret_....json
   ```
   Se abre un link de Google (o se imprime en la consola) — hay que
   loguearse y darle "Permitir". Al final el script imprime 3 valores:
   `client_id`, `client_secret`, `refresh_token`. **Guardarlos** (van al
   punto 5).

   Importante: esa persona tiene que **ser dueña o tener acceso de
   Editor** en todas las carpetas de clientes donde se vaya a escribir —
   la escritura sale de SU cuenta, no de la cuenta de servicio.

## 3. Compartir las carpetas con el bot

- Cada carpeta de cliente en Drive: **Compartir → pegar el email de la
  cuenta de servicio → rol Editor**.
- El archivo `MAESTRO_CLIENTES` (el Google Sheet con la lista de
  clientes): también compartirlo con la cuenta de servicio, rol Editor
  (o al menos Lector, si nunca se va a escribir ahí desde el código).

## 4. Cargar el Maestro de Clientes

Un Google Sheet con 2 solapas (ver contrato completo en el plan):

- **Hoja 1**: columnas `RAZON SOCIAL`, `ESTADO`, `LINK CARPETA`, `CUIT` —
  una fila por cliente.
- **BANCOS**: columna A = `RAZON SOCIAL` (idéntico texto que en Hoja 1),
  columnas siguientes = un código `BANCO_[banco]_[tipo]_[moneda]` por
  cada moneda que maneje ese banco (ej. `BANCO_GALICIA_CTA CTE_$`). Es
  solo informativo — no determina cuántas cuentas físicas tiene el
  cliente (eso sale de los PDFs + el CBU del Plan de Cuentas).

Copiar el ID del Sheet (la parte del link entre `/d/` y `/edit`).

## 5. Cada cliente necesita, en su carpeta, sus 4 archivos de config

`CONCEPTO GASTO Y ASIGNACION CUENTA`, `Listado y Ca Cte de Proveedores`,
`Vep Pagados`, `PLAN DE CUENTAS` — ver el contrato de columnas en
`PLAN_CLAUDE_CODE_conciliacion.md`. Sin estos 4 archivos, ese cliente no
se puede procesar (queda anotado en el reporte, no rompe el lote).

Además, en el `PLAN DE CUENTAS` de cada cliente hay que cargar el **CBU**
de cada cuenta bancaria — sin eso, el sistema no puede identificar a qué
banco pertenece un extracto y esos movimientos quedan "a revisar".

## 6. Streamlit Community Cloud

1. Cuenta gratis en https://streamlit.io/cloud (con GitHub).
2. Repo de GitHub (puede ser privado) con todo el código.
3. New app → elegir el repo → archivo principal: `app.py` → Deploy.
4. **Settings → Secrets**, pegar (con los valores reales de cada uno):

   ```toml
   app_password = "LA-CLAVE-QUE-ELIJAS"
   maestro_clientes_id = "el-id-del-sheet-MAESTRO_CLIENTES"

   [gcp_service_account]
   type = "service_account"
   project_id = "..."
   private_key_id = "..."
   private_key = "-----BEGIN PRIVATE KEY-----\n...\n-----END PRIVATE KEY-----\n"
   client_email = "...@....iam.gserviceaccount.com"
   client_id = "..."
   auth_uri = "https://accounts.google.com/o/oauth2/auth"
   token_uri = "https://oauth2.googleapis.com/token"
   auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
   client_x509_cert_url = "..."
   universe_domain = "googleapis.com"

   [google_oauth_usuario]
   client_id = "el que imprimió autorizar_google_drive.py"
   client_secret = "el que imprimió autorizar_google_drive.py"
   refresh_token = "el que imprimió autorizar_google_drive.py"
   ```

   (El bloque `[gcp_service_account]` es, campo por campo, el mismo JSON
   que se descargó en el paso 1 — cada clave del JSON es una línea acá.)

5. (Recomendado) App privada + lista blanca de emails del equipo
   (Settings → Sharing).
6. Compartir el link (`tu-app.streamlit.app`) con el equipo.

## Uso diario

- **Procesar en lote**: elegir mes/año → recorre TODOS los clientes del
  Maestro. Si un cliente ya está procesado para ese período, lo saltea.
  Si le falta algo (PDF, config, CBU), queda anotado y sigue con el resto.
- **Cliente individual**: elegir cliente (desplegable del Maestro) +
  período → procesa solo ese.

Por cada cliente/período que sí se procesa, quedan 2 archivos en su
carpeta de Drive (en la subcarpeta `AÑO/MES`):
- `IMPORTADOR ASIENTOS - ...xlsx` — una sola hoja, lista para subir a Xubio.
- `Papel de trabajo - ...xlsx` — el detalle, el control de saldo, la hoja
  A REVISAR (pendientes) y las transferencias entre cuentas propias.

## Correr en local (para probar)

```bash
pip install -r requirements.txt
```

Crear `.streamlit/secrets.toml` en la raíz del repo con el mismo
contenido del paso 6 (ese archivo está bloqueado por `.gitignore`, nunca
se sube). Después:

```bash
streamlit run app.py
```

## Mantenimiento — dónde vive cada cosa

- **Reglas de negocio de un cliente puntual** (qué cuenta le corresponde
  a cada concepto, quién es cada proveedor/empleado, etc.): SIEMPRE en
  los 4 archivos de ESE cliente en Drive. El código nunca se toca por
  esto.
- **El circuito de cada cuenta** (a cuál de los 9 grupos de asiento
  pertenece — Gastos Bancarios, Sueldos, Proveedores, etc.): es la única
  config que queda en el código, en `CIRCUITO_POR_CUENTA`
  (`importador_xubio.py`) — es estructural, no cambia por cliente.
- **Lectores de PDF por banco**, detección de banco, control de saldo:
  `conversor_extractos.py` — no se toca al agregar clientes.

## Seguridad

- La clave de la cuenta de servicio y las credenciales OAuth viven SOLO
  en los Secrets de Streamlit (o en variables de entorno para pruebas
  locales) — nunca en un archivo del repo.
- Los PDF y las planillas se procesan en los servidores de Streamlit
  Cloud (EE.UU.). Para datos más sensibles, correr el mismo código en un
  servidor propio con `streamlit run app.py` (los datos no salen de la
  empresa).
