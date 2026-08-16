# Poner la app online (Streamlit Community Cloud, gratis)

Archivos (una carpeta / repositorio):
    app.py                  <- la pantalla (subir PDF -> descargar conciliación)
    conversor_extractos.py  <- el motor (lectura + control + clasificación con los criterios del contador)
    requirements.txt        <- librerías

## Pasos
1. Cuenta gratis en https://streamlit.io/cloud (se entra con GitHub).
2. Subir los 3 archivos a un repo de GitHub (puede ser PRIVADO).
3. New app -> elegir repo -> archivo principal: app.py -> Deploy.
4. Contraseña: Settings -> Secrets -> pegar:  app_password = "LA-CLAVE"
5. (Recomendado) App privada + lista blanca de correos de los empleados (Settings -> Sharing).
6. Compartir el link (tu-app.streamlit.app) con el equipo.

## Uso diario del empleado
Entrar -> contraseña -> subir el PDF del banco -> descargar el Excel conciliado. Nada más.

## Qué resuelve solo (sin cargar nada)
Convierte el PDF, controla el saldo (tiene que dar 0) y asigna la cuenta contable de
cada movimiento según los criterios del contador (ya incorporados): comisiones, IVA,
impuestos, ley 25.413, percepciones, SIRCREB, sueldos, cobranzas (Deudores por Venta),
proveedores (Proveedores). En pruebas reales: 98-100% en Galicia/Macro, ~90% en BBVA.

## Qué queda "a revisar" (marcado, no adivinado)
- Pagos de AFIP: para saber el impuesto exacto se necesita el VEP del cliente.
- Conceptos nuevos que el contador no haya usado antes.
Esos movimientos quedan sin cuenta, señalados, para revisión humana rápida.

## Seguridad
Con la nube de Streamlit los PDF se procesan en sus servidores (EE.UU.), cifrados.
Alternativa para datos sensibles: correr el MISMO código en un servidor propio con
'streamlit run app.py' (los datos no salen de la empresa).

## Nota de mantenimiento
Los criterios del contador viven en conversor_extractos.py (MAPEO_CONTADOR + REGLAS_CONCEPTO).
Si el contador cambia una cuenta o aparece un concepto nuevo, se ajusta ahí (una línea).
