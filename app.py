# -*- coding: utf-8 -*-
"""
App web de conciliación bancaria.
Flujo para el empleado:  subir el PDF del banco  ->  descargar el Excel conciliado.
No hay que elegir cliente ni subir ningún otro archivo: los criterios contables
(cómo imputar cada concepto a su cuenta) ya están dentro del programa.

Bancos soportados: Galicia, BBVA, Macro, Ciudad, Comafi, Supervielle.
Correr local:  streamlit run app.py     Requisitos: pip install -r requirements.txt
"""
import os, tempfile
import streamlit as st
import conversor_extractos as C
import importador_xubio as IX

st.set_page_config(page_title="Conciliación Bancaria", page_icon="🏦", layout="centered")

def acceso_ok() -> bool:
    if st.session_state.get("auth_ok"):
        return True
    st.title("🏦 Conciliación Bancaria")
    pw = st.text_input("Contraseña", type="password")
    if st.button("Entrar"):
        try:
            correcta = st.secrets["app_password"]
        except Exception:
            correcta = None
        if correcta is None:
            st.warning("Falta configurar 'app_password' en los secrets de la app.")
        elif pw == correcta:
            st.session_state["auth_ok"] = True; st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False

if not acceso_ok():
    st.stop()

st.title("🏦 Conciliación Bancaria")
st.caption("Subí el extracto del banco en **PDF** y descargá el Excel conciliado "
           "y el importador de asientos para Xubio. "
           "Bancos: Galicia, BBVA, Macro, Ciudad, Comafi, Supervielle.")

archivos = st.file_uploader("Extracto(s) en PDF", type=["pdf"], accept_multiple_files=True)

for arch in archivos or []:
    st.divider(); st.subheader(f"📄 {arch.name}")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".pdf") as tmp:
        tmp.write(arch.getbuffer()); ruta_pdf = tmp.name
    try:
        resultados = C.leer_extracto(ruta_pdf)
    except Exception as e:
        st.error(f"No pude procesar el PDF (¿banco soportado?). Detalle: {e}")
        os.unlink(ruta_pdf); continue

    st.write(f"**Banco:** {resultados[0].banco}  ·  **Cuentas:** {len(resultados)}")
    todo_ok = True
    for res in resultados:
        C.clasificar(res)                      # criterios del contador, ya incorporados
        ctrl = C.verificar_control(res); ok = ctrl["ok"]; todo_ok = todo_ok and ok
        tot = len(res.movimientos); con = sum(1 for m in res.movimientos if m.cuenta_sugerida)
        c1, c2, c3 = st.columns(3)
        c1.metric(f"Cta {res.cuenta}", f"{tot} movs")
        c2.metric("Control de saldo", "OK ✓" if ok else "REVISAR ✗", f"dif {ctrl['diferencia']:.2f}")
        c3.metric("Con cuenta contable", f"{con}/{tot}")
        if not ok:
            st.warning(f"⚠️ La cuenta {res.cuenta} no cuadra (dif {ctrl['diferencia']:.2f}). "
                       "Revisá antes de importar a Xubio.")
        if tot - con:
            st.caption(f"{tot - con} movimiento(s) sin cuenta asignada — marcados para revisión "
                       "manual (p. ej. pagos de AFIP que requieren el VEP, o conceptos nuevos).")

    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmpx:
        ruta_xlsx = tmpx.name
    C.exportar_excel(resultados, ruta_xlsx)
    with open(ruta_xlsx, "rb") as fh:
        data = fh.read()
    st.download_button("⬇️ Descargar Excel conciliado", data=data,
                       file_name=arch.name.rsplit('.', 1)[0] + " - conciliado.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    if todo_ok:
        st.success("Todas las cuentas cerraron el control de saldo en cero ✓")

    # ---- CAPA 3: importador de asientos para Xubio ----
    st.markdown("**📘 Asientos para Xubio**")
    with tempfile.NamedTemporaryFile(delete=False, suffix=".xlsx") as tmpimp:
        ruta_imp = tmpimp.name
    IX.generar_excel(resultados, ruta_imp)
    cobertura = IX.control_cobertura(resultados)
    for r in cobertura:
        if r["completo"]:
            st.caption(f"Cta {r['cuenta']}: importador COMPLETO ✓ — refleja todo el extracto.")
        else:
            st.warning(f"⚠️ Cta {r['cuenta']}: importador INCOMPLETO — faltan imputar "
                       f"${r['faltan_monto']:,.2f} en {r['faltan_movimientos']} movimiento(s). "
                       "Resolvé la hoja **A REVISAR** antes de subirlo a Xubio.")
    with open(ruta_imp, "rb") as fh:
        data_imp = fh.read()
    st.download_button("⬇️ Descargar importador de asientos (Xubio)", data=data_imp,
                       file_name=arch.name.rsplit('.', 1)[0] + " - importador asientos.xlsx",
                       mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
    st.caption("El importador es un borrador de arranque: mientras el control diga INCOMPLETO, "
               "primero resolvé lo que quede en la hoja A REVISAR antes de importarlo a Xubio.")

    os.unlink(ruta_pdf); os.unlink(ruta_xlsx); os.unlink(ruta_imp)
