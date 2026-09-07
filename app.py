# -*- coding: utf-8 -*-
"""
App web de conciliación bancaria — Fase 4.

Dos pestañas:
  - Lote: recorre TODOS los clientes del Maestro para un período.
  - Individual: procesa un cliente puntual.

Esta app no tiene ninguna regla de negocio de ningún cliente: todo el
criterio de clasificación sale de los 4 archivos que cada cliente tiene en
su propia carpeta de Drive (ver conversor_extractos.py / importador_xubio.py
/ orquestador.py). Acá solo hay interfaz.

Correr local:  streamlit run app.py     Requisitos: pip install -r requirements.txt
"""
import datetime
import os

import streamlit as st

import drive_io
import orquestador as O

st.set_page_config(page_title="Conciliación Bancaria", page_icon="🏦", layout="wide")


# ---------------------------------------------------------------------------
# Acceso
# ---------------------------------------------------------------------------
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
            st.session_state["auth_ok"] = True
            st.rerun()
        else:
            st.error("Contraseña incorrecta.")
    return False


if not acceso_ok():
    st.stop()


# ---------------------------------------------------------------------------
# ID del Maestro de Clientes (infraestructura del sistema, no dato de un
# cliente puntual -- por eso sale de los secrets, igual que las credenciales).
# ---------------------------------------------------------------------------
def _maestro_clientes_id() -> str:
    try:
        v = st.secrets["maestro_clientes_id"]
        if v:
            return v
    except Exception:
        pass
    v = os.environ.get("MAESTRO_CLIENTES_ID")
    if v:
        return v
    st.error("Falta configurar 'maestro_clientes_id' en los secrets de la app "
              "(el ID del Google Sheet MAESTRO_CLIENTES -- la parte del link "
              "entre '/d/' y '/edit').")
    st.stop()


# ---------------------------------------------------------------------------
# Ayudantes de UI
# ---------------------------------------------------------------------------
def _mostrar_reporte(rep: dict) -> None:
    if rep.get("saltado"):
        st.info(f"⏭️ {rep['motivo']}")
        return
    if not rep["ok"]:
        st.warning(f"⚠️ {rep['motivo']}")
        return

    st.success(f"✅ Procesado — {len(rep['pdfs_procesados'])} PDF(s), "
               f"{rep.get('pendientes', 0)} movimiento(s) pendiente(s) de imputar.")
    for nombre in rep["pdfs_procesados"]:
        st.caption(f"📄 {nombre}")
    for nombre, error in rep["pdfs_con_error"]:
        st.error(f"No pude procesar {nombre}: {error}")

    cols = st.columns(max(len(rep["controles"]), 1))
    for c, ctrl in zip(cols, rep["controles"]):
        c.metric(f"Control saldo — Cta {ctrl['cuenta']}",
                  "OK ✓" if ctrl["ok"] else "REVISAR ✗",
                  f"dif {ctrl['diferencia']:,.2f}")

    for cob in rep["cobertura"]:
        if cob["completo"]:
            st.caption(f"Cta {cob['cuenta']} ({cob['banco']}): importador COMPLETO ✓")
        else:
            st.caption(f"⚠️ Cta {cob['cuenta']} ({cob['banco']}): faltan imputar "
                       f"${cob['faltan_monto']:,.2f} en {cob['faltan_movimientos']} "
                       f"movimiento(s) — ver hoja A REVISAR del papel de trabajo.")


MESES_LABEL = [m.title() for m in O.MESES_ES]

st.title("🏦 Conciliación Bancaria")
st.caption("Todo el criterio de clasificación sale de las planillas de cada cliente en Drive — "
           "esta app no decide nada por su cuenta.")

tab_lote, tab_individual = st.tabs(["📦 Procesar en lote", "👤 Cliente individual"])

# ---------------------------------------------------------------------------
# Pestaña LOTE
# ---------------------------------------------------------------------------
with tab_lote:
    st.subheader("Procesar todos los clientes de un período")
    c1, c2, c3 = st.columns([1, 1, 1])
    mes_lote = c1.selectbox("Mes", options=list(range(1, 13)),
                             format_func=lambda m: MESES_LABEL[m - 1],
                             index=datetime.date.today().month - 1, key="lote_mes")
    anio_lote = c2.number_input("Año", min_value=2020, max_value=2100,
                                 value=datetime.date.today().year, step=1, key="lote_anio")
    forzar_lote = c3.checkbox("Reprocesar aunque ya esté hecho", value=False, key="lote_forzar")

    if st.button("▶️ Procesar lote", key="lote_run"):
        maestro_id = _maestro_clientes_id()
        clientes = O.listar_clientes(maestro_id)
        if not clientes:
            st.warning("El Maestro de Clientes no tiene ningún cliente cargado todavía.")
        else:
            progreso = st.progress(0.0, text=f"0 / {len(clientes)}")
            reportes = []
            for i, cliente in enumerate(clientes):
                with st.expander(f"{cliente['razon_social']}", expanded=False):
                    if not cliente["link_carpeta"]:
                        st.info("⏭️ Sin LINK CARPETA cargado en el Maestro — salteado.")
                        reportes.append(dict(cliente=cliente["razon_social"], ok=False, saltado=False))
                    else:
                        rep = O.procesar_cliente(cliente["razon_social"], cliente["link_carpeta"],
                                                  int(anio_lote), int(mes_lote), forzar=forzar_lote)
                        reportes.append(rep)
                        _mostrar_reporte(rep)
                progreso.progress((i + 1) / len(clientes), text=f"{i + 1} / {len(clientes)}")

            ok = sum(1 for r in reportes if r["ok"] and not r.get("saltado"))
            saltados = sum(1 for r in reportes if r.get("saltado"))
            con_error = sum(1 for r in reportes if not r["ok"])
            st.divider()
            st.markdown(f"**Lote terminado.** Procesados: {ok} · Salteados (ya estaban): "
                        f"{saltados} · Con problema: {con_error}")

# ---------------------------------------------------------------------------
# Pestaña INDIVIDUAL
# ---------------------------------------------------------------------------
with tab_individual:
    st.subheader("Procesar un cliente puntual")
    maestro_id_ind = _maestro_clientes_id()
    clientes_ind = O.listar_clientes(maestro_id_ind)

    if not clientes_ind:
        st.warning("El Maestro de Clientes no tiene ningún cliente cargado todavía.")
    else:
        nombres = [c["razon_social"] for c in clientes_ind]
        c1, c2, c3 = st.columns([2, 1, 1])
        elegido = c1.selectbox("Cliente", options=nombres, key="ind_cliente")
        mes_ind = c2.selectbox("Mes", options=list(range(1, 13)),
                                format_func=lambda m: MESES_LABEL[m - 1],
                                index=datetime.date.today().month - 1, key="ind_mes")
        anio_ind = c3.number_input("Año", min_value=2020, max_value=2100,
                                    value=datetime.date.today().year, step=1, key="ind_anio")

        cliente_data = next(c for c in clientes_ind if c["razon_social"] == elegido)
        bancos_declarados = O.leer_bancos_maestro(maestro_id_ind).get(elegido, [])
        if bancos_declarados:
            st.caption("Bancos declarados en el Maestro: " + ", ".join(bancos_declarados))
        if cliente_data["cuit"]:
            st.caption(f"CUIT: {cliente_data['cuit']}")

        forzar_ind = st.checkbox("Reprocesar aunque ya esté hecho", value=False, key="ind_forzar")

        if not cliente_data["link_carpeta"]:
            st.warning("Este cliente no tiene LINK CARPETA cargado en el Maestro todavía.")
        elif st.button("▶️ Procesar este cliente", key="ind_run"):
            rep = O.procesar_cliente(elegido, cliente_data["link_carpeta"],
                                      int(anio_ind), int(mes_ind), forzar=forzar_ind)
            _mostrar_reporte(rep)
