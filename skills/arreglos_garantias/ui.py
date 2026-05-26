import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Arreglos Garantías NASDAQ")
    st.markdown(
        "Genera archivos SI2 (DELIVER / RECEIVE) y resúmenes TXT para arreglos manuales "
        "de garantías en NASDAQ BYMA. Produce además el archivo **Gallo XLSX** con las "
        "hojas de entrega y devoluciones."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivo del día — subir acá:**
        - **INPUTS ARREGLOS GARA.xlsx** — Hoja1 con columnas:
          `CTTE | MOVIMIENTO (ENVIAR/TRAER) | CODIGO (CVSA o Ticker) | VN | FECHA VTO`

        **Archivo de referencia** (de Archivos Compartidos):
        - `ESPECIES.XLS` — maestro de especies (mapeo CVSA ↔ Ticker, tipo de precio)

        **Archivos generados:**
        - `Req-envio de gtias DD-MM-AA.SI2` — instrucciones DELIVER (enviar garantías a NASDAQ)
        - `Resumen DEPOSIT DD-MM-AA.txt` — resumen de depósitos por nodo CLIENT / HOUSE / DEFAULT FUND
        - `Ret-devolucion gtia DD-MM-AA.SI2` — instrucciones RECEIVE (devolver garantías)
        - `Resumen WITHDRAW DD-MM-AA.txt` — resumen de retiros por nodo
        - `Gallo DD-MM-AA.xlsx` — dos hojas: ENTREGA GTIAS + DEVOLUCIONES GTIAS

        **Contadores AGE/AGD:** persisten durante la sesión para evitar referencias duplicadas
        si se ejecuta más de una vez en el día. Usar el botón de reset solo al comenzar el día.

        **Nodos de contraparte:**
        - `80233/555555555` — CLIENT (comitentes estándar)
        - `80233/222222222` — HOUSE (comitentes 1000–1003)
        - `880233/888888888` — DEFAULT FUND (comitente 888888888)
        """)

    # ── Inputs ────────────────────────────────────────────────────────────────
    st.subheader("Archivo del día")
    inputs_file = st.file_uploader(
        "INPUTS ARREGLOS GARA.xlsx",
        type=["xlsx", "xls"],
        key="ag_inputs",
        help="Hoja1: CTTE | MOVIMIENTO | CODIGO | VN | FECHA VTO"
    )

    st.subheader("Archivo de referencia")
    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "ag_esp"
    )

    # ── Contadores AGE / AGD ──────────────────────────────────────────────────
    st.divider()
    if "ag_counter" not in st.session_state:
        st.session_state["ag_counter"] = {}
    if "ag_result" not in st.session_state:
        st.session_state["ag_result"] = None

    counter = st.session_state["ag_counter"]
    col_ctr, col_btn = st.columns([3, 1])
    with col_ctr:
        if counter:
            st.caption(
                "Contadores activos: "
                + ", ".join(f"{k}={v}" for k, v in sorted(counter.items()))
            )
        else:
            st.caption("Contadores en 0 — el próximo proceso arranca desde AGE001 / AGD001.")
    with col_btn:
        if st.button("Reiniciar contadores", key="ag_reset"):
            st.session_state["ag_counter"] = {}
            st.session_state["ag_result"] = None
            st.rerun()

    st.divider()

    # ── Validación ────────────────────────────────────────────────────────────
    faltantes = []
    if not inputs_file:
        faltantes.append("INPUTS ARREGLOS GARA.xlsx")
    if not especies_file:
        faltantes.append("ESPECIES.XLS")

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
        return

    # ── Botón generar ─────────────────────────────────────────────────────────
    if st.button("Generar archivos", type="primary", use_container_width=True, key="ag_gen"):
        with st.spinner("Generando archivos NASDAQ..."):
            try:
                from skills.arreglos_garantias.logic import generar_arreglo
                outputs, new_counter, resumen, advertencias = generar_arreglo(
                    inputs_file=inputs_file,
                    especies_file=especies_file,
                    counter_state=st.session_state["ag_counter"],
                )
                st.session_state["ag_counter"] = new_counter
                st.session_state["ag_result"] = {
                    "outputs": outputs,
                    "resumen": resumen,
                    "advertencias": advertencias,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados — persisten en session_state entre reruns ──────────────────
    result = st.session_state.get("ag_result")
    if not result:
        return

    outputs     = result["outputs"]
    resumen     = result["resumen"]
    advertencias = result["advertencias"]

    n_env = resumen["n_enviar"]
    n_tra = resumen["n_traer"]
    partes = []
    if n_env:
        partes.append(f"{n_env} DELIVER")
    if n_tra:
        partes.append(f"{n_tra} RECEIVE")
    st.success(
        f"Archivos generados — fecha {resumen['fecha']} — "
        + " | ".join(partes)
    )

    # ── Métricas ──────────────────────────────────────────────────────────────
    col_a, col_b, col_c, col_d = st.columns(4)
    col_a.metric("Instrucciones ENVIAR", n_env)
    col_b.metric("Instrucciones TRAER",  n_tra)
    col_c.metric("Refs AGE", resumen["refs_age"] or "—")
    col_d.metric("Refs AGD", resumen["refs_agd"] or "—")

    # ── Descargas ─────────────────────────────────────────────────────────────
    st.subheader("Descargas")
    col1, col2 = st.columns(2)

    with col1:
        if outputs["req_envio"]:
            content, fname = outputs["req_envio"]
            st.download_button(
                label=f"⬇ {fname}",
                data=content, file_name=fname,
                mime="text/plain", use_container_width=True,
                key="ag_dl_req",
            )
        else:
            st.info("Sin instrucciones DELIVER (no hay filas ENVIAR).")

        if outputs["resumen_deposit"]:
            content, fname = outputs["resumen_deposit"]
            st.download_button(
                label=f"⬇ {fname}",
                data=content, file_name=fname,
                mime="text/plain", use_container_width=True,
                key="ag_dl_dep",
            )

    with col2:
        if outputs["ret_devolucion"]:
            content, fname = outputs["ret_devolucion"]
            st.download_button(
                label=f"⬇ {fname}",
                data=content, file_name=fname,
                mime="text/plain", use_container_width=True,
                key="ag_dl_ret",
            )
        else:
            st.info("Sin instrucciones RECEIVE (no hay filas TRAER).")

        if outputs["resumen_withdraw"]:
            content, fname = outputs["resumen_withdraw"]
            st.download_button(
                label=f"⬇ {fname}",
                data=content, file_name=fname,
                mime="text/plain", use_container_width=True,
                key="ag_dl_wit",
            )

    # Gallo XLSX — ancho completo
    if outputs["gallo"]:
        content, fname = outputs["gallo"]
        st.download_button(
            label=f"⬇ {fname}  (ENTREGA GTIAS + DEVOLUCIONES GTIAS)",
            data=content, file_name=fname,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="ag_dl_gallo",
        )

    # ── Advertencias ──────────────────────────────────────────────────────────
    if advertencias:
        with st.expander(f"⚠ {len(advertencias)} advertencia(s) de procesamiento"):
            for adv in advertencias:
                st.markdown(f"- {adv}")
