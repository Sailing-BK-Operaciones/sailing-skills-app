import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Genera TXT Gara NASDAQ")
    st.markdown(
        "Genera los archivos SI2 para NASDAQ BYMA a partir del Excel de Distribución de Garantías. "
        "Produce hasta 4 archivos: **Req-envio** (DELIVER) + **Resumen DEPOSIT** + "
        "**Ret-devolucion** (RECEIVE) + **Resumen WITHDRAW**."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        1. Subí el Excel **Distribucion-Gara-Byma-DD-MM-YYYY.xlsx** generado por la skill anterior.
        2. `ESPECIES.XLS` se toma desde Archivos Compartidos si ya está cargado.
        3. Hacé clic en **Generar** — descargás hasta 4 archivos.

        **Archivos generados:**
        - `Req-envio de gtias DD-MM-AA.SI2` — instrucciones DELIVER (enviar garantías a BYMA)
        - `Resumen DEPOSIT DD-MM-AA.txt` — resumen de depósitos por nodo CLIENT / HOUSE
        - `Ret-devolucion gtia DD-MM-AA.SI2` — instrucciones RECEIVE (retiro de garantías)
        - `Resumen WITHDRAW DD-MM-AA.txt` — resumen de retiros por nodo

        **Contadores TGE/TGD:** persisten durante la sesión para evitar referencias duplicadas
        si se ejecuta más de una vez en el día. Usar el botón de reset solo al comenzar el día.
        """)

    distribucion_file = st.file_uploader(
        "Distribucion-Gara-Byma-*.xlsx",
        type=["xlsx"], key="gtn_distrib"
    )
    especies_file = shared_or_upload("shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "gtn_esp")

    # Contadores de referencia persistidos en session_state
    if "gtn_counter" not in st.session_state:
        st.session_state["gtn_counter"] = {}

    counter = st.session_state["gtn_counter"]
    col_ctr, col_btn = st.columns([3, 1])
    with col_ctr:
        if counter:
            st.caption(f"Contadores activos: {', '.join(f'{k}={v}' for k, v in sorted(counter.items()))}")
        else:
            st.caption("Contadores en 0 — el próximo proceso arranca desde 001.")
    with col_btn:
        if st.button("Reiniciar contadores", key="gtn_reset"):
            st.session_state["gtn_counter"] = {}
            st.rerun()

    st.divider()

    if not distribucion_file:
        st.info("Subí el Excel de Distribución para habilitar la generación.")
        return
    if not especies_file:
        st.info("Falta: ESPECIES.XLS")
        return

    if st.button("Generar", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                from skills.genera_txt_gara_nasdaq.logic import generar_archivos
                outputs, r, new_counter = generar_archivos(
                    distribucion_file,
                    especies_file,
                    counter_state=st.session_state["gtn_counter"],
                )
                st.session_state["gtn_counter"] = new_counter

                st.success(f"Archivos generados — fecha {r['fecha']}")

                col_a, col_b, col_c = st.columns(3)
                col_a.metric("Comitentes", r["n_comitentes"])
                col_b.metric("Instrucciones DELIVER", r["n_enviar"])
                col_c.metric("Instrucciones RECEIVE", r["n_disponible"])

                if r["vto_sheets"]:
                    st.info(f"Hojas VTO procesadas (cauciones → RECEIVE): {', '.join(r['vto_sheets'])}")
                if r["unknown_actions"]:
                    st.warning(f"Acciones no reconocidas — revisar: {', '.join(r['unknown_actions'])}")

                col1, col2 = st.columns(2)
                with col1:
                    if outputs["req_envio"]:
                        content, fname = outputs["req_envio"]
                        _, _, n = outputs["req_envio_refs"]
                        st.download_button(
                            f"Descargar Req-envio ({n} instrucciones)",
                            data=content, file_name=fname,
                            mime="text/plain", use_container_width=True,
                        )
                    else:
                        st.info("Sin instrucciones DELIVER.")

                    if outputs["resumen_deposit"]:
                        content, fname = outputs["resumen_deposit"]
                        nc, nh = outputs["resumen_deposit_stats"]
                        st.download_button(
                            f"Descargar Resumen DEPOSIT (Client: {nc} / House: {nh})",
                            data=content, file_name=fname,
                            mime="text/plain", use_container_width=True,
                        )

                with col2:
                    if outputs["ret_devolucion"]:
                        content, fname = outputs["ret_devolucion"]
                        _, _, n = outputs["ret_devolucion_refs"]
                        st.download_button(
                            f"Descargar Ret-devolucion ({n} instrucciones)",
                            data=content, file_name=fname,
                            mime="text/plain", use_container_width=True,
                        )
                    else:
                        st.info("Sin instrucciones RECEIVE.")

                    if outputs["resumen_withdraw"]:
                        content, fname = outputs["resumen_withdraw"]
                        nc, nh = outputs["resumen_withdraw_stats"]
                        st.download_button(
                            f"Descargar Resumen WITHDRAW (Client: {nc} / House: {nh})",
                            data=content, file_name=fname,
                            mime="text/plain", use_container_width=True,
                        )

            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
