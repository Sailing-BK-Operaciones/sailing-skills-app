import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Arreglo Dev Gara Gallo")
    st.markdown(
        "Convierte el archivo BARRIDO exportado por Gallo (extensión incorrecta) al formato "
        "SI2 válido para subir a NASDAQ, y genera el resumen **Withdraw.txt** agrupado "
        "por nodo CLIENT / HOUSE."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Paso 1 — Subir el archivo de Gallo:**
        - Gallo exporta el BARRIDO con una extensión incorrecta (`BARRIDO.SI2-XXXX.XXXXXXX`).
          Subilo tal cual; la skill lo convierte a `BARRIDO.SI2` válido.

        **Paso 2 — ESPECIES.XLS** se toma automáticamente de Archivos Compartidos.

        **Outputs generados:**
        - `BARRIDO.SI2` — archivo corregido listo para subir a NASDAQ BYMA Clearing
        - `Withdraw.txt` — resumen de retiros por nodo:
          ```
          NODO CLIENT 80233/555555555
          Codigo CVSA;Ticker;VN

          NODO HOUSE 80233/222222222
          Codigo CVSA;Ticker;VN
          ```

        **Formato SI2:** separado por `;` — col 4 = CVSA, col 8 = nodo, col 11 = VN.
        """)

    # ── Inputs ───────────────────────────────────────────────────────────────────
    st.subheader("Archivo del día")
    barrido_file = st.file_uploader(
        "Archivo BARRIDO de Gallo (cualquier extensión — BARRIDO.SI2-XXXX.XXXXXXX o similar)",
        type=None,
        key="adgg_barrido",
        help="Exportado desde Gallo; se convierte automáticamente al formato SI2 correcto.",
    )

    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "adgg_esp"
    )

    st.divider()

    if not barrido_file:
        st.info("Subí el archivo BARRIDO de Gallo para habilitar el procesamiento.")
        return
    if not especies_file:
        st.info("Falta: ESPECIES.XLS (cargalo desde Archivos Compartidos).")
        return

    # ── Session state ─────────────────────────────────────────────────────────────
    if "adgg_result" not in st.session_state:
        st.session_state["adgg_result"] = None

    # ── Botón procesar ─────────────────────────────────────────────────────────────
    if st.button("Procesar", type="primary", use_container_width=True):
        with st.spinner("Procesando BARRIDO..."):
            try:
                from skills.arreglo_dev_gara_gallo.logic import procesar
                si2_bytes, withdraw_text, resumen = procesar(barrido_file, especies_file)
                st.session_state["adgg_result"] = {
                    "si2_bytes":     si2_bytes,
                    "withdraw_text": withdraw_text,
                    "resumen":       resumen,
                    "filename_orig": barrido_file.name,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados ────────────────────────────────────────────────────────────────
    result = st.session_state.get("adgg_result")
    if not result:
        return

    r = result["resumen"]
    st.success(f"✓ Procesado correctamente — origen: `{result['filename_orig']}`")

    col_a, col_b = st.columns(2)
    with col_a:
        st.metric("CLIENT — especies", r["n_client"])
        st.metric("CLIENT — VN total", f"{r['vn_client']:,}")
    with col_b:
        st.metric("HOUSE — especies", r["n_house"])
        st.metric("HOUSE — VN total", f"{r['vn_house']:,}")

    if r["n_sin_ticker"] > 0:
        st.warning(f"{r['n_sin_ticker']} CVSA(s) sin ticker en ESPECIES.XLS (aparecen como '???').")

    col_dl1, col_dl2 = st.columns(2)
    with col_dl1:
        st.download_button(
            label="⬇ Descargar BARRIDO.SI2",
            data=result["si2_bytes"],
            file_name="BARRIDO.SI2",
            mime="text/plain",
            use_container_width=True,
            key="adgg_dl_si2",
        )
    with col_dl2:
        st.download_button(
            label="⬇ Descargar Withdraw.txt",
            data=result["withdraw_text"].encode("utf-8"),
            file_name="Withdraw.txt",
            mime="text/plain",
            use_container_width=True,
            key="adgg_dl_withdraw",
        )
