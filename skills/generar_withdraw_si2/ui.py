import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Genera SI2 NASDAQ")
    st.markdown(
        "Procesa los archivos SI2 de garantías para NASDAQ BYMA Clearing. "
        "Corrige el nodo de la cuenta `233/1000` (HOUSE exportada como CLIENT por Gallo) "
        "y genera los resúmenes **Withdraw** y/o **Deposit**."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        Subí uno o ambos archivos SI2 según el flujo del día:

        | Archivo | Operación | Outputs |
        |---------|-----------|---------|
        | `BARRIDO*.SI2` | **Withdraw** — retiro de garantías | `BARRIDO.SI2` + `Withdraw.txt` |
        | `TRANSFERENCIA*.SI2` | **Deposit** — envío de garantías | `TRANSFERENCIA.SI2` + `Deposit.txt` |

        **ESPECIES.XLS** se toma automáticamente de Archivos Compartidos.

        **Corrección automática de nodo:** si la cuenta `233/1000` aparece bajo el nodo CLIENT
        (error de exportación de Gallo), la skill la reasigna al nodo HOUSE en el SI2 de salida.
        Cuando Gallo lo corrija en origen, la corrección no tendrá efecto.

        **Resumen (.txt):**
        ```
        NODO CLIENT 80233/555555555
        Codigo CVSA;Ticker;VN
        00508;EDN;5000

        NODO HOUSE 80233/222222222
        Codigo CVSA;Ticker;VN
        08691;SLV;755
        ```
        """)

    # ── Inputs ───────────────────────────────────────────────────────────────────
    st.subheader("Archivos del día")
    col1, col2 = st.columns(2)
    with col1:
        barrido_file = st.file_uploader(
            "BARRIDO*.SI2 (Withdraw)",
            type=None,
            key="gwsi2_barrido",
            help="Archivo de retiro de garantías.",
        )
    with col2:
        transferencia_file = st.file_uploader(
            "TRANSFERENCIA*.SI2 (Deposit)",
            type=None,
            key="gwsi2_transf",
            help="Archivo de envío de garantías.",
        )

    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "gwsi2_esp"
    )

    st.divider()

    if not barrido_file and not transferencia_file:
        st.info("Subí al menos un archivo SI2 (BARRIDO o TRANSFERENCIA) para habilitar el procesamiento.")
        return
    if not especies_file:
        st.info("Falta: ESPECIES.XLS (cargalo desde Archivos Compartidos).")
        return

    # ── Session state ─────────────────────────────────────────────────────────────
    if "gwsi2_result" not in st.session_state:
        st.session_state["gwsi2_result"] = None

    # ── Botón procesar ─────────────────────────────────────────────────────────────
    if st.button("Procesar", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                from skills.generar_withdraw_si2.logic import procesar
                withdraw_result, deposit_result = procesar(
                    barrido_file, transferencia_file, especies_file
                )
                st.session_state["gwsi2_result"] = {
                    "withdraw": withdraw_result,
                    "deposit":  deposit_result,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados ────────────────────────────────────────────────────────────────
    result = st.session_state.get("gwsi2_result")
    if not result:
        return

    # ── WITHDRAW ──────────────────────────────────────────────────────────────────
    w = result.get("withdraw")
    if w:
        st.success("✓ WITHDRAW (BARRIDO) procesado correctamente")
        if w["reasignadas"] > 0:
            st.warning(
                f"⚠ {w['reasignadas']} fila(s) de cuenta 233/1000 reasignadas de CLIENT → HOUSE."
            )
        col_a, col_b = st.columns(2)
        with col_a:
            st.metric("CLIENT — especies", w["n_client"])
            st.metric("CLIENT — VN total", f"{w['vn_client']:,}")
        with col_b:
            st.metric("HOUSE — especies", w["n_house"])
            st.metric("HOUSE — VN total", f"{w['vn_house']:,}")
        if w["n_sin_ticker"] > 0:
            st.warning(f"{w['n_sin_ticker']} CVSA(s) sin ticker en ESPECIES.XLS (aparecen como '???').")
        col_dl1, col_dl2 = st.columns(2)
        with col_dl1:
            st.download_button(
                label="⬇ Descargar BARRIDO.SI2",
                data=w["si2_bytes"],
                file_name="BARRIDO.SI2",
                mime="text/plain",
                use_container_width=True,
                key="gwsi2_dl_barrido",
            )
        with col_dl2:
            st.download_button(
                label="⬇ Descargar Withdraw.txt",
                data=w["txt_content"].encode("utf-8"),
                file_name="Withdraw.txt",
                mime="text/plain",
                use_container_width=True,
                key="gwsi2_dl_withdraw",
            )

    # ── DEPOSIT ───────────────────────────────────────────────────────────────────
    d = result.get("deposit")
    if d:
        if w:
            st.divider()
        st.success("✓ DEPOSIT (TRANSFERENCIA) procesado correctamente")
        if d["reasignadas"] > 0:
            st.warning(
                f"⚠ {d['reasignadas']} fila(s) de cuenta 233/1000 reasignadas de CLIENT → HOUSE."
            )
        col_c, col_dd = st.columns(2)
        with col_c:
            st.metric("CLIENT — especies", d["n_client"])
            st.metric("CLIENT — VN total", f"{d['vn_client']:,}")
        with col_dd:
            st.metric("HOUSE — especies", d["n_house"])
            st.metric("HOUSE — VN total", f"{d['vn_house']:,}")
        if d["n_sin_ticker"] > 0:
            st.warning(f"{d['n_sin_ticker']} CVSA(s) sin ticker en ESPECIES.XLS (aparecen como '???').")
        col_dl3, col_dl4 = st.columns(2)
        with col_dl3:
            st.download_button(
                label="⬇ Descargar TRANSFERENCIA.SI2",
                data=d["si2_bytes"],
                file_name="TRANSFERENCIA.SI2",
                mime="text/plain",
                use_container_width=True,
                key="gwsi2_dl_transf",
            )
        with col_dl4:
            st.download_button(
                label="⬇ Descargar Deposit.txt",
                data=d["txt_content"].encode("utf-8"),
                file_name="Deposit.txt",
                mime="text/plain",
                use_container_width=True,
                key="gwsi2_dl_deposit",
            )
