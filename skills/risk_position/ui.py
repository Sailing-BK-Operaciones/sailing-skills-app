import streamlit as st
import traceback


def render():
    st.title("Risk Position")
    st.markdown(
        "Compara saldos de garantía BYMA Clearing vs **Saldos Gara a cubrir** de Gallo. "
        "Completa columnas E (Account ID BC), F (Importe BC) y H (Diferencia) en el Excel."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        1. Descarga `table-riskPositions_*.csv` desde BYMA Clearing → Risk Positions
        2. Descarga `table-accounts_*.csv` desde BYMA Clearing → Accounts
        3. Subi el archivo **Saldos Gara a cubrir.xlsx** (el original de Gallo)
        4. Hace clic en **Generar** — vas a descargar el Excel ya actualizado con las columnas E, F y H completas

        > La fecha de proceso y el TC BYMA se leen automáticamente desde la celda C y G1 del Excel de Gallo.
        """)

    col1, col2 = st.columns(2)
    with col1:
        csv_risk = st.file_uploader(
            "CSV Risk Position (table-riskPositions_*.csv)",
            type=["csv"], key="rp_risk"
        )
    with col2:
        csv_accounts = st.file_uploader(
            "CSV de cuentas (table-accounts_*.csv)",
            type=["csv"], key="rp_accounts"
        )

    saldos_file = st.file_uploader(
        "Saldos Gara a cubrir.xlsx",
        type=["xlsx"], key="rp_saldos"
    )

    st.divider()

    faltantes = []
    if not csv_risk:     faltantes.append("CSV Risk Position")
    if not csv_accounts: faltantes.append("CSV de cuentas")
    if not saldos_file:  faltantes.append("Saldos Gara a cubrir.xlsx")

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
    else:
        if st.button("Generar", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    from skills.risk_position.logic import generar_reporte
                    resultado, r = generar_reporte(csv_risk, csv_accounts, saldos_file)

                    st.success("Excel actualizado correctamente")

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Fecha de proceso", r["fecha_proceso"])
                    col_b.metric("TC BYMA Clearing", f"{r['tc_byma']:,.2f}")
                    col_c.metric("Filas cauciones", r["n_caucion_rows"])

                    col_d, col_e, col_f = st.columns(3)
                    col_d.metric("Comitentes Gallo", r["n_gallo"])
                    col_e.metric("Comitentes BC", r["n_bc"])
                    col_f.metric("Con match exacto", r["n_match"])

                    if r["solo_en_gallo"]:
                        st.warning(
                            f"Solo en Gallo (sin posición BC): "
                            f"{', '.join(r['solo_en_gallo'])}"
                        )
                    if r["solo_en_bc"]:
                        st.error(
                            f"Solo en BC (no están en Saldos Gara): "
                            f"{', '.join(r['solo_en_bc'])}"
                        )

                    st.download_button(
                        label="Descargar Saldos Gara a cubrir actualizado",
                        data=resultado,
                        file_name="Saldos Gara a cubrir.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                    with st.expander("Detalle del error"):
                        st.code(traceback.format_exc())
