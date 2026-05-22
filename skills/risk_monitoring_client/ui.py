import streamlit as st
import traceback


def render():
    st.title("Risk Monitoring Client")
    st.markdown(
        "Panel de riesgo por cuenta BYMA Clearing — "
        "Portfolio Risk, Portfolio Value, Variation Margin, Total Margin Deficit, "
        "Garantías integradas a aforo BYMA."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        **Archivos del dia** (cambian cada ejecucion):
        1. Descarga `table-grouping-details-table_*.csv` desde BYMA Clearing → Risk Monitoring
        2. Descarga `table-accounts_*.csv` desde BYMA Clearing → Accounts

        **Archivos de referencia** (se actualizan ocasionalmente):
        3. `PC*.XLS` — precios de cierre de Gallo
        4. `ESPECIES.XLS` — maestro de especies de Gallo
        5. `SAGACLTE.XLS` — saldos de garantías por comitente (hoja `Saldos_de_Garantias`)
        6. PDF de aforos BYMA (`Listas de Especies en garantia.pdf`)
        """)

    st.subheader("Archivos del dia")
    col1, col2 = st.columns(2)
    with col1:
        csv_grouping = st.file_uploader(
            "CSV Risk Monitoring (table-grouping-details-table_*.csv)",
            type=["csv"], key="rmc_grouping"
        )
    with col2:
        csv_accounts = st.file_uploader(
            "CSV de cuentas (table-accounts_*.csv)",
            type=["csv"], key="rmc_accounts"
        )

    st.subheader("Archivos de referencia")
    col3, col4, col5, col6 = st.columns(4)
    with col3:
        pc_file = st.file_uploader("Precios de cierre (PC*.XLS)", type=["xls", "xlsx"], key="rmc_pc")
    with col4:
        especies_file = st.file_uploader("Maestro de especies (ESPECIES.XLS)", type=["xls", "xlsx"], key="rmc_esp")
    with col5:
        sagaclte_file = st.file_uploader("Garantías por comitente (SAGACLTE.XLS)", type=["xls", "xlsx"], key="rmc_saga")
    with col6:
        pdf_aforos = st.file_uploader("PDF de aforos BYMA", type=["pdf"], key="rmc_pdf")

    st.divider()

    todos_subidos = all([csv_grouping, csv_accounts, pc_file, especies_file, sagaclte_file, pdf_aforos])

    faltantes = []
    if not csv_grouping:  faltantes.append("CSV Risk Monitoring")
    if not csv_accounts:  faltantes.append("CSV de cuentas")
    if not pc_file:       faltantes.append("Precios de cierre")
    if not especies_file: faltantes.append("ESPECIES.XLS")
    if not sagaclte_file: faltantes.append("SAGACLTE.XLS")
    if not pdf_aforos:    faltantes.append("PDF de aforos")

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
    else:
        if st.button("Generar reporte", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    from skills.risk_monitoring_client.logic import generar_reporte
                    resultado, fecha_output, n_data, grand_total, total_garantias = generar_reporte(
                        csv_grouping, csv_accounts, pc_file, especies_file, sagaclte_file, pdf_aforos
                    )

                    nombre = f"Risk Monitoring Client {fecha_output}.xlsx"
                    st.success("Reporte generado correctamente")

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Cuentas procesadas", n_data)
                    col_b.metric(
                        "Total Margin Deficit",
                        f"ARS {grand_total:,.0f}",
                        delta=None,
                    )
                    col_c.metric("Total Garantias integradas", f"ARS {total_garantias:,.0f}")

                    st.download_button(
                        label="Descargar Excel",
                        data=resultado,
                        file_name=nombre,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                    with st.expander("Detalle del error"):
                        st.code(traceback.format_exc())
