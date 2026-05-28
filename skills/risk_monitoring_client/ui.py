import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


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
        **Archivo del dia:**
        - `table-grouping-details-table_*.csv` — descargarlo desde BYMA Clearing → Risk Monitoring

        **Archivos de referencia** (se cargan desde Archivos Compartidos si ya están disponibles):
        - `table-accounts_*.csv` — Account ID por comitente
        - `PC*.XLS` — precios de cierre de Gallo
        - `ESPECIES.XLS` — maestro de especies (col 26 = haircut BYMA API; aforo = (100−haircut)/100)
        - `SAGACLTE.XLS` — saldos de garantías por comitente
        """)

    st.subheader("Archivo del dia")
    csv_grouping = st.file_uploader(
        "CSV Risk Monitoring (table-grouping-details-table_*.csv)",
        type=["csv"], key="rmc_grouping"
    )

    st.subheader("Archivos de referencia")
    col1, col2 = st.columns(2)
    with col1:
        csv_accounts  = shared_or_upload("shared_accounts",   "table-accounts_*.csv",          ["csv"],         "rmc_accounts")
        pc_file       = shared_or_upload("shared_pc",         "Precios de cierre (PC*.XLS)",   ["xls", "xlsx"], "rmc_pc")
    with col2:
        especies_file = shared_or_upload("shared_especies",   "Maestro de especies (ESPECIES.XLS)", ["xls", "xlsx"], "rmc_esp")
        sagaclte_file = shared_or_upload("shared_sagaclte",   "Garantías por comitente (SAGACLTE.XLS)", ["xls", "xlsx"], "rmc_saga")

    st.divider()

    faltantes = []
    if not csv_grouping:  faltantes.append("CSV Risk Monitoring")
    if not csv_accounts:  faltantes.append("table-accounts_*.csv")
    if not pc_file:       faltantes.append("PC*.XLS")
    if not especies_file: faltantes.append("ESPECIES.XLS")
    if not sagaclte_file: faltantes.append("SAGACLTE.XLS")

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
    else:
        if st.button("Generar reporte", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    from skills.risk_monitoring_client.logic import generar_reporte
                    resultado, fecha_output, n_data, grand_total, total_garantias = generar_reporte(
                        csv_grouping, csv_accounts, pc_file, especies_file, sagaclte_file
                    )

                    st.success("Reporte generado correctamente")

                    col_a, col_b, col_c = st.columns(3)
                    col_a.metric("Cuentas procesadas", n_data)
                    col_b.metric("Total Margin Deficit", f"ARS {grand_total:,.0f}")
                    col_c.metric("Total Garantias integradas", f"ARS {total_garantias:,.0f}")

                    st.download_button(
                        label="Descargar Excel",
                        data=resultado,
                        file_name=f"Risk Monitoring Client {fecha_output}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                    with st.expander("Detalle del error"):
                        st.code(traceback.format_exc())
