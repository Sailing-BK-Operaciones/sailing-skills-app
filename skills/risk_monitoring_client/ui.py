import streamlit as st
import traceback
from datetime import date

def render():
    st.title("Risk Monitoring Client")
    st.markdown("Panel de riesgo por cuenta BYMA Clearing — Portfolio Risk, Portfolio Value, Variation Margin, Total Margin Deficit.")
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        1. Descarga el CSV desde BYMA Clearing: `table-grouping-details-table_*.csv`
        2. Descarga el CSV de cuentas: `table-accounts_*.csv`
        3. Subi ambos archivos abajo
        4. Hace clic en **Generar reporte**
        5. Descarga el Excel de resultado
        """)

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

    st.divider()

    if csv_grouping and csv_accounts:
        if st.button("Generar reporte", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    from skills.risk_monitoring_client.logic import generar_reporte
                    resultado = generar_reporte(csv_grouping, csv_accounts)
                    nombre = f"Risk Monitoring Client {date.today().strftime('%d-%m-%Y')}.xlsx"
                    st.success("Reporte generado correctamente")
                    st.download_button(
                        label="Descargar Excel",
                        data=resultado,
                        file_name=nombre,
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True
                    )
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                    with st.expander("Detalle del error"):
                        st.code(traceback.format_exc())
    else:
        st.info("Subi los dos archivos para habilitar el boton de generacion.")
