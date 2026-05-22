import streamlit as st
import traceback
from datetime import date


def render():
    st.title("Collateral Position")
    st.markdown(
        "Valorización de títulos en garantía en los 3 nodos BYMA Clearing "
        "(Client 556 / House 557 / Def Fund 1121). "
        "Si subís SAGACLTE genera además las hojas de conciliación BC vs Gallo."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        **Archivos del día:**
        - CSV de Collateral Positions de BYMA Clearing (`table-collateralPositions_*.csv`)
        - `PC*.XLS` — precios de cierre de Gallo

        **Archivos de referencia** (cambian ocasionalmente):
        - `ESPECIES.XLS` — maestro de especies (tipo de precio: Normal / Porc.)
        - PDF de aforos BYMA (`Listas de Especies Aceptadas como Garantía.pdf`)

        **Opcionales — activan hojas de conciliación BC vs Gallo:**
        - `SAGACLTE.XLS` — stock de garantías por comitente (hoja `Saldos_de_Garantias`)
        - `table-accounts_*.csv` — Account ID por comitente

        **Parámetro:** TC USD/ARS para pesificar cash en dólares (default 1400).
        """)

    # ── Archivos del día ──────────────────────────────────────────────────────
    st.subheader("Archivos del día")
    col1, col2, col3 = st.columns([2, 1, 1])
    with col1:
        csv_file = st.file_uploader(
            "CSV Collateral Positions (table-collateralPositions_*.csv)",
            type=["csv"], key="cp_csv"
        )
    with col2:
        pc_file = st.file_uploader(
            "Precios de cierre (PC*.XLS)",
            type=["xls","xlsx"], key="cp_pc"
        )
    with col3:
        tc_usd = st.number_input(
            "TC USD/ARS", min_value=1.0, value=1400.0, step=10.0, key="cp_tc"
        )

    # ── Archivos de referencia ────────────────────────────────────────────────
    st.subheader("Archivos de referencia")
    col4, col5 = st.columns(2)
    with col4:
        especies_file = st.file_uploader(
            "ESPECIES.XLS", type=["xls","xlsx"], key="cp_esp"
        )
    with col5:
        pdf_file = st.file_uploader(
            "PDF de aforos BYMA", type=["pdf"], key="cp_pdf"
        )

    # ── Archivos opcionales (conciliación) ────────────────────────────────────
    with st.expander("Conciliación BC vs Gallo (opcional)"):
        col6, col7 = st.columns(2)
        with col6:
            sagaclte_file = st.file_uploader(
                "SAGACLTE.XLS (activa hojas de conciliación)",
                type=["xls","xlsx"], key="cp_saga"
            )
        with col7:
            accounts_file = st.file_uploader(
                "table-accounts_*.csv (Account ID por comitente)",
                type=["csv"], key="cp_acc"
            )

    st.divider()

    requeridos = {"CSV Collateral": csv_file, "PC*.XLS": pc_file,
                  "ESPECIES.XLS": especies_file, "PDF aforos": pdf_file}
    faltantes = [k for k, v in requeridos.items() if v is None]

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
    else:
        if sagaclte_file:
            st.success("SAGACLTE cargado — se generarán las hojas de conciliación BC vs Gallo.")

        if st.button("Generar reporte", type="primary", use_container_width=True):
            with st.spinner("Procesando..."):
                try:
                    from skills.Collateral_position.logic import generar_reporte
                    xlsx_buf, r = generar_reporte(
                        csv_file, pc_file, especies_file, pdf_file,
                        sagaclte_file=sagaclte_file,
                        accounts_file=accounts_file,
                        tc_usd=float(tc_usd),
                        fecha_proceso=date.today(),
                    )

                    st.success("Reporte generado correctamente")

                    # Métricas por nodo
                    cols = st.columns(len(r["summary"]) + 1)
                    for i, (short, n_sec, monto) in enumerate(r["summary"]):
                        cols[i].metric(short, f"ARS {monto:,.0f}", f"{n_sec} especies")
                    cols[-1].metric("TOTAL GENERAL", f"ARS {r['grand_monto']:,.0f}")

                    # Alertas
                    if r["n_sin_precio"] > 0:
                        st.warning(f"{r['n_sin_precio']} especies sin precio de cierre (fondo amarillo en el Excel).")
                    if r["n_aforos"] == 0:
                        st.warning("PDF de aforos no produjo datos — verificar formato del archivo.")
                    if r["tiene_conc"]:
                        st.info("Hojas de conciliación BC vs Gallo incluidas.")

                    fecha_str = date.today().strftime("%d-%m-%y")
                    st.download_button(
                        label="Descargar Excel",
                        data=xlsx_buf,
                        file_name=f"Collateral Position {fecha_str}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        use_container_width=True,
                    )

                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                    with st.expander("Detalle del error"):
                        st.code(traceback.format_exc())
