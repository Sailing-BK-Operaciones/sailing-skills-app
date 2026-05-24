import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control Márgenes Gara BYMA")
    st.markdown(
        "Valoriza posiciones en garantía BYMA por comitente, detecta ventas a liquidar "
        "y disponibles para reemplazo. Determina estado **CUBIERTO / DESCUBIERTO** "
        "por comitente y fecha de vencimiento."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivo del día — subir acá:**
        - **SALDOS DEUDORES.xlsx** — comitentes con monto requerido (col B) y fecha VTO (col C).
          Fila 1 = título, fila 2 = encabezados, fila 3 en adelante = datos.
          Un comitente puede tener múltiples filas (una por vencimiento).

        **Archivos de referencia** (de Archivos Compartidos):
        - SAGACLTE.XLS, SATECLTE.XLS, ESPECIES.XLS, PC\\*.XLS, TABCOMPB.XLS,
          PDF de aforos BYMA, table-accounts\\*.csv, CONTBOLE.XLS

        **Output — `Control-Posiciones-Gara-DD-MM-AAAA.xlsx`:**
        - Hoja **RESUMEN**: una fila por comitente con totales, estado CUBIERTO/DESCUBIERTO y Account ID.
        - Hojas **C-{CTTE}**: panel de cobertura por fecha + tabla POSICIONES EN GARANTIA BYMA.
          Si hay VTAs: columnas adicionales de impacto y tabla **GESTIÓN REQUERIDA** con activos de reemplazo.

        **Nota sobre CONTBOLE.XLS:** debe corresponder al día de proceso para que las operaciones CI
        sean detectadas correctamente.
        """)

    # ── Archivo del día ───────────────────────────────────────────────────────
    st.subheader("Archivo del día")
    saldos_file = st.file_uploader(
        "SALDOS DEUDORES.xlsx",
        type=["xlsx", "xls"],
        key="cmgb_saldos",
        help="Comitentes con monto requerido (col B) y fecha VTO (col C)"
    )

    # ── Archivos de referencia (shared) ───────────────────────────────────────
    st.subheader("Archivos de referencia")
    col1, col2, col3 = st.columns(3)
    with col1:
        sagaclte_file = shared_or_upload("shared_sagaclte",   "SAGACLTE.XLS",         ["xls", "xlsx"], "cmgb_saga")
        tabcompb_file = shared_or_upload("shared_tabcompb",   "TABCOMPB.XLS",         ["xls", "xlsx"], "cmgb_tab")
        pc_file       = shared_or_upload("shared_pc",         "PC*.XLS",              ["xls", "xlsx"], "cmgb_pc")
    with col2:
        sateclte_file = shared_or_upload("shared_sateclte",   "SATECLTE.XLS",         ["xls", "xlsx"], "cmgb_sate")
        especies_file = shared_or_upload("shared_especies",   "ESPECIES.XLS",         ["xls", "xlsx"], "cmgb_esp")
        accounts_file = shared_or_upload("shared_accounts",   "table-accounts_*.csv", ["csv"],         "cmgb_acc")
    with col3:
        contbole_file = shared_or_upload("shared_contbole",   "CONTBOLE.XLS",         ["xls", "xlsx"], "cmgb_cont")
        pdf_file      = shared_or_upload("shared_pdf_aforos", "PDF de aforos BYMA",   ["pdf"],         "cmgb_pdf")

    st.divider()

    # ── Validación ────────────────────────────────────────────────────────────
    faltantes = []
    if not saldos_file:   faltantes.append("SALDOS DEUDORES.xlsx")
    if not sagaclte_file: faltantes.append("SAGACLTE.XLS")
    if not sateclte_file: faltantes.append("SATECLTE.XLS")
    if not especies_file: faltantes.append("ESPECIES.XLS")
    if not tabcompb_file: faltantes.append("TABCOMPB.XLS")
    if not pc_file:       faltantes.append("PC*.XLS")
    if not pdf_file:      faltantes.append("PDF de aforos BYMA")
    if not contbole_file: faltantes.append("CONTBOLE.XLS")
    if not accounts_file: faltantes.append("table-accounts_*.csv")

    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
        return

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button("Generar control", type="primary", use_container_width=True):
        with st.spinner("Procesando posiciones en garantía..."):
            try:
                from skills.control_margenes_gara_byma.logic import generar_control
                xlsx_bytes, resumen, advertencias = generar_control(
                    saldos_file   = saldos_file,
                    contbole_file = contbole_file,
                    sagaclte_file = sagaclte_file,
                    sateclte_file = sateclte_file,
                    especies_file = especies_file,
                    tabcompb_file = tabcompb_file,
                    pc_file       = pc_file,
                    pdf_aforos_file = pdf_file,
                    accounts_file = accounts_file,
                )

                # Estado global
                n_desc = resumen["n_descubiertos"]
                if n_desc == 0:
                    st.success(
                        f"✓ Todos los comitentes CUBIERTOS — "
                        f"{resumen['n_comitentes']} procesados"
                    )
                else:
                    st.error(
                        f"⚠ {n_desc} comitente(s) DESCUBIERTO(S) de "
                        f"{resumen['n_comitentes']} procesados"
                    )

                # Métricas
                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Comitentes",    resumen["n_comitentes"])
                col_b.metric("Cubiertos",     resumen["n_cubiertos"])
                col_c.metric("Descubiertos",  resumen["n_descubiertos"])
                col_d.metric(
                    "Diferencia total",
                    f"ARS {resumen['total_diferencia']:,.0f}"
                )

                # Descarga
                fecha_str = resumen["fecha"]
                st.download_button(
                    label=f"⬇ Descargar Control-Posiciones-Gara-{fecha_str}.xlsx",
                    data=xlsx_bytes,
                    file_name=f"Control-Posiciones-Gara-{fecha_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

                # Advertencias
                if advertencias:
                    with st.expander(f"⚠ {len(advertencias)} advertencia(s) de procesamiento"):
                        for adv in advertencias:
                            st.markdown(f"- {adv}")

            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
