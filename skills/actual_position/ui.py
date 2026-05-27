import streamlit as st
import traceback
from datetime import date
from skills.shared_ui import shared_or_upload


def render():
    st.title("Actual Position")
    st.markdown(
        "Panel de liquidación diaria BYMA Clearing — "
        "hojas por moneda (ARS / MEP / Cable) + Movimientos por Especie + Verificación. "
        "También genera el archivo de Opciones en formato Gallo (.xls)."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        st.markdown("""
        **Requerido:**
        - CSV de Actual Positions descargado de BYMA Clearing (`table-currentActualPositions_*.csv`)

        **Opcionales (mejoran el reporte):**
        - `ESPECIES.XLS` — agrega el Código CVSA en la hoja de Movimientos por Especie
          (se toma desde Archivos Compartidos si ya está cargado)
        - `saldos al inicio Nasdaq.csv` — muestra el Saldo proyectado del día en cada hoja de moneda

        La fecha de proceso por defecto es hoy. Si el CSV es de otro día, cambiala abajo.
        """)

    col1, col2 = st.columns([2, 1])
    with col1:
        csv_file = st.file_uploader(
            "CSV Actual Positions (table-currentActualPositions_*.csv)",
            type=["csv"], key="ap_csv"
        )
    with col2:
        process_date = st.date_input("Fecha de proceso", value=date.today(), key="ap_fecha")

    with st.expander("Archivos opcionales"):
        col3, col4 = st.columns(2)
        with col3:
            especies_file = shared_or_upload(
                "shared_especies", "ESPECIES.XLS (para Código CVSA)", ["xls", "xlsx"], "ap_especies"
            )
        with col4:
            saldos_file = shared_or_upload(
                "shared_saldos_nasdaq", "saldos al inicio Nasdaq.csv", ["csv"], "ap_saldos"
            )

    st.divider()

    # ── Inicializar result state ───────────────────────────────────────────────
    if "ap_result" not in st.session_state:
        st.session_state["ap_result"] = None

    if not csv_file:
        st.info("Subí el CSV de Actual Positions para habilitar la generación.")
        return

    # ── Botón generar ─────────────────────────────────────────────────────────
    if st.button("Generar reporte", type="primary", use_container_width=True, key="ap_gen"):
        with st.spinner("Procesando..."):
            try:
                from skills.actual_position.logic import generar_reporte
                xlsx_buf, xls_buf, r = generar_reporte(
                    csv_file, especies_file, saldos_file, process_date
                )
                st.session_state["ap_result"] = {
                    "xlsx_buf":    xlsx_buf,
                    "xls_buf":     xls_buf,
                    "r":           r,
                    "fecha_str":   process_date.strftime("%d-%m-%y"),
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados — persisten en session_state entre reruns ──────────────────
    result = st.session_state.get("ap_result")
    if not result:
        return

    xlsx_buf  = result["xlsx_buf"]
    xls_buf   = result["xls_buf"]
    r         = result["r"]
    fecha_str = result["fecha_str"]

    st.success("Reporte generado correctamente")

    col_a, col_b, col_c = st.columns(3)
    col_a.metric("Fecha de proceso",   r["process_date"])
    col_b.metric("Filas clasificadas", r["clasificados"])
    col_c.metric("Para verificar",     r["verificacion"])

    flags = []
    if not r["tiene_especies"]: flags.append("ESPECIES.XLS no subido — sin Código CVSA")
    if not r["tiene_saldos"]:   flags.append("Saldos inicio no subidos — sin saldo proyectado")
    if r["verificacion"] > 0:
        flags.append(f"Hay {r['verificacion']} filas en Verificación: {', '.join(r['verif_assets'])}")
    for msg in flags:
        st.warning(msg)

    col_d, col_e = st.columns(2)
    with col_d:
        st.download_button(
            label="⬇ Descargar Excel (.xlsx)",
            data=xlsx_buf,
            file_name=f"Actual Position {fecha_str}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="ap_dl_xlsx",
        )
    with col_e:
        if r["n_op"] > 0:
            st.download_button(
                label=f"⬇ Descargar Opciones Gallo (.xls) — {r['n_op']} filas",
                data=xls_buf,
                file_name="Agente 233 - Opciones Sobre Titulos Valores.xls",
                mime="application/vnd.ms-excel",
                use_container_width=True,
                key="ap_dl_xls",
            )
        else:
            st.info("No hay posiciones OP — archivo de Opciones Gallo vacío.")
