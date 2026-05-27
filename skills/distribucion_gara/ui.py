import streamlit as st
import traceback
from datetime import date
from pathlib import Path
from skills.shared_ui import shared_or_upload

_TEMPLATE_PATH = Path(__file__).parent / "templates" / "Saldos_Gara_a_cubrir_template.xlsx"


def render():
    st.title("Distribución Garantías BYMA")
    st.markdown(
        "Distribución de garantías para comitentes con saldo a cubrir. "
        "Genera Excel con hoja RESUMEN + Gallo + hojas por comitente (C-XXXX) "
        "+ hojas VTO para cauciones canceladas."
    )
    st.divider()

    with st.expander("Como usar esta skill"):
        if _TEMPLATE_PATH.exists():
            st.download_button(
                label="⬇ Descargar plantilla Saldos Gara a cubrir.xlsx",
                data=_TEMPLATE_PATH.read_bytes(),
                file_name="Saldos Gara a cubrir.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                key="dg_dl_template",
            )
        st.markdown("""
        **Archivos del día** (cambian cada rueda):
        - `SAGACLTE.XLS` — stock de garantías en BYMA (hoja `Saldos_de_Garantias`) — desde Archivos Compartidos
        - `SATECLTE.XLS` — tenencia disponible por comitente (hoja `Saldos_de_Tenencia`)
        - `PC*.XLS` — precios de cierre del día — desde Archivos Compartidos
        - `Saldos Gara a cubrir.xlsx` — comitentes con monto requerido en ARS

        **Archivos de referencia** (cambian ocasionalmente):
        - `ESPECIES.XLS` — maestro de especies — desde Archivos Compartidos
        - PDF de aforos BYMA — desde Archivos Compartidos
        - `CONTBOLE.XLS` — boletos del día desde Gallo (hoja `Control_de_Boletos`)
        - `tabla listas gallo vs aforos.xlsx` — aforo SAIL por lista Gallo
        - `table-accounts_*.csv` — Account ID por comitente — desde Archivos Compartidos
        - `TABCOMPB.XLS` — clasificador de tipos de operación (opcional, activa ops CI)

        **Opcional:**
        - `Risk Monitoring Client *.xlsx` — agrega columnas BC (deficit, validación) al RESUMEN

        **Rescates/Amortizaciones:** si hubo hoy, ingresarlos en el campo de texto antes de generar.
        """)

    # ── Archivos del día ─────────────────────────────────────────────────────────
    st.subheader("Archivos del día")
    col1, col2 = st.columns(2)
    with col1:
        sagaclte_file = shared_or_upload("shared_sagaclte", "SAGACLTE.XLS (garantías BYMA)", ["xls", "xlsx"], "dg_saga")
        sateclte_file = shared_or_upload("shared_sateclte", "SATECLTE.XLS (tenencia disponible)", ["xls", "xlsx"], "dg_sate")
    with col2:
        pc_file      = shared_or_upload("shared_pc", "Precios de cierre PC*.XLS", ["xls", "xlsx"], "dg_pc")
        saldos_file  = st.file_uploader("Saldos Gara a cubrir.xlsx", type=["xlsx"], key="dg_saldos")

    # ── Archivos de referencia ───────────────────────────────────────────────────
    st.subheader("Archivos de referencia")
    col3, col4 = st.columns(2)
    with col3:
        especies_file   = shared_or_upload("shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "dg_esp")
        pdf_file        = shared_or_upload("shared_pdf_aforos", "PDF aforos BYMA", ["pdf"], "dg_pdf")
        contbole_file   = shared_or_upload("shared_contbole", "CONTBOLE.XLS (boletos del día)", ["xls", "xlsx"], "dg_cont")
    with col4:
        aforo_sail_file = shared_or_upload("shared_aforo_sail", "tabla listas gallo vs aforos.xlsx", ["xlsx"], "dg_aforo")
        accounts_file   = shared_or_upload("shared_accounts", "table-accounts_*.csv", ["csv"], "dg_acc")
        tabcompb_file   = shared_or_upload("shared_tabcompb", "TABCOMPB.XLS (activa ops CI)", ["xls", "xlsx"], "dg_tabcompb")

    # ── Opcional: RMC ────────────────────────────────────────────────────────────
    with st.expander("Archivo opcional — Risk Monitoring Client"):
        rmc_file = st.file_uploader(
            "Risk Monitoring Client *.xlsx (agrega columnas BC al RESUMEN)",
            type=["xlsx"], key="dg_rmc"
        )

    # ── Rescates / Amortizaciones ────────────────────────────────────────────────
    st.subheader("Rescates / Amortizaciones del día")
    rescates_texto = st.text_area(
        "Una línea por especie afectada (dejar vacío si no hubo):",
        placeholder="TICKER RESCATE\nTICKER AMORTIZACION 20",
        height=100,
        key="dg_rescates",
        help="Formato: TICKER RESCATE  |  TICKER AMORTIZACION XX  (XX = porcentaje amortizado)",
    )

    st.divider()

    # ── Validación de campos obligatorios ────────────────────────────────────────
    requeridos = {
        "SAGACLTE.XLS":      sagaclte_file,
        "SATECLTE.XLS":      sateclte_file,
        "PC*.XLS":           pc_file,
        "Saldos Gara":       saldos_file,
        "ESPECIES.XLS":      especies_file,
        "PDF aforos":        pdf_file,
        "CONTBOLE.XLS":      contbole_file,
        "Tabla aforo SAIL":  aforo_sail_file,
        "table-accounts":    accounts_file,
    }
    faltantes = [k for k, v in requeridos.items() if v is None]
    if faltantes:
        st.info(f"Faltan: {', '.join(faltantes)}")
        return

    if not tabcompb_file:
        st.warning(
            "TABCOMPB.XLS no disponible — las operaciones CI de CONTBOLE no se clasificarán "
            "(sin impacto si no hay ops CI hoy)."
        )

    # ── Generar ──────────────────────────────────────────────────────────────────
    if st.button("Generar reporte", type="primary", use_container_width=True):
        with st.spinner("Procesando..."):
            try:
                from skills.distribucion_gara.logic import generar_reporte
                xlsx_buf, r = generar_reporte(
                    sagaclte_file=sagaclte_file,
                    sateclte_file=sateclte_file,
                    especies_file=especies_file,
                    pdf_aforos_file=pdf_file,
                    pc_file=pc_file,
                    saldos_file=saldos_file,
                    contbole_file=contbole_file,
                    aforo_sail_file=aforo_sail_file,
                    accounts_file=accounts_file,
                    tabcompb_file=tabcompb_file,
                    rmc_file=rmc_file,
                    rescates_texto=rescates_texto,
                    fecha_proceso=date.today(),
                )

                st.success("Reporte generado correctamente")

                col_a, col_b, col_c, col_d = st.columns(4)
                col_a.metric("Comitentes procesados", r["n_comitentes"])
                col_b.metric("Superávit", r["n_superavit"])
                col_c.metric("Déficit",   r["n_deficit"])
                col_d.metric("Filas Gallo", r["n_gallo_rows"])

                if r["validacion_issues"]:
                    lines = [
                        f"Ctte {v['ctte']}: Gallo {v['importe_g']:,.0f} | BC {v['deficit_bc']:,.0f} | Dif {v['diferencia']:,.0f}"
                        for v in r["validacion_issues"]
                    ]
                    st.warning(
                        "BYMA requiere más garantía que Gallo — revisar antes de operar:\n"
                        + "\n".join(lines)
                    )

                if r["cpr_found"]:
                    cpr_str = "; ".join(
                        f"Ctte {e['ctte']}: " + ", ".join(f"{t} ({vn:,})" for t, vn in e["items"])
                        for e in r["cpr_found"]
                    )
                    st.info(f"CPR a liquidar encontradas: {cpr_str}")

                if r["vto_cttes"]:
                    st.info(f"Hojas VTO generadas (cauciones canceladas): {', '.join(r['vto_cttes'])}")

                if r.get("contbole_warning"):
                    st.warning(r["contbole_warning"])

                if not r["tiene_rmc"]:
                    st.info("RMC no subido — columnas BC en RESUMEN quedarán sin dato.")

                fecha_str = date.today().strftime("%d-%m-%Y")
                st.download_button(
                    label="Descargar Excel",
                    data=xlsx_buf,
                    file_name=f"Distribucion-Gara-Byma-{fecha_str}.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )

            except ValueError as e:
                st.error(str(e))
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
