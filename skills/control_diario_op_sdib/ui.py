"""
Control diario Op-SDIB — UI Streamlit.
Concilia operaciones PPT BYMA + SENEBI vs CONTBOLE (Gallo) del día.
"""
import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control diario Op-SDIB")
    st.markdown(
        "Concilia las operaciones del día entre **BYMA Clearing** (OPERSECEXT_GARA) "
        "y **Gallo** (CONTBOLE): PPT garantizado + SENEBI bilateral vs boletos. "
        "Detecta diferencias, operaciones sin boleto en Gallo (**FALTA BOLETO**) "
        "y operaciones sin contrapartida en BYMA (**SOLO GALLO**)."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivos del día:**
        - `OPERSECEXT_GARA.DAT` — operaciones del segmento garantizado PPT BYMA — **obligatorio**
        - `OPERBILEXT_GARA.DAT` — operaciones SENEBI bilateral — **opcional** (si no hubo SENEBI, se omite)
        - `CONTBOLE.XLS` — boletos del día exportados de Gallo — se busca en Archivos Compartidos

        **Excel generado — 5 hojas:**
        - **Resumen** — conteo por estado con alertas críticas (FALTA BOLETO / ARANCEL CARTERA)
        - **Operaciones** — PPT + SENEBI vs Gallo: columnas SEGMENTO y VERIFICAR, colores por estado
        - **Cauciones** — VN BYMA (capital) vs capital y total de Gallo
        - **Trading Intraday** — operaciones COTR/VTTR de Gallo sin contrapartida en BYMA
        - **MAV** — operaciones MAV de Gallo (informacional, sin comparación)

        **Colores de estado:**
        - 🟢 Verde — OK
        - 🔴 Rojo — DIFERENCIA o FALTA BOLETO
        - 🟡 Amarillo — SOLO BYMA
        - 🔵 Azul — SOLO GALLO
        - 🟠 Naranja — TRADING INTRADAY

        **Lógica CTA 1002:** si la diferencia de VN de la CTA 1002 es compensada por otros comitentes
        (neto = 0), las filas involucradas se reclasifican automáticamente como **OK** con nota
        "CTA 1002" para trazabilidad.
        """)

    # ── Archivos del día ──────────────────────────────────────────────────────
    st.subheader("Archivos del día")

    dat_file = st.file_uploader(
        "OPERSECEXT_GARA.DAT *",
        type=["dat", "DAT"],
        key="sdib_dat",
        help="Operaciones PPT del segmento garantizado BYMA — obligatorio.",
    )

    bil_file = st.file_uploader(
        "OPERBILEXT_GARA.DAT — opcional",
        type=["dat", "DAT"],
        key="sdib_bil",
        help="Operaciones SENEBI bilateral. Si no hubo SENEBI en el día, se puede omitir.",
    )

    st.subheader("Archivo de referencia")
    contbole_file = shared_or_upload(
        "shared_contbole", "CONTBOLE.XLS", ["xls", "xlsx"], "sdib_contbole"
    )

    st.divider()

    # ── Validación ────────────────────────────────────────────────────────────
    faltantes = []
    if not dat_file:
        faltantes.append("OPERSECEXT_GARA.DAT")
    if not contbole_file:
        faltantes.append("CONTBOLE.XLS")

    if faltantes:
        st.caption(f"Falta para habilitar: {', '.join(faltantes)}")

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button(
        "Conciliar operaciones",
        type="primary",
        use_container_width=True,
        disabled=bool(faltantes),
    ):
        with st.spinner("Conciliando operaciones..."):
            try:
                from skills.control_diario_op_sdib.logic import generar_reporte
                output, resumen = generar_reporte(
                    dat_file=dat_file,
                    xls_file=contbole_file,
                    bil_file=bil_file,
                )

                # ── Alertas críticas ──────────────────────────────────────────
                n_falta  = resumen["n_falta_boleto"]
                n_arancel = resumen["n_arancel"]

                if n_falta > 0:
                    st.error(
                        f"FALTA BOLETO EN GALLO — {n_falta} operación(es) en BYMA "
                        f"sin boleto correspondiente",
                        icon="🔴",
                    )
                    with st.expander(f"Ver detalle — {n_falta} faltante(s)"):
                        for r in resumen["falta_boleto_detail"]:
                            st.markdown(
                                f"- CTTE **{r['ctte']}** · {r['especie']} · {r['moneda']} · "
                                f"{r['sentido']} [{r['segmento']}] — "
                                f"VN: {r['vn_b']:,.0f} | IMP: {r['imp_b']:,.2f}"
                            )
                else:
                    senebi_tag = " | Con SENEBI" if resumen["con_senebi"] else ""
                    st.success(
                        f"✓ Sin faltantes de boleto en Gallo — Fecha {resumen['fecha']}"
                        + senebi_tag
                    )

                if n_arancel > 0:
                    st.error(
                        f"ARANCEL EN CUENTA DE CARTERA PROPIA — {n_arancel} boleto(s) "
                        f"en cuentas 1000-1003 con arancel distinto de cero",
                        icon="🔴",
                    )

                # ── Métricas — Operaciones ────────────────────────────────────
                st.subheader("Operaciones (PPT + SENEBI)")
                c1, c2, c3, c4, c5 = st.columns(5)
                c1.metric("OK",             resumen["n_ops_ok"])
                c2.metric("Diferencia",     resumen["n_ops_dif"])
                c3.metric("Solo BYMA",      resumen["n_solo_byma"])
                c4.metric("Solo Gallo",     resumen["n_solo_gallo"])
                c5.metric("Trading Intrad.", resumen["n_ti"])

                # ── Métricas — Cauciones y MAV ────────────────────────────────
                cc1, cc2, cc3 = st.columns(3)
                cc1.metric("Cauciones OK",        resumen["n_cau_ok"])
                cc2.metric("Cauciones Diferencia", resumen["n_cau_dif"])
                cc3.metric("MAV (informacional)",  resumen["n_mav"])

                # ── Descarga ──────────────────────────────────────────────────
                fname = f"Conciliacion OpersecEXT {resumen['fecha']}.xlsx"
                st.download_button(
                    label=f"⬇ Descargar {fname}",
                    data=output,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="sdib_dl",
                )

            except Exception as e:
                st.error(f"Error al conciliar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
