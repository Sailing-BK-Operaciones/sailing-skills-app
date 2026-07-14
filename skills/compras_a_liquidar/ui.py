"""
Compras a Liquidar — UI Streamlit.
Identifica compras pendientes de liquidar para comitentes con saldo deudor
vencido > $99.900 que van a recibir toma de caución.
"""
import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Compras a Liquidar")
    st.markdown(
        "Identifica las compras pendientes de liquidar para comitentes con "
        "**saldo deudor vencido > $99.900** (que van a recibir toma de caución). "
        "Cruza SALPESO + OPEVEN + CONTBOLE + ESPECIES."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivos del día** (subir manualmente):
        - `SALPESO.XLS` — saldos vencidos por comitente (hoja `Listado_de_Saldos`)
        - `OPEVEN.XLS` — operaciones a vencer (hoja `Operaciones_Vencer`)

        **Archivos de referencia** (Archivos Compartidos):
        - `CONTBOLE.XLS` — boletos del día (hoja `Control_de_Boletos`)
        - `ESPECIES.XLS` — maestro de especies (col `Norm.` para ticker)

        **Excel generado — 2 hojas:**
        1. **Resumen** — una fila por comitente deudor con `Saldo Deudor`,
           `Compras OPEVEN`, `Compras CI` y `Total Compras`. Fila TOTAL al pie.
           Filas con compras CI en naranja.
        2. **Detalle** — una fila por operación (OPEVEN + CONTBOLE CI), con
           ticker, importe, fechas y boleto. AutoFilter en encabezados.

        **Lógica:**
        - Deudor = `SALPESO.Saldo Vencido > 99.900` (positivo = debe dinero)
        - Compras OPEVEN = `Concepto == 'CPRA'` AND `Fec.Liq.` no vacío
        - Compras CI = `peracion == 'CPRA'` AND `Fec_Ope == Fec_Liq` (T+0)
        - Los comitentes deudores sin compras igual aparecen en Resumen (con $0)
        """)

    # ── Archivos del día (locales) ────────────────────────────────────────────
    st.subheader("Archivos del día")
    salpeso_file = st.file_uploader(
        "SALPESO.XLS *",
        type=["xls", "xlsx"],
        key="cal_salpeso",
        help="Saldos vencidos por comitente (hoja Listado_de_Saldos).",
    )
    opeven_file = st.file_uploader(
        "OPEVEN.XLS *",
        type=["xls", "xlsx"],
        key="cal_opeven",
        help="Operaciones a vencer (hoja Operaciones_Vencer).",
    )

    # ── Archivos de referencia (shared) ───────────────────────────────────────
    st.subheader("Archivos de referencia")
    contbole_file = shared_or_upload(
        "shared_contbole", "CONTBOLE.XLS", ["xls", "xlsx"], "cal_contbole"
    )
    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "cal_especies"
    )

    st.divider()

    # ── Validación ────────────────────────────────────────────────────────────
    faltan = []
    if not salpeso_file:  faltan.append("SALPESO.XLS")
    if not opeven_file:   faltan.append("OPEVEN.XLS")
    if not contbole_file: faltan.append("CONTBOLE.XLS")
    if not especies_file: faltan.append("ESPECIES.XLS")
    if faltan:
        st.caption(f"Falta para habilitar: {', '.join(faltan)}")

    # ── Ejecución ─────────────────────────────────────────────────────────────
    if st.button(
        "Generar reporte",
        type="primary",
        use_container_width=True,
        disabled=bool(faltan),
    ):
        with st.spinner("Procesando compras a liquidar..."):
            try:
                from skills.compras_a_liquidar.logic import generar_reporte
                output, resumen = generar_reporte(
                    salpeso_file=salpeso_file,
                    opeven_file=opeven_file,
                    contbole_file=contbole_file,
                    especies_file=especies_file,
                )
                st.session_state["cal_result"] = {
                    "output":  output,
                    "resumen": resumen,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados (persisten entre reruns) ───────────────────────────────────
    result = st.session_state.get("cal_result")
    if not result:
        return

    output  = result["output"]
    resumen = result["resumen"]

    # ── Métricas ──────────────────────────────────────────────────────────────
    if resumen["n_deudores"] == 0:
        st.success("✓ Sin comitentes con saldo deudor vencido > $99.900 hoy.")
    else:
        st.success(
            f"✓ {resumen['n_deudores']} comitente(s) deudor(es) procesados — "
            f"Fecha {resumen['fecha']}"
        )

    c1, c2, c3 = st.columns(3)
    c1.metric("Deudores",       resumen["n_deudores"])
    c2.metric("Compras OPEVEN", resumen["n_ops_opeven"])
    c3.metric("Compras CI",     resumen["n_ops_ci"])

    # Totales por moneda
    m1, m2, m3, m4 = st.columns(4)
    m1.metric("OPEVEN ARS",   f"{resumen['total_opeven_ars']:,.2f}")
    m2.metric("OPEVEN MEP",   f"USD {resumen['total_opeven_mep']:,.2f}")
    m3.metric("OPEVEN Cable", f"USD {resumen['total_opeven_cable']:,.2f}")
    m4.metric("CI (ARS)",     f"{resumen['total_ci']:,.2f}")

    if resumen.get("cttes_con_usd"):
        st.warning(
            f"Comitentes con compras en USD (MEP/Cable): "
            f"{', '.join(resumen['cttes_con_usd'])}",
            icon="💵",
        )
    if resumen["cttes_con_ci"]:
        st.warning(
            f"Comitentes con compras CI hoy (T+0): "
            f"{', '.join(resumen['cttes_con_ci'])}",
            icon="🟧",
        )

    with st.expander(f"Ver detalle por comitente ({resumen['n_deudores']})"):
        for d in resumen["detalle_deudores"]:
            flags = []
            if d["ctte"] in resumen.get("cttes_con_usd", []): flags.append("**USD**")
            if d["ctte"] in resumen["cttes_con_ci"]:          flags.append("**CI HOY**")
            flag_str = " · " + " · ".join(flags) if flags else ""
            st.markdown(
                f"- **{d['ctte']}** · {d['nombre']} · "
                f"Deudor: `{d['saldo_deudor']:,.2f}` · "
                f"ARS: `{d['opeven_ars']:,.2f}` · "
                f"MEP: `{d['opeven_mep']:,.2f}` · "
                f"Cable: `{d['opeven_cable']:,.2f}` · "
                f"CI: `{d['total_ci']:,.2f}`{flag_str}"
            )

    # ── Descarga ──────────────────────────────────────────────────────────────
    fname = f"Compras a Liquidar {resumen['fecha']}.xlsx"
    st.download_button(
        label=f"⬇ Descargar {fname}",
        data=output,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="cal_dl",
    )
