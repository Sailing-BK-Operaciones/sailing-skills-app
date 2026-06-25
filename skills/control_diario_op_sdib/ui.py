"""
Control diario OP-SDIB (MA) — UI Streamlit.
Concilia operaciones PPT BYMA + SENEBI vs CONTBOLE (Gallo) del día.
"""
import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control diario OP-SDIB (MA)")
    st.markdown(
        "Concilia las operaciones del día entre **BYMA Clearing** (OPERSECEXT_GARA) "
        "y **Gallo** (CONTBOLE): PPT garantizado + SENEBI bilateral vs boletos. "
        "Valoriza por moneda, detecta diferencias, faltantes de boleto en Gallo, "
        "operaciones compensadas via CTA 1002 y proyecta el saldo del día (Control 999)."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Archivos del día:**
        - `OPERSECEXT_GARA.DAT` — operaciones PPT BYMA — **obligatorio**
        - `OPERBILEXT_GARA.DAT` — operaciones SENEBI bilateral — **opcional**
        - `CONTBOLE.XLS` — boletos del día desde Gallo — se busca en Archivos Compartidos
        - `Actual Position DD-MM-AA.xlsx` — **input local de esta skill** (no compartido).
          Si está presente, habilita el **Control 999** (saldo proyectado a fin del día por moneda).

        **Excel generado — 10 hojas:**
        1. **Control 999** — saldo Actual Position + movimientos del día = saldo final proyectado
        2-4. **Pesos ARS / Dolar MEP / USD Cable** — valorización por concepto (CI/CN/CAU/TI/MAV) y segmento
        5. **Resumen Ope-Titulos** — conteo por estado + alertas (FALTA BOLETO / ARANCEL CARTERA)
        6. **Diferencias a Verificar** — filtrable, solo filas con discrepancia
        7. **Operaciones** — PPT + SENEBI vs Gallo, con CONCEPTO y SEGMENTO
        8. **Cauciones** — VN BYMA (capital) vs Gallo capital y total
        9. **Trading Intraday** — ops COTR/VTTR de Gallo sin contrapartida BYMA
        10. **MAV** — ops VCHM/CCHM de Gallo (informacional)

        **Lógica CTA 1002:** si la diferencia de la CTA 1002 es compensada por otros comitentes
        (neto ≈ 0), las filas involucradas se reclasifican como **OK** con nota "CTA 1002".
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

    st.subheader("Archivos de referencia")
    contbole_file = shared_or_upload(
        "shared_contbole", "CONTBOLE.XLS", ["xls", "xlsx"], "sdib_contbole"
    )
    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "sdib_especies"
    )

    st.subheader("Actual Position del día (opcional — habilita Control 999)")
    ap_file = st.file_uploader(
        "Actual Position DD-MM-AA.xlsx — opcional",
        type=["xlsx"],
        key="sdib_ap",
        help="Reporte Actual Position generado por la skill homónima. "
             "Aporta el saldo proyectado base por moneda; con eso el Control 999 "
             "calcula el saldo final del día (saldo inicio + movs del día).",
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
                    ap_file=ap_file,
                    especies_file=especies_file,
                )
                st.session_state["sdib_result"] = {
                    "output":  output,
                    "resumen": resumen,
                }
            except Exception as e:
                st.error(f"Error al conciliar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    # ── Resultados (persisten entre reruns) ───────────────────────────────────
    result = st.session_state.get("sdib_result")
    if not result:
        return

    output  = result["output"]
    resumen = result["resumen"]

    # ── Alertas críticas ──────────────────────────────────────────────────────
    n_falta   = resumen["n_falta_boleto"]
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
                    f"{r['sentido']} · {r.get('concepto', '')} [{r.get('segmento_mer', '')}] — "
                    f"VN: {r['vn_b']:,.0f} | IMP: {r['imp_b']:,.2f}"
                )
    else:
        senebi_tag = " | Con SENEBI" if resumen["con_senebi"] else ""
        st.success(
            f"✓ Sin faltantes de boleto en Gallo — Proceso {resumen['process_date']}"
            + senebi_tag
        )

    if n_arancel > 0:
        st.error(
            f"ARANCEL EN CUENTA DE CARTERA PROPIA — {n_arancel} boleto(s) "
            f"en cuentas 1000-1003 con arancel distinto de cero",
            icon="🔴",
        )

    # ── Valorización por moneda ───────────────────────────────────────────────
    st.subheader("Valorización por moneda (Neto BYMA vs Gallo)")
    mc1, mc2, mc3 = st.columns(3)
    for col, code in zip((mc1, mc2, mc3), ("ARS", "USD_MEP", "USD_CABLE")):
        info = resumen["moneda_estado"][code]
        delta = "OK" if info["estado"] == "OK" else f"DIF {info['dif']:,.2f}"
        col.metric(
            info["label"],
            f"{info['neto_g']:,.2f}",
            delta=delta,
            delta_color="off" if info["estado"] == "OK" else "inverse",
        )

    # ── Control 999 (si hay Actual Position) ──────────────────────────────────
    if resumen["saldos_finales"]:
        st.subheader(f"Control 999 — Saldo proyectado del día")
        st.caption(f"Base: {resumen['ap_filename']}")
        sc1, sc2, sc3 = st.columns(3)
        for col, code in zip((sc1, sc2, sc3), ("ARS", "USD_MEP", "USD_CABLE")):
            sf = resumen["saldos_finales"].get(code)
            if sf:
                col.metric(
                    f"Saldo final — {sf['label']}",
                    f"{sf['saldo_final']:,.2f}",
                    delta=f"Inicio: {sf['saldo_ini']:,.2f}",
                    delta_color="off",
                )
    else:
        st.info("Sin Actual Position: Control 999 mostrará 'N/D' (movimientos sí se cargan).")

    # ── Métricas operaciones ──────────────────────────────────────────────────
    st.subheader("Operaciones (PPT + SENEBI)")
    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("OK",              resumen["n_ops_ok"])
    c2.metric("Diferencia",      resumen["n_ops_dif"])
    c3.metric("Solo BYMA",       resumen["n_solo_byma"])
    c4.metric("Solo Gallo",      resumen["n_solo_gallo"])
    c5.metric("Trading Intrad.", resumen["n_ti"])

    cc1, cc2, cc3, cc4 = st.columns(4)
    cc1.metric("Cauciones OK",         resumen["n_cau_ok"])
    cc2.metric("Cauciones Diferencia", resumen["n_cau_dif"])
    cc3.metric("MAV (informacional)",  resumen["n_mav"])
    cc4.metric("Diferencias a verif.", resumen["n_ver_rows"])

    # ── Descarga ──────────────────────────────────────────────────────────────
    fname = f"Conciliacion OpersecEXT {resumen['fecha']}.xlsx"
    st.download_button(
        label=f"⬇ Descargar {fname}",
        data=output,
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="sdib_dl",
    )
