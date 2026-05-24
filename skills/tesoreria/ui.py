import streamlit as st
import traceback
from skills.tesoreria.logic import MESES_ORDER, TC_HISTORICOS, procesar_mes


# Clave de session_state donde se acumulan los meses
_KEY = "tes_months"


def render():
    st.title("Tesorería — Reporte de Gestión")
    st.markdown(
        "Procesa archivos **MOVICTA** (XLS de ARS, MEP y USD de Gallo) y actualiza el reporte "
        "Excel acumulativo anual de Tesorería con los datos del mes seleccionado: "
        "Panel de Control, Análisis de Mercados, Análisis de Clientes, Ranking, "
        "Detalle ARS y Detalle USD."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Flujo de trabajo:**
        1. En la sección **Cargar mes**, seleccioná el mes, ingresá los TC de cierre y subí los 3 archivos MOVICTA.
        2. Hacé clic en **Procesar mes** — el mes queda guardado en la sesión.
        3. Repetí para todos los meses que quieras incluir.
        4. Cuando tengas todos los meses cargados, hacé clic en **Generar reporte**.

        **Archivos MOVICTA necesarios por mes:**
        - `MOVICTA_66_67_MES.XLS` — movimientos en AR$ (hoja `Movimientos_Cuenta`)
        - `MOVICTA_50_51_MEP_MES.XLS` — movimientos en U$D MEP
        - `MOVICTA_50_51_CABLE_MES.XLS` — movimientos en U$D Cable *(opcional)*

        **TC de referencia históricos (pre-cargados como sugerencia):**
        | Mes | TC MEP | TC CCL |
        |-----|--------|--------|
        | Enero 2026 | 1.464,75 | 1.510,53 |
        | Febrero 2026 | 1.436,64 | 1.476,11 |
        | Marzo 2026 | 1.488,00 | 1.499,00 |
        | Abril 2026 | 1.448,93 | 1.503,03 |

        **Clasificación de canal:**
        - **Digital (HB):** TRANSF.RECIBIDA, Solicitado x WEB, TRANSFERENCIA RECIBID, DOLAR MEP
        - **eCheq:** ECHEQ A CTTE *(se muestra por separado en Análisis de Clientes; en Panel se suma a % Manual)*
        - **Manual:** todas las demás referencias
        - **Anulaciones** (ANULADO / ANULA…): excluidas de cantidades y montos, contadas aparte como Errores
        """)

    # ── Inicializar estado ────────────────────────────────────────────────────
    if _KEY not in st.session_state:
        st.session_state[_KEY] = {}   # {mes: {"df":..., "tc_mep":..., "tc_ccl":...}}

    meses_cargados = st.session_state[_KEY]

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 0: Reporte base — siempre visible
    # ─────────────────────────────────────────────────────────────────────────
    st.subheader("Reporte base")
    base_file = st.file_uploader(
        "Reporte acumulativo existente (XLSX) — opcional",
        type=["xlsx"],
        key="tes_base",
        help=(
            "Si subís el reporte del mes anterior, los nuevos meses se agregan al existente. "
            "Sin archivo base, se genera un reporte nuevo desde cero con los meses cargados en sesión."
        ),
    )
    if base_file:
        st.caption(f"📎 Modo actualización — los meses procesados se agregarán a «{base_file.name}»")
    else:
        st.caption("📄 Modo nuevo reporte — se generará desde cero con los meses cargados en sesión")

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 1: Cargar mes
    # ─────────────────────────────────────────────────────────────────────────
    st.subheader("Cargar mes")

    col_mes, col_anio, col_tc1, col_tc2 = st.columns([2, 1, 1.5, 1.5])
    with col_mes:
        mes_sel = st.selectbox(
            "Mes",
            options=[m for m in MESES_ORDER if m not in meses_cargados],
            key="tes_mes_sel",
        )
    with col_anio:
        anio_sel = st.number_input(
            "Año",
            value=2026,
            min_value=2020, max_value=2099, step=1,
            key="tes_anio",
        )
    with col_tc1:
        tc_hist_mep = TC_HISTORICOS.get(mes_sel, {}).get('MEP', 0.0) if mes_sel else 0.0
        tc_mep = st.number_input(
            "TC MEP (cierre de mes)",
            value=tc_hist_mep,
            min_value=0.0, step=0.01, format="%.2f",
            key="tes_tc_mep",
            help="Tipo de cambio MEP de cierre para pesificar USD MEP",
        )
    with col_tc2:
        tc_hist_ccl = TC_HISTORICOS.get(mes_sel, {}).get('CCL', 0.0) if mes_sel else 0.0
        tc_ccl = st.number_input(
            "TC CCL (cierre de mes)",
            value=tc_hist_ccl,
            min_value=0.0, step=0.01, format="%.2f",
            key="tes_tc_ccl",
            help="Tipo de cambio CCL de cierre para pesificar USD Cable (dejar en 0 si no hay Cable)",
        )

    col_ars, col_mep, col_cable = st.columns(3)
    with col_ars:
        ars_file = st.file_uploader(
            "MOVICTA AR$ (XLS) *",
            type=["xls", "xlsx"], key="tes_ars",
        )
    with col_mep:
        mep_file = st.file_uploader(
            "MOVICTA U$D MEP (XLS) *",
            type=["xls", "xlsx"], key="tes_mep",
        )
    with col_cable:
        cable_file = st.file_uploader(
            "MOVICTA U$D Cable (XLS) — opcional",
            type=["xls", "xlsx"], key="tes_cable",
        )

    btn_ok = mes_sel and ars_file and mep_file and tc_mep > 0

    if not btn_ok:
        falt = []
        if not mes_sel:       falt.append("Mes")
        if not ars_file:      falt.append("MOVICTA AR$")
        if not mep_file:      falt.append("MOVICTA U$D MEP")
        if tc_mep == 0:       falt.append("TC MEP")
        if falt:
            st.caption(f"Falta para habilitar: {', '.join(falt)}")

    if st.button(
        f"Actualizar reporte — {mes_sel} {int(anio_sel)}" if mes_sel else "Actualizar reporte",
        disabled=not btn_ok, type="primary", key="tes_procesar",
        use_container_width=True,
    ):
        with st.spinner(f"Procesando {mes_sel} {int(anio_sel)}..."):
            try:
                df, advs = procesar_mes(
                    ars_file   = ars_file,
                    mep_file   = mep_file,
                    tc_mep     = tc_mep,
                    cable_file = cable_file if cable_file else None,
                    tc_ccl     = tc_ccl     if tc_ccl > 0 else None,
                )
                st.session_state[_KEY][mes_sel] = {
                    "df":     df,
                    "tc_mep": tc_mep,
                    "tc_ccl": tc_ccl if tc_ccl > 0 else None,
                }
                n_ops = len(df[df['Canal'] != 'Anulacion'])
                st.success(f"✓ {mes_sel} procesado — {n_ops} operaciones válidas.")
                if advs:
                    with st.expander(f"⚠ {len(advs)} advertencia(s) en {mes_sel}"):
                        for a in advs:
                            st.markdown(f"- {a}")
                st.rerun()
            except Exception as e:
                st.error(f"Error al procesar {mes_sel}: {e}")
                with st.expander("Detalle"):
                    st.code(traceback.format_exc())

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 2: Meses cargados
    # ─────────────────────────────────────────────────────────────────────────
    if not meses_cargados:
        return

    st.divider()
    st.subheader("Meses cargados en la sesión")

    # Mostrar tabla de resumen
    resumen_rows = []
    for mes in MESES_ORDER:
        if mes not in meses_cargados:
            continue
        entry = meses_cargados[mes]
        df    = entry["df"]
        n_ok  = len(df[df['Canal'] != 'Anulacion'])
        n_err = len(df[df['Canal'] == 'Anulacion'])
        n_ars = len(df[(df['Moneda'] == 'ARS') & (df['Canal'] != 'Anulacion')])
        n_usd = len(df[(df['Moneda'] == 'USD') & (df['Canal'] != 'Anulacion')])
        tiene_cable = (df['TipoUSD'] == 'CABLE').any()
        resumen_rows.append({
            "Mes":        mes,
            "Ops válidas": n_ok,
            "Anulaciones": n_err,
            "Ops ARS":     n_ars,
            "Ops USD":     n_usd,
            "Cable":      "✓" if tiene_cable else "—",
            "TC MEP":     f"${entry['tc_mep']:,.2f}",
            "TC CCL":     f"${entry['tc_ccl']:,.2f}" if entry['tc_ccl'] else "—",
        })

    import pandas as _pd
    st.dataframe(
        _pd.DataFrame(resumen_rows).set_index("Mes"),
        use_container_width=True,
    )

    # Botones de eliminación por mes
    cols = st.columns(min(len(meses_cargados), 6))
    for i, mes in enumerate([m for m in MESES_ORDER if m in meses_cargados]):
        with cols[i % len(cols)]:
            if st.button(f"✕ Quitar {mes}", key=f"tes_del_{mes}"):
                del st.session_state[_KEY][mes]
                st.rerun()

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 3: Generar reporte
    # ─────────────────────────────────────────────────────────────────────────
    st.divider()

    # Ordenar meses en el orden del año
    months_ordered = {m: meses_cargados[m] for m in MESES_ORDER if m in meses_cargados}
    n_meses        = len(months_ordered)
    meses_label    = ", ".join(months_ordered.keys())

    st.subheader(f"Generar reporte — {n_meses} mes(es): {meses_label}")

    if base_file:
        st.info(
            f"Modo **actualización**: se agregarán {n_meses} mes(es) al reporte base «{base_file.name}».",
            icon="📎",
        )
    else:
        st.info(
            f"Modo **nuevo reporte**: se generará desde cero con {n_meses} mes(es).",
            icon="📄",
        )

    if st.button("Generar reporte Excel", type="primary", use_container_width=True):
        with st.spinner("Generando Excel..."):
            try:
                if base_file:
                    # ── Patrón update: agregar mes a mes al reporte existente ──
                    from skills.tesoreria.logic import agregar_mes
                    xlsx_bytes = base_file.read()
                    for mes_name, entry in months_ordered.items():
                        with st.spinner(f"Agregando {mes_name}..."):
                            xlsx_bytes = agregar_mes(
                                existing_bytes = xlsx_bytes,
                                mes_name       = mes_name,
                                df_mes         = entry["df"],
                                tc_mep         = entry["tc_mep"],
                                tc_ccl         = entry["tc_ccl"],
                            )
                    fname = base_file.name  # conservar nombre del base
                else:
                    # ── Patrón nuevo: generar desde cero ─────────────────────
                    from skills.tesoreria.logic import generar_reporte
                    xlsx_bytes = generar_reporte(months_ordered)
                    fname = f"Reporte_Tesoreria_{datetime.now().strftime('%b%Y')}.xlsx"

                st.success(
                    f"✓ Reporte {'actualizado' if base_file else 'generado'} "
                    f"con {n_meses} mes(es): {meses_label}."
                )

                # Métricas globales
                total_ops = sum(
                    len(v['df'][v['df']['Canal'] != 'Anulacion'])
                    for v in months_ordered.values()
                )
                total_err = sum(
                    len(v['df'][v['df']['Canal'] == 'Anulacion'])
                    for v in months_ordered.values()
                )
                c1, c2, c3 = st.columns(3)
                c1.metric("Meses incluidos",    n_meses)
                c2.metric("Operaciones válidas", total_ops)
                c3.metric("Anulaciones excluidas", total_err)

                st.download_button(
                    label=f"⬇ Descargar {fname}",
                    data=xlsx_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                )
            except Exception as e:
                st.error(f"Error al generar el reporte: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())


# import local para el nombre del archivo de descarga
from datetime import datetime
