import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload
from skills.reporte_operativo.logic import MESES_ORDER, TC_HISTORICOS


def render():
    st.title("Reporte Operativo — Boletos")
    st.markdown(
        "Procesa el archivo **CONTBOLE** (XLS de Gallo) del mes y actualiza el reporte "
        "Excel acumulativo anual de Operaciones con los datos del mes seleccionado: "
        "Panel de Control, Canal Digital vs Manual, Mercado y Segmento, Rankings TOP 20, "
        "Operaciones Diarias, y Detalle del mes."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Flujo de trabajo:**
        1. Subí el **Reporte existente** (el XLSX acumulativo del año actual).
        2. Verificá que **TABCOMPB.XLS** esté cargado en Archivos Compartidos (o subilo acá).
        3. Subí el **CONTBOLE** del mes a procesar (hoja `Control_de_Boletos`).
        4. Seleccioná el mes, el año y los tipos de cambio, luego clic en **Actualizar reporte**.
        5. Descargá el XLSX generado — reemplaza el anterior como nuevo acumulativo.

        **Primer mes del año:** arrancar desde el template Excel con las 5 hojas fijas vacías:
        - `Panel de Control`, `Canal Digital vs Manual`, `Mercado y Segmento`,
          `Rankings TOP 20`, `Operaciones Diarias`

        **Canal HB:** `WEB`, `APP`, `MgW` (Manager WEB)

        **Cuentas excluidas (Cartera Propia):** 1000, 1003, 1002, 1060, 2583, 1854

        **Tipos de cambio históricos pre-cargados como sugerencia:**
        | Mes | TC MEP | TC CCL |
        |-----|--------|--------|
        | Enero 2026 | 1.464,75 | 1.510,53 |
        | Febrero 2026 | 1.436,64 | 1.476,11 |
        | Marzo 2026 | 1.488,00 | 1.499,00 |
        | Abril 2026 | 1.448,93 | 1.503,03 |
        """)

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 1: Archivos
    # ─────────────────────────────────────────────────────────────────────────
    st.subheader("Archivos")

    reporte_file = st.file_uploader(
        "Reporte acumulativo (XLSX) *",
        type=["xlsx"],
        key="ro_reporte",
        help="El XLSX existente del año. El mes nuevo se agrega sobre este archivo.",
    )

    contbole_file = shared_or_upload(
        "shared_contbole", "CONTBOLE.XLS", ["xls", "xlsx"], "ro_contbole"
    )

    tabcompb_file = shared_or_upload(
        "shared_tabcompb", "TABCOMPB.XLS", ["xls", "xlsx"], "ro_tabcompb"
    )

    # ─────────────────────────────────────────────────────────────────────────
    # SECCIÓN 2: Parámetros
    # ─────────────────────────────────────────────────────────────────────────
    st.subheader("Parámetros del mes")

    col_mes, col_anio, col_tc1, col_tc2 = st.columns([2, 1, 1.5, 1.5])
    with col_mes:
        mes_sel = st.selectbox(
            "Mes a procesar",
            options=MESES_ORDER,
            key="ro_mes",
        )
    with col_anio:
        anio_sel = st.number_input(
            "Año",
            value=2026,
            min_value=2020, max_value=2099, step=1,
            key="ro_anio",
        )
    with col_tc1:
        tc_hist_mep = TC_HISTORICOS.get(mes_sel, {}).get('MEP', 0.0) if mes_sel else 0.0
        tc_mep = st.number_input(
            "TC MEP",
            value=tc_hist_mep,
            min_value=0.0, step=0.01, format="%.2f",
            key="ro_tc_mep",
            help="Tipo de cambio MEP del mes para pesificar operaciones en USD MEP",
        )
    with col_tc2:
        tc_hist_ccl = TC_HISTORICOS.get(mes_sel, {}).get('CCL', 0.0) if mes_sel else 0.0
        tc_ccl = st.number_input(
            "TC CCL",
            value=tc_hist_ccl,
            min_value=0.0, step=0.01, format="%.2f",
            key="ro_tc_ccl",
            help="Tipo de cambio CCL (Cable) del mes",
        )

    st.divider()

    # ─────────────────────────────────────────────────────────────────────────
    # Validación
    # ─────────────────────────────────────────────────────────────────────────
    faltantes = []
    if not reporte_file:  faltantes.append("Reporte acumulativo XLSX")
    if not contbole_file: faltantes.append("CONTBOLE XLS")
    if not tabcompb_file: faltantes.append("TABCOMPB.XLS")
    if tc_mep == 0:       faltantes.append("TC MEP")

    if faltantes:
        st.caption(f"Falta para habilitar: {', '.join(faltantes)}")

    # ─────────────────────────────────────────────────────────────────────────
    # Ejecución
    # ─────────────────────────────────────────────────────────────────────────
    if st.button(
        f"Actualizar reporte — {mes_sel} {int(anio_sel)}",
        type="primary",
        use_container_width=True,
        disabled=bool(faltantes),
    ):
        with st.spinner(f"Procesando {mes_sel} {int(anio_sel)}..."):
            try:
                from skills.reporte_operativo.logic import (
                    load_tabcompb, load_contbole, procesar, actualizar_reporte
                )

                tabcompb = load_tabcompb(tabcompb_file)
                rows     = load_contbole(contbole_file)
                datos, advertencias = procesar(
                    rows,
                    tabcompb,
                    tc_mep=tc_mep,
                    tc_ccl=tc_ccl if tc_ccl > 0 else 1.0,
                )

                reporte_file.seek(0)
                new_bytes = actualizar_reporte(
                    existing_bytes=reporte_file.read(),
                    datos=datos,
                    tabcompb=tabcompb,
                    mes_name=mes_sel,
                    tc_mep=tc_mep,
                    tc_ccl=tc_ccl if tc_ccl > 0 else 1.0,
                    anio=int(anio_sel),
                )

                fname = f"Reporte_Operativo_{mes_sel[:3]}{int(anio_sel)}.xlsx"

                st.success(
                    f"✓ {mes_sel} {int(anio_sel)} procesado — "
                    f"{datos['boletos_analizados']} boletos analizados, "
                    f"{datos['cuentas']} cuentas."
                )

                # ── Métricas ──────────────────────────────────────────────────
                c1, c2, c3, c4, c5, c6 = st.columns(6)
                c1.metric("Boletos totales",    datos['boletos_totales'])
                c2.metric("Boletos analizados", datos['boletos_analizados'])
                c3.metric("Anulados",           datos['anulados_count'])
                c4.metric("Cuentas clientes",   datos['cuentas'])
                c5.metric("Canal HB",           datos['hb'])
                c6.metric("Canal Manual",       datos['manual'])

                # ── Descarga ──────────────────────────────────────────────────
                st.download_button(
                    label=f"⬇ Descargar {fname}",
                    data=new_bytes,
                    file_name=fname,
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True,
                    key="ro_dl",
                )

                # ── Advertencias ──────────────────────────────────────────────
                if advertencias:
                    with st.expander(f"⚠ {len(advertencias)} advertencia(s) de procesamiento"):
                        for adv in advertencias:
                            st.markdown(f"- {adv}")

            except Exception as e:
                st.error(f"Error al procesar {mes_sel}: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())
