"""
Control saldo USD Cable — UI Streamlit.

Dos modos:
  - BASELINE : saldo de partida (una vez al inicio de actividad).
  - DIARIO   : incremental sobre el baseline + CCs + control SALUSD.

El estado (baseline_split.csv, estado_diario.csv, asignaciones_byma_broker.csv,
movimientos_backdated.csv) se maneja via uploads/downloads porque Streamlit
Cloud no persiste filesystem entre sesiones.
"""
import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload


def render():
    st.title("Control saldo USD Cable")
    st.markdown(
        "Determina cuánto saldo en **USD Cable** tiene cada comitente disponible para "
        "operar en **BROKER** (mandato exterior) vs. **BYMA** (mercado local). "
        "Convención positiva (cash disponible)."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Dos modos:**

        - **BASELINE** — se corre **una vez** al inicio de actividad para fijar el saldo de
          partida por comitente. Inputs: `SALUSD*.XLS` + `Saldos y Tenencias*.xlsx`
          (bloque LISTADO GALLO). Genera `baseline_split.csv` (guardar y reusar en modo diario).

        - **DIARIO** — se corre día a día. Inputs: `baseline_split.csv` (del baseline o del
          día anterior), los `HD{CTTE}_*.XLS` del día, `SALUSD*.XLS` (control), y opcionalmente
          los CSV de estado / asignaciones / back-dated del día anterior.
          `TABCOMPB.XLS` se toma automáticamente de Archivos Compartidos.

        **Estado editable (CSV, se descargan al final):**
        - `baseline_split.csv` — saldo de partida.
        - `estado_diario.csv` — último saldo OK por comitente (para el control automático).
        - `asignaciones_byma_broker.csv` — memoria BROKER/BYMA para CU$S y VTUC.
        - `movimientos_backdated.csv` — comprobantes fecha ≤ corte incluidos como incremento.

        **Comprobantes `BYMA O BROKER`** (CU$S, VTUC): el destino lo decide el cliente.
        La app los muestra al final; se puede reprocesar cargando la asignación via
        `asignaciones_byma_broker.csv` editado a mano.
        """)

    modo = st.radio(
        "Modo",
        options=["Diario", "Baseline"],
        index=0,
        horizontal=True,
        key="csuc_modo",
    )
    st.divider()

    if modo == "Baseline":
        _render_baseline()
    else:
        _render_diario()


# ── Modo BASELINE ─────────────────────────────────────────────────────────────
def _render_baseline():
    st.subheader("Modo BASELINE — saldo de partida")
    st.caption(
        "Fija el split BROKER/BYMA a partir de SALUSD + el bloque LISTADO GALLO "
        "de la planilla de tenencias. Correr solo cuando arrancás el ciclo."
    )

    salusd_file = st.file_uploader(
        "SALUSD*.XLS *",
        type=["xls", "xlsx"],
        key="csuc_b_salusd",
        help="Listado de Saldos Dólares (hoja Listado_Saldos_Dolares).",
    )
    listado_file = st.file_uploader(
        "Saldos y Tenencias*.xlsx *",
        type=["xlsx"],
        key="csuc_b_listado",
        help="Hoja 'Saldos CTTES' con el bloque LISTADO GALLO (columnas CTTE / CASH BROKERS / BYMA-VALO).",
    )

    faltan = []
    if not salusd_file:  faltan.append("SALUSD*.XLS")
    if not listado_file: faltan.append("Saldos y Tenencias*.xlsx")
    if faltan:
        st.caption(f"Falta para habilitar: {', '.join(faltan)}")

    if st.button("Generar baseline", type="primary",
                 use_container_width=True, disabled=bool(faltan)):
        with st.spinner("Procesando baseline..."):
            try:
                from skills.control_saldo_usd_cable.logic import procesar_baseline
                xlsx, csv_bytes, ui = procesar_baseline(salusd_file, listado_file)
                st.session_state["csuc_b_result"] = {
                    "xlsx":     xlsx,
                    "csv":      csv_bytes,
                    "ui":       ui,
                }
            except Exception as e:
                st.error(f"Error: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    result = st.session_state.get("csuc_b_result")
    if not result:
        return

    ui = result["ui"]
    st.success(f"✓ Baseline generado — corte {ui['fecha_corte']}")

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Comitentes",     ui["n_baseline"])
    c2.metric("Reconcilian OK", ui["n_ok"])
    c3.metric("REVISAR",        ui["n_revisar"])
    c4.metric("Falta split",    ui["n_falta"])

    if ui["falta_definir"]:
        st.error(
            "Comitentes en SALUSD sin split en LISTADO GALLO — hay que definir BROKER/BYMA:",
            icon="🔴",
        )
        for f in ui["falta_definir"]:
            st.markdown(
                f"- **{f['ctte']}** · {f['nombre']} · "
                f"saldo vencido `{f['saldo']:,.2f}`"
            )
    if ui["excluidas"]:
        with st.expander(f"Excluidas mercado/cartera propia ({ui['n_excluidas']})"):
            for e in ui["excluidas"]:
                st.markdown(f"- {e['ctte']} · {e['nombre']} · `{e['saldo']:,.2f}`")

    d1, d2 = st.columns(2)
    with d1:
        st.download_button(
            label=f"⬇ Descargar Saldos Cable - Baseline {ui['fecha_corte']}.xlsx",
            data=result["xlsx"],
            file_name=f"Saldos Cable - Baseline {ui['fecha_corte']}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="csuc_b_dl_xlsx",
        )
    with d2:
        st.download_button(
            label="⬇ Descargar baseline_split.csv (input del modo Diario)",
            data=result["csv"],
            file_name="baseline_split.csv",
            mime="text/csv",
            use_container_width=True,
            key="csuc_b_dl_csv",
        )


# ── Modo DIARIO ───────────────────────────────────────────────────────────────
def _render_diario():
    st.subheader("Modo DIARIO — incremental sobre el baseline")

    st.markdown("**Estado (obligatorio):**")
    baseline_csv = st.file_uploader(
        "baseline_split.csv *",
        type=["csv"],
        key="csuc_d_baseline",
        help="CSV generado por el modo BASELINE (o el estado del día anterior).",
    )

    st.markdown("**Del día:**")
    cc_files = st.file_uploader(
        "CCs USD Cable — HD{CTTE}_*.XLS (uno o más)",
        type=["xls", "xlsx"],
        key="csuc_d_ccs",
        accept_multiple_files=True,
        help="El número tras 'HD' es el comitente. Subí solo los CCs de las cuentas que operaron hoy.",
    )
    salusd_file = st.file_uploader(
        "SALUSD*.XLS (control) — opcional pero recomendado",
        type=["xls", "xlsx"],
        key="csuc_d_salusd",
    )

    st.markdown("**Estado del día anterior (opcional):**")
    col1, col2, col3 = st.columns(3)
    with col1:
        estado_csv = st.file_uploader(
            "estado_diario.csv", type=["csv"], key="csuc_d_estado",
        )
    with col2:
        asig_csv = st.file_uploader(
            "asignaciones_byma_broker.csv", type=["csv"], key="csuc_d_asig",
        )
    with col3:
        bd_csv = st.file_uploader(
            "movimientos_backdated.csv", type=["csv"], key="csuc_d_bd",
        )

    st.markdown("**Archivo de referencia (Archivos Compartidos):**")
    tabcompb_file = shared_or_upload(
        "shared_tabcompb", "TABCOMPB.XLS", ["xls", "xlsx"], "csuc_d_tabcompb"
    )

    st.divider()

    faltan = []
    if not baseline_csv:  faltan.append("baseline_split.csv")
    if not tabcompb_file: faltan.append("TABCOMPB.XLS")
    if faltan:
        st.caption(f"Falta para habilitar: {', '.join(faltan)}")

    # Ejecución primer pase (sin decisiones)
    if st.button("Procesar día", type="primary",
                 use_container_width=True, disabled=bool(faltan)):
        _run_diario(
            baseline_csv, cc_files, tabcompb_file,
            salusd_file, estado_csv, asig_csv, bd_csv,
            decisiones_ui=None,
        )

    result = st.session_state.get("csuc_d_result")
    if not result:
        return

    ui = result["ui"]

    # ── Métricas ──────────────────────────────────────────────────────────────
    if ui["n_falta_cc"] == 0 and ui["n_discrep"] == 0:
        st.success(
            f"✓ Todo cuadra — {ui['n_ok']}/{ui['n_panel']} OK "
            f"(fecha {ui['fecha_proc']}, corte baseline {ui['cutoff']})"
        )
    else:
        st.warning(
            f"Panel {ui['fecha_proc']} — {ui['n_ok']}/{ui['n_panel']} OK · "
            f"{ui['n_falta_cc']} falta CC · {ui['n_discrep']} discrepancia",
            icon="⚠",
        )

    c1, c2, c3, c4, c5 = st.columns(5)
    c1.metric("Panel",         ui["n_panel"])
    c2.metric("Procesados",    ui["n_procesados"])
    c3.metric("OK",            ui["n_ok"])
    c4.metric("Falta CC",      ui["n_falta_cc"])
    c5.metric("Discrepancias", ui["n_discrep"])

    # ── Alertas ───────────────────────────────────────────────────────────────
    if ui["falta_cc"]:
        with st.expander(f"🔴 OPERO - FALTA CC ({ui['n_falta_cc']})", expanded=True):
            for f in ui["falta_cc"]:
                st.markdown(
                    f"- **{f['ctte']}** · {f['nombre']} · "
                    f"guardado `{f['guardado']:,.2f}` vs SALUSD `{f['salusd']:,.2f}` "
                    f"(cambio `{f['diff']:,.2f}`)"
                )
            st.caption("Subir el CC (HD####) de estas cuentas y volver a procesar.")
    if ui["discrepancias"]:
        with st.expander(f"⚠ DISCREPANCIAS ({ui['n_discrep']})", expanded=True):
            for f in ui["discrepancias"]:
                st.markdown(
                    f"- **{f['ctte']}** · {f['nombre']} · "
                    f"output `{f['output']:,.2f}` vs SALUSD `{f['salusd']:,.2f}` "
                    f"(dif `{f['diff']:,.2f}`)"
                )
    if ui["faltantes_tabcompb"]:
        with st.expander(f"⚠ Comprobantes SIN destino en TABCOMPB ({len(ui['faltantes_tabcompb'])})",
                         expanded=True):
            for f in ui["faltantes_tabcompb"]:
                st.markdown(
                    f"- `{f['cpbt']}` → comitentes {f['comitentes']}"
                )
            st.caption("Actualizá la columna 'analisis CC' de TABCOMPB.XLS y reprocesá.")
    if ui["backdated"]:
        with st.expander(f"🟩 Movimientos back-dated detectados ({len(ui['backdated'])})"):
            for b in ui["backdated"]:
                fs = b["fecha"].strftime("%d/%m/%Y") if hasattr(b["fecha"], "strftime") else str(b["fecha"])
                st.markdown(
                    f"- CTTE **{b['ctte']}** · {fs} · `{b['cpbt']} N.{b['numero']}` · "
                    f"`{b['importe']:,.2f}` (incluido y recordado)"
                )

    # ── Pendientes BYMA O BROKER (asignación interactiva) ─────────────────────
    if ui["pendientes"]:
        st.divider()
        st.subheader(f"Movimientos 'BYMA O BROKER' pendientes ({len(ui['pendientes'])})")
        st.caption(
            "Elegí el destino de cada uno y volvé a procesar. La decisión queda "
            "guardada en `asignaciones_byma_broker.csv` para el próximo día."
        )
        decisiones_ui = {}
        for p in ui["pendientes"]:
            fs = p["fecha"].strftime("%d/%m/%Y") if hasattr(p["fecha"], "strftime") else str(p["fecha"])
            wkey = f"csuc_p_{p['key'][0]}_{p['key'][1]}_{p['key'][2]}"
            cols = st.columns([4, 1])
            with cols[0]:
                st.markdown(
                    f"- CTTE **{p['ctte']}** · {fs} · `{p['cpbt']} N.{p['numero']}` · "
                    f"`{p['importe']:,.2f}` · _{p['referencia'][:60]}_"
                )
            with cols[1]:
                sel = st.selectbox(
                    "Destino",
                    options=["—", "BROKER", "BYMA"],
                    key=wkey,
                    label_visibility="collapsed",
                )
                if sel in ("BROKER", "BYMA"):
                    decisiones_ui[p["key"]] = sel

        if decisiones_ui and st.button(
            f"Reprocesar con {len(decisiones_ui)} decisión(es)",
            key="csuc_d_reprocess",
            use_container_width=True,
        ):
            _run_diario(
                baseline_csv, cc_files, tabcompb_file,
                salusd_file, estado_csv, asig_csv, bd_csv,
                decisiones_ui=decisiones_ui,
            )
            st.rerun()

    # ── Descargas ─────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Descargas")

    fecha = ui["fecha_proc"]
    xlsx_name = f"Saldos Cable por CC {fecha}.xlsx"

    dc1, dc2 = st.columns(2)
    with dc1:
        st.download_button(
            label=f"⬇ {xlsx_name}",
            data=result["xlsx"],
            file_name=xlsx_name,
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True,
            key="csuc_d_dl_xlsx",
        )
    with dc2:
        st.download_button(
            label="⬇ estado_diario.csv (guardar para mañana)",
            data=result["estado"],
            file_name="estado_diario.csv",
            mime="text/csv",
            use_container_width=True,
            key="csuc_d_dl_estado",
        )

    dc3, dc4 = st.columns(2)
    with dc3:
        st.download_button(
            label="⬇ asignaciones_byma_broker.csv",
            data=result["asignaciones"],
            file_name="asignaciones_byma_broker.csv",
            mime="text/csv",
            use_container_width=True,
            key="csuc_d_dl_asig",
        )
    with dc4:
        st.download_button(
            label="⬇ movimientos_backdated.csv",
            data=result["backdated"],
            file_name="movimientos_backdated.csv",
            mime="text/csv",
            use_container_width=True,
            key="csuc_d_dl_bd",
        )

    st.caption(
        "Guardá los 3 CSV para subirlos como estado inicial mañana; "
        "el XLSX es el reporte del día."
    )


def _run_diario(baseline_csv, cc_files, tabcompb_file,
                salusd_file, estado_csv, asig_csv, bd_csv, decisiones_ui):
    """Corre procesar_diario y guarda resultado en session_state."""
    with st.spinner("Procesando día..."):
        try:
            from skills.control_saldo_usd_cable.logic import procesar_diario
            xlsx, estado, asig, bd, ui = procesar_diario(
                baseline_csv_file      = baseline_csv,
                cc_files               = cc_files or [],
                tabcompb_file          = tabcompb_file,
                salusd_file            = salusd_file,
                estado_csv_file        = estado_csv,
                asignaciones_csv_file  = asig_csv,
                backdated_csv_file     = bd_csv,
                decisiones_ui          = decisiones_ui,
            )
            st.session_state["csuc_d_result"] = {
                "xlsx":         xlsx,
                "estado":       estado,
                "asignaciones": asig,
                "backdated":    bd,
                "ui":           ui,
            }
        except Exception as e:
            st.error(f"Error: {e}")
            with st.expander("Detalle del error"):
                st.code(traceback.format_exc())
