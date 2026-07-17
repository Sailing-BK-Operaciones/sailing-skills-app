"""
Control saldo USD Cable — UI Streamlit (modo diario).

El baseline y los CSV de estado (baseline_split.csv, estado_diario.csv,
asignaciones_byma_broker.csv, movimientos_backdated.csv) están bundleados en
skills/control_saldo_usd_cable/state/ — no hace falta subirlos cada día.
Se actualizan automáticamente en cada corrida.
"""
import streamlit as st
import traceback
from skills.shared_ui import shared_or_upload
from skills.control_saldo_usd_cable.logic import get_state_info, read_state_file


def render():
    st.title("Control saldo USD Cable")
    st.markdown(
        "Determina cuánto saldo en **USD Cable** tiene cada comitente disponible para "
        "operar en **BROKER** (mandato exterior) vs. **BYMA** (mercado local). "
        "Modo incremental sobre el **baseline bundled** — solo subir la SALUSD del día "
        "y los CCs de los comitentes que operaron."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Inputs del día:**
        - `SALUSD*.XLS` — Listado de Saldos Dólares (para el control automático vs. output).
        - `HD{CTTE}_*.XLS` — cuentas corrientes de los comitentes que operaron
          (multi-upload: podés subir varios a la vez o ir sumándolos).

        **Archivo de referencia:** `TABCOMPB.XLS` se toma de Archivos Compartidos.

        **Estado bundled** (no hace falta subirlo — se actualiza en cada corrida):
        - `baseline_split.csv` — saldo de partida por comitente.
        - `estado_diario.csv` — último saldo OK guardado.
        - `asignaciones_byma_broker.csv` — memoria de decisiones BROKER/BYMA para CU$S y VTUC.
        - `movimientos_backdated.csv` — comprobantes fecha ≤ corte incluidos como incremento.

        **Comprobantes `BYMA O BROKER`** (CU$S recibos, VTUC ventas cable): si aparecen
        movimientos sin decisión previa, la app los lista al final con un selector
        BROKER / BYMA por movimiento. Al procesar de nuevo con las decisiones, se guardan
        para el próximo día.

        > El estado bundled persiste dentro del contenedor de la app. Si Streamlit Cloud
        > reinicia el contenedor (redeploy, etc.), vuelve a la versión inicial del repo.
        > Los botones de descarga al final permiten backup periódico.
        """)

    _render_diario()


def _render_diario():
    # ── Estado bundled ────────────────────────────────────────────────────────
    info = get_state_info()
    with st.expander("Estado bundled actual", expanded=False):
        c1, c2 = st.columns(2)
        for col, (label, key) in zip(
            [c1, c2, c1, c2],
            [("Baseline",     "baseline"),
             ("Estado diario", "estado"),
             ("Asignaciones",  "asignaciones"),
             ("Backdated",     "backdated")],
        ):
            v = info.get(key)
            if v:
                col.markdown(f"**{label}** — {v['size']:,} B · última mod. {v['mtime']}")
            else:
                col.markdown(f"**{label}** — _(vacío / no existe aún)_")

    # ── Inputs del día ────────────────────────────────────────────────────────
    st.subheader("Archivos del día")
    salusd_file = st.file_uploader(
        "SALUSD*.XLS *",
        type=["xls", "xlsx"],
        key="csuc_d_salusd",
        help="Listado de Saldos Dólares — se usa para el control automático "
             "(Total output = SALUSD por comitente).",
    )
    cc_files = st.file_uploader(
        "CCs USD Cable — HD{CTTE}_*.XLS (uno o más)",
        type=["xls", "xlsx"],
        key="csuc_d_ccs",
        accept_multiple_files=True,
        help="El número tras 'HD' es el comitente. Se pueden seleccionar varios "
             "de una vez o agregar más después.",
    )

    st.subheader("Archivo de referencia (Archivos Compartidos)")
    tabcompb_file = shared_or_upload(
        "shared_tabcompb", "TABCOMPB.XLS", ["xls", "xlsx"], "csuc_d_tabcompb"
    )

    st.divider()

    faltan = []
    if not salusd_file:   faltan.append("SALUSD*.XLS")
    if not tabcompb_file: faltan.append("TABCOMPB.XLS")
    if faltan:
        st.caption(f"Falta para habilitar: {', '.join(faltan)}")

    # ── Ejecución (primer pase, sin decisiones) ───────────────────────────────
    if st.button("Procesar día", type="primary",
                 use_container_width=True, disabled=bool(faltan)):
        _run_diario(
            cc_files, tabcompb_file, salusd_file,
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
            "Elegí BROKER o BYMA para cada movimiento y apretá el botón para reprocesar. "
            "La decisión queda memorizada para el próximo día."
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
            type="primary",
            use_container_width=True,
        ):
            _run_diario(
                cc_files, tabcompb_file, salusd_file,
                decisiones_ui=decisiones_ui,
            )
            st.rerun()

    # ── Descargas ─────────────────────────────────────────────────────────────
    st.divider()
    st.subheader("Descargas")

    fecha = ui["fecha_proc"]
    xlsx_name = f"Saldos Cable por CC {fecha}.xlsx"

    st.download_button(
        label=f"⬇ {xlsx_name} — reporte del día",
        data=result["xlsx"],
        file_name=xlsx_name,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="csuc_d_dl_xlsx",
    )

    with st.expander("Backups del estado (opcional — persiste solo mientras corre el container)"):
        st.caption(
            "Estos CSV se guardan automáticamente en `skills/control_saldo_usd_cable/state/` "
            "dentro del container. Descargalos periódicamente como backup si querés "
            "poder restaurarlos ante un redeploy."
        )
        dc1, dc2, dc3 = st.columns(3)
        with dc1:
            st.download_button(
                label="estado_diario.csv",
                data=result["estado"],
                file_name="estado_diario.csv",
                mime="text/csv",
                use_container_width=True,
                key="csuc_d_dl_estado",
            )
        with dc2:
            st.download_button(
                label="asignaciones_byma_broker.csv",
                data=result["asignaciones"],
                file_name="asignaciones_byma_broker.csv",
                mime="text/csv",
                use_container_width=True,
                key="csuc_d_dl_asig",
            )
        with dc3:
            st.download_button(
                label="movimientos_backdated.csv",
                data=result["backdated"],
                file_name="movimientos_backdated.csv",
                mime="text/csv",
                use_container_width=True,
                key="csuc_d_dl_bd",
            )
        # Baseline: solo descarga (no debería cambiar en la diaria)
        st.download_button(
            label="baseline_split.csv (fixed — solo lectura desde el bundled)",
            data=read_state_file("baseline"),
            file_name="baseline_split.csv",
            mime="text/csv",
            use_container_width=True,
            key="csuc_d_dl_base",
        )


def _run_diario(cc_files, tabcompb_file, salusd_file, decisiones_ui):
    """Corre procesar_diario y guarda resultado en session_state."""
    with st.spinner("Procesando día..."):
        try:
            from skills.control_saldo_usd_cable.logic import procesar_diario
            xlsx, estado, asig, bd, ui = procesar_diario(
                cc_files       = cc_files or [],
                tabcompb_file  = tabcompb_file,
                salusd_file    = salusd_file,
                decisiones_ui  = decisiones_ui,
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
