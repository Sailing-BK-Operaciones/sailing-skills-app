"""
Posicion Aforada por Comitente — UI Streamlit.

Dos alternativas al usuario:
  1) Reporte completo (una fila por comitente + detalle colapsable) — descarga.
  2) Consulta puntual para un comitente específico — se muestra inline.
"""
import streamlit as st
import traceback
from datetime import datetime
import pandas as pd
from skills.shared_ui import shared_or_upload


def render():
    st.title("Posicion aforada por ctte")
    st.markdown(
        "Difunde por comitente la **posición tomadora** (saldo deudor), la **capacidad "
        "de garantizar** con los títulos aceptados en garantía (valor aforado), y la "
        "**diferencia** entre ambos."
    )
    st.divider()

    with st.expander("¿Cómo usar esta skill?"):
        st.markdown("""
        **Input del día:**
        - `TENAFORADA.XLS` — hoja `Tenencia_Aforada`. Descargar diariamente de Gallo.

        **Archivo de referencia** (Archivos Compartidos):
        - `ESPECIES.XLS` — para resolver el ticker de cada especie (col `Norm.`).

        **Dos alternativas de consulta:**
        - **Reporte completo**: descarga un Excel con una fila por comitente
          (filtro estándar: saldo deudor > $100.000 y no cartera propia
          1000-1003). El detalle de las especies queda **agrupado y colapsado**
          por default; se despliega con el botón (+) al lado del comitente.
          Orden: descubiertos primero (rojo), luego cubiertos (verde).
        - **Consulta puntual**: elegís un comitente cualquiera (aún los que
          NO cumplen el filtro estándar) y ves acá mismo su saldo deudor,
          posición garantizable, diferencia y el detalle de las especies.
        """)

    # ── Inputs ────────────────────────────────────────────────────────────────
    st.subheader("Archivos")
    tenaforada_file = st.file_uploader(
        "TENAFORADA.XLS *",
        type=["xls", "xlsx"],
        key="pa_tenaforada",
        help="Tenencia aforada + saldo deudor por comitente (hoja Tenencia_Aforada).",
    )
    especies_file = shared_or_upload(
        "shared_especies", "ESPECIES.XLS", ["xls", "xlsx"], "pa_especies"
    )

    st.divider()

    faltan = []
    if not tenaforada_file: faltan.append("TENAFORADA.XLS")
    if not especies_file:   faltan.append("ESPECIES.XLS")
    if faltan:
        st.caption(f"Falta para habilitar: {', '.join(faltan)}")
        return

    # ── Parseo (una vez, cacheado en session_state por tamaño de input) ───────
    if st.button("Procesar TENAFORADA", type="primary", use_container_width=True):
        with st.spinner("Parseando TENAFORADA..."):
            try:
                from skills.posicion_aforada.logic import (
                    parse_all, generar_reporte_bytes, filtrar_seleccionados,
                )
                comitentes, tickers, fecha = parse_all(
                    tenaforada_file, especies_file, fecha_input=datetime.now(),
                )
                xlsx_bytes = generar_reporte_bytes(comitentes, tickers, fecha)
                sel = filtrar_seleccionados(comitentes)
                st.session_state["pa_result"] = {
                    "comitentes": comitentes,
                    "tickers":    tickers,
                    "fecha":      fecha,
                    "xlsx":       xlsx_bytes,
                    "sel":        sel,
                }
            except Exception as e:
                st.error(f"Error al procesar: {e}")
                with st.expander("Detalle del error"):
                    st.code(traceback.format_exc())

    result = st.session_state.get("pa_result")
    if not result:
        return

    comitentes = result["comitentes"]
    tickers    = result["tickers"]
    fecha      = result["fecha"]
    sel        = result["sel"]

    # ── Métricas + descarga del reporte completo ──────────────────────────────
    n_desc = sum(1 for c in sel if c["diferencia"] < 0)
    n_cub  = len(sel) - n_desc
    tot_saldo = sum(c["saldo"] for c in sel)
    tot_afor  = sum(c["aforado_total"] for c in sel)
    tot_dif   = sum(c["diferencia"] for c in sel)

    st.success(
        f"✓ Procesado — TENAFORADA cargada, {len(comitentes)} comitentes parseados. "
        f"Fecha proceso: {fecha:%d/%m/%Y %H:%M}"
    )

    c1, c2, c3, c4 = st.columns(4)
    c1.metric("Selección (>$100k)", len(sel))
    c2.metric("Cubiertos",  n_cub)
    c3.metric("Descubiertos", n_desc)
    c4.metric("Diferencia total", f"{tot_dif:,.0f}")

    st.subheader("Reporte completo (para difundir al equipo comercial)")
    fname = f"Posicion Aforada por Ctte {fecha:%d-%m-%Y}.xlsx"
    st.download_button(
        label=f"⬇ Descargar {fname}",
        data=result["xlsx"],
        file_name=fname,
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True,
        key="pa_dl",
    )
    st.caption(
        f"Filtro estándar: saldo deudor > $100.000 y no cartera propia (1000-1003). "
        f"Totales — Saldo: `{tot_saldo:,.2f}` · Garantizable: `{tot_afor:,.2f}` · "
        f"Diferencia: `{tot_dif:,.2f}`."
    )

    # ── Lista rápida de descubiertos (si hay) ─────────────────────────────────
    if n_desc:
        with st.expander(f"🔴 Ver comitentes DESCUBIERTOS ({n_desc})", expanded=False):
            for c in sel:
                if c["diferencia"] >= 0:
                    break   # sel viene ordenado ascendente → descubiertos primero
                st.markdown(
                    f"- **{c['ctte']}** · {c['nombre']} · "
                    f"saldo `{c['saldo']:,.2f}` · garantizable `{c['aforado_total']:,.2f}` · "
                    f"**dif `{c['diferencia']:,.2f}`**"
                )

    st.divider()

    # ── Consulta puntual por comitente ────────────────────────────────────────
    st.subheader("Consulta por comitente")
    st.caption(
        "Elegí un comitente para ver su detalle (sin filtro — funciona para cualquier "
        "comitente del TENAFORADA, no solo los del reporte principal)."
    )

    # Opciones: ordenar por número, mostrar CTTE + nombre
    opts = sorted(comitentes, key=lambda x: x["ctte"])
    option_labels = {c["ctte"]: f"{c['ctte']} — {c['nombre']}" for c in opts}
    option_values = [None] + [c["ctte"] for c in opts]

    ctte_sel = st.selectbox(
        "Comitente",
        options=option_values,
        format_func=lambda v: "— seleccionar —" if v is None else option_labels.get(v, str(v)),
        key="pa_ctte_sel",
    )

    if ctte_sel is None:
        return

    from skills.posicion_aforada.logic import consultar_comitente
    c = consultar_comitente(comitentes, ctte_sel)
    if c is None:
        st.warning("Comitente no encontrado.")
        return

    # ── Panel del comitente ───────────────────────────────────────────────────
    st.markdown(f"### {c['ctte']} — {c['nombre']}")

    diff = c["diferencia"]
    color_msg = "🟢 CUBIERTO" if diff >= 0 else "🔴 DESCUBIERTO"

    m1, m2, m3, m4 = st.columns(4)
    m1.metric("Saldo Deudor",          f"{c['saldo']:,.2f}")
    m2.metric("Posición Garantizable", f"{c['aforado_total']:,.2f}")
    m3.metric("Diferencia",            f"{diff:,.2f}",
              delta=color_msg, delta_color=("normal" if diff >= 0 else "inverse"))
    m4.metric("Especies en garantía",  len(c["especies"]))

    if diff >= 0:
        st.success(
            f"✓ Los títulos en garantía **cubren** el saldo deudor "
            f"(excedente `{diff:,.2f}`).",
            icon="🟢",
        )
    else:
        st.error(
            f"⚠ Los títulos en garantía **NO alcanzan** — descubierto por "
            f"`{abs(diff):,.2f}`.",
            icon="🔴",
        )

    # ── Detalle de especies ───────────────────────────────────────────────────
    if not c["especies"]:
        st.info("Este comitente no tiene tenencias aceptadas en garantía.")
        return

    st.markdown("**Detalle de especies en garantía**")
    rows = []
    for e in c["especies"]:
        ticker = tickers.get(str(e["cod"]).zfill(5), "")
        rows.append({
            "Ticker":         ticker,
            "Código CVSA":    e["cod"],
            "Especie":        e["especie"],
            "Tenencia":       e["tenencia"],
            "Precio":         e["precio"],
            "Aforo %":        e["aforo_pct"],
            "Valor Aforado":  e["valor_aforado"],
        })
    df = pd.DataFrame(rows)
    st.dataframe(
        df,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Tenencia":      st.column_config.NumberColumn(format="%.2f"),
            "Precio":        st.column_config.NumberColumn(format="%.4f"),
            "Aforo %":       st.column_config.NumberColumn(format="%.0f"),
            "Valor Aforado": st.column_config.NumberColumn(format="%.2f"),
        },
    )

    # Chequeo: suma del valor aforado del detalle vs aforado_total del subtotal
    suma_det = sum(e["valor_aforado"] for e in c["especies"])
    if abs(suma_det - c["aforado_total"]) > 0.05:
        st.caption(
            f"⚠ Diferencia interna: suma del detalle `{suma_det:,.2f}` vs "
            f"total subtotal `{c['aforado_total']:,.2f}` (diff `{suma_det - c['aforado_total']:,.2f}`)."
        )
