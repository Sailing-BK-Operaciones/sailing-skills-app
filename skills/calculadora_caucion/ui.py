"""
Calculadora de Cauciones — UI Streamlit.

Embebe el HTML autocontenido (mismo motor que el Artifact) dentro del panel.
No procesa archivos: la calculadora corre 100% en el navegador del usuario.
"""
from pathlib import Path
import streamlit as st
import streamlit.components.v1 as components


HTML_PATH = Path(__file__).with_name("calculadora_caucion.html")


def render():
    st.title("Calculadora de Cauciones")
    st.markdown(
        "Estimá una caución **Tomadora** (tomar fondos) o **Colocadora** "
        "(colocar fondos) con interés, gastos, impuestos y aranceles "
        "discriminados, y la tasa neta resultante de cada punta."
    )

    with st.expander("¿Cómo usar la calculadora?"):
        st.markdown("""
        - Editable por el comercial: **monto, tasa y días** (los tres campos en
          ámbar). El resto son parámetros de mercado que actualiza BackOffice
          en el código.
        - Fórmula: `interés = monto × tasa / 365 × días` (base Actual/365).
        - **Tomadora**: costo total = interés + arancel + derecho de registro
          + derechos de mercado + IVA.
        - **Colocadora**: resultado = interés − arancel − derechos − IVA
          − imp. Ganancias (hoy 0%).
        - El cálculo es agnóstico a la moneda; el selector ARS / USD sólo cambia
          el prefijo mostrado.
        """)

    st.divider()

    try:
        html = HTML_PATH.read_text(encoding="utf-8")
    except FileNotFoundError:
        st.error(f"No se encontró el archivo HTML de la calculadora: {HTML_PATH}")
        return

    components.html(html, height=1100, scrolling=True)

    st.caption(
        "Los parámetros de mercado (aranceles, derechos, IVA, prorrateos, "
        "bonificación) están fijados en el HTML. Para actualizarlos, editar el "
        "objeto `K` (`K.tom` / `K.col`) en `calculadora_caucion.html` y "
        "redeployar."
    )
