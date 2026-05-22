import streamlit as st

st.set_page_config(
    page_title="Sailing Inversiones — Skills",
    page_icon="📊",
    layout="wide"
)

# ── Estilos globales ──────────────────────────────────────────────────────────
st.markdown("""
<style>
/* ── Fuente global ── */
html, body, [class*="css"] {
    font-family: 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
}

/* ── Sidebar oscuro estilo plataforma financiera ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1b2a 0%, #142236 100%);
    border-right: 1px solid #1e3a5f;
}
section[data-testid="stSidebar"] * {
    color: #dce8f5 !important;
}
section[data-testid="stSidebar"] h1 {
    color: #ffffff !important;
    font-size: 1.1rem !important;
    letter-spacing: 0.03em;
    padding-bottom: 0.3rem;
    border-bottom: 1px solid #2a4a6e;
    margin-bottom: 0.8rem !important;
}
section[data-testid="stSidebar"] label {
    color: #b8cfe8 !important;
    font-size: 0.88rem !important;
}
/* Radio seleccionado */
section[data-testid="stSidebar"] [data-baseweb="radio"] [aria-checked="true"] ~ div {
    color: #4fc3f7 !important;
    font-weight: 600 !important;
}
/* Alerts en sidebar */
section[data-testid="stSidebar"] [data-testid="stAlert"] {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 6px;
}

/* ── Área principal ── */
.main .block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
    max-width: 1200px;
}

/* ── Títulos ── */
h1 {
    color: #0d1b2a !important;
    font-weight: 700 !important;
    font-size: 1.6rem !important;
    border-bottom: 2px solid #1565c0;
    padding-bottom: 0.4rem;
    margin-bottom: 0.6rem !important;
}
h2 {
    color: #1a3a5c !important;
    font-weight: 600 !important;
    font-size: 1.15rem !important;
}
h3 {
    color: #1565c0 !important;
    font-weight: 600 !important;
    font-size: 1rem !important;
}

/* ── Divider ── */
hr {
    border-color: #d0dae6 !important;
    margin: 1rem 0 !important;
}

/* ── Métricas ── */
[data-testid="stMetric"] {
    background: #f0f5fb;
    border-radius: 8px;
    padding: 0.8rem 1rem !important;
    border-left: 3px solid #1565c0;
}
[data-testid="stMetricLabel"] {
    color: #546e7a !important;
    font-size: 0.78rem !important;
    font-weight: 600 !important;
    text-transform: uppercase;
    letter-spacing: 0.04em;
}
[data-testid="stMetricValue"] {
    color: #0d1b2a !important;
    font-size: 1.35rem !important;
    font-weight: 700 !important;
}

/* ── Botones primarios ── */
button[kind="primary"] {
    background-color: #1565c0 !important;
    border-color: #1565c0 !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em;
    border-radius: 6px !important;
}
button[kind="primary"]:hover {
    background-color: #0d47a1 !important;
    border-color: #0d47a1 !important;
}

/* ── Botones de descarga ── */
button[kind="secondary"] {
    border-radius: 6px !important;
    border-color: #1565c0 !important;
    color: #1565c0 !important;
}
button[kind="secondary"]:hover {
    background-color: #e3f0fc !important;
}

/* ── Alerts ── */
[data-testid="stAlert"] {
    border-radius: 6px !important;
}

/* ── Expanders ── */
[data-testid="stExpander"] {
    border: 1px solid #d0dae6 !important;
    border-radius: 8px !important;
}
[data-testid="stExpander"] summary {
    font-weight: 600;
    color: #1a3a5c;
}

/* ── File uploader ── */
[data-testid="stFileUploader"] {
    border-radius: 8px !important;
}
[data-testid="stFileUploader"] section {
    border-color: #90b4d8 !important;
    border-radius: 8px !important;
    background: #f7fbff !important;
}
[data-testid="stFileUploader"] section:hover {
    border-color: #1565c0 !important;
    background: #eef5fd !important;
}

/* ── Number input ── */
[data-testid="stNumberInput"] input {
    border-radius: 6px !important;
}

/* ── Date input ── */
[data-testid="stDateInput"] input {
    border-radius: 6px !important;
}

/* ── Captions ── */
[data-testid="stCaptionContainer"] p {
    color: #607d8b !important;
    font-size: 0.8rem !important;
}
</style>
""", unsafe_allow_html=True)


# ── Autenticación ─────────────────────────────────────────────────────────────
MAX_INTENTOS = 5

def check_password():
    if "authenticated" not in st.session_state:
        st.session_state.authenticated = False
    if "failed_attempts" not in st.session_state:
        st.session_state.failed_attempts = 0

    if not st.session_state.authenticated:
        st.title("Sailing Inversiones")
        st.subheader("Acceso al panel de skills operativas")

        if st.session_state.failed_attempts >= MAX_INTENTOS:
            st.error(
                f"Acceso bloqueado por {MAX_INTENTOS} intentos fallidos. "
                "Cerrá esta pestaña y volvé a abrir el link para reintentar."
            )
            return False

        restantes = MAX_INTENTOS - st.session_state.failed_attempts
        password = st.text_input("Contraseña:", type="password")
        if st.button("Ingresar"):
            if password == st.secrets["PASSWORD"]:
                st.session_state.authenticated = True
                st.session_state.failed_attempts = 0
                st.rerun()
            else:
                st.session_state.failed_attempts += 1
                restantes -= 1
                if restantes > 0:
                    st.error(f"Contraseña incorrecta. Intentos restantes: {restantes}")
                else:
                    st.rerun()
        return False
    return True

if not check_password():
    st.stop()


# ── Sidebar ───────────────────────────────────────────────────────────────────
st.sidebar.title("Sailing Inversiones")

skill = st.sidebar.radio(
    "Seleccioná la skill:",
    [
        "Archivos Compartidos",
        "Risk Monitoring Client",
        "Risk Position",
        "Distribución Gara BYMA",
        "Genera TXT Gara NASDAQ",
        "Actual Position",
        "Settlement Position",
        "Collateral Position",
    ]
)

# Indicador de archivos compartidos
from shared_store import is_loaded as _is_loaded
_SHARED_KEYS = [
    "shared_pc", "shared_sagaclte", "shared_sateclte", "shared_contbole",
    "shared_saldos_nasdaq", "shared_especies", "shared_accounts",
    "shared_pdf_aforos", "shared_aforo_sail", "shared_tabcompb",
]
_n = sum(1 for k in _SHARED_KEYS if _is_loaded(k))
_total = len(_SHARED_KEYS)
if _n == _total:
    st.sidebar.success(f"✓ {_n}/{_total} archivos compartidos")
elif _n > 0:
    st.sidebar.warning(f"⚠ {_n}/{_total} archivos compartidos")
else:
    st.sidebar.info("Sin archivos compartidos")


# ── Routing ───────────────────────────────────────────────────────────────────
if skill == "Archivos Compartidos":
    from shared_inputs import render
    render()
elif skill == "Risk Monitoring Client":
    from skills.risk_monitoring_client.ui import render
    render()
elif skill == "Risk Position":
    from skills.risk_position.ui import render
    render()
elif skill == "Distribución Gara BYMA":
    from skills.distribucion_gara.ui import render
    render()
elif skill == "Genera TXT Gara NASDAQ":
    from skills.genera_txt_gara_nasdaq.ui import render
    render()
elif skill == "Actual Position":
    from skills.actual_position.ui import render
    render()
elif skill == "Collateral Position":
    from skills.Collateral_position.ui import render
    render()
else:
    st.info("Skill en construcción.")
