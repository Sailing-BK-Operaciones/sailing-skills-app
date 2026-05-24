import streamlit as st

st.set_page_config(
    page_title="Sailing Inversiones — Skills",
    page_icon="📊",
    layout="wide"
)

# ── Modo Day/Dark (persiste en session_state) ─────────────────────────────────
if "dark_mode" not in st.session_state:
    st.session_state.dark_mode = False

dark = st.session_state.dark_mode

# ── CSS base (siempre aplicado) ───────────────────────────────────────────────
st.markdown("""
<style>
html, body, [class*="css"] {
    font-family: 'Segoe UI', system-ui, -apple-system, BlinkMacSystemFont, sans-serif;
}

/* ── Sidebar — siempre oscuro ── */
section[data-testid="stSidebar"] {
    background: linear-gradient(180deg, #0d1b2a 0%, #142236 100%);
    border-right: 1px solid #1e3a5f;
}
section[data-testid="stSidebar"] * { color: #dce8f5 !important; }
section[data-testid="stSidebar"] h1 {
    color: #ffffff !important;
    font-size: 1.1rem !important;
    letter-spacing: 0.03em;
    padding-bottom: 0.3rem;
    border-bottom: 1px solid #2a4a6e;
    margin-bottom: 0.8rem !important;
}
section[data-testid="stSidebar"] label { color: #b8cfe8 !important; font-size: 0.88rem !important; }
section[data-testid="stSidebar"] [data-baseweb="radio"] [aria-checked="true"] ~ div {
    color: #4fc3f7 !important; font-weight: 600 !important;
}
section[data-testid="stSidebar"] [data-testid="stAlert"] {
    background: rgba(255,255,255,0.06) !important;
    border: 1px solid rgba(255,255,255,0.12) !important;
    border-radius: 6px;
}

/* ── Layout ── */
.main .block-container { padding-top: 1.2rem; padding-bottom: 2rem; max-width: 1200px; }

/* ── Botones primarios ── */
button[kind="primary"] {
    background-color: #1565c0 !important; border-color: #1565c0 !important;
    font-weight: 600 !important; letter-spacing: 0.02em; border-radius: 6px !important;
}
button[kind="primary"]:hover { background-color: #0d47a1 !important; border-color: #0d47a1 !important; }

/* ── Alerts ── */
[data-testid="stAlert"] { border-radius: 6px !important; }
</style>
""", unsafe_allow_html=True)

# ── CSS condicional Day / Dark ────────────────────────────────────────────────
if not dark:
    st.markdown("""
<style>
h1 { color: #0d1b2a !important; font-weight: 700 !important; font-size: 1.6rem !important;
     border-bottom: 2px solid #1565c0; padding-bottom: 0.4rem; margin-bottom: 0.6rem !important; }
h2 { color: #1a3a5c !important; font-weight: 600 !important; font-size: 1.15rem !important; }
h3 { color: #1565c0 !important; font-weight: 600 !important; font-size: 1rem !important; }
hr { border-color: #d0dae6 !important; margin: 1rem 0 !important; }
[data-testid="stMetric"] { background: #f0f5fb; border-radius: 8px;
    padding: 0.8rem 1rem !important; border-left: 3px solid #1565c0; }
[data-testid="stMetricLabel"] { color: #546e7a !important; font-size: 0.78rem !important;
    font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.04em; }
[data-testid="stMetricValue"] { color: #0d1b2a !important; font-size: 1.35rem !important; font-weight: 700 !important; }
button[kind="secondary"] { border-radius: 6px !important; border-color: #1565c0 !important; color: #1565c0 !important; }
button[kind="secondary"]:hover { background-color: #e3f0fc !important; }
[data-testid="stExpander"] { border: 1px solid #d0dae6 !important; border-radius: 8px !important; }
[data-testid="stExpander"] summary { font-weight: 600; color: #1a3a5c; }
[data-testid="stFileUploader"] { border-radius: 8px !important; }
[data-testid="stFileUploader"] section { border-color: #90b4d8 !important; border-radius: 8px !important; background: #f7fbff !important; }
[data-testid="stFileUploader"] section:hover { border-color: #1565c0 !important; background: #eef5fd !important; }
[data-testid="stCaptionContainer"] p { color: #607d8b !important; font-size: 0.8rem !important; }
</style>
""", unsafe_allow_html=True)
else:
    st.markdown("""
<style>
/* ── Fondos ── */
.stApp, .main, .main .block-container { background-color: #0f1724 !important; }

/* ── Títulos ── */
h1 { color: #e8f4ff !important; font-weight: 700 !important; font-size: 1.6rem !important;
     border-bottom: 2px solid #4fc3f7; padding-bottom: 0.4rem; margin-bottom: 0.6rem !important; }
h2 { color: #b8d4f0 !important; font-weight: 600 !important; font-size: 1.15rem !important; }
h3 { color: #4fc3f7 !important; font-weight: 600 !important; font-size: 1rem !important; }
hr { border-color: #2a3f5f !important; margin: 1rem 0 !important; }

/* ── Texto general en área principal ── */
.main p, .main li, .main span,
[data-testid="stMarkdownContainer"] p,
[data-testid="stMarkdownContainer"] li,
[data-testid="stMarkdownContainer"] span,
[data-testid="stMarkdownContainer"] strong,
[data-testid="stMarkdownContainer"] em { color: #e0eaf4 !important; }

/* ── Labels de widgets (file uploader, inputs, etc.) ── */
[data-testid="stWidgetLabel"] p,
[data-testid="stWidgetLabel"] label,
[data-testid="stWidgetLabel"] span,
label { color: #e0eaf4 !important; }

/* ── Texto dentro del file uploader ── */
[data-testid="stFileUploader"] p,
[data-testid="stFileUploader"] span,
[data-testid="stFileUploader"] small,
[data-testid="stFileUploaderDropzone"] span { color: #b0cce8 !important; }

/* ── Captions — más visibles ── */
[data-testid="stCaptionContainer"] p,
[data-testid="stCaptionContainer"] span { color: #b0cce8 !important; font-size: 0.8rem !important; }

/* ── Métricas ── */
[data-testid="stMetric"] { background: #1a2535 !important; border-radius: 8px;
    padding: 0.8rem 1rem !important; border-left: 3px solid #4fc3f7; }
[data-testid="stMetricLabel"] p,
[data-testid="stMetricLabel"] { color: #8ab8d8 !important; font-size: 0.78rem !important;
    font-weight: 600 !important; text-transform: uppercase; letter-spacing: 0.04em; }
[data-testid="stMetricValue"] { color: #e8f4ff !important; font-size: 1.35rem !important; font-weight: 700 !important; }

/* ── Alerts (info/warning/success/error) ── */
[data-testid="stAlert"] p,
[data-testid="stAlert"] span,
[data-testid="stAlert"] li { color: #e0eaf4 !important; }

/* ── Expanders ── */
[data-testid="stExpander"] { border: 1px solid #2a3f5f !important; border-radius: 8px !important; background: #1a2535 !important; }
[data-testid="stExpander"] summary,
[data-testid="stExpander"] summary span,
[data-testid="stExpander"] summary p { font-weight: 600; color: #b8d4f0 !important; }
[data-testid="stExpander"] p,
[data-testid="stExpander"] li,
[data-testid="stExpander"] span { color: #e0eaf4 !important; }

/* ── File uploader ── */
[data-testid="stFileUploader"] { border-radius: 8px !important; }
[data-testid="stFileUploader"] section { border-color: #2a4a6e !important; border-radius: 8px !important; background: #1a2535 !important; }
[data-testid="stFileUploader"] section:hover { border-color: #4fc3f7 !important; background: #1e2f45 !important; }

/* ── Inputs ── */
input, textarea { background-color: #1a2535 !important; color: #e0eaf4 !important; border-color: #2a4a6e !important; }
[data-testid="stNumberInput"] input,
[data-testid="stTextInput"] input,
[data-testid="stDateInput"] input { color: #e0eaf4 !important; }

/* ── Botones secundarios/descarga ── */
button[kind="secondary"] { border-radius: 6px !important; border-color: #4fc3f7 !important; color: #4fc3f7 !important; }
button[kind="secondary"]:hover { background-color: rgba(79,195,247,0.1) !important; }

/* ── Código ── */
code, pre { background-color: #1a2535 !important; color: #b0cce8 !important; }
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

# Toggle Day / Dark
st.sidebar.divider()
toggle_label = "☀️ Modo Día" if dark else "🌙 Modo Oscuro"
if st.sidebar.button(toggle_label, use_container_width=True):
    st.session_state.dark_mode = not dark
    st.rerun()


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
