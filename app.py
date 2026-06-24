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
/* Ancho mínimo solo cuando está expandido — al colapsar el panel principal se expande normal */
section[data-testid="stSidebar"][aria-expanded="true"] {
    min-width: 320px !important;
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

/* ── Botones primarios — legibles en modo day y night ── */
button[kind="primary"],
button[kind="primaryFormSubmit"] {
    background-color: #1565c0 !important;
    border-color:     #1565c0 !important;
    color:            #ffffff !important;
    font-weight: 600 !important;
    letter-spacing: 0.02em;
    border-radius: 6px !important;
}
button[kind="primary"] p,
button[kind="primary"] span,
button[kind="primaryFormSubmit"] p,
button[kind="primaryFormSubmit"] span {
    color: #ffffff !important;
}
button[kind="primary"]:hover,
button[kind="primaryFormSubmit"]:hover {
    background-color: #0d47a1 !important;
    border-color:     #0d47a1 !important;
    color:            #ffffff !important;
}
button[kind="primary"]:disabled,
button[kind="primaryFormSubmit"]:disabled {
    background-color: #1565c0 !important;
    border-color:     #1565c0 !important;
    color:            #ffffff !important;
    opacity: 0.42 !important;
    cursor: not-allowed !important;
}
button[kind="primary"]:disabled p,
button[kind="primary"]:disabled span,
button[kind="primaryFormSubmit"]:disabled p,
button[kind="primaryFormSubmit"]:disabled span {
    color: #ffffff !important;
}

/* ── Alerts ── */
[data-testid="stAlert"] { border-radius: 6px !important; }

/* ── Expanders dentro del sidebar ── */
section[data-testid="stSidebar"] [data-testid="stExpander"] {
    border: 1px solid #1e3a5f !important;
    border-radius: 6px !important;
    background: rgba(255,255,255,0.04) !important;
    margin-bottom: 0.35rem !important;
}
/* Colapsado: tono apagado */
section[data-testid="stSidebar"] [data-testid="stExpander"] summary {
    color: #7a9bc0 !important;
    font-size: 0.74rem !important;
    font-weight: 700 !important;
    letter-spacing: 0.07em !important;
    padding: 0.35rem 0.55rem !important;
    text-transform: uppercase;
}
section[data-testid="stSidebar"] details > summary svg { fill: #7a9bc0 !important; }

/* Expandido: fondo gris-azul sutil, texto más visible */
section[data-testid="stSidebar"] details[open] > summary {
    color: #cce8f8 !important;
    background: rgba(255,255,255,0.09) !important;
    border-bottom: 1px solid #1e3a5f !important;
    border-radius: 5px 5px 0 0 !important;
}
section[data-testid="stSidebar"] details[open] > summary svg { fill: #cce8f8 !important; }

/* ── Inputs dentro del sidebar: siempre con fondo oscuro y texto claro ── */
section[data-testid="stSidebar"] input,
section[data-testid="stSidebar"] textarea {
    background-color: #1e3a5f !important;
    color: #e8f4ff !important;
    border: 1px solid #3a6a9f !important;
    border-radius: 4px !important;
}
section[data-testid="stSidebar"] input::placeholder,
section[data-testid="stSidebar"] textarea::placeholder {
    color: #7a9bc0 !important;
}

/* ── Botón de formulario en sidebar (Guardar usuario) ── */
section[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] > button,
section[data-testid="stSidebar"] button[kind="primaryFormSubmit"] {
    background-color: #1565c0 !important;
    color: #ffffff !important;
    border-color: #1565c0 !important;
    font-weight: 600 !important;
    opacity: 1 !important;
}
section[data-testid="stSidebar"] [data-testid="stFormSubmitButton"] > button:hover,
section[data-testid="stSidebar"] button[kind="primaryFormSubmit"]:hover {
    background-color: #0d47a1 !important;
    border-color: #0d47a1 !important;
}

/* ── Labels de inputs dentro del sidebar ── */
section[data-testid="stSidebar"] [data-testid="stWidgetLabel"] p {
    color: #b8d4f0 !important;
    font-size: 0.8rem !important;
}

/* ── Botón toggle tema: compacto en header del sidebar ── */
section[data-testid="stSidebar"] button[kind="secondary"] {
    padding: 0.12rem 0.32rem !important;
    font-size: 0.95rem !important;
    line-height: 1.1 !important;
    border: 1px solid rgba(255,255,255,0.18) !important;
    background: transparent !important;
    box-shadow: none !important;
    min-height: unset !important;
}
section[data-testid="stSidebar"] button[kind="secondary"]:hover {
    background: rgba(255,255,255,0.09) !important;
    border-color: rgba(255,255,255,0.35) !important;
}
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
from auth import check_credentials, render_admin_panel

MAX_INTENTOS = 5

def check_login():
    if "authenticated"  not in st.session_state:
        st.session_state.authenticated  = False
    if "failed_attempts" not in st.session_state:
        st.session_state.failed_attempts = 0
    if "username" not in st.session_state:
        st.session_state.username = ""
    if "role" not in st.session_state:
        st.session_state.role = ""

    if not st.session_state.authenticated:
        st.title("Sailing Inversiones")
        st.subheader("Acceso al panel de skills operativas")

        if st.session_state.failed_attempts >= MAX_INTENTOS:
            st.error(
                f"Acceso bloqueado por {MAX_INTENTOS} intentos fallidos. "
                "Cerrá esta pestaña y volvé a abrir el link para reintentar."
            )
            return False

        with st.form("login_form"):
            username  = st.text_input("Usuario:")
            password  = st.text_input("Contraseña:", type="password")
            submitted = st.form_submit_button("Ingresar", use_container_width=True)

        if submitted:
            role = check_credentials(username, password)
            if role:
                st.session_state.authenticated  = True
                st.session_state.username       = username.strip().lower()
                st.session_state.role           = role
                st.session_state.failed_attempts = 0
                st.rerun()
            else:
                st.session_state.failed_attempts += 1
                restantes = MAX_INTENTOS - st.session_state.failed_attempts
                if restantes > 0:
                    st.error(
                        f"Usuario o contraseña incorrectos. "
                        f"Intentos restantes: {restantes}"
                    )
                else:
                    st.rerun()
        return False
    return True

if not check_login():
    st.stop()


# ── Navegación agrupada ───────────────────────────────────────────────────────
OP_SKILLS = [
    "Archivos Compartidos",
    "Actual Position",
    "Consolida SI2 para NASDAQ",
    "Risk Monitoring Client",
    "Control diario OP-SDIB (MA)",
    "Control Aforos BYMA",
    "Conversión Dólares Renta",
    "Control Márgenes Gara BYMA",
    "Arreglos Garantías",
]
GEST_SKILLS = [
    "Tesorería",
    "Reporte Operativo",
]
SUSP_SKILLS = [
    "Risk Position",
    "Distribución Gara BYMA",
    "Genera TXT Gara NASDAQ",
    "Collateral Position",
    "Settlement Position",
    "Settlement Instruction",
    "Settlement Obligation",
]

# ── Estado de navegación: inicializar keys solo si no existen ─────────────────
if "active_skill" not in st.session_state:
    st.session_state.active_skill = "Archivos Compartidos"

_active = st.session_state.active_skill
if "_nav_op"   not in st.session_state:
    st.session_state._nav_op   = _active if _active in OP_SKILLS   else None
if "_nav_gest" not in st.session_state:
    st.session_state._nav_gest = _active if _active in GEST_SKILLS else None
if "_nav_susp" not in st.session_state:
    st.session_state._nav_susp = _active if _active in SUSP_SKILLS else None

# ── Callbacks: actualizan active_skill y limpian los otros grupos ─────────────
def _on_op_change():
    chosen = st.session_state._nav_op
    if chosen:
        st.session_state.active_skill = chosen
        st.session_state._nav_gest = None
        st.session_state._nav_susp = None

def _on_gest_change():
    chosen = st.session_state._nav_gest
    if chosen:
        st.session_state.active_skill = chosen
        st.session_state._nav_op   = None
        st.session_state._nav_susp = None

def _on_susp_change():
    chosen = st.session_state._nav_susp
    if chosen:
        st.session_state.active_skill = chosen
        st.session_state._nav_op   = None
        st.session_state._nav_gest = None

# ── Header: título + toggle day/dark (icono compacto) ────────────────────────
_hcol_title, _hcol_toggle = st.sidebar.columns([5, 1])
_hcol_title.markdown(
    '<p style="color:#ffffff;font-size:1.05rem;font-weight:700;'
    'letter-spacing:0.03em;padding:0.25rem 0;'
    'border-bottom:1px solid #2a4a6e;margin:0 0 0.55rem 0;">'
    'Sailing Inversiones</p>',
    unsafe_allow_html=True
)
if _hcol_toggle.button("☀️" if dark else "🌙", key="_toggle_dark",
                       help="Modo Día / Modo Oscuro"):
    st.session_state.dark_mode = not dark
    st.rerun()

# ── Usuario conectado + logout ────────────────────────────────────────────────
_uname = st.session_state.get("username", "")
_role  = st.session_state.get("role", "")
_col_usr, _col_logout = st.sidebar.columns([4, 1])
_col_usr.markdown(
    f'<p style="color:#8ab8d8;font-size:0.78rem;margin:0;padding:0.1rem 0;">'
    f'{"👑 " if _role == "admin" else "👤 "}{_uname}</p>',
    unsafe_allow_html=True,
)
if _col_logout.button("↩", key="_logout", help="Cerrar sesión"):
    for k in ["authenticated", "username", "role", "failed_attempts"]:
        st.session_state.pop(k, None)
    st.rerun()

# ── Indicador de archivos compartidos (siempre visible, antes de los grupos) ─
from shared_store import is_loaded as _is_loaded
_SHARED_KEYS = [
    "shared_pc", "shared_sagaclte", "shared_sateclte", "shared_contbole",
    "shared_saldos_nasdaq", "shared_prices", "shared_especies", "shared_accounts",
    "shared_aforo_sail", "shared_tabcompb",
]
_n = sum(1 for k in _SHARED_KEYS if _is_loaded(k))
_total = len(_SHARED_KEYS)
if _n == _total:
    st.sidebar.success(f"✓ {_n}/{_total} archivos compartidos")
elif _n > 0:
    st.sidebar.warning(f"⚠ {_n}/{_total} archivos compartidos")
else:
    st.sidebar.info("Sin archivos compartidos")
st.sidebar.divider()

# ── Grupo 1 ──
with st.sidebar.expander("⚙ OPERATIVAS / GARANTÍAS", expanded=True, key="exp_op"):
    st.radio("", OP_SKILLS, index=None,
             key="_nav_op", on_change=_on_op_change,
             label_visibility="collapsed")

# ── Grupo 2 ──
with st.sidebar.expander("📊 GESTIÓN DEL ÁREA", expanded=False, key="exp_gest"):
    st.radio("", GEST_SKILLS, index=None,
             key="_nav_gest", on_change=_on_gest_change,
             label_visibility="collapsed")

# ── Grupo 3 ──
with st.sidebar.expander("⏸ EN SUSPENSO", expanded=False, key="exp_susp"):
    st.radio("", SUSP_SKILLS, index=None,
             key="_nav_susp", on_change=_on_susp_change,
             label_visibility="collapsed")

skill = st.session_state.active_skill

# ── Panel admin al pie del sidebar (solo visible para el administrador) ───────
if _role == "admin":
    render_admin_panel()



# ── Routing ───────────────────────────────────────────────────────────────────
if skill == "Archivos Compartidos":
    from shared_inputs import render
    render()
elif skill == "Actual Position":
    from skills.actual_position.ui import render
    render()
elif skill == "Risk Position":
    from skills.risk_position.ui import render
    render()
elif skill == "Risk Monitoring Client":
    from skills.risk_monitoring_client.ui import render
    render()
elif skill == "Distribución Gara BYMA":
    from skills.distribucion_gara.ui import render
    render()
elif skill == "Genera TXT Gara NASDAQ":
    from skills.genera_txt_gara_nasdaq.ui import render
    render()
elif skill == "Arreglo Dev Gara Gallo":
    from skills.arreglo_dev_gara_gallo.ui import render
    render()
elif skill == "Consolida SI2 para NASDAQ":
    from skills.generar_withdraw_si2.ui import render
    render()
elif skill == "Collateral Position":
    from skills.Collateral_position.ui import render
    render()
elif skill == "Control Aforos BYMA":
    from skills.control_aforos_byma.ui import render
    render()
elif skill == "Control Márgenes Gara BYMA":
    from skills.control_margenes_gara_byma.ui import render
    render()
elif skill == "Arreglos Garantías":
    from skills.arreglos_garantias.ui import render
    render()
elif skill == "Control diario OP-SDIB (MA)":
    from skills.control_diario_op_sdib.ui import render
    render()
elif skill == "Conversión Dólares Renta":
    from skills.conversion_dolares_renta.ui import render
    render()
elif skill == "Tesorería":
    from skills.tesoreria.ui import render
    render()
elif skill == "Reporte Operativo":
    from skills.reporte_operativo.ui import render
    render()
else:
    st.info(f"**{skill}** — skill en construcción.", icon="🚧")
