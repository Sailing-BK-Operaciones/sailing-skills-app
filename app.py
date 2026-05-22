import streamlit as st

st.set_page_config(
    page_title="Sailing Inversiones — Skills",
    page_icon="📊",
    layout="wide"
)

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

# ── Sidebar ──────────────────────────────────────────────────────────────────
st.sidebar.title("Skills disponibles")

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

# Indicador de archivos compartidos cargados
_SHARED_KEYS = ["shared_especies", "shared_accounts", "shared_pdf_aforos",
                "shared_pc", "shared_sagaclte"]
_n = sum(1 for k in _SHARED_KEYS if st.session_state.get(k))
if _n == len(_SHARED_KEYS):
    st.sidebar.success(f"✓ {_n}/{len(_SHARED_KEYS)} archivos compartidos")
elif _n > 0:
    st.sidebar.warning(f"⚠ {_n}/{len(_SHARED_KEYS)} archivos compartidos")
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
