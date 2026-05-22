import streamlit as st
from shared_store import save_file, get_meta, is_loaded

# ── Definición de archivos compartidos ────────────────────────────────────────
# Grupo 1: cambian cada rueda / cada día
DEL_DIA = [
    {
        "key":   "shared_pc",
        "label": "PC*.XLS",
        "desc":  "Precios de cierre del día",
        "types": ["xls", "xlsx"],
        "icon":  "📈",
    },
    {
        "key":   "shared_sagaclte",
        "label": "SAGACLTE.XLS",
        "desc":  "Garantías en BYMA por comitente",
        "types": ["xls", "xlsx"],
        "icon":  "🔒",
    },
    {
        "key":   "shared_sateclte",
        "label": "SATECLTE.XLS",
        "desc":  "Tenencia disponible por comitente",
        "types": ["xls", "xlsx"],
        "icon":  "📦",
    },
    {
        "key":   "shared_contbole",
        "label": "CONTBOLE.XLS",
        "desc":  "Boletos del día desde Gallo",
        "types": ["xls", "xlsx"],
        "icon":  "📋",
    },
    {
        "key":   "shared_saldos_nasdaq",
        "label": "saldos al inicio Nasdaq.csv",
        "desc":  "Saldos de inicio de cuentas NASDAQ",
        "types": ["csv"],
        "icon":  "💰",
    },
]

# Grupo 2: cambian ocasionalmente (referencia)
REFERENCIA = [
    {
        "key":   "shared_especies",
        "label": "ESPECIES.XLS",
        "desc":  "Maestro de especies de Gallo",
        "types": ["xls", "xlsx"],
        "icon":  "📗",
    },
    {
        "key":   "shared_accounts",
        "label": "table-accounts_*.csv",
        "desc":  "Account ID por comitente (BYMA Clearing)",
        "types": ["csv"],
        "icon":  "🔑",
    },
    {
        "key":   "shared_pdf_aforos",
        "label": "PDF aforos BYMA",
        "desc":  "Listas de Especies Aceptadas como Garantía",
        "types": ["pdf"],
        "icon":  "📄",
    },
    {
        "key":   "shared_aforo_sail",
        "label": "tabla listas gallo vs aforos.xlsx",
        "desc":  "Aforo SAIL por lista Gallo",
        "types": ["xlsx"],
        "icon":  "📊",
    },
    {
        "key":   "shared_tabcompb",
        "label": "TABCOMPB.XLS",
        "desc":  "Clasificador de tipos de operación",
        "types": ["xls", "xlsx"],
        "icon":  "🗂️",
    },
]

ALL_SHARED = DEL_DIA + REFERENCIA


def _render_group(files: list):
    """Renderiza una grilla de 2 columnas con uploaders + estado."""
    cols = st.columns(2)
    for i, info in enumerate(files):
        key   = info["key"]
        icon  = info["icon"]
        label = info["label"]
        desc  = info["desc"]
        types = info["types"]

        with cols[i % 2]:
            nombre, fecha = get_meta(key)
            loaded = is_loaded(key)

            # Cabecera del card
            estado_html = (
                f'<span style="color:#1e8449;font-weight:600">✓ {fecha}</span>'
                if loaded else
                '<span style="color:#b0bec5">Sin cargar</span>'
            )
            st.markdown(
                f'<div style="margin-bottom:2px">'
                f'{icon} <strong>{label}</strong>&nbsp;&nbsp;{estado_html}'
                f'</div>',
                unsafe_allow_html=True,
            )
            if loaded and nombre:
                st.caption(f"📁 {nombre}")

            uploaded = st.file_uploader(
                desc,
                type=types,
                key=f"uploader_{key}",
                label_visibility="visible",
            )
            if uploaded is not None:
                upload_id = f"{uploaded.name}_{uploaded.size}"
                if st.session_state.get(f"_sid_{key}") != upload_id:
                    save_file(key, uploaded)
                    st.session_state[f"_sid_{key}"] = upload_id
                    st.rerun()

            st.write("")  # separación entre cards


def render():
    st.title("Archivos Compartidos")
    st.markdown(
        "Subí los archivos una sola vez y quedan disponibles para **todas las skills** "
        "y para **todos los usuarios** conectados.  \n"
        "Re-subí cualquier archivo si se actualizó durante la rueda — reemplaza al anterior."
    )

    # Banner de estado global
    n_cargados = sum(1 for f in ALL_SHARED if is_loaded(f["key"]))
    n_total    = len(ALL_SHARED)
    if n_cargados == n_total:
        st.success(f"✓ {n_cargados}/{n_total} archivos cargados — todo listo para operar")
    elif n_cargados > 0:
        faltantes = [f["label"] for f in ALL_SHARED if not is_loaded(f["key"])]
        st.warning(f"⚠ {n_cargados}/{n_total} cargados — faltan: {', '.join(faltantes)}")
    else:
        st.info("Sin archivos cargados — subí los archivos del día para comenzar")

    st.divider()

    # ── Grupo 1: Del día ──────────────────────────────────────────────────────
    st.markdown("#### 📅 Del día")
    st.caption("Archivos que se actualizan en cada rueda")
    _render_group(DEL_DIA)

    st.divider()

    # ── Grupo 2: Referencia ───────────────────────────────────────────────────
    st.markdown("#### 📚 Referencia")
    st.caption("Archivos que cambian ocasionalmente")
    _render_group(REFERENCIA)
