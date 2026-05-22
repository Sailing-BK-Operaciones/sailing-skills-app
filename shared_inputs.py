import streamlit as st
from shared_store import save_file, get_meta, is_loaded

# ── Archivos compartidos ───────────────────────────────────────────────────────
DEL_DIA = [
    {"key": "shared_pc",            "label": "PC*.XLS",                     "desc": "Precios de cierre del día (hoja Precios_de_Cierre)",          "types": ["xls", "xlsx"]},
    {"key": "shared_sagaclte",      "label": "SAGACLTE.XLS",                "desc": "Garantías en BYMA por comitente (hoja Saldos_de_Garantias)",  "types": ["xls", "xlsx"]},
    {"key": "shared_sateclte",      "label": "SATECLTE.XLS",                "desc": "Tenencia disponible por comitente (hoja Saldos_de_Tenencia)", "types": ["xls", "xlsx"]},
    {"key": "shared_contbole",      "label": "CONTBOLE.XLS",                "desc": "Boletos del día desde Gallo (hoja Control_de_Boletos)",       "types": ["xls", "xlsx"]},
    {"key": "shared_saldos_nasdaq", "label": "saldos al inicio Nasdaq.csv", "desc": "Saldos de inicio de cuentas NASDAQ",                         "types": ["csv"]},
]

REFERENCIA = [
    {"key": "shared_especies",   "label": "ESPECIES.XLS",                       "desc": "Maestro de especies de Gallo (lista, aforo SAIL, tipo precio)",  "types": ["xls", "xlsx"]},
    {"key": "shared_accounts",   "label": "table-accounts_*.csv",               "desc": "Account ID por comitente (BYMA Clearing)",                       "types": ["csv"]},
    {"key": "shared_pdf_aforos", "label": "PDF aforos BYMA",                    "desc": "Listas de Especies Aceptadas como Garantía (circular BYMA)",     "types": ["pdf"]},
    {"key": "shared_aforo_sail", "label": "tabla listas gallo vs aforos.xlsx",  "desc": "Aforo SAIL por lista Gallo",                                     "types": ["xlsx"]},
    {"key": "shared_tabcompb",   "label": "TABCOMPB.XLS",                       "desc": "Clasificador de tipos de operación (activa ops CI en CONTBOLE)", "types": ["xls", "xlsx"]},
]

ALL_SHARED = DEL_DIA + REFERENCIA


def _render_group(files: list):
    """Grilla compacta de 3 columnas. Descripción como tooltip (help)."""
    for i in range(0, len(files), 3):
        chunk = files[i:i + 3]
        cols = st.columns(3)
        for j, info in enumerate(chunk):
            key    = info["key"]
            loaded = is_loaded(key)
            nombre, fecha = get_meta(key)

            with cols[j]:
                uploaded = st.file_uploader(
                    info["label"],
                    type=info["types"],
                    key=f"uploader_{key}",
                    help=info["desc"],
                )
                # Estado compacto bajo el uploader
                if loaded:
                    st.caption(f"✅ {fecha} · {nombre}" if nombre else f"✅ {fecha}")
                else:
                    st.caption("⬜ Sin cargar")

                # Guardar solo si es un archivo nuevo (evita loop)
                if uploaded is not None:
                    upload_id = f"{uploaded.name}_{uploaded.size}"
                    if st.session_state.get(f"_sid_{key}") != upload_id:
                        save_file(key, uploaded)
                        st.session_state[f"_sid_{key}"] = upload_id
                        st.rerun()


def render():
    st.title("Archivos Compartidos")
    st.markdown(
        "Subí los archivos **una sola vez** — quedan disponibles para todas las skills "
        "y para todos los usuarios. Re-subí cualquier archivo si se actualizó durante la rueda."
    )

    # Banner de estado global (compacto, una línea)
    n  = sum(1 for f in ALL_SHARED if is_loaded(f["key"]))
    nt = len(ALL_SHARED)
    if n == nt:
        st.success(f"✓ {n}/{nt} archivos cargados — todo listo")
    elif n > 0:
        faltantes = [f["label"] for f in ALL_SHARED if not is_loaded(f["key"])]
        st.warning(f"⚠ {n}/{nt} cargados — faltan: {', '.join(faltantes)}")
    else:
        st.info(f"0/{nt} archivos cargados")

    st.divider()

    st.markdown("**📅 Del día** — se actualizan cada rueda")
    _render_group(DEL_DIA)

    st.divider()

    st.markdown("**📚 Referencia** — cambian ocasionalmente")
    _render_group(REFERENCIA)
