import streamlit as st

# Claves en session_state y metadata de cada archivo compartido
SHARED_FILES = [
    {
        "key":   "shared_especies",
        "label": "ESPECIES.XLS",
        "desc":  "Maestro de especies de Gallo (lista, aforo SAIL, tipo precio)",
        "types": ["xls", "xlsx"],
    },
    {
        "key":   "shared_accounts",
        "label": "table-accounts_*.csv",
        "desc":  "Account ID por comitente (tabla BYMA Clearing)",
        "types": ["csv"],
    },
    {
        "key":   "shared_pdf_aforos",
        "label": "PDF aforos BYMA",
        "desc":  "Listas de Especies Aceptadas como Garantía (circular BYMA)",
        "types": ["pdf"],
    },
    {
        "key":   "shared_pc",
        "label": "Precios de cierre PC*.XLS",
        "desc":  "Precios de cierre del día (hoja Precios_de_Cierre)",
        "types": ["xls", "xlsx"],
    },
    {
        "key":   "shared_sagaclte",
        "label": "SAGACLTE.XLS",
        "desc":  "Stock de garantías por comitente (hoja Saldos_de_Garantias)",
        "types": ["xls", "xlsx"],
    },
]


def render():
    st.title("Archivos Compartidos")
    st.markdown(
        "Estos archivos se usan en múltiples skills. Subílos acá una sola vez y "
        "quedan disponibles durante toda la sesión sin necesidad de volver a subirlos "
        "al cambiar de skill.  \n"
        "**Podés re-subir cualquier archivo en cualquier momento** si se actualizó "
        "durante la rueda — el nuevo reemplaza al anterior automáticamente."
    )
    st.divider()

    for info in SHARED_FILES:
        key   = info["key"]
        label = info["label"]
        desc  = info["desc"]
        types = info["types"]

        current = st.session_state.get(key)

        col_up, col_status = st.columns([3, 1])

        with col_up:
            uploaded = st.file_uploader(
                f"**{label}** — {desc}",
                type=types,
                key=f"uploader_{key}",
            )
            if uploaded is not None:
                st.session_state[key] = uploaded

        with col_status:
            # Vertical alignment trick
            st.write("")
            st.write("")
            if st.session_state.get(key) is not None:
                st.success("✓ Cargado")
            else:
                st.warning("Sin cargar")

    st.divider()

    # ── Resumen de estado ────────────────────────────────────────────────────
    cargados  = [f["label"] for f in SHARED_FILES if st.session_state.get(f["key"])]
    faltantes = [f["label"] for f in SHARED_FILES if not st.session_state.get(f["key"])]

    if cargados:
        st.success(f"Cargados ({len(cargados)}/{len(SHARED_FILES)}): {', '.join(cargados)}")
    if faltantes:
        st.info(f"Sin cargar: {', '.join(faltantes)}")
    if not faltantes:
        st.balloons()
