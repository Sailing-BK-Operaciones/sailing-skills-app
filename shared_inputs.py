import streamlit as st
from shared_store import save_file, get_meta, is_loaded

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
        "quedan disponibles para **todos los usuarios** durante toda la sesión.  \n"
        "**Podés re-subir cualquier archivo en cualquier momento** si se actualizó "
        "durante la rueda — el nuevo reemplaza al anterior automáticamente."
    )
    st.info(
        "🔄 Los archivos son compartidos entre usuarios: lo que sube un miembro del equipo "
        "lo ven todos los demás en su próxima acción.",
        icon=None,
    )
    st.divider()

    for info in SHARED_FILES:
        key   = info["key"]
        label = info["label"]
        desc  = info["desc"]
        types = info["types"]

        nombre, fecha = get_meta(key)

        col_up, col_status = st.columns([3, 1])

        with col_up:
            uploaded = st.file_uploader(
                f"**{label}** — {desc}",
                type=types,
                key=f"uploader_{key}",
            )
            if uploaded is not None:
                # Identificamos el archivo por nombre + tamaño para evitar el loop
                # pero permitir re-subir el mismo nombre con contenido actualizado
                upload_id = f"{uploaded.name}_{uploaded.size}"
                if st.session_state.get(f"_sid_{key}") != upload_id:
                    save_file(key, uploaded)
                    st.session_state[f"_sid_{key}"] = upload_id
                    st.rerun()

        with col_status:
            st.write("")
            st.write("")
            if is_loaded(key):
                nombre, fecha = get_meta(key)
                st.success(f"✓ Cargado  \n{fecha}")
                st.caption(nombre)
            else:
                st.warning("Sin cargar")

    st.divider()

    cargados  = [f["label"] for f in SHARED_FILES if is_loaded(f["key"])]
    faltantes = [f["label"] for f in SHARED_FILES if not is_loaded(f["key"])]

    if cargados:
        st.success(f"Cargados ({len(cargados)}/{len(SHARED_FILES)}): {', '.join(cargados)}")
    if faltantes:
        st.info(f"Sin cargar: {', '.join(faltantes)}")
    if not faltantes:
        st.balloons()
