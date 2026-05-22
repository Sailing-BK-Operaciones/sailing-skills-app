import streamlit as st
from shared_store import get_file, get_meta


def shared_or_upload(session_key, label, file_types, upload_key):
    """
    Si el archivo ya está en el store compartido (visible para todos los usuarios),
    lo muestra con fecha de última carga y lo devuelve.
    Si no, muestra un uploader individual.
    """
    f = get_file(session_key)
    if f is not None:
        _, fecha = get_meta(session_key)
        st.caption(f"✓ **{label}** — Archivos Compartidos ({fecha})")
        return f
    return st.file_uploader(label, type=file_types, key=upload_key)
