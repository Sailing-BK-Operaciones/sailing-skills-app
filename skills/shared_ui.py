import streamlit as st


def shared_or_upload(session_key, label, file_types, upload_key):
    """Returns file from shared session_state, or shows a fresh uploader."""
    f = st.session_state.get(session_key)
    if f is not None:
        st.caption(f"✓ **{label}** — desde Archivos Compartidos")
        return f
    return st.file_uploader(label, type=file_types, key=upload_key)
