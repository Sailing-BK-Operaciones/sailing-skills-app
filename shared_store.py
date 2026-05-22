"""
Store compartido entre todas las sesiones (usuarios) de la app.
Usa st.cache_resource para mantener un único objeto en memoria del servidor.
Los archivos persisten hasta el próximo re-deploy de la app.
"""
import io
from datetime import datetime, timezone, timedelta
import streamlit as st

AR_TZ = timezone(timedelta(hours=-3))  # UTC-3 Buenos Aires


@st.cache_resource
def _store() -> dict:
    """Diccionario único en memoria, compartido por todas las sesiones."""
    return {}


def save_file(key: str, uploaded_file) -> None:
    """Guarda un archivo subido en el store compartido."""
    uploaded_file.seek(0)
    _store()[key] = {
        "bytes":       uploaded_file.read(),
        "name":        uploaded_file.name,
        "uploaded_at": datetime.now(AR_TZ).strftime("%d/%m/%Y %H:%M"),
    }


def get_file(key: str):
    """Devuelve un BytesIO con el archivo, o None si no existe."""
    entry = _store().get(key)
    if entry is None:
        return None
    buf = io.BytesIO(entry["bytes"])
    buf.name = entry["name"]
    return buf


def get_meta(key: str):
    """Devuelve (nombre, fecha_hora_carga) o (None, None) si no existe."""
    entry = _store().get(key)
    if entry:
        return entry["name"], entry["uploaded_at"]
    return None, None


def is_loaded(key: str) -> bool:
    return key in _store()
