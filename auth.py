"""
Sistema de autenticación multi-usuario — Sailing Skills App.

Credenciales:
  - Admin : ADMIN_USER + ADMIN_PASS en Streamlit Secrets
  - Equipo: sección [USERS] en Streamlit Secrets (base permanente)
             + altas/bajas runtime vía panel admin
             (las modificaciones runtime persisten en memoria del servidor
              hasta el próximo redeploy de la app)
"""
import streamlit as st


# ── Store de usuarios en memoria compartida ───────────────────────────────────

@st.cache_resource
def _user_store() -> dict:
    """
    Dict {username_lower: password} — único en memoria del servidor.
    Se inicializa una sola vez desde la sección [USERS] de Secrets.
    El admin puede agregar/quitar usuarios en runtime.
    """
    users: dict = {}
    try:
        if "USERS" in st.secrets:
            for uname, pwd in st.secrets["USERS"].items():
                users[uname.strip().lower()] = str(pwd)
    except Exception:
        pass
    return users


# ── Helpers internos ──────────────────────────────────────────────────────────

def _admin_user() -> str:
    return str(st.secrets.get("ADMIN_USER", "admin")).strip().lower()


def _admin_pass() -> str:
    return str(st.secrets.get("ADMIN_PASS", ""))


# ── API pública ───────────────────────────────────────────────────────────────

def check_credentials(username: str, password: str):
    """
    Valida usuario + contraseña.
    Retorna 'admin', 'user', o None si las credenciales son incorrectas.
    """
    uname = username.strip().lower()
    if not uname or not password:
        return None
    if uname == _admin_user() and password == _admin_pass():
        return "admin"
    users = _user_store()
    if uname in users and users[uname] == password:
        return "user"
    return None


def add_user(username: str, password: str) -> str | None:
    """
    Agrega o modifica un usuario en el store runtime.
    Retorna un mensaje de error (str) o None si fue exitoso.
    """
    uname = username.strip().lower()
    if not uname:
        return "El nombre de usuario no puede estar vacío."
    if not password:
        return "La contraseña no puede estar vacía."
    if uname == _admin_user():
        return "No se puede crear un usuario con el nombre del administrador."
    _user_store()[uname] = password
    return None


def remove_user(username: str) -> None:
    _user_store().pop(username.strip().lower(), None)


def get_usernames() -> list[str]:
    """Retorna lista ordenada de nombres de usuario (sin contraseñas)."""
    return sorted(_user_store().keys())


# ── Panel admin ───────────────────────────────────────────────────────────────

def render_admin_panel() -> None:
    """Renderiza el panel de administración de usuarios en el sidebar."""
    with st.sidebar.expander("👤 ADMINISTRADOR DE USUARIOS", expanded=False):

        # ── Lista de usuarios actuales ────────────────────────────────────────
        usernames = get_usernames()
        if usernames:
            st.markdown("**Usuarios activos:**")
            for uname in usernames:
                col_u, col_del = st.columns([4, 1])
                col_u.markdown(f"• `{uname}`")
                if col_del.button(
                    "✕", key=f"_del_usr_{uname}",
                    help=f"Eliminar usuario {uname}",
                ):
                    remove_user(uname)
                    st.rerun()
        else:
            st.caption("Sin usuarios del equipo configurados.")

        st.divider()

        # ── Formulario alta / modificación ───────────────────────────────────
        st.markdown("**Agregar / modificar usuario:**")
        with st.form("_admin_add_user", clear_on_submit=True):
            new_uname = st.text_input(
                "Nombre de usuario",
                placeholder="ej: ana.garcia",
            )
            new_pass  = st.text_input("Contraseña", type="password")
            new_pass2 = st.text_input("Repetir contraseña", type="password")
            guardar   = st.form_submit_button("Guardar usuario", use_container_width=True)

        if guardar:
            if new_pass != new_pass2:
                st.error("Las contraseñas no coinciden.")
            else:
                err = add_user(new_uname, new_pass)
                if err:
                    st.error(err)
                else:
                    action = "modificado" if new_uname.strip().lower() in get_usernames() else "creado"
                    st.success(f"✓ Usuario `{new_uname.strip().lower()}` {action}.")
                    st.rerun()

        # ── Nota de persistencia ─────────────────────────────────────────────
        st.caption(
            "⚠ Los cambios hechos acá persisten en memoria del servidor "
            "hasta el próximo redeploy de la app. "
            "Para usuarios permanentes (que sobrevivan redeploys), "
            "agregalos también en la sección `[USERS]` de los Secrets "
            "en el dashboard de Streamlit Cloud."
        )
