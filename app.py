import streamlit as st

# ==========================================
# 1. CONFIGURACIÓN DE PÁGINA (Debe ir primero)
# ==========================================
st.set_page_config(page_title="Xavilu Network", layout="wide", page_icon="🌱")

# ==========================================
# 2. ESTÉTICA: OCULTAR BARRA BLANCA Y MENÚS
# ==========================================
st.markdown("""
    <style>
        header {visibility: hidden;}
        #MainMenu {visibility: hidden;}
        footer {visibility: hidden;}
        .block-container {padding-top: 2rem;}
    </style>
""", unsafe_allow_html=True)

# ==========================================
# 3. SISTEMA DE LOGIN
# ==========================================
def check_password():
    # Si ya se logueó antes en esta sesión, lo dejamos pasar
    if st.session_state.get("password_correct", False):
        return True
    
    # Pantalla de login
    st.markdown("## 🔐 Xavilu Network")
    st.caption("Introduce la contraseña para acceder")
    
    with st.form("login_form"):
        pwd = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Entrar", use_container_width=True)
        
    if submitted:
        # Comprueba la contraseña que tienes en los Secrets de Streamlit
        if pwd == st.secrets["PASSWORD"]:
            st.session_state["password_correct"] = True
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta. Inténtalo de nuevo.")
    return False

# Ejecutamos el login. Si falla, st.stop() detiene la app por completo aquí.
if not check_password():
    st.stop()

# Botón para cerrar sesión en la barra lateral
if st.sidebar.button("🚪 Cerrar Sesión", use_container_width=True):
    st.session_state["password_correct"] = False
    st.rerun()


# ==========================================
# 👇 AQUÍ ABAJO PEGA TU CÓDIGO ORIGINAL 👇
# ==========================================
# (A partir de aquí pon todo lo que ya tenías: tus columnas, 
# los botones de XaviWork, Gestión de Plantas, etc.)
#
# st.title("XAVILU NETWORK")
# ...
