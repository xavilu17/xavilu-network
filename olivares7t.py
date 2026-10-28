import streamlit as st
import sqlite3
import pandas as pd
import datetime
import os
import urllib.parse
import streamlit.components.v1 as components

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="Olivares 4T", page_icon="🏍️", layout="wide", initial_sidebar_state="collapsed")

# --- BASE DE DATOS ---
def init_db():
    conn = sqlite3.connect('olivares4t.db')
    c = conn.cursor()
    # Tabla Clientes
    c.execute('''CREATE TABLE IF NOT EXISTS clientes 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT, telefono TEXT, email TEXT, notas TEXT)''')
    # Tabla Motos
    c.execute('''CREATE TABLE IF NOT EXISTS motos 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, cliente_id INTEGER, matricula TEXT, marca TEXT, modelo TEXT, anio TEXT,
                 FOREIGN KEY(cliente_id) REFERENCES clientes(id))''')
    # Tabla Intervenciones
    c.execute('''CREATE TABLE IF NOT EXISTS intervenciones 
                 (id INTEGER PRIMARY KEY AUTOINCREMENT, moto_id INTEGER, fecha TEXT, km TEXT, descripcion TEXT, precio REAL,
                 FOREIGN KEY(moto_id) REFERENCES motos(id))''')
    conn.commit()
    conn.close()

init_db()

def get_db_data(query, params=()):
    try:
        conn = sqlite3.connect('olivares4t.db')
        df = pd.read_sql_query(query, conn, params=params)
        conn.close()
        return df
    except Exception as e:
        st.error(f"Error DB: {e}")
        return pd.DataFrame()

def execute_db(query, params=()):
    try:
        conn = sqlite3.connect('olivares4t.db')
        conn.execute(query, params)
        conn.commit()
        conn.close()
        return True
    except Exception as e:
        st.error(f"Error DB: {e}")
        return False

# --- ESTADO DE SESIÓN ---
_defaults = {
    "password_correct": False,
    "rol": "taller",
    "sidebar_visible": True
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# --- SISTEMA DE LOGIN ---
def check_password():
    if st.session_state.get("password_correct", False):
        return True
    st.markdown("## 🔐 Olivares 4T")
    st.caption("Introduce la contraseña para acceder")
    with st.form("login_form"):
        pwd = st.text_input("Contraseña", type="password")
        submitted = st.form_submit_button("Entrar", use_container_width=True)
    if submitted:
        try:
            pwd_admin = st.secrets.get("PASSWORD", "admin")
            pwd_taller = st.secrets.get("PASSWORD_TALLER", "taller")
        except:
            pwd_admin = "admin"
            pwd_taller = "taller"
        
        if pwd == pwd_admin:
            st.session_state["password_correct"] = True
            st.session_state["rol"] = "admin"
            st.rerun()
        elif pwd == pwd_taller:
            st.session_state["password_correct"] = True
            st.session_state["rol"] = "taller"
            st.rerun()
        else:
            st.error("❌ Contraseña incorrecta.")
    return False

if not check_password():
    st.stop()

# --- CSS ESTILO XAVILU ---
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
.stApp { background-color: #000000 !important; background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
header[data-testid="stHeader"] { visibility: hidden !important; }
#MainMenu { visibility: hidden !important; }
footer { visibility: hidden !important; }
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700, 0 0 35px #FF8C00 !important; text-align: center; letter-spacing: 3px; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #ffffff !important; text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff, 0 0 30px #FF007F !important; text-align: center; letter-spacing: 3px; }
[data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; text-transform: uppercase !important; background: rgba(0, 240, 255, 0.1) !important; padding: 5px 10px !important; border-left: 4px solid #FF007F !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 8px !important; }
div.stButton > button[kind="secondary"] { background-color: #0a0a0a !important; color: #ffffff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; padding: 10px 15px !important; transition: all 0.3s ease !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.2) !important; width: 100% !important; }
div.stButton > button[kind="secondary"]:hover { background-color: rgba(0, 240, 255, 0.15) !important; color: #00f0ff !important; box-shadow: 0 0 20px rgba(0, 240, 255, 0.7) !important; transform: translateY(-2px) !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000000 !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: 900 !important; border: 2px solid #ffffff !important; border-radius: 25px !important; padding: 10px 15px !important; transition: all 0.3s ease !important; box-shadow: 0 0 18px rgba(0, 240, 255, 0.9) !important; width: 100% !important; }
.stTextArea textarea, .stTextInput input { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; color: #ffffff !important; font-size: 16px !important; }
</style>
""", unsafe_allow_html=True)

# --- FUNCIÓN DE PESTAÑAS (MACRO) ---
def custom_tabs(tab_names, key_name, default_idx=0):
    s_key = f"tab_{key_name}"
    if s_key not in st.session_state: st.session_state[s_key] = default_idx
    
    cols = st.columns(len(tab_names))
    for i, name in enumerate(tab_names):
        with cols[i]:
            if st.button(name, key=f"btn_{key_name}_{i}", use_container_width=True, type="primary" if st.session_state[s_key] == i else "secondary"):
                st.session_state[s_key] = i
                st.rerun()
    return st.session_state[s_key]

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.markdown("<h1>OLIVARES 4T</h1>", unsafe_allow_html=True)

# Sidebar básico para administración
with st.sidebar:
    st.markdown("### 🏍️ Olivares 4T")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state["password_correct"] = False
        st.rerun()
    st.info(f"Rol actual: **{st.session_state['rol']}**")

# MENÚ PRINCIPAL
macro_names = ["👥 CLIENTES", "🏍️ TALLER", "🧾 FACTURACIÓN", "💬 ASISTENTE IA", "📁 DOCUMENTOS", "⚙️ AJUSTES"]
macro_idx = custom_tabs(macro_names, "macro", 0)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# ==========================================
# CATEGORÍA 0: CLIENTES (CON BUSCADOR INTEGRADO)
# ==========================================
if macro_idx == 0:
    st.markdown("#### 👥 Directorio de Clientes")
    
    # BUSCADOR INTEGRADO
    busqueda = st.text_input("🔍 Buscar por nombre, matrícula o modelo:", placeholder="Ej: Juan o 1234ABC")
    
    if busqueda:
        # Buscamos en clientes o motos
        query = """
            SELECT c.id, c.nombre, c.telefono, m.matricula, m.marca, m.modelo 
            FROM clientes c
            LEFT JOIN motos m ON c.id = m.cliente_id
            WHERE c.nombre LIKE ? OR m.matricula LIKE ? OR m.modelo LIKE ?
        """
        param = f"%{busqueda}%"
        resultados = get_db_data(query, (param, param, param))
        
        if not resultados.empty:
            st.success(f"Se encontraron {len(resultados)} resultados:")
            for _, row in resultados.iterrows():
                with st.expander(f"👤 {row['nombre']} - 🏍️ {row['matricula'] if pd.notna(row['matricula']) else 'Sin moto'}"):
                    st.write(f"**Teléfono:** {row['telefono']}")
                    if pd.notna(row['matricula']):
                        st.write(f"**Moto:** {row['marca']} {row['modelo']}")
                    # Aquí irá el botón de ver ficha completa
        else:
            st.warning("No se encontraron clientes ni motos con esa búsqueda.")
    
    st.markdown("---")
    
    # FORMULARIO NUEVO CLIENTE
    with st.expander("➕ Dar de Alta Nuevo Cliente", expanded=False):
        with st.form("form_nuevo_cliente"):
            c1, c2 = st.columns(2)
            with c1:
                nombre = st.text_input("Nombre y Apellidos:")
                telefono = st.text_input("Teléfono:")
            with c2:
                email = st.text_input("Email (opcional):")
                notas = st.text_area("Notas del cliente:", height=100)
            
            btn_guardar = st.form_submit_button("Guardar Cliente", type="primary")
            if btn_guardar and nombre:
                if execute_db("INSERT INTO clientes (nombre, telefono, email, notas) VALUES (?, ?, ?, ?)", (nombre, telefono, email, notas)):
                    st.success(f"✅ Cliente {nombre} guardado correctamente.")
                    st.rerun()
            elif btn_guardar:
                st.error("⚠️ El nombre es obligatorio.")

    st.markdown("---")
    
    # LISTADO DE TODOS LOS CLIENTES
    st.markdown("##### Últimos clientes registrados")
    df_clientes = get_db_data("SELECT id, nombre, telefono, email FROM clientes ORDER BY id DESC LIMIT 10")
    
    if not df_clientes.empty:
        for _, row in df_clientes.iterrows():
            with st.expander(f"👤 {row['nombre']} - 📞 {row['telefono']}"):
                st.write(f"**Email:** {row['email']}")
                # Aquí buscaremos las motos de este cliente
                df_motos = get_db_data("SELECT id, matricula, marca, modelo FROM motos WHERE cliente_id = ?", (row['id'],))
                if not df_motos.empty:
                    st.markdown("**Motos asociadas:**")
                    for _, moto in df_motos.iterrows():
                        st.info(f"🏍️ {moto['matricula']} - {moto['marca']} {moto['modelo']}")
                else:
                    st.caption("Este cliente aún no tiene motos registradas.")
    else:
        st.info("Aún no hay clientes. Usa el formulario de arriba para crear el primero.")

# ==========================================
# RESTO DE CATEGORÍAS (VACÍAS POR AHORA)
# ==========================================
elif macro_idx == 1:
    st.info("🚧 Módulo de Taller en construcción (Fichas de motos e intervenciones).")
elif macro_idx == 2:
    st.info("🚧 Módulo de Facturación en construcción.")
elif macro_idx == 3:
    st.info("🚧 Módulo de Asistente IA en construcción.")
elif macro_idx == 4:
    st.info("🚧 Módulo de Documentos en construcción.")
elif macro_idx == 5:
    st.markdown("#### ⚙️ Ajustes del Sistema")
    if st.session_state.get('rol') == 'admin':
        st.write("Aquí irán los ajustes de la empresa, logo y claves API.")
    else:
        st.warning("No tienes permisos para ver esta sección.")
