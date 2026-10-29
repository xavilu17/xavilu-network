import streamlit as st
# import sqlite3  # Descomenta si lo usas

# ==========================================
# 1. CONFIGURACIÓN Y LOGO
# ==========================================
st.set_page_config(page_title="Taller Olivares 4T", layout="wide", page_icon="🏍️")

# LOGO TRANSPARENTE
# (Si el archivo no existe todavía, no romperá la app gracias al try/except)
try:
    st.logo("logo_transparent.png", size="large")
except Exception:
    pass

# ==========================================
# 2. ESTADO DE NAVEGACIÓN
# ==========================================
if "pagina" not in st.session_state:
    st.session_state.pagina = "Inicio"

def ir_a(pagina):
    st.session_state.pagina = pagina

# ==========================================
# 3. BOTONERA (Grid 4x2)
# ==========================================
st.markdown("### 🧰 Panel de Control")

col1, col2, col3, col4 = st.columns(4)
col5, col6, col7, col8 = st.columns(4)

with col1:
    st.button("🏠\nInicio", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Inicio" else "secondary", 
              on_click=ir_a, args=("Inicio",))
with col2:
    st.button("📅\nAvisos ITV", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Avisos" else "secondary", 
              on_click=ir_a, args=("Avisos",))
with col3:
    st.button("🔧\nPiezas", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Piezas" else "secondary", 
              on_click=ir_a, args=("Piezas",))
with col4:
    st.button("📄\nPDFs", use_container_width=True, 
              type="primary" if st.session_state.pagina == "PDFs" else "secondary", 
              on_click=ir_a, args=("PDFs",))
with col5:
    st.button("📦\nInventario", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Inventario" else "secondary", 
              on_click=ir_a, args=("Inventario",))
with col6:
    st.button("🎙️\nVoz/IA", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Voz/IA" else "secondary", 
              on_click=ir_a, args=("Voz/IA",))
with col7:
    st.button("👥\nClientes", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Clientes" else "secondary", 
              on_click=ir_a, args=("Clientes",))
with col8:
    st.button("⚙️\nAjustes", use_container_width=True, 
              type="primary" if st.session_state.pagina == "Ajustes" else "secondary", 
              on_click=ir_a, args=("Ajustes",))

st.divider()

# ==========================================
# 4. CONTENIDO DINÁMICO
# ==========================================

if st.session_state.pagina == "Inicio":
    st.title("Bienvenido al sistema de gestión")
    st.info("Selecciona una opción en la botonera superior para empezar a trabajar.")
    # [AQUÍ VA TU CÓDIGO DE INICIO]

elif st.session_state.pagina == "Avisos":
    st.header("📅 Avisos ITV")
    # [AQUÍ VA TU CÓDIGO DE AVISOS ITV]

elif st.session_state.pagina == "Piezas":
    st.header("🔧 Gestión de Piezas")
    # [AQUÍ VA TU CÓDIGO DE PIEZAS]

elif st.session_state.pagina == "PDFs":
    st.header("📄 Generación de PDFs")
    # [AQUÍ VA TU CÓDIGO DE PDFs]

elif st.session_state.pagina == "Inventario":
    st.header("📦 Inventario")
    # [AQUÍ VA TU CÓDIGO DE INVENTARIO]

elif st.session_state.pagina == "Voz/IA":
    st.header("🎙️ Asistente Voz/IA")
    # [AQUÍ VA TU CÓDIGO DE VOZ/IA]

elif st.session_state.pagina == "Clientes":
    st.header("👥 Clientes")
    # [AQUÍ VA TU CÓDIGO DE CLIENTES]
    # Ejemplo:
    # if clientes:
    #     st.table(clientes)
    # else:
    #     st.info("No hay clientes registrados aún.")

elif st.session_state.pagina == "Ajustes":
    st.header("⚙️ Ajustes de la Aplicación")
    st.write("Configuraciones generales del sistema.")
    st.write("- Base de datos: SQLite local")
    st.write("- Tema: Rojo / Negro")
    st.write("- IA: Gemini API")
    
    if st.button("Cerrar Conexión Base de Datos"):
        # conn.close()  # Descomenta si tienes la variable global 'conn'
        st.success("Conexión cerrada.")
