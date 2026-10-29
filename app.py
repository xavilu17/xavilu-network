import streamlit as st
import urllib.parse # Para el botón de WhatsApp
import datetime # Para la fecha del albarán

# ==========================================
# 1. CONFIGURACIÓN Y LOGO
# ==========================================
st.set_page_config(page_title="Taller Olivares 4T", layout="wide", page_icon="🏍️")

# LOGO TRANSPARENTE (Descomenta cuando tengas el archivo listo)
# st.logo("logo_transparent.png", size="large")

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
    st.button("📄\nPDFs / Albaranes", use_container_width=True, 
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

# --- PESTAÑA 1: INICIO ---
if st.session_state.pagina == "Inicio":
    st.title("Bienvenido al sistema de gestión")
    st.info("Selecciona una opción en la botonera superior para empezar a trabajar.")
    # [AQUÍ VA TU CÓDIGO DE INICIO SI TENÍAS ALGO MÁS]

# --- PESTAÑA 2: AVISOS ITV ---
elif st.session_state.pagina == "Avisos":
    st.header("📅 Avisos ITV")
    st.write("Gestión de avisos y recordatorios de ITV.")
    
    st.subheader("Ejemplo de Aviso por WhatsApp")
    # Datos de ejemplo (luego los sacarías de tu base de datos)
    tel_ejemplo = "34600000000" 
    nombre_ejemplo = "Cliente Prueba"
    moto_ejemplo = "Yamaha T-Max"
    
    # Crear el mensaje
    mensaje = f"Hola {nombre_ejemplo}, te escribimos de Taller Olivares 4T. Tu {moto_ejemplo} tiene la ITV próxima a caducar. ¿Quieres que te pidamos cita? 🏍️"
    mensaje_url = urllib.parse.quote(mensaje)
    url_whatsapp = f"https://wa.me/{tel_ejemplo}?text={mensaje_url}"
    
    st.link_button("📱 Avisar ITV por WhatsApp", url_whatsapp, type="primary")
    
    # [AQUÍ VA TU CÓDIGO DE AVISOS ITV REAL]

# --- PESTAÑA 3: PIEZAS ---
elif st.session_state.pagina == "Piezas":
    st.header("🔧 Gestión de Piezas")
    st.write("Control y búsqueda de repuestos.")
    # [AQUÍ VA TU CÓDIGO DE PIEZAS]

# --- PESTAÑA 4: PDFs / ALBARANES ---
elif st.session_state.pagina == "PDFs":
    st.header("📄 Generación de Albaranes")
    st.write("Crea y descarga el albarán del trabajo realizado.")
    
    with st.form("form_albaran"):
        cliente = st.text_input("Nombre del Cliente")
        matricula = st.text_input("Matrícula de la Moto")
        trabajo = st.text_area("Trabajo realizado / Piezas")
        submit = st.form_submit_button("Generar Albarán")
        
        if submit:
            # Generar texto para TXT
            fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")
            texto_albaran = f"""
            =========================================
                    TALLER OLIVARES 4T
            =========================================
            Fecha: {fecha_hoy}
            Cliente: {cliente}
            Matrícula: {matricula}
            -----------------------------------------
            DETALLES DEL TRABAJO:
            {trabajo}
            -----------------------------------------
            * Documento no válido como factura.
            """
            
            st.success("¡Albarán generado con éxito!")
            # Botón para descargar en TXT
            st.download_button(
                label="⬇️ Descargar Albarán en TXT",
                data=texto_albaran,
                file_name=f"Albaran_{cliente}_{fecha_hoy.replace('/','-')}.txt",
                mime="text/plain",
                type="primary"
            )
            # [AQUÍ AÑADIREMOS EL BOTÓN DE DESCARGAR PDF CON EL LOGO]

# --- PESTAÑA 5: INVENTARIO ---
elif st.session_state.pagina == "Inventario":
    st.header("📦 Inventario")
    st.write("Gestión de stock del taller.")
    # [AQUÍ VA TU CÓDIGO DE INVENTARIO]

# --- PESTAÑA 6: VOZ/IA ---
elif st.session_state.pagina == "Voz/IA":
    st.header("🎙️ Asistente Voz/IA")
    st.write("Módulo de inteligencia artificial y dictado por voz.")
    # [AQUÍ VA TU CÓDIGO DE VOZ/IA]

# --- PESTAÑA 7: CLIENTES ---
elif st.session_state.pagina == "Clientes":
    st.header("👥 Clientes")
    st.write("Base de datos de clientes del taller.")
    # [AQUÍ VA TU CÓDIGO DE CLIENTES]
    # Ejemplo:
    # if clientes:
    #     st.table(clientes)
    # else:
    #     st.info("No hay clientes registrados aún.")

# --- PESTAÑA 8: AJUSTES ---
elif st.session_state.pagina == "Ajustes":
    st.header("⚙️ Ajustes de la Aplicación")
    st.write("Configuraciones generales del sistema.")
    st.write("- Base de datos: SQLite local")
    st.write("- Tema: Rojo / Negro")
    st.write("- IA: Gemini API")
    
    if st.button("Cerrar Conexión Base de Datos"):
        # conn.close()  # Descomenta si tienes la variable global 'conn'
        st.success("Conexión cerrada.")
