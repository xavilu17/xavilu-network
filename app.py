import streamlit as st
import urllib.parse # Para el botón de WhatsApp
import datetime # Para la fecha del albarán
import os # Para comprobar si existe el logo
from fpdf import FPDF # Para generar el PDF

# ==========================================
# FUNCIÓN PARA GENERAR EL PDF
# ==========================================
def crear_pdf_albaran(cliente, matricula, trabajo, fecha):
    pdf = FPDF()
    pdf.add_page()
    
    # --- CABECERA CON LOGO ---
    logo_path = "logo_transparent.png"
    if os.path.exists(logo_path):
        # Ajusta el ancho (w) y posición (x, y) según tu logo
        pdf.image(logo_path, x=10, y=10, w=35) 
    
    # Texto de la cabecera (a la derecha del logo)
    pdf.set_font("Helvetica", "B", 16)
    pdf.set_xy(50, 12)
    pdf.cell(0, 8, "TALLER OLIVARES 4T", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.set_x(50)
    pdf.cell(0, 5, "C/ Ejemplo, 123 - Tel: 600 000 000", new_x="LMARGIN", new_y="NEXT")
    
    pdf.ln(15) # Espacio debajo de la cabecera
    
    # --- DATOS DEL ALBARÁN ---
    pdf.set_font("Helvetica", "B", 12)
    pdf.cell(0, 10, f"ALBARÁN - Fecha: {fecha}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(2)
    
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 8, f"Cliente: {cliente}", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 8, f"Matrícula: {matricula}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # --- DETALLES DEL TRABAJO ---
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "Detalles del trabajo:", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("Helvetica", "", 11)
    # multi_cell permite escribir texto largo y que salte de línea solo
    pdf.multi_cell(0, 8, trabajo) 
    
    pdf.ln(15)
    
    # --- PIE DE PÁGINA ---
    pdf.set_font("Helvetica", "I", 9)
    pdf.cell(0, 10, "* Documento no válido como factura.", new_x="LMARGIN", new_y="NEXT")
    
    # Devolvemos el PDF en bytes para que Streamlit lo pueda descargar
    return bytes(pdf.output())

# ==========================================
# 1. CONFIGURACIÓN Y LOGO
# ==========================================
st.set_page_config(page_title="Taller Olivares 4T", layout="wide", page_icon="🏍️")

# LOGO TRANSPARENTE
st.logo("logo_transparent.png", size="large")

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

# --- PESTAÑA 2: AVISOS ITV ---
elif st.session_state.pagina == "Avisos":
    st.header("📅 Avisos ITV")
    st.write("Gestión de avisos y recordatorios de ITV.")
    
    st.subheader("Ejemplo de Aviso por WhatsApp")
    tel_ejemplo = "34600000000" 
    nombre_ejemplo = "Cliente Prueba"
    moto_ejemplo = "Yamaha T-Max"
    
    mensaje = f"Hola {nombre_ejemplo}, te escribimos de Taller Olivares 4T. Tu {moto_ejemplo} tiene la ITV próxima a caducar. ¿Quieres que te pidamos cita? 🏍️"
    mensaje_url = urllib.parse.quote(mensaje)
    url_whatsapp = f"https://wa.me/{tel_ejemplo}?text={mensaje_url}"
    
    st.link_button("📱 Avisar ITV por WhatsApp", url_whatsapp, type="primary")

# --- PESTAÑA 3: PIEZAS ---
elif st.session_state.pagina == "Piezas":
    st.header("🔧 Gestión de Piezas")
    st.write("Control y búsqueda de repuestos.")

# --- PESTAÑA 4: PDFs / ALBARANES ---
elif st.session_state.pagina == "PDFs":
    st.header("📄 Generación de Albaranes")
    st.write("Crea y descarga el albarán del trabajo realizado.")
    
    # Formulario para recoger datos
    with st.form("form_albaran"):
        cliente = st.text_input("Nombre del Cliente")
        matricula = st.text_input("Matrícula de la Moto")
        trabajo = st.text_area("Trabajo realizado / Piezas")
        submit = st.form_submit_button("Generar Albarán")
    
    # El botón de descarga va FUERA del formulario
    if submit:
        fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")
        
        # Generamos el texto para TXT
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
        
        # Generamos el PDF pasando los datos
        pdf_bytes = crear_pdf_albaran(cliente, matricula, trabajo, fecha_hoy)
        
        st.success("¡Albarán generado con éxito! Elige el formato de descarga:")
        
        # Colocamos los botones en dos columnas
        col_desc1, col_desc2 = st.columns(2)
        
        with col_desc1:
            st.download_button(
                label="⬇️ Descargar en PDF",
                data=pdf_bytes,
                file_name=f"Albaran_{cliente}_{fecha_hoy.replace('/','-')}.pdf",
                mime="application/pdf",
                type="primary",
                use_container_width=True
            )
            
        with col_desc2:
            st.download_button(
                label="⬇️ Descargar en TXT",
                data=texto_albaran,
                file_name=f"Albaran_{cliente}_{fecha_hoy.replace('/','-')}.txt",
                mime="text/plain",
                use_container_width=True
            )

# --- PESTAÑA 5: INVENTARIO ---
elif st.session_state.pagina == "Inventario":
    st.header("📦 Inventario")
    st.write("Gestión de stock del taller.")

# --- PESTAÑA 6: VOZ/IA ---
elif st.session_state.pagina == "Voz/IA":
    st.header("🎙️ Asistente Voz/IA")
    st.write("Módulo de inteligencia artificial y dictado por voz.")

# --- PESTAÑA 7: CLIENTES ---
elif st.session_state.pagina == "Clientes":
    st.header("👥 Clientes")
    st.write("Base de datos de clientes del taller.")

# --- PESTAÑA 8: AJUSTES ---
elif st.session_state.pagina == "Ajustes":
    st.header("⚙️ Ajustes de la Aplicación")
    st.write("Configuraciones generales del sistema.")
    st.write("- Base de datos: SQLite local")
    st.write("- Tema: Rojo / Negro")
    st.write("- IA: Gemini API")
    
    if st.button("Cerrar Conexión Base de Datos"):
        st.success("Conexión cerrada.")
