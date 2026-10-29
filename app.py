import streamlit as st
import urllib.parse # Para el botón de WhatsApp
import datetime # Para la fecha del albarán
import os # Para comprobar si existe el logo
from fpdf import FPDF # Para generar el PDF
import db # Nuestro gestor de base de datos SQLite
import pandas as pd # Para manejar datos de la BD

# ==========================================
# FUNCIÓN PARA GENERAR EL PDF
# ==========================================
def crear_pdf_albaran(cliente, matricula, trabajo, fecha):
    pdf = FPDF()
    pdf.add_page()
    
    # --- CABECERA CON LOGO CENTRADO ---
    logo_path = "logo_transparent.png"
    logo_w = 35 # Ancho del logo en mm
    if os.path.exists(logo_path):
        x_logo = (210 - logo_w) / 2
        pdf.image(logo_path, x=x_logo, y=10, w=logo_w) 
    
    pdf.set_y(45) 
    
    # --- TEXTO DE CABECERA CENTRADO ---
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 8, "TALLER OLIVARES 4T", align="C", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("Helvetica", "", 10)
    pdf.cell(0, 5, "C/ Ejemplo, 123 - Tel: 600 000 000", align="C", new_x="LMARGIN", new_y="NEXT")
    pdf.cell(0, 5, "NIF: B-12345678", align="C", new_x="LMARGIN", new_y="NEXT")
    
    # --- LÍNEA SEPARADORA ---
    pdf.ln(10)
    pdf.set_draw_color(200, 200, 200)
    pdf.line(10, pdf.get_y(), 200, pdf.get_y())
    pdf.ln(10)
    
    # --- DATOS DEL ALBARÁN ---
    pdf.set_font("Helvetica", "B", 14)
    pdf.cell(0, 10, f"ALBARÁN Nº: 0001/{datetime.date.today().year}", new_x="LMARGIN", new_y="NEXT")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 6, f"Fecha: {fecha}", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(5)
    
    # --- DATOS DEL CLIENTE ---
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(30, 7, "Cliente:")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, cliente if cliente else "-", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(30, 7, "Matrícula:")
    pdf.set_font("Helvetica", "", 11)
    pdf.cell(0, 7, matricula if matricula else "-", new_x="LMARGIN", new_y="NEXT")
    pdf.ln(10)
    
    # --- DETALLES DEL TRABAJO ---
    pdf.set_font("Helvetica", "B", 11)
    pdf.cell(0, 8, "DETALLES DEL TRABAJO:", new_x="LMARGIN", new_y="NEXT")
    
    pdf.set_fill_color(245, 245, 245)
    pdf.multi_cell(0, 10, trabajo if trabajo else "Sin detalles especificados.", border=1, fill=True)
    
    pdf.ln(20)
    
    # --- PIE DE PÁGINA ---
    pdf.set_font("Helvetica", "I", 9)
    pdf.set_text_color(150, 150, 150)
    pdf.cell(0, 10, "* Documento no válido como factura. Generado por Taller Olivares 4T.", align="C", new_x="LMARGIN", new_y="NEXT")
    
    return bytes(pdf.output())

# ==========================================
# 1. CONFIGURACIÓN Y LOGO
# ==========================================
st.set_page_config(page_title="Taller Olivares 4T", layout="wide", page_icon="🏍️")
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
    st.button("📅\nAvisos / Tareas", use_container_width=True, 
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

# --- PESTAÑA 2: AVISOS / TAREAS ---
elif st.session_state.pagina == "Avisos":
    st.header("📅 Avisos y Tareas Pendientes")
    st.write("Programa futuros trabajos para tus clientes y no dejes que se te pase ninguno.")
    
    df_clientes = db.obtener_clientes()
    
    if df_clientes.empty:
        st.warning("⚠️ No hay clientes en la base de datos. Añade uno primero en la pestaña 'Clientes'.")
    else:
        tab_avisos1, tab_avisos2 = st.tabs(["➕ Crear Aviso", "📋 Ver Avisos"])
        
        with tab_avisos1:
            with st.form("form_nuevo_aviso", clear_on_submit=True):
                st.subheader("Programar un aviso")
                
                opciones_clientes = {
                    f"{row['nombre']} - {row['matricula'] if pd.notna(row['matricula']) else 'Sin matrícula'}": row['id'] 
                    for _, row in df_clientes.iterrows()
                }
                
                cliente_aviso = st.selectbox("Selecciona el Cliente", options=list(opciones_clientes.keys()))
                
                col_f1, col_f2 = st.columns(2)
                with col_f1:
                    fecha_aviso = st.date_input("Fecha del aviso")
                with col_f2:
                    motivo = st.text_input("Motivo (Ej: ITV, Cambio de neumáticos, Revisión)")
                
                submit_aviso = st.form_submit_button("Guardar Aviso", type="primary")
                
                if submit_aviso:
                    if motivo.strip() == "":
                        st.error("El motivo es obligatorio.")
                    else:
                        id_aviso_cliente = opciones_clientes[cliente_aviso]
                        fecha_str = fecha_aviso.strftime("%Y-%m-%d") 
                        db.añadir_aviso(id_aviso_cliente, fecha_str, motivo)
                        st.success("¡Aviso guardado correctamente!")
        
        with tab_avisos2:
            st.subheader("Lista de Avisos")
            df_avisos = db.obtener_avisos()
            
            if not df_avisos.empty:
                df_avisos['fecha'] = pd.to_datetime(df_avisos['fecha']).dt.strftime('%d/%m/%Y')
                
                st.dataframe(df_avisos, use_container_width=True, hide_index=True)
                
                st.divider()
                st.subheader("Gestionar un aviso")
                
                opciones_avisos = {
                    f"{row['id']} - {row['nombre']} ({row['fecha']} - {row['motivo']})": row['id']
                    for _, row in df_avisos.iterrows()
                }
                
                aviso_seleccionado = st.selectbox("Selecciona un aviso para gestionar:", options=list(opciones_avisos.keys()))
                
                if aviso_seleccionado:
                    id_aviso = opciones_avisos[aviso_seleccionado]
                    datos_aviso = df_avisos[df_avisos['id'] == id_aviso].iloc[0]
                    
                    nombre_cliente_aviso = datos_aviso['nombre']
                    mensaje_wa = f"Hola {nombre_cliente_aviso}, te escribimos de Taller Olivares 4T. Te recordamos que tienes pendiente: {datos_aviso['motivo']}. ¿Quieres que te pidamos cita? 🏍️"
                    mensaje_url = urllib.parse.quote(mensaje_wa)
                    url_whatsapp = f"https://wa.me/?text={mensaje_url}" 
                    
                    col_g1, col_g2 = st.columns(2)
                    with col_g1:
                        st.link_button("📱 Enviar Recordatorio WhatsApp", url_whatsapp, use_container_width=True)
                    with col_g2:
                        if st.button("✅ Marcar como Completado (Borrar)", type="primary", use_container_width=True):
                            db.eliminar_aviso(id_aviso)
                            st.success("Aviso marcado como completado y eliminado de la lista.")
                            st.rerun()
            else:
                st.info("No hay avisos pendientes. ¡Todo al día!")

# --- PESTAÑA 3: PIEZAS ---
elif st.session_state.pagina == "Piezas":
    st.header("🔧 Buscador de Piezas y Precios")
    st.write("Consulta disponibilidad y precios de repuestos. (Datos de ejemplo)")

    piezas_db = [
        {"Ref": "ACE-10W40", "Pieza": "Aceite Motor 10W40 (1L)", "Marca": "Castrol", "Precio (€)": 12.50, "Stock": 15},
        {"Ref": "PAS-BRE-01", "Pieza": "Pastillas de freno delanteras", "Marca": "Brembo", "Precio (€)": 35.00, "Stock": 8},
        {"Ref": "NEU-MIC-17", "Pieza": "Neumático 120/70-17", "Marca": "Michelin", "Precio (€)": 110.00, "Stock": 4},
        {"Ref": "BUJ-NGK-01", "Pieza": "Bujía de iridio", "Marca": "NGK", "Precio (€)": 18.20, "Stock": 20},
        {"Ref": "CAD-DID-520", "Pieza": "Kit de arrastre 520", "Marca": "DID", "Precio (€)": 85.00, "Stock": 3},
        {"Ref": "FIL-AIR-02", "Pieza": "Filtro de aire", "Marca": "K&N", "Precio (€)": 45.00, "Stock": 6},
    ]

    busqueda = st.text_input("🔍 Buscar por pieza, marca o referencia:", placeholder="Ej: Pastillas, Michelin, o NEU-MIC-17")

    if busqueda:
        resultados = [
            p for p in piezas_db 
            if busqueda.lower() in p["Pieza"].lower() 
            or busqueda.lower() in p["Marca"].lower() 
            or busqueda.lower() in p["Ref"].lower()
        ]
    else:
        resultados = piezas_db

    st.divider()
    if resultados:
        st.dataframe(
            resultados, 
            use_container_width=True, 
            hide_index=True,
            column_config={
                "Precio (€)": st.column_config.NumberColumn(format="%.2f €"),
                "Stock": st.column_config.ProgressColumn(min_value=0, max_value=25)
            }
        )
    else:
        st.warning("No se encontraron piezas que coincidan con la búsqueda. Prueba con otra palabra.")

# --- PESTAÑA 4: PDFs / ALBARANES ---
elif st.session_state.pagina == "PDFs":
    st.header("📄 Generación de Albaranes")
    st.write("Crea el albarán del trabajo. Si el cliente tenía una tarea pendiente, se rellenará sola.")
    
    df_clientes = db.obtener_clientes()
    
    if df_clientes.empty:
        st.warning("⚠️ No hay clientes en la base de datos. Ve a la pestaña de 'Clientes' y añade al menos uno.")
    else:
        # SELECCIÓN DE CLIENTE (Fuera del formulario para que actualice el texto dinámicamente)
        opciones_clientes = {
            f"{row['nombre']} - {row['matricula'] if pd.notna(row['matricula']) else 'Sin matrícula'}": row['id'] 
            for _, row in df_clientes.iterrows()
        }
        
        cliente_seleccionado = st.selectbox("1. Selecciona el Cliente", options=list(opciones_clientes.keys()))
        
        id_cliente = opciones_clientes[cliente_seleccionado]
        datos_cliente = df_clientes[df_clientes['id'] == id_cliente].iloc[0]
        
        nombre_cliente = datos_cliente['nombre']
        matricula_cliente = datos_cliente['matricula'] if pd.notna(datos_cliente['matricula']) else ""
        
        # BUSCAMOS AVISOS PENDIENTES DE ESTE CLIENTE
        df_avisos_pendientes = db.obtener_avisos_cliente(id_cliente)
        texto_autocompletado = ""
        aviso_a_borrar = None
        
        if not df_avisos_pendientes.empty:
            st.info("💡 Este cliente tiene tareas pendientes de la pestaña de Avisos.")
            opciones_pendientes = {row['motivo']: row['id'] for _, row in df_avisos_pendientes.iterrows()}
            
            # Desplegable para elegir cuál de las tareas pendientes estamos haciendo
            tarea_seleccionada = st.selectbox(
                "2. ¿Qué trabajo vas a realizar hoy?", 
                options=["-- Selecciona una tarea para autocompletar --"] + list(opciones_pendientes.keys())
            )
            
            if tarea_seleccionada != "-- Selecciona una tarea para autocompletar --":
                texto_autocompletado = tarea_seleccionada
                aviso_a_borrar = opciones_pendientes[tarea_seleccionada]
        
        # FORMULARIO DEL ALBARÁN
        with st.form("form_albaran"):
            st.subheader("3. Detalles del trabajo")
            
            # El texto se rellena solo si has elegido una tarea pendiente arriba
            trabajo = st.text_area("Trabajo realizado / Piezas", value=texto_autocompletado)
            
            submit = st.form_submit_button("Generar Albarán", type="primary")
        
        if submit:
            fecha_hoy = datetime.date.today().strftime("%d/%m/%Y")
            
            # 1. Guardamos el trabajo en el historial de la moto
            db.guardar_historial(id_cliente, fecha_hoy, trabajo)
            
            # 2. Si venía de un aviso, lo eliminamos de pendientes
            if aviso_a_borrar:
                db.eliminar_aviso(aviso_a_borrar)
            
            # 3. Generamos los documentos
            texto_albaran = f"""
            =========================================
                    TALLER OLIVARES 4T
            =========================================
            Fecha: {fecha_hoy}
            Cliente: {nombre_cliente}
            Matrícula: {matricula_cliente}
            -----------------------------------------
            DETALLES DEL TRABAJO:
            {trabajo}
            -----------------------------------------
            * Documento no válido como factura.
            """
            
            pdf_bytes = crear_pdf_albaran(nombre_cliente, matricula_cliente, trabajo, fecha_hoy)
            
            st.success(f"¡Albarán generado! El trabajo se ha guardado en el historial y la tarea pendiente ha sido eliminada.")
            
            col_desc1, col_desc2 = st.columns(2)
            
            with col_desc1:
                st.download_button(
                    label="⬇️ Descargar en PDF",
                    data=pdf_bytes,
                    file_name=f"Albaran_{nombre_cliente}_{fecha_hoy.replace('/','-')}.pdf",
                    mime="application/pdf",
                    type="primary",
                    use_container_width=True
                )
                
            with col_desc2:
                st.download_button(
                    label="⬇️ Descargar en TXT",
                    data=texto_albaran,
                    file_name=f"Albaran_{nombre_cliente}_{fecha_hoy.replace('/','-')}.txt",
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
    st.header("👥 Gestión de Clientes")
    st.write("Añade, consulta, modifica o revisa el historial de los clientes.")

    tab1, tab2, tab3, tab4 = st.tabs(["➕ Añadir", "📋 Ver", "✏️ Modificar", "📜 Historial"])

    with tab1:
        with st.form("form_nuevo_cliente", clear_on_submit=True):
            st.subheader("Datos del nuevo cliente")
            nombre = st.text_input("Nombre completo *")
            telefono = st.text_input("Teléfono")
            matricula = st.text_input("Matrícula")
            moto = st.text_input("Modelo de Moto")
            
            st.markdown("---")
            st.markdown("**¿Tiene algún trabajo pendiente? (Opcional)**")
            trabajo_pendiente = st.text_area("Trabajo a realizar (Ej: Cambio de aceite, ITV...)")
            fecha_trabajo = st.date_input("Fecha prevista para el trabajo")
            
            submit_cliente = st.form_submit_button("Guardar Cliente", type="primary")
            
            if submit_cliente:
                if nombre.strip() == "":
                    st.error("El nombre es obligatorio.")
                else:
                    nuevo_id = db.añadir_cliente(nombre, telefono, matricula, moto)
                    
                    if trabajo_pendiente.strip() != "":
                        fecha_str = fecha_trabajo.strftime("%Y-%m-%d")
                        db.añadir_aviso(nuevo_id, fecha_str, trabajo_pendiente)
                        st.success(f"¡Cliente '{nombre}' guardado y aviso programado!")
                    else:
                        st.success(f"¡Cliente '{nombre}' guardado con éxito!")

    with tab2:
        st.subheader("Listado de clientes")
        df_clientes = db.obtener_clientes()
        
        if not df_clientes.empty:
            st.dataframe(df_clientes, use_container_width=True, hide_index=True)
        else:
            st.info("Todavía no hay clientes en la base de datos. ¡Añade el primero!")

    with tab3:
        st.subheader("Editar un cliente existente")
        df_clientes = db.obtener_clientes()
        
        if not df_clientes.empty:
            opciones = {
                f"{row['id']} - {row['nombre']} ({row['matricula'] if row['matricula'] else 'Sin matrícula'})": row['id'] 
                for _, row in df_clientes.iterrows()
            }
            
            cliente_seleccionado = st.selectbox("Selecciona el cliente a modificar:", options=list(opciones.keys()))
            
            if cliente_seleccionado:
                id_sel = opciones[cliente_seleccionado]
                datos_actuales = df_clientes[df_clientes['id'] == id_sel].iloc[0]
                
                with st.form("form_editar_cliente"):
                    edit_nombre = st.text_input("Nombre completo *", value=datos_actuales['nombre'])
                    
                    val_tel = datos_actuales['telefono'] if pd.notna(datos_actuales['telefono']) else ""
                    val_mat = datos_actuales['matricula'] if pd.notna(datos_actuales['matricula']) else ""
                    val_mot = datos_actuales['moto'] if pd.notna(datos_actuales['moto']) else ""
                    
                    edit_telefono = st.text_input("Teléfono", value=val_tel)
                    edit_matricula = st.text_input("Matrícula", value=val_mat)
                    edit_moto = st.text_input("Modelo de Moto", value=val_mot)
                    
                    submit_edit = st.form_submit_button("Actualizar Cliente", type="primary")
                    
                    if submit_edit:
                        if edit_nombre.strip() == "":
                            st.error("El nombre es obligatorio.")
                        else:
                            db.actualizar_cliente(id_sel, edit_nombre, edit_telefono, edit_matricula, edit_moto)
                            st.success(f"¡Cliente '{edit_nombre}' actualizado con éxito!")
                            st.rerun() 
        else:
            st.info("No hay clientes registrados para modificar.")

    with tab4:
        st.subheader("Historial de la Moto")
        df_clientes = db.obtener_clientes()
        
        if not df_clientes.empty:
            opciones_hist = {
                f"{row['nombre']} - {row['moto'] if pd.notna(row['moto']) else 'Moto'} ({row['matricula'] if pd.notna(row['matricula']) else 'Sin matrícula'})": row['id'] 
                for _, row in df_clientes.iterrows()
            }
            
            cliente_hist = st.selectbox("Selecciona una moto para ver su historial:", options=list(opciones_hist.keys()))
            id_hist = opciones_hist[cliente_hist]
            
            df_historial = db.obtener_historial(id_hist)
            
            if not df_historial.empty:
                st.dataframe(df_historial, use_container_width=True, hide_index=True)
            else:
                st.info("Esta moto todavía no tiene trabajos registrados. Genera un albarán para empezar su historial.")
        else:
            st.info("No hay clientes en la base de datos.")

# --- PESTAÑA 8: AJUSTES ---
elif st.session_state.pagina == "Ajustes":
    st.header("⚙️ Ajustes de la Aplicación")
    st.write("Configuraciones generales del sistema.")
    st.write("- Base de datos: SQLite local")
    st.write("- Tema: Rojo / Negro")
    st.write("- IA: Gemini API")
    
    if st.button("Cerrar Conexión Base de Datos"):
        st.success("Conexión cerrada.")
