import streamlit as st
import ollama
import asyncio
import edge_tts
import pypdf
from docx import Document
import openpyxl
from openpyxl.styles import Font
from pptx import Presentation
from io import BytesIO
from PIL import Image
from fpdf import FPDF
import urllib.parse
import urllib.request
import csv
import os
import base64
import pandas as pd
import plotly.express as px
from youtube_transcript_api import YouTubeTranscriptApi
import re
import requests
import tempfile
import time
import qrcode
import gc
import html
import numpy as np

# Intentar importar OpenCV para el lector de QR
try:
    import cv2
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

# --- LIBRERÍAS EXTERNAS ---
try:
    import yt_dlp
    HAS_YTDLP = True
except ImportError:
    HAS_YTDLP = False

try:
    from ddgs import DDGS
except ImportError:
    try:
        from duckduckgo_search import DDGS
    except ImportError:
        DDGS = None

try:
    from google import genai
    from google.genai import types
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

# --- CONFIGURACIÓN ---
st.set_page_config(page_title="Xavilu Network Pro", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")
MAX_FILE_MB = 50
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-3.8-flash"

# --- ESTADO DE SESIÓN ---
_defaults = {
    "chat_history": [], "img_url": "", "app_code": "", "pdf_chat_history": [],
    "mermaid_code": "", "analisis_ia": "", "dl_bytes": None, "dl_filename": "", "dl_mime": "",
    "cuento_generado": "", "receta_generada": "", "menu_generado": "",
    "redes_resultado": "", "pelo_resultado": "", "deberes_resultado": "",
    "qr_resultado": "", "code_resultado": "", "entrevista_resultado": "",
    "presupuesto_resultado": "", "noticias_deportes": [], "resumen_deportivo": "",
    "evento_resultados": None, "office_text": "", "pdf_text": "", "clip_text": "",
    "dep_resultados": "", "partido_analisis": "", "plan_deporte": "",
    "cine_rec": "", "tv_prog": "", "entradas": "", "trivia": "", "yt_text": "", "web_text": "",
    "letra_cancion": "", "titulo_cancion": "", "portada_cancion": "",
    "doc_opcion": "Resumen Ejecutivo",
    "global_context": "", "global_file_name": "",
    "agente_activo": None, "agente_chat_history": []
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# --- CSS CYBERPUNK INTUITIVO ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
    .stApp { background-color: #050505 !important; background-image: radial-gradient(circle at 50% 50%, #1a1a1a 0%, #050505 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
    h1, h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #f7dfa3 !important; text-shadow: 0 0 10px #d4af37, 0 0 20px #d4af37 !important; text-align: center; letter-spacing: 2px; }
    [data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; text-transform: uppercase !important; background: rgba(0, 240, 255, 0.1) !important; padding: 5px 10px !important; border-left: 4px solid #00f0ff !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 8px !important; }
    p, .stMarkdown p { color: #d4af37 !important; font-size: 16px; }
    [data-testid="stSidebar"] { background-color: rgba(10, 10, 15, 0.9) !important; border-right: 1px solid #d4af37 !important; }
    button[data-baseweb="tab"] { background: rgba(20, 20, 25, 0.8) !important; border: 1px solid #d4af37 !important; border-radius: 25px !important; color: #d4af37 !important; font-family: 'Orbitron', sans-serif !important; font-size: 12px !important; padding: 10px 15px !important; margin: 5px !important; transition: all 0.3s; }
    button[data-baseweb="tab"]:hover { background: rgba(212, 175, 55, 0.2) !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background: linear-gradient(135deg, #d4af37 0%, #b8860b 100%) !important; color: #000 !important; border: 1px solid #fff !important; box-shadow: 0 0 15px #d4af37 !important; }
    .stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { background-color: rgba(0, 0, 0, 0.8) !important; border: 1px solid #d4af37 !important; border-radius: 8px !important; color: #f7dfa3 !important; font-size: 16px !important; }
    div.stButton > button, div[data-testid="stDownloadButton"] > button { background: rgba(0, 0, 0, 0.8) !important; color: #d4af37 !important; border: 1px solid #d4af37 !important; border-radius: 25px !important; font-family: 'Orbitron', sans-serif !important; font-weight: bold !important; padding: 10px 20px !important; }
    div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover { background: #d4af37 !important; color: #000 !important; transform: translateY(-2px); box-shadow: 0 0 15px #d4af37 !important; }
    [data-testid="stChatMessage"] { background: rgba(10, 10, 15, 0.7) !important; border: 1px solid rgba(212, 175, 55, 0.4) !important; border-radius: 15px !important; padding: 10px; }
    .stAlert { border-radius: 10px; font-weight: bold; }
    </style>
    """, unsafe_allow_html=True)

# --- FUNCIONES BASE ---
def buscar_en_internet(query):
    if DDGS is None: return []
    try:
        res = DDGS().text(query, region='es-es', max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in res]
    except: return []

def buscar_noticias_deportivas(deporte, liga):
    if DDGS is None: return []
    try:
        res = DDGS().news(f"últimas noticias {deporte} {liga} hoy", max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("url", ""), "resumen": r.get("body", ""), "fecha": r.get("date", "")} for r in res]
    except: return []

def leer_archivo_universal(upload, modelo_ia):
    nombre = upload.name.lower()
    if nombre.endswith(".txt"): return upload.read().decode("utf-8", errors="ignore")
    if nombre.endswith(".pdf"): return "\n".join([p.extract_text() for p in pypdf.PdfReader(BytesIO(upload.getvalue())).pages if p.extract_text()])
    if nombre.endswith(".docx"): return "\n".join([p.text for p in Document(upload).paragraphs if p.text])
    if nombre.endswith(".xlsx"):
        wb = openpyxl.load_workbook(upload, data_only=True)
        return "\n".join([" | ".join([str(c) for c in f if c is not None]) for h in wb.sheetnames for f in wb[h].iter_rows(values_only=True) if any(c is not None for c in f)])
    return "Formato no reconocido."

@st.cache_resource
def get_gemini_client(api_key): return genai.Client(api_key=api_key)

def analizar_archivo_con_gemini(archivo_bytes, mime_type, prompt, api_key_gemini):
    if not HAS_GEMINI: return "[ERROR] Falta google-genai."
    for i in range(3):
        try:
            client = get_gemini_client(api_key_gemini)
            parte_archivo = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
            response = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, parte_archivo])
            return response.text
        except Exception as e:
            if "503" in str(e) and i < 2: time.sleep(3); continue
            else: return f"[ERROR GEMINI] {e}"

def llamar_ia(modelo, prompt, api_key_gemini=""):
    prompt_final = prompt[:MAX_TEXTO_IA]
    
    # 🧠 INYECCIÓN DEL CONTEXTO GLOBAL
    if st.session_state.get("global_context"):
        nombre_archivo = st.session_state.get("global_file_name", "Archivo_Adjunto")
        prompt_final = f"[EL USUARIO HA ADJUNTADO EL SIGUIENTE ARCHIVO: {nombre_archivo}]\nCONTENIDO EXTRAÍDO DEL ARCHIVO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN DEL USUARIO]:\n{prompt_final}\n\n(Responde teniendo en cuenta el archivo adjunto si es relevante)."

    if api_key_gemini.strip():
        if not HAS_GEMINI: return "[ERROR] Falta google-genai."
        for intento in range(3):
            try:
                client = get_gemini_client(api_key_gemini)
                response = client.models.generate_content(model=GEMINI_MODEL, contents=prompt_final)
                return response.text
            except Exception as e:
                if "503" in str(e) and intento < 2: time.sleep(3); continue
                return f"[ERROR GEMINI] {e}"
    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=300)
        response = client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})
        return response["message"]["content"].split("</think>")[-1].strip()
    except Exception as e: return f"[ERROR LOCAL] {e}"

def generar_pdf_pro(titulo, contenido):
    pdf = FPDF(); pdf.add_page()
    font_path = "C:/Windows/Fonts/arial.ttf"
    if os.path.exists(font_path): pdf.add_font("Arial", "", font_path, uni=True); pdf.set_font("Arial", size=11)
    else: pdf.set_font("Helvetica", size=11); contenido = contenido.encode('latin-1', 'replace').decode('latin-1')
    pdf.set_font("Helvetica", "B", 24); pdf.cell(0, 20, txt="", ln=True); pdf.multi_cell(0, 10, txt=titulo, align='C')
    pdf.set_font("Helvetica", "I", 12); pdf.cell(0, 10, txt="Generado por Xavilu Network Pro", align='C', ln=True); pdf.add_page()
    if os.path.exists(font_path): pdf.set_font("Arial", size=11)
    else: pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 8, txt=contenido)
    return bytes(pdf.output())

def generar_word(texto):
    doc = Document(); doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); out.seek(0); return out.getvalue()

def generar_excel(texto_csv):
    wb = openpyxl.Workbook(); ws = wb.active
    for row_idx, row in enumerate(csv.reader(texto_csv.replace("```csv", "").replace("```", "").strip().split('\n')), 1):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1: cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

def generar_pptx(titulo, texto):
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[0]); slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1]); lineas = bloque.split("\n"); slide.shapes.title.text = lineas[0][:50]
        if len(slide.placeholders) > 1: slide.placeholders[1].text = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
    out = BytesIO(); prs.save(out); out.seek(0); return out.getvalue()

def botones_descarga(texto, nombre_por_defecto="documento"):
    if texto:
        st.markdown("---")
        key_unica = f"{nombre_por_defecto}".replace(" ", "_")
        
        nombre_elegido = st.text_input(
            "📝 Ponle nombre a tu archivo (sin extensión):", 
            value=nombre_por_defecto, 
            key=f"input_name_{key_unica}",
            help="Escribe cómo quieres que se llame el archivo antes de descargarlo."
        )
        
        nombre_limpio = re.sub(r'[^a-zA-Z0-9_\- ]', '', nombre_elegido).strip().replace(' ', '_')
        if not nombre_limpio: nombre_limpio = nombre_por_defecto

        c1, c2 = st.columns(2)
        with c1: 
            st.download_button("📥 Descargar (TXT)", texto.encode('utf-8'), f"{nombre_limpio}.txt", "text/plain", key=f"dl_txt_{key_unica}", use_container_width=True)
        with c2: 
            st.download_button("📄 Descargar (PDF)", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"dl_pdf_{key_unica}", use_container_width=True)

def generar_audio_sincrono(texto, voz, exp=None):
    loop = asyncio.new_event_loop(); asyncio.set_event_loop(loop)
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(texto, voz).stream():
            if chunk["type"] == "audio": buf.write(chunk["data"])
        buf.seek(0); return buf.getvalue()
    try: return loop.run_until_complete(_gen())
    except: return None
    finally: loop.close()

def boton_superpoder(texto_base, instruccion_poder, key_suffix, modelo, api_key):
    if st.button(f"🚀 Activar Superpoder: {instruccion_poder}", key=f"super_{key_suffix}", type="secondary"):
        if not texto_base or len(str(texto_base)) < 10:
            st.warning("⚠️ Necesitas generar o subir algo primero para usar los superpoderes.")
            return
        with st.spinner("⚡ Canalizando energía de IA..."):
            prompt = f"Actúa como un experto de élite. {instruccion_poder}. Basándote en este contenido:\n\n{texto_base}"
            resultado = llamar_ia(modelo, prompt, api_key)
            st.success("¡Superpoder activado!")
            st.info(resultado)
            botones_descarga(resultado, f"Superpoder_{key_suffix}")

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.markdown("<h1>XAVILU NETWORK</h1><p style='text-align:center; color:#00f0ff;'>● SISTEMA CONECTADO</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🪐 PANEL DE CONTROL")
    modelo_seleccionado = st.selectbox("Motor Local (Ollama)", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])
    try: api_key_gemini = st.secrets["GEMINI_API_KEY"]; st.success("🔑 Gemini API OK")
    except: api_key_gemini = st.text_input("Gemini API Key", type="password")
    
    st.markdown("---")
    # 🧠 CEREBRO GLOBAL
    st.markdown("### 🧠 Cerebro Global")
    st.caption("Sube cualquier archivo (foto, audio, pdf). La IA lo leerá y lo usará en TODAS las pestañas.")
    archivo_global = st.file_uploader("📁 Contexto Global", type=["pdf", "txt", "docx", "png", "jpg", "jpeg", "mp3", "wav", "mp4"], key="up_global")
    
    if archivo_global:
        if st.session_state.get("global_file_name") != archivo_global.name:
            with st.spinner("Procesando archivo para el cerebro global..."):
                nombre = archivo_global.name.lower()
                if nombre.endswith((".png", ".jpg", ".jpeg")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "image/jpeg", "Describe esta imagen con todo detalle (qué es, texto visible, etc.) para que otra IA sepa de qué trata.", api_key_gemini)
                elif nombre.endswith((".mp3", ".wav")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "audio/mp3", "Transcribe este audio.", api_key_gemini)
                elif nombre.endswith((".mp4")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "video/mp4", "Resume el contenido de este vídeo.", api_key_gemini)
                else:
                    ctx = leer_archivo_universal(archivo_global, modelo_seleccionado)
                
                st.session_state.global_context = ctx[:5000] # Limitamos a 5000 para no saturar
                st.session_state.global_file_name = archivo_global.name
                st.success("✅ Archivo memorizado.")
                
    if st.session_state.get("global_file_name"):
        st.info(f"Archivo Activo: **{st.session_state.global_file_name}**")
        if st.button("🗑️ Olvidar Archivo"):
            st.session_state.global_context = ""
            st.session_state.global_file_name = ""
            st.rerun()

    st.markdown("---")
    st.caption("📁 Informes guardados en: Escritorio/Informes_Xavilu")
    if not HAS_YTDLP: st.warning("⚠️ `yt-dlp` no instalado")
    if not HAS_CV2: st.warning("⚠️ `opencv-python` no instalado")

macro_tabs = st.tabs([
    "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA",
    "🤖 ASISTENTES IA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR",
    "🎵 MÚSICA Y ARTE"
])

# ==========================================
# CATEGORÍA 1: DOCUMENTOS Y TEXTO
# ==========================================
with macro_tabs[0]:
    sub = st.tabs(["📁 ANÁLISIS", "📊 OFFICE", "📄 PDF", "🕸️ MAPAS", "📇 QR & CV", "📋 PORTAPAPELES"])
    with sub[0]:
        st.markdown("#### Análisis Avanzado de Documentos")
        
        with st.expander("💡 Cajón de Ideas (¿Qué puedo hacer con mi archivo?)"):
            st.write("- **Auditoría:** Busca contradicciones o letra pequeña.")
            st.write("- **Estudio:** Convierte el texto en un test de preguntas.")
            st.write("- **Datos:** Saca solo las fechas, nombres y cifras.")
            st.write("- **Traducción:** Pasa todo el documento a otro idioma.")
        
        archivo = st.file_uploader("Sube un PDF, Word, Excel o TXT", type=["pdf", "txt", "docx", "xlsx"], key="up_dat")
        
        st.session_state.doc_opcion = st.selectbox(
            "🎯 Selecciona el tipo de análisis:", 
            ["Resumen Ejecutivo", "Extraer Fechas, Nombres y Cifras", "Buscar Contradicciones o Errores", "Crear un Cuestionario (Test)", "Traducir el documento al Inglés"]
        )
        
        if archivo and st.button("🔎 Ejecutar Análisis Seleccionado", key="btn_dat", type="primary"):
            with st.spinner("Analizando..."):
                texto = leer_archivo_universal(archivo, modelo_seleccionado)
                
                if st.session_state.doc_opcion == "Resumen Ejecutivo":
                    prompt = f"Haz un resumen ejecutivo muy claro de este documento. Destaca las ideas principales:\n{texto}"
                elif st.session_state.doc_opcion == "Extraer Fechas, Nombres y Cifras":
                    prompt = f"Actúa como un analista de datos. Extrae en forma de lista todas las fechas, nombres propios y cifras numéricas importantes de este texto:\n{texto}"
                elif st.session_state.doc_opcion == "Buscar Contradicciones o Errores":
                    prompt = f"Actúa como un auditor. Lee este documento y busca contradicciones, datos incoherentes o cosas que no cuadren:\n{texto}"
                elif st.session_state.doc_opcion == "Crear un Cuestionario (Test)":
                    prompt = f"Basándote en este documento, genera un cuestionario de 5 preguntas tipo test (con 4 opciones) y da las respuestas correctas al final para evaluar mi comprensión:\n{texto}"
                else:
                    prompt = f"Traduce todo este documento al inglés manteniendo el tono y formato:\n{texto}"
                
                st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt, api_key_gemini)
                st.text_area("Resultado:", value=st.session_state.analisis_ia, height=400)
                botones_descarga(st.session_state.analisis_ia, "Analisis_Doc")

    with sub[1]:
        st.markdown("#### Creación de Documentos (Word/Excel/PPT)")
        col_form, col_prompt = st.columns([1, 2])
        with col_form:
            formato = st.radio("Formato final:", ["Word", "Excel", "PowerPoint"])
            st.info("Dile a la IA qué necesitas y lo generará en el formato elegido.")
        with col_prompt:
            prompt_office = st.text_area("¿Qué documento quieres crear?", placeholder="Ej: Un presupuesto para un cliente de reformas...")
        if st.button("Generar Documento", key="btn_office", type="primary") and prompt_office:
            with st.spinner("Generando..."):
                texto = llamar_ia(modelo_seleccionado, prompt_office, api_key_gemini)
                st.session_state.office_text = texto
                if formato == "Word": st.download_button("📥 Descargar Word", generar_word(texto), "doc.docx")
                elif formato == "Excel": st.download_button("📥 Descargar Excel", generar_excel(texto), "doc.xlsx")
                else: st.download_button("📥 Descargar PPTX", generar_pptx("Título", texto), "doc.pptx")
        if st.session_state.office_text:
            boton_superpoder(st.session_state.office_text, "Reescribe este documento para que sea un 200% más persuasivo, estructurado y profesional", "office", modelo_seleccionado, api_key_gemini)
            
    with sub[2]:
        st.markdown("#### PDF Pro")
        titulo_pdf = st.text_input("Título del documento:")
        contenido_pdf = st.text_area("Contenido a incluir:")
        if st.button("Generar PDF", key="btn_pdf", type="primary") and titulo_pdf and contenido_pdf:
            texto = llamar_ia(modelo_seleccionado, contenido_pdf, api_key_gemini)
            st.session_state.pdf_text = texto
            st.download_button("📥 Descargar PDF", generar_pdf_pro(titulo_pdf, texto), f"{titulo_pdf}.pdf")
        if st.session_state.pdf_text:
            boton_superpoder(st.session_state.pdf_text, "Añade un resumen ejecutivo al principio y amplía las conclusiones", "pdf", modelo_seleccionado, api_key_gemini)
            
    with sub[3]:
        st.markdown("#### Mapas Mentales")
        col_tema, col_modo = st.columns([2, 1])
        with col_tema: tema = st.text_input("Tema Central:", placeholder="Ej: Redes informáticas")
        with col_modo: modo_infantil = st.checkbox("🎨 Modo Infantil (Emojis)", value=False)
        if st.button("🎯 Trazar Mapa", type="primary", key="btn_mapa") and tema:
            prompt = f"Genera código Mermaid para un mapa mental de '{tema}'."
            if modo_infantil: prompt += " Usa emojis y colores 'classDef rojo fill:#FF6B6B'."
            codigo = llamar_ia(modelo_seleccionado, prompt, api_key_gemini)
            st.session_state.mermaid_code = html.escape(codigo.replace("```mermaid", "").replace("```", "").strip())
        if st.session_state.mermaid_code:
            st.html(f"<div class='mermaid'>{st.session_state.mermaid_code}</div><script type='module'>import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs'; mermaid.initialize({{ startOnLoad: true }});</script>")
            boton_superpoder(f"Añade 3 ramas innovadoras y disruptivas al siguiente mapa mental (código Mermaid):\n{st.session_state.mermaid_code}", "Expandir Mapa con Ramas Creativas", "mapa", modelo_seleccionado, api_key_gemini)
            
    with sub[4]:
        col_qr1, col_qr2 = st.columns(2)
        with col_qr1:
            st.markdown("#### Generador QR")
            txt_qr = st.text_input("Texto o URL para el QR:")
            if st.button("Generar QR") and txt_qr:
                qr = qrcode.QRCode(version=1, box_size=10, border=5); qr.add_data(txt_qr); qr.make(fit=True)
                buf = BytesIO(); qr.make_image().save(buf, format="PNG"); st.image(buf.getvalue())
        with col_qr2:
            st.markdown("#### Lector QR")
            img_qr = st.file_uploader("Sube una imagen de QR:", type=["png", "jpg", "jpeg"])
            if img_qr:
                if not HAS_CV2: st.error("Instala opencv-python: `pip install opencv-python`")
                else:
                    img = Image.open(img_qr).convert('RGB')
                    detector = cv2.QRCodeDetector()
                    data, _, _ = detector.detectAndDecode(np.array(img))
                    if data: st.success(f"Contenido: **{data}**")
                    else: st.warning("No se detectó ningún QR.")
                    
    with sub[5]:
        st.markdown("#### Portapapeles Universal")
        texto_uni = st.text_area("Pega cualquier texto aquí:")
        instruccion = st.text_input("¿Qué quieres que haga la IA con este texto?")
        if st.button("Procesar Texto") and texto_uni and instruccion:
            res = llamar_ia(modelo_seleccionado, f"Instrucción: {instruccion}\n\nTexto: {texto_uni}", api_key_gemini)
            st.session_state.clip_text = res
            st.info(res)
        if st.session_state.clip_text:
            boton_superpoder(st.session_state.clip_text, "Convierte esto en un artículo de blog estructurado y atractivo", "clip", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 2: WEB Y ACTUALIDAD
# ==========================================
with macro_tabs[1]:
    sub = st.tabs(["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV & EVENTOS", "🎥 YOUTUBE", "📰 LECTOR DE WEBS"])
    with sub[0]:
        st.markdown("#### Buscador Global")
        q = st.text_input("Busca en internet:", key="web_q")
        if st.button("🔍 Buscar", key="btn_web", type="primary") and q:
            with st.spinner("Buscando..."):
                res = buscar_en_internet(q)
                for r in res:
                    st.markdown(f"**[{r['titulo']}]({r['url']})**")
                    st.caption(r['resumen'])
                    st.markdown("---")
                    
    with sub[1]:
        st.markdown("#### ⚽ Central de Deportes")
        tab_dep1, tab_dep2, tab_dep3, tab_dep4 = st.tabs(["📰 Noticias", "⚡ Resultados", "📊 Análisis", "🏋️ Entrenamiento"])
        with tab_dep1:
            col_dep, col_liga = st.columns(2)
            with col_dep: deporte = st.selectbox("Deporte:", ["Fútbol", "Baloncesto", "Tenis", "Motor", "Ciclismo"], key="dep_sel_1")
            with col_liga: liga = st.selectbox("Liga:", ["Todas", "La Liga", "Premier League", "Champions League", "NBA"], key="dep_sel_2")
            if st.button("🔍 Buscar Noticias", key="btn_buscar_deportes"):
                with st.spinner(f"Buscando..."): st.session_state.noticias_deportes = buscar_noticias_deportivas(deporte, liga)
            if st.session_state.noticias_deportes:
                for noti in st.session_state.noticias_deportes:
                    st.markdown(f"**{noti['titulo']}** - *{noti['fecha']}*")
                    st.write(noti['resumen'])
                    st.markdown("---")
                boton_superpoder(str(st.session_state.noticias_deportes), "Analiza estas noticias y dime cuál es la noticia más impactante y por qué", "noticias", modelo_seleccionado, api_key_gemini)
        with tab_dep2:
            st.markdown("##### ⚡ Resultados y Directo")
            evento_directo = st.text_input("¿Qué resultados buscas?", placeholder="Ej: Resultados de LaLiga hoy")
            if st.button("Buscar Resultados", type="primary", key="btn_res_directo") and evento_directo:
                with st.spinner("Buscando marcadores..."):
                    res_web = buscar_en_internet(evento_directo)
                    contexto = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_web])
                    prompt_directo = f"Dame los resultados de: '{evento_directo}'. Datos: {contexto}"
                    st.session_state.dep_resultados = llamar_ia(modelo_seleccionado, prompt_directo, api_key_gemini)
                    st.info(st.session_state.dep_resultados)
            if st.session_state.dep_resultados:
                boton_superpoder(st.session_state.dep_resultados, "Haz un análisis táctico profundo de estos resultados como si fueras un entrenador de élite", "tactico", modelo_seleccionado, api_key_gemini)
        with tab_dep3:
            st.markdown("##### 📊 Análisis Táctico y Clasificaciones")
            col_eq1, col_eq2 = st.columns(2)
            with col_eq1: equipo_local = st.text_input("Equipo Local:", placeholder="Ej: Real Madrid")
            with col_eq2: equipo_visitante = st.text_input("Equipo Visitante:", placeholder="Ej: FC Barcelona")
            if st.button("⚔️ Analizar Partido", key="btn_analisis_partido") and equipo_local and equipo_visitante:
                with st.spinner("Analizando rachas..."):
                    prompt_partido = f"Analiza el partido {equipo_local} vs {equipo_visitante}. Pronóstico final."
                    st.session_state.partido_analisis = llamar_ia(modelo_seleccionado, prompt_partido, api_key_gemini)
                    st.info(st.session_state.partido_analisis)
            if st.session_state.partido_analisis:
                boton_superpoder(st.session_state.partido_analisis, "Dame 3 apuestas seguras y el jugador clave que decidirá el partido", "partido", modelo_seleccionado, api_key_gemini)
            st.markdown("---")
            liga_tabla = st.selectbox("Consultar Clasificación:", ["LaLiga EA Sports", "Premier League", "NBA", "Euroliga", "Fórmula 1"], key="sel_tabla")
            if st.button("🏆 Ver Tabla", key="btn_ver_tabla"):
                with st.spinner("Buscando clasificación..."):
                    prompt_tabla = f"Dame la clasificación actual de {liga_tabla} (Posición, Equipo/Piloto, Puntos) y un breve resumen de la lucha por el título."
                    st.info(llamar_ia(modelo_seleccionado, prompt_tabla, api_key_gemini))
        with tab_dep4:
            st.markdown("##### 🏋️ Entrenamiento y Nutrición")
            objetivo = st.text_input("¿Qué necesitas?", placeholder="Ej: Rutina para ganar masa muscular")
            deporte_pract = st.selectbox("Deporte principal:", ["Fútbol", "Running", "Gimnasio/Musculación", "Ciclismo", "Baloncesto", "Tenis"])
            if st.button("Generar Plan", key="btn_plan_deporte") and objetivo:
                with st.spinner("Diseñando plan..."):
                    prompt_plan = f"Entrenador personal. El usuario practica {deporte_pract} y necesita: '{objetivo}'. Dame un plan semanal."
                    st.session_state.plan_deporte = llamar_ia(modelo_seleccionado, prompt_plan, api_key_gemini)
                    st.info(st.session_state.plan_deporte)
            if st.session_state.plan_deporte:
                boton_superpoder(st.session_state.plan_deporte, "Añade una lista de suplementos naturales y una rutina de estiramientos preventivos", "entrenamiento", modelo_seleccionado, api_key_gemini)
    with sub[2]:
        st.markdown("#### 📺 Guía de TV y Eventos")
        tab_tv1, tab_tv2, tab_tv3, tab_tv4 = st.tabs(["🍿 Cine", "📅 TV", "🎫 Entradas", "🎮 Trivia"])
        with tab_tv1:
            estado_animo = st.text_input("¿Qué te apetece ver?", placeholder="Ej: Comedia familiar", key="tv_animo")
            plataformas = st.multiselect("Tus plataformas:", ["Netflix", "Prime Video", "Disney+", "HBO Max", "Movistar+", "Filmin"], key="tv_plat")
            if st.button("🍿 Recomendar", type="primary", key="btn_rec_cine") and estado_animo:
                with st.spinner("Buscando películas..."):
                    prompt_cine = f"Recomienda 3 películas o series. Busco: '{estado_animo}'. Plataformas: {', '.join(plataformas)}."
                    st.session_state.cine_rec = llamar_ia(modelo_seleccionado, prompt_cine, api_key_gemini)
                    st.info(st.session_state.cine_rec)
            if st.session_state.cine_rec:
                boton_superpoder(st.session_state.cine_rec, "Recomienda 3 joyas ocultas (películas no tan famosas) que encajen con esto", "cine", modelo_seleccionado, api_key_gemini)
        with tab_tv2:
            canal_tv = st.text_input("¿Qué canal buscas?", placeholder="Ej: Programación de esta noche en Telecinco")
            if st.button("Buscar Programación", key="btn_tv_prog") and canal_tv:
                with st.spinner("Buscando parrilla..."):
                    res_web_tv = buscar_en_internet(canal_tv)
                    contexto_tv = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_web_tv])
                    prompt_tv = f"Programación de hoy para: '{canal_tv}'. Datos: {contexto_tv}"
                    st.session_state.tv_prog = llamar_ia(modelo_seleccionado, prompt_tv, api_key_gemini)
                    st.info(st.session_state.tv_prog)
            if st.session_state.tv_prog:
                boton_superpoder(st.session_state.tv_prog, "Dime cuál es el programa más recomendado de esta parrilla y por qué", "tv", modelo_seleccionado, api_key_gemini)
        with tab_tv3:
            evento_buscar = st.text_input("¿Qué evento buscas?", placeholder="Ej: Concierto de Taylor Swift")
            ciudad = st.text_input("¿En qué ciudad?", placeholder="Ej: Madrid")
            if st.button("Buscar Entradas", key="btn_entradas") and evento_buscar:
                with st.spinner("Buscando disponibilidad..."):
                    res_entradas = buscar_en_internet(f"entradas {evento_buscar} {ciudad}")
                    contexto_entradas = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_entradas])
                    prompt_entradas = f"Info de: '{evento_buscar}' en {ciudad}. Datos: {contexto_entradas}"
                    st.session_state.entradas = llamar_ia(modelo_seleccionado, prompt_entradas, api_key_gemini)
                    st.info(st.session_state.entradas)
            if st.session_state.entradas:
                boton_superpoder(st.session_state.entradas, "Dame un plan de viaje de 24 horas alrededor de este evento (dónde comer, qué visitar)", "entradas", modelo_seleccionado, api_key_gemini)
        with tab_tv4:
            juego = st.radio("Elige el juego:", ["Adivinar la película", "Curiosidades de series", "Preguntas de deportes"], horizontal=True)
            if st.button("¡Jugar!", key="btn_jugar_trivia"):
                with st.spinner("Generando juego..."):
                    prompt_juego = f"Juego de 3 preguntas trivia sobre '{juego}'. Da respuestas al final."
                    st.session_state.trivia = llamar_ia(modelo_seleccionado, prompt_juego, api_key_gemini)
                    st.info(st.session_state.trivia)
    with sub[3]:
        st.markdown("#### 🎥 Resumidor de YouTube")
        url_yt = st.text_input("Pega la URL del vídeo:")
        if st.button("Extraer y Resumir", key="btn_yt") and url_yt:
            try:
                video_id = url_yt.split("v=")[-1].split("&")[0]
                transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['es', 'en'])
                resumen = llamar_ia(modelo_seleccionado, f"Resume esto:\n{' '.join([t['text'] for t in transcript])[:5000]}", api_key_gemini)
                st.session_state.yt_text = resumen
                st.text_area("Resumen:", value=resumen, height=300)
            except Exception as e: st.error(f"Error: {e}")
        if st.session_state.yt_text:
            boton_superpoder(st.session_state.yt_text, "Saca 3 ideas de negocio o proyectos que se puedan derivar de este resumen", "yt", modelo_seleccionado, api_key_gemini)
    with sub[4]:
        st.markdown("#### 📰 Lector de Páginas Web")
        url_web = st.text_input("Pega el enlace de una noticia o web:")
        if st.button("Leer y Resumir Web", key="btn_read_web") and url_web:
            with st.spinner("Leyendo la web..."):
                try:
                    headers = {'User-Agent': 'Mozilla/5.0'}
                    response = requests.get(url_web, headers=headers, timeout=10)
                    texto_web = re.sub('<[^<]+?>', ' ', response.text)
                    texto_web = " ".join(texto_web.split())[:5000]
                    resumen = llamar_ia(modelo_seleccionado, f"Resume el contenido de esta web:\n\n{texto_web}", api_key_gemini)
                    st.session_state.web_text = resumen
                    st.text_area("Resumen:", value=resumen, height=300)
                except Exception as e: st.error(f"Error: {e}")
        if st.session_state.web_text:
            boton_superpoder(st.session_state.web_text, "Dame una opinión crítica sobre el sesgo o la veracidad de este artículo", "web", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 3: MULTIMEDIA
# ==========================================
with macro_tabs[2]:
    sub = st.tabs(["🖼️ ANÁLISIS MULTIMEDIA", "🎨 GENERADOR VISUAL", "🛡️ CREADOR DE LOGOS", "📥 DESCARGADOR"])
    with sub[0]:
        st.markdown("#### Análisis Inteligente de Archivos")
        
        with st.expander("💡 Cajón de Ideas (¿Qué puedo hacer con este archivo?)"):
            st.write("- **Si es foto:** Puedo pedirle que me extraiga el texto (OCR) o que me diga qué ropa lleva la gente.")
            st.write("- **Si es audio:** Puedo pedirle que me haga un resumen o que traduzca una entrevista.")
            st.write("- **Si es vídeo:** Puedo pedirle que me diga qué escenarios aparecen o de qué hablan los protagonistas.")
            
        archivo_multi = st.file_uploader("📎 Sube tu archivo aquí", type=["png", "jpg", "jpeg", "webp", "mp4", "mov", "mp3", "wav"], key="up_multi")
        if archivo_multi:
            nombre = archivo_multi.name.lower()
            if nombre.endswith((".png", ".jpg", ".jpeg", ".webp")):
                st.image(archivo_multi, use_container_width=True)
                mime_type = "image/jpeg" if nombre.endswith((".jpg", ".jpeg")) else f"image/{nombre.split('.')[-1]}"
                opcion_img = st.radio("Opciones para Imagen:", [
                    "👁️ Describir la imagen detalladamente", 
                    "📝 Extraer todo el texto (OCR)", 
                    "🎨 Generar un prompt para recrearla",
                    "🕵️‍♂️ Analizar si tiene fotomontaje o IA"
                ])
                if st.button("Ejecutar Acción en Imagen", type="primary"):
                    if "Describir" in opcion_img: prompt = "Describe detalladamente esta imagen."
                    elif "texto" in opcion_img: prompt = "Extrae todo el texto visible (OCR)."
                    elif "prompt" in opcion_img: prompt = "Escribe un prompt en inglés para Midjourney que replique esta imagen."
                    else: prompt = "Analiza la imagen y dime si parece real o generada por IA / editada. Fíjate en las sombras, manos y texturas."
                    
                    with st.spinner("Procesando..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            elif nombre.endswith((".mp3", ".wav")):
                st.audio(archivo_multi)
                mime_type = "audio/mp3" if nombre.endswith(".mp3") else "audio/wav"
                opcion_aud = st.radio("Opciones para Audio:", [
                    "📝 Transcribir el audio a texto", 
                    "📋 Hacer un resumen del audio", 
                    "🌍 Traducir el audio al inglés",
                    "🎭 Analizar el tono emocional (¿está enfadado, triste?)"
                ])
                if st.button("Ejecutar Acción en Audio", type="primary"):
                    if "Transcribir" in opcion_aud: prompt = "Transcribe exactamente lo que se dice."
                    elif "Resumir" in opcion_aud: prompt = "Haz un resumen de lo que se habla."
                    elif "Traducir" in opcion_aud: prompt = "Traduce al inglés lo que se dice."
                    else: prompt = "Escucha este audio y dime qué tono emocional tiene (alegre, triste, agresivo, neutro) y por qué."
                    
                    with st.spinner("Procesando..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            elif nombre.endswith((".mp4", ".mov")):
                st.video(archivo_multi)
                mime_type = "video/mp4"
                opcion_vid = st.radio("Opciones para Vídeo:", [
                    "📋 Resumir lo que ocurre en el vídeo", 
                    "🔍 Describir los elementos visuales y escenas", 
                    "🗣️ Analizar el tono o los diálogos"
                ])
                if st.button("Ejecutar Acción en Vídeo", type="primary"):
                    if "Resumir" in opcion_vid: prompt = "Resume el contenido de este vídeo."
                    elif "visuales" in opcion_vid: prompt = "Describe detalladamente las escenas."
                    else: prompt = "Analiza los diálogos y el tono del vídeo."
                    
                    with st.spinner("Procesando vídeo (puede tardar)..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            if st.session_state.analisis_ia:
                st.success("✅ Análisis completado:")
                st.info(st.session_state.analisis_ia)
                botones_descarga(st.session_state.analisis_ia, "Analisis_Multimedia")
                boton_superpoder(st.session_state.analisis_ia, "Extrae información forense de este análisis (qué se oculta, qué es sospechoso)", "multi", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### Generador Visual")
        prompt_img = st.text_input("Describe la imagen:")
        tipo_img = st.radio("Estilo:", ["Foto/Realista", "Dibujo para Colorear"], horizontal=True)
        if st.button("Generar Imagen", key="btn_gen_visual") and prompt_img:
            prompt_final = prompt_img + (", coloring book page, thick black outlines, pure white background, no color" if "Colorear" in tipo_img else "")
            st.session_state.img_url = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_final)}?width=1024&height=1024&nologo=true&enhance=false"
        if st.session_state.img_url:
            st.image(st.session_state.img_url, use_container_width=True)
            boton_superpoder(f"Genera un prompt para recrear esta imagen pero en un estilo cinematográfico y oscuro: {prompt_img}", "Crear Prompt Alternativo", "visual", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### 🛡️ Generador de Logotipos e Identidad")
        st.caption("Crea logotipos profesionales para tu marca, negocio o proyecto.")
        
        with st.expander("💡 Cajón de Ideas para Logos"):
            st.write("- **Cafetería:** Minimalista, líneas finas, tonos tierra.")
            st.write("- **App de Tecnología:** 3D, colores neón, moderno.")
            st.write("- **Marca de Ropa:** Vintage, estilo sello, blanco y negro.")
            
        col_logo1, col_logo2 = st.columns(2)
        with col_logo1:
            nombre_marca = st.text_input("Nombre de la marca:")
            estilo_logo = st.selectbox("Estilo de diseño:", ["Minimalista y limpio", "Corporativo moderno", "3D e Impactante", "Vintage y Retro", "Cyberpunk / Neón"])
        with col_logo2:
            color_logo = st.color_picker("Color principal de la marca:", "#d4af37")
            icono_logo = st.text_input("Icono o símbolo (opcional):", placeholder="Ej: Un lobo, una taza, un rayo...")
            
        if st.button("🛡️ Generar Logo", type="primary", key="btn_gen_logo") and nombre_marca:
            with st.spinner("Diseñando logotipo..."):
                # Creamos un prompt MUY específico para logos
                prompt_logo = f"Professional brand logo design for '{nombre_marca}', {estilo_logo} style, primary color {color_logo}, "
                if icono_logo: prompt_logo += f"with an icon of {icono_logo}, "
                prompt_logo += "flat vector graphic, minimalist, clean white background, centered, high quality, no text, corporate identity"
                
                url_logo = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_logo)}?width=1024&height=1024&nologo=true"
                st.image(url_logo, caption=f"Propuesta de Logo para {nombre_marca}", use_container_width=True)
                st.markdown(f"[Descargar Logo en Alta Resolución]({url_logo})")

    with sub[3]:
        st.markdown("#### Descargador")
        url_dl = st.text_input("Pega la URL del vídeo/audio (YouTube, TikTok...):")
        if st.button("⬇️ Descargar", type="primary") and url_dl:
            if not HAS_YTDLP: st.error("Instala yt-dlp: `pip install yt-dlp`")
            else:
                with st.spinner("Descargando..."):
                    temp_dir = tempfile.mkdtemp()
                    ydl_opts = {'outtmpl': os.path.join(temp_dir, "%(title)s.%(ext)s"), 'quiet': True, 'format': 'best'}
                    try:
                        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                            info = ydl.extract_info(url_dl, download=True)
                            nombre_archivo = ydl.prepare_filename(info)
                        with open(nombre_archivo, "rb") as f: st.session_state.dl_bytes = f.read()
                        st.session_state.dl_filename = os.path.basename(nombre_archivo)
                        st.success(f"✅ **{info.get('title', 'Archivo')}** descargado.")
                    except Exception as e: st.error(f"Error: {e}")
                    finally:
                        for archivo in os.listdir(temp_dir): os.remove(os.path.join(temp_dir, archivo))
                        os.rmdir(temp_dir)
        if st.session_state.dl_bytes:
            st.download_button("💾 Guardar archivo", st.session_state.dl_bytes, st.session_state.dl_filename)

# ==========================================
# CATEGORÍA 4: ASISTENTES IA
# ==========================================
with macro_tabs[3]:
    sub = st.tabs(["💬 CHAT", "🎙️ VOZ", "💬 CHAT CON PDF", "🕵️‍♂️ AGENTES CUSTOM", "🍳 COCINA", "📖 CUENTOS", "💻 CÓDIGO", "🌍 TRADUCTOR"])
    with sub[0]:
        st.markdown("#### Chat con IA")
        for msg in st.session_state.chat_history: st.chat_message(msg["role"]).write(msg["content"])
        if prompt := st.chat_input("Escribe tu mensaje..."):
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            st.chat_message("user").write(prompt)
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, prompt, api_key_gemini)
                st.session_state.chat_history.append({"role": "assistant", "content": res})
                st.chat_message("assistant").write(res)
                
    with sub[1]:
        st.markdown("#### Sintetizador de Voz")
        texto_pod = st.text_area("Escribe el texto que quieres escuchar:")
        voz = st.selectbox("Selecciona la voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"])
        if st.button("🔊 Generar Audio", key="btn_voz") and texto_pod:
            with st.spinner("Generando voz..."):
                audio = generar_audio_sincrono(texto_pod, voz)
                if audio: st.audio(audio, format="audio/mp3")
                
    with sub[2]:
        st.markdown("#### Chat con Documentos")
        archivos = st.file_uploader("Sube documentos (PDF, Word, TXT):", type=["pdf", "txt", "docx"], accept_multiple_files=True)
        if archivos:
            texto_completo = "".join([leer_archivo_universal(a, modelo_seleccionado) for a in archivos])[:MAX_TEXTO_IA]
            st.success(f"✅ {len(archivos)} documentos leídos. Puedes preguntar.")
            for msg in st.session_state.pdf_chat_history: st.chat_message(msg["role"]).write(msg["content"])
            if prompt := st.chat_input("Pregúntale a los documentos..."):
                st.session_state.pdf_chat_history.append({"role": "user", "content": prompt})
                st.chat_message("user").write(prompt)
                with st.spinner("Buscando en el texto..."):
                    res = llamar_ia(modelo_seleccionado, f"Basándote en esto:\n{texto_completo}\nResponde: {prompt}", api_key_gemini)
                    st.session_state.pdf_chat_history.append({"role": "assistant", "content": res})
                    st.chat_message("assistant").write(res)

    with sub[3]:
        st.markdown("#### 🕵️‍♂️ Creador de Agentes IA")
        st.caption("Crea un personaje o experto con una personalidad única y chatea con él.")
        
        # Configuración del Agente
        with st.container():
            nombre_agente = st.text_input("Nombre del Agente:", placeholder="Ej: Albert, el Profesor de Física")
            rol_agente = st.text_area("Personalidad y Reglas (System Prompt):", 
                                      placeholder="Ej: Eres un profesor de física muy divertido. Explicas conceptos complejos usando ejemplos de dibujos animados. Usas muchos emojis.",
                                      height=100)
            
            if st.button("✨ Crear / Activar Agente", type="primary"):
                if nombre_agente and rol_agente:
                    st.session_state.agente_activo = {"nombre": nombre_agente, "rol": rol_agente}
                    st.session_state.agente_chat_history = [] # Limpiar chat al crear nuevo agente
                    st.success(f"Agente '{nombre_agente}' activado y listo para hablar.")
                else:
                    st.warning("Ponle un nombre y una personalidad al agente para activarlo.")
                    
        # Interfaz de Chat del Agente
        if st.session_state.get("agente_activo"):
            agente = st.session_state.agente_activo
            st.markdown(f"---")
            st.markdown(f"### 💬 Hablando con: **{agente['nombre']}**")
            
            for msg in st.session_state.agente_chat_history: 
                st.chat_message(msg["role"]).write(msg["content"])
                
            if prompt := st.chat_input(f"Escribe un mensaje para {agente['nombre']}..."):
                st.session_state.agente_chat_history.append({"role": "user", "content": prompt})
                st.chat_message("user").write(prompt)
                
                with st.spinner(f"{agente['nombre']} está pensando..."):
                    # Inyectamos la personalidad en el prompt
                    prompt_con_rol = f"INSTRUCCIONES DEL SISTEMA (ERES ESTE AGENTE):\n{agente['rol']}\n\nEl usuario dice: {prompt}"
                    
                    res = llamar_ia(modelo_seleccionado, prompt_con_rol, api_key_gemini)
                    st.session_state.agente_chat_history.append({"role": "assistant", "content": res})
                    st.chat_message("assistant").write(res)

    with sub[4]:
        st.markdown("#### 🍳 El Chef Virtual")
        
        with st.expander("💡 Cajón de Ideas para Cocinar"):
            st.write("- Pídele un **menú semanal** equilibrado para 4 personas.")
            st.write("- Dile los ingredientes que tienes y pídele **una receta express** de 15 minutos.")
            st.write("- Pídele **alternativas saludables** para una receta de postre.")
            
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            ingredientes = st.text_area("¿Qué tienes en la nevera?", placeholder="Pollo, arroz, tomate...")
            if st.button("Dame una receta") and ingredientes:
                st.session_state.receta_generada = llamar_ia(modelo_seleccionado, f"Tengo {ingredientes}. Dame una receta rápida.", api_key_gemini)
            if st.session_state.receta_generada: st.info(st.session_state.receta_generada)
        with col_c2:
            preferencias = st.text_input("¿Menú semanal? (Alergias, gustos)")
            if st.button("Generar Menú + Compra") and preferencias:
                st.session_state.menu_generado = llamar_ia(modelo_seleccionado, f"Haz un menú semanal para 4 personas. {preferencias}. Añade lista de la compra.", api_key_gemini)
            if st.session_state.menu_generado: st.info(st.session_state.menu_generado)
            
    with sub[5]:
        st.markdown("#### 📖 Cuentacuentos")
        prota = st.text_input("Nombre del protagonista:")
        lugar = st.text_input("Lugar mágico:")
        if st.button("✨ Crear Cuento") and prota and lugar:
            cuento = llamar_ia(modelo_seleccionado, f"Cuento infantil. Protagonista: {prota}, Lugar: {lugar}. Moraleja al final.", api_key_gemini)
            st.text_area("Cuento:", value=cuento, height=300)
            if st.button("🔊 Leer en voz alta"):
                audio = generar_audio_sincrono(cuento, "es-ES-ElviraNeural")
                if audio: st.audio(audio, format="audio/mp3")
                
    with sub[6]:
        st.markdown("#### Asistente de Código")
        req_code = st.text_area("Pega un código para explicarlo o pide uno nuevo:")
        if st.button("Procesar Código") and req_code:
            res = llamar_ia(modelo_seleccionado, f"Actúa como programador senior. Explica o mejora esto:\n{req_code}", api_key_gemini)
            st.code(res, language="python")
            
    with sub[7]:
        st.markdown("#### Traductor")
        texto_trad = st.text_area("Texto a traducir:")
        idioma = st.selectbox("Idioma destino:", ["Inglés", "Francés", "Alemán", "Italiano"])
        if st.button("Traducir") and texto_trad:
            res = llamar_ia(modelo_seleccionado, f"Traduce al {idioma}: {texto_trad}", api_key_gemini)
            st.info(res)

# ==========================================
# CATEGORÍA 5: ESTUDIO Y OFICIOS
# ==========================================
with macro_tabs[4]:
    sub = st.tabs(["🌐 REDES Y SISTEMAS", "💇‍♀️ PELUQUERÍA", "📚 DEBERES Y ESTUDIO"])
    with sub[0]:
        st.markdown("#### 🌐 Herramientas de Redes")
        
        with st.expander("💡 Cajón de Ideas para Redes"):
            st.write("- Pídele que te explique la diferencia entre **TCP y UDP** con una analogía.")
            st.write("- Pídele el comando de Linux para **ver las IPs conectadas** a tu red.")
            st.write("- Pídele **3 preguntas de examen** sobre subnetting.")
            
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            ip = st.text_input("IP/CIDR (Ej: 192.168.1.0/24)")
            if st.button("Calcular Red") and ip:
                st.session_state.redes_resultado = llamar_ia(modelo_seleccionado, f"Calcula máscara, rango, broadcast y hosts para: {ip}", api_key_gemini)
                st.info(st.session_state.redes_resultado)
        with col_r2:
            tarea = st.text_input("¿Qué necesitas hacer? (Ej: ver puertos abiertos)")
            so = st.selectbox("Sistema:", ["Windows", "Linux", "Cisco"])
            if st.button("Buscar Comando") and tarea:
                st.session_state.redes_resultado = llamar_ia(modelo_seleccionado, f"Comando para {so} que haga: {tarea}. Explícalo.", api_key_gemini)
                st.info(st.session_state.redes_resultado)
        st.markdown("---")
        st.markdown("**Simulador de Entrevista Técnica**")
        if st.button("Comenzar Entrevista (Nivel CCNA)"):
            st.session_state.entrevista_resultado = llamar_ia(modelo_seleccionado, "Hazme 3 preguntas técnicas de nivel CCNA (redes) como si fueras un reclutador. Espera mi respuesta.", api_key_gemini)
        if st.session_state.entrevista_resultado: st.info(st.session_state.entrevista_resultado)
        
    with sub[1]:
        st.markdown("#### 💇‍♀️ Asistente de Peluquería")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            actual = st.text_input("Color actual:")
            deseado = st.text_input("Color deseado:")
            estado = st.selectbox("Estado del pelo:", ["Sano", "Teñido", "Decolorado", "Dañado"])
            if st.button("Calcular Fórmula") and actual and deseado:
                st.session_state.pelo_resultado = llamar_ia(modelo_seleccionado, f"Fórmula para pasar de {actual} a {deseado}. Pelo {estado}. Da tintes, oxidantes y pasos.", api_key_gemini)
                st.info(st.session_state.pelo_resultado)
        with col_p2:
            servicios = st.text_area("Servicios a realizar (Ej: Mechas, corte, peinado):")
            tiempo = st.number_input("Tiempo estimado (horas):", min_value=1.0, step=0.5)
            if st.button("Calcular Presupuesto") and servicios:
                st.session_state.presupuesto_resultado = llamar_ia(modelo_seleccionado, f"Crea un presupuesto profesional para: {servicios}. Tiempo: {tiempo}h. Sugiere precio competitivo y desglose.", api_key_gemini)
                st.info(st.session_state.presupuesto_resultado)
                
    with sub[2]:
        st.markdown("#### 📚 Asistente de Deberes")
        nivel = st.selectbox("Nivel Educativo:", ["Primaria", "ESO", "Bachillerato", "Universidad"])
        materia = st.selectbox("Materia:", ["Matemáticas", "Lengua", "Ciencias", "Historia", "Idiomas"])
        duda = st.text_area("Escribe tu duda o ejercicio:")
        if st.button("Explicar Duda") and duda:
            res = llamar_ia(modelo_seleccionado, f"Eres un profesor. Nivel: {nivel}. Materia: {materia}. Explica esto paso a paso y con ejemplos: {duda}", api_key_gemini)
            st.success("Explicación:")
            st.info(res)

# ==========================================
# CATEGORÍA 6: VIAJES Y HOGAR
# ==========================================
with macro_tabs[5]:
    sub = st.tabs(["🌍 VIAJES", "💰 FINANZAS", "✉️ CARTAS Y MENSAJES"])
    with sub[0]:
        st.markdown("#### 🌍 Planificador de Viajes")
        
        with st.expander("💡 Cajón de Ideas para Viajar"):
            st.write("- Pide un **itinerario con niños** para no aburrirlos.")
            st.write("- Pide una **lista de la maleta** según el clima del destino.")
            st.write("- Pide **restaurantes económicos** típicos del lugar.")
            
        destino = st.text_input("¿A dónde quieres ir?", placeholder="Ej: Roma, Nueva York, un pueblo de Asturias...")
        col_v1, col_v2 = st.columns(2)
        with col_v1: dias = st.number_input("¿Cuántos días?", min_value=1, max_value=30, value=3)
        with col_v2: compania = st.selectbox("¿Con quién viajas?", ["Pareja", "Familia con niños", "Amigos", "Solo/a", "Tercera edad"])
        if st.button("✈️ Generar Itinerario", type="primary") and destino:
            with st.spinner("Planificando tu viaje..."):
                prompt_viaje = f"Actúa como un guía de viajes experto. Diseña un itinerario de {dias} días para {destino} viajando con {compania}. Incluye qué ver cada día (mañana, tarde y noche), recomendaciones de comida, un presupuesto aproximado y una lista de 5 cosas imprescindibles para la maleta."
                st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt_viaje, api_key_gemini)
                st.info(st.session_state.analisis_ia)
                botones_descarga(st.session_state.analisis_ia, f"Viaje_{destino}")
                
    with sub[1]:
        st.markdown("#### 💰 Finanzas del Hogar")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.markdown("**Calculadora de Ahorro**")
            ingresos = st.number_input("Ingresos netos mensuales (€):", min_value=0)
            gastos = st.number_input("Gastos fijos mensuales (€):", min_value=0)
            if st.button("Calcular Ahorro"):
                ahorro = ingresos - gastos
                if ahorro > 0: st.success(f"Te quedan **{ahorro}€** al mes. ¡Puedes ahorrar!")
                else: st.error(f"Te faltan **{abs(ahorro)}€** al mes. Hay que ajustar gastos.")
        with col_f2:
            st.markdown("**Calculadora de Hipoteca**")
            prestamo = st.number_input("Importe del préstamo (€):", value=150000, step=10000)
            interes = st.number_input("Interés anual (%):", value=3.0, step=0.1)
            años = st.number_input("Años a pagar:", value=30, step=1)
            if st.button("Calcular Cuota"):
                i = (interes / 100) / 12
                n = años * 12
                cuota = prestamo * (i * (1 + i)**n) / ((1 + i)**n - 1)
                st.success(f"Tu cuota mensual sería de **{cuota:.2f}€**.")
                
    with sub[2]:
        st.markdown("#### ✉️ Cartas y Mensajes")
        
        with st.expander("💡 Cajón de Ideas"):
            st.write("- Pide un **discurso emotivo** para el cumpleaños de tu pareja.")
            st.write("- Pide una **carta formal** para reclamar una factura del banco.")
            
        tipo_mensaje = st.selectbox("¿Qué tipo de mensaje necesitas?", ["Carta para los Reyes Magos", "Carta a Papá Noel", "Discurso de cumpleaños", "Mensaje de aniversario", "Carta formal de reclamación"])
        detalles = st.text_area("Añade detalles (ej: nombre del niño, lo que ha hecho bien este año, o qué quieres reclamar):")
        if st.button("✍️ Redactar Mensaje", type="primary") and detalles:
            with st.spinner("Redactando..."):
                prompt_mensaje = f"Actúa como un redactor experto. Escribe un/a {tipo_mensaje} basándote en estos detalles: {detalles}. Usa el tono adecuado (emotivo, divertido o formal según corresponda)."
                st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt_mensaje, api_key_gemini)
                st.text_area("Resultado:", value=st.session_state.analisis_ia, height=300)
                botones_descarga(st.session_state.analisis_ia, f"{tipo_mensaje.replace(' ', '_')}")

# ==========================================
# CATEGORÍA 7: MÚSICA Y ARTE
# ==========================================
with macro_tabs[6]:
    sub = st.tabs(["🎵 CREADOR DE CANCIONES", "🎨 POESÍA Y RIMAS"])
    
    with sub[0]:
        st.markdown("#### 🎵 Estudio de Creación Musical")
        st.caption("Crea la letra, la portada y escucha una maqueta de tu canción (voz recitada).")
        
        with st.expander("💡 Cajón de Ideas Musicales"):
            st.write("- Pídele **una canción de superación** para dedicar en un cumpleaños.")
            st.write("- Pídele una **canción de rap** que hable de la historia de España.")
            
        col_m1, col_m2 = st.columns(2)
        with col_m1:
            tema_cancion = st.text_input("¿De qué trata la canción?", placeholder="Ej: Un amor de verano, superación personal...")
        with col_m2:
            genero = st.selectbox("Género Musical:", ["Pop", "Rock", "Rap/Trap", "Reggaetón", "Balada", "Indie"])
            
        tono = st.select_slider("Tono de la canción:", options=["Muy triste", "Melancólico", "Alegre", "Fiestero", "Épico"])
        
        if st.button("🎼 Escribir Canción", type="primary", key="btn_crear_cancion") and tema_cancion:
            with st.spinner("Componiendo la letra..."):
                prompt_cancion = f"Actúa como un compositor de éxito. Escribe la letra de una canción de género {genero} sobre: '{tema_cancion}'. El tono debe ser {tono}. Estructura la canción claramente con etiquetas: [Verso 1], [Coro], [Verso 2], [Puente], [Coro Final]. Sé creativo y rítmico."
                letra = llamar_ia(modelo_seleccionado, prompt_cancion, api_key_gemini)
                st.session_state.letra_cancion = letra
                
                # Extraer un título aproximado
                titulo_prompt = f"Basándote en esta letra, dame SOLO un título corto y pegadizo (máximo 5 palabras):\n{letra[:500]}"
                st.session_state.titulo_cancion = llamar_ia(modelo_seleccionado, titulo_prompt, api_key_gemini).replace('"', '').strip()
                
                # Generar portada
                prompt_img_cancion = f"Album cover art, {genero} style, mood {tono}, about {tema_cancion}, high quality, 4k, artistic"
                st.session_state.portada_cancion = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img_cancion)}?width=1024&height=1024&nologo=true"
                
        if st.session_state.letra_cancion:
            st.markdown(f"### 🎶 Título: {st.session_state.titulo_cancion}")
            
            col_portada, col_letra = st.columns([1, 2])
            with col_portada:
                st.image(st.session_state.portada_cancion, caption="Portada del Single", use_container_width=True)
                
                # Opciones de audio
                st.markdown("#### 🎙️ Escuchar Maqueta")
                voz_cancion = st.selectbox("Elige la voz:", [
                    "es-MX-JorgeNeural (Recomendado para Rap)", 
                    "es-ES-ElviraNeural (Recomendado para Baladas)", 
                    "es-AR-TomasNeural (Tono más grave)"
                ], key="voz_cancion")
                
                voz_id = voz_cancion.split(" ")[0]
                
                if st.button("🔊 Generar Audio de la Letra", key="btn_audio_cancion"):
                    with st.spinner("Grabando maqueta..."):
                        audio_letra = generar_audio_sincrono(st.session_state.letra_cancion, voz_id)
                        if audio_letra:
                            st.audio(audio_letra, format="audio/mp3")
                            st.download_button("💾 Descargar Audio (.mp3)", audio_letra, f"{st.session_state.titulo_cancion}.mp3", "audio/mpeg")

            with col_letra:
                st.text_area("Letra de la Canción:", value=st.session_state.letra_cancion, height=500)
                botones_descarga(st.session_state.letra_cancion, f"Letra_{st.session_state.titulo_cancion}")

    with sub[1]:
        st.markdown("#### 🎨 Poesía y Rimas")
        st.caption("Genera poemas personalizados para regalar o para inspirarte.")
        
        tema_poema = st.text_input("¿Sobre qué quieres el poema?", placeholder="Ej: El mar, mi mascota, el paso del tiempo...")
        estilo_poema = st.selectbox("Estilo:", ["Clásico y Romántico", "Divertido y Loco", "Épico y Épico", "Haiku (Corto)"])
        
        if st.button("✒️ Escribir Poema", type="primary") and tema_poema:
            with st.spinner("Inspirándose..."):
                prompt_poema = f"Escribe un poema {estilo_poema} sobre: '{tema_poema}'. Que tenga rima y una estructura bonita."
                poema = llamar_ia(modelo_seleccionado, prompt_poema, api_key_gemini)
                st.text_area("Poema:", value=poema, height=300)
                botones_descarga(poema, f"Poema_{tema_poema}")
                if st.button("🔊 Recitar Poema"):
                    audio_poema = generar_audio_sincrono(poema, "es-ES-ElviraNeural")
                    if audio_poema: st.audio(audio_poema, format="audio/mp3")
