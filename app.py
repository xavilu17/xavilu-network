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
import random

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
MAX_FILE_MB = 1000
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-3.8-flash"

# --- ESTADO DE SESIÓN ---
_defaults = {
    "chat_history": [], "img_url": "", "app_code": "", "pdf_chat_history": [],
    "mermaid_code": "", "analisis_ia": "", "dl_bytes": None, "dl_filename": "", "dl_mime": "",
    "cuento_generado": "", "receta_generada": "", "menu_generado": "",
    "redes_resultado": "", "pelo_resultado": "", "deberes_resultado": "", "prl_resultado": "",
    "qr_resultado": "", "code_resultado": "", "entrevista_resultado": "",
    "presupuesto_resultado": "", "noticias_deportes": [], "resumen_deportivo": "",
    "evento_resultados": None, "office_text": "", "pdf_text": "", "clip_text": "",
    "dep_resultados": "", "partido_analisis": "", "plan_deporte": "",
    "cine_rec": "", "tv_prog": "", "entradas": "", "trivia": "", "yt_text": "", "web_text": "",
    "letra_cancion": "", "titulo_cancion": "", "portada_cancion": "",
    "doc_opcion": "Resumen Ejecutivo",
    "global_context": "", "global_file_name": "",
    "agente_activo": None, "agente_chat_history": [],
    "debate_resultado": "", "img_prompt_en": ""
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# --- 🎨 CSS CYBERPUNK: PESTAÑAS TIPO BOTÓN NEÓN ---
st.markdown("""
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
    
    /* Fondo Negro Espacial */
    .stApp { 
        background-color: #000000 !important; 
        background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; 
        color: #f7dfa3 !important; 
        font-family: 'Rajdhani', sans-serif !important; 
    }
    
    /* Títulos con brillo neón */
    h1, h2, h3 { 
        font-family: 'Orbitron', sans-serif !important; 
        color: #ffffff !important; 
        text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff, 0 0 30px #FF007F !important; 
        text-align: center; 
        letter-spacing: 3px; 
    }
    
    /* Etiquetas de los inputs */
    [data-testid="stWidgetLabel"] p { 
        color: #00f0ff !important; 
        font-family: 'Orbitron', sans-serif !important; 
        font-size: 13px !important; 
        font-weight: bold !important; 
        text-transform: uppercase !important; 
        background: rgba(0, 240, 255, 0.1) !important; 
        padding: 5px 10px !important; 
        border-left: 4px solid #FF007F !important; 
        border-radius: 4px !important; 
        display: inline-block !important; 
        margin-bottom: 8px !important; 
    }
    
    /* Barra lateral */
    [data-testid="stSidebar"] { 
        background-color: rgba(5, 5, 10, 0.98) !important; 
        border-right: 2px solid #00f0ff !important; 
        box-shadow: 5px 0 15px rgba(0, 240, 255, 0.2) !important; 
    }
    
    /* 🌟 PESTAÑAS ESTILO BOTÓN NEÓN 🌟 */
    .stTabs [data-baseweb="tab-list"] {
        gap: 10px !important;
        border-bottom: 2px solid rgba(0, 240, 255, 0.2) !important;
        padding-bottom: 10px !important;
    }
    
    /* Pestañas inactivas (Igual que el botón "Ejecutar mi Idea") */
    .stTabs [data-baseweb="tab"] {
        background: rgba(10, 10, 10, 0.8) !important;
        border: 2px solid #00f0ff !important;
        border-radius: 25px !important;
        padding: 8px 20px !important;
        transition: all 0.3s ease !important;
        box-shadow: 0 0 10px rgba(0, 240, 255, 0.3) !important;
        opacity: 0.7 !important;
    }
    
    /* Texto de las inactivas */
    .stTabs [data-baseweb="tab"] p,
    .stTabs [data-baseweb="tab"] div {
        color: #00f0ff !important;
        font-family: 'Orbitron', sans-serif !important;
        font-weight: bold !important;
        font-size: 13px !important;
        opacity: 1 !important;
    }
    
    /* Al pasar el ratón */
    .stTabs [data-baseweb="tab"]:hover {
        background: rgba(0, 240, 255, 0.1) !important;
        box-shadow: 0 0 15px rgba(0, 240, 255, 0.6) !important;
        opacity: 1 !important;
        transform: translateY(-2px);
    }
    
    /* Pestaña activa (ENCENDIDA: Degradado rosa/cian y letras negras) */
    .stTabs [data-baseweb="tab"][aria-selected="true"] {
        background: linear-gradient(135deg, #FF007F 0%, #00f0ff 100%) !important;
        border: 2px solid #ffffff !important;
        box-shadow: 0 0 25px rgba(255, 0, 127, 0.9) !important;
        opacity: 1 !important;
        transform: translateY(-2px);
    }
    
    /* Texto de la pestaña activa */
    .stTabs [data-baseweb="tab"][aria-selected="true"] p,
    .stTabs [data-baseweb="tab"][aria-selected="true"] div {
        color: #000000 !important;
        font-weight: 900 !important;
    }
    
    /* Ocultar la línea de Streamlit por defecto */
    [data-baseweb="tab-highlight"] {
        display: none !important;
    }
    
    /* Cajas de texto e inputs */
    .stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { 
        background-color: rgba(10, 10, 10, 0.9) !important; 
        border: 1px solid #00f0ff !important; 
        border-radius: 8px !important; 
        color: #ffffff !important; 
        font-size: 16px !important; 
        box-shadow: inset 0 0 10px rgba(0, 240, 255, 0.1) !important;
    }
    .stTextArea textarea:focus, .stTextInput input:focus { 
        border: 1px solid #FF007F !important; 
        box-shadow: 0 0 15px rgba(255, 0, 127, 0.5) !important; 
    }
    
    /* Botones */
    div.stButton > button, div[data-testid="stDownloadButton"] > button { 
        background: rgba(10, 10, 10, 0.8) !important; 
        color: #00f0ff !important; 
        border: 2px solid #00f0ff !important; 
        border-radius: 25px !important; 
        font-family: 'Orbitron', sans-serif !important; 
        font-weight: bold !important; 
        padding: 10px 20px !important; 
        transition: all 0.3s !important;
        box-shadow: 0 0 10px rgba(0, 240, 255, 0.3) !important;
    }
    div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover { 
        background: linear-gradient(135deg, #FF007F 0%, #00f0ff 100%) !important; 
        color: #000 !important; 
        border: 2px solid #fff !important; 
        transform: translateY(-2px); 
        box-shadow: 0 0 20px rgba(255, 0, 127, 0.8) !important; 
    }
    
    /* Mensajes del Chat */
    [data-testid="stChatMessage"] { 
        background: rgba(10, 10, 15, 0.8) !important; 
        border: 1px solid #FF007F !important; 
        border-radius: 15px !important; 
        padding: 10px; 
        box-shadow: 0 0 15px rgba(255, 0, 127, 0.3) !important;
    }
    
    /* Alertas */
    .stAlert { 
        border-radius: 10px; 
        font-weight: bold; 
        background-color: rgba(0, 240, 255, 0.1) !important;
        border-left: 5px solid #00f0ff !important;
    }
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

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None):
    prompt_final = prompt[:MAX_TEXTO_IA]
    
    if st.session_state.get("global_context"):
        nombre_archivo = st.session_state.get("global_file_name", "Archivo_Adjunto")
        prompt_final = f"[EL USUARIO HA ADJUNTADO EL SIGUIENTE ARCHIVO: {nombre_archivo}]\nCONTENIDO EXTRAÍDO DEL ARCHIVO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN DEL USUARIO]:\n{prompt_final}\n\n(Responde teniendo en cuenta el archivo adjunto si es relevante)."

    if historial_chat:
        contexto_chat = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in historial_chat[-10:]])
        prompt_final = f"Historial de la conversación hasta ahora:\n{contexto_chat}\n\nNuevo mensaje del usuario:\n{prompt_final}\n\nResponde al nuevo mensaje."

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
        
        nombre_elegido = st.text_input("📝 Ponle nombre a tu archivo (sin extensión):", value=nombre_por_defecto, key=f"input_name_{key_unica}")
        nombre_limpio = re.sub(r'[^a-zA-Z0-9_\- ]', '', nombre_elegido).strip().replace(' ', '_')
        if not nombre_limpio: nombre_limpio = nombre_por_defecto

        c1, c2 = st.columns(2)
        with c1: st.download_button("📥 Descargar (TXT)", texto.encode('utf-8'), f"{nombre_limpio}.txt", "text/plain", key=f"dl_txt_{key_unica}", use_container_width=True)
        with c2: st.download_button("📄 Descargar (PDF)", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"dl_pdf_{key_unica}", use_container_width=True)

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

def cajon_ideas_personalizado(contexto_pestana, modelo, api_key):
    st.markdown("---")
    st.markdown("### 🧠 Tu Idea Propia (Control Total)")
    st.caption(f"¿Tienes otra idea para {contexto_pestana}? Escríbela aquí y la IA la ejecutará usando el contexto de la app.")
    
    idea_libre = st.text_area("Escribe tu instrucción personalizada:", placeholder="Ej: Haz un resumen enfocado en los costes... / Escribe un poema sobre esto...", key=f"idea_libre_{contexto_pestana}")
    
    if st.button("🚀 Ejecutar mi Idea", key=f"btn_idea_libre_{contexto_pestana}", type="primary"):
        if idea_libre:
            with st.spinner("Procesando tu idea..."):
                res = llamar_ia(modelo, idea_libre, api_key)
                st.success("¡Resultado de tu idea!")
                st.info(res)
                botones_descarga(res, f"Idea_Libre_{contexto_pestana}")
        else:
            st.warning("Escribe primero tu idea.")

# ==========================================
# INTERFAZ PRINCIPAL
# ==========================================
st.markdown("<h1>XAVILU NETWORK</h1><p style='text-align:center; color:#00f0ff; font-weight:bold; letter-spacing:2px;'>● SISTEMA CONECTADO</p>", unsafe_allow_html=True)

with st.sidebar:
    st.markdown("### 🪐 PANEL DE CONTROL")
    modelo_seleccionado = st.selectbox("Motor Local (Ollama)", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])
    try: api_key_gemini = st.secrets["GEMINI_API_KEY"]; st.success("🔑 Gemini API OK")
    except: api_key_gemini = st.text_input("Gemini API Key", type="password")
    
    st.markdown("---")
    st.markdown("### 🧠 Cerebro Global")
    st.caption("Sube cualquier archivo (foto, audio, pdf). La IA lo leerá y lo usará en TODAS las pestañas.")
    archivo_global = st.file_uploader("📁 Contexto Global", type=["pdf", "txt", "docx", "png", "jpg", "jpeg", "mp3", "wav", "mp4"], key="up_global")
    
    if archivo_global:
        if st.session_state.get("global_file_name") != archivo_global.name:
            with st.spinner("Procesando archivo para el cerebro global..."):
                nombre = archivo_global.name.lower()
                if nombre.endswith((".png", ".jpg", ".jpeg")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "image/jpeg", "Describe esta imagen con todo detalle.", api_key_gemini)
                elif nombre.endswith((".mp3", ".wav")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "audio/mp3", "Transcribe este audio.", api_key_gemini)
                elif nombre.endswith((".mp4")):
                    ctx = analizar_archivo_con_gemini(archivo_global.getvalue(), "video/mp4", "Resume el contenido de este vídeo.", api_key_gemini)
                else:
                    ctx = leer_archivo_universal(archivo_global, modelo_seleccionado)
                
                st.session_state.global_context = ctx[:5000]
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
    "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD",
    "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR",
    "🎵 MÚSICA Y ARTE"
])

# ==========================================
# CATEGORÍA 1: CHAT Y ASISTENTES IA
# ==========================================
with macro_tabs[0]:
    sub = st.tabs(["💬 CHAT", "🎙️ VOZ", "💬 CHAT CON PDF", "🕵️‍♂️ AGENTES CUSTOM", "⚖️ EL DEBATE", "🍳 COCINA", "📖 CUENTOS", "💻 CÓDIGO", "🌍 TRADUCTOR"])
    
    with sub[0]:
        st.markdown("#### 💬 Chat con IA")
        st.caption("Este chat tiene memoria. Recuerda lo que habéis hablado anteriormente.")
        
        if st.session_state.chat_history:
            col_q1, col_q2, col_q3 = st.columns(3)
            quick_prompt = ""
            with col_q1:
                if st.button("👶 Explicar como a un niño", use_container_width=True, key="quick_nino"): 
                    quick_prompt = "Coge tu última respuesta y explícamela como si tuviera 5 años."
            with col_q2:
                if st.button("📝 Resumir en 3 puntos", use_container_width=True, key="quick_resumen"): 
                    quick_prompt = "Resume tu última respuesta en 3 puntos clave."
            with col_q3:
                if st.button("🌍 Traducir al inglés", use_container_width=True, key="quick_trad"): 
                    quick_prompt = "Traduce tu última respuesta al inglés."
        else:
            quick_prompt = ""

        for msg in st.session_state.chat_history: 
            st.chat_message(msg["role"]).write(msg["content"])
            
        if st.session_state.chat_history:
            col_u1, col_u2, col_u3 = st.columns(3)
            with col_u1:
                texto_historial = "\n\n".join([f"{m['role'].upper()}: {m['content']}" for m in st.session_state.chat_history])
                st.download_button("📥 Guardar Chat (TXT)", texto_historial.encode('utf-8'), "chat.txt", "text/plain", use_container_width=True, key="dl_chat_txt")
            with col_u2:
                if st.button("🗑️ Limpiar Conversación", use_container_width=True, key="btn_clear_chat"):
                    st.session_state.chat_history = []
                    st.rerun()
            with col_u3:
                if st.button("🔊 Leer última respuesta", use_container_width=True, key="btn_tts_chat") and st.session_state.chat_history:
                    with st.spinner("Generando voz..."):
                        ultimo_msg = st.session_state.chat_history[-1]["content"]
                        audio_chat = generar_audio_sincrono(ultimo_msg, "es-ES-ElviraNeural")
                        if audio_chat: st.audio(audio_chat, format="audio/mp3")

        prompt = st.chat_input("Escribe tu mensaje...", key="chat_main_input")
        
        if quick_prompt:
            prompt = quick_prompt
            
        if prompt:
            st.session_state.chat_history.append({"role": "user", "content": prompt})
            st.chat_message("user").write(prompt)
            
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, prompt, api_key_gemini, historial_chat=st.session_state.chat_history)
                st.session_state.chat_history.append({"role": "assistant", "content": res})
                st.chat_message("assistant").write(res)
                st.rerun()
                
        cajon_ideas_personalizado("Chat IA", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### Sintetizador de Voz")
        texto_pod = st.text_area("Escribe el texto que quieres escuchar:", key="voz_texto")
        voz = st.selectbox("Selecciona la voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"], key="voz_select")
        if st.button("🔊 Generar Audio", key="btn_voz") and texto_pod:
            with st.spinner("Generando voz..."):
                audio = generar_audio_sincrono(texto_pod, voz)
                if audio: st.audio(audio, format="audio/mp3")
        cajon_ideas_personalizado("Sintetizador de Voz", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### Chat con Documentos")
        archivos = st.file_uploader("Sube documentos (PDF, Word, TXT):", type=["pdf", "txt", "docx"], accept_multiple_files=True, key="chat_pdf_up")
        if archivos:
            texto_completo = "".join([leer_archivo_universal(a, modelo_seleccionado) for a in archivos])[:MAX_TEXTO_IA]
            st.success(f"✅ {len(archivos)} documentos leídos.")
            for msg in st.session_state.pdf_chat_history: st.chat_message(msg["role"]).write(msg["content"])
            if prompt := st.chat_input("Pregúntale a los documentos...", key="chat_pdf_input"):
                st.session_state.pdf_chat_history.append({"role": "user", "content": prompt})
                st.chat_message("user").write(prompt)
                with st.spinner("Buscando..."):
                    res = llamar_ia(modelo_seleccionado, f"Basándote en esto:\n{texto_completo}\nResponde: {prompt}", api_key_gemini, historial_chat=st.session_state.pdf_chat_history)
                    st.session_state.pdf_chat_history.append({"role": "assistant", "content": res})
                    st.chat_message("assistant").write(res)
        cajon_ideas_personalizado("Chat con PDFs", modelo_seleccionado, api_key_gemini)

    with sub[3]:
        st.markdown("#### 🕵️‍♂️ Creador de Agentes IA")
        nombre_agente = st.text_input("Nombre del Agente:", placeholder="Ej: Albert, el Profesor de Física", key="agent_name")
        rol_agente = st.text_area("Personalidad:", placeholder="Ej: Eres un profesor divertido. Explicas usando dibujos animados.", height=100, key="agent_role")
        
        if st.button("✨ Crear / Activar Agente", type="primary", key="btn_create_agent"):
            if nombre_agente and rol_agente:
                st.session_state.agente_activo = {"nombre": nombre_agente, "rol": rol_agente}
                st.session_state.agente_chat_history = []
                st.success(f"Agente '{nombre_agente}' listo.")
                    
        if st.session_state.get("agente_activo"):
            agente = st.session_state.agente_activo
            st.markdown(f"### 💬 Hablando con: **{agente['nombre']}**")
            for msg in st.session_state.agente_chat_history: st.chat_message(msg["role"]).write(msg["content"])
            if prompt := st.chat_input(f"Mensaje para {agente['nombre']}...", key="chat_agent_input"):
                st.session_state.agente_chat_history.append({"role": "user", "content": prompt})
                st.chat_message("user").write(prompt)
                with st.spinner("Pensando..."):
                    prompt_con_rol = f"INSTRUCCIONES DEL SISTEMA (ERES ESTE AGENTE):\n{agente['rol']}\n\nUsuario: {prompt}"
                    res = llamar_ia(modelo_seleccionado, prompt_con_rol, api_key_gemini, historial_chat=st.session_state.agente_chat_history)
                    st.session_state.agente_chat_history.append({"role": "assistant", "content": res})
                    st.chat_message("assistant").write(res)
        cajon_ideas_personalizado("Agentes Custom", modelo_seleccionado, api_key_gemini)

    with sub[4]:
        st.markdown("#### ⚖️ La Balanza: El Debate")
        st.caption("Pon a prueba tus ideas. La IA defenderá una postura y luego la contraria.")
        
        tema_debate = st.text_input("¿Sobre qué quieres debatir?", placeholder="Ej: ¿Debería comprarme un coche eléctrico?", key="debate_q")
        
        if st.button("Iniciar Debate", type="primary", key="btn_debate") and tema_debate:
            with st.spinner("Preparando argumentos..."):
                prompt_debate = f"Actúa como un debate estructurado. El tema es: '{tema_debate}'. " \
                                 f"1. Presenta los 3 argumentos MÁS FUERTES a favor. " \
                                 f"2. Presenta los 3 argumentos MÁS FUERTES en contra. " \
                                 f"3. Concluye con una reflexión neutral."
                st.session_state.debate_resultado = llamar_ia(modelo_seleccionado, prompt_debate, api_key_gemini)
                st.markdown(st.session_state.debate_resultado)
                botones_descarga(st.session_state.debate_resultado, f"Debate_{tema_debate[:20]}")
        cajon_ideas_personalizado("Debate", modelo_seleccionado, api_key_gemini)

    with sub[5]:
        st.markdown("#### 🍳 El Chef Virtual")
        col_c1, col_c2 = st.columns(2)
        with col_c1:
            ingredientes = st.text_area("¿Qué tienes en la nevera?", key="chef_ing")
            if st.button("Dame una receta", key="btn_chef_receta") and ingredientes:
                st.session_state.receta_generada = llamar_ia(modelo_seleccionado, f"Tengo {ingredientes}. Dame una receta.", api_key_gemini)
            if st.session_state.receta_generada: st.info(st.session_state.receta_generada)
        with col_c2:
            preferencias = st.text_input("¿Menú semanal?", key="chef_menu")
            if st.button("Generar Menú + Compra", key="btn_chef_menu") and preferencias:
                st.session_state.menu_generado = llamar_ia(modelo_seleccionado, f"Menú semanal para 4. {preferencias}.", api_key_gemini)
            if st.session_state.menu_generado: st.info(st.session_state.menu_generado)
        cajon_ideas_personalizado("Cocina", modelo_seleccionado, api_key_gemini)

    with sub[6]:
        st.markdown("#### 📖 Cuentacuentos")
        prota = st.text_input("Protagonista:", key="cuento_prota")
        lugar = st.text_input("Lugar mágico:", key="cuento_lugar")
        if st.button("✨ Crear Cuento", key="btn_cuento") and prota and lugar:
            cuento = llamar_ia(modelo_seleccionado, f"Cuento. Protagonista: {prota}, Lugar: {lugar}.", api_key_gemini)
            st.text_area("Cuento:", value=cuento, height=300, key="cuento_text")
            if st.button("🔊 Leer en voz alta", key="btn_cuento_tts"):
                audio = generar_audio_sincrono(cuento, "es-ES-ElviraNeural")
                if audio: st.audio(audio, format="audio/mp3")
        cajon_ideas_personalizado("Cuentos", modelo_seleccionado, api_key_gemini)

    with sub[7]:
        st.markdown("#### Asistente de Código")
        req_code = st.text_area("Pega un código o pide uno nuevo:", key="code_req")
        if st.button("Procesar Código", key="btn_code") and req_code:
            res = llamar_ia(modelo_seleccionado, f"Programador senior. Explica o mejora:\n{req_code}", api_key_gemini)
            st.code(res, language="python")
        cajon_ideas_personalizado("Código", modelo_seleccionado, api_key_gemini)

    with sub[8]:
        st.markdown("#### Traductor")
        texto_trad = st.text_area("Texto a traducir:", key="trad_text")
        idioma = st.selectbox("Idioma destino:", ["Inglés", "Francés", "Alemán", "Italiano"], key="trad_lang")
        if st.button("Traducir", key="btn_trad") and texto_trad:
            st.info(llamar_ia(modelo_seleccionado, f"Traduce al {idioma}: {texto_trad}", api_key_gemini))
        cajon_ideas_personalizado("Traductor", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 2: DOCUMENTOS Y TEXTO
# ==========================================
with macro_tabs[1]:
    sub = st.tabs(["📁 ANÁLISIS", "⚖️ COMPARAR", "📊 OFFICE", "📄 PDF", "🕸️ MAPAS", "📇 QR & CV", "📋 PORTAPAPELES"])
    with sub[0]:
        st.markdown("#### Análisis Avanzado de Documentos")
        archivo = st.file_uploader("Sube un PDF, Word, Excel o TXT", type=["pdf", "txt", "docx", "xlsx"], key="up_dat")
        st.session_state.doc_opcion = st.selectbox("🎯 Selección de análisis:", ["Resumen Ejecutivo", "Extraer Fechas y Cifras", "Buscar Errores", "Crear Test", "Traducir al Inglés"])
        
        if archivo and st.button("🔎 Ejecutar Análisis Seleccionado", key="btn_dat", type="primary"):
            with st.spinner("Analizando..."):
                texto = leer_archivo_universal(archivo, modelo_seleccionado)
                if st.session_state.doc_opcion == "Resumen Ejecutivo": prompt = f"Haz un resumen ejecutivo de este documento:\n{texto}"
                elif st.session_state.doc_opcion == "Extraer Fechas y Cifras": prompt = f"Extrae todas las fechas, nombres y cifras de este texto:\n{texto}"
                elif st.session_state.doc_opcion == "Buscar Errores": prompt = f"Busca contradicciones o datos incoherentes en este texto:\n{texto}"
                elif st.session_state.doc_opcion == "Crear Test": prompt = f"Crea un test de 5 preguntas (con respuestas al final) de este texto:\n{texto}"
                else: prompt = f"Traduce al inglés este documento:\n{texto}"
                
                st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt, api_key_gemini)
                st.text_area("Resultado:", value=st.session_state.analisis_ia, height=400)
                botones_descarga(st.session_state.analisis_ia, "Analisis_Doc")
        cajon_ideas_personalizado("Análisis de Documentos", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### ⚖️ Comparador de Documentos")
        col_c1, col_c2 = st.columns(2)
        with col_c1: doc_original = st.file_uploader("Documento Original (A):", type=["pdf", "txt", "docx"], key="doc_a")
        with col_c2: doc_nuevo = st.file_uploader("Documento Modificado (B):", type=["pdf", "txt", "docx"], key="doc_b")
            
        if doc_original and doc_nuevo and st.button("Comparar Documentos", type="primary", key="btn_compare"):
            with st.spinner("Buscando diferencias..."):
                texto_a = leer_archivo_universal(doc_original, modelo_seleccionado)
                texto_b = leer_archivo_universal(doc_nuevo, modelo_seleccionado)
                prompt_comp = f"Auditor legal. Encuentra las diferencias entre estos dos documentos.\n\nDOC A:\n{texto_a[:10000]}\n\nDOC B:\n{texto_b[:10000]}"
                st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt_comp, api_key_gemini)
                st.text_area("Diferencias Encontradas:", value=st.session_state.analisis_ia, height=400)
                botones_descarga(st.session_state.analisis_ia, "Comparativa_Docs")
        cajon_ideas_personalizado("Comparador de Documentos", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### Creación de Documentos")
        col_form, col_prompt = st.columns([1, 2])
        with col_form: formato = st.radio("Formato final:", ["Word", "Excel", "PowerPoint"], key="office_format")
        with col_prompt: prompt_office = st.text_area("¿Qué documento quieres crear?", placeholder="Ej: Un presupuesto para un cliente...", key="office_prompt")
        
        if st.button("Generar Documento", key="btn_office", type="primary") and prompt_office:
            with st.spinner("Generando..."):
                texto = llamar_ia(modelo_seleccionado, prompt_office, api_key_gemini)
                st.session_state.office_text = texto
                if formato == "Word": st.download_button("📥 Descargar Word", generar_word(texto), "doc.docx")
                elif formato == "Excel": st.download_button("📥 Descargar Excel", generar_excel(texto), "doc.xlsx")
                else: st.download_button("📥 Descargar PPTX", generar_pptx("Título", texto), "doc.pptx")
        if st.session_state.office_text:
            boton_superpoder(st.session_state.office_text, "Haz este documento un 200% más persuasivo", "office", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("Creación de Documentos", modelo_seleccionado, api_key_gemini)

    with sub[3]:
        st.markdown("#### PDF Pro")
        titulo_pdf = st.text_input("Título del documento:", key="pdf_title")
        contenido_pdf = st.text_area("Contenido a incluir:", key="pdf_content")
        if st.button("Generar PDF", key="btn_pdf", type="primary") and titulo_pdf and contenido_pdf:
            texto = llamar_ia(modelo_seleccionado, contenido_pdf, api_key_gemini)
            st.session_state.pdf_text = texto
            st.download_button("📥 Descargar PDF", generar_pdf_pro(titulo_pdf, texto), f"{titulo_pdf}.pdf")
        if st.session_state.pdf_text:
            boton_superpoder(st.session_state.pdf_text, "Añade resumen ejecutivo y conclusiones", "pdf", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("PDF Pro", modelo_seleccionado, api_key_gemini)

    with sub[4]:
        st.markdown("#### Mapas Mentales")
        col_tema, col_modo = st.columns([2, 1])
        with col_tema: tema = st.text_input("Tema Central:", placeholder="Ej: Redes informáticas", key="map_tema")
        with col_modo: modo_infantil = st.checkbox("🎨 Modo Infantil (Emojis)", value=False, key="map_kids")
        
        if st.button("🎯 Trazar Mapa", type="primary", key="btn_mapa") and tema:
            prompt = f"Genera código Mermaid para un mapa mental de '{tema}'."
            if modo_infantil: prompt += " Usa emojis y colores 'classDef rojo fill:#FF6B6B'."
            codigo = llamar_ia(modelo_seleccionado, prompt, api_key_gemini)
            st.session_state.mermaid_code = html.escape(codigo.replace("```mermaid", "").replace("```", "").strip())
        if st.session_state.mermaid_code:
            st.html(f"<div class='mermaid'>{st.session_state.mermaid_code}</div><script type='module'>import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs'; mermaid.initialize({{ startOnLoad: true }});</script>")
            boton_superpoder(f"Añade 3 ramas innovadoras a este mapa mental (Mermaid):\n{st.session_state.mermaid_code}", "Expandir Mapa", "mapa", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("Mapas Mentales", modelo_seleccionado, api_key_gemini)

    with sub[5]:
        col_qr1, col_qr2 = st.columns(2)
        with col_qr1:
            st.markdown("#### Generador QR")
            txt_qr = st.text_input("Texto o URL para el QR:", key="qr_text")
            if st.button("Generar QR", key="btn_gen_qr") and txt_qr:
                qr = qrcode.QRCode(version=1, box_size=10, border=5); qr.add_data(txt_qr); qr.make(fit=True)
                buf = BytesIO(); qr.make_image().save(buf, format="PNG"); st.image(buf.getvalue())
        with col_qr2:
            st.markdown("#### Lector QR")
            img_qr = st.file_uploader("Sube una imagen de QR:", type=["png", "jpg", "jpeg"], key="qr_img")
            if img_qr:
                if not HAS_CV2: st.error("Instala opencv-python.")
                else:
                    img = Image.open(img_qr).convert('RGB')
                    detector = cv2.QRCodeDetector()
                    data, _, _ = detector.detectAndDecode(np.array(img))
                    if data: st.success(f"Contenido: **{data}**")
                    else: st.warning("No se detectó ningún QR.")
        cajon_ideas_personalizado("QR y CV", modelo_seleccionado, api_key_gemini)

    with sub[6]:
        st.markdown("#### Portapapeles Universal")
        texto_uni = st.text_area("Pega cualquier texto aquí:", key="clip_text")
        instruccion = st.text_input("¿Qué quieres que haga la IA con este texto?", key="clip_inst")
        if st.button("Procesar Texto", key="btn_clip") and texto_uni and instruccion:
            res = llamar_ia(modelo_seleccionado, f"Instrucción: {instruccion}\n\nTexto: {texto_uni}", api_key_gemini)
            st.session_state.clip_text = res
            st.info(res)
        if st.session_state.clip_text:
            boton_superpoder(st.session_state.clip_text, "Convierte esto en un artículo de blog", "clip", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("Portapapeles", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 3: WEB Y ACTUALIDAD
# ==========================================
with macro_tabs[2]:
    sub = st.tabs(["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV & EVENTOS", "📊 WEB A EXCEL", "🎥 YOUTUBE", "📰 LECTOR DE WEBS"])
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
        cajon_ideas_personalizado("Buscador Web", modelo_seleccionado, api_key_gemini)
                    
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
                boton_superpoder(str(st.session_state.noticias_deportes), "Analiza la noticia más impactante", "noticias", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Noticias de Deportes", modelo_seleccionado, api_key_gemini)
            
        with tab_dep2:
            st.markdown("##### ⚡ Resultados y Directo")
            evento_directo = st.text_input("¿Qué resultados buscas?", placeholder="Ej: Resultados de LaLiga hoy", key="dep_results_q")
            if st.button("Buscar Resultados", type="primary", key="btn_res_directo") and evento_directo:
                with st.spinner("Buscando marcadores..."):
                    res_web = buscar_en_internet(evento_directo)
                    contexto = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_web])
                    st.session_state.dep_resultados = llamar_ia(modelo_seleccionado, f"Dame los resultados de: '{evento_directo}'. Datos: {contexto}", api_key_gemini)
                    st.info(st.session_state.dep_resultados)
            if st.session_state.dep_resultados:
                boton_superpoder(st.session_state.dep_resultados, "Haz un análisis táctico de élite", "tactico", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Resultados Deportivos", modelo_seleccionado, api_key_gemini)
            
        with tab_dep3:
            st.markdown("##### 📊 Análisis Táctico y Clasificaciones")
            col_eq1, col_eq2 = st.columns(2)
            with col_eq1: equipo_local = st.text_input("Equipo Local:", placeholder="Ej: Real Madrid", key="dep_eq1")
            with col_eq2: equipo_visitante = st.text_input("Equipo Visitante:", placeholder="Ej: FC Barcelona", key="dep_eq2")
            if st.button("⚔️ Analizar Partido", key="btn_analisis_partido") and equipo_local and equipo_visitante:
                with st.spinner("Analizando rachas..."):
                    prompt_partido = f"Analiza el partido {equipo_local} vs {equipo_visitante}. Pronóstico final."
                    st.session_state.partido_analisis = llamar_ia(modelo_seleccionado, prompt_partido, api_key_gemini)
                    st.info(st.session_state.partido_analisis)
            if st.session_state.partido_analisis:
                boton_superpoder(st.session_state.partido_analisis, "Dame el jugador clave y las apuestas seguras", "partido", modelo_seleccionado, api_key_gemini)
            st.markdown("---")
            liga_tabla = st.selectbox("Consultar Clasificación:", ["LaLiga EA Sports", "Premier League", "NBA", "Euroliga", "Fórmula 1"], key="sel_tabla")
            if st.button("🏆 Ver Tabla", key="btn_ver_tabla"):
                with st.spinner("Buscando clasificación..."):
                    prompt_tabla = f"Dame la clasificación actual de {liga_tabla} (Posición, Equipo/Piloto, Puntos)."
                    st.info(llamar_ia(modelo_seleccionado, prompt_tabla, api_key_gemini))
            cajon_ideas_personalizado("Análisis Deportivo", modelo_seleccionado, api_key_gemini)
            
        with tab_dep4:
            st.markdown("##### 🏋️ Entrenamiento y Nutrición")
            objetivo = st.text_input("¿Qué necesitas?", placeholder="Ej: Rutina para ganar masa muscular", key="dep_train_q")
            deporte_pract = st.selectbox("Deporte principal:", ["Fútbol", "Running", "Gimnasio", "Ciclismo", "Baloncesto", "Tenis"], key="dep_train_s")
            if st.button("Generar Plan", key="btn_plan_deporte") and objetivo:
                with st.spinner("Diseñando plan..."):
                    st.session_state.plan_deporte = llamar_ia(modelo_seleccionado, f"Entrenador personal. Deporte: {deporte_pract}. Objetivo: '{objetivo}'. Dame un plan semanal.", api_key_gemini)
                    st.info(st.session_state.plan_deporte)
            if st.session_state.plan_deporte:
                boton_superpoder(st.session_state.plan_deporte, "Añade suplementos y estiramientos preventivos", "entrenamiento", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Entrenamiento", modelo_seleccionado, api_key_gemini)
            
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
                boton_superpoder(st.session_state.cine_rec, "Recomienda 3 joyas ocultas", "cine", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Cine y Series", modelo_seleccionado, api_key_gemini)
            
        with tab_tv2:
            canal_tv = st.text_input("¿Qué canal buscas?", placeholder="Ej: Programación de esta noche en Telecinco", key="tv_chan")
            if st.button("Buscar Programación", key="btn_tv_prog") and canal_tv:
                with st.spinner("Buscando parrilla..."):
                    res_web_tv = buscar_en_internet(canal_tv)
                    contexto_tv = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_web_tv])
                    st.session_state.tv_prog = llamar_ia(modelo_seleccionado, f"Programación de hoy para: '{canal_tv}'. Datos: {contexto_tv}", api_key_gemini)
                    st.info(st.session_state.tv_prog)
            if st.session_state.tv_prog:
                boton_superpoder(st.session_state.tv_prog, "Dime el programa más recomendado", "tv", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Programación TV", modelo_seleccionado, api_key_gemini)
            
        with tab_tv3:
            evento_buscar = st.text_input("¿Qué evento buscas?", placeholder="Ej: Concierto de Taylor Swift", key="tv_event")
            ciudad = st.text_input("¿En qué ciudad?", placeholder="Ej: Madrid", key="tv_city")
            if st.button("Buscar Entradas", key="btn_entradas") and evento_buscar:
                with st.spinner("Buscando disponibilidad..."):
                    res_entradas = buscar_en_internet(f"entradas {evento_buscar} {ciudad}")
                    contexto_entradas = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in res_entradas])
                    st.session_state.entradas = llamar_ia(modelo_seleccionado, f"Info de: '{evento_buscar}' en {ciudad}. Datos: {contexto_entradas}", api_key_gemini)
                    st.info(st.session_state.entradas)
            if st.session_state.entradas:
                boton_superpoder(st.session_state.entradas, "Plan de viaje de 24h alrededor del evento", "entradas", modelo_seleccionado, api_key_gemini)
            cajon_ideas_personalizado("Entradas de Eventos", modelo_seleccionado, api_key_gemini)
            
        with tab_tv4:
            juego = st.radio("Elige el juego:", ["Adivinar la película", "Curiosidades de series", "Preguntas de deportes"], horizontal=True, key="tv_trivia_game")
            if st.button("¡Jugar!", key="btn_jugar_trivia"):
                with st.spinner("Generando juego..."):
                    st.session_state.trivia = llamar_ia(modelo_seleccionado, f"Juego de 3 preguntas trivia sobre '{juego}'. Da respuestas al final.", api_key_gemini)
                    st.info(st.session_state.trivia)
            cajon_ideas_personalizado("Trivia", modelo_seleccionado, api_key_gemini)

    with sub[3]:
        st.markdown("#### 📊 Extractor de Tablas Web a Excel")
        st.caption("Pega la URL de una página que tenga una tabla (ej. clasificación, precios) y la IA extraerá los datos.")
        
        url_tabla = st.text_input("Pega la URL de la página web:", key="table_url")
        if st.button("Extraer Tabla", type="primary", key="btn_table") and url_tabla:
            with st.spinner("Buscando tablas en la web..."):
                try:
                    tablas = pd.read_html(url_tabla)
                    if len(tablas) > 0:
                        st.success(f"¡Se han encontrado {len(tablas)} tablas!")
                        tabla_elegida = st.selectbox("Selecciona la tabla que quieres descargar:", range(len(tablas)), format_func=lambda x: f"Tabla {x+1}")
                        
                        df = tablas[tabla_elegida]
                        st.dataframe(df)
                        
                        out = BytesIO()
                        with pd.ExcelWriter(out, engine='openpyxl') as writer:
                            df.to_excel(writer, index=False, sheet_name='Datos_Web')
                        st.download_button("📥 Descargar en Excel", out.getvalue(), "tabla_web.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
                    else:
                        st.warning("No se encontraron tablas en esa página.")
                except Exception as e:
                    st.error(f"Error al leer la web: {e}. Asegúrate de que la URL es correcta y contiene tablas.")
        cajon_ideas_personalizado("Extractor Web", modelo_seleccionado, api_key_gemini)

    with sub[4]:
        st.markdown("#### 🎥 Resumidor de YouTube")
        url_yt = st.text_input("Pega la URL del vídeo:", key="yt_url")
        if st.button("Extraer y Resumir", key="btn_yt") and url_yt:
            try:
                video_id = url_yt.split("v=")[-1].split("&")[0]
                transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['es', 'en'])
                resumen = llamar_ia(modelo_seleccionado, f"Resume esto:\n{' '.join([t['text'] for t in transcript])[:5000]}", api_key_gemini)
                st.session_state.yt_text = resumen
                st.text_area("Resumen:", value=resumen, height=300)
            except Exception as e: st.error(f"Error: {e}")
        if st.session_state.yt_text:
            boton_superpoder(st.session_state.yt_text, "Saca 3 ideas de negocio de este resumen", "yt", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("YouTube", modelo_seleccionado, api_key_gemini)
        
    with sub[5]:
        st.markdown("#### 📰 Lector de Páginas Web")
        url_web = st.text_input("Pega el enlace de una noticia o web:", key="web_url")
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
            boton_superpoder(st.session_state.web_text, "Opinión crítica sobre el sesgo de este artículo", "web", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("Lector de Webs", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 4: MULTIMEDIA
# ==========================================
with macro_tabs[3]:
    sub = st.tabs(["🖼️ ANÁLISIS MULTIMEDIA", "🎨 GENERADOR VISUAL", "🛡️ CREADOR DE LOGOS", "📥 DESCARGADOR"])
    with sub[0]:
        st.markdown("#### Análisis Inteligente de Archivos")
        with st.expander("💡 Cajón de Ideas (¿Qué puedo hacer con este archivo?)"):
            st.write("- **Si es foto:** Extraer texto (OCR) o describir ropa.")
            st.write("- **Si es audio:** Resumen o traducción de entrevista.")
            st.write("- **Si es vídeo:** Qué escenarios aparecen o de qué hablan.")
            
        archivo_multi = st.file_uploader("📎 Sube tu archivo aquí", type=["png", "jpg", "jpeg", "webp", "mp4", "mov", "mp3", "wav"], key="up_multi")
        if archivo_multi:
            nombre = archivo_multi.name.lower()
            if nombre.endswith((".png", ".jpg", ".jpeg", ".webp")):
                st.image(archivo_multi, use_container_width=True)
                mime_type = "image/jpeg" if nombre.endswith((".jpg", ".jpeg")) else f"image/{nombre.split('.')[-1]}"
                opcion_img = st.radio("Opciones para Imagen:", ["👁️ Describir", "📝 Extraer texto (OCR)", "🎨 Prompt para recrearla", "🕵️‍♂️ Analizar si es IA"], key="radio_img")
                if st.button("Ejecutar Acción en Imagen", type="primary", key="btn_img_action"):
                    if "Describir" in opcion_img: prompt = "Describe detalladamente esta imagen."
                    elif "texto" in opcion_img: prompt = "Extrae todo el texto visible (OCR)."
                    elif "prompt" in opcion_img: prompt = "Escribe un prompt en inglés para Midjourney que replique esta imagen."
                    else: prompt = "Analiza la imagen y dime si parece real o generada por IA."
                    with st.spinner("Procesando..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            elif nombre.endswith((".mp3", ".wav")):
                st.audio(archivo_multi)
                mime_type = "audio/mp3" if nombre.endswith(".mp3") else "audio/wav"
                opcion_aud = st.radio("Opciones para Audio:", ["📝 Transcribir", "📋 Resumen", "🌍 Traducir al inglés", "🎭 Analizar tono emocional"], key="radio_aud")
                if st.button("Ejecutar Acción en Audio", type="primary", key="btn_aud_action"):
                    if "Transcribir" in opcion_aud: prompt = "Transcribe exactamente lo que se dice."
                    elif "Resumen" in opcion_aud: prompt = "Haz un resumen de lo que se habla."
                    elif "Traducir" in opcion_aud: prompt = "Traduce al inglés lo que se dice."
                    else: prompt = "Escucha este audio y dime qué tono emocional tiene y por qué."
                    with st.spinner("Procesando..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            elif nombre.endswith((".mp4", ".mov")):
                st.video(archivo_multi)
                mime_type = "video/mp4"
                opcion_vid = st.radio("Opciones para Vídeo:", ["📋 Resumir", "🔍 Describir escenas", "🗣️ Analizar diálogos"], key="radio_vid")
                if st.button("Ejecutar Acción en Vídeo", type="primary", key="btn_vid_action"):
                    if "Resumir" in opcion_vid: prompt = "Resume el contenido de este vídeo."
                    elif "escenas" in opcion_vid: prompt = "Describe detalladamente las escenas."
                    else: prompt = "Analiza los diálogos y el tono del vídeo."
                    with st.spinner("Procesando vídeo..."): st.session_state.analisis_ia = analizar_archivo_con_gemini(archivo_multi.getvalue(), mime_type, prompt, api_key_gemini)
            
            if st.session_state.analisis_ia:
                st.success("✅ Análisis completado:")
                st.info(st.session_state.analisis_ia)
                botones_descarga(st.session_state.analisis_ia, "Analisis_Multimedia")
                boton_superpoder(st.session_state.analisis_ia, "Información forense de este análisis", "multi", modelo_seleccionado, api_key_gemini)
                
        cajon_ideas_personalizado("Análisis Multimedia", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### Generador Visual")
        prompt_img = st.text_input("Describe la imagen (en español):", key="img_prompt")
        tipo_img = st.radio("Estilo:", ["Foto/Realista", "Dibujo para Colorear"], horizontal=True, key="img_style")
        
        if st.button("Generar Imagen", key="btn_gen_visual") and prompt_img:
            with st.spinner("Traduciendo y generando imagen..."):
                prompt_traducido = llamar_ia(modelo_seleccionado, f"Traduce el siguiente texto al inglés para un generador de imágenes. Devuelve SOLO la traducción, sin comillas ni explicaciones: '{prompt_img}'", api_key_gemini)
                prompt_en_ingles = prompt_traducido.replace('"', '').replace("'", "").strip()
                st.session_state.img_prompt_en = prompt_en_ingles
                
                if "Colorear" in tipo_img:
                    prompt_final = prompt_en_ingles + ", coloring book page, thick black outlines, pure white background, no color"
                else:
                    prompt_final = prompt_en_ingles
                
                st.session_state.img_url = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_final)}?width=1024&height=1024&nologo=true&enhance=true"
        
        if st.session_state.img_url:
            st.image(st.session_state.img_url, use_container_width=True)
            st.caption(f"Prompt enviado a la IA: *{st.session_state.img_prompt_en}*")
            boton_superpoder(f"Genera un prompt cinematográfico para: {prompt_img}", "Crear Prompt Alternativo", "visual", modelo_seleccionado, api_key_gemini)
        cajon_ideas_personalizado("Generador Visual", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### 🛡️ Generador de Logotipos e Identidad")
        col_logo1, col_logo2 = st.columns(2)
        with col_logo1:
            nombre_marca = st.text_input("Nombre de la marca:", key="logo_name")
            estilo_logo = st.selectbox("Estilo de diseño:", ["Minimalista y limpio", "Corporativo moderno", "3D e Impactante", "Vintage y Retro", "Cyberpunk / Neón"], key="logo_style")
        with col_logo2:
            color_logo = st.color_picker("Color principal de la marca:", "#d4af37", key="logo_color")
            icono_logo = st.text_input("Icono o símbolo (opcional):", placeholder="Ej: Un lobo, una taza...", key="logo_icon")
            
        if st.button("🛡️ Generar Logo", type="primary", key="btn_gen_logo") and nombre_marca:
            with st.spinner("Diseñando logotipo..."):
                prompt_logo = f"Professional brand logo design for '{nombre_marca}', {estilo_logo} style, primary color {color_logo}, "
                if icono_logo: prompt_logo += f"with an icon of {icono_logo}, "
                prompt_logo += "flat vector graphic, minimalist, clean white background, centered, high quality, no text, corporate identity"
                url_logo = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_logo)}?width=1024&height=1024&nologo=true"
                st.image(url_logo, caption=f"Propuesta para {nombre_marca}", use_container_width=True)
                st.markdown(f"[Descargar Logo en Alta Resolución]({url_logo})")
        cajon_ideas_personalizado("Creación de Logos", modelo_seleccionado, api_key_gemini)

    with sub[3]:
        st.markdown("#### Descargador")
        url_dl = st.text_input("Pega la URL del vídeo/audio (YouTube, TikTok...):", key="dl_url")
        if st.button("⬇️ Descargar", type="primary", key="btn_dl") and url_dl:
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
        cajon_ideas_personalizado("Descargas", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 5: ESTUDIO Y OFICIOS
# ==========================================
with macro_tabs[4]:
    sub = st.tabs(["🌐 REDES Y SISTEMAS", "💇‍♀️ PELUQUERÍA", "📚 DEBERES Y ESTUDIO", "🦺 PRL (RIESGOS LABORALES)"])
    with sub[0]:
        st.markdown("#### 🌐 Herramientas de Redes")
        col_r1, col_r2 = st.columns(2)
        with col_r1:
            ip = st.text_input("IP/CIDR (Ej: 192.168.1.0/24)", key="net_ip")
            if st.button("Calcular Red", key="btn_net") and ip:
                st.info(llamar_ia(modelo_seleccionado, f"Calcula máscara, rango, broadcast y hosts para: {ip}", api_key_gemini))
        with col_r2:
            tarea = st.text_input("¿Qué necesitas hacer?", key="net_task")
            so = st.selectbox("Sistema:", ["Windows", "Linux", "Cisco"], key="net_os")
            if st.button("Buscar Comando", key="btn_net_cmd") and tarea:
                st.info(llamar_ia(modelo_seleccionado, f"Comando para {so}: {tarea}. Explícalo.", api_key_gemini))
        
        if st.button("Comenzar Entrevista (Nivel CCNA)", key="btn_net_interview"):
            st.info(llamar_ia(modelo_seleccionado, "Hazme 3 preguntas de nivel CCNA.", api_key_gemini))
        cajon_ideas_personalizado("Redes", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### 💇‍♀️ Asistente de Peluquería")
        col_p1, col_p2 = st.columns(2)
        with col_p1:
            actual = st.text_input("Color actual:", key="hair_actual")
            deseado = st.text_input("Color deseado:", key="hair_desired")
            estado = st.selectbox("Estado del pelo:", ["Sano", "Teñido", "Decolorado", "Dañado"], key="hair_state")
            if st.button("Calcular Fórmula", key="btn_hair_formula") and actual and deseado:
                st.info(llamar_ia(modelo_seleccionado, f"Fórmula de {actual} a {deseado}. Pelo {estado}.", api_key_gemini))
        with col_p2:
            servicios = st.text_area("Servicios:", key="hair_services")
            tiempo = st.number_input("Horas:", min_value=1.0, key="hair_time")
            if st.button("Calcular Presupuesto", key="btn_hair_budget") and servicios:
                st.info(llamar_ia(modelo_seleccionado, f"Presupuesto para: {servicios}. Tiempo: {tiempo}h.", api_key_gemini))
        cajon_ideas_personalizado("Peluquería", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### 📚 Asistente de Deberes")
        nivel = st.selectbox("Nivel:", ["Primaria", "ESO", "Bachillerato", "Universidad"], key="home_nivel")
        materia = st.selectbox("Materia:", ["Matemáticas", "Lengua", "Ciencias", "Historia", "Idiomas"], key="home_materia")
        duda = st.text_area("Duda:", key="home_duda")
        if st.button("Explicar Duda", key="btn_home") and duda:
            st.info(llamar_ia(modelo_seleccionado, f"Profesor. Nivel: {nivel}. Materia: {materia}. Explica: {duda}", api_key_gemini))
        cajon_ideas_personalizado("Deberes", modelo_seleccionado, api_key_gemini)

    with sub[3]:
        st.markdown("#### 🦺 Asistente de Prevención de Riesgos Laborales")
        tipo_prl = st.selectbox("¿Qué necesitas generar?", ["Charla de 5 Minutos", "Evaluación de Riesgos", "Recomendación de EPIs", "Consulta Legal / Normativa"], key="prl_tipo")
        contexto_prl = st.text_area("Describe la tarea o duda:", placeholder="Ej: Trabajo en altura con escalera de tijera.", key="prl_contexto")
        
        if st.button("🦺 Generar Documento PRL", type="primary", key="btn_prl") and contexto_prl:
            if "Charla" in tipo_prl: prompt_prl = f"Charla de 5 min de PRL sobre: '{contexto_prl}'."
            elif "Evaluación" in tipo_prl: prompt_prl = f"Evaluación de riesgos simplificada para: '{contexto_prl}'."
            elif "EPIs" in tipo_prl: prompt_prl = f"EPIs obligatorios para: '{contexto_prl}'."
            else: prompt_prl = f"Consulta legal en PRL (España): '{contexto_prl}'."
            
            st.session_state.prl_resultado = llamar_ia(modelo_seleccionado, prompt_prl, api_key_gemini)
            st.text_area("Documento:", value=st.session_state.prl_resultado, height=400)
            botones_descarga(st.session_state.prl_resultado, "Documento_PRL")
        cajon_ideas_personalizado("PRL", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 6: VIAJES Y HOGAR
# ==========================================
with macro_tabs[5]:
    sub = st.tabs(["🌍 VIAJES", "💰 FINANZAS", "✉️ CARTAS Y MENSAJES"])
    with sub[0]:
        st.markdown("#### 🌍 Planificador de Viajes")
        destino = st.text_input("Destino:", key="travel_dest")
        col_v1, col_v2 = st.columns(2)
        with col_v1: dias = st.number_input("Días:", min_value=1, value=3, key="travel_days")
        with col_v2: compania = st.selectbox("Compañía:", ["Pareja", "Familia", "Amigos", "Solo/a"], key="travel_comp")
        if st.button("✈️ Generar Itinerario", type="primary", key="btn_travel") and destino:
            st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, f"Itinerario {dias} días en {destino} con {compania}.", api_key_gemini)
            st.info(st.session_state.analisis_ia)
            botones_descarga(st.session_state.analisis_ia, f"Viaje_{destino}")
        cajon_ideas_personalizado("Viajes", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### 💰 Finanzas del Hogar")
        col_f1, col_f2 = st.columns(2)
        with col_f1:
            st.markdown("**Ahorro**")
            ingresos = st.number_input("Ingresos (€):", key="fin_in")
            gastos = st.number_input("Gastos (€):", key="fin_out")
            if st.button("Calcular Ahorro", key="btn_fin_save"):
                ahorro = ingresos - gastos
                if ahorro > 0: st.success(f"Ahorras **{ahorro}€**.")
                else: st.error(f"Faltan **{abs(ahorro)}€**.")
        with col_f2:
            st.markdown("**Hipoteca**")
            prestamo = st.number_input("Préstamo (€):", value=150000, key="fin_mort")
            interes = st.number_input("Interés (%):", value=3.0, key="fin_int")
            años = st.number_input("Años:", value=30, key="fin_years")
            if st.button("Calcular Cuota", key="btn_fin_mort"):
                i = (interes / 100) / 12; n = años * 12
                cuota = prestamo * (i * (1 + i)**n) / ((1 + i)**n - 1)
                st.success(f"Cuota: **{cuota:.2f}€**.")
        cajon_ideas_personalizado("Finanzas", modelo_seleccionado, api_key_gemini)

    with sub[2]:
        st.markdown("#### ✉️ Cartas y Mensajes")
        tipo_mensaje = st.selectbox("Tipo:", ["Reyes Magos", "Papá Noel", "Discurso cumpleaños", "Carta formal de reclamación"], key="letter_type")
        detalles = st.text_area("Detalles:", key="letter_details")
        if st.button("✍️ Redactar", type="primary", key="btn_letter") and detalles:
            st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, f"Escribe {tipo_mensaje} con estos datos: {detalles}.", api_key_gemini)
            st.text_area("Resultado:", value=st.session_state.analisis_ia, height=300)
            botones_descarga(st.session_state.analisis_ia, "Mensaje")
        cajon_ideas_personalizado("Cartas y Mensajes", modelo_seleccionado, api_key_gemini)

# ==========================================
# CATEGORÍA 7: MÚSICA Y ARTE
# ==========================================
with macro_tabs[6]:
    sub = st.tabs(["🎵 CREADOR DE CANCIONES", "🎨 POESÍA Y RIMAS"])
    with sub[0]:
        st.markdown("#### 🎵 Estudio de Creación Musical")
        col_m1, col_m2 = st.columns(2)
        with col_m1: tema_cancion = st.text_input("¿De qué trata?", key="song_theme")
        with col_m2: genero = st.selectbox("Género:", ["Pop", "Rock", "Rap", "Balada"], key="song_genre")
        tono = st.select_slider("Tono:", options=["Triste", "Melancólico", "Alegre", "Épico"], key="song_tone")
        
        if st.button("🎼 Escribir Canción", type="primary", key="btn_song") and tema_cancion:
            letra = llamar_ia(modelo_seleccionado, f"Canción {genero} sobre '{tema_cancion}', tono {tono}.", api_key_gemini)
            st.session_state.letra_cancion = letra
            titulo = llamar_ia(modelo_seleccionado, f"Título para esta letra:\n{letra[:500]}", api_key_gemini).strip()
            st.session_state.titulo_cancion = titulo
            st.session_state.portada_cancion = f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Album cover ' + genero + ' ' + tema_cancion)}?width=1024&height=1024&nologo=true"
                
        if st.session_state.letra_cancion:
            st.markdown(f"### 🎶 {st.session_state.titulo_cancion}")
            col_portada, col_letra = st.columns([1, 2])
            with col_portada:
                st.image(st.session_state.portada_cancion, use_container_width=True)
                voz_cancion = st.selectbox("Voz:", ["es-MX-JorgeNeural", "es-ES-ElviraNeural"], key="song_voice")
                if st.button("🔊 Generar Audio", key="btn_song_audio"):
                    audio_letra = generar_audio_sincrono(st.session_state.letra_cancion, voz_cancion)
                    if audio_letra:
                        st.audio(audio_letra, format="audio/mp3")
                        st.download_button("💾 MP3", audio_letra, f"{st.session_state.titulo_cancion}.mp3")
            with col_letra:
                st.text_area("Letra:", value=st.session_state.letra_cancion, height=500)
                botones_descarga(st.session_state.letra_cancion, "Letra")
        cajon_ideas_personalizado("Música", modelo_seleccionado, api_key_gemini)

    with sub[1]:
        st.markdown("#### 🎨 Poesía y Rimas")
        tema_poema = st.text_input("¿Sobre qué quieres el poema?", key="poem_theme")
        estilo_poema = st.selectbox("Estilo:", ["Romántico", "Divertido", "Épico", "Haiku"], key="poem_style")
        if st.button("✒️ Escribir Poema", type="primary", key="btn_poem") and tema_poema:
            poema = llamar_ia(modelo_seleccionado, f"Poema {estilo_poema} sobre: '{tema_poema}'.", api_key_gemini)
            st.text_area("Poema:", value=poema, height=300)
            botones_descarga(poema, "Poema")
            if st.button("🔊 Recitar Poema", key="btn_poem_audio"):
                audio_poema = generar_audio_sincrono(poema, "es-ES-ElviraNeural")
                if audio_poema: st.audio(audio_poema, format="audio/mp3")
        cajon_ideas_personalizado("Poesía", modelo_seleccionado, api_key_gemini)
