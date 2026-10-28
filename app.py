import streamlit as st
import streamlit.components.v1 as components
from streamlit_markmap import markmap 
import ollama
import asyncio
import edge_tts
import pypdf
from docx import Document
import openpyxl
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Inches
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
import sqlite3
import datetime
import concurrent.futures
import json

try: import cv2; HAS_CV2 = True
except ImportError: HAS_CV2 = False
try: import yt_dlp; HAS_YTDLP = True
except ImportError: HAS_YTDLP = False
try: from ddgs import DDGS
except ImportError:
    try: from duckduckgo_search import DDGS
    except ImportError: DDGS = None
try: from google import genai; from google.genai import types; HAS_GEMINI = True
except ImportError: HAS_GEMINI = False
try: import openai; HAS_OPENAI = True
except ImportError: HAS_OPENAI = False
try: import anthropic; HAS_ANTHROPIC = True
except ImportError: HAS_ANTHROPIC = False

# ==========================================
# 1. CONFIGURACIÓN INICIAL Y CONSTANTES
# ==========================================
st.set_page_config(page_title="Xavilu Network Pro", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")
MAX_FILE_MB = 1000
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-3.8-flash"

if os.name == 'nt':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# ==========================================
# 2. BASES DE DATOS (PLANTAS E HISTORIAL)
# ==========================================
def init_plant_db():
    conn = sqlite3.connect('xavilu_planta.db')
    c = conn.cursor()
    c.execute('''CREATE TABLE IF NOT EXISTS bascula (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, matricula TEXT, residuo TEXT, peso REAL, operacion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS lotes (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, residuo TEXT, peso_in REAL, estado TEXT, peso_out REAL)''')
    c.execute('''CREATE TABLE IF NOT EXISTS mantenimiento (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, maquina TEXT, descripcion TEXT, estado TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS stock (id INTEGER PRIMARY KEY AUTOINCREMENT, nombre TEXT, categoria TEXT, cantidad REAL, unidad TEXT, stock_minimo REAL, ubicacion TEXT, fecha_actualizacion TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS compras (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, proveedor TEXT, material TEXT, cantidad REAL, precio_unitario REAL, total REAL, archivo TEXT)''')
    c.execute('''CREATE TABLE IF NOT EXISTS ventas (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, cliente TEXT, material TEXT, cantidad REAL, precio_unitario REAL, total REAL, archivo TEXT)''')
    conn.commit(); conn.close()
init_plant_db()

def get_db_data(table):
    try:
        conn = sqlite3.connect('xavilu_planta.db')
        df = pd.read_sql_query(f"SELECT * FROM {table}", conn)
        conn.close(); return df
    except: return pd.DataFrame()

def insert_db(query, params):
    try:
        conn = sqlite3.connect('xavilu_planta.db')
        conn.execute(query, params)
        conn.commit(); conn.close()
    except Exception as e: st.error(f"Error DB: {e}")

def update_stock(item_id, nueva_cantidad):
    try:
        conn = sqlite3.connect('xavilu_planta.db')
        fecha = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        conn.execute("UPDATE stock SET cantidad = ?, fecha_actualizacion = ? WHERE id = ?", (nueva_cantidad, fecha, item_id))
        conn.commit(); conn.close()
    except Exception as e: st.error(f"Error al actualizar: {e}")

def init_db():
    conn = sqlite3.connect('xavilu_memoria.db')
    conn.execute('CREATE TABLE IF NOT EXISTS historial (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, tipo TEXT, contenido TEXT)')
    conn.commit(); conn.close()
def guardar_en_db(tipo, contenido):
    try:
        conn = sqlite3.connect('xavilu_memoria.db')
        conn.execute("INSERT INTO historial (fecha, tipo, contenido) VALUES (?, ?, ?)", (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo, contenido))
        conn.commit(); conn.close()
    except: pass
def obtener_de_db():
    try:
        conn = sqlite3.connect('xavilu_memoria.db')
        rows = conn.execute("SELECT fecha, tipo, contenido FROM historial ORDER BY id DESC LIMIT 5").fetchall()
        conn.close(); return rows
    except: return []
init_db()

# ==========================================
# 3. ESTADO DE SESIÓN (VARIABLES GLOBALES)
# ==========================================
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
    "doc_opcion": "Resumen Ejecutivo", "global_context": "", "global_file_name": "",
    "agente_activo": None, "agente_chat_history": [], "debate_resultado": "", "img_prompt_en": "",
    "podcast_script": "", "podcast_ideas": [], "sidebar_visible": True, "whatsapp_number": "",
    "office_borrador": "", "pdf_borrador": "", "plant_borrador": "", "pod_tema_input": "", "pod_sugs": [],
    "api_key_gemini_manual": "", "api_key_groq_manual": "", "api_key_silicon_manual": "", "api_key_mistral_manual": "", "api_key_openrouter_manual": "",
    # XaviWork
    "xw_mapa": "", "xw_ppt": "", "xw_informe": "", "xw_tarjetas": "", "xw_cuestionario": "", "xw_tabla": "", 
    "xw_podcast": "", "xw_podcast_audio": None, "xw_timeline": "", "xw_glosario": "", "xw_dafo": "", "xw_cover_img": None, "xw_copilot_history": [],
    "xw_multi_word": "", "xw_multi_excel": "", "xw_multi_ppt": "", "xw_multi_pdf": "",
    # Facturas IA
    "proveedor_temp": "", "material_temp": "", "cant_temp": 0.0, "precio_temp": 0.0,
    "cliente_temp": "", "material_v_temp": "", "cant_v_temp": 0.0, "precio_v_temp": 0.0
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# ==========================================
# 4. CSS - ESTILO VISUAL MEJORADO (MÁS CLARO)
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');

/* Fondo general */
.stApp { 
    background-color: #050505 !important; 
    background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; 
    color: #ffffff !important; 
    font-family: 'Rajdhani', sans-serif !important; 
}

/* Títulos */
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700 !important; text-align: center; letter-spacing: 3px; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #00f0ff !important; text-shadow: 0 0 10px rgba(0, 240, 255, 0.5) !important; text-align: center; }
h4 { font-family: 'Orbitron', sans-serif !important; color: #FF007F !important; }

/* Textos generales y etiquetas */
p, span, div, label, .stMarkdown { color: #e2e8f0 !important; }
[data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 14px !important; font-weight: bold !important; }

/* Inputs y TextAreas (Fondo más claro para leer mejor) */
.stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { 
    background-color: #1a1e24 !important; 
    border: 1px solid #00f0ff !important; 
    border-radius: 8px !important; 
    color: #ffffff !important; 
    font-size: 16px !important; 
}
.stTextArea textarea:focus, .stTextInput input:focus { border: 1px solid #FF007F !important; box-shadow: 0 0 10px rgba(255, 0, 127, 0.5) !important; }

/* Tablas de datos (Muy importante para claridad) */
[data-testid="stDataFrame"] { background-color: #1a1e24 !important; border-radius: 8px; padding: 5px; }

/* Sidebar */
[data-testid="stSidebar"] { background-color: #0a0a0a !important; border-right: 2px solid #FFD700 !important; }

/* Botones (Primario y Secundario) */
div.stButton > button[kind="secondary"] { background-color: #1a1e24 !important; color: #00f0ff !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; width: 100% !important; font-weight: bold !important; transition: all 0.3s; }
div.stButton > button[kind="secondary"]:hover { background-color: #00f0ff !important; color: #000 !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000000 !important; border: none !important; border-radius: 8px !important; font-weight: 900 !important; width: 100% !important; box-shadow: 0 0 15px rgba(0, 240, 255, 0.5) !important; transition: all 0.3s; }
div.stButton > button[kind="primary"]:hover { transform: scale(1.02); }

/* Cajas de alerta */
.stAlert { border-radius: 8px; background-color: rgba(0, 240, 255, 0.1) !important; border-left: 5px solid #00f0ff !important; color: #fff !important; }
</style>
""", unsafe_allow_html=True)

def inyectar_js_fixer():
    components.html("<script>(function(){var doc=window.parent.document;function f(){doc.querySelectorAll('[data-testid=\"stTextInput\"] *').forEach(function(el){if(el.tagName.toLowerCase()==='input'||el.tagName.toLowerCase()==='div'){el.style.setProperty('background-color','#1a1e24','important');el.style.setProperty('color','#ffffff','important');}});}f();setInterval(f,300);})();</script>", height=0)
inyectar_js_fixer()

# ==========================================
# 5. FUNCIONES LÓGICAS (IA, ARCHIVOS, AUDIO)
# ==========================================
def buscar_en_internet(query):
    if DDGS is None: return []
    try:
        res = DDGS().text(query, region='es-es', max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in res]
    except: return []

def extraer_youtube_id(url):
    if not url: return None
    m = re.search(r"(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([A-Za-z0-9_-]{11})", url)
    if m: return m.group(1)
    if len(url.strip()) == 11: return url.strip()
    return None

def transcribir_audio_youtube(url, api_key_groq):
    if not HAS_YTDLP or not HAS_OPENAI: return "[ERROR] Faltan yt-dlp u openai instalados."
    if not api_key_groq: return "[ERROR] No hay Groq API Key."
    cookies_path = None
    try:
        cookie_content = st.secrets.get("YT_COOKIES")
        if cookie_content:
            tmp_cookie_file = tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.txt')
            tmp_cookie_file.write(cookie_content); tmp_cookie_file.close(); cookies_path = tmp_cookie_file.name
    except Exception: pass
    tmp = tempfile.mkdtemp(); audio_path = None
    try:
        ydl_opts = {"format": "bestaudio[ext=m4a]/bestaudio/best", "outtmpl": os.path.join(tmp, "audio.%(ext)s"), "quiet": True, "noplaylist": True, "impersonate": "chrome", "extractor_args": {"youtube": ["player_client=android"]}}
        if cookies_path: ydl_opts["cookiefile"] = cookies_path
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            info = ydl.extract_info(url, download=True); audio_path = ydl.prepare_filename(info)
        if os.path.getsize(audio_path) / (1024 * 1024) > 24: return "[ERROR] Audio mayor a 25 MB."
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e: return f"[ERROR] {e}"
    finally:
        if audio_path and os.path.exists(audio_path):
            try: os.remove(audio_path)
            except: pass
        if cookies_path and os.path.exists(cookies_path):
            try: os.remove(cookies_path)
            except: pass
        try: os.rmdir(tmp)
        except: pass

def transcribir_archivo_local(archivo_bytes, nombre_archivo, api_key_groq):
    if not HAS_OPENAI: return "[ERROR] Falta openai."
    if not api_key_groq: return "[ERROR] No hay Groq API Key."
    ext = os.path.splitext(nombre_archivo)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(archivo_bytes); tmp.close(); audio_path = tmp.name
    try:
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e: return f"[ERROR] {str(e)}"
    finally:
        if os.path.exists(audio_path):
            try: os.remove(audio_path)
            except: pass

def leer_archivo_universal(upload, modelo_ia):
    if upload.size > MAX_FILE_MB * 1024 * 1024: return ""
    nombre = upload.name.lower()
    if nombre.endswith(".txt"): return upload.read().decode("utf-8", errors="ignore")
    if nombre.endswith(".pdf"): return "\n".join([p.extract_text() for p in pypdf.PdfReader(BytesIO(upload.getvalue())).pages if p.extract_text()])
    if nombre.endswith(".docx"): return "\n".join([p.text for p in Document(upload).paragraphs if p.text])
    return "Formato no reconocido."

@st.cache_resource
def get_gemini_client(api_key): return genai.Client(api_key=api_key)

def analizar_archivo_con_gemini(archivo_bytes, mime_type, prompt, api_key_gemini):
    if not HAS_GEMINI: return "[ERROR] Falta google-genai."
    size_mb = len(archivo_bytes) / (1024 * 1024); tmp_path = None
    try:
        client = get_gemini_client(api_key_gemini)
        if size_mb > 15:
            ext = mime_type.split('/')[-1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
            tmp.write(archivo_bytes); tmp.close(); tmp_path = tmp.name
            uploaded_file = client.files.upload(file=tmp_path)
            while uploaded_file.state.name == "PROCESSING": time.sleep(3); uploaded_file = client.files.get(name=uploaded_file.name)
            if uploaded_file.state.name == "FAILED": return "[ERROR] Gemini falló."
            res = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, uploaded_file]).text
            client.files.delete(name=uploaded_file.name); return res
        else:
            parte = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
            return client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, parte]).text
    except Exception as e: return f"[ERROR GEMINI] {str(e)}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None, api_key_openai="", api_key_anthropic=""):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if st.session_state.get("global_context"):
        prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}"
    if historial_chat:
        contexto_chat = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in historial_chat[-10:]])
        prompt_final = f"Historial:\n{contexto_chat}\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."

    try: key_gemini = st.session_state.api_key_gemini_manual or st.secrets.get("GEMINI_API_KEY", api_key_gemini)
    except: key_gemini = st.session_state.api_key_gemini_manual or api_key_gemini
    try: key_groq = st.session_state.api_key_groq_manual or st.secrets.get("GROQ_API_KEY", "")
    except: key_groq = st.session_state.api_key_groq_manual or ""
    try: key_silicon = st.session_state.api_key_silicon_manual or st.secrets.get("SILICONFLOW_API_KEY", "")
    except: key_silicon = st.session_state.api_key_silicon_manual or ""
    try: key_mistral = st.session_state.api_key_mistral_manual or st.secrets.get("MISTRAL_API_KEY", "")
    except: key_mistral = st.session_state.api_key_mistral_manual or ""
    try: key_openrouter = st.session_state.api_key_openrouter_manual or st.secrets.get("OPENROUTER_API_KEY", "")
    except: key_openrouter = st.session_state.api_key_openrouter_manual or ""

    res_exitosa = False; res = None
    if key_gemini.strip() and HAS_GEMINI and not res_exitosa:
        for intento in range(2):
            try:
                res = get_gemini_client(key_gemini).models.generate_content(model=GEMINI_MODEL, contents=prompt_final).text
                res_exitosa = True; break
            except Exception as e:
                if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e): break 
                elif "503" in str(e) and intento < 1: time.sleep(2); continue
                else: break
    if key_groq.strip() and HAS_OPENAI and not res_exitosa:
        try:
            client = openai.OpenAI(api_key=key_groq, base_url="https://api.groq.com/openai/v1")
            res = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            res_exitosa = True
        except Exception: pass
    if key_silicon.strip() and HAS_OPENAI and not res_exitosa:
        try:
            client = openai.OpenAI(api_key=key_silicon, base_url="https://api.siliconflow.cn/v1")
            res = client.chat.completions.create(model="Qwen/Qwen2.5-7B-Instruct", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            res_exitosa = True
        except Exception: pass
    if key_mistral.strip() and HAS_OPENAI and not res_exitosa:
        try:
            client = openai.OpenAI(api_key=key_mistral, base_url="https://api.mistral.ai/v1")
            res = client.chat.completions.create(model="mistral-small-latest", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            res_exitosa = True
        except Exception: pass
    if key_openrouter.strip() and HAS_OPENAI and not res_exitosa:
        try:
            client = openai.OpenAI(api_key=key_openrouter, base_url="https://openrouter.ai/api/v1")
            res = client.chat.completions.create(model="deepseek/deepseek-r1:free", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            res_exitosa = True
        except Exception: pass
    if res_exitosa and res: return res
    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=600)
        # CORRECCIÓN AQUÍ: Eliminado el .split("") que causaba el error
        return client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})["message"]["content"].strip()
    except Exception as e: return f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}"

def obtener_ruta_fuente():
    for ruta in ["C:/Windows/Fonts/arial.ttf", "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "/Library/Fonts/Arial.ttf"]:
        if os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo, contenido):
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente()
    try:
        url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Cover ' + titulo)}?width=512&height=512&nologo=true"
        req = urllib.request.Request(url_img, headers={'User-Agent': 'Mozilla/5.0'})
        with urllib.request.urlopen(req) as response: img_data = response.read()
        with open("temp_cover.jpg", "wb") as f: f.write(img_data)
        pdf.image("temp_cover.jpg", x=40, y=50, w=130); pdf.ln(150); os.remove("temp_cover.jpg")
    except: pass
    if font_path: pdf.add_font("Arial", "", font_path, uni=True); pdf.set_font("Arial", size=11)
    else: pdf.set_font("Helvetica", size=11); contenido = contenido.encode('latin-1', 'replace').decode('latin-1')
    pdf.set_font("Helvetica", "B", 24); pdf.multi_cell(0, 10, txt=titulo, align='C')
    pdf.set_font("Helvetica", "I", 12); pdf.cell(0, 10, txt="Generado por Xavilu", align='C', ln=True)
    pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)
    if font_path: pdf.set_font("Arial", size=11)
    else: pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 8, txt=contenido); return bytes(pdf.output())

def generar_word(texto):
    doc = Document(); doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); out.seek(0); return out.getvalue()

def generar_excel(texto_csv):
    wb = openpyxl.Workbook(); ws = wb.active
    for r_i, row in enumerate(csv.reader(texto_csv.replace("```csv", "").replace("```", "").strip().split('\n')), 1):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            if r_i == 1: cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

def generar_pptx(titulo, texto):
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[0]); slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1]); lineas = bloque.split("\n")
        if slide.shapes.title: slide.shapes.title.text = lineas[0][:50]
        contenido_texto = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
        tiene_ph = False
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1: shape.text = contenido_texto; tiene_ph = True
        if not tiene_ph:
            txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5)); txBox.text_frame.text = contenido_texto
    out = BytesIO(); prs.save(out); out.seek(0); return out.getvalue()

def botones_descarga(texto, nombre_por_defecto="documento", formatos_extra=False):
    if texto:
        st.markdown("---")
        key_u = f"{nombre_por_defecto}".replace(" ", "_")
        nombre_elegido = st.text_input("📝 Nombre del archivo:", value=nombre_por_defecto, key=f"name_{key_u}")
        nombre_limpio = re.sub(r'[^a-zA-Z0-9_\- ]', '', nombre_elegido).strip().replace(' ', '_') or nombre_por_defecto
        
        if formatos_extra:
            c1, c2, c3, c4 = st.columns(4)
            with c1: st.download_button("📥 TXT", texto.encode('utf-8'), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
            with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
            with c3: st.download_button("📝 Word", generar_word(texto), f"{nombre_limpio}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"wd_{key_u}", use_container_width=True)
            with c4: st.download_button("📊 Excel", generar_excel(texto), f"{nombre_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"ex_{key_u}", use_container_width=True)
        else:
            c1, c2, c3 = st.columns(3)
            with c1: st.download_button("📥 TXT", texto.encode('utf-8'), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
            with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
            with c3:
                if st.button("📲 Enviar a WhatsApp", key=f"wa_{key_u}", use_container_width=True):
                    num = st.session_state.get("whatsapp_number", "").strip()
                    if num:
                        url_wa = f"https://wa.me/{num}?text={urllib.parse.quote(texto[:1500])}"
                        components.html(f"<script>window.open('{url_wa}', '_blank');</script>", height=0)
                        st.success("✅ Abriendo WhatsApp...")
                    else: st.error("⚠️ Configura tu número.")

def run_async_safe(coro):
    def _run(): return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor() as pool: return pool.submit(_run).result()

def generar_audio_sincrono(texto, voz):
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(texto, voz).stream():
            if chunk["type"] == "audio": buf.write(chunk["data"])
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error TTS: {e}"); return None

def generar_podcast_3_voces(guion, v1, v2, v3):
    async def _gen():
        buf = BytesIO()
        texto_procesado = re.sub(r'\(risas\)|\[risas\]|\*risas\*', 'jajaja', guion, flags=re.IGNORECASE)
        texto_procesado = re.sub(r'\(pausa\)|\[pausa\]', '...', texto_procesado, flags=re.IGNORECASE)
        texto_procesado = re.sub(r'[*#_]', '', texto_procesado)
        lineas = texto_procesado.split('\n'); voz_actual = v1
        for linea in lineas:
            linea = linea.strip()
            if not linea: continue
            match = re.search(r'^Locutor\s*(\d)', linea, re.IGNORECASE)
            if match:
                num = int(match.group(1))
                if num == 1: voz_actual = v1
                elif num == 2: voz_actual = v2
                elif num == 3: voz_actual = v3
                texto_a_leer = re.sub(r'^Locutor\s*\d\s*:?', '', linea, flags=re.IGNORECASE).strip()
            else: texto_a_leer = linea
            if not texto_a_leer or texto_a_leer.lower() in ["elvira", "jorge", "tomas", "carlos"]: continue
            try:
                texto_final = texto_a_leer + ", "
                async for chunk in edge_tts.Communicate(texto_final, voz_actual).stream():
                    if chunk["type"] == "audio": buf.write(chunk["data"])
            except Exception: pass
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error crítico 3 voces: {e}"); return None

def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="...", api_key_openai="", api_key_anthropic=""):
    input_key = f"{key_prefix}_input"
    if input_key not in st.session_state: st.session_state[input_key] = ""
    def agregar_idea(idea_seleccionada):
        texto_actual = st.session_state[input_key].strip()
        if texto_actual: st.session_state[input_key] = f"{texto_actual}\n\n{idea_seleccionada}"
        else: st.session_state[input_key] = idea_seleccionada
        if f"{key_prefix}_sugs" in st.session_state: del st.session_state[f"{key_prefix}_sugs"]
    if st.button("💡 Generar Ideas con IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
        with st.spinner("Pensando ideas..."):
            prompt = f"Dame 3 ideas creativas para {contexto_ia}. Solo lista separada por saltos de línea."
            res = llamar_ia(modelo, prompt, api_key)
            st.session_state[f"{key_prefix}_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
    if st.session_state.get(f"{key_prefix}_sugs"):
        st.markdown("✨ **Haz clic para añadir:**")
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]: st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True, on_click=agregar_idea, args=(idea,))
    return st.text_area(label, key=input_key, placeholder=placeholder, height=100)

def boton_superpoder_borrador(texto_base, instruccion, key_suffix, modelo, api_key, state_key):
    if st.button(f"🚀 {instruccion}", key=f"super_{key_suffix}", type="secondary", use_container_width=True):
        if not texto_base: return
        with st.spinner("Aplicando superpoder..."):
            prompt = f"Actúa como experto. {instruccion}. Devuelve SOLO el texto mejorado:\n\n{texto_base}"
            res = llamar_ia(modelo, prompt, api_key)
            if "[ERROR" in res: st.error(res)
            else: st.session_state[state_key] = res; st.rerun()

def chat_input_pro(key, placeholder="Mensaje..."):
    t_k = f"{key}_txt"; m_k = f"{key}_msg"
    if t_k not in st.session_state: st.session_state[t_k] = ""
    def _send():
        msg = st.session_state.get(t_k, "").strip()
        if msg: st.session_state[m_k] = msg; st.session_state[t_k] = ""
    c1, c2 = st.columns([6, 1])
    with c1: st.text_input("msg", placeholder=placeholder, key=t_k, label_visibility="collapsed")
    with c2: st.button("➤", key=f"{key}_btn", on_click=_send, use_container_width=True, type="primary")
    return st.session_state.pop(m_k, None)

def custom_tabs(tab_names, key_name, default_idx=0):
    s_key = f"tab_{key_name}"
    if s_key not in st.session_state: st.session_state[s_key] = default_idx
    for i in range(0, len(tab_names), 5):
        chunk = tab_names[i:i+5]; cols = st.columns(5) 
        for j, name in enumerate(chunk):
            idx = i + j
            with cols[j]:
                if st.button(name, key=f"btn_{key_name}_{idx}", use_container_width=True, type="primary" if st.session_state[s_key] == idx else "secondary"):
                    st.session_state[s_key] = idx; st.rerun()
    return st.session_state[s_key]

def cajita_adjuntar_archivo(key_suffix, modelo, api_key):
    with st.expander("📎 Adjuntar archivo (Opcional)"):
        archivo = st.file_uploader("Sube PDF, Word, TXT o Imagen", type=["pdf", "txt", "docx", "png", "jpg", "jpeg"], key=f"adj_{key_suffix}")
        if archivo:
            if archivo.name.lower().endswith((".png", ".jpg", ".jpeg")):
                with st.spinner("Analizando imagen..."):
                    ctx = analizar_archivo_con_gemini(archivo.getvalue(), "image/jpeg", "Describe esta imagen o extrae el texto (OCR).", api_key)
                    st.success("✅ Imagen leída."); return f"\n[TEXTO DE IMAGEN]:\n{ctx}\n"
            else:
                texto = leer_archivo_universal(archivo, modelo)
                st.success(f"✅ Archivo '{archivo.name}' leído."); return f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"
    return ""

# ==========================================
# 6. INTERFAZ PRINCIPAL (SIDEBAR Y MACROS)
# ==========================================
if not st.session_state.sidebar_visible:
    st.markdown("""<style>[data-testid="collapsedControl"] {display: none !important;}</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⛶ MOSTRAR MENÚ Y KEYS", type="primary", use_container_width=True):
            st.session_state.sidebar_visible = True; st.rerun()

st.markdown("<h1>XAVILU NETWORK</h1><p style='text-align:center; color:#FFD700; font-weight:bold;'>● SISTEMA CONECTADO</p>", unsafe_allow_html=True)

if st.session_state.sidebar_visible:
    with st.sidebar:
        try: st.image("logo.png", use_container_width=True)
        except: st.markdown("### 🪐 PANEL DE CONTROL")
        if st.button("⛶ Ocultar Menú", use_container_width=True): 
            st.session_state.sidebar_visible = False; st.rerun()
        st.markdown("---")
        modelo_seleccionado = st.selectbox("Ollama Local", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])
        
        st.markdown("#### 🔑 API Keys (Fallback)")
        st.caption("Si Ollama falla, la app usará estas IAs en orden.")
        st.session_state.api_key_gemini_manual = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key_gemini_manual)
        st.session_state.api_key_groq_manual = st.text_input("Groq API Key", type="password", value=st.session_state.api_key_groq_manual)
        st.session_state.api_key_silicon_manual = st.text_input("SiliconFlow API Key", type="password", value=st.session_state.api_key_silicon_manual)
        st.session_state.api_key_mistral_manual = st.text_input("Mistral API Key", type="password", value=st.session_state.api_key_mistral_manual)
        st.session_state.api_key_openrouter_manual = st.text_input("OpenRouter API Key", type="password", value=st.session_state.api_key_openrouter_manual)
            
        st.session_state.whatsapp_number = st.text_input("📲 WhatsApp (ej. 34600000000):", value=st.session_state.whatsapp_number)
        
        archivo_global = st.file_uploader("📁 Cerebro Global", type=["pdf", "txt", "docx", "png", "jpg"], key="up_global")
        if archivo_global and st.session_state.get("global_file_name") != archivo_global.name:
            with st.spinner("Procesando..."):
                st.session_state.global_context = leer_archivo_universal(archivo_global, modelo_seleccionado)[:5000]
                st.session_state.global_file_name = archivo_global.name; st.success("✅ Memorizado.")
        if st.session_state.get("global_file_name"):
            st.info(f"Activo: **{st.session_state.global_file_name}**")
            if st.button("🗑️ Olvidar"): st.session_state.global_context = ""; st.session_state.global_file_name = ""; st.rerun()
else:
    modelo_seleccionado = "llama3.2:latest"

macro_names = ["🛠️ XAVIWORK", "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR", "♻️ GESTIÓN PLANTAS", "🎵 MÚSICA Y ARTE"]
macro_idx = custom_tabs(macro_names, "macro", 0)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# ==========================================
# 7. MÓDULOS (CADA MACRO CATEGORÍA)
# ==========================================

# --- MACRO 0: XAVIWORK ---
if macro_idx == 0:
    st.markdown("#### 🛠️ XaviWork - El Estudio Definitivo")
    with st.container():
        col1, col2 = st.columns([2, 1])
        with col1:
            contexto_xw = cajita_adjuntar_archivo("xw_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
            tema_xw = st.text_area("💡 ¿Sobre qué quieres trabajar?", height=100, key="xw_tema_input", placeholder="Ej: Estrategia de marketing...")
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💡 Generar 3 Ideas Base", use_container_width=True):
                with st.spinner("Pensando..."):
                    res = llamar_ia(modelo_seleccionado, f"Dame 3 ideas creativas sobre: '{tema_xw}'.", st.session_state.api_key_gemini_manual)
                    st.session_state["xw_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
            if st.session_state.get("xw_sugs"):
                for i, idea in enumerate(st.session_state["xw_sugs"]):
                    if st.button(f"➕ {idea}", key=f"xw_idea_{i}", use_container_width=True):
                        st.session_state["xw_tema_input"] += f"\n{idea}"; del st.session_state["xw_sugs"]; st.rerun()
    st.markdown("---")
    
    tabs_xw = st.tabs(["🕸️ Mapa", "📊 PPT", "📄 Informe", "🃏 Tarjetas", "❓ Test", "📋 Tabla", "🎙️ Podcast", "💬 Copiloto", "📅 Timeline", "📖 Glosario", "📊 DAFO", "🖼️ Portada", "🏢 Multioffice"])
    
    with tabs_xw[0]:
        if st.button("✨ Generar Mapa Mental", key="btn_xw_mapa", type="primary"):
            with st.spinner("Estructurando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGenera un mapa mental jerárquico en Markdown sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_mapa = res.replace("```markdown", "").replace("```", "").strip()
        if st.session_state.xw_mapa:
            markmap(st.session_state.xw_mapa, height=400)
            st.text_area("Editar Markdown:", value=st.session_state.xw_mapa, height=150, key="edit_xw_mapa")
            st.session_state.xw_mapa = st.session_state.edit_xw_mapa
            c1, c2 = st.columns(2)
            with c1: boton_superpoder_borrador(st.session_state.xw_mapa, "Hazlo más profesional", "xw_map_p1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "xw_mapa")
            with c2: st.download_button("📥 Descargar Markdown", st.session_state.xw_mapa.encode(), "mapa.md", use_container_width=True)

    with tabs_xw[1]:
        if st.button("✨ Generar Presentación", key="btn_xw_ppt", type="primary"):
            with st.spinner("Estructurando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nCrea una estructura de presentación para PowerPoint sobre: {st.session_state.xw_tema_input}. Separa cada diapositiva con doble salto de línea.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_ppt = res
        if st.session_state.xw_ppt:
            st.text_area("Editar Presentación:", value=st.session_state.xw_ppt, height=250, key="edit_xw_ppt")
            st.session_state.xw_ppt = st.session_state.edit_xw_ppt
            st.download_button("📥 Descargar PPTX", generar_pptx("Presentación XaviWork", st.session_state.xw_ppt), "presentacion.pptx", use_container_width=True)

    with tabs_xw[2]:
        if st.button("✨ Generar Informe", key="btn_xw_informe", type="primary"):
            with st.spinner("Redactando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nRedacta un informe profesional sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_informe = res
        if st.session_state.xw_informe:
            st.text_area("Editar Informe:", value=st.session_state.xw_informe, height=250, key="edit_xw_informe")
            st.session_state.xw_informe = st.session_state.edit_xw_informe
            botones_descarga(st.session_state.xw_informe, "Informe_XaviWork", formatos_extra=True)

    with tabs_xw[3]:
        if st.button("✨ Generar Tarjetas", key="btn_xw_tarjetas", type="primary"):
            with st.spinner("Creando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nCrea 5 tarjetas didácticas (pregunta y respuesta) sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_tarjetas = res
        if st.session_state.xw_tarjetas:
            st.text_area("Editar Tarjetas:", value=st.session_state.xw_tarjetas, height=250, key="edit_xw_tarjetas")
            st.session_state.xw_tarjetas = st.session_state.edit_xw_tarjetas
            st.download_button("📥 Descargar Tarjetas", st.session_state.xw_tarjetas.encode(), "tarjetas.txt", use_container_width=True)

    with tabs_xw[4]:
        if st.button("✨ Generar Cuestionario", key="btn_xw_cuest", type="primary"):
            with st.spinner("Creando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nCrea un cuestionario de 5 preguntas tipo test sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_cuestionario = res
        if st.session_state.xw_cuestionario:
            st.text_area("Editar Cuestionario:", value=st.session_state.xw_cuestionario, height=250, key="edit_xw_cuest")
            st.session_state.xw_cuestionario = st.session_state.edit_xw_cuest
            st.download_button("📥 Descargar Test", st.session_state.xw_cuestionario.encode(), "cuestionario.txt", use_container_width=True)

    with tabs_xw[5]:
        if st.button("✨ Generar Tabla de Datos", key="btn_xw_tabla", type="primary"):
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGenera una tabla de datos en formato CSV sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_tabla = res.replace("```csv", "").replace("```", "").strip()
        if st.session_state.xw_tabla:
            st.text_area("Editar Tabla (CSV):", value=st.session_state.xw_tabla, height=250, key="edit_xw_tabla")
            st.session_state.xw_tabla = st.session_state.edit_xw_tabla
            st.download_button("📥 Descargar Excel", generar_excel(st.session_state.xw_tabla), "tabla.xlsx", use_container_width=True)

    with tabs_xw[6]:
        if st.button("✨ Generar Guion Podcast", key="btn_xw_pod", type="primary"):
            with st.spinner("Escribiendo..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGuion podcast sobre: '{st.session_state.xw_tema_input}'. 3 locutores. Formato exacto: 'Locutor 1: [texto]'.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_podcast = res
        if st.session_state.xw_podcast:
            st.text_area("Guion (Editable):", value=st.session_state.xw_podcast, height=250, key="edit_xw_pod")
            st.session_state.xw_podcast = st.session_state.edit_xw_pod
            if st.button("🔊 Generar Audio", key="btn_xw_pod_aud"):
                with st.spinner("Generando audio..."):
                    aud = generar_podcast_3_voces(st.session_state.xw_podcast, "es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural")
                    if aud: st.session_state.xw_podcast_audio = aud; st.success("✅ Audio listo.")
            if st.session_state.xw_podcast_audio:
                st.audio(st.session_state.xw_podcast_audio, format="audio/mp3")
                st.download_button("💾 Descargar MP3", st.session_state.xw_podcast_audio, "podcast_xaviwork.mp3")

    with tabs_xw[7]:
        chat_container = st.container(height=400)
        with chat_container:
            for msg in st.session_state.xw_copilot_history:
                with st.chat_message(msg["role"]): st.write(msg["content"])
        if prompt_chat := st.chat_input("Pregunta sobre el tema..."):
            st.session_state.xw_copilot_history.append({"role": "user", "content": prompt_chat})
            with chat_container:
                with st.chat_message("user"): st.write(prompt_chat)
                with st.chat_message("assistant"):
                    with st.spinner("Pensando..."):
                        res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nTema: {st.session_state.xw_tema_input}\n\nPregunta: {prompt_chat}", st.session_state.api_key_gemini_manual)
                        st.write(res); st.session_state.xw_copilot_history.append({"role": "assistant", "content": res})

    with tabs_xw[8]:
        if st.button("✨ Generar Línea de Tiempo", key="btn_xw_time", type="primary"):
            with st.spinner("Extrayendo fechas..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGenera una línea de tiempo sobre: {st.session_state.xw_tema_input}. Formato: Año - Evento.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_timeline = res
        if st.session_state.xw_timeline:
            st.text_area("Editar Línea de Tiempo:", value=st.session_state.xw_timeline, height=250, key="edit_xw_time")
            st.session_state.xw_timeline = st.session_state.edit_xw_time
            botones_descarga(st.session_state.xw_timeline, "Timeline", formatos_extra=True)

    with tabs_xw[9]:
        if st.button("✨ Generar Glosario", key="btn_xw_glos", type="primary"):
            with st.spinner("Extrayendo..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGenera un glosario de 5 términos sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_glosario = res
        if st.session_state.xw_glosario:
            st.text_area("Editar Glosario:", value=st.session_state.xw_glosario, height=250, key="edit_xw_glos")
            st.session_state.xw_glosario = st.session_state.edit_xw_glos
            botones_descarga(st.session_state.xw_glosario, "Glosario", formatos_extra=True)

    with tabs_xw[10]:
        if st.button("✨ Generar Análisis DAFO", key="btn_xw_dafo", type="primary"):
            with st.spinner("Analizando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nRealiza un Análisis DAFO sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                st.session_state.xw_dafo = res
        if st.session_state.xw_dafo:
            st.text_area("Editar DAFO:", value=st.session_state.xw_dafo, height=250, key="edit_xw_dafo")
            st.session_state.xw_dafo = st.session_state.edit_xw_dafo
            botones_descarga(st.session_state.xw_dafo, "DAFO", formatos_extra=True)

    with tabs_xw[11]:
        prompt_img = st.text_input("Describe la imagen:", value=st.session_state.xw_tema_input, key="xw_img_prompt")
        if st.button("🖼️ Generar Portada", key="btn_xw_img", type="primary") and prompt_img:
            with st.spinner("Generando imagen..."):
                try:
                    url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img + ', cyberpunk')}?width=1024&height=1024&nologo=true"
                    req = urllib.request.Request(url_img, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req) as response: img_data = response.read()
                    st.session_state.xw_cover_img = img_data
                except Exception as e: st.error(f"Error: {e}")
        if st.session_state.xw_cover_img:
            st.image(st.session_state.xw_cover_img, use_container_width=True)
            st.download_button("📥 Descargar Imagen", st.session_state.xw_cover_img, "portada.png", "image/png", use_container_width=True)

    with tabs_xw[12]:
        st.markdown("#### 🏢 Multioffice - Generación Simultánea")
        st.markdown("Genera el mismo contenido adaptado a diferentes formatos corporativos de una sola vez.")
        
        col_f1, col_f2, col_f3, col_f4 = st.columns(4)
        with col_f1: opt_word = st.checkbox("Word (.docx)", value=True)
        with col_f2: opt_excel = st.checkbox("Excel (.xlsx)", value=False)
        with col_f3: opt_ppt = st.checkbox("PowerPoint (.pptx)", value=False)
        with col_f4: opt_pdf = st.checkbox("PDF (.pdf)", value=True)
        
        if st.button("✨ Generar Documentos Múltiples", type="primary", use_container_width=True):
            if not st.session_state.xw_tema_input.strip():
                st.warning("Por favor, introduce un tema principal en la caja superior.")
            else:
                with st.spinner("Procesando y estructurando para cada formato..."):
                    if opt_word or opt_pdf:
                        res_texto = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nRedacta un documento formal y completo sobre: {st.session_state.xw_tema_input}.", st.session_state.api_key_gemini_manual)
                        st.session_state.xw_multi_word = res_texto
                        st.session_state.xw_multi_pdf = res_texto
                    if opt_excel:
                        res_csv = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nGenera una tabla de datos en formato CSV sobre: {st.session_state.xw_tema_input}. Devuelve SOLO el código CSV.", st.session_state.api_key_gemini_manual)
                        st.session_state.xw_multi_excel = res_csv.replace("```csv", "").replace("```", "").strip()
                    if opt_ppt:
                        res_ppt = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nCrea una estructura de diapositivas para PowerPoint sobre: {st.session_state.xw_tema_input}. Usa doble salto de línea entre diapositivas.", st.session_state.api_key_gemini_manual)
                        st.session_state.xw_multi_ppt = res_ppt
                st.success("✅ Contenidos generados. Revisa y descarga abajo.")

        st.markdown("---")
        st.markdown("### 📥 Resultados y Descargas")
        
        if opt_word and st.session_state.get("xw_multi_word"):
            with st.expander("📝 Documento Word (Editable)", expanded=False):
                st.text_area("Texto Word:", value=st.session_state.xw_multi_word, height=200, key="edit_multi_word")
                st.session_state.xw_multi_word = st.session_state.edit_multi_word
                st.download_button("📥 Descargar Word", generar_word(st.session_state.xw_multi_word), "multioffice_doc.docx", use_container_width=True)

        if opt_excel and st.session_state.get("xw_multi_excel"):
            with st.expander("📊 Tabla Excel (Editable)", expanded=False):
                st.text_area("CSV Excel:", value=st.session_state.xw_multi_excel, height=200, key="edit_multi_excel")
                st.session_state.xw_multi_excel = st.session_state.edit_multi_excel
                st.download_button("📥 Descargar Excel", generar_excel(st.session_state.xw_multi_excel), "multioffice_tabla.xlsx", use_container_width=True)

        if opt_ppt and st.session_state.get("xw_multi_ppt"):
            with st.expander("📉 Presentación PPT (Editable)", expanded=False):
                st.text_area("Estructura PPT:", value=st.session_state.xw_multi_ppt, height=200, key="edit_multi_ppt")
                st.session_state.xw_multi_ppt = st.session_state.edit_multi_ppt
                st.download_button("📥 Descargar PowerPoint", generar_pptx("Presentación Multioffice", st.session_state.xw_multi_ppt), "multioffice_ppt.pptx", use_container_width=True)

        if opt_pdf and st.session_state.get("xw_multi_pdf"):
            with st.expander("📄 Documento PDF (Editable)", expanded=False):
                st.text_area("Texto PDF:", value=st.session_state.xw_multi_pdf, height=200, key="edit_multi_pdf")
                st.session_state.xw_multi_pdf = st.session_state.edit_multi_pdf
                st.download_button("📥 Descargar PDF", generar_pdf_pro("Multioffice PDF", st.session_state.xw_multi_pdf), "multioffice_pdf.pdf", use_container_width=True)

# --- MACRO 1: CHAT Y ASISTENTES ---
elif macro_idx == 1:
    sub_names = ["🎙️ PODCAST", "💬 CHAT", "🎙️ VOZ", "💬 CHAT CON PDF", "🕵️‍♂️ AGENTES", "⚖️ DEBATE", "🍳 COCINA", "📖 CUENTOS", "💻 CÓDIGO", "🌍 TRADUCTOR"]
    sub_idx = custom_tabs(sub_names, "chat_sub", 0)
    
    if sub_idx == 0:
        st.markdown("#### 🎙️ Podcast")
        contexto = cajita_adjuntar_archivo("podcast_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        tema = st.text_area("Temática:", key="pod_tema_input", height=100)
        if st.button("✨ Crear Guion", type="primary", key="btn_pod") and st.session_state["pod_tema_input"]:
            with st.spinner("Escribiendo..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\nGuion podcast Tema: '{st.session_state['pod_tema_input']}'. 3 locutores. Formato exacto: 'Locutor 1: [texto]'.", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.session_state.podcast_script = res
        if st.session_state.get("podcast_script"):
            st.text_area("Guion:", value=st.session_state.podcast_script, height=300, key="pod_edit")
            st.session_state.podcast_script = st.session_state.pod_edit
            if st.button("🔊 Audio 3 Voces", key="btn_pod_aud"):
                with st.spinner("Generando audio..."):
                    aud = generar_podcast_3_voces(st.session_state.podcast_script, "es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural")
                    if aud: st.audio(aud, format="audio/mp3"); st.download_button("💾 MP3", aud, "podcast.mp3")
            botones_descarga(st.session_state.podcast_script, "Podcast")

    elif sub_idx == 1:
        st.markdown("#### 💬 Chat")
        for m in st.session_state.chat_history: st.chat_message(m["role"]).write(m["content"])
        if st.button("🗑️ Limpiar Chat"): st.session_state.chat_history = []; st.rerun()
        ctx = cajita_adjuntar_archivo("chat_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        p = chat_input_pro("main")
        if p:
            st.session_state.chat_history.append({"role": "user", "content": p}); st.chat_message("user").write(p)
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"{ctx}\n\n{p}", st.session_state.api_key_gemini_manual, st.session_state.chat_history)
                if "[ERROR" not in res: st.session_state.chat_history.append({"role": "assistant", "content": res}); st.chat_message("assistant").write(res)
            st.rerun()

    elif sub_idx == 2:
        st.markdown("#### Voz")
        txt = st.text_area("Texto:"); voz = st.selectbox("Voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural"])
        if st.button("🔊 Audio") and txt:
            aud = generar_audio_sincrono(txt, voz)
            if aud: st.audio(aud, format="audio/mp3")

    elif sub_idx == 3:
        st.markdown("#### Chat con PDFs")
        archs = st.file_uploader("Sube docs:", type=["pdf", "txt", "docx"], accept_multiple_files=True)
        if archs:
            texto = "".join([leer_archivo_universal(a, modelo_seleccionado) for a in archs])[:MAX_TEXTO_IA]
            for m in st.session_state.pdf_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("pdf")
            if p:
                st.session_state.pdf_chat_history.append({"role": "user", "content": p})
                with st.spinner("Buscando..."):
                    res = llamar_ia(modelo_seleccionado, f"Contexto:\n{texto}\n\nPregunta: {p}", st.session_state.api_key_gemini_manual, st.session_state.pdf_chat_history)
                    if "[ERROR" not in res: st.session_state.pdf_chat_history.append({"role": "assistant", "content": res}); st.rerun()

    elif sub_idx == 4:
        st.markdown("#### Agentes")
        nom = st.text_input("Nombre:"); rol = st.text_area("Personalidad:")
        if st.button("Crear"): st.session_state.agente_activo = {"nombre": nom, "rol": rol}; st.session_state.agente_chat_history = []
        if st.session_state.get("agente_activo"):
            ag = st.session_state.agente_activo
            for m in st.session_state.agente_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("agent")
            if p:
                st.session_state.agente_chat_history.append({"role": "user", "content": p})
                with st.spinner("Pensando..."):
                    res = llamar_ia(modelo_seleccionado, f"ERES: {ag['rol']}\nUsuario: {p}", st.session_state.api_key_gemini_manual, st.session_state.agente_chat_history)
                    if "[ERROR" not in res: st.session_state.agente_chat_history.append({"role": "assistant", "content": res}); st.rerun()

    elif sub_idx == 5:
        st.markdown("#### Debate")
        t = campo_con_ideas("Tema:", "deb", "temas", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Iniciar") and t:
            res = llamar_ia(modelo_seleccionado, f"Debate {t}: 3 a favor, 3 en contra.", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.markdown(res); botones_descarga(res, "Debate")

    elif sub_idx == 6:
        st.markdown("#### Cocina")
        ing = campo_con_ideas("Ingredientes:", "chef1", "recetas", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Receta") and ing: 
            res = llamar_ia(modelo_seleccionado, f"Receta con {ing}", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.info(res)

    elif sub_idx == 7:
        st.markdown("#### Cuentos")
        p = campo_con_ideas("Protagonista:", "c1", "protagonistas", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Crear") and p:
            res = llamar_ia(modelo_seleccionado, f"Cuento sobre {p}.", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.text_area("Cuento:", res, height=300); botones_descarga(res, "Cuento")

    elif sub_idx == 8:
        st.markdown("#### Código")
        req = campo_con_ideas("Código:", "code", "scripts", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Procesar") and req:
            res = llamar_ia(modelo_seleccionado, f"Programador senior:\n{req}", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.code(res, language="python")

    elif sub_idx == 9:
        st.markdown("#### Traductor")
        txt = campo_con_ideas("Texto:", "trad", "frases", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Traducir") and txt:
            res = llamar_ia(modelo_seleccionado, f"Traduce al inglés:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.info(res); botones_descarga(res, "Traduccion")

# --- MACRO 2: DOCUMENTOS Y TEXTO ---
elif macro_idx == 2:
    sub_names = ["📁 ANÁLISIS", "⚖️ COMPARAR", "📊 OFFICE", "📄 PDF", "📇 QR & CV", "📋 PORTAPAPELES"]
    sub_idx = custom_tabs(sub_names, "doc_sub", 0)
    
    if sub_idx == 0:
        st.markdown("#### Análisis de Documentos")
        arch = st.file_uploader("Sube PDF, Word o TXT", type=["pdf", "txt", "docx"])
        opc = st.selectbox("Análisis:", ["Resumen Ejecutivo", "Extraer Fechas", "Buscar Errores", "Crear Test"])
        if arch and st.button("🔎 Ejecutar", type="primary"):
            with st.spinner("Analizando..."):
                texto = leer_archivo_universal(arch, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"{opc}:\n{texto}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.text_area("Resultado:", res, height=400); botones_descarga(res, "Analisis")

    elif sub_idx == 1:
        st.markdown("#### Comparador")
        a = st.file_uploader("Original:", type=["pdf", "txt"]); b = st.file_uploader("Modificado:", type=["pdf", "txt"])
        if a and b and st.button("Comparar"):
            with st.spinner("Buscando..."):
                ta = leer_archivo_universal(a, modelo_seleccionado); tb = leer_archivo_universal(b, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"Diferencias:\nA:\n{ta[:10000]}\nB:\n{tb[:10000]}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.text_area("Diferencias:", res, height=400); botones_descarga(res, "Comparativa")

    elif sub_idx == 2:
        st.markdown("#### Creación de Documentos")
        contexto = cajita_adjuntar_archivo("office_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        prompt_office = campo_con_ideas("¿Qué documento crear?", "office", "documentos", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        formato = st.radio("Formato:", ["Word", "Excel", "PowerPoint"], horizontal=True)
        if st.button("✨ Generar Borrador", type="primary") and prompt_office:
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\n{prompt_office}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.session_state.office_borrador = res
        if st.session_state.get("office_borrador"):
            st.text_area("Editar Borrador:", value=st.session_state.office_borrador, height=250, key="off_b")
            st.session_state.office_borrador = st.session_state.off_b
            if formato == "Word": st.download_button("📥 Word", generar_word(st.session_state.office_borrador), "doc.docx")
            elif formato == "Excel": st.download_button("📥 Excel", generar_excel(st.session_state.office_borrador), "doc.xlsx")
            else: st.download_button("📥 PPTX", generar_pptx("Presentación", st.session_state.office_borrador), "doc.pptx")

    elif sub_idx == 3:
        st.markdown("#### PDF Pro")
        titulo = st.text_input("Título del PDF:")
        contenido = campo_con_ideas("Contenido:", "pdf_c", "contenido PDF", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("✨ Generar PDF", type="primary") and titulo and contenido:
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contenido}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.session_state.pdf_borrador = res
        if st.session_state.get("pdf_borrador"):
            st.text_area("Editar:", value=st.session_state.pdf_borrador, height=250, key="pdf_b")
            st.session_state.pdf_borrador = st.session_state.pdf_b
            st.download_button("📥 PDF", generar_pdf_pro(titulo, st.session_state.pdf_borrador), f"{titulo}.pdf")

    elif sub_idx == 4:
        c1, c2 = st.columns(2)
        with c1:
            txt = st.text_input("Texto/URL QR:")
            if st.button("Generar QR") and txt:
                qr = qrcode.QRCode(); qr.add_data(txt); qr.make(fit=True)
                buf = BytesIO(); qr.make_image().save(buf, format="PNG"); st.image(buf.getvalue())
        with c2:
            img = st.file_uploader("Leer QR:", type=["png", "jpg"])
            if img and HAS_CV2:
                d, _, _ = cv2.QRCodeDetector().detectAndDecode(np.array(Image.open(img).convert('RGB')))
                if d: st.success(d)

    elif sub_idx == 5:
        st.markdown("#### Portapapeles")
        txt = st.text_area("Pega texto:")
        inst = campo_con_ideas("¿Qué hago?", "clip", "instrucciones", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Procesar") and txt and inst:
            res = llamar_ia(modelo_seleccionado, f"{inst}:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in res: st.info(res)

# --- MACRO 3: WEB Y ACTUALIDAD ---
elif macro_idx == 3:
    sub_names = ["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV", "📊 WEB A EXCEL", "🎥 YOUTUBE", "📰 LECTOR"]
    sub_idx = custom_tabs(sub_names, "web_sub", 0)
    if sub_idx == 0:
        q = campo_con_ideas("Buscar:", "web", "búsquedas", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("🔍 Buscar") and q:
            for r in buscar_en_internet(q): st.markdown(f"**[{r['titulo']}]({r['url']})**"); st.caption(r['resumen'])
    elif sub_idx == 1:
        d = st.selectbox("Deporte:", ["Fútbol", "Baloncesto", "Tenis"])
        if st.button("Noticias"):
            for n in buscar_noticias_deportivas(d, "Todas"): st.markdown(f"**{n['titulo']}** - {n['resumen']}")
    elif sub_idx == 2:
        c = campo_con_ideas("Canal:", "tv", "TV", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        if st.button("Programación") and c:
            res = buscar_en_internet(c)
            r_ia = llamar_ia(modelo_seleccionado, f"Programación {c}: {res}", st.session_state.api_key_gemini_manual)
            if "[ERROR" not in r_ia: st.info(r_ia)
    elif sub_idx == 3:
        url = st.text_input("URL Tabla:")
        if st.button("Extraer") and url:
            try:
                df = pd.read_html(url)[0]; st.dataframe(df)
                out = BytesIO(); df.to_excel(out); st.download_button("📥 Excel", out.getvalue(), "tabla.xlsx")
            except: st.error("Fallo al leer.")
    elif sub_idx == 4:
        st.markdown("#### 🎥 YouTube")
        url = st.text_input("URL del vídeo:")
        if st.button("Ejecutar", type="primary") and url:
            vid = extraer_youtube_id(url)
            key_groq = st.session_state.api_key_groq_manual
            try:
                tr = YouTubeTranscriptApi.get_transcript(vid, languages=['es', 'en'])
                texto_original = " ".join([t['text'] for t in tr])
            except:
                with st.spinner("Transcribiendo con Whisper..."): texto_original = transcribir_audio_youtube(url, key_groq)
            if not texto_original.startswith("[ERROR]"):
                res = llamar_ia(modelo_seleccionado, f"Resume:\n{texto_original[:8000]}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.text_area("Resumen:", res, height=300); botones_descarga(res, "YT")
            else: st.error(texto_original)
    elif sub_idx == 5:
        url = st.text_input("URL Web:")
        if st.button("Leer") and url:
            try:
                txt = " ".join(re.sub('<[^<]+?>', ' ', requests.get(url).text).split())[:5000]
                res = llamar_ia(modelo_seleccionado, f"Resume: {txt}", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.text_area("Resumen:", res, height=200)
            except: st.error("Error.")

# --- MACRO 4: MULTIMEDIA ---
elif macro_idx == 4:
    st.markdown("### 🎬 MULTIMEDIA")
    st.info("Módulo de imagen, vídeo y generación visual listo para expandir.")
    
# --- MACRO 5: ESTUDIO Y OFICIOS ---
elif macro_idx == 5:
    st.markdown("### 🎓 ESTUDIO Y OFICIOS")
    st.info("Módulo de deberes, PRL, entrevistas y más listo para expandir.")

# --- MACRO 6: VIAJES Y HOGAR ---
elif macro_idx == 6:
    st.markdown("### 🌍 VIAJES Y HOGAR")
    st.info("Módulo de viajes, recetas de hogar y organización listo para expandir.")

# --- MACRO 7: GESTIÓN PLANTAS ---
elif macro_idx == 7:
    st.markdown("### ♻️ GESTIÓN PLANTAS")
    sub_plant = ["📦 STOCK", "🛒 COMPRAS", "💰 VENTAS", "⚖️ BÁSCULA", "📋 LOTES", "🔧 MANTENIMIENTO"]
    sub_plant_idx = custom_tabs(sub_plant, "plant_sub", 0)
    
    df_stock = get_db_data('stock')
    opciones_stock = {f"{row['id']} - {row['nombre']} ({row['unidad']})": row['id'] for _, row in df_stock.iterrows()} if not df_stock.empty else {}

    if sub_plant_idx == 0:
        st.markdown("#### 📦 Inventario Actual")
        with st.expander("➕ Añadir Nuevo Artículo"):
            with st.form("form_nuevo_stock"):
                c1, c2, c3 = st.columns(3)
                with c1: nuevo_nombre = st.text_input("Nombre"); nueva_categoria = st.selectbox("Categoría", ["Repuestos", "Consumibles", "EPIs", "Material Procesado"])
                with c2: nueva_cantidad = st.number_input("Cantidad", min_value=0.0); nueva_unidad = st.selectbox("Unidad", ["Unidades", "kg", "Litros", "Metros"])
                with c3: nuevo_minimo = st.number_input("Stock Mínimo", min_value=0.0); nueva_ubicacion = st.text_input("Ubicación")
                if st.form_submit_button("💾 Guardar", type="primary"):
                    if nuevo_nombre:
                        fecha_ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        insert_db("INSERT INTO stock (nombre, categoria, cantidad, unidad, stock_minimo, ubicacion, fecha_actualizacion) VALUES (?, ?, ?, ?, ?, ?, ?)", (nuevo_nombre, nueva_categoria, nueva_cantidad, nueva_unidad, nuevo_minimo, nueva_ubicacion, fecha_ahora))
                        st.success("✅ Añadido."); st.rerun()
        if not df_stock.empty:
            df_bajo_stock = df_stock[df_stock['cantidad'] <= df_stock['stock_minimo']]
            if not df_bajo_stock.empty: st.error(f"⚠️ **{len(df_bajo_stock)} artículos bajo mínimo.**")
            st.dataframe(df_stock[['id', 'nombre', 'categoria', 'cantidad', 'unidad', 'stock_minimo', 'ubicacion']], use_container_width=True)
        else: st.info("Catálogo vacío.")

    elif sub_plant_idx == 1:
        st.markdown("#### 🛒 Compras (Entradas)")
        if not opciones_stock: st.warning("Añade artículos en STOCK primero.")
        else:
            with st.expander("🤖 Autocompletar con IA (Sube Factura)"):
                factura_compra = st.file_uploader("Sube Factura (PDF/IMG):", type=["pdf", "png", "jpg", "jpeg"], key="fac_compra")
                if factura_compra and st.button("✨ Extraer datos", key="extraer_compra"):
                    with st.spinner("IA leyendo..."):
                        if factura_compra.name.lower().endswith(".pdf"): texto_factura = leer_archivo_universal(factura_compra, modelo_seleccionado)
                        else: texto_factura = analizar_archivo_con_gemini(factura_compra.getvalue(), "image/jpeg", "Extrae el texto literal.", st.session_state.api_key_gemini_manual)
                        prompt = f"Extrae datos de esta factura. JSON: {{'proveedor': 'Nom', 'material': 'Nom', 'cantidad': 10.5, 'precio_unitario': 15.0}}. Texto: {texto_factura[:5000]}"
                        res_ia = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                        try:
                            json_limpio = res_ia.replace("```json", "").replace("```", "").strip()
                            datos = json.loads(json_limpio)
                            st.session_state.proveedor_temp = datos.get("proveedor", "")
                            st.session_state.material_temp = datos.get("material", "")
                            st.session_state.cant_temp = float(datos.get("cantidad", 0.0))
                            st.session_state.precio_temp = float(datos.get("precio_unitario", 0.0))
                            st.success("✅ Datos extraídos. Revisa abajo.")
                        except Exception as e: st.error(f"Rellena a mano. (Error: {e})")

            with st.form("form_compra"):
                c1, c2, c3 = st.columns(3)
                with c1: proveedor = st.text_input("Proveedor", value=st.session_state.get("proveedor_temp", ""))
                with c2: 
                    idx_material = 0
                    mat_temp = st.session_state.get("material_temp", "")
                    if mat_temp and mat_temp in [v.split(" - ")[1].split(" (")[0] for v in opciones_stock.values()]:
                        idx_material = [v.split(" - ")[1].split(" (")[0] for v in opciones_stock.values()].index(mat_temp)
                    item_compra = st.selectbox("Material:", list(opciones_stock.keys()), index=idx_material)
                with c3: cant_compra = st.number_input("Cantidad:", min_value=0.1, step=1.0, value=st.session_state.get("cant_temp", 1.0))
                precio_u = st.number_input("Precio Unitario (€):", min_value=0.0, step=0.01, value=st.session_state.get("precio_temp", 0.0))
                archivo_ref = st.text_input("Referencia archivo:", value=factura_compra.name if 'factura_compra' in locals() and factura_compra else "")
                
                if st.form_submit_button("💾 Registrar Compra", type="primary"):
                    id_item = opciones_stock[item_compra]
                    cant_actual = float(df_stock[df_stock['id'] == id_item]['cantidad'].values[0])
                    update_stock(id_item, cant_actual + cant_compra)
                    total = cant_compra * precio_u
                    fecha_ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                    insert_db("INSERT INTO compras (fecha, proveedor, material, cantidad, precio_unitario, total, archivo) VALUES (?, ?, ?, ?, ?, ?, ?)", (fecha_ahora, proveedor, item_compra, cant_compra, precio_u, total, archivo_ref))
                    st.session_state.proveedor_temp = ""; st.session_state.material_temp = ""; st.session_state.cant_temp = 0.0; st.session_state.precio_temp = 0.0
                    st.success("✅ Compra registrada. Stock actualizado."); st.rerun()
        st.markdown("##### Historial")
        st.dataframe(get_db_data('compras'), use_container_width=True)

    elif sub_plant_idx == 2:
        st.markdown("#### 💰 Ventas (Salidas)")
        if not opciones_stock: st.warning("Añade artículos en STOCK primero.")
        else:
            with st.expander("🤖 Autocompletar con IA (Sube Albarán)"):
                factura_venta = st.file_uploader("Sube Albarán (PDF/IMG):", type=["pdf", "png", "jpg", "jpeg"], key="fac_venta")
                if factura_venta and st.button("✨ Extraer datos", key="extraer_venta"):
                    with st.spinner("IA leyendo..."):
                        if factura_venta.name.lower().endswith(".pdf"): texto_factura = leer_archivo_universal(factura_venta, modelo_seleccionado)
                        else: texto_factura = analizar_archivo_con_gemini(factura_venta.getvalue(), "image/jpeg", "Extrae el texto literal.", st.session_state.api_key_gemini_manual)
                        prompt = f"Extrae datos de este albarán. JSON: {{'cliente': 'Nom', 'material': 'Nom', 'cantidad': 10.5, 'precio_unitario': 15.0}}. Texto: {texto_factura[:5000]}"
                        res_ia = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                        try:
                            json_limpio = res_ia.replace("```json", "").replace("```", "").strip()
                            datos = json.loads(json_limpio)
                            st.session_state.cliente_temp = datos.get("cliente", "")
                            st.session_state.material_v_temp = datos.get("material", "")
                            st.session_state.cant_v_temp = float(datos.get("cantidad", 0.0))
                            st.session_state.precio_v_temp = float(datos.get("precio_unitario", 0.0))
                            st.success("✅ Datos extraídos.")
                        except Exception as e: st.error(f"Rellena a mano. (Error: {e})")

            with st.form("form_venta"):
                c1, c2, c3 = st.columns(3)
                with c1: cliente = st.text_input("Cliente", value=st.session_state.get("cliente_temp", ""))
                with c2: 
                    idx_material_v = 0
                    mat_v_temp = st.session_state.get("material_v_temp", "")
                    if mat_v_temp and mat_v_temp in [v.split(" - ")[1].split(" (")[0] for v in opciones_stock.values()]:
                        idx_material_v = [v.split(" - ")[1].split(" (")[0] for v in opciones_stock.values()].index(mat_v_temp)
                    item_venta = st.selectbox("Material:", list(opciones_stock.keys()), index=idx_material_v)
                with c3: cant_venta = st.number_input("Cantidad:", min_value=0.1, step=1.0, value=st.session_state.get("cant_v_temp", 1.0))
                precio_v = st.number_input("Precio Unitario (€):", min_value=0.0, step=0.01, value=st.session_state.get("precio_v_temp", 0.0))
                archivo_ref_v = st.text_input("Referencia archivo:", value=factura_venta.name if 'factura_venta' in locals() and factura_venta else "")
                
                if st.form_submit_button("💾 Registrar Venta", type="primary"):
                    id_item = opciones_stock[item_venta]
                    cant_actual = float(df_stock[df_stock['id'] == id_item]['cantidad'].values[0])
                    if cant_venta > cant_actual: st.error(f"❌ Stock insuficiente. Quedan {cant_actual}.")
                    else:
                        update_stock(id_item, cant_actual - cant_venta)
                        total = cant_venta * precio_v
                        fecha_ahora = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                        insert_db("INSERT INTO ventas (fecha, cliente, material, cantidad, precio_unitario, total, archivo) VALUES (?, ?, ?, ?, ?, ?, ?)", (fecha_ahora, cliente, item_venta, cant_venta, precio_v, total, archivo_ref_v))
                        st.session_state.cliente_temp = ""; st.session_state.material_v_temp = ""; st.session_state.cant_v_temp = 0.0; st.session_state.precio_v_temp = 0.0
                        st.success("✅ Venta registrada. Stock actualizado."); st.rerun()
        st.markdown("##### Historial")
        st.dataframe(get_db_data('ventas'), use_container_width=True)

    elif sub_plant_idx == 3:
        st.markdown("#### ⚖️ Báscula"); st.dataframe(get_db_data('bascula'), use_container_width=True)
    elif sub_plant_idx == 4:
        st.markdown("#### 📋 Lotes"); st.dataframe(get_db_data('lotes'), use_container_width=True)
    elif sub_plant_idx == 5:
        st.markdown("#### 🔧 Mantenimiento"); st.dataframe(get_db_data('mantenimiento'), use_container_width=True)

# --- MACRO 8: MÚSICA Y ARTE ---
elif macro_idx == 8:
    st.markdown("### 🎵 MÚSICA Y ARTE")
    st.info("Módulo de letras, canciones y portadas listo para expandir.")
