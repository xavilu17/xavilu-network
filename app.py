# ==========================================
# XAVILU NETWORK PRO — VERSIÓN UNIFICADA
# Streamlit + GitHub ready
# ==========================================
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
import socket
import ipaddress

# --- Importaciones opcionales ---
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
# 1. CONFIGURACIÓN
# ==========================================
st.set_page_config(
    page_title="Xavilu Network Pro",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="expanded"
)

XAVILU_VERSION = "2.1 - Unificada"
MAX_FILE_MB = 100
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-2.0-flash"  # modelo real y estable

if os.name == 'nt':
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# ==========================================
# 2. UTILIDADES DE SEGURIDAD
# ==========================================
def url_es_segura(url):
    """Valida que la URL no apunte a recursos internos (anti-SSRF)."""
    try:
        if not url or not isinstance(url, str): return False
        if not url.startswith(("http://", "https://")): return False
        hostname = urllib.parse.urlparse(url).hostname
        if not hostname: return False
        ip = socket.gethostbyname(hostname)
        ip_obj = ipaddress.ip_address(ip)
        if (ip_obj.is_private or ip_obj.is_loopback or ip_obj.is_link_local
                or ip_obj.is_reserved or ip_obj.is_multicast):
            return False
        return True
    except Exception:
        return False

# ==========================================
# 3. MEMORIA GENERAL (Historial)
# ==========================================
def init_db():
    try:
        with sqlite3.connect('xavilu_memoria.db', timeout=10) as conn:
            conn.execute('''CREATE TABLE IF NOT EXISTS historial (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fecha TEXT, tipo TEXT, contenido TEXT)''')
    except Exception as e:
        st.warning(f"No se pudo inicializar la memoria: {e}")

def guardar_en_db(tipo, contenido):
    try:
        with sqlite3.connect('xavilu_memoria.db', timeout=10) as conn:
            conn.execute(
                "INSERT INTO historial (fecha, tipo, contenido) VALUES (?, ?, ?)",
                (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo, contenido)
            )
    except Exception:
        pass

def obtener_de_db():
    try:
        with sqlite3.connect('xavilu_memoria.db', timeout=10) as conn:
            return conn.execute(
                "SELECT fecha, tipo, contenido FROM historial ORDER BY id DESC LIMIT 5"
            ).fetchall()
    except Exception:
        return []

init_db()

# ==========================================
# 4. ESTADO DE SESIÓN
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
    "doc_opcion": "Resumen Ejecutivo",
    "global_context": "", "global_file_name": "",
    "agente_activo": None, "agente_chat_history": [],
    "debate_resultado": "", "img_prompt_en": "",
    "podcast_script": "", "podcast_ideas": [],
    "sidebar_visible": True,
    "whatsapp_number": "",
    "office_borrador": "", "pdf_borrador": "",
    "pod_tema_input": "", "pod_sugs": [],
    "api_key_gemini_manual": "", "api_key_groq_manual": "", "api_key_silicon_manual": "",
    "api_key_mistral_manual": "", "api_key_openrouter_manual": "",
    "xw_mapa": "", "xw_ppt": "", "xw_informe": "", "xw_tarjetas": "", "xw_cuestionario": "",
    "xw_tabla": "", "xw_podcast": "", "xw_podcast_audio": None, "xw_timeline": "",
    "xw_glosario": "", "xw_dafo": "", "xw_cover_img": None, "xw_copilot_history": [],
    "xw_sugs": []
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

# ==========================================
# 5. CSS — ESTÉTICA NEGRO ESPACIAL
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
.stApp { background-color: #000000 !important; background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
header[data-testid="stHeader"] { background-color: #000000 !important; }
[data-testid="stAppDeployButton"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700, 0 0 35px #FF8C00 !important; text-align: center; letter-spacing: 3px; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #ffffff !important; text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff !important; text-align: center; letter-spacing: 3px; }
h4 { font-family: 'Orbitron', sans-serif !important; color: #00f0ff !important; text-shadow: 0 0 8px #00f0ff !important; letter-spacing: 1px; }
[data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; text-transform: uppercase !important; background: rgba(0, 240, 255, 0.1) !important; padding: 5px 10px !important; border-left: 4px solid #FF007F !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 8px !important; }
[data-testid="stSidebar"] { background-color: rgba(5, 5, 10, 0.98) !important; border-right: 2px solid #00f0ff !important; box-shadow: 5px 0 15px rgba(0, 240, 255, 0.2) !important; }
[data-testid="stRadio"] > div[role="radiogroup"] label div p, [data-testid="stRadio"] > div[role="radiogroup"] label div span { color: #ffffff !important; font-size: 15px !important; font-weight: 600 !important; }
[data-testid="stFileUploader"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; padding: 10px !important; }
[data-testid="stFileUploaderDropzone"] { background-color: #0a0a0a !important; color: #ffffff !important; }
[data-testid="stFileUploaderDropzone"] p { color: #00f0ff !important; font-weight: bold !important; }
[data-testid="stExpander"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; }
[data-testid="stExpander"] summary { color: #00f0ff !important; font-weight: bold !important; }
div.stButton > button[kind="secondary"] { background-color: #0a0a0a !important; color: #ffffff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; padding: 10px 15px !important; margin-bottom: 5px !important; transition: all 0.3s ease !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.2) !important; width: 100% !important; }
div.stButton > button[kind="secondary"]:hover { background-color: rgba(0, 240, 255, 0.15) !important; color: #00f0ff !important; box-shadow: 0 0 20px rgba(0, 240, 255, 0.7) !important; transform: translateY(-2px) !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000000 !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: 900 !important; border: 2px solid #ffffff !important; border-radius: 25px !important; padding: 10px 15px !important; margin-bottom: 5px !important; transition: all 0.3s ease !important; box-shadow: 0 0 18px rgba(0, 240, 255, 0.9) !important; width: 100% !important; }
.stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; color: #ffffff !important; font-size: 16px !important; box-shadow: inset 0 0 10px rgba(0, 240, 255, 0.1) !important; }
.stTextArea textarea:focus, .stTextInput input:focus { border: 1px solid #FF007F !important; box-shadow: 0 0 15px rgba(255, 0, 127, 0.5) !important; }
[data-testid="stTextInput"] input { caret-color: #00f0ff !important; border: 1px solid #00f0ff !important; border-radius: 10px !important; padding: 10px 14px !important; font-size: 15px !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.2) !important; }
[data-testid="stTextInput"] input:focus { border: 1px solid #FF007F !important; box-shadow: 0 0 18px rgba(255, 0, 127, 0.6) !important; }
div[data-testid="stDownloadButton"] > button { background: #0a0a0a !important; color: #00f0ff !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; font-family: 'Orbitron', sans-serif !important; font-weight: bold !important; padding: 10px 20px !important; transition: all 0.3s !important; }
div[data-testid="stDownloadButton"] > button:hover { background: linear-gradient(135deg, #FF007F 0%, #00f0ff 100%) !important; color: #000 !important; }
[data-testid="stChatMessage"] { background: rgba(10, 10, 15, 0.8) !important; border: 1px solid #FF007F !important; border-radius: 15px !important; padding: 10px; box-shadow: 0 0 15px rgba(255, 0, 127, 0.3) !important; }
.stAlert { border-radius: 10px; font-weight: bold; background-color: rgba(0, 240, 255, 0.1) !important; border-left: 5px solid #00f0ff !important; }
</style>
""", unsafe_allow_html=True)

def inyectar_js_fixer():
    components.html("<script>(function(){var doc=window.parent.document;function f(){doc.querySelectorAll('[data-testid=\"stTextInput\"] *').forEach(function(el){if(el.tagName.toLowerCase()==='input'||el.tagName.toLowerCase()==='div'){el.style.setProperty('background-color','#0a0a0a','important');el.style.setProperty('color','#ffffff','important');}});}f();setInterval(f,300);})();</script>", height=0)
inyectar_js_fixer()

# ==========================================
# 6. UTILIDADES IA / ARCHIVOS
# ==========================================
def buscar_en_internet(query):
    if DDGS is None: return []
    try:
        res = DDGS().text(query, region='es-es', max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in res]
    except Exception: return []

def buscar_noticias_deportivas(deporte, liga):
    if DDGS is None: return []
    try:
        res = DDGS().news(f"últimas noticias {deporte} {liga} hoy", max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("url", ""), "resumen": r.get("body", ""), "fecha": r.get("date", "")} for r in res]
    except Exception: return []

def extraer_youtube_id(url):
    if not url: return None
    m = re.search(r"(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([A-Za-z0-9_-]{11})", url)
    if m: return m.group(1)
    if len(url.strip()) == 11: return url.strip()
    return None

def transcribir_audio_youtube(url, api_key_groq):
    if not HAS_YTDLP or not HAS_OPENAI:
        return "[ERROR] Faltan yt-dlp u openai instalados."
    if not api_key_groq:
        return "[ERROR] No hay Groq API Key configurada."
    cookies_path = None
    try:
        cookie_content = st.secrets.get("YT_COOKIES")
        if cookie_content:
            tmp_cookie_file = tempfile.NamedTemporaryFile(delete=False, mode='w', suffix='.txt')
            tmp_cookie_file.write(cookie_content); tmp_cookie_file.close()
            cookies_path = tmp_cookie_file.name
    except Exception: pass
    tmp = tempfile.mkdtemp(); audio_path = None
    try:
        ydl_opts = {
            "format": "bestaudio[filesize<20M]/bestaudio/best",
            "outtmpl": os.path.join(tmp, "audio.%(ext)s"),
            "quiet": True, "noplaylist": True,
            "match_filter": yt_dlp.utils.match_filter_func("duration < 1800"),
            "extractor_args": {"youtube": ["player_client=android"]},
        }
        if cookies_path: ydl_opts["cookiefile"] = cookies_path
        try:
            ydl_opts["impersonate"] = "chrome"
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=True)
                audio_path = ydl.prepare_filename(info)
        except Exception:
            ydl_opts.pop("impersonate", None)
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=True)
                audio_path = ydl.prepare_filename(info)
        if not audio_path or not os.path.exists(audio_path):
            return "[ERROR] No se pudo descargar el audio."
        size_mb = os.path.getsize(audio_path) / (1024 * 1024)
        if size_mb > 24: return f"[ERROR] Audio de {size_mb:.1f} MB, supera 25 MB de Groq."
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e:
        return f"[ERROR] {e}"
    finally:
        if audio_path and os.path.exists(audio_path):
            try: os.remove(audio_path)
            except Exception: pass
        if cookies_path and os.path.exists(cookies_path):
            try: os.remove(cookies_path)
            except Exception: pass
        try: os.rmdir(tmp)
        except Exception: pass

def transcribir_archivo_local(archivo_bytes, nombre_archivo, api_key_groq):
    if not HAS_OPENAI: return "[ERROR] Falta openai instalado."
    if not api_key_groq: return "[ERROR] No hay Groq API Key."
    ext = os.path.splitext(nombre_archivo)[1]
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=ext)
    tmp.write(archivo_bytes); tmp.close(); audio_path = tmp.name
    try:
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e:
        return f"[ERROR] {str(e)}"
    finally:
        if os.path.exists(audio_path):
            try: os.remove(audio_path)
            except Exception: pass

def leer_archivo_universal(upload, modelo_ia):
    try:
        if upload.size > MAX_FILE_MB * 1024 * 1024:
            return f"[ERROR] Archivo supera {MAX_FILE_MB} MB."
        nombre = upload.name.lower()
        if nombre.endswith(".txt"): return upload.read().decode("utf-8", errors="ignore")
        if nombre.endswith(".pdf"): return "\n".join([p.extract_text() for p in pypdf.PdfReader(BytesIO(upload.getvalue())).pages if p.extract_text()])
        if nombre.endswith(".docx"): return "\n".join([p.text for p in Document(upload).paragraphs if p.text])
        return "Formato no reconocido."
    except Exception as e:
        return f"[ERROR] {e}"

def cajita_adjuntar_archivo(key_suffix, modelo, api_key):
    with st.expander("📎 Adjuntar archivo para esta acción (Opcional)"):
        archivo = st.file_uploader("Sube PDF, Word, TXT o Imagen", type=["pdf", "txt", "docx", "png", "jpg", "jpeg"], key=f"adj_{key_suffix}")
        if archivo:
            nombre = archivo.name.lower()
            if nombre.endswith((".png", ".jpg", ".jpeg")):
                with st.spinner("Analizando imagen..."):
                    ctx = analizar_archivo_con_gemini(archivo.getvalue(), "image/jpeg", "Describe esta imagen o extrae el texto (OCR).", api_key)
                    st.success("✅ Imagen leída."); return f"\n[TEXTO DE IMAGEN]:\n{ctx}\n"
            else:
                texto = leer_archivo_universal(archivo, modelo)
                if texto.startswith("[ERROR]"): st.error(texto); return ""
                st.success(f"✅ Archivo '{archivo.name}' leído.")
                return f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"
    return ""

@st.cache_resource(show_spinner=False)
def get_gemini_client(api_key): return genai.Client(api_key=api_key)

def analizar_archivo_con_gemini(archivo_bytes, mime_type, prompt, api_key_gemini):
    if not HAS_GEMINI: return "[ERROR] Falta google-genai."
    size_mb = len(archivo_bytes) / (1024 * 1024)
    tmp_path = None
    try:
        client = get_gemini_client(api_key_gemini)
        if size_mb > 15:
            ext = mime_type.split('/')[-1]
            tmp = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}")
            tmp.write(archivo_bytes); tmp.close(); tmp_path = tmp.name
            uploaded_file = client.files.upload(file=tmp_path)
            while uploaded_file.state.name == "PROCESSING":
                time.sleep(3); uploaded_file = client.files.get(name=uploaded_file.name)
            if uploaded_file.state.name == "FAILED": return "[ERROR] Gemini no pudo procesar el archivo."
            res = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, uploaded_file]).text
            client.files.delete(name=uploaded_file.name); return res
        else:
            parte = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
            return client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, parte]).text
    except Exception as e:
        return f"[ERROR GEMINI] {str(e)}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except Exception: pass

def _leer_keys():
    """Recupera todas las claves desde session_state > secrets > vacío."""
    def _get(manual_key, secret_key):
        manual = st.session_state.get(manual_key, "").strip()
        if manual: return manual
        try: return st.secrets.get(secret_key, "")
        except Exception: return ""
    return {
        "gemini": _get("api_key_gemini_manual", "GEMINI_API_KEY"),
        "groq": _get("api_key_groq_manual", "GROQ_API_KEY"),
        "silicon": _get("api_key_silicon_manual", "SILICONFLOW_API_KEY"),
        "mistral": _get("api_key_mistral_manual", "MISTRAL_API_KEY"),
        "openrouter": _get("api_key_openrouter_manual", "OPENROUTER_API_KEY"),
    }

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None, api_key_openai="", api_key_anthropic=""):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if st.session_state.get("global_context"):
        prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}\n\n(Responde teniendo en cuenta el archivo)."
    if historial_chat:
        contexto_chat = "\n".join([f"{msg['role'].upper()}: {msg['content']}" for msg in historial_chat[-10:]])
        prompt_final = f"Historial:\n{contexto_chat}\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."

    keys = _leer_keys()
    if not keys["gemini"] and api_key_gemini: keys["gemini"] = api_key_gemini

    # 1. Gemini
    if keys["gemini"] and HAS_GEMINI:
        for intento in range(2):
            try:
                res = get_gemini_client(keys["gemini"]).models.generate_content(model=GEMINI_MODEL, contents=prompt_final).text
                return res
            except Exception as e:
                s = str(e)
                if "429" in s or "RESOURCE_EXHAUSTED" in s:
                    st.toast("🔋 Cuota Gemini agotada → Groq"); break
                if "503" in s and intento < 1: time.sleep(2); continue
                st.toast(f"⚠️ Gemini: {s[:60]} → Groq"); break
    # 2. Groq
    if keys["groq"] and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=keys["groq"], base_url="https://api.groq.com/openai/v1")
            return client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
        except Exception as e: st.toast(f"⚠️ Groq: {str(e)[:60]} → SiliconFlow")
    # 3. SiliconFlow
    if keys["silicon"] and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=keys["silicon"], base_url="https://api.siliconflow.cn/v1")
            return client.chat.completions.create(model="Qwen/Qwen2.5-7B-Instruct", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
        except Exception as e: st.toast(f"⚠️ SiliconFlow: {str(e)[:60]} → Mistral")
    # 4. Mistral
    if keys["mistral"] and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=keys["mistral"], base_url="https://api.mistral.ai/v1")
            return client.chat.completions.create(model="mistral-small-latest", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
        except Exception as e: st.toast(f"⚠️ Mistral: {str(e)[:60]} → OpenRouter")
    # 5. OpenRouter
    if keys["openrouter"] and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=keys["openrouter"], base_url="https://openrouter.ai/api/v1")
            return client.chat.completions.create(model="deepseek/deepseek-r1:free", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
        except Exception as e: st.toast(f"⚠️ OpenRouter: {str(e)[:60]} → Ollama local")
    # 6. Ollama (último fallback)
    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=600)
        content = client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})["message"]["content"]
        return content.split("```")[-1].strip() if "```" in content else content.strip()
    except Exception as e:
        return f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}"

# ==========================================
# 7. GENERADORES DE DOCUMENTOS
# ==========================================
def obtener_ruta_fuente():
    rutas = [
        "assets/DejaVuSans.ttf",
        "assets/arial.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
    ]
    for ruta in rutas:
        if os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo, contenido):
    pdf = FPDF(); pdf.add_page()
    try:
        if url_es_segura(f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Book cover for ' + titulo + ', cyberpunk')}?width=512&height=512&nologo=true"):
            url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Book cover for ' + titulo + ', cyberpunk')}?width=512&height=512&nologo=true"
            req = urllib.request.Request(url_img, headers={'User-Agent': 'Mozilla/5.0'})
            with urllib.request.urlopen(req, timeout=15) as r:
                with open("temp_cover.jpg", "wb") as f: f.write(r.read())
            pdf.image("temp_cover.jpg", x=40, y=50, w=130); pdf.ln(150); os.remove("temp_cover.jpg")
    except Exception: pass
    font_path = obtener_ruta_fuente()
    if font_path:
        pdf.add_font("Arial", "", font_path, uni=True)
        pdf.set_font("Arial", "B", 24)
    else:
        pdf.set_font("Helvetica", "B", 24)
    pdf.multi_cell(0, 10, txt=titulo, align='C')
    pdf.set_font(pdf.font_family, "I", 12)
    pdf.cell(0, 10, txt="Generado por Xavilu", align='C', ln=True)
    pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)
    if font_path:
        pdf.set_font("Arial", size=11)
        pdf.multi_cell(0, 8, txt=contenido)
    else:
        pdf.set_font("Helvetica", size=11)
        pdf.multi_cell(0, 8, txt=contenido.encode('latin-1', 'replace').decode('latin-1'))
    return bytes(pdf.output())

def generar_word(texto):
    doc = Document(); doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); out.seek(0); return out.getvalue()

def generar_excel(texto_csv):
    wb = openpyxl.Workbook(); ws = wb.active
    limpio = texto_csv.replace("```csv", "").replace("```", "").strip()
    for r_i, row in enumerate(csv.reader(limpio.split('\n')), 1):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            if r_i == 1: cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); out.seek(0); return out.getvalue()

def generar_pptx(titulo, texto):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        lineas = bloque.split("\n")
        if slide.shapes.title: slide.shapes.title.text = lineas[0][:50]
        contenido_texto = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
        tiene_ph = False
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1:
                shape.text = contenido_texto; tiene_ph = True
        if not tiene_ph:
            txBox = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5))
            txBox.text_frame.text = contenido_texto
    out = BytesIO(); prs.save(out); out.seek(0); return out.getvalue()

def botones_descarga(texto, nombre_por_defecto="documento", formatos_extra=False):
    if not texto: return
    st.markdown("---")
    key_u = nombre_por_defecto.replace(" ", "_")
    nombre_elegido = st.text_input("📝 Nombre del archivo (sin extensión):", value=nombre_por_defecto, key=f"name_{key_u}")
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
                else:
                    st.error("⚠️ Configura tu número en la barra lateral.")

# ==========================================
# 8. AUDIO (TTS y Podcasts)
# ==========================================
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
    except Exception as e:
        st.error(f"Error TTS: {e}"); return None

def generar_podcast_3_voces(guion, v1, v2, v3):
    async def _gen():
        buf = BytesIO()
        texto = re.sub(r'\(risas\)|\[risas\]|\*risas\*', 'jajaja', guion, flags=re.IGNORECASE)
        texto = re.sub(r'\(pausa\)|\[pausa\]', '...', texto, flags=re.IGNORECASE)
        texto = re.sub(r'[*#_]', '', texto)
        voz_actual = v1
        for linea in texto.split('\n'):
            linea = linea.strip()
            if not linea: continue
            match = re.search(r'^Locutor\s*(\d)', linea, re.IGNORECASE)
            if match:
                num = int(match.group(1))
                voz_actual = v1 if num == 1 else v2 if num == 2 else v3 if num == 3 else voz_actual
                texto_a_leer = re.sub(r'^Locutor\s*\d\s*:?', '', linea, flags=re.IGNORECASE).strip()
            else:
                texto_a_leer = linea
            if not texto_a_leer: continue
            try:
                async for chunk in edge_tts.Communicate(texto_a_leer + ", ", voz_actual).stream():
                    if chunk["type"] == "audio": buf.write(chunk["data"])
            except Exception: pass
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e:
        st.error(f"Error crítico en 3 voces: {e}"); return None

# ==========================================
# 9. HELPERS DE UI
# ==========================================
def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="...", api_key_openai="", api_key_anthropic=""):
    input_key = f"{key_prefix}_input"
    if input_key not in st.session_state: st.session_state[input_key] = ""
    def agregar_idea(idea_seleccionada):
        actual = st.session_state[input_key].strip()
        st.session_state[input_key] = f"{actual}\n\n{idea_seleccionada}" if actual else idea_seleccionada
        if f"{key_prefix}_sugs" in st.session_state: del st.session_state[f"{key_prefix}_sugs"]
    if st.button("💡 Generar Ideas con IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
        with st.spinner("Pensando ideas..."):
            prompt = f"Dame 3 ideas creativas para {contexto_ia}. Solo lista separada por saltos de línea, sin numeración."
            res = llamar_ia(modelo, prompt, api_key, api_key_openai=api_key_openai, api_key_anthropic=api_key_anthropic)
            if not res.startswith("[ERROR"):
                st.session_state[f"{key_prefix}_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
    if st.session_state.get(f"{key_prefix}_sugs"):
        st.markdown("✨ **Haz clic para añadir:**")
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]:
                st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True, on_click=agregar_idea, args=(idea,))
    return st.text_area(label, key=input_key, placeholder=placeholder, height=100)

def boton_superpoder_borrador(texto_base, instruccion, key_suffix, modelo, api_key, state_key, api_key_openai="", api_key_anthropic=""):
    if st.button(f"🚀 {instruccion}", key=f"super_{key_suffix}", type="secondary", use_container_width=True):
        if not texto_base: return
        with st.spinner("Aplicando superpoder..."):
            prompt = f"Actúa como experto. {instruccion}. Devuelve SOLO el texto mejorado, sin introducciones:\n\n{texto_base}"
            res = llamar_ia(modelo, prompt, api_key, api_key_openai=api_key_openai, api_key_anthropic=api_key_anthropic)
            if res.startswith("[ERROR"): st.error(res)
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
                if st.button(name, key=f"btn_{key_name}_{idx}", use_container_width=True,
                             type="primary" if st.session_state[s_key] == idx else "secondary"):
                    st.session_state[s_key] = idx; st.rerun()
    return st.session_state[s_key]

# ==========================================
# 10. INTERFAZ PRINCIPAL — SIDEBAR
# ==========================================
if not st.session_state.sidebar_visible:
    st.markdown("""<style>[data-testid="collapsedControl"] {display: none !important;}</style>""", unsafe_allow_html=True)
    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        if st.button("⛶ MOSTRAR MENÚ Y KEYS", type="primary", use_container_width=True):
            st.session_state.sidebar_visible = True; st.rerun()

st.markdown(f"<h1>XAVILU NETWORK</h1><p style='text-align:center; color:#FFD700; font-weight:bold; letter-spacing:2px;'>● SISTEMA CONECTADO · v{XAVILU_VERSION}</p>", unsafe_allow_html=True)

if st.session_state.sidebar_visible:
    with st.sidebar:
        try: st.image("logo.png", use_container_width=True)
        except Exception: st.markdown("### 🪐 PANEL DE CONTROL")
        if st.button("⛶ Ocultar Menú (Pantalla Completa)", use_container_width=True):
            st.session_state.sidebar_visible = False; st.rerun()
        st.markdown("---")
        modelo_seleccionado = st.selectbox("Ollama Local", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])
        st.markdown("#### 🔑 Estado de Conexiones")
        try:
            if st.secrets.get("GEMINI_API_KEY"): st.success("✅ Gemini Conectado (vía secrets)")
        except Exception: pass
        with st.expander("⚙️ Configurar Keys Manualmente"):
            st.caption("Si no usas secrets.toml, pega aquí tus claves para esta sesión:")
            st.session_state.api_key_gemini_manual = st.text_input("Gemini API Key", type="password", value=st.session_state.api_key_gemini_manual)
            st.session_state.api_key_groq_manual = st.text_input("Groq API Key", type="password", value=st.session_state.api_key_groq_manual)
            st.session_state.api_key_silicon_manual = st.text_input("SiliconFlow API Key", type="password", value=st.session_state.api_key_silicon_manual)
            st.session_state.api_key_mistral_manual = st.text_input("Mistral API Key", type="password", value=st.session_state.api_key_mistral_manual)
            st.session_state.api_key_openrouter_manual = st.text_input("OpenRouter API Key", type="password", value=st.session_state.api_key_openrouter_manual)
        st.markdown("---")
        st.session_state.whatsapp_number = st.text_input("📲 WhatsApp (ej. 34600000000):", value=st.session_state.whatsapp_number)
        st.markdown("---")
        archivo_global = st.file_uploader("📁 Cerebro Global", type=["pdf", "txt", "docx", "png", "jpg"], key="up_global")
        if archivo_global and st.session_state.get("global_file_name") != archivo_global.name:
            with st.spinner("Procesando..."):
                st.session_state.global_context = leer_archivo_universal(archivo_global, modelo_seleccionado)[:5000]
                st.session_state.global_file_name = archivo_global.name
                st.success("✅ Memorizado.")
        if st.session_state.get("global_file_name"):
            st.info(f"Activo: **{st.session_state.global_file_name}**")
            if st.button("🗑️ Olvidar"):
                st.session_state.global_context = ""; st.session_state.global_file_name = ""; st.rerun()
        st.markdown("---")
        for h in obtener_de_db():
            with st.expander(f"{h[1]} - {h[0][:10]}"): st.caption(h[2][:200] + "...")
else:
    modelo_seleccionado = "llama3.2:latest"

# ==========================================
# 11. NAVEGACIÓN DE MACRO-CATEGORÍAS
# ==========================================
macro_names = ["🛠️ XAVIWORK", "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR", "🎵 MÚSICA Y ARTE"]
macro_idx = custom_tabs(macro_names, "macro", 0)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

# ==========================================
# CATEGORÍA 0: XAVIWORK
# ==========================================
if macro_idx == 0:
    st.markdown("#### 🛠️ XaviWork - El Estudio Definitivo")

    with st.container():
        col1, col2 = st.columns([2, 1])
        with col1:
            contexto_xw = cajita_adjuntar_archivo("xw_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
            tema_xw = st.text_area("💡 ¿Sobre qué quieres trabajar?", height=100, key="xw_tema_input", placeholder="Ej: Estrategia de marketing para CreatorFlow...")
        with col2:
            st.markdown("<br>", unsafe_allow_html=True)
            if st.button("💡 Generar 3 Ideas Base", use_container_width=True):
                with st.spinner("Pensando..."):
                    res = llamar_ia(modelo_seleccionado, f"Dame 3 ideas creativas y concisas sobre: '{tema_xw}'. Sepáralas por saltos de línea.", st.session_state.api_key_gemini_manual)
                    if not res.startswith("[ERROR"):
                        st.session_state["xw_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
            def _add_idea(idea):
                actual = st.session_state["xw_tema_input"].strip()
                st.session_state["xw_tema_input"] = f"{actual}\n{idea}" if actual else idea
                st.session_state["xw_sugs"] = []
            if st.session_state.get("xw_sugs"):
                for i, idea in enumerate(st.session_state["xw_sugs"]):
                    st.button(f"➕ {idea}", key=f"xw_idea_{i}", use_container_width=True, on_click=_add_idea, args=(idea,))

    st.markdown("---")

    tabs_xw = st.tabs(["🕸️ Mapa", "📊 PPT", "📄 Informe", "🃏 Tarjetas", "❓ Test", "📋 Tabla", "🎙️ Podcast", "💬 Copiloto", "📅 Timeline", "📖 Glosario", "📊 DAFO", "🖼️ Portada"])

    def _panel_borrador(nombre, prompt_gen, state_key, edit_key, superpowers, download_label=None):
        """Renderiza botón-generar + editor + superpoderes + descarga."""
        if st.button(f"✨ Generar {nombre}", key=f"btn_{state_key}", type="primary"):
            if not st.session_state.xw_tema_input.strip():
                st.warning("Define el tema arriba primero.")
            else:
                with st.spinner(f"Generando {nombre}..."):
                    full = f"{contexto_xw}\n\n{prompt_gen}\n\nTEMA: {st.session_state.xw_tema_input}"
                    res = llamar_ia(modelo_seleccionado, full, st.session_state.api_key_gemini_manual)
                    if res.startswith("[ERROR"): st.error(res)
                    else: st.session_state[state_key] = res.replace("```markdown","").replace("```csv","").replace("```","").strip() if state_key in ("xw_mapa","xw_tabla") else res

        if st.session_state.get(state_key):
            if edit_key not in st.session_state:
                st.session_state[edit_key] = st.session_state[state_key]
            st.text_area("Editar:", key=edit_key, height=250)
            st.session_state[state_key] = st.session_state[edit_key]

            n = len(superpowers) + (1 if download_label else 0)
            if n == 0: return
            cols = st.columns(n)
            for i, (label, prompt) in enumerate(superpowers):
                with cols[i]:
                    boton_superpoder_borrador(st.session_state[state_key], prompt,
                                              f"{state_key}_p{i}", modelo_seleccionado,
                                              st.session_state.api_key_gemini_manual, state_key)
            if download_label:
                with cols[-1]:
                    if state_key == "xw_tabla":
                        st.download_button(download_label, generar_excel(st.session_state[state_key]),
                                           "tabla.xlsx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key == "xw_ppt":
                        st.download_button(download_label, generar_pptx("Presentación XaviWork", st.session_state[state_key]),
                                           "presentacion.pptx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key in ("xw_informe", "xw_timeline", "xw_glosario", "xw_dafo"):
                        botones_descarga(st.session_state[state_key], state_key, formatos_extra=True)
                    else:
                        st.download_button(download_label, st.session_state[state_key].encode(),
                                           f"{state_key}.txt", key=f"dl_{state_key}", use_container_width=True)

    with tabs_xw[0]:
        _panel_borrador("Mapa Mental", "Genera un mapa mental jerárquico en Markdown. Usa #, ## y -. Devuelve SOLO el Markdown.", "xw_mapa", "edit_xw_mapa",
                        [("Hazlo más profesional", "Hazlo más profesional"), ("Añade más detalles", "Añade más detalles")],
                        "📥 Descargar Markdown")
        if st.session_state.get("xw_mapa"):
            with st.expander("🕸️ Ver mapa interactivo"): markmap(st.session_state.xw_mapa, height=400)

    with tabs_xw[1]:
        _panel_borrador("Presentación (PPTX)", "Crea una estructura de presentación. Separa cada diapositiva con doble salto de línea. Primer renglón = título, guiones = puntos.", "xw_ppt", "edit_xw_ppt",
                        [("Hazlo más visual y directo", "Hazlo más visual y directo")], "📥 Descargar PPTX")

    with tabs_xw[2]:
        _panel_borrador("Informe Detallado", "Redacta un informe profesional y detallado. Incluye introducción, desarrollo y conclusiones.", "xw_informe", "edit_xw_informe",
                        [("Hazlo más técnico", "Hazlo más técnico"), ("Resumir a la mitad", "Resumir a la mitad"), ("Traducir al Inglés", "Traducir al Inglés")],
                        "📥 Descargar")

    with tabs_xw[3]:
        _panel_borrador("Tarjetas (Flashcards)", "Crea 5 tarjetas didácticas (pregunta y respuesta). Formato: Pregunta -> Respuesta.", "xw_tarjetas", "edit_xw_tarjetas",
                        [("Hazlas más difíciles", "Hazlas más difíciles")], "📥 Descargar")

    with tabs_xw[4]:
        _panel_borrador("Cuestionario", "Crea un cuestionario de 5 preguntas tipo test (3 opciones + respuesta al final).", "xw_cuestionario", "edit_xw_cuest",
                        [("Añade 5 preguntas más", "Añade 5 preguntas más")], "📥 Descargar")

    with tabs_xw[5]:
        _panel_borrador("Tabla de Datos", "Genera una tabla de datos en formato CSV (separado por comas). Devuelve SOLO el código CSV.", "xw_tabla", "edit_xw_tabla",
                        [("Ordena por primera columna", "Ordena por primera columna")], "📥 Descargar Excel")

    with tabs_xw[6]:
        _panel_borrador("Guion Podcast", "Guion podcast. 3 locutores. Formato exacto: 'Locutor 1: [texto]'.", "xw_podcast", "edit_xw_pod",
                        [("Hazlo más gracioso", "Hazlo más gracioso")], None)
        if st.session_state.get("xw_podcast"):
            if st.button("🔊 Generar Audio", key="btn_xw_pod_aud"):
                with st.spinner("Generando audio..."):
                    aud = generar_podcast_3_voces(st.session_state.xw_podcast, "es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural")
                    if aud:
                        st.session_state.xw_podcast_audio = aud; st.success("✅ Audio listo.")
            if st.session_state.get("xw_podcast_audio"):
                st.audio(st.session_state.xw_podcast_audio, format="audio/mp3")
                st.download_button("💾 Descargar MP3", st.session_state.xw_podcast_audio, "podcast_xaviwork.mp3")

    with tabs_xw[7]:
        chat_container = st.container(height=400)
        with chat_container:
            for msg in st.session_state.xw_copilot_history:
                with st.chat_message(msg["role"]): st.write(msg["content"])
        if prompt_chat := st.chat_input("Pregunta sobre el tema o archivos..."):
            st.session_state.xw_copilot_history.append({"role": "user", "content": prompt_chat})
            with chat_container:
                with st.chat_message("user"): st.write(prompt_chat)
                with st.chat_message("assistant"):
                    with st.spinner("Pensando..."):
                        ctx_chat = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.xw_copilot_history])
                        res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nTema: {st.session_state.xw_tema_input}\n\nHistorial:\n{ctx_chat}\n\nResponde a la última pregunta.", st.session_state.api_key_gemini_manual)
                        st.write(res)
                        st.session_state.xw_copilot_history.append({"role": "assistant", "content": res})

    with tabs_xw[8]:
        _panel_borrador("Línea de Tiempo", "Extrae los eventos cronológicos. Formato: Año/Fecha - Evento.", "xw_timeline", "edit_xw_time",
                        [("Añade más fechas", "Añade más fechas")], "📥 Descargar")

    with tabs_xw[9]:
        _panel_borrador("Glosario", "Extrae los 5 conceptos más complejos y genera un glosario con definiciones sencillas.", "xw_glosario", "edit_xw_glos",
                        [("Explícalo para un niño", "Explícalo para un niño")], "📥 Descargar")

    with tabs_xw[10]:
        _panel_borrador("Análisis DAFO", "Realiza un Análisis DAFO (Debilidades, Amenazas, Fortalezas, Oportunidades). Sé conciso.", "xw_dafo", "edit_xw_dafo",
                        [("Hazlo más estratégico", "Hazlo más estratégico")], "📥 Descargar")

    with tabs_xw[11]:
        st.markdown("Genera una imagen para ilustrar tu proyecto.")
        prompt_img = st.text_input("Describe la imagen:", value=st.session_state.xw_tema_input, key="xw_img_prompt")
        if st.button("🖼️ Generar Portada", key="btn_xw_img", type="primary") and prompt_img:
            with st.spinner("Generando imagen..."):
                try:
                    url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img + ', cyberpunk, high quality')}?width=1024&height=1024&nologo=true"
                    if url_es_segura(url_img):
                        req = urllib.request.Request(url_img, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=30) as response:
                            st.session_state.xw_cover_img = response.read()
                except Exception as e: st.error(f"Error al generar imagen: {e}")
        if st.session_state.get("xw_cover_img"):
            st.image(st.session_state.xw_cover_img, caption="Portada Generada", use_container_width=True)
            st.download_button("📥 Descargar Imagen", st.session_state.xw_cover_img, "portada.png", "image/png", use_container_width=True)

# ==========================================
# CATEGORÍA 1: CHAT Y ASISTENTES
# ==========================================
elif macro_idx == 1:
    sub_names = ["🎙️ PODCAST", "💬 CHAT", "🎙️ VOZ", "💬 CHAT CON PDF", "🕵️‍♂️ AGENTES", "⚖️ DEBATE", "🍳 COCINA", "📖 CUENTOS", "💻 CÓDIGO", "🌍 TRADUCTOR"]
    sub_idx = custom_tabs(sub_names, "chat_sub", 0)

    if sub_idx == 0:
        st.markdown("#### 🎙️ Generador de Podcasts")
        contexto = cajita_adjuntar_archivo("podcast_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        tema = st.text_area("Temática principal:", placeholder="Ej: El futuro de la IA, La pizza con piña...", key="pod_tema_input", height=100)
        c1, c2 = st.columns(2)
        with c1: estilo = st.selectbox("Estilo:", ["Entrevista", "Debate", "Educativo", "Narrativo"], key="p_est")
        with c2: dur = st.slider("Duración (min):", 1, 10, 3, key="p_dur")
        st.markdown("---")
        if st.button("💡 Generar 3 Ideas con IA", use_container_width=True, key="pod_btn_ideas"):
            with st.spinner("Pensando ideas..."):
                res = llamar_ia(modelo_seleccionado, f"Dame 3 ideas creativas y concisas para un podcast de {dur} min, estilo {estilo}, sobre: '{tema}'. Solo ideas separadas por saltos de línea, sin numeración.", st.session_state.api_key_gemini_manual)
                if not res.startswith("[ERROR"):
                    st.session_state["pod_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
        def _add_idea_pod(idea):
            actual = st.session_state["pod_tema_input"].strip()
            st.session_state["pod_tema_input"] = f"{actual}\n\n{idea}" if actual else idea
            st.session_state["pod_sugs"] = []
        if st.session_state.get("pod_sugs"):
            st.markdown("✨ **Haz clic para añadir:**")
            cols = st.columns(len(st.session_state["pod_sugs"]))
            for i, idea in enumerate(st.session_state["pod_sugs"]):
                with cols[i]:
                    st.button(idea, key=f"pod_idea_{i}", use_container_width=True, on_click=_add_idea_pod, args=(idea,))
        st.markdown("---")
        if st.button("✨ Crear Guion", type="primary", key="btn_pod") and st.session_state["pod_tema_input"]:
            with st.spinner("Escribiendo..."):
                prompt = f"{contexto}\n\nGuion podcast {dur} min. Tema: '{st.session_state['pod_tema_input']}'. Estilo: {estilo}. 3 locutores. REGLA ESTRICTA: natural y coloquial. Usa SOLO formato 'Locutor 1: [texto]'."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if res.startswith("[ERROR"): st.error(res)
                else: st.session_state.podcast_script = res; guardar_en_db("Podcast", res)
        if st.session_state.get("podcast_script"):
            st.markdown("### 📜 Guion Generado")
            if "pod_edit" not in st.session_state:
                st.session_state.pod_edit = st.session_state.podcast_script
            st.text_area("Guion (Editable):", key="pod_edit", height=300)
            st.session_state.podcast_script = st.session_state.pod_edit
            c1, c2, c3 = st.columns(3)
            with c1: v1 = st.selectbox("Voz L1:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural"], key="v1")
            with c2: v2 = st.selectbox("Voz L2:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural"], index=1, key="v2")
            with c3: v3 = st.selectbox("Voz L3:", ["es-AR-TomasNeural", "es-ES-ElviraNeural"], key="v3")
            if st.button("🔊 Audio 3 Voces", key="btn_pod_aud"):
                with st.spinner("Generando audio..."):
                    aud = generar_podcast_3_voces(st.session_state.podcast_script, v1, v2, v3)
                    if aud: st.audio(aud, format="audio/mp3"); st.download_button("💾 MP3", aud, "podcast.mp3")
                    else: st.error("❌ Error al generar el audio.")
            botones_descarga(st.session_state.podcast_script, "Podcast")

    elif sub_idx == 1:
        st.markdown("#### 💬 Chat")
        for m in st.session_state.chat_history: st.chat_message(m["role"]).write(m["content"])
        if st.session_state.chat_history:
            c1, c2, c3 = st.columns(3)
            with c1: st.download_button("📥 TXT", "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history]), "chat.txt")
            with c2:
                if st.button("🗑️ Limpiar"): st.session_state.chat_history = []; st.rerun()
            with c3:
                if st.button("🔊 Leer"):
                    aud = generar_audio_sincrono(st.session_state.chat_history[-1]["content"], "es-ES-ElviraNeural")
                    if aud: st.audio(aud, format="audio/mp3")
        ctx = cajita_adjuntar_archivo("chat_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        p = chat_input_pro("main")
        if p:
            st.session_state.chat_history.append({"role": "user", "content": p}); st.chat_message("user").write(p)
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"{ctx}\n\n{p}", st.session_state.api_key_gemini_manual, st.session_state.chat_history)
                if res.startswith("[ERROR"): st.error(res)
                else:
                    st.session_state.chat_history.append({"role": "assistant", "content": res}); st.chat_message("assistant").write(res); guardar_en_db("Chat", res)
            st.rerun()

    elif sub_idx == 2:
        st.markdown("#### Voz")
        txt = st.text_area("Texto:")
        voz = st.selectbox("Voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural"])
        if st.button("🔊 Audio") and txt:
            aud = generar_audio_sincrono(txt, voz)
            if aud: st.audio(aud, format="audio/mp3")

    elif sub_idx == 3:
        st.markdown("#### Chat con PDFs")
        archs = st.file_uploader("Sube docs:", type=["pdf", "txt", "docx"], accept_multiple_files=True)
        if archs:
            texto = "".join([leer_archivo_universal(a, modelo_seleccionado) for a in archs])[:MAX_TEXTO_IA]
            st.success("Docs cargados.")
            for m in st.session_state.pdf_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("pdf")
            if p:
                st.session_state.pdf_chat_history.append({"role": "user", "content": p})
                with st.spinner("Buscando..."):
                    res = llamar_ia(modelo_seleccionado, f"Contexto:\n{texto}\n\nPregunta: {p}", st.session_state.api_key_gemini_manual, st.session_state.pdf_chat_history)
                    if res.startswith("[ERROR"): st.error(res)
                    else: st.session_state.pdf_chat_history.append({"role": "assistant", "content": res}); st.rerun()

    elif sub_idx == 4:
        st.markdown("#### Agentes")
        nom = st.text_input("Nombre:"); rol = st.text_area("Personalidad:")
        if st.button("Crear"):
            st.session_state.agente_activo = {"nombre": nom, "rol": rol}; st.session_state.agente_chat_history = []
        if st.session_state.get("agente_activo"):
            ag = st.session_state.agente_activo
            for m in st.session_state.agente_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("agent")
            if p:
                st.session_state.agente_chat_history.append({"role": "user", "content": p})
                with st.spinner("Pensando..."):
                    res = llamar_ia(modelo_seleccionado, f"ERES: {ag['rol']}\nUsuario: {p}", st.session_state.api_key_gemini_manual, st.session_state.agente_chat_history)
                    if res.startswith("[ERROR"): st.error(res)
                    else: st.session_state.agente_chat_history.append({"role": "assistant", "content": res}); st.rerun()

    elif sub_idx == 5:
        st.markdown("#### Debate")
        t = campo_con_ideas("Tema:", "deb", "temas de debate", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Coche eléctrico")
        if st.button("Iniciar") and t:
            res = llamar_ia(modelo_seleccionado, f"Debate {t}: 3 a favor, 3 en contra, conclusión.", st.session_state.api_key_gemini_manual)
            if res.startswith("[ERROR"): st.error(res)
            else: st.markdown(res); botones_descarga(res, "Debate")

    elif sub_idx == 6:
        st.markdown("#### Cocina")
        ing = campo_con_ideas("Ingredientes:", "chef1", "recetas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Pollo")
        pref = campo_con_ideas("Menú:", "chef2", "menús", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Vegetariano")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Receta") and ing:
                res = llamar_ia(modelo_seleccionado, f"Receta con {ing}", st.session_state.api_key_gemini_manual)
                if res.startswith("[ERROR"): st.error(res)
                else: st.info(res)
        with c2:
            if st.button("Menú") and pref:
                res = llamar_ia(modelo_seleccionado, f"Menú semanal {pref}", st.session_state.api_key_gemini_manual)
                if res.startswith("[ERROR"): st.error(res)
                else: st.info(res)

    elif sub_idx == 7:
        st.markdown("#### Cuentos")
        p = campo_con_ideas("Protagonista:", "c1", "protagonistas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Dragón")
        l = campo_con_ideas("Lugar:", "c2", "lugares mágicos", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Bosque")
        if st.button("Crear") and p and l:
            res = llamar_ia(modelo_seleccionado, f"Cuento sobre {p} en {l}.", st.session_state.api_key_gemini_manual)
            if res.startswith("[ERROR"): st.error(res)
            else:
                st.text_area("Cuento:", res, height=300); botones_descarga(res, "Cuento")
                if st.button("Leer"):
                    aud = generar_audio_sincrono(res, "es-ES-ElviraNeural")
                    if aud: st.audio(aud, format="audio/mp3")

    elif sub_idx == 8:
        st.markdown("#### Código")
        req = campo_con_ideas("Código:", "code", "scripts", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Renombrar archivos")
        if st.button("Procesar") and req:
            res = llamar_ia(modelo_seleccionado, f"Programador senior. Explica o mejora:\n{req}", st.session_state.api_key_gemini_manual)
            if res.startswith("[ERROR"): st.error(res)
            else: st.code(res, language="python")

    elif sub_idx == 9:
        st.markdown("#### Traductor")
        txt = campo_con_ideas("Texto:", "trad", "frases", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Texto...")
        c1, c2 = st.columns(2)
        with c1: o = st.selectbox("Origen:", ["Auto", "Español", "Inglés"], key="o")
        with c2: d = st.selectbox("Destino:", ["Español", "Inglés", "Francés"], key="d")
        if st.button("Traducir") and txt:
            res = llamar_ia(modelo_seleccionado, f"Traduce del {o} al {d}:\n{txt}", st.session_state.api_key_gemini_manual)
            if res.startswith("[ERROR"): st.error(res)
            else: st.info(res); botones_descarga(res, "Traduccion")

# ==========================================
# CATEGORÍA 2: DOCUMENTOS Y TEXTO
# ==========================================
elif macro_idx == 2:
    sub_names = ["📁 ANÁLISIS", "⚖️ COMPARAR", "📊 OFFICE", "📄 PDF", "📇 QR & CV", "📋 PORTAPAPELES"]
    sub_idx = custom_tabs(sub_names, "doc_sub", 0)

    if sub_idx == 0:
        st.markdown("#### Análisis de Documentos")
        arch = st.file_uploader("Sube PDF, Word, Excel o TXT", type=["pdf", "txt", "docx", "xlsx"])
        opc = st.selectbox("Análisis:", ["Resumen Ejecutivo", "Extraer Fechas y Cifras", "Buscar Errores", "Crear Test", "Traducir al Inglés"])
        if arch and st.button("🔎 Ejecutar", type="primary"):
            with st.spinner("Analizando..."):
                texto = leer_archivo_universal(arch, modelo_seleccionado)
                prompt = {"Resumen Ejecutivo": f"Resumen:\n{texto}", "Extraer Fechas y Cifras": f"Extrae fechas:\n{texto}",
                          "Buscar Errores": f"Errores:\n{texto}", "Crear Test": f"Test 5 preguntas:\n{texto}",
                          "Traducir al Inglés": f"Traduce al inglés:\n{texto}"}[opc]
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if res.startswith("[ERROR"): st.error(res)
                else: st.text_area("Resultado:", res, height=400); botones_descarga(res, "Analisis")

    elif sub_idx == 1:
        st.markdown("#### Comparador")
        a = st.file_uploader("Original:", type=["pdf", "txt"]); b = st.file_uploader("Modificado:", type=["pdf", "txt"])
        if a and b and st.button("Comparar"):
            with st.spinner("Buscando..."):
                ta = leer_archivo_universal(a, modelo_seleccionado); tb = leer_archivo_universal(b, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"Diferencias:\nA:\n{ta[:10000]}\nB:\n{tb[:10000]}", st.session_state.api_key_gemini_manual)
                if res.startswith("[ERROR"): st.error(res)
                else: st.text_area("Diferencias:", res, height=400); botones_descarga(res, "Comparativa")

    elif sub_idx == 2:
        st.markdown("#### Creación de Documentos (Borrador)")
        contexto = cajita_adjuntar_archivo("office_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        prompt_office = campo_con_ideas("¿Qué documento quieres crear?", "office", "documentos", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Presupuesto para cliente")
        formato = st.radio("Formato final:", ["Word", "Excel", "PowerPoint"], key="office_fmt", horizontal=True)

        if st.button("✨ Generar Borrador de Texto", type="primary", key="btn_office_gen"):
            if prompt_office:
                with st.spinner("Generando borrador..."):
                    prompt_ia = f"{contexto}\n\n{prompt_office}"
                    if formato == "PowerPoint": prompt_ia += "\n\nIMPORTANTE: Formatea para diapositivas. Título, salto de línea, guiones para puntos. Doble salto de línea entre diapositivas."
                    elif formato == "Excel": prompt_ia += "\n\nIMPORTANTE: Devuelve los datos en formato CSV separados por comas."
                    res = llamar_ia(modelo_seleccionado, prompt_ia, st.session_state.api_key_gemini_manual)
                    if res.startswith("[ERROR"): st.error(res)
                    else: st.session_state.office_borrador = res
            else: st.warning("Escribe un prompt primero.")

        if st.session_state.get("office_borrador"):
            st.markdown("### 📝 Borrador Editable")
            if "office_txt" not in st.session_state: st.session_state.office_txt = st.session_state.office_borrador
            st.text_area("Modifica tu documento antes de descargar:", key="office_txt", height=350)
            st.session_state.office_borrador = st.session_state.office_txt
            c1, c2, c3 = st.columns(3)
            with c1: boton_superpoder_borrador(st.session_state.office_txt, "Hazlo más profesional", "of_sup1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador")
            with c2: boton_superpoder_borrador(st.session_state.office_txt, "Resúmelo", "of_sup2", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador")
            with c3: boton_superpoder_borrador(st.session_state.office_txt, "Añade más detalles", "of_sup3", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador")
            st.markdown("---")
            if formato == "Word": st.download_button("📥 Descargar Word", generar_word(st.session_state.office_txt), "doc.docx")
            elif formato == "Excel": st.download_button("📥 Descargar Excel", generar_excel(st.session_state.office_txt
