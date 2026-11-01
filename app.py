"""
Xavilu Network Pro - Fusión 3.0 (Completo)
Interfaz unificada de asistentes IA, documentos, web y multimedia.
Las API keys van en secrets.toml (Streamlit Cloud → Settings → Secrets).
"""
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
import ipaddress
import socket
import csv
import os
import pandas as pd
from youtube_transcript_api import YouTubeTranscriptApi
import re
import requests
import tempfile
import time
import qrcode
import numpy as np
import sqlite3
import datetime
import concurrent.futures

try:
    import cv2
    HAS_CV2 = True
except ImportError:
    HAS_CV2 = False

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

try:
    import openai
    HAS_OPENAI = True
except ImportError:
    HAS_OPENAI = False

try:
    import anthropic
    HAS_ANTHROPIC = True
except ImportError:
    HAS_ANTHROPIC = False


# ==========================================
# 1. CONFIGURACIÓN
# ==========================================
st.set_page_config(
    page_title="Xavilu Network Pro",
    page_icon="logo.png",
    layout="wide",
    initial_sidebar_state="expanded",
)

XAVILU_VERSION = "3.0 - Fusión general"
MAX_FILE_MB = 100
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-2.0-flash"

if os.name == "nt":
    asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())


# ==========================================
# 2. BASE DE DATOS (solo historial)
# ==========================================
def init_db():
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        conn.execute(
            "CREATE TABLE IF NOT EXISTS historial "
            "(id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, tipo TEXT, contenido TEXT)"
        )


def guardar_en_db(tipo, contenido):
    try:
        with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
            conn.execute(
                "INSERT INTO historial (fecha, tipo, contenido) VALUES (?, ?, ?)",
                (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo, contenido),
            )
    except sqlite3.Error as e:
        st.warning(f"DB historial (guardar): {e}")


def obtener_de_db():
    try:
        with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
            return conn.execute(
                "SELECT fecha, tipo, contenido FROM historial ORDER BY id DESC LIMIT 5"
            ).fetchall()
    except sqlite3.Error as e:
        st.warning(f"DB historial (leer): {e}")
        return []


init_db()


# ==========================================
# 3. ESTADO DE SESIÓN
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
    # --- XaviWork ---
    "xw_mapa": "", "xw_ppt": "", "xw_informe": "", "xw_tarjetas": "", "xw_cuestionario": "",
    "xw_tabla": "", "xw_podcast": "", "xw_podcast_audio": None, "xw_timeline": "", "xw_glosario": "",
    "xw_dafo": "", "xw_cover_img": None, "xw_copilot_history": [], "xw_tema_input": "",
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v


# ==========================================
# 4. CSS
# ==========================================
st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
.stApp { background-color: #000000 !important; background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
header[data-testid="stHeader"] { background-color: #000000 !important; background: #000000 !important; }
[data-testid="stAppDeployButton"] { display: none !important; }
[data-testid="stDecoration"] { display: none !important; }
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700, 0 0 35px #FF8C00 !important; text-align: center; letter-spacing: 3px; margin-top: 0px !important; padding-top: 0px !important; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #ffffff !important; text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff, 0 0 30px #FF007F !important; text-align: center; letter-spacing: 3px; }
h4 { font-family: 'Orbitron', sans-serif !important; color: #00f0ff !important; text-shadow: 0 0 8px #00f0ff, 0 0 15px #FF007F !important; letter-spacing: 1px; }
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
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000000 !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: 900 !important; border: 2px solid #ffffff !important; border-radius: 25px !important; padding: 10px 15px !important; margin-bottom: 5px !important; transition: all 0.3s ease !important; box-shadow: 0 0 18px rgba(0, 240, 255, 0.9), 0 0 32px rgba(255, 0, 127, 0.7) !important; transform: translateY(-1px) !important; width: 100% !important; }
.stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; color: #ffffff !important; font-size: 16px !important; box-shadow: inset 0 0 10px rgba(0, 240, 255, 0.1) !important; }
.stTextArea textarea:focus, .stTextInput input:focus { border: 1px solid #FF007F !important; box-shadow: 0 0 15px rgba(255, 0, 127, 0.5) !important; }
[data-testid="stTextInput"] div, [data-testid="stTextInput"] input { background-color: #0a0a0a !important; color: #ffffff !important; }
[data-testid="stTextInput"] > div > div { border: none !important; box-shadow: none !important; }
[data-testid="stTextInput"] input { caret-color: #00f0ff !important; border: 1px solid #00f0ff !important; border-radius: 10px !important; padding: 10px 14px !important; font-size: 15px !important; outline: none !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.2) !important; transition: all 0.3s ease !important; width: 100% !important; }
[data-testid="stTextInput"] input:focus { border: 1px solid #FF007F !important; box-shadow: 0 0 18px rgba(255, 0, 127, 0.6) !important; }
[data-testid="stTextInput"] input::placeholder { color: #888888 !important; opacity: 1 !important; }
div.stButton > button:not([kind="primary"]):not([kind="secondary"]), div[data-testid="stDownloadButton"] > button { background: #0a0a0a !important; color: #00f0ff !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; font-family: 'Orbitron', sans-serif !important; font-weight: bold !important; padding: 10px 20px !important; transition: all 0.3s !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.3) !important; }
div.stButton > button:not([kind="primary"]):not([kind="secondary"]):hover, div[data-testid="stDownloadButton"] > button:hover { background: linear-gradient(135deg, #FF007F 0%, #00f0ff 100%) !important; color: #000 !important; border: 1px solid #fff !important; transform: translateY(-2px); box-shadow: 0 0 20px rgba(255, 0, 127, 0.8) !important; }
[data-testid="stChatMessage"] { background: rgba(10, 10, 15, 0.8) !important; border: 1px solid #FF007F !important; border-radius: 15px !important; padding: 10px; box-shadow: 0 0 15px rgba(255, 0, 127, 0.3) !important; }
.stAlert { border-radius: 10px; font-weight: bold; background-color: rgba(0, 240, 255, 0.1) !important; border-left: 5px solid #00f0ff !important; }
[data-testid="stVerticalBlock"] > div { padding-bottom: 0.5rem !important; }

/* ELIMINAR HUECO SUPERIOR */
.block-container {
    padding-top: 1rem !important;
    padding-bottom: 1rem !important;
}

@media (max-width: 600px) {
    .block-container { padding-top: 0.5rem !important; }
    h1 { font-size: 2.2rem !important; letter-spacing: 1px !important; margin-top: 0px !important; }
    [data-testid="stWidgetLabel"] p { font-size: 11px !important; padding: 3px 8px !important; }
    div.stButton > button { padding: 6px 10px !important; font-size: 12px !important; margin-bottom: 3px !important; }
    .stTextArea textarea { font-size: 14px !important; }
}
</style>
""", unsafe_allow_html=True)

def inyectar_js_fixer():
    components.html(
        "<script>(function(){var doc=window.parent.document;"
        "function f(){doc.querySelectorAll('[data-testid=\"stTextInput\"] *').forEach(function(el){"
        "if(el.tagName.toLowerCase()==='input'||el.tagName.toLowerCase()==='div'){"
        "el.style.setProperty('background-color','#0a0a0a','important');"
        "el.style.setProperty('color','#ffffff','important');}});}f();setInterval(f,300);})();</script>",
        height=0,
    )

inyectar_js_fixer()


# ==========================================
# 5. UTILIDADES Y SEGURIDAD
# ==========================================
def url_segura(url: str) -> bool:
    """Bloquea SSRF: solo http/https, sin IPs privadas, loopback ni link-local."""
    try:
        p = urllib.parse.urlparse(url)
        if p.scheme not in ("http", "https") or not p.hostname:
            return False
        host = socket.gethostbyname(p.hostname)
        ip = ipaddress.ip_address(host)
        return not (
            ip.is_private or ip.is_loopback or ip.is_link_local
            or ip.is_reserved or ip.is_multicast
        )
    except Exception:
        return False

def fetch_seguro(url: str, timeout: int = 15, headers: dict = None):
    if not url_segura(url):
        raise ValueError(f"URL bloqueada por seguridad: {url}")
    h = {"User-Agent": "Mozilla/5.0 (compatible; Xavilu/3.0)"}
    if headers:
        h.update(headers)
    return requests.get(url, headers=h, timeout=timeout)

def leer_archivo_universal(upload, modelo_ia: str) -> str:
    if upload.size > MAX_FILE_MB * 1024 * 1024:
        st.error(f"Archivo supera {MAX_FILE_MB} MB")
        return ""
    nombre = upload.name.lower()
    try:
        if nombre.endswith(".txt"):
            return upload.read().decode("utf-8", errors="ignore")
        if nombre.endswith(".pdf"):
            reader = pypdf.PdfReader(BytesIO(upload.getvalue()))
            return "\n".join(p.extract_text() for p in reader.pages if p.extract_text())
        if nombre.endswith(".docx"):
            return "\n".join(p.text for p in Document(upload).paragraphs if p.text)
        if nombre.endswith(".xlsx"):
            df = pd.read_excel(upload, sheet_name=None)
            return "\n\n".join(f"=== {sheet} ===\n{data.to_csv(index=False)}" for sheet, data in df.items())
    except Exception as e:
        st.warning(f"No pude leer '{upload.name}': {e}")
        return ""
    return "Formato no reconocido."

@st.cache_resource(show_spinner=False)
def get_gemini_client(api_key: str):
    return genai.Client(api_key=api_key)

def buscar_en_internet(query: str):
    if DDGS is None: return []
    try:
        res = DDGS().text(query, region="es-es", max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in res]
    except Exception as e:
        st.warning(f"Búsqueda web: {e}")
        return []

def buscar_noticias_deportivas(deporte: str, liga: str):
    if DDGS is None: return []
    try:
        res = DDGS().news(f"últimas noticias {deporte} {liga} hoy", max_results=5)
        return [{"titulo": r.get("title", ""), "url": r.get("url", ""), "resumen": r.get("body", ""), "fecha": r.get("date", "")} for r in res]
    except Exception as e:
        st.warning(f"Noticias: {e}")
        return []

def extraer_youtube_id(url: str):
    if not url: return None
    m = re.search(r"(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([A-Za-z0-9_-]{11})", url)
    if m: return m.group(1)
    if len(url.strip()) == 11: return url.strip()
    return None

def transcribir_audio_youtube(url: str, api_key_groq: str) -> str:
    if not HAS_YTDLP or not HAS_OPENAI:
        return "[ERROR] Faltan yt-dlp u openai instalados."
    if not api_key_groq:
        return "[ERROR] No hay Groq API Key."

    cookies_path = None
    try:
        cookie_content = st.secrets.get("YT_COOKIES")
        if cookie_content:
            tmp_cookie = tempfile.NamedTemporaryFile(delete=False, mode="w", suffix=".txt")
            tmp_cookie.write(cookie_content)
            tmp_cookie.close()
            cookies_path = tmp_cookie.name
    except Exception:
        pass

    tmp_dir = tempfile.mkdtemp()
    audio_path = None
    try:
        ydl_opts = {
            "format": "bestaudio[abr<=64]/worstaudio[abr<=64]/worstaudio",
            "outtmpl": os.path.join(tmp_dir, "audio.%(ext)s"),
            "quiet": True, "noplaylist": True,
            "postprocessors": [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "48"}],
        }
        if cookies_path:
            ydl_opts["cookiefile"] = cookies_path

        try:
            ydl_opts["impersonate"] = "chrome"
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=True)
        except Exception:
            ydl_opts.pop("impersonate", None)
            with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                info = ydl.extract_info(url, download=True)

        audio_path = ydl.prepare_filename(info)
        if not audio_path.endswith(".mp3"):
            base = os.path.splitext(audio_path)[0]
            if os.path.exists(base + ".mp3"):
                audio_path = base + ".mp3"

        if not os.path.exists(audio_path):
            return "[ERROR] No se generó el archivo de audio."

        size_mb = os.path.getsize(audio_path) / (1024 * 1024)
        if size_mb > 24:
            return f"[ERROR] Audio de {size_mb:.1f} MB, supera 25 MB de Groq."

        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e:
        return f"[ERROR] {e}"
    finally:
        if audio_path and os.path.exists(audio_path):
            try: os.remove(audio_path)
            except OSError: pass
        if cookies_path and os.path.exists(cookies_path):
            try: os.remove(cookies_path)
            except OSError: pass
        try: os.rmdir(tmp_dir)
        except OSError: pass

def transcribir_archivo_local(archivo_bytes: bytes, nombre_archivo: str, api_key_groq: str) -> str:
    if not HAS_OPENAI: return "[ERROR] Falta openai instalado."
    if not api_key_groq: return "[ERROR] No hay Groq API Key."
    ext = os.path.splitext(nombre_archivo)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp:
        tmp.write(archivo_bytes)
        audio_path = tmp.name
    try:
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f:
            resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e:
        return f"[ERROR] {e}"
    finally:
        if os.path.exists(audio_path):
            try: os.remove(audio_path)
            except OSError: pass

def cajita_adjuntar_archivo(key_suffix: str, modelo: str, api_key: str) -> str:
    with st.expander("📎 Adjuntar archivo para esta acción (Opcional)"):
        archivo = st.file_uploader("Sube PDF, Word, TXT o Imagen", type=["pdf", "txt", "docx", "png", "jpg", "jpeg"], key=f"adj_{key_suffix}")
        if not archivo: return ""
        nombre = archivo.name.lower()
        try:
            if nombre.endswith((".png", ".jpg", ".jpeg")):
                with st.spinner("Analizando imagen..."):
                    ctx = analizar_archivo_con_gemini(archivo.getvalue(), "image/jpeg", "Describe esta imagen o extrae el texto (OCR).", api_key)
                st.success("✅ Imagen leída.")
                return f"\n[TEXTO DE IMAGEN]:\n{ctx}\n"
            texto = leer_archivo_universal(archivo, modelo)
            st.success(f"✅ Archivo '{archivo.name}' leído.")
            return f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"
        except Exception as e:
            st.warning(f"No pude procesar '{archivo.name}': {e}")
            return ""

def analizar_archivo_con_gemini(archivo_bytes: bytes, mime_type: str, prompt: str, api_key_gemini: str) -> str:
    if not HAS_GEMINI or not api_key_gemini: return "[ERROR] Falta google-genai o API key."
    size_mb = len(archivo_bytes) / (1024 * 1024)
    tmp_path = None
    try:
        client = get_gemini_client(api_key_gemini)
        if size_mb > 15:
            ext = mime_type.split("/")[-1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp:
                tmp.write(archivo_bytes)
                tmp_path = tmp.name
            uploaded = client.files.upload(file=tmp_path)
            while uploaded.state.name == "PROCESSING":
                time.sleep(3)
                uploaded = client.files.get(name=uploaded.name)
            if uploaded.state.name == "FAILED": return "[ERROR] Gemini no pudo procesar el archivo."
            res = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, uploaded]).text
            try: client.files.delete(name=uploaded.name)
            except Exception: pass
            return res
        parte = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
        return client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, parte]).text
    except Exception as e:
        return f"[ERROR GEMINI] {e}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except OSError: pass

def _get_key(nombre: str, arg_directo: str = "") -> str:
    manual = st.session_state.get(f"api_key_{nombre.lower()}_manual", "").strip()
    if manual: return manual
    try:
        secret = st.secrets.get(f"{nombre.upper()}_API_KEY", "")
        if secret: return secret
    except Exception: pass
    return arg_directo or ""

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None, api_key_openai="", api_key_anthropic=""):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if st.session_state.get("global_context"):
        prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}\n\n(Responde teniendo en cuenta el archivo)."
    if historial_chat:
        contexto_chat = "\n".join(f"{msg['role'].upper()}: {msg['content']}" for msg in historial_chat[-10:])
        prompt_final = f"Historial:\n{contexto_chat}\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."

    key_gemini = _get_key("gemini", api_key_gemini)
    key_groq = _get_key("groq")
    key_silicon = _get_key("silicon")
    key_mistral = _get_key("mistral")
    key_openrouter = _get_key("openrouter")

    if key_gemini and HAS_GEMINI:
        for intento in range(2):
            try:
                res = get_gemini_client(key_gemini).models.generate_content(model=GEMINI_MODEL, contents=prompt_final).text
                if res: return res
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg: break
                if "503" in msg and intento < 1:
                    time.sleep(2)
                    continue
                break

    if key_groq and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_groq, base_url="https://api.groq.com/openai/v1")
            res = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except Exception: pass

    if key_silicon and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_silicon, base_url="https://api.siliconflow.cn/v1")
            res = client.chat.completions.create(model="Qwen/Qwen2.5-7B-Instruct", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except Exception: pass

    if key_mistral and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_mistral, base_url="https://api.mistral.ai/v1")
            res = client.chat.completions.create(model="mistral-small-latest", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except Exception: pass

    if key_openrouter and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_openrouter, base_url="https://openrouter.ai/api/v1")
            res = client.chat.completions.create(model="deepseek/deepseek-r1:free", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except Exception: pass

    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=600)
        raw = client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})["message"]["content"]
        raw = re.sub(r"^```\w*\n", "", raw)
        raw = re.sub(r"\n```\s*$", "", raw)
        return raw.strip()
    except Exception as e:
        return f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}"


# ==========================================
# 6. GENERADORES DE DOCUMENTOS
# ==========================================
def obtener_ruta_fuente():
    candidatos = [
        "/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf",
        "/usr/share/fonts/dejavu/DejaVuSans.ttf",
        "C:/Windows/Fonts/arial.ttf",
        "C:/Windows/Fonts/DejaVuSans.ttf",
        "/Library/Fonts/Arial.ttf",
        "/System/Library/Fonts/Supplemental/Arial.ttf",
        os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf") if "__file__" in globals() else "",
    ]
    for ruta in candidatos:
        if ruta and os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo: str, contenido: str) -> bytes:
    pdf = FPDF()
    pdf.add_page()
    font_path = obtener_ruta_fuente()
    try:
        url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Book cover for ' + titulo + ', cyberpunk')}?width=512&height=512&nologo=true"
        with urllib.request.urlopen(urllib.request.Request(url_img, headers={"User-Agent": "Mozilla/5.0"}), timeout=15) as response:
            img_data = response.read()
        with open("temp_cover.jpg", "wb") as f:
            f.write(img_data)
        pdf.image("temp_cover.jpg", x=40, y=50, w=130)
        pdf.ln(150)
        os.remove("temp_cover.jpg")
    except Exception: pass

    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True)
        pdf.set_font("XaviFont", size=11)
    else:
        st.warning("No se encontró fuente Unicode. Pon 'DejaVuSans.ttf' junto al script para evitar pérdida de acentos.")
        pdf.set_font("Helvetica", size=11)

    pdf.set_font("Helvetica", "B", 24)
    pdf.multi_cell(0, 10, txt=titulo, align="C")
    pdf.set_font("Helvetica", "I", 12)
    pdf.cell(0, 10, txt="Generado por Xavilu Network", align="C", ln=True)
    pdf.set_draw_color(0, 240, 255)
    pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5)
    pdf.ln(15)

    if font_path: pdf.set_font("XaviFont", size=11)
    else: contenido = contenido.encode("latin-1", "replace").decode("latin-1")

    pdf.multi_cell(0, 8, txt=contenido)
    return bytes(pdf.output())

def generar_word(texto: str) -> bytes:
    doc = Document()
    doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO()
    doc.save(out)
    return out.getvalue()

def generar_excel(texto_csv: str) -> bytes:
    wb = openpyxl.Workbook()
    ws = wb.active
    csv_limpio = texto_csv.replace("```csv", "").replace("```", "").strip()
    for r_i, row in enumerate(csv.reader(csv_limpio.split("\n")), 1):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            if r_i == 1: cell.font = Font(bold=True)
    out = BytesIO()
    wb.save(out)
    return out.getvalue()

def generar_pptx(titulo: str, texto: str) -> bytes:
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        lineas = bloque.split("\n")
        if slide.shapes.title: slide.shapes.title.text = lineas[0][:50]
        contenido = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
        tiene_ph = False
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1:
                shape.text = contenido
                tiene_ph = True
        if not tiene_ph:
            tx = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5))
            tx.text_frame.text = contenido
    out = BytesIO()
    prs.save(out)
    return out.getvalue()

def botones_descarga(texto: str, nombre_por_defecto: str = "documento", formatos_extra: bool = False):
    if not texto: return
    st.markdown("---")
    key_u = nombre_por_defecto.replace(" ", "_")
    nombre_elegido = st.text_input("📝 Nombre del archivo (sin extensión):", value=nombre_por_defecto, key=f"name_{key_u}")
    nombre_limpio = re.sub(r"[^a-zA-Z0-9_\- ]", "", nombre_elegido).strip().replace(" ", "_") or nombre_por_defecto

    if formatos_extra:
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.download_button("📥 TXT", texto.encode("utf-8"), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
        with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
        with c3: st.download_button("📝 Word", generar_word(texto), f"{nombre_limpio}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"wd_{key_u}", use_container_width=True)
        with c4: st.download_button("📊 Excel", generar_excel(texto), f"{nombre_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"ex_{key_u}", use_container_width=True)
    else:
        c1, c2, c3 = st.columns(3)
        with c1: st.download_button("📥 TXT", texto.encode("utf-8"), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
        with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
        with c3:
            if st.button("📲 Enviar a WhatsApp", key=f"wa_{key_u}", use_container_width=True):
                num = st.session_state.get("whatsapp_number", "").strip()
                if num:
                    url_wa = f"https://wa.me/{num}?text={urllib.parse.quote(texto[:1500])}"
                    components.html(f"<script>window.open('{url_wa}', '_blank');</script>", height=0)
                    st.success("✅ Abriendo WhatsApp...")
                else:
                    st.error("⚠️ Configura tu número en la barra lateral izquierda.")


# ==========================================
# 7. AUDIO / TTS
# ==========================================
def run_async_safe(coro):
    def _run(): return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor() as pool:
        return pool.submit(_run).result()

def generar_audio_sincrono(texto: str, voz: str) -> bytes:
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(texto, voz).stream():
            if chunk["type"] == "audio": buf.write(chunk["data"])
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e:
        st.error(f"Error TTS: {e}")
        return None

def generar_podcast_3_voces(guion: str, v1: str, v2: str, v3: str) -> bytes:
    async def _gen():
        buf = BytesIO()
        texto = re.sub(r"\(risas\)|\[risas\]|\*risas\*", "jajaja", guion, flags=re.IGNORECASE)
        texto = re.sub(r"\(pausa\)|\[pausa\]", "...", texto, flags=re.IGNORECASE)
        texto = re.sub(r"[*#_]", "", texto)
        voz_actual = v1
        for linea in texto.split("\n"):
            linea = linea.strip()
            if not linea: continue
            m = re.search(r"^Locutor\s*(\d)", linea, re.IGNORECASE)
            if m:
                num = int(m.group(1))
                voz_actual = {1: v1, 2: v2, 3: v3}.get(num, v1)
                linea = re.sub(r"^Locutor\s*\d\s*:?", "", linea, flags=re.IGNORECASE).strip()
            if not linea or linea.lower() in {"elvira", "jorge", "tomas", "carlos"}: continue
            try:
                async for chunk in edge_tts.Communicate(linea + ", ", voz_actual).stream():
                    if chunk["type"] == "audio": buf.write(chunk["data"])
            except Exception: continue
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e:
        st.error(f"Error crítico en generación 3 voces: {e}")
        return None


# ==========================================
# 8. HELPERS DE UI
# ==========================================
def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="...", api_key_openai="", api_key_anthropic=""):
    input_key = f"{key_prefix}_input"
    if input_key not in st.session_state: st.session_state[input_key] = ""

    def agregar_idea(idea):
        txt = st.session_state[input_key].strip()
        st.session_state[input_key] = f"{txt}\n\n{idea}" if txt else idea
        if f"{key_prefix}_sugs" in st.session_state: del st.session_state[f"{key_prefix}_sugs"]

    if st.button("💡 Generar Ideas con IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
        with st.spinner("Pensando ideas..."):
            prompt = f"Dame 3 ideas creativas para {contexto_ia}. Solo lista separada por saltos de línea, sin numeración."
            res = llamar_ia(modelo, prompt, api_key, api_key_openai=api_key_openai, api_key_anthropic=api_key_anthropic)
            st.session_state[f"{key_prefix}_sugs"] = [l.strip("-*1234567890. ") for l in res.split("\n") if len(l.strip()) > 5][:3]

    if st.session_state.get(f"{key_prefix}_sugs"):
        st.markdown("✨ **Haz clic en una o varias ideas para añadirlas a tu texto:**")
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]:
                st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True, on_click=agregar_idea, args=(idea,))
    return st.text_area(label, key=input_key, placeholder=placeholder, height=100)

def boton_superpoder_borrador(texto_base, instruccion, key_suffix, modelo, api_key, state_key, api_key_openai="", api_key_anthropic="", widget_key=None):
    if st.button(f"🚀 {instruccion}", key=f"super_{key_suffix}", type="secondary", use_container_width=True):
        if not texto_base: return
        with st.spinner("Aplicando superpoder..."):
            prompt = f"Actúa como experto. {instruccion}. Devuelve SOLO el texto mejorado, sin introducciones:\n\n{texto_base}"
            res = llamar_ia(modelo, prompt, api_key, api_key_openai=api_key_openai, api_key_anthropic=api_key_anthropic)
            if "[ERROR" in res: st.error(res)
            else:
                st.session_state[state_key] = res
                if widget_key and widget_key in st.session_state: del st.session_state[widget_key]
                st.rerun()

def borrador_editable(state_key: str, widget_key: str, height: int = 300) -> str:
    if widget_key not in st.session_state: st.session_state[widget_key] = st.session_state[state_key]
    st.text_area("Editar:", key=widget_key, height=height)
    st.session_state[state_key] = st.session_state[widget_key]
    return st.session_state[state_key]

def chat_input_pro(key: str, placeholder: str = "Mensaje..."):
    t_k, m_k = f"{key}_txt", f"{key}_msg"
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
        chunk = tab_names[i:i + 5]
        cols = st.columns(5)
        for j, name in enumerate(chunk):
            idx = i + j
            with cols[j]:
                if st.button(name, key=f"btn_{key_name}_{idx}", use_container_width=True, type="primary" if st.session_state[s_key] == idx else "secondary"):
                    st.session_state[s_key] = idx
                    st.rerun()
    return st.session_state[s_key]


# ==========================================
# 9. INTERFAZ
# ==========================================
if not st.session_state.sidebar_visible:
    st.markdown("<style>[data-testid=\"collapsedControl\"] {display: none !important;}</style>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if st.button("⛶ MOSTRAR MENÚ Y KEYS", type="primary", use_container_width=True):
            st.session_state.sidebar_visible = True
            st.rerun()

st.markdown(f"<h1>XAVILU NETWORK</h1><p style='text-align:center; color:#FFD700; font-weight:bold; letter-spacing:2px; text-shadow: 0 0 10px #FFD700;'>● v{XAVILU_VERSION}</p>", unsafe_allow_html=True)

if st.session_state.sidebar_visible:
    with st.sidebar:
        try: st.image("logo.png", use_container_width=True)
        except Exception: st.markdown("### 🪐 PANEL DE CONTROL")
        if st.button("⛶ Ocultar Menú", use_container_width=True):
            st.session_state.sidebar_visible = False
            st.rerun()

        st.markdown("---")
        modelo_seleccionado = st.selectbox("Ollama Local", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])

        st.markdown("#### 🔑 Estado de Conexiones")
        for prov in ["gemini", "groq", "silicon", "mistral", "openrouter"]:
            conectado = bool(_get_key(prov))
            etiqueta = "✅" if conectado else "⚠️"
            st.caption(f"{etiqueta} {prov.capitalize()}")

        with st.expander("⚙️ Keys manuales (solo si no usas secrets.toml)"):
            st.caption("En Streamlit Cloud las keys van en **Settings → Secrets**.")
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
                st.session_state.global_context = ""
                st.session_state.global_file_name = ""
                st.rerun()

        st.markdown("---")
        for h in obtener_de_db():
            with st.expander(f"{h[1]} - {h[0][:10]}"): st.caption(h[2][:200] + "...")
else:
    modelo_seleccionado = "llama3.2:latest"

macro_names = ["🛠️ XAVIWORK", "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR", "🎵 MÚSICA Y ARTE"]
macro_idx = custom_tabs(macro_names, "macro", 0)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)


# ==========================================
# CATEGORÍA 0: XAVIWORK
# ==========================================
if macro_idx == 0:
    st.markdown("#### 🛠️ XaviWork - El Estudio Definitivo")
    contexto_xw = cajita_adjuntar_archivo("xw_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
    
    col1, col2 = st.columns([2, 1])
    with col1:
        st.text_area("💡 ¿Sobre qué quieres trabajar? (Ideas propias, texto o archivo):", height=100, key="xw_tema_input", placeholder="Ej: Estrategia de marketing para CreatorFlow...")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💡 Generar 3 Ideas Base", use_container_width=True):
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"Dame 3 ideas creativas y concisas sobre: '{st.session_state.xw_tema_input}'. Sepáralas por saltos de línea.", st.session_state.api_key_gemini_manual)
                st.session_state["xw_sugs"] = [l.strip("-*1234567890. ") for l in res.split("\n") if len(l.strip()) > 5][:3]
        def _xw_add(idea):
            actual = st.session_state["xw_tema_input"].strip()
            st.session_state["xw_tema_input"] = f"{actual}\n{idea}" if actual else idea
            st.session_state["xw_sugs"] = []
        if st.session_state.get("xw_sugs"):
            for i, idea in enumerate(st.session_state["xw_sugs"]):
                st.button(f"➕ {idea}", key=f"xw_idea_{i}", use_container_width=True, on_click=_xw_add, args=(idea,))

    st.markdown("---")
    tabs_xw = st.tabs(["🕸️ Mapa", "📊 PPT", "📄 Informe", "🃏 Tarjetas", "❓ Test", "📋 Tabla", "🎙️ Podcast", "💬 Copiloto", "📅 Timeline", "📖 Glosario", "📊 DAFO", "🖼️ Portada"])

    def _panel_borrador(nombre, prompt_gen, state_key, edit_key, superpowers, download_label=None):
        if st.button(f"✨ Generar {nombre}", key=f"btn_{state_key}", type="primary"):
            if not st.session_state.xw_tema_input.strip():
                st.warning("Define el tema arriba primero.")
            else:
                with st.spinner(f"Generando {nombre}..."):
                    full = f"{contexto_xw}\n\n{prompt_gen}\n\nTEMA: {st.session_state.xw_tema_input}"
                    res = llamar_ia(modelo_seleccionado, full, st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state[state_key] = res.replace("```markdown","").replace("```csv","").replace("```","").strip() if state_key in ("xw_mapa","xw_tabla") else res

        if st.session_state.get(state_key):
            borrador_editable(state_key, edit_key, height=250)
            n = len(superpowers) + (1 if download_label else 0)
            if n == 0: return
            cols = st.columns(n)
            for i, (label, prompt) in enumerate(superpowers):
                with cols[i]:
                    boton_superpoder_borrador(st.session_state[state_key], prompt, f"{state_key}_p{i}", modelo_seleccionado, st.session_state.api_key_gemini_manual, state_key, widget_key=edit_key)
            if download_label:
                with cols[-1]:
                    if state_key == "xw_tabla": st.download_button(download_label, generar_excel(st.session_state[state_key]), "tabla.xlsx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key == "xw_ppt": st.download_button(download_label, generar_pptx("Presentación XaviWork", st.session_state[state_key]), "presentacion.pptx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key in ("xw_informe", "xw_timeline", "xw_glosario", "xw_dafo"): botones_descarga(st.session_state[state_key], state_key, formatos_extra=True)
                    else: st.download_button(download_label, st.session_state[state_key].encode(), f"{state_key}.txt", key=f"dl_{state_key}", use_container_width=True)

    with tabs_xw[0]:
        _panel_borrador("Mapa Mental", "Genera un mapa mental jerárquico en Markdown. Usa #, ## y -. Devuelve SOLO el Markdown.", "xw_mapa", "edit_xw_mapa", [("Hazlo más profesional", "Hazlo más profesional"), ("Añade más detalles", "Añade más detalles")], "📥 Descargar Markdown")
        if st.session_state.get("xw_mapa"):
            with st.expander("🕸️ Ver mapa interactivo"): markmap(st.session_state.xw_mapa, height=400)

    with tabs_xw[1]:
        _panel_borrador("Presentación (PPTX)", "Crea una estructura de presentación. Separa cada diapositiva con doble salto de línea. Primer renglón = título, guiones = puntos.", "xw_ppt", "edit_xw_ppt", [("Hazlo más visual y directo", "Hazlo más visual y directo")], "📥 Descargar PPTX")

    with tabs_xw[2]:
        _panel_borrador("Informe Detallado", "Redacta un informe profesional y detallado. Incluye introducción, desarrollo y conclusiones.", "xw_informe", "edit_xw_informe", [("Hazlo más técnico", "Hazlo más técnico"), ("Resumir a la mitad", "Resumir a la mitad"), ("Traducir al Inglés", "Traducir al Inglés")], "📥 Descargar")

    with tabs_xw[3]:
        _panel_borrador("Tarjetas (Flashcards)", "Crea 5 tarjetas didácticas (pregunta y respuesta). Formato: Pregunta -> Respuesta.", "xw_tarjetas", "edit_xw_tarjetas", [("Hazlas más difíciles", "Hazlas más difíciles")], "📥 Descargar")

    with tabs_xw[4]:
        _panel_borrador("Cuestionario", "Crea un cuestionario de 5 preguntas tipo test (3 opciones + respuesta al final).", "xw_cuestionario", "edit_xw_cuest", [("Añade 5 preguntas más", "Añade 5 preguntas más")], "📥 Descargar")

    with tabs_xw[5]:
        _panel_borrador("Tabla de Datos", "Genera una tabla de datos en formato CSV (separado por comas). Devuelve SOLO el código CSV.", "xw_tabla", "edit_xw_tabla", [("Ordena por primera columna", "Ordena por primera columna")], "📥 Descargar Excel")

    with tabs_xw[6]:
        _panel_borrador("Guion Podcast", "Guion podcast. 3 locutores. Formato exacto: 'Locutor 1: [texto]'.", "xw_podcast", "edit_xw_pod", [("Hazlo más gracioso", "Hazlo más gracioso")], None)
        if st.session_state.get("xw_podcast"):
            if st.button("🔊 Generar Audio", key="btn_xw_pod_aud"):
                with st.spinner("Generando audio..."):
                    aud = generar_podcast_3_voces(st.session_state.xw_podcast, "es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural")
                    if aud: st.session_state.xw_podcast_audio = aud; st.success("✅ Audio listo.")
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
        _panel_borrador("Línea de Tiempo", "Extrae los eventos cronológicos. Formato: Año/Fecha - Evento.", "xw_timeline", "edit_xw_time", [("Añade más fechas", "Añade más fechas")], "📥 Descargar")

    with tabs_xw[9]:
        _panel_borrador("Glosario", "Extrae los 5 conceptos más complejos y genera un glosario con definiciones sencillas.", "xw_glosario", "edit_xw_glos", [("Explícalo para un niño", "Explícalo para un niño")], "📥 Descargar")

    with tabs_xw[10]:
        _panel_borrador("Análisis DAFO", "Realiza un Análisis DAFO (Debilidades, Amenazas, Fortalezas, Oportunidades). Sé conciso.", "xw_dafo", "edit_xw_dafo", [("Hazlo más estratégico", "Hazlo más estratégico")], "📥 Descargar")

    with tabs_xw[11]:
        st.markdown("Genera una imagen para ilustrar tu proyecto.")
        prompt_img = st.text_input("Describe la imagen:", value=st.session_state.xw_tema_input, key="xw_img_prompt")
        if st.button("🖼️ Generar Portada", key="btn_xw_img", type="primary") and prompt_img:
            with st.spinner("Generando imagen..."):
                try:
                    url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img + ', cyberpunk, high quality')}?width=1024&height=1024&nologo=true"
                    if url_segura(url_img):
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
        st.text_area("Temática principal:", placeholder="Ej: El futuro de la IA, La pizza con piña...", key="pod_tema_input", height=100)
        c1, c2 = st.columns(2)
        with c1: estilo = st.selectbox("Estilo:", ["Entrevista", "Debate", "Educativo", "Narrativo"], key="p_est")
        with c2: dur = st.slider("Duración (min):", 1, 10, 3, key="p_dur")
        st.markdown("---")
        if st.button("💡 Generar 3 Ideas con IA", use_container_width=True, key="pod_btn_ideas"):
            with st.spinner("Pensando ideas..."):
                res = llamar_ia(modelo_seleccionado, f"Dame 3 ideas creativas y concisas para un podcast de {dur} min, estilo {estilo}, sobre: '{st.session_state.pod_tema_input}'. Solo ideas separadas por saltos de línea, sin numeración.", st.session_state.api_key_gemini_manual)
                if "[ERROR" not in res: st.session_state["pod_sugs"] = [l.strip("-*1234567890. ") for l in res.split('\n') if len(l.strip()) > 5][:3]
        def _add_idea_pod(idea):
            actual = st.session_state["pod_tema_input"].strip()
            st.session_state["pod_tema_input"] = f"{actual}\n\n{idea}" if actual else idea
            st.session_state["pod_sugs"] = []
        if st.session_state.get("pod_sugs"):
            st.markdown("✨ **Haz clic para añadir:**")
            cols = st.columns(len(st.session_state["pod_sugs"]))
            for i, idea in enumerate(st.session_state["pod_sugs"]):
                with cols[i]: st.button(idea, key=f"pod_idea_{i}", use_container_width=True, on_click=_add_idea_pod, args=(idea,))
        st.markdown("---")
        if st.button("✨ Crear Guion", type="primary", key="btn_pod") and st.session_state["pod_tema_input"]:
            with st.spinner("Escribiendo..."):
                prompt = f"{contexto}\n\nGuion podcast {dur} min. Tema: '{st.session_state['pod_tema_input']}'. Estilo: {estilo}. 3 locutores. REGLA ESTRICTA: natural y coloquial. Usa SOLO formato 'Locutor 1: [texto]'."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.podcast_script = res; guardar_en_db("Podcast", res)
        if st.session_state.get("podcast_script"):
            st.markdown("### 📜 Guion Generado")
            borrador_editable("podcast_script", "pod_edit", height=300)
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
                if "[ERROR" in res: st.error(res)
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
                    if "[ERROR" in res: st.error(res)
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
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state.agente_chat_history.append({"role": "assistant", "content": res}); st.rerun()

    elif sub_idx == 5:
        st.markdown("#### Debate")
        t = campo_con_ideas("Tema:", "deb", "temas de debate", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Coche eléctrico")
        if st.button("Iniciar") and t:
            res = llamar_ia(modelo_seleccionado, f"Debate {t}: 3 a favor, 3 en contra, conclusión.", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
            else: st.markdown(res); botones_descarga(res, "Debate")

    elif sub_idx == 6:
        st.markdown("#### Cocina")
        ing = campo_con_ideas("Ingredientes:", "chef1", "recetas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Pollo")
        pref = campo_con_ideas("Menú:", "chef2", "menús", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Vegetariano")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Receta") and ing:
                res = llamar_ia(modelo_seleccionado, f"Receta con {ing}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.info(res)
        with c2:
            if st.button("Menú") and pref:
                res = llamar_ia(modelo_seleccionado, f"Menú semanal {pref}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.info(res)

    elif sub_idx == 7:
        st.markdown("#### Cuentos")
        p = campo_con_ideas("Protagonista:", "c1", "protagonistas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Dragón")
        l = campo_con_ideas("Lugar:", "c2", "lugares mágicos", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Bosque")
        if st.button("Crear") and p and l:
            res = llamar_ia(modelo_seleccionado, f"Cuento sobre {p} en {l}.", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
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
            if "[ERROR" in res: st.error(res)
            else: st.code(res, language="python")

    elif sub_idx == 9:
        st.markdown("#### Traductor")
        txt = campo_con_ideas("Texto:", "trad", "frases", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Texto...")
        c1, c2 = st.columns(2)
        with c1: o = st.selectbox("Origen:", ["Auto", "Español", "Inglés"], key="o")
        with c2: d = st.selectbox("Destino:", ["Español", "Inglés", "Francés"], key="d")
        if st.button("Traducir") and txt:
            res = llamar_ia(modelo_seleccionado, f"Traduce del {o} al {d}:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
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
                prompt = {"Resumen Ejecutivo": f"Resumen:\n{texto}", "Extraer Fechas y Cifras": f"Extrae fechas:\n{texto}", "Buscar Errores": f"Errores:\n{texto}", "Crear Test": f"Test 5 preguntas:\n{texto}", "Traducir al Inglés": f"Traduce al inglés:\n{texto}"}[opc]
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resultado:", res, height=400); botones_descarga(res, "Analisis")

    elif sub_idx == 1:
        st.markdown("#### Comparador")
        a = st.file_uploader("Original:", type=["pdf", "txt"]); b = st.file_uploader("Modificado:", type=["pdf", "txt"])
        if a and b and st.button("Comparar"):
            with st.spinner("Buscando..."):
                ta = leer_archivo_universal(a, modelo_seleccionado); tb = leer_archivo_universal(b, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"Diferencias:\nA:\n{ta[:10000]}\nB:\n{tb[:10000]}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
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
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state.office_borrador = res
            else: st.warning("Escribe un prompt primero.")

        if st.session_state.get("office_borrador"):
            st.markdown("### 📝 Borrador Editable")
            borrador_editable("office_borrador", "office_txt", height=350)
            c1, c2, c3 = st.columns(3)
            with c1: boton_superpoder_borrador(st.session_state.office_txt, "Hazlo más profesional", "of_sup1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            with c2: boton_superpoder_borrador(st.session_state.office_txt, "Resúmelo", "of_sup2", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            with c3: boton_superpoder_borrador(st.session_state.office_txt, "Añade más detalles", "of_sup3", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            st.markdown("---")
            if formato == "Word": st.download_button("📥 Descargar Word", generar_word(st.session_state.office_txt), "doc.docx")
            elif formato == "Excel": st.download_button("📥 Descargar Excel", generar_excel(st.session_state.office_txt), "doc.xlsx")
            else: st.download_button("📥 Descargar PPTX", generar_pptx("Presentación", st.session_state.office_txt), "doc.pptx")

    elif sub_idx == 3:
        st.markdown("#### PDF Pro (Flujo Borrador)")
        contexto = cajita_adjuntar_archivo("pdf_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        titulo = st.text_input("Título del PDF:")
        contenido = campo_con_ideas("Contenido:", "pdf_c", "contenido PDF", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Escribe el contenido...")
        if st.button("✨ Generar Borrador PDF", type="primary", key="btn_pdf_gen") and titulo and contenido:
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\n{contenido}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.pdf_borrador = res
        if st.session_state.get("pdf_borrador"):
            st.markdown("### 📝 Borrador Editable")
            borrador_editable("pdf_borrador", "pdf_txt", height=350)
            c1, c2 = st.columns(2)
            with c1: boton_superpoder_borrador(st.session_state.pdf_txt, "Hazlo más técnico", "pdf_sup1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "pdf_borrador", widget_key="pdf_txt")
            with c2: boton_superpoder_borrador(st.session_state.pdf_txt, "Añade conclusiones", "pdf_sup2", modelo_seleccionado, st.session_state.api_key_gemini_manual, "pdf_borrador", widget_key="pdf_txt")
            st.markdown("---")
            st.download_button("📥 Descargar PDF (con portada)", generar_pdf_pro(titulo, st.session_state.pdf_txt), f"{titulo}.pdf")

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
        inst = campo_con_ideas("¿Qué hago?", "clip", "instrucciones", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Resume")
        if st.button("Procesar") and txt and inst:
            res = llamar_ia(modelo_seleccionado, f"{inst}:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
            else: st.info(res)


# ==========================================
# CATEGORÍA 3: WEB Y ACTUALIDAD
# ==========================================
elif macro_idx == 3:
    sub_names = ["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV", "📊 WEB A EXCEL", "🎥 YOUTUBE", "📰 LECTOR"]
    sub_idx = custom_tabs(sub_names, "web_sub", 0)
    if sub_idx == 0:
        q = campo_con_ideas("Buscar:", "web", "búsquedas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Noticias IA")
        if st.button("🔍 Buscar") and q:
            for r in buscar_en_internet(q): st.markdown(f"**[{r['titulo']}]({r['url']})**"); st.caption(r['resumen'])
    elif sub_idx == 1:
        d = st.selectbox("Deporte:", ["Fútbol", "Baloncesto", "Tenis"])
        if st.button("Noticias"):
            for n in buscar_noticias_deportivas(d, "Todas"): st.markdown(f"**{n['titulo']}** - {n['resumen']}")
    elif sub_idx == 2:
        c = campo_con_ideas("Canal:", "tv", "TV", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Telecinco")
        if st.button("Programación") and c:
            res = buscar_en_internet(c)
            r_ia = llamar_ia(modelo_seleccionado, f"Programación {c}: {res}", st.session_state.api_key_gemini_manual)
            if "[ERROR" in r_ia: st.error(r_ia)
            else: st.info(r_ia)
    elif sub_idx == 3:
        url = st.text_input("URL Tabla:")
        if st.button("Extraer") and url:
            try:
                if url_segura(url):
                    df = pd.read_html(url)[0]; st.dataframe(df)
                    out = BytesIO(); df.to_excel(out); st.download_button("📥 Excel", out.getvalue(), "tabla.xlsx")
                else: st.error("URL bloqueada.")
            except: st.error("Fallo al leer.")

    elif sub_idx == 4:
        st.markdown("#### 🎥 YouTube / Audio")
        fuente_tipo = st.radio("Elige el método:", ["Subir archivo (Recomendado)", "Pegar enlace de YouTube"], horizontal=True)
        archivo_subido = None; url = ""
        if fuente_tipo == "Subir archivo (Recomendado)": archivo_subido = st.file_uploader("Sube un audio (.mp3) o vídeo (.mp4):", type=["mp3", "wav", "m4a", "mp4", "mov"])
        else: url = st.text_input("URL del vídeo:", placeholder="https://www.youtube.com/watch?v=...")
        modo = st.radio("¿Qué quieres hacer?", ["Resumir", "Traducir al español"], horizontal=True)

        if st.button("Ejecutar", type="primary") and (archivo_subido or url):
            texto_original = ""; fuente = ""
            vid = extraer_youtube_id(url) if url else "archivo_local"
            key_groq = _get_key("groq")
            if archivo_subido:
                with st.spinner("Transcribiendo archivo con Whisper..."):
                    texto_original = transcribir_archivo_local(archivo_subido.getvalue(), archivo_subido.name, key_groq)
                fuente = "archivo subido"
            else:
                if not vid:
                    st.error("No he podido extraer el ID del vídeo. Revisa el enlace."); st.stop()
                try:
                    with st.spinner("Buscando subtítulos..."):
                        tr = YouTubeTranscriptApi.get_transcript(vid, languages=['es', 'en'])
                    texto_original = " ".join([t['text'] for t in tr]); fuente = "subtítulos de YouTube"
                except Exception:
                    st.info("Este vídeo no tiene subtítulos. Transcribiendo el audio con Whisper...")
                    with st.spinner("Descargando audio y transcribiendo..."):
                        texto_original = transcribir_audio_youtube(url, key_groq)
                    fuente = "whisper"

            if texto_original.startswith("[ERROR]"): st.error(texto_original)
            else:
                st.caption(f"Fuente: {fuente} · {len(texto_original)} caracteres")
                with st.expander("Ver transcripción original"): st.text_area("Texto detectado:", value=texto_original, height=150, key=f"yt_orig_{vid}")
                if modo == "Resumir": prompt = f"Resume este vídeo de YouTube en español, con tono claro y directo:\n{texto_original[:8000]}"
                else: prompt = f"Traduce al español el siguiente texto manteniendo el tono natural del original, sin añadir ni quitar contenido. Si el texto ya está en español, devuélvelo tal cual. Devuelve solo la traducción:\n{texto_original[:8000]}"
                with st.spinner("Procesando con IA..."):
                    res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resultado:", res, height=300); botones_descarga(res, f"YT_{vid}")

    elif sub_idx == 5:
        url = st.text_input("URL Web:")
        if st.button("Leer") and url:
            try:
                if not url_segura(url): st.error("URL no permitida."); st.stop()
                txt = " ".join(re.sub('<[^<]+?>', ' ', fetch_seguro(url).text).split())[:5000]
                res = llamar_ia(modelo_seleccionado, f"Resume: {txt}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resumen:", res, height=200)
            except Exception as e: st.error(f"Error: {e}")


# ==========================================
# CATEGORÍA 4: MULTIMEDIA
# ==========================================
elif macro_idx == 4:
    st.markdown("#### 🎬 Multimedia")
    sub_names = ["🖼️ IMAGEN", "🎙️ TEXTO A VOZ", "🎬 VIDEO (Pronto)"]
    sub_idx = custom_tabs(sub_names, "multi_sub", 0)

    if sub_idx == 0:
        st.markdown("### Generador de Imágenes")
        prompt_img = st.text_input("Describe la imagen que quieres crear:", key="img_gen_prompt")
        estilo_img = st.selectbox("Estilo:", ["Cyberpunk", "Realista", "Anime", "3D Render", "Fotografía"], key="img_style")
        if st.button("🖼️ Generar Imagen", type="primary") and prompt_img:
            with st.spinner("Generando imagen..."):
                try:
                    query = urllib.parse.quote(f"{prompt_img}, {estilo_img}, high quality, 8k")
                    url = f"https://image.pollinations.ai/prompt/{query}?width=1024&height=1024&nologo=true"
                    if url_segura(url):
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=30) as response:
                            img_data = response.read()
                        st.image(img_data, caption=prompt_img, use_container_width=True)
                        st.download_button("📥 Descargar Imagen", img_data, "imagen_generada.png", "image/png")
                except Exception as e:
                    st.error(f"Error al generar la imagen: {e}")

    elif sub_idx == 1:
        st.markdown("### Texto a Voz (TTS)")
        txt_tts = st.text_area("Escribe el texto que quieres escuchar:", height=150)
        voz_tts = st.selectbox("Selecciona la voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"])
        if st.button("🔊 Generar Audio", type="primary") and txt_tts:
            with st.spinner("Sintetizando voz..."):
                aud = generar_audio_sincrono(txt_tts, voz_tts)
                if aud:
                    st.audio(aud, format="audio/mp3")
                    st.download_button("💾 Descargar MP3", aud, "audio.mp3")


# ==========================================
# CATEGORÍA 5: ESTUDIO Y OFICIOS
# ==========================================
elif macro_idx == 5:
    st.markdown("#### 🎓 Estudio y Oficios")
    sub_names = ["📚 DEBERES", "👔 ENTREVISTA", "🛠️ PRL (Riesgos)"]
    sub_idx = custom_tabs(sub_names, "study_sub", 0)

    if sub_idx == 0:
        st.markdown("### Resolver Deberes")
        materia = st.selectbox("Materia:", ["Matemáticas", "Historia", "Ciencias", "Lengua", "Física y Química", "Inglés"])
        pregunta = st.text_area("Escribe tu duda o pega el enunciado:", height=150)
        if st.button("🔎 Resolver paso a paso", type="primary") and pregunta:
            with st.spinner("Pensando..."):
                prompt = f"Actúa como un profesor experto en {materia}. Resuelve el siguiente problema paso a paso, explicando cada parte de forma sencilla para que lo entienda:\n\n{pregunta}"
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.markdown(res)
                    botones_descarga(res, "Deberes_Resueltos", formatos_extra=True)

    elif sub_idx == 1:
        st.markdown("### Simulador de Entrevista de Trabajo")
        puesto = st.text_input("¿A qué puesto aspiras?", placeholder="Ej: Camarero, Programador, Administrativo...")
        if st.button("🤖 Empezar Entrevista", type="primary") and puesto:
            st.session_state.entrevista_puesto = puesto
            st.session_state.entrevista_historial = [{"role": "user", "content": f"Actúa como un reclutador de RRHH para el puesto de {puesto}. Preséntate y hazme la primera pregunta."}]
            with st.spinner("El reclutador está pensando..."):
                res = llamar_ia(modelo_seleccionado, st.session_state.entrevista_historial[0]["content"], st.session_state.api_key_gemini_manual)
                st.session_state.entrevista_historial.append({"role": "assistant", "content": res})
                st.rerun()

        if "entrevista_historial" in st.session_state and st.session_state.entrevista_historial:
            for msg in st.session_state.entrevista_historial:
                if msg["role"] == "user" and "Actúa como" not in msg["content"]:
                    st.chat_message("user").write(msg["content"])
                elif msg["role"] == "assistant":
                    st.chat_message("assistant").write(msg["content"])
            
            resp = chat_input_pro("entrevista_input", "Escribe tu respuesta...")
            if resp:
                st.session_state.entrevista_historial.append({"role": "user", "content": resp})
                with st.spinner("Evaluando tu respuesta..."):
                    contexto = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.entrevista_historial])
                    prompt = f"Continúa la entrevista de trabajo para {st.session_state.entrevista_puesto}. Evalúa brevemente mi última respuesta y hazme la siguiente pregunta. Sé profesional.\n\nHistorial:\n{contexto}"
                    res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                    st.session_state.entrevista_historial.append({"role": "assistant", "content": res})
                    st.rerun()

    elif sub_idx == 2:
        st.markdown("### Prevención de Riesgos Laborales (PRL)")
        tarea = st.text_area("Describe la tarea o el entorno de trabajo:", placeholder="Ej: Trabajo en altura en una obra, o Manejo de productos químicos en laboratorio...")
        if st.button("🦺 Analizar Riesgos", type="primary") and tarea:
            with st.spinner("Analizando riesgos..."):
                prompt = f"Actúa como un técnico de Prevención de Riesgos Laborales. Analiza la siguiente situación y dime: 1. Riesgos principales, 2. Medidas preventivas, 3. Equipos de Protección Individual (EPIs) necesarios:\n\n{tarea}"
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.markdown(res)
                    botones_descarga(res, "Informe_PRL", formatos_extra=True)


# ==========================================
# CATEGORÍA 6: VIAJES Y HOGAR
# ==========================================
elif macro_idx == 6:
    st.markdown("#### 🌍 Viajes y Hogar")
    sub_names = ["✈️ ITINERARIO", "🏠 COCINA"]
    sub_idx = custom_tabs(sub_names, "travel_sub", 0)

    if sub_idx == 0:
        st.markdown("### Planificador de Viajes")
        destino = st.text_input("¿A dónde quieres ir?", placeholder="Ej: Roma, Tokio, Nueva York...")
        c1, c2 = st.columns(2)
        with c1: dias = st.number_input("¿Cuántos días?", min_value=1, max_value=30, value=3)
        with c2: tipo_viaje = st.selectbox("Estilo de viaje:", ["Cultural", "Gastronómico", "Aventura", "Relax", "Familiar", "Económico"])
        
        if st.button("🗺️ Crear Itinerario", type="primary") and destino:
            with st.spinner("Planificando tu viaje..."):
                prompt = f"Crea un itinerario detallado de {dias} días para viajar a {destino}. El estilo del viaje es {tipo_viaje}. Incluye qué ver cada día (mañana/tarde/noche), dónde comer y consejos prácticos. Usa un formato estructurado con títulos."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.markdown(res)
                    botones_descarga(res, f"Viaje_{destino}", formatos_extra=True)

    elif sub_idx == 1:
        st.markdown("### Recetas con lo que tengo en casa")
        ingredientes = campo_con_ideas("Ingredientes en la nevera:", "hogar", "recetas de cocina", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Huevos, patatas, cebolla...")
        if st.button("🍳 Crear Receta", type="primary") and ingredientes:
            with st.spinner("Cocinando ideas..."):
                prompt = f"Tengo estos ingredientes: {ingredientes}. Dame 2 recetas fáciles y rápidas de preparar. Incluye los pasos y el tiempo de preparación."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.markdown(res)
                    botones_descarga(res, "Recetas_Hogar")


# ==========================================
# CATEGORÍA 7: MÚSICA Y ARTE
# ==========================================
elif macro_idx == 7:
    st.markdown("### 🎵 MÚSICA Y ARTE")
    st.info("Módulo de letras, canciones y portadas listo para expandir.")
