"""
Xavilu Network Pro - Edición Logo Opcional
"""
import streamlit as st
import streamlit.components.v1 as components
from streamlit_markmap import markmap
import ollama, asyncio, edge_tts, pypdf, openpyxl, csv, os, re, requests, tempfile, time, qrcode, numpy as np, sqlite3, datetime, concurrent.futures
from docx import Document
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
from PIL import Image
from fpdf import FPDF
import urllib.parse, urllib.request, ipaddress, socket
import pandas as pd
from youtube_transcript_api import YouTubeTranscriptApi
from pydub import AudioSegment

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

st.set_page_config(page_title="Xavilu Network", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")
MAX_FILE_MB = 100
MAX_TEXTO_IA = 60_000
GEMINI_MODEL = "gemini-2.0-flash"

if os.name == "nt": asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

def init_db():
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        conn.execute("CREATE TABLE IF NOT EXISTS historial (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, tipo TEXT, contenido TEXT)")
def guardar_en_db(tipo, contenido):
    try:
        with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
            conn.execute("INSERT INTO historial (fecha, tipo, contenido) VALUES (?, ?, ?)", (datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S"), tipo, contenido))
    except: pass
def obtener_de_db():
    try:
        with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
            return conn.execute("SELECT fecha, tipo, contenido FROM historial ORDER BY id DESC LIMIT 5").fetchall()
    except: return []
init_db()

_defaults = {
    "chat_history": [], "pdf_chat_history": [], "agente_activo": None, "agente_chat_history": [],
    "sidebar_visible": True, "whatsapp_number": "", "api_key_gemini_manual": "", "api_key_groq_manual": "",
    "api_key_silicon_manual": "", "api_key_mistral_manual": "", "api_key_openrouter_manual": "",
    "api_key_deepseek_manual": "", "api_key_together_manual": "", "api_key_cerebras_manual": "",
    "xw_mapa": "", "xw_conceptual": "", "xw_ppt": "", "xw_informe": "", "xw_tarjetas": "", "xw_cuestionario": "",
    "xw_tabla": "", "xw_podcast": "", "xw_podcast_audio": None, "xw_timeline": "", "xw_glosario": "",
    "xw_dafo": "", "xw_cover_img": None, "xw_copilot_history": [], "xw_tema_input": "",
    "office_borrador": "", "pdf_borrador": "", "pod_tema_input": "", "podcast_script": "", "entrevista_historial": [],
    "alex_chat_history": [], "letra_generada": ""
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
.stApp { background-color: #000000 !important; background-image: radial-gradient(circle at 50% 50%, #111111 0%, #000000 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
header[data-testid="stHeader"] { background-color: #000000 !important; }
[data-testid="stAppDeployButton"], [data-testid="stDecoration"] { display: none !important; }
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700, 0 0 35px #FF8C00 !important; text-align: center; letter-spacing: 3px; margin-top: 0px !important; padding-top: 0px !important; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #ffffff !important; text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff, 0 0 30px #FF007F !important; text-align: center; letter-spacing: 3px; }
h4 { font-family: 'Orbitron', sans-serif !important; color: #00f0ff !important; text-shadow: 0 0 8px #00f0ff, 0 0 15px #FF007F !important; letter-spacing: 1px; }
[data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; text-transform: uppercase !important; background: rgba(0, 240, 255, 0.1) !important; padding: 5px 10px !important; border-left: 4px solid #FF007F !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 8px !important; }
[data-testid="stSidebar"] { background-color: rgba(5, 5, 10, 0.98) !important; border-right: 2px solid #00f0ff !important; box-shadow: 5px 0 15px rgba(0, 240, 255, 0.2) !important; }
[data-testid="stRadio"] > div[role="radiogroup"] label div p, [data-testid="stRadio"] > div[role="radiogroup"] label div span { color: #ffffff !important; font-size: 15px !important; font-weight: 600 !important; }
[data-testid="stFileUploader"], [data-testid="stExpander"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; }
[data-testid="stFileUploaderDropzone"] { background-color: #0a0a0a !important; color: #ffffff !important; }
[data-testid="stFileUploaderDropzone"] p, [data-testid="stExpander"] summary { color: #00f0ff !important; font-weight: bold !important; }
div.stButton > button[kind="secondary"] { background-color: #0a0a0a !important; color: #ffffff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; padding: 10px 15px !important; margin-bottom: 5px !important; transition: all 0.3s ease !important; box-shadow: 0 0 10px rgba(0, 240, 255, 0.2) !important; width: 100% !important; }
div.stButton > button[kind="secondary"]:hover { background-color: rgba(0, 240, 255, 0.15) !important; color: #00f0ff !important; box-shadow: 0 0 20px rgba(0, 240, 255, 0.7) !important; transform: translateY(-2px) !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000000 !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: 900 !important; border: 2px solid #ffffff !important; border-radius: 25px !important; padding: 10px 15px !important; margin-bottom: 5px !important; transition: all 0.3s ease !important; box-shadow: 0 0 18px rgba(0, 240, 255, 0.9), 0 0 32px rgba(255, 0, 127, 0.7) !important; transform: translateY(-1px) !important; width: 100% !important; }
.stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"], .stNumberInput input { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; color: #ffffff !important; font-size: 16px !important; box-shadow: inset 0 0 10px rgba(0, 240, 255, 0.1) !important; }
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
.block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
@media (max-width: 600px) { .block-container { padding-top: 0.5rem !important; } h1 { font-size: 2.2rem !important; letter-spacing: 1px !important; margin-top: 0px !important; } [data-testid="stWidgetLabel"] p { font-size: 11px !important; padding: 3px 8px !important; } div.stButton > button { padding: 6px 10px !important; font-size: 12px !important; margin-bottom: 3px !important; } .stTextArea textarea { font-size: 14px !important; } }
</style>
""", unsafe_allow_html=True)

def inyectar_js_fixer():
    components.html("<script>(function(){var doc=window.parent.document;function f(){doc.querySelectorAll('[data-testid=\"stTextInput\"] *').forEach(function(el){if(el.tagName.toLowerCase()==='input'||el.tagName.toLowerCase()==='div'){el.style.setProperty('background-color','#0a0a0a','important');el.style.setProperty('color','#ffffff','important');}});}f();setInterval(f,300);})();</script>", height=0)
inyectar_js_fixer()

def url_segura(url: str) -> bool:
    try:
        p = urllib.parse.urlparse(url)
        if p.scheme not in ("http", "https") or not p.hostname: return False
        host = socket.gethostbyname(p.hostname)
        ip = ipaddress.ip_address(host)
        return not (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast)
    except: return False
def fetch_seguro(url: str, timeout: int = 15, headers: dict = None):
    if not url_segura(url): raise ValueError(f"URL bloqueada: {url}")
    h = {"User-Agent": "Mozilla/5.0 (compatible; Xavilu)"}
    if headers: h.update(headers)
    return requests.get(url, headers=h, timeout=timeout)
def leer_archivo_universal(upload, modelo_ia: str) -> str:
    if upload.size > MAX_FILE_MB * 1024 * 1024: return ""
    nombre = upload.name.lower()
    try:
        if nombre.endswith(".txt"): return upload.read().decode("utf-8", errors="ignore")
        if nombre.endswith(".pdf"): return "\n".join(p.extract_text() for p in pypdf.PdfReader(BytesIO(upload.getvalue())).pages if p.extract_text())
        if nombre.endswith(".docx"): return "\n".join(p.text for p in Document(upload).paragraphs if p.text)
        if nombre.endswith(".xlsx"): return "\n\n".join(f"=== {s} ===\n{d.to_csv(index=False)}" for s, d in pd.read_excel(upload, sheet_name=None).items())
    except Exception as e: st.warning(f"No pude leer '{upload.name}': {e}"); return ""
    return "Formato no reconocido."
@st.cache_resource(show_spinner=False)
def get_gemini_client(api_key: str): return genai.Client(api_key=api_key)
def buscar_en_internet(query: str):
    if DDGS is None: return []
    try: return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in DDGS().text(query, region="es-es", max_results=5)]
    except: return []
def buscar_noticias_deportivas(deporte: str, liga: str):
    if DDGS is None: return []
    try: return [{"titulo": r.get("title", ""), "url": r.get("url", ""), "resumen": r.get("body", ""), "fecha": r.get("date", "")} for r in DDGS().news(f"últimas noticias {deporte} {liga} hoy", max_results=5)]
    except: return []
def extraer_youtube_id(url: str):
    if not url: return None
    m = re.search(r"(?:v=|/v/|youtu\.be/|/embed/|/shorts/)([A-Za-z0-9_-]{11})", url)
    return m.group(1) if m else (url.strip() if len(url.strip()) == 11 else None)
def transcribir_audio_youtube(url: str, api_key_groq: str) -> str:
    if not HAS_YTDLP or not HAS_OPENAI or not api_key_groq: return "[ERROR] Faltan herramientas."
    tmp_dir = tempfile.mkdtemp(); audio_path = None
    try:
        ydl_opts = {"format": "bestaudio[abr<=64]/worstaudio", "outtmpl": os.path.join(tmp_dir, "audio.%(ext)s"), "quiet": True, "noplaylist": True, "postprocessors": [{"key": "FFmpegExtractAudio", "preferredcodec": "mp3", "preferredquality": "48"}]}
        try:
            ydl_opts["impersonate"] = "chrome"
            with yt_dlp.YoutubeDL(ydl_opts) as ydl: info = ydl.extract_info(url, download=True)
        except Exception:
            ydl_opts.pop("impersonate", None)
            with yt_dlp.YoutubeDL(ydl_opts) as ydl: info = ydl.extract_info(url, download=True)
        audio_path = ydl.prepare_filename(info)
        if not audio_path.endswith(".mp3"):
            base = os.path.splitext(audio_path)[0]
            if os.path.exists(base + ".mp3"): audio_path = base + ".mp3"
        size_mb = os.path.getsize(audio_path) / (1024 * 1024)
        if size_mb > 24: return f"[ERROR] Audio de {size_mb:.1f} MB, supera 25 MB."
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f: resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e: return f"[ERROR] {e}"
    finally:
        if audio_path and os.path.exists(audio_path):
            try: os.remove(audio_path)
            except: pass
        try: os.rmdir(tmp_dir)
        except: pass
def transcribir_archivo_local(archivo_bytes: bytes, nombre_archivo: str, api_key_groq: str) -> str:
    if not HAS_OPENAI or not api_key_groq: return "[ERROR] Faltan herramientas."
    ext = os.path.splitext(nombre_archivo)[1]
    with tempfile.NamedTemporaryFile(delete=False, suffix=ext) as tmp: tmp.write(archivo_bytes); audio_path = tmp.name
    try:
        client = openai.OpenAI(api_key=api_key_groq, base_url="https://api.groq.com/openai/v1")
        with open(audio_path, "rb") as f: resp = client.audio.transcriptions.create(model="whisper-large-v3-turbo", file=f, response_format="text")
        return resp if isinstance(resp, str) else resp.text
    except Exception as e: return f"[ERROR] {e}"
    finally:
        if os.path.exists(audio_path):
            try: os.remove(audio_path)
            except: pass
def cajita_adjuntar_archivo(key_suffix: str, modelo: str, api_key: str) -> str:
    with st.expander("📎 Adjuntar archivo para esta acción (Opcional)"):
        archivo = st.file_uploader("Sube PDF, Word, TXT o Imagen", type=["pdf", "txt", "docx", "png", "jpg", "jpeg"], key=f"adj_{key_suffix}")
        if not archivo: return ""
        nombre = archivo.name.lower()
        try:
            if nombre.endswith((".png", ".jpg", ".jpeg")):
                with st.spinner("Analizando imagen..."): ctx = analizar_archivo_con_gemini(archivo.getvalue(), "image/jpeg", "Describe esta imagen o extrae el texto (OCR).", api_key)
                st.success("✅ Imagen leída."); return f"\n[TEXTO DE IMAGEN]:\n{ctx}\n"
            texto = leer_archivo_universal(archivo, modelo)
            st.success(f"✅ Archivo '{archivo.name}' leído."); return f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"
        except Exception as e: st.warning(f"No pude procesar '{archivo.name}': {e}"); return ""
def analizar_archivo_con_gemini(archivo_bytes: bytes, mime_type: str, prompt: str, api_key_gemini: str) -> str:
    if not HAS_GEMINI or not api_key_gemini: return "[ERROR] Falta google-genai o API key."
    size_mb = len(archivo_bytes) / (1024 * 1024); tmp_path = None
    try:
        client = get_gemini_client(api_key_gemini)
        if size_mb > 15:
            ext = mime_type.split("/")[-1]
            with tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}") as tmp: tmp.write(archivo_bytes); tmp_path = tmp.name
            uploaded = client.files.upload(file=tmp_path)
            while uploaded.state.name == "PROCESSING": time.sleep(3); uploaded = client.files.get(name=uploaded.name)
            if uploaded.state.name == "FAILED": return "[ERROR] Gemini no pudo procesar el archivo."
            res = client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, uploaded]).text
            try: client.files.delete(name=uploaded.name)
            except: pass
            return res
        parte = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
        return client.models.generate_content(model=GEMINI_MODEL, contents=[prompt, parte]).text
    except Exception as e: return f"[ERROR GEMINI] {e}"
    finally:
        if tmp_path and os.path.exists(tmp_path):
            try: os.remove(tmp_path)
            except: pass
def _get_key(nombre: str, arg_directo: str = "") -> str:
    manual = st.session_state.get(f"api_key_{nombre.lower()}_manual", "").strip()
    if manual: return manual
    try:
        secret = st.secrets.get(f"{nombre.upper()}_API_KEY", "")
        if secret: return secret
    except: pass
    return arg_directo or ""
def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None, api_key_openai="", api_key_anthropic=""):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if st.session_state.get("global_context"): prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}\n\n(Responde teniendo en cuenta el archivo)."
    if historial_chat:
        contexto_chat = "\n".join(f"{msg['role'].upper()}: {msg['content']}" for msg in historial_chat[-10:])
        prompt_final = f"Historial:\n{contexto_chat}\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."
    key_gemini, key_groq, key_silicon, key_mistral, key_openrouter = _get_key("gemini", api_key_gemini), _get_key("groq"), _get_key("silicon"), _get_key("mistral"), _get_key("openrouter")
    key_deepseek, key_together, key_cerebras = _get_key("deepseek"), _get_key("together"), _get_key("cerebras")
    if key_gemini and HAS_GEMINI:
        for intento in range(2):
            try:
                res = get_gemini_client(key_gemini).models.generate_content(model=GEMINI_MODEL, contents=prompt_final).text
                if res: return res
            except Exception as e:
                msg = str(e)
                if "429" in msg or "RESOURCE_EXHAUSTED" in msg: break
                if "503" in msg and intento < 1: time.sleep(2); continue
                break
    if key_groq and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_groq, base_url="https://api.groq.com/openai/v1")
            res = client.chat.completions.create(model="llama-3.3-70b-versatile", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_silicon and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_silicon, base_url="https://api.siliconflow.cn/v1")
            res = client.chat.completions.create(model="Qwen/Qwen2.5-7B-Instruct", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_mistral and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_mistral, base_url="https://api.mistral.ai/v1")
            res = client.chat.completions.create(model="mistral-small-latest", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_openrouter and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_openrouter, base_url="https://openrouter.ai/api/v1")
            res = client.chat.completions.create(model="deepseek/deepseek-r1:free", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_deepseek and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_deepseek, base_url="https://api.deepseek.com/v1")
            res = client.chat.completions.create(model="deepseek-chat", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_together and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_together, base_url="https://api.together.xyz/v1")
            res = client.chat.completions.create(model="meta-llama/Llama-3-70b-chat-hf", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    if key_cerebras and HAS_OPENAI:
        try:
            client = openai.OpenAI(api_key=key_cerebras, base_url="https://api.cerebras.ai/v1")
            res = client.chat.completions.create(model="llama3.1-70b", messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
            if res: return res
        except: pass
    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=600)
        raw = client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})["message"]["content"]
        raw = re.sub(r"^```\w*\n", "", raw); raw = re.sub(r"\n```\s*$", "", raw)
        return raw.strip()
    except Exception as e: return f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}"

def render_mermaid(mermaid_code):
    html_code = f"<div class='mermaid'>{mermaid_code}</div><script type='module'>import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';mermaid.initialize({{ startOnLoad: true, theme: 'dark', securityLevel: 'loose' }});</script>"
    components.html(html_code, height=600, scrolling=True)
def obtener_ruta_fuente():
    for ruta in ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/arial.ttf", "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf", os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf") if "__file__" in globals() else ""]:
        if ruta and os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo: str, contenido: str, logo_bytes: bytes = None) -> bytes:
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente()
    
    try:
        url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote('Book cover for ' + titulo + ', cyberpunk')}?width=512&height=512&nologo=true"
        with urllib.request.urlopen(urllib.request.Request(url_img, headers={"User-Agent": "Mozilla/5.0"}), timeout=15) as response: img_data = response.read()
        with open("temp_cover.jpg", "wb") as f: f.write(img_data)
        pdf.image("temp_cover.jpg", x=40, y=50, w=130); pdf.ln(150); os.remove("temp_cover.jpg")
    except: pass
    
    if logo_bytes:
        try:
            with open("temp_logo.jpg", "wb") as f: f.write(logo_bytes)
            pdf.image("temp_logo.jpg", x=160, y=10, w=30)
            os.remove("temp_logo.jpg")
        except: pass

    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.add_font("XaviFont", "I", font_path, uni=True); pdf.set_font("XaviFont", size=11)
    else: pdf.set_font("Helvetica", size=11)
    if font_path: pdf.set_font("XaviFont", "B", 24)
    else: pdf.set_font("Helvetica", "B", 24)
    pdf.multi_cell(0, 10, text=titulo, align="C")
    if font_path: pdf.set_font("XaviFont", "I", 12)
    else: pdf.set_font("Helvetica", "I", 12)
    pdf.cell(0, 10, text="Generado por Xavilu Network", align="C", ln=True); pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)
    if font_path: pdf.set_font("XaviFont", size=11)
    else: contenido = contenido.encode("latin-1", "replace").decode("latin-1"); pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 8, text=contenido); return bytes(pdf.output())

def generar_manual_pdf() -> bytes:
    texto_manual = "¡Bienvenido a Xavilu Network Pro!\n\nEsta app conecta con Google, Meta, y modelos abiertos. Necesitas API Keys.\n\n1. XaviWork: Mapas, informes, presentaciones.\n2. Chat: Agentes y podcasts.\n3. Documentos: Resúmenes y facturas.\n4. Web: Búsquedas y YouTube.\n5. Multimedia: Imágenes y audios.\n6. Estudio: Deberes, Entrevistas y PRL.\n7. Viajes: Itinerarios y Rutinas.\n8. Música: Letras, Poesía y Canciones.\n9. Comunidad: Gestor de fincas PRO.\n\n¡Disfruta!"
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente()
    if font_path: pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.set_font("XaviFont", "B", 20)
    else: pdf.set_font("Helvetica", "B", 20)
    pdf.multi_cell(0, 10, text="Manual de Instrucciones", align="C"); pdf.ln(10)
    if font_path: pdf.set_font("XaviFont", size=12)
    else: pdf.set_font("Helvetica", size=12); texto_manual = texto_manual.encode('latin-1', 'replace').decode('latin-1')
    pdf.multi_cell(0, 7, text=texto_manual); return bytes(pdf.output())
def generar_word(texto: str) -> bytes:
    doc = Document(); doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); return out.getvalue()
def generar_excel(texto_csv: str) -> bytes:
    wb = openpyxl.Workbook(); ws = wb.active; csv_limpio = texto_csv.replace("```csv", "").replace("```", "").strip()
    for r_i, row in enumerate(csv.reader(csv_limpio.split("\n")), 1):
        for c_i, val in enumerate(row, 1):
            cell = ws.cell(row=r_i, column=c_i, value=val)
            if r_i == 1: cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); return out.getvalue()
def generar_pptx(titulo: str, texto: str) -> bytes:
    prs = Presentation(); slide = prs.slides.add_slide(prs.slide_layouts[0]); slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1]); lineas = bloque.split("\n")
        if slide.shapes.title: slide.shapes.title.text = lineas[0][:50]
        contenido = "\n".join(lineas[1:]) if len(lineas) > 1 else ""; tiene_ph = False
        for shape in slide.placeholders:
            if shape.placeholder_format.idx == 1: shape.text = contenido; tiene_ph = True
        if not tiene_ph: tx = slide.shapes.add_textbox(Inches(1), Inches(2), Inches(8), Inches(5)); tx.text_frame.text = contenido
    out = BytesIO(); prs.save(out); return out.getvalue()

def botones_descarga(texto: str, nombre_por_defecto: str = "documento", formatos_extra: bool = False):
    if not texto: return
    st.markdown("---"); key_u = nombre_por_defecto.replace(" ", "_")
    nombre_elegido = st.text_input("📝 Nombre del archivo (sin extensión):", value=nombre_por_defecto, key=f"name_{key_u}")
    nombre_limpio = re.sub(r"[^a-zA-Z0-9_\- ]", "", nombre_elegido).strip().replace(" ", "_") or nombre_por_defecto
    
    logo_bytes = None
    with st.expander("📎 Adjuntar Logo a este PDF (Opcional)"):
        logo_file = st.file_uploader("Sube una imagen (PNG/JPG)", type=["png", "jpg", "jpeg"], key=f"logo_{key_u}", label_visibility="collapsed")
        if logo_file: logo_bytes = logo_file.getvalue()

    if formatos_extra:
        c1, c2, c3, c4 = st.columns(4)
        with c1: st.download_button("📥 TXT", texto.encode("utf-8"), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
        with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto, logo_bytes), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
        with c3: st.download_button("📝 Word", generar_word(texto), f"{nombre_limpio}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"wd_{key_u}", use_container_width=True)
        with c4: st.download_button("📊 Excel", generar_excel(texto), f"{nombre_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"ex_{key_u}", use_container_width=True)
    else:
        c1, c2, c3 = st.columns(3)
        with c1: st.download_button("📥 TXT", texto.encode("utf-8"), f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True)
        with c2: st.download_button("📄 PDF", generar_pdf_pro(nombre_limpio, texto, logo_bytes), f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True)
        with c3:
            if st.button("📲 Enviar a WhatsApp", key=f"wa_{key_u}", use_container_width=True):
                num = st.session_state.get("whatsapp_number", "").strip()
                num_limpio = re.sub(r'[^0-9]', '', num)
                if num_limpio:
                    url_wa = f"https://wa.me/{num_limpio}?text={urllib.parse.quote(texto[:1500])}"
                    components.html(f"<script>window.open('{url_wa}', '_blank');</script>", height=0); st.success("✅ Abriendo WhatsApp...")
                else: st.error("⚠️ Configura tu número en la barra lateral izquierda.")

def run_async_safe(coro):
    def _run(): return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor() as pool: return pool.submit(_run).result()
def generar_audio_sincrono(texto: str, voz: str) -> bytes:
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(texto, voz).stream():
            if chunk["type"] == "audio": buf.write(chunk["data"])
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error TTS: {e}"); return None
def generar_podcast_3_voces(guion: str, v1: str, v2: str, v3: str) -> bytes:
    async def _gen():
        buf = BytesIO(); texto = re.sub(r"\(risas\)|\[risas\]|\*risas\*", "jajaja", guion, flags=re.IGNORECASE)
        texto = re.sub(r"\(pausa\)|\[pausa\]", "...", texto, flags=re.IGNORECASE); texto = re.sub(r"[*#_]", "", texto)
        voz_actual = v1
        for linea in texto.split("\n"):
            linea = linea.strip()
            if not linea: continue
            m = re.search(r"^Locutor\s*(\d)", linea, re.IGNORECASE)
            if m:
                num = int(m.group(1)); voz_actual = {1: v1, 2: v2, 3: v3}.get(num, v1)
                linea = re.sub(r"^Locutor\s*\d\s*:?", "", linea, flags=re.IGNORECASE).strip()
            if not linea or linea.lower() in {"elvira", "jorge", "tomas", "carlos"}: continue
            try:
                async for chunk in edge_tts.Communicate(linea + ", ", voz_actual).stream():
                    if chunk["type"] == "audio": buf.write(chunk["data"])
            except: continue
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error crítico: {e}"); return None
def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="...", api_key_openai="", api_key_anthropic=""):
    input_key = f"{key_prefix}_input"
    if input_key not in st.session_state: st.session_state[input_key] = ""
    def agregar_idea(idea):
        txt = st.session_state[input_key].strip(); st.session_state[input_key] = f"{txt}\n\n{idea}" if txt else idea
        if f"{key_prefix}_sugs" in st.session_state: del st.session_state[f"{key_prefix}_sugs"]
    if st.button("💡 Generar Ideas con IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
        with st.spinner("Pensando ideas..."):
            prompt = f"Dame 3 ideas creativas para {contexto_ia}. Solo lista separada por saltos de línea, sin numeración."
            res = llamar_ia(modelo, prompt, api_key, api_key_openai=api_key_openai, api_key_anthropic=api_key_anthropic)
            if "[ERROR" in res: st.error(res)
            else: st.session_state[f"{key_prefix}_sugs"] = [l.strip("-*1234567890. ") for l in res.split("\n") if len(l.strip()) > 5][:3]
    if st.session_state.get(f"{key_prefix}_sugs"):
        st.markdown("✨ **Haz clic en una o varias ideas para añadirlas a tu texto:**")
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]: st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True, on_click=agregar_idea, args=(idea,))
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
    st.text_area("Editar:", key=widget_key, height=height); st.session_state[state_key] = st.session_state[widget_key]
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
def custom_tabs(tab_names, key_name, default_idx=0, cols_per_row=5):
    s_key = f"tab_{key_name}"
    if s_key not in st.session_state: st.session_state[s_key] = default_idx
    for i in range(0, len(tab_names), cols_per_row):
        chunk = tab_names[i:i + cols_per_row]; cols = st.columns(cols_per_row)
        for j, name in enumerate(chunk):
            idx = i + j
            with cols[j]:
                if st.button(name, key=f"btn_{key_name}_{idx}", use_container_width=True, type="primary" if st.session_state[s_key] == idx else "secondary"):
                    st.session_state[s_key] = idx; st.rerun()
    return st.session_state[s_key]

def generar_cancion_con_base(letra: str, voz: str, estilo_base: str) -> bytes:
    try:
        texto_limpio = re.sub(r"\[.*?\]", "", letra)
        voz_bytes = generar_audio_sincrono(texto_limpio, voz)
        if not voz_bytes: return None
        voz_audio = AudioSegment.from_file(BytesIO(voz_bytes), format="mp3")
        urls = {
            "Rap / Hip-Hop": "https://cdn.pixabay.com/download/audio/2022/10/25/audio_946bc9e7eb.mp3",
            "Lo-Fi Chill": "https://cdn.pixabay.com/download/audio/2022/05/27/audio_1808fbf07a.mp3",
            "Épico / Trap": "https://cdn.pixabay.com/download/audio/2022/11/22/audio_1b6d6c0b34.mp3"
        }
        url_base = urls.get(estilo_base, urls["Rap / Hip-Hop"])
        r = requests.get(url_base, headers={"User-Agent": "Mozilla/5.0"})
        base_audio = AudioSegment.from_file(BytesIO(r.content), format="mp3")
        base_audio = base_audio - 12
        if len(voz_audio) > len(base_audio):
            repeticiones = (len(voz_audio) // len(base_audio)) + 1
            base_audio = base_audio * repeticiones
        base_audio = base_audio[:len(voz_audio)]
        cancion = base_audio.overlay(voz_audio)
        out = BytesIO()
        cancion.export(out, format="mp3")
        return out.getvalue()
    except Exception as e:
        st.error(f"Error al mezclar el audio. Detalle: {e}")
        return None

if not st.session_state.sidebar_visible:
    st.markdown("<style>[data-testid=\"collapsedControl\"] {display: none !important;}</style>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
