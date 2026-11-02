import streamlit as st
import streamlit.components.v1 as components
from streamlit_markmap import markmap
import ollama, asyncio, edge_tts, pypdf, openpyxl, csv, os, re, requests, tempfile, time, qrcode, numpy as np, sqlite3, datetime, concurrent.futures, json
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
try:
    import imageio_ffmpeg
    AudioSegment.converter = imageio_ffmpeg.get_ffmpeg_exe()
    AudioSegment.ffprobe = imageio_ffmpeg.get_ffmpeg_exe()
except ImportError: pass

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
    "alex_chat_history": [], "letra_edit": ""
}
for _k, _v in _defaults.items():
    if _k not in st.session_state: st.session_state[_k] = _v

st.markdown("""
<style>
@import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
.stApp { background-color: #000 !important; background-image: radial-gradient(circle at 50% 50%, #111 0%, #000 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
header[data-testid="stHeader"] { background-color: #000 !important; }
[data-testid="stAppDeployButton"], [data-testid="stDecoration"] { display: none !important; }
h1 { font-family: 'Orbitron', sans-serif !important; color: #FFD700 !important; text-shadow: 0 0 15px #FF8C00, 0 0 25px #FFD700 !important; text-align: center; letter-spacing: 3px; margin-top: 0px !important; padding-top: 0px !important; }
h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #fff !important; text-shadow: 0 0 10px #FF007F, 0 0 20px #00f0ff !important; text-align: center; }
[data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: bold !important; text-transform: uppercase !important; background: rgba(0, 240, 255, 0.1) !important; padding: 5px 10px !important; border-left: 4px solid #FF007F !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 8px !important; }
[data-testid="stSidebar"] { background-color: rgba(5, 5, 10, 0.98) !important; border-right: 2px solid #00f0ff !important; }
[data-testid="stFileUploader"], [data-testid="stExpander"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; }
[data-testid="stFileUploaderDropzone"] { background-color: #0a0a0a !important; color: #fff !important; }
div.stButton > button[kind="secondary"] { background-color: #0a0a0a !important; color: #fff !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; border: 1px solid #00f0ff !important; border-radius: 25px !important; transition: all 0.3s !important; width: 100% !important; }
div.stButton > button[kind="secondary"]:hover { background-color: rgba(0, 240, 255, 0.15) !important; color: #00f0ff !important; }
div.stButton > button[kind="primary"] { background: linear-gradient(135deg, #00f0ff 0%, #FF007F 100%) !important; color: #000 !important; font-family: 'Orbitron', sans-serif !important; font-size: 13px !important; font-weight: 900 !important; border: 2px solid #fff !important; border-radius: 25px !important; width: 100% !important; }
.stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"] { background-color: #0a0a0a !important; border: 1px solid #00f0ff !important; border-radius: 8px !important; color: #fff !important; }
[data-testid="stTextInput"] input { border: 1px solid #00f0ff !important; border-radius: 10px !important; color: #fff !important; background: #0a0a0a !important; }
[data-testid="stChatMessage"] { background: rgba(10, 10, 15, 0.8) !important; border: 1px solid #FF007F !important; border-radius: 15px !important; }
.stAlert { border-radius: 10px; font-weight: bold; background-color: rgba(0, 240, 255, 0.1) !important; border-left: 5px solid #00f0ff !important; }
.block-container { padding-top: 1rem !important; padding-bottom: 1rem !important; }
</style>
""", unsafe_allow_html=True)

def inyectar_js_fixer():
    components.html("<script>(function(){var doc=window.parent.document;function f(){doc.querySelectorAll('[data-testid=\"stTextInput\"] *').forEach(function(el){if(el.tagName.toLowerCase()==='input'||el.tagName.toLowerCase()==='div'){el.style.setProperty('background-color','#0a0a0a','important');el.style.setProperty('color','#fff','important');}});}f();setInterval(f,300);})();</script>", height=0)
inyectar_js_fixer()

def limpiar_texto_para_voz(texto):
    if not texto: return ""
    texto = re.sub(r'[\*\#\_\`\~\"\']', '', texto)
    texto = re.sub(r'\[.*?\]', '', texto, flags=re.DOTALL)
    texto = re.sub(r'\(.*?\)', '', texto, flags=re.DOTALL)
    etiquetas = r'(?i)\b(Título|Intro|Verso|Coro|Puente|Outro|Producción|Instrumentos|Ritmo|Estilo|Género|Duración|Tema)\s*\d*\s*:?'
    texto = re.sub(etiquetas, '', texto)
    texto = texto.replace('\n', '. ')
    texto = re.sub(r'\.{2,}', '.', texto)
    texto = re.sub(r'\s+', ' ', texto)
    return texto.strip()

def url_segura(url: str) -> bool:
    try:
        p = urllib.parse.urlparse(url)
        if p.scheme not in ("http", "https") or not p.hostname: return False
        host = socket.gethostbyname(p.hostname)
        ip = ipaddress.ip_address(host)
        return not (ip.is_private or ip.is_loopback or ip.is_link_local or ip.is_reserved or ip.is_multicast)
    except: return False

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
    with st.expander("📎 Adjuntar archivo (Opcional)"):
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

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if st.session_state.get("global_context"):
        prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}"
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
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente(); tmp_img = None
    if logo_bytes:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp:
                tmp.write(logo_bytes); tmp_img = tmp.name
            pdf.image(tmp_img, x=40, y=50, w=130); pdf.ln(150)
        except Exception as e: st.warning(f"No se pudo usar la portada subida: {e}")
        finally:
            if tmp_img and os.path.exists(tmp_img):
                try: os.remove(tmp_img)
                except: pass
    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.add_font("XaviFont", "I", font_path, uni=True); pdf.set_font("XaviFont", size=11)
    else: pdf.set_font("Helvetica", size=11)
    if font_path: pdf.set_font("XaviFont", "B", 24)
    else: pdf.set_font("Helvetica", "B", 24)
    pdf.multi_cell(0, 10, text=titulo, align="C")
    if font_path: pdf.set_font("XaviFont", "I", 12)
    else: pdf.set_font("Helvetica", "I", 12)
    pdf.cell(0, 10, text="Generado por Xavilu Network", align="C", ln=True)
    pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)
    if font_path: pdf.set_font("XaviFont", size=11)
    else: contenido = contenido.encode("latin-1", "replace").decode("latin-1"); pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 8, text=contenido); return bytes(pdf.output())

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
    nombre_elegido = st.text_input("📝 Nombre del archivo:", value=nombre_por_defecto, key=f"name_{key_u}")
    nombre_limpio = re.sub(r"[^a-zA-Z0-9_\- ]", "", nombre_elegido).strip().replace(" ", "_") or nombre_por_defecto
    logo_bytes = None
    with st.expander("📸 Cambiar Portada (Sube tu propia foto)"):
        st.caption("Si no subes foto, el PDF saldrá sin portada (solo título y texto).")
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
                else: st.error("⚠️ Configura tu número en la barra lateral.")

def run_async_safe(coro):
    def _run(): return asyncio.run(coro)
    with concurrent.futures.ThreadPoolExecutor() as pool: return pool.submit(_run).result()

def generar_audio_sincrono(texto: str, voz: str) -> bytes:
    texto_limpio = limpiar_texto_para_voz(texto)
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(texto_limpio, voz).stream():
            if chunk["type"] == "audio": buf.write(chunk["data"])
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error TTS: {e}"); return None

def generar_podcast_3_voces(guion: str, v1: str, v2: str, v3: str) -> bytes:
    async def _gen():
        buf = BytesIO(); voz_actual = v1
        for linea in guion.split("\n"):
            linea = linea.strip()
            if not linea: continue
            m = re.search(r"^Locutor\s*(\d)", linea, re.IGNORECASE)
            if m:
                num = int(m.group(1)); voz_actual = {1: v1, 2: v2, 3: v3}.get(num, v1)
                linea = re.sub(r"^Locutor\s*\d\s*:?", "", linea, flags=re.IGNORECASE).strip()
            if not linea or linea.lower() in {"elvira", "jorge", "tomas", "carlos"}: continue
            linea_limpia = limpiar_texto_para_voz(linea)
            if not linea_limpia: continue
            try:
                async for chunk in edge_tts.Communicate(linea_limpia + ", ", voz_actual).stream():
                    if chunk["type"] == "audio": buf.write(chunk["data"])
            except: continue
        return buf.getvalue()
    try: return run_async_safe(_gen())
    except Exception as e: st.error(f"Error: {e}"); return None

def generar_cancion_con_base(letra: str, voz1: str, estilo_base: str, voz2: str = None) -> bytes:
    try:
        if voz2:
            estrofas = [e.strip() for e in letra.split('\n\n') if e.strip()]
            if len(estrofas) <= 1: estrofas = [e.strip() for e in letra.split('\n') if e.strip()]
            async def _gen_dueto():
                buf_total = BytesIO(); voz_actual = voz1
                for estrofa in estrofas:
                    estrofa_limpia = limpiar_texto_para_voz(estrofa)
                    if not estrofa_limpia: continue
                    try:
                        async for chunk in edge_tts.Communicate(estrofa_limpia, voz_actual).stream():
                            if chunk["type"] == "audio": buf_total.write(chunk["data"])
                    except: pass
                    voz_actual = voz2 if voz_actual == voz1 else voz1
                return buf_total.getvalue()
            voz_bytes = run_async_safe(_gen_dueto())
        else: voz_bytes = generar_audio_sincrono(limpiar_texto_para_voz(letra), voz1)
        if not voz_bytes: return None
        voz_audio = AudioSegment.from_file(BytesIO(voz_bytes), format="mp3")
        urls = {
            "Reggaeton / Urbano": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-1.mp3",
            "Pop / Alegre": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-2.mp3",
            "Balada / Acústico": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-3.mp3",
            "Rap / Hip-Hop": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-4.mp3",
            "Lo-Fi Chill": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-5.mp3",
            "Épico / Trap": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-6.mp3",
            "Rock": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-7.mp3",
            "Electrónica / EDM": "https://www.soundhelix.com/examples/mp3/SoundHelix-Song-8.mp3"
        }
        url_base = urls.get(estilo_base, urls["Reggaeton / Urbano"])
        r = requests.get(url_base, headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if len(r.content) < 100_000: st.error("❌ No se pudo descargar la música."); return None
        base_audio = AudioSegment.from_file(BytesIO(r.content), format="mp3") - 14
        if len(voz_audio) > len(base_audio):
            base_audio = base_audio * ((len(voz_audio) // len(base_audio)) + 1)
        base_audio = base_audio[:len(voz_audio)]
        cancion = base_audio.overlay(voz_audio)
        out = BytesIO(); cancion.export(out, format="mp3"); return out.getvalue()
    except Exception as e:
        st.error(f"Error al mezclar: {e}"); return None

def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="..."):
    input_key = f"{key_prefix}_input"
    if input_key not in st.session_state: st.session_state[input_key] = ""
    def agregar_idea(idea):
        txt = st.session_state[input_key].strip(); st.session_state[input_key] = f"{txt}\n\n{idea}" if txt else idea
        if f"{key_prefix}_sugs" in st.session_state: del st.session_state[f"{key_prefix}_sugs"]
    if st.button("💡 Generar Ideas con IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
        with st.spinner("Pensando..."):
            res = llamar_ia(modelo, f"Dame 3 ideas creativas para {contexto_ia}. Solo lista por líneas.", api_key)
            if "[ERROR" in res: st.error(res)
            else: st.session_state[f"{key_prefix}_sugs"] = [l.strip("-*1234567890. ") for l in res.split("\n") if len(l.strip()) > 5][:3]
    if st.session_state.get(f"{key_prefix}_sugs"):
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]: st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True, on_click=agregar_idea, args=(idea,))
    return st.text_area(label, key=input_key, placeholder=placeholder, height=100)

def boton_superpoder_borrador(texto_base, instruccion, key_suffix, modelo, api_key, state_key, widget_key=None):
    if st.button(f"🚀 {instruccion}", key=f"super_{key_suffix}", type="secondary", use_container_width=True):
        if not texto_base: return
        with st.spinner("Aplicando..."):
            res = llamar_ia(modelo, f"{instruccion}. Devuelve SOLO el texto mejorado:\n\n{texto_base}", api_key)
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

if not st.session_state.sidebar_visible:
    st.markdown("<style>[data-testid=\"collapsedControl\"] {display: none !important;}</style>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if st.button("⛶ MOSTRAR MENÚ Y KEYS", type="primary", use_container_width=True): st.session_state.sidebar_visible = True; st.rerun()

st.markdown("<h1>XAVILU NETWORK</h1>", unsafe_allow_html=True)

if st.session_state.sidebar_visible:
    with st.sidebar:
        try: st.image("logo.png", use_container_width=True)
        except: st.markdown("### 🪐 PANEL DE CONTROL")
        if st.button("⛶ Ocultar Menú", use_container_width=True): st.session_state.sidebar_visible = False; st.rerun()
        st.markdown("---")
        modelo_seleccionado = st.selectbox("Ollama Local", ["llama3.2:latest", "gemma2:2b", "deepseek-r1:1.5b"])
        st.markdown("#### 🔑 Conexiones")
        for prov in ["gemini", "groq", "silicon", "mistral", "openrouter", "deepseek", "together", "cerebras"]:
            st.caption(f"{'✅' if _get_key(prov) else '⚠️'} {prov.capitalize()}")
        with st.expander("⚙️ Keys manuales"):
            st.session_state.api_key_gemini_manual = st.text_input("Gemini", type="password", value=st.session_state.api_key_gemini_manual)
            st.session_state.api_key_groq_manual = st.text_input("Groq", type="password", value=st.session_state.api_key_groq_manual)
            st.session_state.api_key_silicon_manual = st.text_input("SiliconFlow", type="password", value=st.session_state.api_key_silicon_manual)
            st.session_state.api_key_mistral_manual = st.text_input("Mistral", type="password", value=st.session_state.api_key_mistral_manual)
            st.session_state.api_key_openrouter_manual = st.text_input("OpenRouter", type="password", value=st.session_state.api_key_openrouter_manual)
            st.session_state.api_key_deepseek_manual = st.text_input("DeepSeek", type="password", value=st.session_state.api_key_deepseek_manual)
            st.session_state.api_key_together_manual = st.text_input("Together", type="password", value=st.session_state.api_key_together_manual)
            st.session_state.api_key_cerebras_manual = st.text_input("Cerebras", type="password", value=st.session_state.api_key_cerebras_manual)
        st.markdown("---")
        st.session_state.whatsapp_number = st.text_input("📲 WhatsApp:", value=st.session_state.whatsapp_number)
        st.markdown("---")
        archivo_global = st.file_uploader("📁 Cerebro Global", type=["pdf", "txt", "docx", "png", "jpg"], key="up_global")
        if archivo_global and st.session_state.get("global_file_name") != archivo_global.name:
            with st.spinner("Procesando..."):
                st.session_state.global_context = leer_archivo_universal(archivo_global, modelo_seleccionado)[:5000]
                st.session_state.global_file_name = archivo_global.name; st.success("✅ Memorizado.")
        if st.session_state.get("global_file_name"):
            st.info(f"Activo: **{st.session_state.global_file_name}**")
            if st.button("🗑️ Olvidar"): st.session_state.global_context = ""; st.session_state.global_file_name = ""; st.rerun()
        st.markdown("---")
        for h in obtener_de_db():
            with st.expander(f"{h[1]} - {h[0][:10]}"): st.caption(h[2][:200] + "...")
else: modelo_seleccionado = "llama3.2:latest"

macro_names = ["🛠️ XAVIWORK", "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR", "🎵 MÚSICA Y ARTE", "🏢 GESTOR VECINOS"]
macro_idx = custom_tabs(macro_names, "macro", 0, cols_per_row=4)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

if macro_idx == 0:
    st.markdown("#### 🛠️ XaviWork")
    contexto_xw = cajita_adjuntar_archivo("xw_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
    col1, col2 = st.columns([2, 1])
    with col1: st.text_area("💡 ¿Sobre qué quieres trabajar?", height=100, key="xw_tema_input", placeholder="Ej: Estrategia de marketing...")
    with col2:
        st.markdown("<br>", unsafe_allow_html=True)
        if st.button("💡 3 Ideas Base", use_container_width=True):
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"3 ideas sobre: '{st.session_state.xw_tema_input}'. Por líneas.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state["xw_sugs"] = [l.strip("-*1234567890. ") for l in res.split("\n") if len(l.strip()) > 5][:3]
        def _xw_add(idea):
            actual = st.session_state["xw_tema_input"].strip(); st.session_state["xw_tema_input"] = f"{actual}\n{idea}" if actual else idea; st.session_state["xw_sugs"] = []
        if st.session_state.get("xw_sugs"):
            for i, idea in enumerate(st.session_state["xw_sugs"]): st.button(f"➕ {idea}", key=f"xw_idea_{i}", use_container_width=True, on_click=_xw_add, args=(idea,))
    st.markdown("---")
    tabs_xw = st.tabs(["🕸️ Mapa Mental", "🧠 Mapa Conceptual", "📊 PPT", "📄 Informe", "🃏 Tarjetas", "❓ Test", "📋 Tabla", "🎙️ Podcast", "💬 Copiloto", "📅 Timeline", "📖 Glosario", "📊 DAFO", "🖼️ Portada"])
    def _panel_borrador(nombre, prompt_gen, state_key, edit_key, superpowers, download_label=None):
        if st.button(f"✨ Generar {nombre}", key=f"btn_{state_key}", type="primary"):
            if not st.session_state.xw_tema_input.strip(): st.warning("Define el tema arriba.")
            else:
                with st.spinner(f"Generando {nombre}..."):
                    res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\n{prompt_gen}\n\nTEMA: {st.session_state.xw_tema_input}", st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state[state_key] = res.replace("```markdown","").replace("```csv","").replace("```mermaid","").replace("```","").strip() if state_key in ("xw_mapa","xw_tabla","xw_conceptual") else res
        if st.session_state.get(state_key):
            if state_key != "xw_conceptual": borrador_editable(state_key, edit_key, height=250)
            else:
                st.markdown("### 🧠 Mapa Visual"); render_mermaid(st.session_state[state_key])
                with st.expander("🛠️ Editar código"): st.text_area("Mermaid:", key=edit_key, height=200); st.session_state[state_key] = st.session_state[edit_key]
            n = len(superpowers) + (1 if download_label else 0)
            if n == 0: return
            cols = st.columns(n)
            for i, (label, prompt) in enumerate(superpowers):
                with cols[i]: boton_superpoder_borrador(st.session_state[state_key], prompt, f"{state_key}_p{i}", modelo_seleccionado, st.session_state.api_key_gemini_manual, state_key, widget_key=edit_key)
            if download_label:
                with cols[-1]:
                    if state_key == "xw_tabla": st.download_button(download_label, generar_excel(st.session_state[state_key]), "tabla.xlsx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key == "xw_ppt": st.download_button(download_label, generar_pptx("PPT", st.session_state[state_key]), "presentacion.pptx", key=f"dl_{state_key}", use_container_width=True)
                    elif state_key in ("xw_informe","xw_timeline","xw_glosario","xw_dafo"): botones_descarga(st.session_state[state_key], state_key, formatos_extra=True)
                    else: st.download_button(download_label, st.session_state[state_key].encode(), f"{state_key}.txt", key=f"dl_{state_key}", use_container_width=True)
    with tabs_xw[0]:
        _panel_borrador("Mapa Mental", "Mapa mental jerárquico en Markdown (#, ##, -). Solo el Markdown.", "xw_mapa", "edit_xw_mapa", [("Más profesional", "Hazlo más profesional"), ("Más detalles", "Añade más detalles")], "📥 Descargar Markdown")
        if st.session_state.get("xw_mapa"):
            with st.expander("🕸️ Ver mapa interactivo"): markmap(st.session_state.xw_mapa, height=400)
    with tabs_xw[1]: _panel_borrador("Mapa Conceptual", "Mermaid 'graph TD' con emojis y colores 'style'. SOLO código Mermaid.", "xw_conceptual", "edit_xw_conceptual", [("Más colorido", "Más colorido y emojis"), ("Más ramas", "Añade ramas")], "📥 Descargar Código")
    with tabs_xw[2]: _panel_borrador("Presentación (PPTX)", "Estructura de presentación. Diapositivas separadas por doble salto. 1ª línea=título, guiones=puntos.", "xw_ppt", "edit_xw_ppt", [("Más visual", "Más visual")], "📥 Descargar PPTX")
    with tabs_xw[3]: _panel_borrador("Informe", "Informe profesional: introducción, desarrollo, conclusiones.", "xw_informe", "edit_xw_informe", [("Más técnico", "Más técnico"), ("Resumir", "Resumir a la mitad"), ("Inglés", "Traducir al Inglés")], "📥 Descargar")
    with tabs_xw[4]: _panel_borrador("Tarjetas", "5 flashcards Pregunta -> Respuesta.", "xw_tarjetas", "edit_xw_tarjetas", [("Más difíciles", "Hazlas más difíciles")], "📥 Descargar")
    with tabs_xw[5]: _panel_borrador("Test", "5 preguntas tipo test (3 opciones + respuesta al final).", "xw_cuestionario", "edit_xw_cuest", [("+5 preguntas", "Añade 5 preguntas más")], "📥 Descargar")
    with tabs_xw[6]: _panel_borrador("Tabla", "Tabla CSV. SOLO el CSV.", "xw_tabla", "edit_xw_tabla", [("Ordenar", "Ordena por 1ª columna")], "📥 Descargar Excel")
    with tabs_xw[7]:
        _panel_borrador("Guion Podcast", "Guion 3 locutores. Formato: 'Locutor 1: [texto]'.", "xw_podcast", "edit_xw_pod", [("Más gracioso", "Más gracioso")], None)
        if st.session_state.get("xw_podcast"):
            if st.button("🔊 Generar Audio", key="btn_xw_pod_aud"):
                with st.spinner("Generando..."):
                    aud = generar_podcast_3_voces(st.session_state.xw_podcast, "es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural")
                    if aud: st.session_state.xw_podcast_audio = aud; st.success("✅ Listo.")
            if st.session_state.get("xw_podcast_audio"):
                st.audio(st.session_state.xw_podcast_audio, format="audio/mp3"); st.download_button("💾 MP3", st.session_state.xw_podcast_audio, "podcast_xaviwork.mp3")
    with tabs_xw[8]:
        cc = st.container(height=400)
        with cc:
            for msg in st.session_state.xw_copilot_history:
                with st.chat_message(msg["role"]): st.write(msg["content"])
        if prompt_chat := st.chat_input("Pregunta sobre el tema..."):
            st.session_state.xw_copilot_history.append({"role": "user", "content": prompt_chat})
            with cc:
                with st.chat_message("user"): st.write(prompt_chat)
                with st.chat_message("assistant"):
                    with st.spinner("Pensando..."):
                        ctx_chat = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.xw_copilot_history])
                        res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nTema: {st.session_state.xw_tema_input}\n\n{ctx_chat}", st.session_state.api_key_gemini_manual)
                        st.write(res); st.session_state.xw_copilot_history.append({"role": "assistant", "content": res})
    with tabs_xw[9]: _panel_borrador("Timeline", "Eventos cronológicos: Año/Fecha - Evento.", "xw_timeline", "edit_xw_time", [("+fechas", "Añade más fechas")], "📥 Descargar")
    with tabs_xw[10]: _panel_borrador("Glosario", "5 conceptos complejos con definiciones sencillas.", "xw_glosario", "edit_xw_glos", [("Para niño", "Explícalo para un niño")], "📥 Descargar")
    with tabs_xw[11]: _panel_borrador("DAFO", "Análisis DAFO conciso.", "xw_dafo", "edit_xw_dafo", [("Más estratégico", "Más estratégico")], "📥 Descargar")
    with tabs_xw[12]:
        st.markdown("Genera una imagen ilustrativa.")
        prompt_img = st.text_input("Describe la imagen:", value=st.session_state.xw_tema_input, key="xw_img_prompt")
        if st.button("🖼️ Generar Portada", key="btn_xw_img", type="primary") and prompt_img:
            with st.spinner("Generando..."):
                try:
                    url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img + ', cyberpunk, high quality')}?width=1024&height=1024&nologo=true"
                    if url_segura(url_img):
                        req = urllib.request.Request(url_img, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=30) as response: st.session_state.xw_cover_img = response.read()
                except Exception as e: st.error(f"Error: {e}")
        if st.session_state.get("xw_cover_img"):
            st.image(st.session_state.xw_cover_img, use_container_width=True)
            st.download_button("📥 Descargar Imagen", st.session_state.xw_cover_img, "portada.png", "image/png", use_container_width=True)

elif macro_idx == 1:
    sub_names = ["🎙️ PODCAST", "💬 CHAT", "🎙️ VOZ", "💬 CHAT CON PDF", "🕵️ AGENTES", "⚖️ DEBATE", "🍳 COCINA", "📖 CUENTOS", "💻 CÓDIGO", "🌍 TRADUCTOR"]
    sub_idx = custom_tabs(sub_names, "chat_sub", 0)
    if sub_idx == 0:
        st.markdown("#### 🎙️ Podcasts")
        contexto = cajita_adjuntar_archivo("podcast_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        st.text_area("Temática:", key="pod_tema_input", height=100)
        c1, c2 = st.columns(2)
        with c1: estilo = st.selectbox("Estilo:", ["Entrevista","Debate","Educativo","Narrativo"], key="p_est")
        with c2: dur = st.slider("Dur (min):", 1, 10, 3, key="p_dur")
        st.markdown("---")
        if st.button("✨ Crear Guion", type="primary", key="btn_pod") and st.session_state["pod_tema_input"]:
            with st.spinner("Escribiendo..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\nGuion podcast {dur} min. Tema: '{st.session_state['pod_tema_input']}'. Estilo: {estilo}. 3 locutores. Formato 'Locutor 1: [texto]'.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.podcast_script = res
        if st.session_state.get("podcast_script"):
            st.markdown("### 📜 Guion"); borrador_editable("podcast_script", "pod_edit", height=300)
            c1, c2, c3 = st.columns(3)
            with c1: v1 = st.selectbox("Voz L1:", ["es-ES-ElviraNeural","es-MX-JorgeNeural"], key="v1")
            with c2: v2 = st.selectbox("Voz L2:", ["es-ES-ElviraNeural","es-MX-JorgeNeural"], index=1, key="v2")
            with c3: v3 = st.selectbox("Voz L3:", ["es-AR-TomasNeural","es-ES-ElviraNeural"], key="v3")
            if st.button("🔊 Audio 3 Voces", key="btn_pod_aud"):
                with st.spinner("Generando..."):
                    aud = generar_podcast_3_voces(st.session_state.podcast_script, v1, v2, v3)
                    if aud: st.audio(aud, format="audio/mp3"); st.download_button("💾 MP3", aud, "podcast.mp3")
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
            st.session_state.chat_history.append({"role":"user","content":p}); st.chat_message("user").write(p)
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"{ctx}\n\n{p}", st.session_state.api_key_gemini_manual, st.session_state.chat_history)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.chat_history.append({"role":"assistant","content":res}); st.chat_message("assistant").write(res)
            st.rerun()
    elif sub_idx == 2:
        st.markdown("#### Voz")
        txt = st.text_area("Texto:"); voz = st.selectbox("Voz:", ["es-ES-ElviraNeural","es-MX-JorgeNeural"])
        if st.button("🔊 Audio") and txt:
            aud = generar_audio_sincrono(txt, voz)
            if aud: st.audio(aud, format="audio/mp3")
    elif sub_idx == 3:
        st.markdown("#### Chat con PDFs")
        archs = st.file_uploader("Docs:", type=["pdf","txt","docx"], accept_multiple_files=True)
        if archs:
            texto = "".join([leer_archivo_universal(a, modelo_seleccionado) for a in archs])[:MAX_TEXTO_IA]
            st.success("Cargados.")
            for m in st.session_state.pdf_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("pdf")
            if p:
                st.session_state.pdf_chat_history.append({"role":"user","content":p})
                with st.spinner("Buscando..."):
                    res = llamar_ia(modelo_seleccionado, f"Contexto:\n{texto}\n\nPregunta: {p}", st.session_state.api_key_gemini_manual, st.session_state.pdf_chat_history)
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state.pdf_chat_history.append({"role":"assistant","content":res}); st.rerun()
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
                st.session_state.agente_chat_history.append({"role":"user","content":p})
                with st.spinner("Pensando..."):
                    res = llamar_ia(modelo_seleccionado, f"ERES: {ag['rol']}\nUsuario: {p}", st.session_state.api_key_gemini_manual, st.session_state.agente_chat_history)
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state.agente_chat_history.append({"role":"assistant","content":res}); st.rerun()
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
        with c1: o = st.selectbox("Origen:", ["Auto","Español","Inglés"], key="o")
        with c2: d = st.selectbox("Destino:", ["Español","Inglés","Francés"], key="d")
        if st.button("Traducir") and txt:
            res = llamar_ia(modelo_seleccionado, f"Traduce del {o} al {d}:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
            else: st.info(res); botones_descarga(res, "Traduccion")

elif macro_idx == 2:
    sub_names = ["📁 ANÁLISIS", "⚖️ COMPARAR", "📊 OFFICE", "📄 PDF", "🧾 FACTURAS", "📇 QR & CV", "📋 PORTAPAPELES"]
    sub_idx = custom_tabs(sub_names, "doc_sub", 0)
    if sub_idx == 0:
        st.markdown("#### Análisis")
        arch = st.file_uploader("PDF, Word, Excel o TXT", type=["pdf","txt","docx","xlsx"])
        opc = st.selectbox("Análisis:", ["Resumen Ejecutivo","Extraer Fechas y Cifras","Buscar Errores","Crear Test","Traducir al Inglés"])
        if arch and st.button("🔎 Ejecutar", type="primary"):
            with st.spinner("Analizando..."):
                texto = leer_archivo_universal(arch, modelo_seleccionado)
                prompt = {"Resumen Ejecutivo": f"Resumen:\n{texto}", "Extraer Fechas y Cifras": f"Fechas:\n{texto}", "Buscar Errores": f"Errores:\n{texto}", "Crear Test": f"Test 5 preguntas:\n{texto}", "Traducir al Inglés": f"Traduce al inglés:\n{texto}"}[opc]
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resultado:", res, height=400); botones_descarga(res, "Analisis")
    elif sub_idx == 1:
        st.markdown("#### Comparador")
        a = st.file_uploader("Original:", type=["pdf","txt"]); b = st.file_uploader("Modificado:", type=["pdf","txt"])
        if a and b and st.button("Comparar"):
            with st.spinner("Buscando..."):
                ta = leer_archivo_universal(a, modelo_seleccionado); tb = leer_archivo_universal(b, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"Diferencias:\nA:\n{ta[:10000]}\nB:\n{tb[:10000]}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Diferencias:", res, height=400); botones_descarga(res, "Comparativa")
    elif sub_idx == 2:
        st.markdown("#### Crear Documentos")
        contexto = cajita_adjuntar_archivo("office_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        prompt_office = campo_con_ideas("¿Qué documento?", "office", "documentos", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Presupuesto")
        formato = st.radio("Formato:", ["Word","Excel","PowerPoint"], key="office_fmt", horizontal=True)
        if st.button("✨ Generar", type="primary", key="btn_office_gen"):
            if prompt_office:
                with st.spinner("Generando..."):
                    prompt_ia = f"{contexto}\n\n{prompt_office}"
                    if formato == "PowerPoint": prompt_ia += "\n\nFormatea para diapositivas con doble salto."
                    elif formato == "Excel": prompt_ia += "\n\nDevuelve CSV."
                    res = llamar_ia(modelo_seleccionado, prompt_ia, st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.session_state.office_borrador = res
            else: st.warning("Escribe un prompt.")
        if st.session_state.get("office_borrador"):
            st.markdown("### 📝 Borrador"); borrador_editable("office_borrador", "office_txt", height=350)
            c1, c2, c3 = st.columns(3)
            with c1: boton_superpoder_borrador(st.session_state.office_txt, "Más profesional", "of_sup1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            with c2: boton_superpoder_borrador(st.session_state.office_txt, "Resúmelo", "of_sup2", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            with c3: boton_superpoder_borrador(st.session_state.office_txt, "+detalles", "of_sup3", modelo_seleccionado, st.session_state.api_key_gemini_manual, "office_borrador", widget_key="office_txt")
            st.markdown("---")
            if formato == "Word": st.download_button("📥 Word", generar_word(st.session_state.office_txt), "doc.docx")
            elif formato == "Excel": st.download_button("📥 Excel", generar_excel(st.session_state.office_txt), "doc.xlsx")
            else: st.download_button("📥 PPTX", generar_pptx("Presentación", st.session_state.office_txt), "doc.pptx")
    elif sub_idx == 3:
        st.markdown("#### PDF Pro")
        contexto = cajita_adjuntar_archivo("pdf_adj", modelo_seleccionado, st.session_state.api_key_gemini_manual)
        titulo = st.text_input("Título PDF:")
        contenido = campo_con_ideas("Contenido:", "pdf_c", "contenido PDF", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Escribe...")
        if st.button("✨ Generar PDF", type="primary", key="btn_pdf_gen") and titulo and contenido:
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\n{contenido}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.pdf_borrador = res
        if st.session_state.get("pdf_borrador"):
            st.markdown("### 📝 Borrador"); borrador_editable("pdf_borrador", "pdf_txt", height=350)
            c1, c2 = st.columns(2)
            with c1: boton_superpoder_borrador(st.session_state.pdf_txt, "Más técnico", "pdf_sup1", modelo_seleccionado, st.session_state.api_key_gemini_manual, "pdf_borrador", widget_key="pdf_txt")
            with c2: boton_superpoder_borrador(st.session_state.pdf_txt, "+conclusiones", "pdf_sup2", modelo_seleccionado, st.session_state.api_key_gemini_manual, "pdf_borrador", widget_key="pdf_txt")
            st.markdown("---"); botones_descarga(st.session_state.pdf_borrador, titulo or "PDF_Pro")
    elif sub_idx == 4:
        st.markdown("#### 🧾 Facturas")
        arch_factura = st.file_uploader("Factura (JPG/PNG/PDF):", type=["jpg","jpeg","png","pdf"])
        if st.button("🪄 Extraer", type="primary") and arch_factura:
            with st.spinner("Escaneando..."):
                prompt_factura = "Extrae datos. Devuelve CSV: Proveedor, Fecha, Concepto, Base_Imponible, IVA, Total. Sin markdown."
                if arch_factura.name.lower().endswith(".pdf"):
                    texto_pdf = leer_archivo_universal(arch_factura, modelo_seleccionado)
                    res = llamar_ia(modelo_seleccionado, f"{prompt_factura}\n\n{texto_pdf}", st.session_state.api_key_gemini_manual)
                else: res = analizar_archivo_con_gemini(arch_factura.getvalue(), "image/jpeg", prompt_factura, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.success("✅ Listo."); st.text_area("Datos:", value=res, height=150, key="factura_data")
                    st.download_button("📊 Excel", generar_excel(st.session_state.factura_data), "factura.xlsx", use_container_width=True)
    elif sub_idx == 5:
        c1, c2 = st.columns(2)
        with c1:
            txt = st.text_input("Texto/URL QR:")
            if st.button("Generar QR") and txt:
                qr = qrcode.QRCode(); qr.add_data(txt); qr.make(fit=True)
                buf = BytesIO(); qr.make_image().save(buf, format="PNG"); st.image(buf.getvalue())
        with c2:
            img = st.file_uploader("Leer QR:", type=["png","jpg"])
            if img and HAS_CV2:
                d, _, _ = cv2.QRCodeDetector().detectAndDecode(np.array(Image.open(img).convert('RGB')))
                if d: st.success(d)
    elif sub_idx == 6:
        st.markdown("#### Portapapeles")
        txt = st.text_area("Texto:"); inst = campo_con_ideas("¿Qué hago?", "clip", "instrucciones", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Resume")
        if st.button("Procesar") and txt and inst:
            res = llamar_ia(modelo_seleccionado, f"{inst}:\n{txt}", st.session_state.api_key_gemini_manual)
            if "[ERROR" in res: st.error(res)
            else: st.info(res)

elif macro_idx == 3:
    sub_names = ["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV", "📊 WEB A EXCEL", "🎥 YOUTUBE", "📰 LECTOR"]
    sub_idx = custom_tabs(sub_names, "web_sub", 0)
    if sub_idx == 0:
        q = campo_con_ideas("Buscar:", "web", "búsquedas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Noticias IA")
        if st.button("🔍 Buscar") and q:
            for r in buscar_en_internet(q): st.markdown(f"**[{r['titulo']}]({r['url']})**"); st.caption(r['resumen'])
    elif sub_idx == 1:
        d = st.selectbox("Deporte:", ["Fútbol","Baloncesto","Tenis"])
        if st.button("Noticias"):
            for n in buscar_noticias_deportivas(d, "Todas"): st.markdown(f"**{n['titulo']}** - {n['resumen']}")
    elif sub_idx == 2:
        c = campo_con_ideas("Canal:", "tv", "TV", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Telecinco")
        if st.button("Programación") and c:
            res = buscar_en_internet(c); r_ia = llamar_ia(modelo_seleccionado, f"Programación {c}: {res}", st.session_state.api_key_gemini_manual)
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
        fuente_tipo = st.radio("Método:", ["Subir archivo", "Enlace de YouTube"], horizontal=True)
        archivo_subido = None; url = ""
        if fuente_tipo == "Subir archivo": archivo_subido = st.file_uploader("Sube audio/vídeo:", type=["mp3","wav","m4a","mp4","mov"])
        else: url = st.text_input("URL del vídeo:")
        modo = st.radio("Acción:", ["Resumir", "Traducir al español"], horizontal=True)
        if st.button("Ejecutar", type="primary") and (archivo_subido or url):
            texto_original = ""; vid = extraer_youtube_id(url) if url else "local"; key_groq = _get_key("groq")
            if archivo_subido:
                with st.spinner("Transcribiendo..."): texto_original = transcribir_archivo_local(archivo_subido.getvalue(), archivo_subido.name, key_groq)
            else:
                if not vid: st.error("No hay ID válido."); st.stop()
                try:
                    with st.spinner("Buscando subtítulos..."): tr = YouTubeTranscriptApi.get_transcript(vid, languages=['es', 'en'])
                    texto_original = " ".join([t['text'] for t in tr])
                except:
                    st.info("Sin subtítulos. Transcribiendo audio...")
                    with st.spinner("Descargando y transcribiendo..."): texto_original = transcribir_audio_youtube(url, key_groq)
            if texto_original.startswith("[ERROR]"): st.error(texto_original)
            else:
                with st.expander("Ver transcripción"): st.text_area("Texto:", value=texto_original, height=150, key=f"yt_orig_{vid}")
                prompt = f"Resume esto en español:\n{texto_original[:8000]}" if modo == "Resumir" else f"Traduce al español:\n{texto_original[:8000]}"
                with st.spinner("Procesando..."): res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resultado:", res, height=300); botones_descarga(res, f"YT_{vid}")
    elif sub_idx == 5:
        url = st.text_input("URL Web:")
        if st.button("Leer") and url:
            try:
                if not url_segura(url): st.error("URL no permitida."); st.stop()
                txt = " ".join(re.sub('<[^<]+?>', ' ', requests.get(url, headers={"User-Agent":"Mozilla/5.0"}).text).split())[:5000]
                res = llamar_ia(modelo_seleccionado, f"Resume: {txt}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Resumen:", res, height=200)
            except Exception as e: st.error(f"Error: {e}")

elif macro_idx == 4:
    st.markdown("#### 🎬 Multimedia")
    sub_names = ["🖼️ IMAGEN", "🎙️ TEXTO A VOZ", "🎬 VIDEO (Pronto)"]
    sub_idx = custom_tabs(sub_names, "multi_sub", 0)
    if sub_idx == 0:
        st.markdown("### Generador de Imágenes")
        prompt_img = st.text_input("Describe la imagen:", key="img_gen_prompt"); estilo_img = st.selectbox("Estilo:", ["Cyberpunk", "Realista", "Anime", "3D Render", "Fotografía"], key="img_style")
        if st.button("🖼️ Generar Imagen", type="primary") and prompt_img:
            with st.spinner("Generando..."):
                try:
                    query = urllib.parse.quote(f"{prompt_img}, {estilo_img}, high quality, 8k")
                    url = f"https://image.pollinations.ai/prompt/{query}?width=1024&height=1024&nologo=true"
                    if url_segura(url):
                        req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                        with urllib.request.urlopen(req, timeout=30) as response: img_data = response.read()
                        st.image(img_data, caption=prompt_img, use_container_width=True); st.download_button("📥 Descargar Imagen", img_data, "imagen_generada.png", "image/png")
                except Exception as e: st.error(f"Error: {e}")
    elif sub_idx == 1:
        st.markdown("### Texto a Voz (TTS)")
        txt_tts = st.text_area("Texto:", height=150); voz_tts = st.selectbox("Voz:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"])
        if st.button("🔊 Generar Audio", type="primary") and txt_tts:
            with st.spinner("Sintetizando..."):
                aud = generar_audio_sincrono(txt_tts, voz_tts)
                if aud: st.audio(aud, format="audio/mp3"); st.download_button("💾 Descargar MP3", aud, "audio.mp3")

elif macro_idx == 5:
    st.markdown("#### 🎓 Estudio y Oficios")
    sub_names = ["📚 DEBERES", "👔 ENTREVISTA", "🛠️ PRL PRO"]
    sub_idx = custom_tabs(sub_names, "study_sub", 0)
    if sub_idx == 0:
        st.markdown("### Resolver Deberes")
        materia = st.selectbox("Materia:", ["Matemáticas", "Historia", "Ciencias", "Lengua", "Física y Química", "Inglés"]); pregunta = st.text_area("Duda:", height=150)
        if st.button("🔎 Resolver", type="primary") and pregunta:
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"Profesor experto en {materia}. Resuelve paso a paso:\n\n{pregunta}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, "Deberes_Resueltos", formatos_extra=True)
    elif sub_idx == 1:
        st.markdown("### Simulador de Entrevista")
        puesto = st.text_input("¿Puesto?", placeholder="Ej: Camarero...")
        if st.button("🤖 Empezar", type="primary") and puesto:
            st.session_state.entrevista_puesto = puesto
            st.session_state.entrevista_historial = [{"role": "user", "content": f"Actúa como reclutador para {puesto}. Preséntate y haz la primera pregunta."}]
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, st.session_state.entrevista_historial[0]["content"], st.session_state.api_key_gemini_manual)
                st.session_state.entrevista_historial.append({"role": "assistant", "content": res}); st.rerun()
        if "entrevista_historial" in st.session_state and st.session_state.entrevista_historial:
            for msg in st.session_state.entrevista_historial:
                if msg["role"] == "user" and "Actúa como" not in msg["content"]: st.chat_message("user").write(msg["content"])
                elif msg["role"] == "assistant": st.chat_message("assistant").write(msg["content"])
            resp = chat_input_pro("entrevista_input", "Respuesta...")
            if resp:
                st.session_state.entrevista_historial.append({"role": "user", "content": resp})
                with st.spinner("Evaluando..."):
                    contexto = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.entrevista_historial])
                    prompt = f"Continúa la entrevista. Evalúa brevemente mi respuesta y haz otra pregunta.\n\nHistorial:\n{contexto}"
                    res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                    st.session_state.entrevista_historial.append({"role": "assistant", "content": res}); st.rerun()
    elif sub_idx == 2:
        st.markdown("### 🦺 PRL PRO")
        tabs_prl = st.tabs(["📊 Evaluación", "📝 Test", "📖 Normativa"])
        with tabs_prl[0]:
            puesto = st.text_input("Puesto:", placeholder="Ej: Soldador..."); entorno = st.text_area("Entorno:")
            if st.button("🔍 Evaluar", type="primary") and puesto:
                with st.spinner("Analizando..."):
                    res = llamar_ia(modelo_seleccionado, f"Técnico PRL. Evalúa riesgos para '{puesto}'. Contexto: '{entorno}'.\nTabla Markdown: | Riesgo | Probabilidad | Consecuencia | Nivel | Medidas | EPIs |\nAñade 'Recomendaciones'.", st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.markdown(res); botones_descarga(res, f"Evaluacion_PRL", formatos_extra=True)
        with tabs_prl[1]:
            tema_test = st.selectbox("Área:", ["Ley 31/1995", "EPIs", "Primeros Auxilios", "Ergonomía", "Máquinas"])
            if st.button("📝 Generar Test", type="primary"):
                with st.spinner("Preparando..."):
                    res = llamar_ia(modelo_seleccionado, f"Test 10 preguntas (A,B,C,D) sobre '{tema_test}'. Incluye 'Soluciones'. Normativa española.", st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.markdown(res); botones_descarga(res, "Test_PRL", formatos_extra=True)
        with tabs_prl[2]:
            if st.button("📖 Resumen Leyes"):
                with st.spinner("Resumiendo..."):
                    res = llamar_ia(modelo_seleccionado, "Resumen estructurado Ley 31/1995 PRL: 1. Objeto, 2. Principios, 3. Derechos y obligaciones.", st.session_state.api_key_gemini_manual)
                    if "[ERROR" in res: st.error(res)
                    else: st.markdown(res); botones_descarga(res, "Resumen_Ley_PRL", formatos_extra=True)

elif macro_idx == 6:
    st.markdown("#### 🌍 Viajes y Hogar")
    sub_names = ["✈️ ITINERARIO", "🏠 COCINA", "💪 ENTRENADOR"]
    sub_idx = custom_tabs(sub_names, "travel_sub", 0)
    if sub_idx == 0:
        st.markdown("### Planificador de Viajes")
        destino = st.text_input("Destino:"); c1, c2 = st.columns(2)
        with c1: dias = st.number_input("Días:", min_value=1, max_value=30, value=3)
        with c2: tipo_viaje = st.selectbox("Estilo:", ["Cultural", "Gastronómico", "Aventura", "Relax", "Familiar", "Económico"])
        if st.button("🗺️ Crear", type="primary") and destino:
            with st.spinner("Planificando..."):
                res = llamar_ia(modelo_seleccionado, f"Itinerario {dias} días para {destino}. Estilo: {tipo_viaje}. Incluye qué ver, dónde comer y consejos.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, f"Viaje_{destino}", formatos_extra=True)
    elif sub_idx == 1:
        st.markdown("### Recetas")
        ingredientes = campo_con_ideas("Ingredientes:", "hogar", "recetas", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: Huevos...")
        if st.button("🍳 Crear", type="primary") and ingredientes:
            with st.spinner("Cocinando..."):
                res = llamar_ia(modelo_seleccionado, f"Tengo: {ingredientes}. Dame 2 recetas fáciles. Incluye pasos y tiempo.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, "Recetas_Hogar")
    elif sub_idx == 2:
        st.markdown("### 💪 Entrenador IA")
        c1, c2 = st.columns(2)
        with c1:
            edad = st.number_input("Edad:", min_value=14, max_value=99, value=25); peso = st.number_input("Peso:", min_value=40, max_value=200, value=75); altura = st.number_input("Altura:", min_value=130, max_value=220, value=175)
        with c2:
            objetivo = st.selectbox("Objetivo:", ["Perder grasa", "Ganar músculo", "Mantenerme en forma", "Resistencia"])
            equipo = st.selectbox("Material:", ["Gimnasio", "Mancuernas", "Sin material", "Aire libre"]); dias_entreno = st.slider("Días:", 2, 6, 4)
        if st.button("📋 Generar Plan", type="primary"):
            with st.spinner("Diseñando..."):
                res = llamar_ia(modelo_seleccionado, f"Entrenador. Usuario: {edad}a, {peso}kg, {altura}cm. Objetivo: '{objetivo}'. Entrena {dias_entreno} días en '{equipo}'. Genera: 1. Rutina semanal. 2. Consejos nutricionales.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, "Plan_Entrenamiento", formatos_extra=True)

elif macro_idx == 7:
    st.markdown("#### 🎵 Música y Arte")
    sub_names = ["🎤 LETRAS", "🌞 SUNO AI", "🖼️ PORTADAS", "📝 POESÍA"]
    sub_idx = custom_tabs(sub_names, "music_sub", 0)
    if sub_idx == 0:
        st.markdown("### Generador de Letras")
        if "letra_edit" not in st.session_state: st.session_state.letra_edit = ""
        tema_cancion = st.text_input("Tema (para la IA):", placeholder="Ej: Un desamor en verano...")
        c1, c2 = st.columns(2)
        with c1: genero = st.selectbox("Género:", ["Reggaeton", "Pop", "Rap/Trap", "Rock", "Bachata", "Indie"])
        with c2: artista_estilo = st.text_input("Estilo artista:", placeholder="Ej: Bad Bunny...")
        if st.button("🪄 Generar Letra IA", use_container_width=True) and tema_cancion:
            with st.spinner("Componiendo..."):
                prompt = f"Letra de canción {genero} sobre: '{tema_cancion}'. "
                if artista_estilo: prompt += f"Estilo {artista_estilo}. "
                prompt += "Estructura (Intro, Verso 1, Coro, etc.) y añade indicaciones de producción entre corchetes."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.session_state.letra_edit = res; st.rerun()
        letra_final = st.text_area("📜 Letra (Editable):", key="letra_edit", height=300)
        if letra_final.strip():
            botones_descarga(letra_final, "Letra_Cancion")
            st.markdown("---"); st.markdown("### 🎵 Ponerle Música"); hacer_dueto = st.checkbox("🎤 Hacer Dúo")
            c1, c2 = st.columns(2)
            with c1: voz_cancion1 = st.selectbox("Voz 1:", ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"])
            with c2: 
                if hacer_dueto: voz_cancion2 = st.selectbox("Voz 2:", ["es-MX-JorgeNeural", "es-ES-ElviraNeural", "es-AR-TomasNeural"])
                else: voz_cancion2 = None; st.selectbox("Base:", ["Reggaeton / Urbano", "Pop / Alegre", "Balada / Acústico", "Rap / Hip-Hop", "Lo-Fi Chill", "Épico / Trap", "Rock", "Electrónica / EDM"], key="beat_select_solo")
            if hacer_dueto:
                st.info("💡 *El sistema alternará voces por párrafo.*")
                beat_estilo = st.selectbox("Base:", ["Reggaeton / Urbano", "Pop / Alegre", "Balada / Acústico", "Rap / Hip-Hop", "Lo-Fi Chill", "Épico / Trap", "Rock", "Electrónica / EDM"], key="beat_select_duo")
            if st.button("🎶 Generar Canción", type="primary"):
                with st.spinner("Grabando..."):
                    audio_final = generar_cancion_con_base(letra_final, voz_cancion1, beat_estilo, voz_cancion2)
                    if audio_final:
                        st.success("✅ ¡Lista!"); st.audio(audio_final, format="audio/mp3"); st.download_button("💾 MP3", audio_final, "cancion_xavilu.mp3", "audio/mp3", use_container_width=True)
    elif sub_idx == 1:
        st.markdown("### 🌞 Generador de Prompts para Suno AI")
        st.caption("Suno AI canta de verdad. Genera el texto perfecto para pegarlo en su web.")
        c1, c2 = st.columns(2)
        with c1:
            st_tema = st.text_input("Tema:", placeholder="Ej: Marta y Luz en el campamento...")
            st_genero = st.selectbox("Género:", ["Bachata", "Reggaeton", "Pop", "Rock", "Rap", "Cumbia", "Balada", "Electrónica"])
            st_voz = st.selectbox("Voz:", ["Dúo (Hombre y Mujer)", "Masculina", "Femenina", "Coro"])
        with c2:
            st_instr = st.text_input("Instrumentos:", placeholder="Ej: Guitarra, marimba...")
            st_artista = st.text_input("Artista similar:", placeholder="Ej: Romeo Santos...")
        if st.button("✨ Generar Prompt para Suno", type="primary") and st_tema:
            with st.spinner("Preparando el prompt..."):
                prompt_ia = f"Actúa como experto en Suno AI. Tema: '{st_tema}'. Género: {st_genero}. Voz: {st_voz}. Instrumentos: {st_instr}. Artista similar: {st_artista}. Devuelve un JSON con dos campos: 'style_prompt' (en inglés, máximo 200 caracteres, describiendo estilo, instrumentos y voz) y 'lyrics' (letra en español con etiquetas [Verse], [Chorus], etc.)."
                res = llamar_ia(modelo_seleccionado, prompt_ia, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else:
                    st.success("✅ Prompt generado.")
                    st.text_area("Copia esto en el apartado 'Style of Music':", value=res.split("'lyrics':")[0].replace("{", "").replace("'style_prompt':", "").replace('"', '').strip(), height=100)
                    st.text_area("Copia esto en 'Lyrics' (Custom Mode):", value=res.split("'lyrics':")[1].replace("}", "").replace('"', '').strip() if "'lyrics':" in res else res, height=250)
                    st.markdown("[🚀 Ir a Suno AI a crear la canción](https://suno.com/create)")
    elif sub_idx == 2:
        st.markdown("### 🖼️ Portadas")
        prompt_img = st.text_input("Describe la portada:")
        if st.button("Generar") and prompt_img:
            with st.spinner("Pintando..."):
                url = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img + ', album cover, high quality')}?width=1024&height=1024&nologo=true"
                if url_segura(url):
                    req = urllib.request.Request(url, headers={'User-Agent': 'Mozilla/5.0'})
                    with urllib.request.urlopen(req, timeout=30) as response: img_data = response.read()
                    st.image(img_data, use_container_width=True); st.download_button("📥 Descargar", img_data, "portada.png", "image/png")
    elif sub_idx == 3:
        st.markdown("### 📝 Poesía")
        tema_poema = st.text_input("Tema del poema:")
        if st.button("Escribir Poema") and tema_poema:
            with st.spinner("Escribiendo..."):
                res = llamar_ia(modelo_seleccionado, f"Escribe un poema profundo sobre: {tema_poema}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, "Poema")

elif macro_idx == 8:
    st.markdown("#### 🏢 Gestor de Comunidad PRO")
    sub_names = ["🚨 INCIDENCIAS", "📜 NORMATIVA", "🧾 ANALIZAR PRESUPUESTOS", "📊 ACTAS", "📝 CITACIONES"]
    sub_idx = custom_tabs(sub_names, "vecinos_sub", 0)
    
    if sub_idx == 0:
        st.markdown("### 🚨 Gestión de Incidencias")
        c1, c2 = st.columns(2)
        with c1:
            tipo_inc = st.selectbox("Tipo:", ["Avería general", "Ruidos", "Filtración", "Daños", "Otro"])
            urgencia = st.selectbox("Urgencia:", ["Baja", "Media", "Alta", "Urgente"])
        with c2:
            vecino_afectado = st.text_input("¿A quién afecta? (Ej: A mi vivienda, a todo el bloque...):", value="A mi vivienda")
            ubicacion = st.text_input("Ubicación:", placeholder="Ej: Portal 2, 3ºB")
        
        descripcion_inc = st.text_area("Describe lo ocurrido:", height=150)
        
        if st.button("📝 Redactar Reclamación", type="primary") and descripcion_inc:
            with st.spinner("Redactando..."):
                prompt = f"Administrador de fincas. Reclamación formal. Tipo: {tipo_inc}. Urgencia: {urgencia}. Afecta a: {vecino_afectado}. Ubicación: {ubicacion}. Descripción: {descripcion_inc}. Incluye encabezado, cuerpo y pie solicitando solución."
                res = llamar_ia(modelo_seleccionado, prompt, st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Documento:", value=res, height=300); botones_descarga(res, f"Reclamacion_{tipo_inc.replace(' ', '_')}")
                
    elif sub_idx == 1:
        st.markdown("### 📜 Normativa y Leyes")
        consulta = campo_con_ideas("Duda legal:", "norm", "dudas legales de comunidad", modelo_seleccionado, st.session_state.api_key_gemini_manual, "Ej: ¿Puedo poner cámara en el rellano?")
        if st.button("🔎 Consultar") and consulta:
            with st.spinner("Buscando en la Ley..."):
                res = llamar_ia(modelo_seleccionado, f"Abogado experto en propiedad horizontal (España). Resuelve: {consulta}", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.markdown(res); botones_descarga(res, "Consulta_Normativa")

    elif sub_idx == 2:
        st.markdown("### 🧾 Analizador Múltiple de Presupuestos")
        st.caption("Sube varios presupuestos (fotos o PDFs) a la vez. La IA los analizará y generará una tabla comparativa final.")
        
        arch_presupuestos = st.file_uploader("Sube los presupuestos (puedes seleccionar varios):", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)
        
        if st.button("🔍 Analizar Presupuestos", type="primary") and arch_presupuestos:
            resultados_individuales = []
            progreso = st.progress(0)
            status_text = st.empty()
            
            # 1. ANÁLISIS INDIVIDUAL EN BUCLE
            for i, arch in enumerate(arch_presupuestos):
                status_text.text(f"Analizando presupuesto {i+1} de {len(arch_presupuestos)}: {arch.name}...")
                
                # Pedimos a la IA que devuelva un JSON con el nombre de la empresa
                prompt_individual = """Extrae de este presupuesto: Proveedor, Precio Total, Qué incluye y Qué NO incluye o falta. 
                IMPORTANTE: Devuelve EXCLUSIVAMENTE un objeto JSON válido con dos claves:
                1. "empresa": (Solo el nombre corto de la empresa que presenta el presupuesto, ej: "Pinturas Pepe")
                2. "analisis": (El texto detallado con los precios, lo que incluye y lo que falta).
                No añadas markdown ni texto extra, solo el JSON puro."""
                
                try:
                    if arch.name.lower().endswith(".pdf"):
                        texto_pdf = leer_archivo_universal(arch, modelo_seleccionado)
                        res = llamar_ia(modelo_seleccionado, f"{prompt_individual}\n\nTexto del PDF:\n{texto_pdf}", st.session_state.api_key_gemini_manual)
                    else:
                        res = analizar_archivo_con_gemini(arch.getvalue(), "image/jpeg", prompt_individual, st.session_state.api_key_gemini_manual)
                    
                    # Limpiamos el JSON por si la IA añade ```json
                    res_limpia = res.replace("```json", "").replace("```", "").strip()
                    
                    try:
                        data_json = json.loads(res_limpia)
                        nombre_empresa = data_json.get("empresa", arch.name)
                        texto_analisis = data_json.get("analisis", res)
                    except json.JSONDecodeError:
                        # Si la IA falla y no devuelve JSON, usamos el nombre del archivo como respaldo
                        nombre_empresa = arch.name
                        texto_analisis = res
                        
                    resultados_individuales.append({"empresa": nombre_empresa, "archivo": arch.name, "analisis": texto_analisis})
                except Exception as e:
                    resultados_individuales.append({"empresa": arch.name, "archivo": arch.name, "analisis": f"Error al procesar: {e}"})
                
                progreso.progress((i + 1) / len(arch_presupuestos))
            
            status_text.text("✅ Análisis individual completado. Generando comparativa final...")
            
            # 2. COMPARATIVA FINAL
            if len(resultados_individuales) > 0:
                texto_para_comparar = ""
                for r in resultados_individuales:
                    texto_para_comparar += f"\n--- PROVEEDOR: {r['empresa']} (Archivo: {r['archivo']}) ---\n{r['analisis']}\n"
                
                prompt_comparativa = f"""Actúa como un Administrador de Fincas. Aquí tienes el análisis de {len(resultados_individuales)} presupuestos:
{texto_para_comparar}

Genera un informe final que contenga:
1. Una TABLA COMPARATIVA en formato Markdown con las columnas: Proveedor, Precio, Qué incluye, Qué falta, y Valoración.
2. Un párrafo final de RECOMENDACIÓN explicando cuál es la mejor opción calidad/precio para la comunidad y por qué.
"""
                res_comparativa = llamar_ia(modelo_seleccionado, prompt_comparativa, st.session_state.api_key_gemini_manual)
                status_text.empty()
                progreso.empty()
                
                if "[ERROR" not in res_comparativa:
                    st.success("✅ Informe comparativo generado con éxito.")
                    st.markdown("### 🏆 Comparativa y Recomendación Final")
                    st.markdown(res_comparativa)
                    botones_descarga(res_comparativa, "Comparativa_Presupuestos", formatos_extra=True)
                    
                    st.markdown("---")
                    st.markdown("### 📂 Descargar Análisis Individuales")
                    # 3. DESPLEGABLES CON ANÁLISIS Y DESCARGA INDIVIDUAL
                    for idx, r in enumerate(resultados_individuales):
                        with st.expander(f"📄 Análisis: {r['empresa']}"):
                            st.caption(f"Archivo original: {r['archivo']}")
                            st.markdown(r['analisis'])
                            st.download_button(
                                label=f"📥 Descargar análisis de {r['empresa']}",
                                data=r['analisis'].encode('utf-8'),
                                file_name=f"Analisis_{r['empresa'].replace(' ', '_')}.txt",
                                mime="text/plain",
                                key=f"dl_ind_{idx}"
                            )
                else:
                    st.error("Error al generar la comparativa final.")

    elif sub_idx == 3:
        st.markdown("### 📊 Redactar Acta")
        puntos_tratados = st.text_area("Puntos tratados:", height=200, placeholder="1. Aprobación de cuentas...\n2. Cambio de ascensor...")
        if st.button("Generar Acta", type="primary") and puntos_tratados:
            with st.spinner("Redactando..."):
                res = llamar_ia(modelo_seleccionado, f"Acta de Junta de Vecinos formal. Puntos: {puntos_tratados}. Añade fórmulas de aprobación y firmas.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Acta:", value=res, height=300); botones_descarga(res, "Acta_Reunion")

    elif sub_idx == 4:
        st.markdown("### 📝 Convocar Junta")
        motivo = st.text_input("Motivo:", placeholder="Ej: Obras de fachada")
        fecha = st.text_input("Fecha y hora:", placeholder="Ej: 15 de Noviembre a las 19:00")
        if st.button("Generar Citación", type="primary") and motivo:
            with st.spinner("Redactando..."):
                res = llamar_ia(modelo_seleccionado, f"Citación formal para junta de vecinos. Motivo: {motivo}. Fecha: {fecha}. Incluye orden del día y firma del presidente.", st.session_state.api_key_gemini_manual)
                if "[ERROR" in res: st.error(res)
                else: st.text_area("Citación:", value=res, height=300); botones_descarga(res, "Citacion_Junta")
