import streamlit as st
import streamlit.components.v1 as components
from streamlit_markmap import markmap
import ollama, asyncio, edge_tts, pypdf, openpyxl, csv, os, re, requests, tempfile, time, qrcode, numpy as np, sqlite3, datetime, concurrent.futures, json, subprocess
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
    FFMPEG_EXE = imageio_ffmpeg.get_ffmpeg_exe()
    AudioSegment.converter = FFMPEG_EXE
    AudioSegment.ffprobe = FFMPEG_EXE
except ImportError: FFMPEG_EXE = "ffmpeg"

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
MAX_FILE_MB = 200
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
    "alex_chat_history": [], "letra_edit": "", "presu_resultados": None, "presu_comparativa": ""
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
    texto = re.sub(r'(?i)\b(Título|Intro|Verso|Coro|Puente|Outro|Producción|Instrumentos|Ritmo|Estilo|Género|Duración|Tema)\s*\d*\s*:?', '', texto)
    return re.sub(r'\s+', ' ', texto.replace('\n', '. ')).strip()

def url_segura(url: str) -> bool:
    try:
        p = urllib.parse.urlparse(url)
        if p.scheme not in ("http", "https") or not p.hostname: return False
        ip = ipaddress.ip_address(socket.gethostbyname(p.hostname))
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
    except Exception as e: st.warning(f"No pude leer '{upload.name}': {e}")
    return ""

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
    if not HAS_OPENAI or not api_key_groq: return "[ERROR] Faltan herramientas (Groq API Key)."
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
        if archivo.name.lower().endswith((".png", ".jpg", ".jpeg")):
            with st.spinner("Analizando imagen..."): ctx = analizar_archivo_con_gemini(archivo.getvalue(), "image/jpeg", "Describe esta imagen o extrae el texto (OCR).", api_key)
            st.success("✅ Imagen leída."); return f"\n[TEXTO DE IMAGEN]:\n{ctx}\n"
        texto = leer_archivo_universal(archivo, modelo)
        st.success(f"✅ Archivo '{archivo.name}' leído."); return f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"

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
    if st.session_state.get("global_context"): prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}"
    if historial_chat: prompt_final = f"Historial:\n" + "\n".join(f"{m['role'].upper()}: {m['content']}" for m in historial_chat[-10:]) + f"\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."
    
    keys = {k: _get_key(k, api_key_gemini if k=="gemini" else "") for k in ["gemini", "groq", "silicon", "mistral", "openrouter", "deepseek", "together", "cerebras"]}
    
    if keys["gemini"] and HAS_GEMINI:
        for intento in range(2):
            try:
                res = get_gemini_client(keys["gemini"]).models.generate_content(model=GEMINI_MODEL, contents=prompt_final).text
                if res: return res
            except Exception as e:
                if "429" in str(e) or "RESOURCE_EXHAUSTED" in str(e): break
                if "503" in str(e) and intento < 1: time.sleep(2); continue
                break
    urls = {"groq": "https://api.groq.com/openai/v1", "silicon": "https://api.siliconflow.cn/v1", "mistral": "https://api.mistral.ai/v1", "openrouter": "https://openrouter.ai/api/v1", "deepseek": "https://api.deepseek.com/v1", "together": "https://api.together.xyz/v1", "cerebras": "https://api.cerebras.ai/v1"}
    models = {"groq": "llama-3.3-70b-versatile", "silicon": "Qwen/Qwen2.5-7B-Instruct", "mistral": "mistral-small-latest", "openrouter": "deepseek/deepseek-r1:free", "deepseek": "deepseek-chat", "together": "meta-llama/Llama-3-70b-chat-hf", "cerebras": "llama3.1-70b"}
    for prov in ["groq", "silicon", "mistral", "openrouter", "deepseek", "together", "cerebras"]:
        if keys[prov] and HAS_OPENAI:
            try:
                client = openai.OpenAI(api_key=keys[prov], base_url=urls[prov])
                res = client.chat.completions.create(model=models[prov], messages=[{"role": "user", "content": prompt_final}]).choices[0].message.content
                if res: return res
            except: pass
    try:
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=600)
        raw = client.chat(model=modelo, messages=[{"role": "user", "content": prompt_final}], options={"temperature": 0.2})["message"]["content"]
        return re.sub(r"\n```\s*$", "", re.sub(r"^```\w*\n", "", raw)).strip()
    except Exception as e: return f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}"

def render_mermaid(mermaid_code):
    components.html(f"<div class='mermaid'>{mermaid_code}</div><script type='module'>import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';mermaid.initialize({{ startOnLoad: true, theme: 'dark', securityLevel: 'loose' }});</script>", height=600, scrolling=True)

def obtener_ruta_fuente():
    for ruta in ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/arial.ttf", "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf", os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf") if "__file__" in globals() else ""]:
        if ruta and os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo: str, contenido: str, logo_bytes: bytes = None) -> bytes:
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente(); tmp_img = None
    if logo_bytes:
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=".jpg") as tmp: tmp.write(logo_bytes); tmp_img = tmp.name
            pdf.image(tmp_img, x=40, y=50, w=130); pdf.ln(150)
        except Exception as e: st.warning(f"Error portada: {e}")
        finally:
            if tmp_img and os.path.exists(tmp_img):
                try: os.remove(tmp_img)
                except: pass
    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.add_font("XaviFont", "I", font_path, uni=True); pdf.set_font("XaviFont", size=11)
    else: pdf.set_font("Helvetica", size=11)
    pdf.set_font("XaviFont" if font_path else "Helvetica", "B", 24)
    pdf.multi_cell(0, 10, text=titulo, align="C")
    pdf.set_font("XaviFont" if font_path else "Helvetica", "I", 12)
    pdf.cell(0, 10, text="Generado por Xavilu Network", align="C", ln=True)
    pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)
    if not font_path: contenido = contenido.encode("latin-1", "replace").decode("latin-1")
    pdf.set_font("XaviFont" if font_path else "Helvetica", size=11)
    pdf.multi_cell(0, 8, text=contenido); return bytes(pdf.output())

def generar_manual_pdf() -> bytes:
    texto_manual = "¡Bienvenido a Xavilu Network Pro!\n\nEsta app conecta con Google, Meta, y modelos abiertos. Necesitas API Keys.\n\n1. XaviWork: Mapas, informes, presentaciones.\n2. Chat: Agentes y podcasts.\n3. Documentos: Resúmenes y facturas.\n4. Web: Búsquedas y YouTube.\n5. Multimedia: Imágenes, audios y CONVERSOR PRO.\n6. Estudio: Deberes, Entrevistas y PRL.\n7. Viajes: Itinerarios y Rutinas.\n8. Música: Letras, Poesía, Canciones y Suno.\n9. Comunidad: Gestor de fincas PRO (Múltiples presupuestos).\n\nCONVERSOR PRO:\nPermite convertir formatos de audio, vídeo e imagen. También permite recortar, extraer audio de vídeos, transcribir a texto, unir archivos y extraer fotogramas. Puedes editar el nombre del archivo antes de descargarlo.\n\n¡Disfruta!"
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente()
    if font_path: pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.set_font("XaviFont", "B", 20)
    else: pdf.set_font("Helvetica", "B", 20)
    pdf.multi_cell(0, 10, text="Manual de Instrucciones", align="C"); pdf.ln(10)
    if not font_path: texto_manual = texto_manual.encode('latin-1', 'replace').decode('latin-1')
    pdf.set_font("XaviFont" if font_path else "Helvetica", size=12)
    pdf.multi_cell(0, 7, text=texto_manual); return bytes(pdf.output())

def generar_word(texto: str) -> bytes:
    doc = Document(); doc.add_heading("Documento Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); return out.getvalue()

def generar_excel(texto_csv: str) -> bytes:
    wb = openpyxl.Workbook(); ws = wb.active
    for r_i, row in enumerate(csv.reader(texto_csv.replace("```csv", "").replace("```", "").strip().split("\n")), 1):
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
        logo_file = st.file_uploader("Sube imagen (PNG/JPG)", type=["png", "jpg", "jpeg"], key=f"logo_{key_u}", label_visibility="collapsed")
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
                num_limpio = re.sub(r'[^0-9]', '', st.session_state.get("whatsapp_number", ""))
                if num_limpio:
                    components.html(f"<script>window.open('https://wa.me/{num_limpio}?text={urllib.parse.quote(texto[:1500])}', '_blank');</script>", height=0); st.success("✅ Abriendo WhatsApp...")
                else: st.error("⚠️ Configura tu número en la barra lateral.")

def run_async_safe(coro):
    with concurrent.futures.ThreadPoolExecutor() as pool: return pool.submit(lambda: asyncio.run(coro)).result()

def generar_audio_sincrono(texto: str, voz: str) -> bytes:
    async def _gen():
        buf = BytesIO()
        async for chunk in edge_tts.Communicate(limpiar_texto_para_voz(texto), voz).stream():
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
            if m: voz_actual = {1: v1, 2: v2, 3: v3}.get(int(m.group(1)), v1); linea = re.sub(r"^Locutor\s*\d\s*:?", "", linea, flags=re.IGNORECASE).strip()
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
        urls = {"Reggaeton / Urbano": "SoundHelix-Song-1", "Pop / Alegre": "SoundHelix-Song-2", "Balada / Acústico": "SoundHelix-Song-3", "Rap / Hip-Hop": "SoundHelix-Song-4", "Lo-Fi Chill": "SoundHelix-Song-5", "Épico / Trap": "SoundHelix-Song-6", "Rock": "SoundHelix-Song-7", "Electrónica / EDM": "SoundHelix-Song-8"}
        r = requests.get(f"https://www.soundhelix.com/examples/mp3/{urls.get(estilo_base, 'SoundHelix-Song-1')}.mp3", headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if len(r.content) < 100_000: st.error("❌ No se pudo descargar la música."); return None
        base_audio = AudioSegment.from_file(BytesIO(r.content), format="mp3") - 14
        if len(voz_audio) > len(base_audio): base_audio = base_audio * ((len(voz_audio) // len(base_audio)) + 1)
        out = BytesIO(); base_audio[:len(voz_audio)].overlay(voz_audio).export(out, format="mp3"); return out.getvalue()
    except Exception as e: st.error(f"Error al mezclar: {e}"); return None

def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="..."):
    if f"{key_prefix}_input" not in st.session_state: st.session_state[f"{key_prefix}_input"] = ""
    def agregar_idea(idea):
        txt = st.session_state[f"{key_prefix}_input"].strip(); st.session_state[f"{key_prefix}_input"] = f"{txt}\n\n{idea}" if txt else idea
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
    return st.text_area(label, key=f"{key_prefix}_input", placeholder=placeholder, height=100)

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
    if f"{key}_txt" not in st.session_state: st.session_state[f"{key}_txt"] = ""
    def _send():
        msg = st.session_state.get(f"{key}_txt", "").strip()
        if msg: st.session_state[f"{key}_msg"] = msg; st.session_state[f"{key}_txt"] = ""
    c1, c2 = st.columns([6, 1])
    with c1: st.text_input("msg", placeholder=placeholder, key=f"{key}_txt", label_visibility="collapsed")
    with c2: st.button("➤", key=f"{key}_btn", on_click=_send, use_container_width=True, type="primary")
    return st.session_state.pop(f"{key}_msg", None)

def custom_tabs(tab_names, key_name, default_idx=0, cols_per_row=5):
    if f"tab_{key_name}" not in st.session_state: st.session_state[f"tab_{key_name}"] = default_idx
    for i in range(0, len(tab_names), cols_per_row):
        chunk = tab_names[i:i + cols_per_row]; cols = st.columns(cols_per_row)
        for j, name in enumerate(chunk):
            with cols[j]:
                if st.button(name, key=f"btn_{key_name}_{i+j}", use_container_width=True, type="primary" if st.session_state[f"tab_{key_name}"] == i+j else "secondary"):
                    st.session_state[f"tab_{key_name}"] = i+j; st.rerun()
    return st.session_state[f"tab_{key_name}"]
