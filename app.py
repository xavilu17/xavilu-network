import streamlit as st
import streamlit.components.v1 as components
from streamlit_markmap import markmap
import ollama, asyncio, edge_tts, pypdf, openpyxl, csv, os, re, requests, tempfile, time, qrcode, numpy as np, sqlite3, datetime, concurrent.futures, json, subprocess, base64
from pathlib import Path
from docx import Document
from openpyxl.styles import Font
from pptx import Presentation
from pptx.util import Inches
from io import BytesIO
import textwrap
import xml.etree.ElementTree as ET
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
try: import openai; HAS_OPENAI = True
except ImportError: HAS_OPENAI = False

st.set_page_config(page_title="Xavilu Network", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")
MAX_FILE_MB = 200
MAX_TEXTO_IA = 60_000

if os.name == "nt": asyncio.set_event_loop_policy(asyncio.WindowsSelectorEventLoopPolicy())

# ============================================================
# 🔧 HELPERS
# ============================================================
CARPETA_DESCARGAS = Path("descargas_xavilu")
CARPETA_DESCARGAS.mkdir(exist_ok=True)

def guardar_descarga(nombre_archivo: str, datos):
    sello = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    base = Path(nombre_archivo)
    destino = CARPETA_DESCARGAS / f"{base.stem}_{sello}{base.suffix}"
    try:
        if isinstance(datos, str): destino.write_text(datos, encoding="utf-8")
        else: destino.write_bytes(datos)
    except Exception: pass
    return destino

def limpiar_texto_analisis(texto: str) -> str:
    if not texto: return ""
    t = str(texto)
    t = re.sub(r"[*_`#>]+", "", t)
    t = re.sub(r"[\U0001F300-\U0001FAFF\u2600-\u27BF]", "", t)
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()

def generar_guia_pdf(version: str = "1.0") -> bytes:
    pdf = FPDF()
    pdf.set_auto_page_break(auto=True, margin=15)
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 16)
    pdf.cell(0, 10, "Guia rapida - Xavilu Network Pro", ln=True)
    pdf.ln(4)
    pdf.set_font("Helvetica", size=11)
    contenido = f"""1. QUE ES
Xavilu Network Pro es tu centro de trabajo para gestionar presupuestos,
generar prompts de musica con IA y convertir audio/video, todo desde una
misma app.

2. PRESUPUESTOS
- Crea varios presupuestos y comparalos.
- Descarga cada presupuesto de forma individual.
- Revisa el analisis individual antes de enviarlo.

3. PROMPTS PARA SUNO AI
- Rellena estilo, ambiente y estructura.
- Copia el prompt y pegalo en Suno AI.

4. CONVERSOR PRO (AUDIO/VIDEO)
- Sube tu archivo (MP3, WAV, MP4...).
- Elige formato y opciones.
- Descarga el resultado al terminar.

5. COMPARTIR
- Descarga el PDF o el analisis en .txt.
- Pulsa "Compartir por WhatsApp" y adjunta el archivo.

6. SOPORTE
Manual completo disponible en la app.
Version {version}"""
    for linea in contenido.split("\n"):
        linea = linea.strip()
        if not linea:
            pdf.ln(3)
            continue
        for fragmento in textwrap.wrap(linea, width=85, break_long_words=True, break_on_hyphens=True) or [""]:
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(w=0, h=6, text=fragmento)
    resultado = pdf.output(dest="S")
    return resultado.encode("latin-1") if isinstance(resultado, str) else bytes(resultado)

def generar_manual_sencillo_pdf() -> bytes:
    pdf = FPDF()
    pdf.add_page()
    font_path = obtener_ruta_fuente()
    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True)
        pdf.add_font("XaviFont", "B", font_path, uni=True)
        pdf.add_font("XaviFont", "I", font_path, uni=True)
        font_name = "XaviFont"
    else:
        font_name = "Helvetica"

    ancho_util = pdf.w - pdf.l_margin - pdf.r_margin

    pdf.set_font(font_name, "B", 24)
    pdf.set_text_color(255, 140, 0)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=20, text="XAVILU NETWORK PRO", align="C")

    pdf.set_font(font_name, "I", 14)
    pdf.set_text_color(0, 0, 0)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=10, text="Tu centro de trabajo y diversion todo en uno", align="C")
    pdf.ln(10)

    pdf.set_font(font_name, "", 12)
    intro = ("Imagina que tienes una oficina entera dentro de tu ordenador. Xavilu Network es una aplicacion "
             "que reune a los mejores asistentes virtuales (Inteligencia Artificial) para ayudarte con las "
             "tareas del dia a dia, los deberes, el trabajo o para crear musica. Todo desde una sola pantalla.")
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text=intro)
    pdf.ln(10)

    pdf.set_font(font_name, "B", 14)
    pdf.set_text_color(200, 0, 150)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=10, text="1. COMO ENTRAR")

    pdf.set_font(font_name, "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="Al abrir la aplicacion, lo primero que veras es una pantalla oscura que te pedira una Contrasena.")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- Adultos (Administrador): Acceso a absolutamente todo.")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- Ninos: La app escondera las herramientas delicadas y dejara solo lo divertido.")
    pdf.ln(8)

    pdf.set_font(font_name, "B", 14)
    pdf.set_text_color(200, 0, 150)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=10, text="2. EL MENU DE ARRIBA (Tus herramientas)")

    pdf.set_font(font_name, "", 11)
    pdf.set_text_color(0, 0, 0)

    herramientas = [
        ("XaviWork", "Tu bloc de notas inteligente. Crea mapas, resumenes y esquemas."),
        ("Chat", "Habla con la IA como si fuera una persona. Crea guiones de podcast."),
        ("Documentos", "Sube fotos de facturas o papeles y la IA te los resume o traduce."),
        ("Web", "Busca en internet o pide que resuman videos de YouTube."),
        ("Multimedia", "Crea imagenes, convierte textos a voz o cambia formatos de video."),
        ("Estudio", "Un profesor particular. Le escribes dudas y te las resuelve paso a paso."),
        ("Viajes", "Organiza tus viajes dia por dia y te inventa recetas con lo que tengas."),
        ("Musica", "Escribe letras de canciones, poemas y crea la musica para cantarlas."),
        ("Gestor Vecinos", "(Solo Admin). Para papeles del edificio y presupuestos.")
    ]

    for nombre, desc in herramientas:
        pdf.set_font(font_name, "B", 11)
        pdf.set_x(pdf.l_margin)
        pdf.cell(w=0, h=6, text=nombre, ln=True)
        pdf.set_font(font_name, "", 11)
        pdf.set_x(pdf.l_margin + 5)
        pdf.multi_cell(w=ancho_util - 5, h=6, text=desc)
        pdf.ln(2)

    pdf.ln(5)

    pdf.set_font(font_name, "B", 14)
    pdf.set_text_color(200, 0, 150)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=10, text="3. COMO GUARDAR LO QUE CREO")

    pdf.set_font(font_name, "", 12)
    pdf.set_text_color(0, 0, 0)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="Cuando pides a la IA que escriba algo, justo debajo apareceran unos botones de colores para descargarlo:")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- TXT: Solo el texto.")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- PDF: Un documento listo para imprimir.")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- Word / Excel: Para abrirlo en los programas de toda la vida.")

    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="- WhatsApp: Para enviartelo al movil.")

    pdf.ln(10)
    pdf.set_font(font_name, "I", 12)
    pdf.set_text_color(100, 100, 100)
    pdf.set_x(pdf.l_margin)
    pdf.multi_cell(w=ancho_util, h=7, text="Consejo: No tengas miedo, no puedes romper nada. Cuanto mas detalle le des a la IA, mejor sera el resultado.")

    resultado = pdf.output(dest="S")
    return resultado.encode("latin-1") if isinstance(resultado, str) else bytes(resultado)

import urllib.parse as _urlparse
MENSAJES_WA = {
    "guia": "Hola, te comparto la guia rapida de Xavilu Network Pro. Te adjunto el PDF. Cualquier duda, aqui estoy.",
    "presupuesto": "Hola, te envio la comparativa de presupuestos de Xavilu Network Pro. Te adjunto el detalle. Cualquier duda, aqui estoy.",
    "analisis": "Hola, te envio el analisis individual de {nombre}. Te adjunto el archivo de texto. Cualquier duda, aqui estoy.",
    "manual": "Hola, te comparto el manual de Xavilu Network Pro. Te adjunto el PDF. Cualquier duda, aqui estoy.",
}
def wa_url(mensaje: str) -> str: return f"https://wa.me/?text={_urlparse.quote(mensaje)}"

# ============================================================
# BASE DE DATOS Y CONFIG
# ============================================================
def init_db():
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        conn.execute("CREATE TABLE IF NOT EXISTS historial (id INTEGER PRIMARY KEY AUTOINCREMENT, fecha TEXT, tipo TEXT, contenido TEXT)")
        conn.execute("""CREATE TABLE IF NOT EXISTS incidencias (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            creada TEXT NOT NULL,
            tipo TEXT NOT NULL,
            urgencia TEXT NOT NULL,
            afectado TEXT,
            ubicacion TEXT,
            descripcion TEXT NOT NULL,
            responsable TEXT,
            fecha_limite TEXT,
            estado TEXT NOT NULL DEFAULT 'Abierta',
            actualizada TEXT NOT NULL
        )""")

def guardar_incidencia(tipo, urgencia, afectado, ubicacion, descripcion, responsable="", fecha_limite=""):
    ahora = datetime.datetime.now().isoformat(timespec="minutes")
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        cur = conn.execute(
            "INSERT INTO incidencias (creada,tipo,urgencia,afectado,ubicacion,descripcion,responsable,fecha_limite,estado,actualizada) VALUES (?,?,?,?,?,?,?,?,?,?)",
            (ahora, tipo, urgencia, afectado, ubicacion, descripcion, responsable, fecha_limite, "Abierta", ahora)
        )
        return cur.lastrowid

def listar_incidencias():
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        return conn.execute("SELECT id,creada,tipo,urgencia,afectado,ubicacion,descripcion,responsable,fecha_limite,estado,actualizada FROM incidencias ORDER BY CASE urgencia WHEN 'Urgente' THEN 0 WHEN 'Alta' THEN 1 WHEN 'Media' THEN 2 ELSE 3 END, id DESC").fetchall()

def actualizar_incidencia(incidencia_id, estado, responsable):
    if estado not in {"Abierta", "En curso", "Resuelta", "Cerrada"}: return False
    ahora = datetime.datetime.now().isoformat(timespec="minutes")
    with sqlite3.connect("xavilu_memoria.db", timeout=10) as conn:
        cur = conn.execute("UPDATE incidencias SET estado=?, responsable=?, actualizada=? WHERE id=?", (estado, responsable, ahora, incidencia_id))
        return cur.rowcount == 1
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

# ============================================================
# 🔐 SISTEMA DE LOGIN Y ROLES
# ============================================================
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

def check_login():
    if "rol" not in st.session_state:
        st.markdown("<h2 style='text-align: center; margin-top: 50px;'>🔐 Acceso Restringido</h2>", unsafe_allow_html=True)
        st.markdown("<p style='text-align: center;'>Introduce tu contraseña para entrar a Xavilu Network.</p>", unsafe_allow_html=True)

        c1, c2, c3 = st.columns([1, 1, 1])
        with c2:
            pwd = st.text_input("Contraseña:", type="password", key="login_pwd")
            if st.button("Entrar", use_container_width=True, type="primary"):
                try:
                    pwd_admin = st.secrets.get("PASSWORD", "admin")
                    pwd_ninos = st.secrets.get("PASSWORD_NINOS", "ninos")
                except Exception:
                    pwd_admin = "admin"
                    pwd_ninos = "ninos"

                if pwd and pwd == pwd_admin:
                    st.session_state["rol"] = "admin"
                    st.rerun()
                elif pwd and pwd == pwd_ninos:
                    st.session_state["rol"] = "ninos"
                    st.rerun()
                elif pwd:
                    st.error("❌ Contraseña incorrecta.")
        st.stop()

check_login()

# ============================================================
# FUNCIONES CORE
# ============================================================
def mensaje_error_amigable(texto: str) -> str:
    if not texto or "[ERROR" not in texto:
        return texto

    t = texto.upper()
    mapa = [
        ("API KEY",           "🔑 Revisa tu API Key: no es válida o no está configurada en el panel lateral."),
        ("API_KEY",           "🔑 Revisa tu API Key: no es válida o no está configurada en el panel lateral."),
        ("PERMISSION",        "🚫 La API Key no tiene permisos para este modelo. Revisa tu cuenta."),
        ("RESOURCE_EXHAUSTED","💳 Se agotó la cuota (429). Cambia de proveedor en el panel lateral."),
        ("QUOTA",             "💳 Se agotó la cuota gratuita. Espera unos minutos o usa otro proveedor."),
        ("429",               "💳 Demasiadas peticiones seguidas (429). Espera 1 minuto y reintenta."),
        ("503",               "🛠️ El servidor de IA está saturado (503). Reintenta en unos segundos."),
        ("TIMEOUT",           "⏱️ La IA tardó demasiado. Prueba con un texto más corto."),
        ("CONNECTION",        "🌐 Sin conexión con el servidor de IA. Revisa tu internet."),
        ("OLLAMA",            "🖥️ Ollama no responde. ¿Está arrancado en localhost:11434?"),
        ("MEMORY",            "📦 El archivo es demasiado grande. Prueba con uno más pequeño."),
        ("TOKEN",             "📏 El texto supera el límite del modelo. Resúmelo o divídelo."),
        ("GEMINI",            "🤖 Gemini falló. Revisa la API Key o cambia de proveedor."),
        ("404",               "🔍 El modelo de Gemini no está disponible en tu cuenta."),
    ]
    for clave, msg in mapa:
        if clave in t:
            return f"{msg}\n\n_Detalle técnico: {texto[:400]}_"
    return f"⚠️ Algo salió mal con la IA.\n\n_Detalle técnico: {texto[:400]}_"

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

def buscar_legislacion_boe(consulta: str, max_normas: int = 3):
    """Recupera normativa consolidada vigente desde la API oficial del BOE."""
    api = "https://www.boe.es/datosabiertos/api/legislacion-consolidada"
    stop = {"para", "como", "sobre", "puede", "puedo", "debe", "deben", "cuando", "donde", "cual", "cuáles", "comunidad", "propietarios", "ley", "legal", "normativa", "una", "unos", "unas", "los", "las", "del", "por", "que", "qué", "con", "sin", "esta", "este", "hay", "ser", "son", "tiene", "tienen", "hacer", "poner", "tener"}
    tokens = [t for t in re.findall(r"[a-záéíóúüñ]{4,}", consulta.lower()) if t not in stop][:5]
    hoy = datetime.date.today().strftime("%Y%m%d")
    normas = []
    vistos = set()

    # La LPH es la norma marco de esta sección; se obtiene siempre en su última versión consolidada.
    ids = ["BOE-A-1960-10906"]
    if tokens:
        query_string = " and ".join(f'texto:"{t}"' for t in tokens[:3])
        query_json = json.dumps({"query": {"query_string": {"query": query_string}}, "sort": [{"fecha_actualizacion": "desc"}]}, ensure_ascii=False)
        try:
            respuesta = requests.get(api, params={"query": query_json, "limit": max_normas + 2}, headers={"Accept": "application/json"}, timeout=(5, 20))
            respuesta.raise_for_status()
            datos = respuesta.json().get("data", [])
            if isinstance(datos, dict): datos = datos.get("item", []) if isinstance(datos.get("item", []), list) else [datos.get("item", {})]
            for item in datos if isinstance(datos, list) else []:
                ident = item.get("identificador")
                if ident and item.get("vigencia_agotada", "N") != "S" and item.get("estatus_derogacion", "N") != "S" and item.get("estatus_anulacion", "N") != "S" and ident not in ids:
                    ids.append(ident)
                if len(ids) >= max_normas: break
        except Exception:
            # Si falla la búsqueda temática, aún se intenta la LPH consolidada.
            pass

    for ident in ids[:max_normas]:
        try:
            respuesta = requests.get(f"{api}/id/{ident}", headers={"Accept": "application/xml"}, timeout=(5, 20))
            respuesta.raise_for_status()
            root = ET.fromstring(respuesta.content)
            data = root.find("data")
            if data is None: continue
            meta = data.find("metadatos")
            if meta is None: continue
            def meta_text(tag):
                node = meta.find(tag)
                return (node.text or "").strip() if node is not None else ""
            titulo = meta_text("titulo") or ident
            actualizado = meta_text("fecha_actualizacion")[:8]
            url = meta_text("url_html_consolidada") or f"{api}/id/{ident}/texto"
            texto_node = data.find("texto")
            bloques = []
            if texto_node is not None:
                for bloque in texto_node.findall("bloque"):
                    versiones = bloque.findall("version")
                    aplicables = [v for v in versiones if (v.get("fecha_vigencia") or v.get("fecha_publicacion") or "00000000") <= hoy]
                    if not aplicables: continue
                    version = max(aplicables, key=lambda v: (v.get("fecha_vigencia") or v.get("fecha_publicacion") or "00000000"))
                    fragmento = " ".join(x.strip() for x in version.itertext() if x and x.strip())
                    encabezado = bloque.get("titulo", "")
                    if fragmento: bloques.append((encabezado, fragmento))
            if not bloques: continue
            scored = []
            for titulo_bloque, fragmento in bloques:
                haystack = f"{titulo_bloque} {fragmento}".lower()
                score = sum(1 for t in tokens if t in haystack)
                if score: scored.append((score, titulo_bloque, fragmento))
            seleccion = sorted(scored, key=lambda x: x[0], reverse=True)[:12]
            if not seleccion: seleccion = [(0, h, t) for h, t in bloques[:8]]
            texto_norma = "\n".join(f"{h}: {t}" if h else t for _, h, t in seleccion)
            if texto_norma.strip():
                normas.append({"id": ident, "titulo": titulo, "actualizado": actualizado, "url": url, "texto": texto_norma[:16000]})
                vistos.add(ident)
        except Exception:
            continue
    return normas

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
    with st.expander("📎 Adjuntar archivo(s) (Opcional)"):
        archivos = st.file_uploader("Sube PDF, Word, TXT o Imagen (puedes seleccionar varios)", type=["pdf", "txt", "docx", "png", "jpg", "jpeg"], accept_multiple_files=True, key=f"adj_{key_suffix}")
        if not archivos: return ""
        
        contexto_total = ""
        for i, archivo in enumerate(archivos):
            if archivo.name.lower().endswith((".png", ".jpg", ".jpeg")):
                with st.spinner(f"Analizando imagen {i+1}/{len(archivos)}..."):
                    mime = "image/png" if archivo.name.lower().endswith(".png") else "image/jpeg"
                    ctx = analizar_archivo_con_gemini(archivo.getvalue(), mime, "Describe esta imagen o extrae el texto (OCR).", api_key, modelo_local=modelo)
                    contexto_total += f"\n[TEXTO DE IMAGEN {i+1} - {archivo.name}]:\n{ctx}\n"
            else:
                texto = leer_archivo_universal(archivo, modelo)
                contexto_total += f"\n[CONTENIDO DE '{archivo.name}']:\n{texto[:10000]}\n"
        
        st.success(f"✅ {len(archivos)} archivo(s) leído(s) correctamente.")
        return contexto_total

# ============================================================
# ⚡ CONEXIÓN DIRECTA A GEMINI (Modelos ACTUALIZADOS)
# ============================================================
def llamar_gemini_rest(prompt: str, api_key: str, archivo_bytes: bytes = None, mime_type: str = None) -> str:
    if not api_key or len(api_key.strip()) < 15:
        return "[ERROR GEMINI] No hay API Key configurada o es demasiado corta. Añádela en el panel lateral o en st.secrets como GEMINI_API_KEY."

    api_key = api_key.strip()

    combinaciones = [
        ("v1beta", "gemini-3.8-flash"),
        ("v1beta", "gemini-3.6-flash"),
        ("v1beta", "gemini-3.5-flash"),
    ]

    errores_detalle = []

    for version, modelo in combinaciones:
        url = f"https://generativelanguage.googleapis.com/{version}/models/{modelo}:generateContent"
        headers = {'Content-Type': 'application/json', 'x-goog-api-key': api_key}

        if archivo_bytes and mime_type:
            b64_data = base64.b64encode(archivo_bytes).decode('utf-8')
            contents = [{
                "parts": [
                    {"text": prompt},
                    {"inline_data": {"mime_type": mime_type, "data": b64_data}}
                ]
            }]
        else:
            contents = [{"parts": [{"text": prompt}]}]

        payload = {"contents": contents}

        try:
            response = requests.post(url, headers=headers, json=payload, timeout=60)
            if response.status_code == 200:
                res_json = response.json()
                try:
                    return res_json['candidates'][0]['content']['parts'][0]['text']
                except (KeyError, IndexError):
                    errores_detalle.append(f"{version}/{modelo}: HTTP 200 sin texto -> {str(res_json)[:150]}")
                    continue
            else:
                txt_err = response.text[:250]
                errores_detalle.append(f"{version}/{modelo}: HTTP {response.status_code} -> {txt_err}")
                if response.status_code in (401, 403) or "API_KEY_INVALID" in txt_err or "API key not valid" in txt_err:
                    return f"[ERROR GEMINI] API Key inválida o sin permisos. HTTP {response.status_code}. Detalle: {txt_err}"
                if response.status_code == 429:
                    return f"[ERROR GEMINI] Cuota agotada (HTTP 429). Espera un poco o usa otro proveedor. Detalle: {txt_err}"
        except requests.exceptions.Timeout:
            errores_detalle.append(f"{version}/{modelo}: Timeout (60s)")
        except requests.exceptions.ConnectionError as e:
            errores_detalle.append(f"{version}/{modelo}: Error de conexión -> {str(e)[:120]}")
        except Exception as e:
            errores_detalle.append(f"{version}/{modelo}: {type(e).__name__} -> {str(e)[:120]}")

    detalle = " | ".join(errores_detalle[-3:]) if errores_detalle else "Sin detalles"
    return f"[ERROR GEMINI] Todos los intentos fallaron. Detalles: {detalle}"

def analizar_archivo_con_gemini(archivo_bytes: bytes, mime_type: str, prompt: str, api_key_gemini: str, modelo_local: str = "llama3.2:latest") -> str:
    """Analiza adjuntos con Gemini y usa proveedores de texto/Ollama como respaldo para PDF legible."""
    error_gemini = "[ERROR] No se pudo usar Gemini."
    if api_key_gemini:
        res = llamar_gemini_rest(prompt, api_key_gemini, archivo_bytes, mime_type)
        if "[ERROR" not in res:
            return res
        error_gemini = res

    if mime_type == "application/pdf":
        try:
            paginas = pypdf.PdfReader(BytesIO(archivo_bytes)).pages
            texto = "\n\n".join((pagina.extract_text() or "") for pagina in paginas).strip()
        except Exception as e:
            return f"{error_gemini}\nNo se pudo extraer texto del PDF: {e}"

        if texto:
            prompt_respaldo = (
                f"{prompt}\n\nEl proveedor principal no pudo procesar el archivo. "
                "Analiza únicamente el texto extraído a continuación; si falta información, indícalo.\n\n"
                f"TEXTO DEL PDF:\n{texto[:MAX_TEXTO_IA]}"
            )
            return llamar_ia(modelo_local, prompt_respaldo, api_key_gemini, intentar_gemini=False, incluir_contexto_global=False)
        return f"{error_gemini}\nEl PDF no contiene texto seleccionable. Si es un escaneo, hace falta OCR o un modelo con visión."

    if mime_type.startswith("image/"):
        return f"{error_gemini}\nEl respaldo automático de imágenes requiere un modelo con visión; los modelos locales configurados actualmente son de texto."
    return error_gemini

def _get_key(nombre: str, arg_directo: str = "") -> str:
    manual = st.session_state.get(f"api_key_{nombre.lower()}_manual", "").strip()
    if manual: return manual
    try:
        secret = st.secrets.get(f"{nombre.upper()}_API_KEY", "")
        if secret: return secret
    except: pass
    return arg_directo or ""

def llamar_ia(modelo, prompt, api_key_gemini="", historial_chat=None, intentar_gemini=True, incluir_contexto_global=True):
    prompt_final = prompt[:MAX_TEXTO_IA]
    if incluir_contexto_global and st.session_state.get("global_context"):
        prompt_final = f"[ARCHIVO ADJUNTO: {st.session_state.global_file_name}]\nCONTENIDO:\n{st.session_state.global_context}\n\n[INSTRUCCIÓN]:\n{prompt_final}"
    if historial_chat:
        prompt_final = f"Historial:\n" + "\n".join(f"{m['role'].upper()}: {m['content']}" for m in historial_chat[-10:]) + f"\n\nNuevo mensaje:\n{prompt_final}\n\nResponde."

    keys = {k: _get_key(k, api_key_gemini if k=="gemini" else "") for k in ["gemini", "groq", "silicon", "mistral", "openrouter", "deepseek", "together", "cerebras"]}

    if intentar_gemini and keys["gemini"]:
        res = llamar_gemini_rest(prompt_final, keys["gemini"])
        if "[ERROR" not in res:
            return res

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
    except Exception as e:
        return mensaje_error_amigable(f"[ERROR CRÍTICO] Todas las IAs fallaron. Detalle Ollama: {e}")

def render_mermaid(mermaid_code):
    components.html(f"<div class='mermaid'>{mermaid_code}</div><script type='module'>import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';mermaid.initialize({{ startOnLoad: true, theme: 'dark', securityLevel: 'loose' }});</script>", height=600, scrolling=True)

def obtener_ruta_fuente():
    for ruta in ["/usr/share/fonts/truetype/dejavu/DejaVuSans.ttf", "C:/Windows/Fonts/arial.ttf", "/Library/Fonts/Arial.ttf", "/System/Library/Fonts/Supplemental/Arial.ttf", os.path.join(os.path.dirname(__file__), "DejaVuSans.ttf") if "__file__" in globals() else ""]:
        if ruta and os.path.exists(ruta): return ruta
    return None

def generar_pdf_pro(titulo: str, contenido: str, logo_bytes: bytes = None) -> bytes:
    pdf = FPDF(); pdf.set_auto_page_break(auto=True, margin=18); pdf.add_page(); font_path = obtener_ruta_fuente(); tmp_img = None

    if logo_bytes:
        try:
            imagen_logo = Image.open(BytesIO(logo_bytes))
            formato_logo = (imagen_logo.format or "PNG").lower()
            extension_logo = ".jpg" if formato_logo in ("jpeg", "jpg") else ".png"
            with tempfile.NamedTemporaryFile(delete=False, suffix=extension_logo) as tmp: tmp.write(logo_bytes); tmp_img = tmp.name
            pdf.image(tmp_img, x=10, y=10, w=32)
            pdf.set_y(45)
        except Exception as e: st.warning(f"Error portada: {e}")
        finally:
            if tmp_img and os.path.exists(tmp_img):
                try: os.remove(tmp_img)
                except: pass
    else:
        pdf.set_y(20)

    if font_path:
        pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.add_font("XaviFont", "I", font_path, uni=True); pdf.set_font("XaviFont", size=11)
    else: pdf.set_font("Helvetica", size=11)

    pdf.set_font("XaviFont" if font_path else "Helvetica", "B", 20)
    pdf.multi_cell(w=0, h=10, text=titulo, align="C")
    pdf.set_font("XaviFont" if font_path else "Helvetica", "I", 10)
    pdf.cell(w=0, h=10, text="Generado por Xavilu Network", align="C", ln=True)
    pdf.set_draw_color(0, 240, 255); pdf.line(10, pdf.get_y() + 5, 200, pdf.get_y() + 5); pdf.ln(15)

    contenido_limpio = limpiar_texto_analisis(contenido)
    if not font_path: contenido_limpio = contenido_limpio.encode("latin-1", "replace").decode("latin-1")
    pdf.set_font("XaviFont" if font_path else "Helvetica", size=11)

    for linea in contenido_limpio.split("\n"):
        linea = linea.strip()
        if not linea:
            pdf.ln(3)
            continue
        for fragmento in textwrap.wrap(linea, width=85, break_long_words=True, break_on_hyphens=True) or [""]:
            pdf.set_x(pdf.l_margin)
            pdf.multi_cell(w=0, h=6, text=fragmento)

    resultado = pdf.output(dest="S")
    return resultado.encode("latin-1") if isinstance(resultado, str) else bytes(resultado)

def generar_manual_pdf() -> bytes:
    texto_manual = "¡Bienvenido a Xavilu Network Pro!\n\nEsta app conecta con Google, Meta, y modelos abiertos. Necesitas API Keys.\n\n1. XaviWork: Mapas, informes, presentaciones.\n2. Chat: Agentes y podcasts.\n3. Documentos: Resúmenes y facturas.\n4. Web: Búsquedas y YouTube.\n5. Multimedia: Imágenes, audios y CONVERSOR PRO.\n6. Estudio: Deberes, Entrevistas y PRL.\n7. Viajes: Itinerarios y Rutinas.\n8. Música: Letras, Poesía, Canciones y Suno.\n9. Comunidad: Gestor de fincas PRO (Múltiples presupuestos).\n\nCONVERSOR PRO:\nPermite convertir formatos de audio, vídeo e imagen. También permite recortar, extraer audio de vídeos, transcribir a texto, unir archivos y extraer fotogramas. Puedes editar el nombre del archivo antes de descargarlo.\n\n¡Disfruta!"
    pdf = FPDF(); pdf.add_page(); font_path = obtener_ruta_fuente()
    if font_path: pdf.add_font("XaviFont", "", font_path, uni=True); pdf.add_font("XaviFont", "B", font_path, uni=True); pdf.set_font("XaviFont", "B", 20)
    else: pdf.set_font("Helvetica", "B", 20)
    pdf.multi_cell(w=0, h=10, text="Manual de Instrucciones", align="C"); pdf.ln(10)
    if not font_path: texto_manual = texto_manual.encode('latin-1', 'replace').decode('latin-1')
    pdf.set_font("XaviFont" if font_path else "Helvetica", size=12)
    pdf.multi_cell(w=0, h=7, text=texto_manual); return bytes(pdf.output())

def generar_word(texto: str, logo_bytes: bytes = None) -> bytes:
    doc = Document()
    if logo_bytes:
        try:
            parrafo_logo = doc.add_paragraph()
            parrafo_logo.alignment = 1
            parrafo_logo.add_run().add_picture(BytesIO(logo_bytes), width=Inches(1.8))
        except Exception:
            pass
    doc.add_heading("Documento Xavilu", 0)
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
    logo_general = st.session_state.get("logo_empresa_upload")
    logo_bytes = logo_general.getvalue() if logo_general else None
    with st.expander("🏷️ Logo del informe (opcional)"):
        st.caption("Se aplicará el logo general configurado en el menú lateral. Puedes sustituirlo solo para este PDF.")
        logo_file = st.file_uploader("Elegir otro logo para este PDF", type=["png", "jpg", "jpeg"], key=f"logo_{key_u}", label_visibility="collapsed")
        if logo_file: logo_bytes = logo_file.getvalue()
        
    if formatos_extra:
        c1, c2, c3, c4, c5 = st.columns(5)
        with c1:
            _txt = texto.encode("utf-8")
            if st.download_button("📥 TXT", _txt, f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.txt", _txt)
        with c2:
            _pdf = generar_pdf_pro(nombre_limpio, texto, logo_bytes)
            if st.download_button("📄 PDF", _pdf, f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.pdf", _pdf)
        with c3:
            _docx = generar_word(texto, logo_bytes)
            if st.download_button("📝 Word", _docx, f"{nombre_limpio}.docx", "application/vnd.openxmlformats-officedocument.wordprocessingml.document", key=f"wd_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.docx", _docx)
        with c4:
            _xlsx = generar_excel(texto)
            if st.download_button("📊 Excel", _xlsx, f"{nombre_limpio}.xlsx", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", key=f"ex_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.xlsx", _xlsx)
        with c5:
            if st.button("📲 WhatsApp", key=f"wa_{key_u}", use_container_width=True):
                num_limpio = re.sub(r'[^0-9]', '', st.session_state.get("whatsapp_number", ""))
                if num_limpio:
                    components.html(f"<script>window.open('https://wa.me/{num_limpio}?text={urllib.parse.quote(texto[:1500])}', '_blank');</script>", height=0)
                    st.success("✅ Abriendo WhatsApp...")
                else:
                    st.error("⚠️ Configura tu número en la barra lateral.")
    else:
        c1, c2, c3 = st.columns(3)
        with c1:
            _txt = texto.encode("utf-8")
            if st.download_button("📥 TXT", _txt, f"{nombre_limpio}.txt", "text/plain", key=f"tx_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.txt", _txt)
        with c2:
            _pdf = generar_pdf_pro(nombre_limpio, texto, logo_bytes)
            if st.download_button("📄 PDF", _pdf, f"{nombre_limpio}.pdf", "application/pdf", key=f"pd_{key_u}", use_container_width=True):
                guardar_descarga(f"{nombre_limpio}.pdf", _pdf)
        with c3:
            if st.button("📲 Enviar a WhatsApp", key=f"wa_{key_u}", use_container_width=True):
                num_limpio = re.sub(r'[^0-9]', '', st.session_state.get("whatsapp_number", ""))
                if num_limpio:
                    components.html(f"<script>window.open('https://wa.me/{num_limpio}?text={urllib.parse.quote(texto[:1500])}', '_blank');</script>", height=0)
                    st.success("✅ Abriendo WhatsApp...")
                else:
                    st.error("⚠️ Configura tu número en la barra lateral.")

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

        mapeo_bases = {
            "Bachata": "SoundHelix-Song-3", "Reggaeton": "SoundHelix-Song-1", "Pop": "SoundHelix-Song-2",
            "Rock": "SoundHelix-Song-7", "Rap": "SoundHelix-Song-4", "Cumbia": "SoundHelix-Song-2",
            "Balada": "SoundHelix-Song-3", "Electrónica": "SoundHelix-Song-8", "Rumba": "SoundHelix-Song-2",
            "Flamenco": "SoundHelix-Song-3", "Salsa": "SoundHelix-Song-2", "Merengue": "SoundHelix-Song-2",
            "Tango": "SoundHelix-Song-3", "Reggae": "SoundHelix-Song-5", "Metal": "SoundHelix-Song-7",
            "Jazz": "SoundHelix-Song-5", "Blues": "SoundHelix-Song-5", "Country": "SoundHelix-Song-7",
            "Clásica": "SoundHelix-Song-3", "Trap": "SoundHelix-Song-6", "Techno": "SoundHelix-Song-8",
            "Indie": "SoundHelix-Song-8", "Punk": "SoundHelix-Song-7", "Funk": "SoundHelix-Song-8",
            "Bossa Nova": "SoundHelix-Song-5"
        }
        track_elegido = mapeo_bases.get(estilo_base, "SoundHelix-Song-1")

        r = requests.get(f"https://www.soundhelix.com/examples/mp3/{track_elegido}.mp3", headers={"User-Agent": "Mozilla/5.0"}, timeout=15)
        if len(r.content) < 100_000: st.error("❌ No se pudo descargar la música."); return None
        base_audio = AudioSegment.from_file(BytesIO(r.content), format="mp3") - 14
        if len(voz_audio) > len(base_audio): base_audio = base_audio * ((len(voz_audio) // len(base_audio)) + 1)
        out = BytesIO(); base_audio[:len(voz_audio)].overlay(voz_audio).export(out, format="mp3"); return out.getvalue()
    except Exception as e: st.error(f"Error al mezclar: {e}"); return None

def campo_con_ideas(label, key_prefix, contexto_ia, modelo, api_key, placeholder="..."):
    if f"{key_prefix}_input" not in st.session_state: st.session_state[f"{key_prefix}_input"] = ""
    if f"{key_prefix}_sugs" not in st.session_state: st.session_state[f"{key_prefix}_sugs"] = []

    texto = st.text_area(label, key=f"{key_prefix}_input", placeholder=placeholder, height=100)

    c1, c2, c3 = st.columns(3)
    with c1:
        if st.button("💡 3 Ideas IA", key=f"{key_prefix}_btn_ideas", use_container_width=True):
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo, f"Dame 3 ideas creativas y concisas para {contexto_ia}. Solo lista por líneas.", api_key)
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else:
                    st.session_state[f"{key_prefix}_sugs"] = [re.sub(r'^[\-\*\d\.]+\s*', '', l).strip() for l in res.split("\n") if len(l.strip()) > 5][:3]
    with c2:
        if st.button("🧬 Mezclar con IA", key=f"{key_prefix}_btn_mix", use_container_width=True):
            if not texto.strip() and not st.session_state[f"{key_prefix}_sugs"]:
                st.warning("Escribe algo o genera ideas IA primero.")
            else:
                with st.spinner("Fusionando ideas..."):
                    ideas_ia = ", ".join(st.session_state[f"{key_prefix}_sugs"]) if st.session_state[f"{key_prefix}_sugs"] else "Ninguna (usa tu creatividad)"
                    prompt_mix = f"Mi idea base es: '{texto}'.\nIdeas de la IA: {ideas_ia}.\n\nFusiona todo en un único prompt o idea final mucho más detallada, profesional y creativa. Devuelve SOLO el texto final, sin introducciones ni saludos."
                    res = llamar_ia(modelo, prompt_mix, api_key)
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else:
                        st.session_state[f"{key_prefix}_input"] = res
                        st.session_state[f"{key_prefix}_sugs"] = []
                        st.rerun()
    with c3:
        if st.button("✍️ Usar mi idea", key=f"{key_prefix}_btn_use", use_container_width=True):
            st.session_state[f"{key_prefix}_sugs"] = []
            st.rerun()

    if st.session_state[f"{key_prefix}_sugs"]:
        st.caption("👉 Haz clic en una idea para reemplazar tu texto, o pulsa 'Mezclar' para fusionarla:")
        cols = st.columns(len(st.session_state[f"{key_prefix}_sugs"]))
        for i, idea in enumerate(st.session_state[f"{key_prefix}_sugs"]):
            with cols[i]:
                if st.button(idea, key=f"{key_prefix}_idea_{i}", use_container_width=True):
                    st.session_state[f"{key_prefix}_input"] = idea
                    st.session_state[f"{key_prefix}_sugs"] = []
                    st.rerun()

    return st.session_state[f"{key_prefix}_input"]

def boton_superpoder_borrador(texto_base, instruccion, key_suffix, modelo, api_key, state_key, widget_key=None):
    if st.button(f"🚀 {instruccion}", key=f"super_{key_suffix}", type="secondary", use_container_width=True):
        if not texto_base: return
        with st.spinner("Aplicando..."):
            res = llamar_ia(modelo, f"{instruccion}. Devuelve SOLO el texto mejorado:\n\n{texto_base}", api_key)
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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

if "sidebar_visible" not in st.session_state: st.session_state.sidebar_visible = True

if not st.session_state.sidebar_visible:
    st.markdown("<style>[data-testid=\"collapsedControl\"] {display: none !important;}</style>", unsafe_allow_html=True)
    c1, c2, c3 = st.columns([1, 2, 1])
    with c2:
        if st.button("⛶ MOSTRAR MENÚ Y KEYS", type="primary", use_container_width=True): st.session_state.sidebar_visible = True; st.rerun()

st.markdown("<div style='margin-top: 40px;'></div>", unsafe_allow_html=True)
st.markdown("<h1 style='margin-bottom: 0;'>XAVILU NETWORK</h1>", unsafe_allow_html=True)

if st.session_state.sidebar_visible:
    with st.sidebar:
        st.markdown(f"👤 **Rol:** `{st.session_state.get('rol', 'N/A').upper()}`")
        if st.button("🚪 Cerrar Sesión", use_container_width=True):
            del st.session_state["rol"]
            st.rerun()
        st.markdown("---")

        if st.session_state.get("rol") == "admin":
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

            # Botón para probar la misma conexión Gemini que usa el resto de la app.
            if st.button("🧪 Probar conexión Gemini", use_container_width=True, key="btn_test_gemini"):
                key_test = _get_key("gemini")
                if not key_test:
                    st.error("⚠️ No hay API Key de Gemini configurada.")
                else:
                    with st.spinner("Probando conexión..."):
                        resultado_test = llamar_gemini_rest("Di solo: OK", key_test)
                        if resultado_test.startswith("[ERROR"):
                            st.error(resultado_test)
                        else:
                            st.success(f"✅ Conexión Gemini OK. Respuesta: {resultado_test[:80]}")
            st.markdown("---")
            st.session_state.whatsapp_number = st.text_input("📲 WhatsApp:", value=st.session_state.whatsapp_number)
            st.file_uploader("🏷️ Logo para informes PDF (opcional)", type=["png", "jpg", "jpeg"], key="logo_empresa_upload", help="Se usará automáticamente en los PDF que descargues. Puedes cambiarlo o quitarlo cuando quieras.")
            st.markdown("---")

            _manual_bytes = generar_manual_pdf()
            st.download_button("📖 Descargar Manual (PDF)", _manual_bytes, "Manual_Xavilu.pdf", "application/pdf", use_container_width=True)

            _guia_bytes = generar_guia_pdf(version="1.0")
            st.download_button("⚡ Descargar Guía Rápida (PDF)", _guia_bytes, "Guia_Rapida_Xavilu.pdf", "application/pdf", use_container_width=True, key="dl_guia_pdf")

            _manual_sencillo_bytes = generar_manual_sencillo_pdf()
            st.download_button("📘 Manual para Principiantes (PDF)", _manual_sencillo_bytes, "Manual_Sencillo_Xavilu.pdf", "application/pdf", use_container_width=True, key="dl_manual_sencillo")

            st.link_button("💬 Compartir guía por WhatsApp", wa_url(MENSAJES_WA["guia"]), use_container_width=True)

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
        else:
            st.info("👋 ¡Hola! Estás en el modo Niños. Disfruta de las herramientas seguras.")
            modelo_seleccionado = "llama3.2:latest"
else: modelo_seleccionado = "llama3.2:latest"

macro_names = ["🛠️ XAVIWORK", "💬 CHAT Y ASISTENTES", "📁 DOCUMENTOS Y TEXTO", "🌐 WEB Y ACTUALIDAD", "🎬 MULTIMEDIA", "🎓 ESTUDIO Y OFICIOS", "🌍 VIAJES Y HOGAR", "🎵 MÚSICA Y ARTE"]
if st.session_state.get("rol") == "admin":
    macro_names.append("🏢 GESTOR VECINOS")

macro_idx = custom_tabs(macro_names, "macro", 0, cols_per_row=4)
st.markdown("<div style='margin-bottom: 15px;'></div>", unsafe_allow_html=True)

if macro_idx == 0:
    st.markdown("#### 🛠️ XaviWork")
    contexto_xw = cajita_adjuntar_archivo("xw_adj", modelo_seleccionado, _get_key("gemini"))
    
    campo_con_ideas("💡 ¿Sobre qué quieres trabajar? (Define el tema central)", "xw_tema", "temas de trabajo o estudio", modelo_seleccionado, _get_key("gemini"), "Ej: Estrategia de marketing...")
    
    st.markdown("---")
    tabs_xw = st.tabs(["🕸️ Mapa Mental", "🧠 Mapa Conceptual", "📊 PPT", "📄 Informe", "🃏 Tarjetas", "❓ Test", "📋 Tabla", "🎙️ Podcast", "💬 Copiloto", "📅 Timeline", "📖 Glosario", "📊 DAFO", "🖼️ Portada"])
    def _panel_borrador(nombre, prompt_gen, state_key, edit_key, superpowers, download_label=None):
        if st.button(f"✨ Generar {nombre}", key=f"btn_{state_key}", type="primary"):
            if not st.session_state.xw_tema_input.strip(): st.warning("Define el tema arriba.")
            else:
                with st.spinner(f"Generando {nombre}..."):
                    res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\n{prompt_gen}\n\nTEMA: {st.session_state.xw_tema_input}", _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
                with cols[i]: boton_superpoder_borrador(st.session_state[state_key], prompt, f"{state_key}_p{i}", modelo_seleccionado, _get_key("gemini"), state_key, widget_key=edit_key)
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
                        res = llamar_ia(modelo_seleccionado, f"{contexto_xw}\n\nTema: {st.session_state.xw_tema_input}\n\n{ctx_chat}", _get_key("gemini"))
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
        contexto = cajita_adjuntar_archivo("podcast_adj", modelo_seleccionado, _get_key("gemini"))
        st.text_area("Temática:", key="pod_tema_input", height=100)
        c1, c2 = st.columns(2)
        with c1: estilo = st.selectbox("Estilo:", ["Entrevista","Debate","Educativo","Narrativo"], key="p_est")
        with c2: dur = st.slider("Dur (min):", 1, 10, 3, key="p_dur")
        st.markdown("---")
        if st.button("✨ Crear Guion", type="primary", key="btn_pod") and st.session_state["pod_tema_input"]:
            with st.spinner("Escribiendo..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\nGuion podcast {dur} min. Tema: '{st.session_state['pod_tema_input']}'. Estilo: {estilo}. 3 locutores. Formato 'Locutor 1: [texto]'.", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
        ctx = cajita_adjuntar_archivo("chat_adj", modelo_seleccionado, _get_key("gemini"))
        p = chat_input_pro("main")
        if p:
            st.session_state.chat_history.append({"role":"user","content":p}); st.chat_message("user").write(p)
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"{ctx}\n\n{p}", _get_key("gemini"), st.session_state.chat_history)
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
            texto = ""
            for a in archs:
                texto += f"\n--- {a.name} ---\n" + leer_archivo_universal(a, modelo_seleccionado)
            texto = texto[:MAX_TEXTO_IA]
            st.success("Cargados.")
            for m in st.session_state.pdf_chat_history: st.chat_message(m["role"]).write(m["content"])
            p = chat_input_pro("pdf")
            if p:
                st.session_state.pdf_chat_history.append({"role":"user","content":p})
                with st.spinner("Buscando..."):
                    res = llamar_ia(modelo_seleccionado, f"Contexto:\n{texto}\n\nPregunta: {p}", _get_key("gemini"), st.session_state.pdf_chat_history)
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
                    res = llamar_ia(modelo_seleccionado, f"ERES: {ag['rol']}\nUsuario: {p}", _get_key("gemini"), st.session_state.agente_chat_history)
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state.agente_chat_history.append({"role":"assistant","content":res}); st.rerun()
    elif sub_idx == 5:
        st.markdown("#### Debate")
        t = campo_con_ideas("Tema:", "deb", "temas de debate", modelo_seleccionado, _get_key("gemini"), "Ej: Coche eléctrico")
        if st.button("Iniciar") and t:
            res = llamar_ia(modelo_seleccionado, f"Debate {t}: 3 a favor, 3 en contra, conclusión.", _get_key("gemini"))
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
            else: st.markdown(res); botones_descarga(res, "Debate")
    elif sub_idx == 6:
        st.markdown("#### Cocina")
        ing = campo_con_ideas("Ingredientes:", "chef1", "recetas", modelo_seleccionado, _get_key("gemini"), "Ej: Pollo")
        pref = campo_con_ideas("Menú:", "chef2", "menús", modelo_seleccionado, _get_key("gemini"), "Ej: Vegetariano")
        c1, c2 = st.columns(2)
        with c1:
            if st.button("Receta") and ing:
                res = llamar_ia(modelo_seleccionado, f"Receta con {ing}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.info(res)
        with c2:
            if st.button("Menú") and pref:
                res = llamar_ia(modelo_seleccionado, f"Menú semanal {pref}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.info(res)
    elif sub_idx == 7:
        st.markdown("#### Cuentos")
        p = campo_con_ideas("Protagonista:", "c1", "protagonistas", modelo_seleccionado, _get_key("gemini"), "Ej: Dragón")
        l = campo_con_ideas("Lugar:", "c2", "lugares mágicos", modelo_seleccionado, _get_key("gemini"), "Ej: Bosque")
        if st.button("Crear") and p and l:
            res = llamar_ia(modelo_seleccionado, f"Cuento sobre {p} en {l}.", _get_key("gemini"))
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
            else:
                st.text_area("Cuento:", res, height=300); botones_descarga(res, "Cuento")
                if st.button("Leer"):
                    aud = generar_audio_sincrono(res, "es-ES-ElviraNeural")
                    if aud: st.audio(aud, format="audio/mp3")
    elif sub_idx == 8:
        st.markdown("#### Código")
        req = campo_con_ideas("Código:", "code", "scripts", modelo_seleccionado, _get_key("gemini"), "Ej: Renombrar archivos")
        if st.button("Procesar") and req:
            res = llamar_ia(modelo_seleccionado, f"Programador senior. Explica o mejora:\n{req}", _get_key("gemini"))
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
            else: st.code(res, language="python")
    elif sub_idx == 9:
        st.markdown("#### Traductor")
        txt = campo_con_ideas("Texto:", "trad", "frases", modelo_seleccionado, _get_key("gemini"), "Texto...")
        c1, c2 = st.columns(2)
        with c1: o = st.selectbox("Origen:", ["Auto","Español","Inglés"], key="o")
        with c2: d = st.selectbox("Destino:", ["Español","Inglés","Francés"], key="d")
        if st.button("Traducir") and txt:
            res = llamar_ia(modelo_seleccionado, f"Traduce del {o} al {d}:\n{txt}", _get_key("gemini"))
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
            else: st.info(res); botones_descarga(res, "Traduccion")

elif macro_idx == 2:
    sub_names = ["📁 ANÁLISIS", "⚖️ COMPARAR", "📊 OFFICE", "📄 PDF", "🧾 FACTURAS", "📇 QR & CV", "📋 PORTAPAPELES"]
    sub_idx = custom_tabs(sub_names, "doc_sub", 0)
    if sub_idx == 0:
        st.markdown("#### Análisis")
        archs = st.file_uploader("PDF, Word, Excel o TXT (Puedes subir varios)", type=["pdf","txt","docx","xlsx"], accept_multiple_files=True)
        opc = st.selectbox("Análisis:", ["Resumen Ejecutivo","Extraer Fechas y Cifras","Buscar Errores","Crear Test","Traducir al Inglés"])
        if archs and st.button("🔎 Ejecutar", type="primary"):
            with st.spinner("Analizando..."):
                texto = ""
                for a in archs:
                    texto += f"\n--- DOCUMENTO: {a.name} ---\n" + leer_archivo_universal(a, modelo_seleccionado)
                
                prompt = {"Resumen Ejecutivo": f"Resumen:\n{texto}", "Extraer Fechas y Cifras": f"Fechas:\n{texto}", "Buscar Errores": f"Errores:\n{texto}", "Crear Test": f"Test 5 preguntas:\n{texto}", "Traducir al Inglés": f"Traduce al inglés:\n{texto}"}[opc]
                res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.text_area("Resultado:", res, height=400); botones_descarga(res, "Analisis")
    elif sub_idx == 1:
        st.markdown("#### Comparador")
        a = st.file_uploader("Original:", type=["pdf","txt"]); b = st.file_uploader("Modificado:", type=["pdf","txt"])
        if a and b and st.button("Comparar"):
            with st.spinner("Buscando..."):
                ta = leer_archivo_universal(a, modelo_seleccionado); tb = leer_archivo_universal(b, modelo_seleccionado)
                res = llamar_ia(modelo_seleccionado, f"Diferencias:\nA:\n{ta[:10000]}\nB:\n{tb[:10000]}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.text_area("Diferencias:", res, height=400); botones_descarga(res, "Comparativa")
    elif sub_idx == 2:
        st.markdown("#### Crear Documentos")
        contexto = cajita_adjuntar_archivo("office_adj", modelo_seleccionado, _get_key("gemini"))
        prompt_office = campo_con_ideas("¿Qué documento?", "office", "documentos", modelo_seleccionado, _get_key("gemini"), "Ej: Presupuesto")
        formato = st.radio("Formato:", ["Word","Excel","PowerPoint"], key="office_fmt", horizontal=True)
        if st.button("✨ Generar", type="primary", key="btn_office_gen"):
            if prompt_office:
                with st.spinner("Generando..."):
                    prompt_ia = f"{contexto}\n\n{prompt_office}"
                    if formato == "PowerPoint": prompt_ia += "\n\nFormatea para diapositivas con doble salto."
                    elif formato == "Excel": prompt_ia += "\n\nDevuelve CSV."
                    res = llamar_ia(modelo_seleccionado, prompt_ia, _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state.office_borrador = res
            else: st.warning("Escribe un prompt.")
        if st.session_state.get("office_borrador"):
            st.markdown("### 📝 Borrador"); borrador_editable("office_borrador", "office_txt", height=350)
            c1, c2, c3 = st.columns(3)
            with c1: boton_superpoder_borrador(st.session_state.office_txt, "Más profesional", "of_sup1", modelo_seleccionado, _get_key("gemini"), "office_borrador", widget_key="office_txt")
            with c2: boton_superpoder_borrador(st.session_state.office_txt, "Resúmelo", "of_sup2", modelo_seleccionado, _get_key("gemini"), "office_borrador", widget_key="office_txt")
            with c3: boton_superpoder_borrador(st.session_state.office_txt, "+detalles", "of_sup3", modelo_seleccionado, _get_key("gemini"), "office_borrador", widget_key="office_txt")
            st.markdown("---")
            if formato == "Word": st.download_button("📥 Word", generar_word(st.session_state.office_txt, (st.session_state.get("logo_empresa_upload").getvalue() if st.session_state.get("logo_empresa_upload") else None)), "doc.docx")
            elif formato == "Excel": st.download_button("📥 Excel", generar_excel(st.session_state.office_txt), "doc.xlsx")
            else: st.download_button("📥 PPTX", generar_pptx("Presentación", st.session_state.office_txt), "doc.pptx")
    elif sub_idx == 3:
        st.markdown("#### PDF Pro")
        contexto = cajita_adjuntar_archivo("pdf_adj", modelo_seleccionado, _get_key("gemini"))
        titulo = st.text_input("Título PDF:")
        contenido = campo_con_ideas("Contenido:", "pdf_c", "contenido PDF", modelo_seleccionado, _get_key("gemini"), "Escribe...")
        if st.button("✨ Generar PDF", type="primary", key="btn_pdf_gen") and titulo and contenido:
            with st.spinner("Generando..."):
                res = llamar_ia(modelo_seleccionado, f"{contexto}\n\n{contenido}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.session_state.pdf_borrador = res
        if st.session_state.get("pdf_borrador"):
            st.markdown("### 📝 Borrador"); borrador_editable("pdf_borrador", "pdf_txt", height=350)
            c1, c2 = st.columns(2)
            with c1: boton_superpoder_borrador(st.session_state.pdf_txt, "Más técnico", "pdf_sup1", modelo_seleccionado, _get_key("gemini"), "pdf_borrador", widget_key="pdf_txt")
            with c2: boton_superpoder_borrador(st.session_state.pdf_txt, "+conclusiones", "pdf_sup2", modelo_seleccionado, _get_key("gemini"), "pdf_borrador", widget_key="pdf_txt")
            st.markdown("---"); botones_descarga(st.session_state.pdf_borrador, titulo or "PDF_Pro")
    elif sub_idx == 4:
        st.markdown("#### 🧾 Facturas")
        archs_factura = st.file_uploader("Factura(s) (JPG/PNG/PDF):", type=["jpg","jpeg","png","pdf"], accept_multiple_files=True)
        if st.button("🪄 Extraer", type="primary") and archs_factura:
            with st.spinner("Escaneando..."):
                prompt_factura = "Extrae datos. Devuelve CSV: Proveedor, Fecha, Concepto, Base_Imponible, IVA, Total. Sin markdown."
                resultados_totales = ""
                for a in archs_factura:
                    if a.name.lower().endswith(".pdf"):
                        texto_pdf = leer_archivo_universal(a, modelo_seleccionado)
                        res = llamar_ia(modelo_seleccionado, f"{prompt_factura}\n\n{texto_pdf}", _get_key("gemini"))
                    else:
                        res = analizar_archivo_con_gemini(a.getvalue(), "image/jpeg", prompt_factura, _get_key("gemini"))
                    resultados_totales += f"\n[Factura: {a.name}]\n{res}\n"
                
                if "[ERROR" in resultados_totales: st.error(mensaje_error_amigable(resultados_totales))
                else:
                    st.success("✅ Listo."); st.text_area("Datos:", value=resultados_totales, height=150, key="factura_data")
                    st.download_button("📊 Excel", generar_excel(st.session_state.factura_data), "facturas.xlsx", use_container_width=True)
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
        txt = st.text_area("Texto:"); inst = campo_con_ideas("¿Qué hago?", "clip", "instrucciones", modelo_seleccionado, _get_key("gemini"), "Ej: Resume")
        if st.button("Procesar") and txt and inst:
            res = llamar_ia(modelo_seleccionado, f"{inst}:\n{txt}", _get_key("gemini"))
            if "[ERROR" in res: st.error(mensaje_error_amigable(res))
            else: st.info(res)

elif macro_idx == 3:
    sub_names = ["🌐 BUSCADOR", "⚽ DEPORTES", "📺 TV", "📊 WEB A EXCEL", "🎥 YOUTUBE", "📰 LECTOR"]
    sub_idx = custom_tabs(sub_names, "web_sub", 0)
    if sub_idx == 0:
        q = campo_con_ideas("Buscar:", "web", "búsquedas", modelo_seleccionado, _get_key("gemini"), "Ej: Noticias IA")
        if st.button("🔍 Buscar") and q:
            for r in buscar_en_internet(q): st.markdown(f"**[{r['titulo']}]({r['url']})**"); st.caption(r['resumen'])
    elif sub_idx == 1:
        d = st.selectbox("Deporte:", ["Fútbol","Baloncesto","Tenis"])
        if st.button("Noticias"):
            for n in buscar_noticias_deportivas(d, "Todas"): st.markdown(f"**{n['titulo']}** - {n['resumen']}")
    elif sub_idx == 2:
        c = campo_con_ideas("Canal:", "tv", "TV", modelo_seleccionado, _get_key("gemini"), "Ej: Telecinco")
        if st.button("Programación") and c:
            res = buscar_en_internet(c); r_ia = llamar_ia(modelo_seleccionado, f"Programación {c}: {res}", _get_key("gemini"))
            if "[ERROR" in r_ia: st.error(mensaje_error_amigable(r_ia))
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
            if texto_original.startswith("[ERROR]"): st.error(mensaje_error_amigable(texto_original))
            else:
                with st.expander("Ver transcripción"): st.text_area("Texto:", value=texto_original, height=150, key=f"yt_orig_{vid}")
                prompt = f"Resume esto en español:\n{texto_original[:8000]}" if modo == "Resumir" else f"Traduce al español:\n{texto_original[:8000]}"
                with st.spinner("Procesando..."): res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.text_area("Resultado:", res, height=300); botones_descarga(res, f"YT_{vid}")
    elif sub_idx == 5:
        url = st.text_input("URL Web:")
        if st.button("Leer") and url:
            try:
                if not url_segura(url): st.error("URL no permitida."); st.stop()
                txt = " ".join(re.sub('<[^<]+?>', ' ', requests.get(url, headers={"User-Agent":"Mozilla/5.0"}).text).split())[:5000]
                res = llamar_ia(modelo_seleccionado, f"Resume: {txt}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.text_area("Resumen:", res, height=200)
            except Exception as e: st.error(f"Error: {e}")
elif macro_idx == 4:
    st.markdown("#### 🎬 Multimedia y Conversor PRO")
    sub_names = ["🖼️ IMAGEN", "🎙️ TEXTO A VOZ", "🔄 CONVERSOR PRO"]
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
    elif sub_idx == 2:
        st.markdown("### 🔄 Conversor Multiformato PRO")
        st.caption("Convierte audio, vídeo, extrae fotogramas, transcribe y más.")

        tipo_conv = st.radio("¿Qué quieres hacer?", ["Convertir Audio", "Convertir Vídeo", "Extraer Audio de Vídeo", "Cortar/Recortar", "Transcribir a Texto", "Extraer Fotograma"], horizontal=True)

        def procesar_ffmpeg(comando, nombre_defecto, extension, es_audio=False, es_video=False, es_imagen=False):
            with st.spinner("Procesando archivo..."):
                try:
                    subprocess.run(comando, check=True, stdout=subprocess.PIPE, stderr=subprocess.PIPE)
                    with open(comando[-1], "rb") as f: res_bytes = f.read()

                    nuevo_nombre = st.text_input("📝 Renombrar archivo antes de descargar:", value=nombre_defecto)
                    nombre_limpio = re.sub(r'[^a-zA-Z0-9_\- ]', '', nuevo_nombre).strip().replace(" ", "_")

                    if es_audio: st.audio(res_bytes)
                    elif es_video: st.video(res_bytes)
                    elif es_imagen: st.image(res_bytes)

                    st.download_button(f"💾 Descargar {extension.upper()}", res_bytes, f"{nombre_limpio}.{extension}", use_container_width=True)
                except Exception as e:
                    st.error(f"Error al procesar: {e}")
                finally:
                    for file in comando:
                        if isinstance(file, str) and os.path.exists(file) and file != FFMPEG_EXE:
                            try: os.remove(file)
                            except: pass

        if tipo_conv == "Convertir Audio":
            arch_audio = st.file_uploader("Sube tu audio:", type=["mp3", "wav", "m4a", "ogg", "flac", "opus", "aac"])
            fmt_audio = st.selectbox("Formato de salida:", ["mp3", "wav", "m4a", "ogg"])
            if arch_audio and st.button("Convertir", type="primary"):
                tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=f".{arch_audio.name.split('.')[-1]}").name
                with open(tmp_in, "wb") as f: f.write(arch_audio.getvalue())
                tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=f".{fmt_audio}").name
                procesar_ffmpeg([FFMPEG_EXE, "-y", "-i", tmp_in, tmp_out], arch_audio.name.split('.')[0], fmt_audio, es_audio=True)

        elif tipo_conv == "Convertir Vídeo":
            arch_video = st.file_uploader("Sube tu vídeo:", type=["mp4", "mov", "avi", "mkv", "webm"])
            fmt_video = st.selectbox("Formato de salida:", ["mp4", "avi", "mkv", "webm", "gif"])
            if arch_video and st.button("Convertir", type="primary"):
                tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=f".{arch_video.name.split('.')[-1]}").name
                with open(tmp_in, "wb") as f: f.write(arch_video.getvalue())
                tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=f".{fmt_video}").name
                procesar_ffmpeg([FFMPEG_EXE, "-y", "-i", tmp_in, tmp_out], arch_video.name.split('.')[0], fmt_video, es_video=True)

        elif tipo_conv == "Extraer Audio de Vídeo":
            arch_video = st.file_uploader("Sube un vídeo:", type=["mp4", "mov", "avi", "mkv"])
            fmt_audio = st.selectbox("Formato de audio:", ["mp3", "m4a", "wav"])
            if arch_video and st.button("Extraer Audio", type="primary"):
                tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=f".{arch_video.name.split('.')[-1]}").name
                with open(tmp_in, "wb") as f: f.write(arch_video.getvalue())
                tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=f".{fmt_audio}").name
                procesar_ffmpeg([FFMPEG_EXE, "-y", "-i", tmp_in, "-vn", tmp_out], f"Audio_{arch_video.name.split('.')[0]}", fmt_audio, es_audio=True)

        elif tipo_conv == "Cortar/Recortar":
            arch_media = st.file_uploader("Sube audio o vídeo:", type=["mp3", "wav", "m4a", "mp4", "mov"])
            if arch_media:
                c1, c2 = st.columns(2)
                with c1: inicio = st.text_input("Inicio (HH:MM:SS):", "00:00:00")
                with c2: fin = st.text_input("Fin (HH:MM:SS):", "00:00:10")
                if st.button("Cortar", type="primary"):
                    ext = arch_media.name.split('.')[-1]
                    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}").name
                    with open(tmp_in, "wb") as f: f.write(arch_media.getvalue())
                    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=f".{ext}").name
                    es_aud = ext in ["mp3", "wav", "m4a"]
                    procesar_ffmpeg([FFMPEG_EXE, "-y", "-i", tmp_in, "-ss", inicio, "-to", fin, "-c", "copy", tmp_out], f"Cortado_{arch_media.name.split('.')[0]}", ext, es_audio=es_aud, es_video=not es_aud)

        elif tipo_conv == "Transcribir a Texto":
            st.info("💡 Requiere tener configurada la API Key de Groq.")
            arch_audio = st.file_uploader("Sube audio o vídeo para transcribir:", type=["mp3", "wav", "m4a", "mp4", "opus"])
            if arch_audio and st.button("Transcribir", type="primary"):
                with st.spinner("Transcribiendo con IA (Whisper)..."):
                    texto = transcribir_archivo_local(arch_audio.getvalue(), arch_audio.name, _get_key("groq"))
                    if "[ERROR]" in texto: st.error(mensaje_error_amigable(texto))
                    else:
                        nuevo_nombre = st.text_input("📝 Renombrar archivo:", value=f"Transcripcion_{arch_audio.name.split('.')[0]}")
                        st.text_area("Texto detectado:", value=texto, height=250)
                        st.download_button("💾 Descargar TXT", texto.encode("utf-8"), f"{nuevo_nombre}.txt", "text/plain", use_container_width=True)

        elif tipo_conv == "Extraer Fotograma":
            arch_video = st.file_uploader("Sube un vídeo:", type=["mp4", "mov", "avi"])
            if arch_video:
                segundo = st.number_input("Segundo exacto del vídeo:", min_value=0, value=1)
                if st.button("Extraer Foto", type="primary"):
                    tmp_in = tempfile.NamedTemporaryFile(delete=False, suffix=f".{arch_video.name.split('.')[-1]}").name
                    with open(tmp_in, "wb") as f: f.write(arch_video.getvalue())
                    tmp_out = tempfile.NamedTemporaryFile(delete=False, suffix=".jpg").name
                    procesar_ffmpeg([FFMPEG_EXE, "-y", "-i", tmp_in, "-ss", str(segundo), "-vframes", "1", tmp_out], f"Foto_{arch_video.name.split('.')[0]}", "jpg", es_imagen=True)

elif macro_idx == 5:
    st.markdown("#### 🎓 Estudio y Oficios")
    sub_names = ["📚 DEBERES", "👔 ENTREVISTA", "🛠️ PRL PRO"]
    sub_idx = custom_tabs(sub_names, "study_sub", 0)
    if sub_idx == 0:
        st.markdown("### Resolver Deberes")
        materia = st.selectbox("Materia:", ["Matemáticas", "Historia", "Ciencias", "Lengua", "Física y Química", "Inglés"]); pregunta = st.text_area("Duda:", height=150)
        if st.button("🔎 Resolver", type="primary") and pregunta:
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, f"Profesor experto en {materia}. Resuelve paso a paso:\n\n{pregunta}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.markdown(res); botones_descarga(res, "Deberes_Resueltos", formatos_extra=True)
    elif sub_idx == 1:
        st.markdown("### Simulador de Entrevista")
        puesto = st.text_input("¿Puesto?", placeholder="Ej: Camarero...")
        if st.button("🤖 Empezar", type="primary") and puesto:
            st.session_state.entrevista_puesto = puesto
            st.session_state.entrevista_historial = [{"role": "user", "content": f"Actúa como reclutador para {puesto}. Preséntate y haz la primera pregunta."}]
            with st.spinner("Pensando..."):
                res = llamar_ia(modelo_seleccionado, st.session_state.entrevista_historial[0]["content"], _get_key("gemini"))
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
                    res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                    st.session_state.entrevista_historial.append({"role": "assistant", "content": res}); st.rerun()
    elif sub_idx == 2:
        st.markdown("### 🦺 PRL PRO")
        tabs_prl = st.tabs(["📊 Evaluación", "📝 Test", "📖 Normativa"])
        with tabs_prl[0]:
            puesto = st.text_input("Puesto:", placeholder="Ej: Soldador..."); entorno = st.text_area("Entorno:")
            if st.button("🔍 Evaluar", type="primary") and puesto:
                with st.spinner("Analizando..."):
                    res = llamar_ia(modelo_seleccionado, f"Técnico PRL. Evalúa riesgos para '{puesto}'. Contexto: '{entorno}'.\nTabla Markdown: | Riesgo | Probabilidad | Consecuencia | Nivel | Medidas | EPIs |\nAñade 'Recomendaciones'.", _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.markdown(res); botones_descarga(res, f"Evaluacion_PRL", formatos_extra=True)
        with tabs_prl[1]:
            tema_test = st.selectbox("Área:", ["Ley 31/1995", "EPIs", "Primeros Auxilios", "Ergonomía", "Máquinas"])
            if st.button("📝 Generar Test", type="primary"):
                with st.spinner("Preparando..."):
                    res = llamar_ia(modelo_seleccionado, f"Test 10 preguntas (A,B,C,D) sobre '{tema_test}'. Incluye 'Soluciones'. Normativa española.", _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.markdown(res); botones_descarga(res, "Test_PRL", formatos_extra=True)
        with tabs_prl[2]:
            if st.button("📖 Resumen Leyes"):
                with st.spinner("Resumiendo..."):
                    res = llamar_ia(modelo_seleccionado, "Resumen estructurado Ley 31/1995 PRL: 1. Objeto, 2. Principios, 3. Derechos y obligaciones.", _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
                res = llamar_ia(modelo_seleccionado, f"Itinerario {dias} días para {destino}. Estilo: {tipo_viaje}. Incluye qué ver, dónde comer y consejos.", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.markdown(res); botones_descarga(res, f"Viaje_{destino}", formatos_extra=True)
    elif sub_idx == 1:
        st.markdown("### Recetas")
        ingredientes = campo_con_ideas("Ingredientes:", "hogar", "recetas", modelo_seleccionado, _get_key("gemini"), "Ej: Huevos...")
        if st.button("🍳 Crear", type="primary") and ingredientes:
            with st.spinner("Cocinando..."):
                res = llamar_ia(modelo_seleccionado, f"Tengo: {ingredientes}. Dame 2 recetas fáciles. Incluye pasos y tiempo.", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
                res = llamar_ia(modelo_seleccionado, f"Entrenador. Usuario: {edad}a, {peso}kg, {altura}cm. Objetivo: '{objetivo}'. Entrena {dias_entreno} días en '{equipo}'. Genera: 1. Rutina semanal. 2. Consejos nutricionales.", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
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
        with c1: genero = st.selectbox("Género:", ["Bachata", "Reggaeton", "Pop", "Rock", "Rap", "Cumbia", "Balada", "Electrónica", "Rumba", "Flamenco", "Salsa", "Merengue", "Tango", "Reggae", "Metal", "Jazz", "Blues", "Country", "Clásica", "Trap", "Techno", "Indie", "Punk", "Funk", "Bossa Nova"], key="gen_letra")
        with c2: artista_estilo = st.text_input("Estilo artista:", placeholder="Ej: Bad Bunny...", key="art_letra")

        st.markdown("#### ⚙️ Ajustes de Voz y Música (Elige antes de generar)")
        hacer_dueto = st.checkbox("🎤 Hacer Dúo")

        voces_disponibles = ["es-ES-ElviraNeural", "es-MX-JorgeNeural", "es-AR-TomasNeural"]
        bases_disponibles = ["Bachata", "Reggaeton", "Pop", "Rock", "Rap", "Cumbia", "Balada", "Electrónica", "Rumba", "Flamenco", "Salsa", "Merengue", "Tango", "Reggae", "Metal", "Jazz", "Blues", "Country", "Clásica", "Trap", "Techno", "Indie", "Punk", "Funk", "Bossa Nova"]

        cv1, cv2, cb = st.columns(3)
        with cv1: voz_config1 = st.selectbox("Voz 1:", voces_disponibles, key="vc_v1")
        with cv2:
            if hacer_dueto: voz_config2 = st.selectbox("Voz 2:", voces_disponibles, index=1, key="vc_v2")
            else:
                voz_config2 = None
                st.selectbox("Voz 2:", ["(Desactivado)"], disabled=True, key="vc_v2_off")
        with cb: base_config = st.selectbox("Base musical:", bases_disponibles, key="vc_base")

        if st.button("🪄 Generar Letra IA", use_container_width=True, type="primary") and tema_cancion:
            with st.spinner("Componiendo..."):
                prompt = f"Letra de canción {genero} sobre: '{tema_cancion}'. "
                if artista_estilo: prompt += f"Estilo {artista_estilo}. "
                prompt += "Estructura (Intro, Verso 1, Coro, etc.) y añade indicaciones de producción entre corchetes."
                res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.session_state.letra_edit = res; st.rerun()

        if st.session_state.letra_edit.strip():
            st.markdown("---")
            letra_final = st.text_area("📜 Letra (Editable):", key="letra_edit", height=300)
            botones_descarga(letra_final, "Letra_Cancion")

            st.markdown("### 🎵 Ponerle Música")
            st.info("💡 Puedes cambiar las voces o la base musical aquí abajo antes de generar el audio final.")

            ca1, ca2, ca3 = st.columns(3)
            with ca1:
                idx_v1 = voces_disponibles.index(voz_config1) if voz_config1 in voces_disponibles else 0
                voz_final1 = st.selectbox("Voz 1 Final:", voces_disponibles, index=idx_v1, key="vf_v1")
            with ca2:
                if hacer_dueto:
                    idx_v2 = voces_disponibles.index(voz_config2) if voz_config2 in voces_disponibles else 1
                    voz_final2 = st.selectbox("Voz 2 Final:", voces_disponibles, index=idx_v2, key="vf_v2")
                else:
                    voz_final2 = None
                    st.selectbox("Voz 2 Final:", ["(Desactivado)"], disabled=True, key="vf_v2_off")
            with ca3:
                idx_b = bases_disponibles.index(base_config) if base_config in bases_disponibles else 0
                base_final = st.selectbox("Base Final:", bases_disponibles, index=idx_b, key="vf_base")

            if st.button("🎶 Generar Canción Completa", type="primary", use_container_width=True):
                with st.spinner("Grabando..."):
                    audio_final = generar_cancion_con_base(letra_final, voz_final1, base_final, voz_final2)
                    if audio_final:
                        st.success("✅ ¡Lista!")
                        st.audio(audio_final, format="audio/mp3")
                        st.download_button("💾 Descargar MP3", audio_final, "cancion_xavilu.mp3", "audio/mp3", use_container_width=True)

    elif sub_idx == 1:
        st.markdown("### 🌞 Generador de Prompts para Suno AI")
        st.caption("Suno AI canta de verdad. Genera el texto perfecto para pegarlo en su web.")
        c1, c2 = st.columns(2)
        with c1:
            st_tema = st.text_input("Tema:", placeholder="Ej: Marta y Luz en el campamento...")
            st_genero = st.selectbox("Género:", ["Bachata", "Reggaeton", "Pop", "Rock", "Rap", "Cumbia", "Balada", "Electrónica", "Rumba", "Flamenco", "Salsa", "Merengue", "Tango", "Reggae", "Metal", "Jazz", "Blues", "Country", "Clásica", "Trap", "Techno", "Indie", "Punk", "Funk", "Bossa Nova"])
            st_voz = st.selectbox("Voz:", ["Dúo (Hombre y Mujer)", "Masculina", "Femenina", "Coro"])
        with c2:
            st_instr = st.text_input("Instrumentos:", placeholder="Ej: Guitarra, marimba...")
            st_artista = st.text_input("Artista similar:", placeholder="Ej: Romeo Santos...")
        if st.button("✨ Generar Prompt para Suno", type="primary") and st_tema:
            with st.spinner("Preparando el prompt..."):
                prompt_ia = f"""Actúa como un experto en ingeniería de prompts para Suno AI.
                Tema: "{st_tema}"
                Género: {st_genero}
                Voz: {st_voz}
                Instrumentos: {st_instr}
                Artista similar: {st_artista}

                Tu tarea es generar EXCLUSIVAMENTE el texto listo para copiar y pegar en Suno AI.
                REGLAS ESTRICTAS:
                1. NO escribas saludos, introducciones ni comentarios (ej: "Aquí tienes...").
                2. NO uses bloques de código markdown (```) ni formato JSON.
                3. Sigue el formato exacto usando los separadores que se indican a continuación.

                Formato de salida:

                ---STYLE---
                [Aquí va el prompt de estilo en inglés, máximo 200 caracteres, describiendo género, instrumentos y voz]

                ---LYRICS---
                [Aquí va la letra en español con las etiquetas [Verse], [Chorus], [Bridge], etc.]
                """
                res = llamar_ia(modelo_seleccionado, prompt_ia, _get_key("gemini"))
                if "[ERROR" in res:
                    st.error(mensaje_error_amigable(res))
                else:
                    st.success("✅ Prompt generado. Listo para copiar.")
                    res_limpia = res.replace("```markdown", "").replace("```", "").strip()
                    if "---STYLE---" in res_limpia and "---LYRICS---" in res_limpia:
                        partes = res_limpia.split("---LYRICS---")
                        style_part = partes[0].replace("---STYLE---", "").strip()
                        lyrics_part = partes[1].strip()
                        st.markdown("### 🎧 Estilo de Música (Copia y pega en 'Style of Music'):")
                        st.code(style_part, language="text")
                        st.markdown("### 📜 Letra (Copia y pega en 'Lyrics' con Custom Mode activado):")
                        st.code(lyrics_part, language="text")
                    else:
                        st.warning("La IA no siguió el formato exacto, pero aquí tienes el texto limpio:")
                        st.text_area("Copia esto:", value=res_limpia, height=300)
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
                res = llamar_ia(modelo_seleccionado, f"Escribe un poema profundo sobre: {tema_poema}", _get_key("gemini"))
                if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                else: st.markdown(res); botones_descarga(res, "Poema")

elif macro_idx == 8:
    st.markdown("#### 🏢 Gestor de Comunidad PRO")
    sub_names = ["🚨 INCIDENCIAS", "📜 NORMATIVA", "🧾 ANALIZAR PRESUPUESTOS", "📊 ACTAS", "📝 CITACIONES"]
    sub_idx = custom_tabs(sub_names, "vecinos_sub", 0)

    if sub_idx == 0:
        st.markdown("### 🚨 Gestión de Incidencias")
        st.caption("Registra cada aviso y sigue su responsable, plazo y estado. Los registros se guardan en la base de datos local.")
        c1, c2 = st.columns(2)
        with c1:
            tipo_inc = st.selectbox("Tipo:", ["Avería general", "Ruidos", "Filtración", "Daños", "Otro"], key="inc_tipo")
            urgencia = st.selectbox("Urgencia:", ["Baja", "Media", "Alta", "Urgente"], key="inc_urgencia")
            vecino_afectado = st.text_input("¿A quién afecta?", value="A mi vivienda", key="inc_afectado")
        with c2:
            ubicacion = st.text_input("Ubicación:", placeholder="Ej: Portal 2, 3ºB", key="inc_ubicacion")
            responsable = st.text_input("Responsable de seguimiento", key="inc_responsable")
            fecha_limite = st.text_input("Fecha objetivo (AAAA-MM-DD)", placeholder="2026-11-20", key="inc_fecha_limite")
        descripcion_inc = st.text_area("Describe lo ocurrido:", height=130, key="inc_descripcion")
        col_guardar, col_redactar = st.columns(2)
        with col_guardar:
            if st.button("💾 Registrar incidencia", type="primary", use_container_width=True, key="inc_guardar"):
                if not descripcion_inc.strip():
                    st.warning("Describe la incidencia antes de registrarla.")
                else:
                    fecha_ok = True
                    if fecha_limite.strip():
                        try: datetime.date.fromisoformat(fecha_limite.strip())
                        except ValueError:
                            fecha_ok = False
                            st.warning("Usa el formato de fecha AAAA-MM-DD, por ejemplo 2026-11-20.")
                    if fecha_ok:
                        try:
                            incidencia_id = guardar_incidencia(tipo_inc, urgencia, vecino_afectado, ubicacion, descripcion_inc.strip(), responsable.strip(), fecha_limite.strip())
                            st.success(f"Incidencia #{incidencia_id} registrada.")
                        except Exception as e:
                            st.error(f"No se pudo guardar la incidencia: {e}")
        with col_redactar:
            if st.button("📝 Redactar reclamación", use_container_width=True, key="inc_redactar") and descripcion_inc.strip():
                st.session_state.pop("incidencia_res", None)
                st.session_state.pop("inc_documento", None)
                with st.spinner("Redactando..."):
                    prompt = f"Administrador de fincas. Redacta una reclamación formal. Tipo: {tipo_inc}. Urgencia: {urgencia}. Afecta a: {vecino_afectado}. Ubicación: {ubicacion}. Descripción: {descripcion_inc}. No inventes fechas, normas ni datos de contacto. Incluye encabezado, solicitud de solución y espacios identificados para los datos que falten."
                    res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["incidencia_res"] = res

        if st.session_state.get("incidencia_res"):
            documento_inc_editado = st.text_area("Documento (puedes editarlo):", value=st.session_state["incidencia_res"], height=300, key="inc_documento")
            botones_descarga(documento_inc_editado, f"Reclamacion_{tipo_inc.replace(' ', '_')}")

        st.markdown("#### Seguimiento de incidencias")
        try:
            incidencias = listar_incidencias()
            if not incidencias:
                st.info("Aún no hay incidencias registradas.")
            for incidencia in incidencias:
                iid, creada, tipo, urg, afectado, lugar, desc, resp, limite, estado, actualizada = incidencia
                with st.expander(f"#{iid} · {urg} · {tipo} · {lugar or 'Ubicación pendiente'} · {estado}"):
                    st.caption(f"Creada: {creada} · Última actualización: {actualizada}")
                    st.write(f"**Afectado:** {afectado or 'No indicado'}")
                    st.write(f"**Descripción:** {desc}")
                    st.write(f"**Fecha objetivo:** {limite or 'No indicada'}")
                    c_estado, c_responsable, c_actualizar = st.columns([1, 2, 1])
                    with c_estado:
                        estado_nuevo = st.selectbox("Estado", ["Abierta", "En curso", "Resuelta", "Cerrada"], index=["Abierta", "En curso", "Resuelta", "Cerrada"].index(estado), key=f"inc_estado_{iid}")
                    with c_responsable:
                        responsable_nuevo = st.text_input("Responsable", value=resp or "", key=f"inc_resp_{iid}")
                    with c_actualizar:
                        if st.button("Guardar", key=f"inc_update_{iid}", use_container_width=True):
                            if actualizar_incidencia(iid, estado_nuevo, responsable_nuevo.strip()):
                                st.success("Seguimiento actualizado.")
                                st.rerun()
        except Exception as e:
            st.error(f"No se pudo cargar el seguimiento: {e}")

    elif sub_idx == 1:
        st.markdown("### 📜 Normativa y Leyes")
        consulta = campo_con_ideas("Duda legal:", "norm", "dudas legales de comunidad", modelo_seleccionado, _get_key("gemini"), "Ej: ¿Puedo poner cámara en el rellano?")
        if st.button("🔎 Consultar en fuentes oficiales") and consulta:
            st.session_state["normativa_res"] = ""
            st.session_state["normativa_fuentes"] = []
            with st.spinner("Consultando legislación consolidada del BOE..."):
                fuentes = buscar_legislacion_boe(consulta)
                if not fuentes:
                    st.error("No se pudo verificar legislación en el BOE. Comprueba la conexión y vuelve a intentarlo; no se generó una respuesta presentada como vigente.")
                else:
                    contexto_legal = "\n\n".join(
                        f"FUENTE OFICIAL: {f['titulo']}\nÚltima actualización que informa el BOE: {f['actualizado'] or 'no indicada'}\nTexto relevante actual:\n{f['texto']}"
                        for f in fuentes
                    )
                    prompt_legal = f"""Responde en español a la consulta sobre normativa de comunidades de propietarios en España usando EXCLUSIVAMENTE las fuentes oficiales del BOE incluidas abajo.
No inventes artículos, obligaciones, mayorías ni fechas. Cita cada norma y artículo que sustente la respuesta. Si las fuentes no resuelven la cuestión, dilo expresamente y señala qué normativa autonómica, municipal o jurisprudencia habría que verificar. Separa texto legal, interpretación orientativa y datos que faltan. No presentes la respuesta como asesoramiento jurídico vinculante.

CONSULTA: {consulta}

FUENTES OFICIALES RECUPERADAS EN DIRECTO:
{contexto_legal}"""
                    res = llamar_ia(modelo_seleccionado, prompt_legal, _get_key("gemini"), incluir_contexto_global=False)
                    if "[ERROR" in res:
                        st.error(mensaje_error_amigable(res))
                    else:
                        st.session_state["normativa_res"] = res
                        st.session_state["normativa_fuentes"] = fuentes

        if st.session_state.get("normativa_res"):
            st.markdown("**Respuesta orientativa basada en legislación consolidada recuperada del BOE.**")
            st.markdown(st.session_state["normativa_res"])
            st.markdown("#### Fuentes oficiales consultadas")
            for fuente in st.session_state.get("normativa_fuentes", []):
                fecha_fuente = fuente["actualizado"]
                fecha_legible = f"{fecha_fuente[6:8]}/{fecha_fuente[4:6]}/{fecha_fuente[:4]}" if len(fecha_fuente) == 8 else "fecha no indicada"
                st.markdown(f"- [{fuente['titulo']}]({fuente['url']}) — actualización BOE: {fecha_legible}")
            st.caption("El texto consolidado del BOE es informativo. Para decisiones o procedimientos, confirma el texto oficial y la normativa autonómica aplicable.")
            fuentes_export = "\n\nFUENTES OFICIALES CONSULTADAS\n" + "\n".join(
                f"- {f['titulo']} | Actualización BOE: {f['actualizado'] or 'no indicada'} | {f['url']}"
                for f in st.session_state.get("normativa_fuentes", [])
            )
            botones_descarga(st.session_state["normativa_res"] + fuentes_export, "Consulta_Normativa")

    elif sub_idx == 2:
        st.markdown("### 🧾 Analizador Múltiple de Presupuestos")
        st.caption("Sube varios presupuestos (fotos o PDFs) a la vez. Los análisis se guardarán en la memoria hasta que reinicies o limpies.")

        arch_presupuestos = st.file_uploader("Sube los presupuestos (puedes seleccionar varios):", type=["jpg", "jpeg", "png", "pdf"], accept_multiple_files=True)
        incluir_reputacion = st.checkbox("Añadir búsqueda pública de reputación (opcional; resultados orientativos y con enlaces)", value=False, key="presu_incluir_reputacion")

        instrucciones_extra = st.text_area(
            "📝 ¿En qué quieres que se fije la IA? (Opcional):",
            placeholder="Ej: Céntrate en los plazos de entrega, busca si incluyen limpieza o gestión de escombros, revisa las garantías...",
            height=80
        )

        c_btn1, c_btn2 = st.columns(2)
        with c_btn1:
            if st.button("🔍 Analizar Presupuestos", type="primary") and arch_presupuestos:
                st.session_state["presu_resultados"] = None
                st.session_state["presu_comparativa"] = ""
                resultados_individuales = []
                errores_presupuestos = []
                progreso = st.progress(0)
                status_text = st.empty()

                for i, arch in enumerate(arch_presupuestos):
                    status_text.text(f"Analizando presupuesto {i+1} de {len(arch_presupuestos)}: {arch.name}...")

                    prompt_individual = """Analiza este presupuesto desde una perspectiva técnica y administrativa orientativa. No te presentes como perito ni como profesional que ha inspeccionado la obra. Separa los datos que aparecen literalmente de las observaciones; si un dato no está, escribe «No consta». No inventes importes, partidas, permisos ni garantías.

                    IMPORTANTE: Devuelve EXCLUSIVAMENTE un objeto JSON válido con dos claves:
                    1. "empresa": (Solo el nombre corto de la empresa, ej: "Pinturas Pepe").
                    2. "analisis": (El informe técnico redactado en texto plano, estructurado y muy profesional).

                    En el texto del "analisis", DEBES seguir estrictamente esta estructura formal, usando guiones o números para las listas:

                    INFORME TÉCNICO DE EVALUACIÓN

                    1. RESUMEN EJECUTIVO: Un párrafo de 2 o 3 líneas con el veredicto general sobre la oferta (viabilidad, riesgos principales).

                    2. ANÁLISIS ECONÓMICO:
                    - Precio Total y desglose del IVA (si lo lleva o está exento).
                    - Advertencia sobre posibles partidas ocultas, sobrecostes o falta de desglose.

                    3. CUMPLIMIENTO NORMATIVO Y LICENCIAS:
                    - Evaluación de las licencias de obra/actividad y tasas municipales. ¿Se encarga la empresa o lo deja a la comunidad?

                    4. SEGURIDAD Y SALUD (PRL) Y RESIDUOS:
                    - Indica qué documentación preventiva se menciona. No supongas que un plan específico es exigible en todos los casos; marca como «requiere revisión técnica según el tipo de obra» cuando no se pueda determinar.
                    - ¿Se especifica la gestión y retirada de escombros/residuos? Si no consta, indícalo expresamente.

                    5. PLAZOS Y GARANTÍAS:
                    - Plazos de ejecución.
                    - Garantías ofrecidas por escrito (seguro, años de garantía).

                    6. LA LETRA PEQUEÑA (PUNTOS CRÍTICOS):
                    - Lista de carencias, ambigüedades o cosas que NO incluye el presupuesto y que el presidente de la comunidad debería exigir antes de firmar.

                    7. RECOMENDACIONES PARA REVISIÓN:
                    - 2 o 3 acciones concretas a tomar antes de contratar a esta empresa.

                    No añadas markdown extra (como asteriscos) ni símbolos raros. Redacta con un tono serio, formal y jurídico."""

                    if instrucciones_extra.strip():
                        prompt_individual += f"\n\nINSTRUCCIONES ESPECÍFICAS DEL USUARIO A TENER EN CUENTA:\n{instrucciones_extra}\n"

                    try:
                        if arch.name.lower().endswith(".pdf"):
                            texto_pdf = leer_archivo_universal(arch, modelo_seleccionado)
                            if texto_pdf.strip():
                                res = llamar_ia(modelo_seleccionado, f"{prompt_individual}\n\nTexto del PDF:\n{texto_pdf[:MAX_TEXTO_IA]}", _get_key("gemini"))
                            else:
                                # PDF escaneado o vacío: intentar el analizador de adjuntos; si no hay OCR, no inventar datos.
                                res = analizar_archivo_con_gemini(arch.getvalue(), "application/pdf", prompt_individual, _get_key("gemini"), modelo_local=modelo_seleccionado)
                        else:
                            mime = "image/png" if arch.name.lower().endswith(".png") else "image/jpeg"
                            res = analizar_archivo_con_gemini(arch.getvalue(), mime, prompt_individual, _get_key("gemini"), modelo_local=modelo_seleccionado)

                        if "[ERROR" in res:
                            errores_presupuestos.append(arch.name)
                            progreso.progress((i + 1) / len(arch_presupuestos))
                            continue

                        res_limpia = res.replace("```json", "").replace("```", "").strip()
                        try:
                            data_json = json.loads(res_limpia)
                            nombre_empresa = data_json.get("empresa", arch.name)
                            texto_analisis = data_json.get("analisis", res)
                        except json.JSONDecodeError:
                            nombre_empresa = arch.name
                            texto_analisis = res

                        if incluir_reputacion:
                            status_text.text(f"Buscando referencias públicas para {nombre_empresa}...")
                            res_busqueda = buscar_en_internet(f'"{nombre_empresa}" opiniones trabajos realizados')
                            if res_busqueda:
                                fuentes_web = "\n".join(f"- {r['titulo']} | {r['url']} | {r['resumen']}" for r in res_busqueda)
                                prompt_rep = f"Evalúa con cautela estos resultados sobre la empresa '{nombre_empresa}'. No afirmes que corresponden a la misma empresa salvo que el propio resultado lo confirme claramente. Resume en un máximo de 3 líneas y conserva enlaces para revisión humana. Datos encontrados:\n{fuentes_web}"
                                rep_ia = llamar_ia(modelo_seleccionado, prompt_rep, _get_key("gemini"))
                                texto_analisis += "\n\nREFERENCIAS PÚBLICAS (pendientes de verificación de identidad):\n" + (rep_ia if "[ERROR" not in rep_ia else "No se pudo resumir la búsqueda; revisa las fuentes enlazadas.") + "\n\nFuentes consultadas:\n" + fuentes_web
                            else:
                                texto_analisis += "\n\nReferencias públicas: búsqueda sin resultados concluyentes."
                        else:
                            texto_analisis += "\n\nReferencias públicas: no consultadas."

                        resultados_individuales.append({"empresa": nombre_empresa, "archivo": arch.name, "analisis": texto_analisis})
                    except Exception:
                        errores_presupuestos.append(arch.name)

                    progreso.progress((i + 1) / len(arch_presupuestos))

                status_text.text("✅ Análisis individual completado. Generando comparativa final...")
                for archivo_error in errores_presupuestos:
                    st.warning(f"No se pudo analizar {archivo_error}; se excluyó del informe para evitar resultados incompletos o inventados.")

                if len(resultados_individuales) > 0:
                    texto_para_comparar = ""
                    for r in resultados_individuales:
                        texto_para_comparar += f"\n--- PROVEEDOR: {r['empresa']} (Archivo: {r['archivo']}) ---\n{r['analisis']}\n"

                    prompt_comparativa = f"""Redacta una comparativa técnica orientativa de ofertas para una comunidad de propietarios. No te presentes como perito ni completes datos ausentes por inferencia. Usa «No consta» cuando el documento no indique un importe, IVA, licencia, garantía o plazo. Distingue hechos del documento de recomendaciones.
                    Aquí tienes los informes periciales de {len(resultados_individuales)} presupuestos para una comunidad de vecinos:

                    {texto_para_comparar}

                    Tu tarea es redactar un INFORME COMPARATIVO FINAL EXTREMADAMENTE DETALLADO. No resumas demasiado. Sigue esta estructura exacta:

                    1. TABLA COMPARATIVA DETALLADA
                    Crea una tabla en Markdown con estas columnas: Proveedor, Presupuesto (€), IVA Desglosado (Sí/No), Licencias/PRL (Sí/No), Garantías (Sí/No), Valoración Técnica y Nivel de Riesgo (Bajo/Medio/Alto).

                    2. ANÁLISIS PORMENORIZADO POR EMPRESA
                    Haz un pequeño párrafo para cada empresa (usando su nombre como título). Explica qué ofrece, cuáles son sus puntos fuertes y qué carencias graves tiene respecto al resto.

                    3. ANÁLISIS DE RIESGOS (SEMÁFORO)
                    - RIESGOS CRÍTICOS: ¿Alguna empresa tiene un precio anormalmente bajo que sugiera "bicicleta" o falta de calidad? ¿Falta PRL obligatoria?
                    - RIESGOS MODERADOS: ¿Faltan detalles de garantías o licencias?
                    - PUNTOS A FAVOR: ¿Qué empresa da más confianza técnica?

                    4. RECOMENDACIÓN ORIENTATIVA
                    Explica con argumentos técnicos y jurídicos cuál es la mejor opción calidad/precio. Justifica si merece la pena pagar más por una empresa que cumpla estrictamente con la normativa.

                    5. ACCIONES ANTES DE FIRMAR
                    Lista de 3 a 5 exigencias que la comunidad debe pedir por escrito a la empresa ganadora antes de dar el visto bueno (ej: "Exigir el Plan de Seguridad y Salud", "Aclarar quién paga la tasa de escombros", etc.).

                    IMPORTANTE: Usa un tono formal, serio y profesional. No uses emojis ni símbolos raros. Formatea el texto para que sea fácil de leer en un PDF.
                    """

                    if instrucciones_extra.strip():
                        prompt_comparativa += f"\n\nINSTRUCCIONES ESPECÍFICAS DEL USUARIO A TENER EN CUENTA EN LA COMPARATIVA:\n{instrucciones_extra}\n"

                    res_comparativa = llamar_ia(modelo_seleccionado, prompt_comparativa, _get_key("gemini"))
                    status_text.empty()
                    progreso.empty()

                    if "[ERROR" not in res_comparativa:
                        st.session_state["presu_resultados"] = resultados_individuales
                        st.session_state["presu_comparativa"] = res_comparativa
                        st.success("✅ Informe generado y guardado en memoria.")
                        st.rerun()
                    else:
                        st.error("No se pudo generar la comparativa. Revisa las conexiones de IA y vuelve a intentarlo.")
                elif errores_presupuestos:
                    status_text.empty()
                    progreso.empty()
                    st.info("No se generó informe porque ningún presupuesto pudo analizarse correctamente.")

        with c_btn2:
            if st.session_state["presu_resultados"]:
                if st.button("🗑️ Limpiar Análisis Guardados", use_container_width=True):
                    st.session_state["presu_resultados"] = None
                    st.session_state["presu_comparativa"] = ""
                    st.rerun()

        if st.session_state["presu_resultados"]:
            st.markdown("---")
            st.markdown("### 🏆 Comparativa y Recomendación Final")
            st.markdown(st.session_state["presu_comparativa"])
            botones_descarga(st.session_state["presu_comparativa"], "Comparativa_Presupuestos", formatos_extra=True)

            st.link_button(
                "💬 Compartir comparativa por WhatsApp",
                wa_url(MENSAJES_WA["presupuesto"]),
                use_container_width=True,
            )

            st.markdown("---")
            st.markdown("### 📂 Informes Periciales Individuales")
            for idx, r in enumerate(st.session_state["presu_resultados"]):
                with st.expander(f"📄 {r['empresa']} (Archivo: {r['archivo']})"):
                    st.markdown(r['analisis'])
                    botones_descarga(r['analisis'], f"Analisis_{r['empresa'].replace(' ', '_')}", formatos_extra=True)

    elif sub_idx == 3:
        st.markdown("### 📊 Actas de Comunidad")
        tab_act = st.tabs(["📝 Crear Nueva", "🔍 Analizar / Preguntar"])

        with tab_act[0]:
            puntos_dia = st.text_area("Puntos del día a tratar (uno por línea):", height=150)
            if st.button("📝 Redactar Acta", type="primary") and puntos_dia:
                with st.spinner("Redactando..."):
                    prompt = f"Actúa como Secretario-Administrador de Fincas. Redacta un Acta de Reunión formal para una comunidad de vecinos. Los puntos del día a tratar son:\n{puntos_dia}\n\nSigue el formato legal español: Encabezado, Asistentes (deja espacio), Orden del día, Deliberaciones y Acuerdos, Cierre y Firmas."
                    res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["acta_crear_res"] = res
            
            if st.session_state.get("acta_crear_res"):
                st.text_area("Acta Redactada:", value=st.session_state["acta_crear_res"], height=400)
                botones_descarga(st.session_state["acta_crear_res"], "Acta_Comunidad")

        with tab_act[1]:
            st.markdown("Sube una o varias actas (fotos o PDFs) y pregúntale a la IA lo que necesites.")
            archs_acta = st.file_uploader("Sube los documentos (puedes seleccionar varios):", type=["pdf", "jpg", "jpeg", "png"], accept_multiple_files=True, key="up_acta")
            pregunta_acta = st.text_area("¿Qué quieres saber sobre estas actas?", placeholder="Ej: Resume los acuerdos, ¿qué se aprobó sobre la derrama?, mejora la redacción...", height=100)

            if archs_acta and st.button("🔍 Analizar Documentos", type="primary", key="btn_analizar_acta"):
                st.session_state.pop("acta_analisis_res", None)
                st.session_state["acta_analisis_errores"] = []
                with st.spinner("Analizando los documentos..."):
                    prompt_analisis = f"""Actúa como un Administrador de Fincas experto. Analiza el siguiente documento (Acta de Comunidad).
                    IMPORTANTE: Responde de forma clara y concisa a la instrucción del usuario. No inventes información que no esté en el documento.

                    INSTRUCCIÓN DEL USUARIO: {pregunta_acta if pregunta_acta.strip() else 'Haz un resumen ejecutivo y extrae los acuerdos principales.'}
                    """
                    
                    resultados_combinados = []
                    errores_documentos = []
                    progreso = st.progress(0)
                    
                    for i, arch in enumerate(archs_acta):
                        if arch.name.lower().endswith(".pdf"):
                            res = analizar_archivo_con_gemini(arch.getvalue(), "application/pdf", prompt_analisis, _get_key("gemini"), modelo_local=modelo_seleccionado)
                        else:
                            mime = "image/png" if arch.name.lower().endswith(".png") else "image/jpeg"
                            res = analizar_archivo_con_gemini(arch.getvalue(), mime, prompt_analisis, _get_key("gemini"), modelo_local=modelo_seleccionado)
                        
                        if "[ERROR" in res:
                            errores_documentos.append(arch.name)
                        else:
                            resultados_combinados.append(f"### Análisis de: {arch.name}\n\n{res.strip()}")
                            
                        progreso.progress((i + 1) / len(archs_acta))
                        
                    st.session_state["acta_analisis_res"] = "\n\n---\n\n".join(resultados_combinados)
                    st.session_state["acta_analisis_errores"] = errores_documentos
                    progreso.empty()

            for nombre_error in st.session_state.get("acta_analisis_errores", []):
                st.warning(f"No se pudo analizar {nombre_error}. Se omitió del informe; comprueba si el PDF tiene texto seleccionable o inténtalo más tarde.")

            if st.session_state.get("acta_analisis_res"):
                st.markdown(st.session_state["acta_analisis_res"])
                botones_descarga(st.session_state["acta_analisis_res"], "Analisis_Actas", formatos_extra=True)

    elif sub_idx == 4:
        st.markdown("### 📝 Citaciones y Notificaciones")
        tab_cit = st.tabs(["📝 Crear Nueva", "🔍 Analizar / Preguntar"])

        with tab_cit[0]:
            motivo = st.text_input("Motivo de la citación:", placeholder="Ej: Reunión extraordinaria por derrama...")
            fecha_hora = st.text_input("Fecha y Hora:", placeholder="Ej: 15 de Noviembre a las 19:00h")
            lugar = st.text_input("Lugar:", value="Portal del edificio o Sala de juntas")
            if st.button("📝 Redactar Citación", type="primary", key="btn_redactar_cit") and motivo:
                with st.spinner("Redactando..."):
                    prompt = f"Actúa como Administrador de Fincas. Redacta una notificación formal para los vecinos. Motivo: {motivo}. Fecha y hora: {fecha_hora}. Lugar: {lugar}. Incluye un lenguaje formal, la obligación de asistencia o delegación de voto, y espacio para la firma."
                    res = llamar_ia(modelo_seleccionado, prompt, _get_key("gemini"))
                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["citacion_crear_res"] = res
            
            if st.session_state.get("citacion_crear_res"):
                st.text_area("Citación:", value=st.session_state["citacion_crear_res"], height=300)
                botones_descarga(st.session_state["citacion_crear_res"], "Citacion_Vecinos")

        with tab_cit[1]:
            st.markdown("Sube una citación o notificación (foto o PDF) para que la IA la revise en busca de defectos de forma o errores legales.")
            arch_citacion = st.file_uploader("Sube el documento:", type=["pdf", "jpg", "jpeg", "png"], key="up_citacion")
            pregunta_cit = st.text_area("¿Qué quieres saber sobre esta citación?", placeholder="Ej: ¿Falta el orden del día?, ¿Es válida legalmente?, redáctala de nuevo corrigiendo errores...", height=100)

            if arch_citacion and st.button("🔍 Revisar Documento", type="primary", key="btn_analizar_cit"):
                with st.spinner("Analizando el documento..."):
                    prompt_revision = f"""Actúa como un Administrador de Fincas y Abogado experto en propiedad horizontal en España.
                    Tu objetivo es revisar la siguiente citación.

                    INSTRUCCIÓN DEL USUARIO: {pregunta_cit if pregunta_cit.strip() else 'Analiza los defectos de forma y propón una versión corregida.'}

                    REGLAS ESTRICTAS:
                    1. NO TE INVENTES FALLOS. Si el documento tiene fecha, hora, lugar y orden del día, indica que es correcto.
                    2. Si el usuario te ha pedido algo concreto, responde a eso.
                    3. Si propones una versión corregida, usa exactamente los datos que aparecen en el documento.

                    Estructura tu respuesta así:
                    VEREDICTO: (Es válido o tiene fallos reales)
                    ANÁLISIS DE FORMA: (Puntos clave que sí tiene o que faltan de verdad)
                    PROPUESTA DE MEJORA: (Si es necesario, redacta la versión final con los datos del documento)
                    """

                    if arch_citacion.name.lower().endswith(".pdf"):
                        res = analizar_archivo_con_gemini(arch_citacion.getvalue(), "application/pdf", prompt_revision, _get_key("gemini"), modelo_local=modelo_seleccionado)
                    else:
                        mime = "image/png" if arch_citacion.name.lower().endswith(".png") else "image/jpeg"
                        res = analizar_archivo_con_gemini(arch_citacion.getvalue(), mime, prompt_revision, _get_key("gemini"), modelo_local=modelo_seleccionado)

                    if "[ERROR" in res: st.error(mensaje_error_amigable(res))
                    else: st.session_state["citacion_res"] = res

            if st.session_state.get("citacion_res"):
                st.markdown(st.session_state["citacion_res"])
                botones_descarga(st.session_state["citacion_res"], "Revision_Citacion", formatos_extra=True)        
