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
import csv
import streamlit.components.v1 as components
import os
import base64
import pandas as pd
import plotly.express as px
from youtube_transcript_api import YouTubeTranscriptApi
import re

try:
    from ddgs import DDGS
except ImportError:
    DDGS = None

try:
    from google import genai
    HAS_GEMINI = True
except ImportError:
    HAS_GEMINI = False

st.set_page_config(page_title="Xavilu Network", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")

MAX_FILE_MB = 25
MAX_TEXTO_IA = 60_000

# --- ESTADO DE SESIÓN ---
_defaults = {
    "guion_podcast": "", "resultados_busqueda": [], "texto_archivo": "",
    "analisis_ia": "", "audio_podcast": None, "contexto_podcast": "", "tema_podcast": "",
    "chat_history": [], "img_url": "", "app_code": "", "video_script": "",
    "office_file_bytes": None, "office_file_ext": "", "office_file_mime": "",
    "pdf_bytes": None, "mermaid_code": "", "yt_resumen": "", "traduccion_audio": None
}
for _k, _v in _defaults.items():
    if _k not in st.session_state:
        st.session_state[_k] = _v

# --- CSS CYBERPUNK / RED XAVILU ---
st.markdown(
    """
    <style>
    @import url('https://fonts.googleapis.com/css2?family=Orbitron:wght@400;700&family=Rajdhani:wght@500;700&display=swap');
    .stApp { background-color: #050505 !important; background-image: radial-gradient(circle at 50% 50%, #1a1a1a 0%, #050505 100%) !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; }
    [data-testid="stHeader"] { background-color: transparent !important; }
    [data-testid="stSidebarCollapsedControl"] { color: #d4af37 !important; background-color: #0b0d13 !important; border: 2px solid #d4af37 !important; border-radius: 50% !important; box-shadow: 0 0 15px rgba(212, 175, 55, 0.8) !important; display: flex !important; visibility: visible !important; }
    h1, h2, h3 { font-family: 'Orbitron', sans-serif !important; color: #f7dfa3 !important; text-shadow: 0 0 10px #d4af37, 0 0 20px #d4af37 !important; text-align: center; letter-spacing: 3px; }
    [data-testid="stWidgetLabel"] p { color: #00f0ff !important; font-family: 'Orbitron', sans-serif !important; font-size: 12px !important; font-weight: bold !important; text-transform: uppercase !important; letter-spacing: 1.5px !important; text-shadow: 0 0 8px rgba(0, 240, 255, 0.7) !important; background: rgba(0, 240, 255, 0.1) !important; padding: 4px 10px !important; border-left: 4px solid #00f0ff !important; border-radius: 4px !important; display: inline-block !important; margin-bottom: 5px !important; }
    p, .stMarkdown p { color: #d4af37 !important; font-size: 15px; }
    [data-testid="stSidebar"] { background-color: rgba(10, 10, 15, 0.8) !important; border-right: 1px solid #d4af37 !important; box-shadow: 5px 0 15px rgba(0,0,0,0.5) !important; backdrop-filter: blur(10px); }
    button[data-baseweb="tab"] { background: rgba(20, 20, 25, 0.7) !important; border: 1px solid #d4af37 !important; border-radius: 25px !important; color: #d4af37 !important; font-family: 'Orbitron', sans-serif !important; font-size: 11px !important; padding: 8px 12px !important; margin: 5px !important; box-shadow: 0 0 10px rgba(212, 175, 55, 0.2) !important; transition: all 0.3s ease; }
    button[data-baseweb="tab"]:hover { background: rgba(212, 175, 55, 0.2) !important; box-shadow: 0 0 15px rgba(212, 175, 55, 0.6) !important; }
    button[data-baseweb="tab"][aria-selected="true"] { background: linear-gradient(135deg, #d4af37 0%, #b8860b 100%) !important; color: #000 !important; border: 1px solid #fff !important; box-shadow: 0 0 20px rgba(212, 175, 55, 0.8) !important; }
    [data-baseweb="tab-list"] { border-bottom: 1px solid rgba(212, 175, 55, 0.3) !important; justify-content: flex-start; overflow-x: auto; white-space: nowrap; gap: 10px; }
    .stTextArea textarea, .stTextInput input, .stSelectbox [data-baseweb="select"], .stChatInput textarea { background-color: rgba(0, 0, 0, 0.8) !important; border: 1px solid #d4af37 !important; border-radius: 8px !important; color: #f7dfa3 !important; font-family: 'Rajdhani', sans-serif !important; font-size: 16px !important; box-shadow: inset 0 0 10px rgba(212, 175, 55, 0.1) !important; }
    .stTextArea textarea:focus, .stTextInput input:focus { border: 1px solid #00f0ff !important; box-shadow: 0 0 15px rgba(0, 240, 255, 0.4) !important; color: #00f0ff !important; }
    div.stButton > button, div[data-testid="stDownloadButton"] > button { background: rgba(0, 0, 0, 0.8) !important; color: #d4af37 !important; border: 1px solid #d4af37 !important; border-radius: 20px !important; font-family: 'Orbitron', sans-serif !important; font-weight: bold !important; box-shadow: 0 0 10px rgba(212, 175, 55, 0.3) !important; transition: all 0.3s ease; }
    div.stButton > button:hover, div[data-testid="stDownloadButton"] > button:hover { background: #d4af37 !important; color: #000 !important; box-shadow: 0 0 20px rgba(212, 175, 55, 0.8) !important; transform: translateY(-2px); }
    [data-testid="stChatMessage"] { background: rgba(10, 10, 15, 0.6) !important; border: 1px solid rgba(212, 175, 55, 0.5) !important; border-radius: 15px !important; box-shadow: 0 0 15px rgba(0,0,0,0.8) !important; backdrop-filter: blur(5px); }
    .streamlit-expanderHeader { background-color: rgba(20,20,20,0.8) !important; border-radius: 10px !important; color: #d4af37 !important; border: 1px solid rgba(212, 175, 55, 0.3) !important; }
    ::-webkit-scrollbar { width: 8px; }
    ::-webkit-scrollbar-track { background: #050505; }
    ::-webkit-scrollbar-thumb { background: #d4af37; border-radius: 10px; }
    </style>
    """,
    unsafe_allow_html=True,
)

# --- EFECTO DE SONIDO SCI-FI ---
components.html(
    """
    <script>
    const audioCtx = new (window.AudioContext || window.webkitAudioContext)();
    document.addEventListener('click', function(e) {
        if(e.target.closest('button')) {
            if (audioCtx.state === 'suspended') audioCtx.resume();
            const oscillator = audioCtx.createOscillator();
            const gainNode = audioCtx.createGain();
            oscillator.type = 'sine';
            oscillator.frequency.setValueAtTime(800, audioCtx.currentTime);
            oscillator.frequency.exponentialRampToValueAtTime(300, audioCtx.currentTime + 0.1);
            gainNode.gain.setValueAtTime(0.1, audioCtx.currentTime);
            gainNode.gain.exponentialRampToValueAtTime(0.01, audioCtx.currentTime + 0.1);
            oscillator.connect(gainNode);
            gainNode.connect(audioCtx.destination);
            oscillator.start();
            oscillator.stop(audioCtx.currentTime + 0.1);
        }
    });
    </script>
    """,
    height=0,
    width=0
)

# --- FUNCIONES BASE ---
def buscar_en_internet(query):
    if DDGS is None: return [{"titulo": "⚠️ Falta librería", "url": "", "resumen": "Instala: pip install ddgs"}]
    for region in ['es-es', 'us-en', 'wt-wt']:
        for motor in ['duckduckgo', 'bing', 'brave']:
            try:
                resultados_raw = DDGS().text(query, region=region, backend=motor, max_results=5)
                if resultados_raw:
                    return [{"titulo": r.get("title", ""), "url": r.get("href", ""), "resumen": r.get("body", "")} for r in resultados_raw]
            except: continue
    return [{"titulo": "⚠️ Bloqueo temporal", "url": "", "resumen": "Espera 1 minuto."}]

def leer_archivo_universal(upload, modelo_ia):
    nombre = upload.name.lower()
    if nombre.endswith(".txt"): return upload.read().decode("utf-8", errors="ignore")
    if nombre.endswith(".pdf"): return "\n".join([p.extract_text() for p in pypdf.PdfReader(BytesIO(upload.getvalue())).pages if p.extract_text()])
    if nombre.endswith(".docx"): return "\n".join([p.text for p in Document(upload).paragraphs if p.text])
    if nombre.endswith(".xlsx"):
        wb = openpyxl.load_workbook(upload, data_only=True)
        return "\n".join([" | ".join([str(c) for c in f if c is not None]) for h in wb.sheetnames for f in wb[h].iter_rows(values_only=True) if any(c is not None for c in f)])
    if nombre.endswith(".pptx"):
        return "\n".join([sh.text for s in Presentation(upload).slides for sh in s.shapes if hasattr(sh, "text") and sh.text])
    
    if nombre.endswith((".png", ".jpg", ".jpeg", ".webp")):
        try:
            imagen = Image.open(upload)
            out_img = BytesIO()
            imagen.save(out_img, format="PNG")
            bytes_img = out_img.getvalue()
            client = ollama.Client(host="http://127.0.0.1:11434")
            mod_vision = modelo_ia if "vision" in modelo_ia or "bakllava" in modelo_ia else "llama3.2-vision:latest"
            res = client.chat(model=mod_vision, messages=[{'role': 'user', 'content': 'Extrae absolutamente todo el texto y describe los elementos visuales de esta imagen para un análisis forense.', 'images': [bytes_img]}])
            return f"[Texto extraído de la imagen]: {res['message']['content']}"
        except Exception as e: 
            return f"Error procesando imagen: {e}. (Asegúrate de tener un modelo de visión como llama3.2-vision en Ollama o usa Gemini)."
            
    return "Formato no reconocido."

def llamar_ia(modelo, prompt, api_key_gemini=""):
    prompt = prompt[:MAX_TEXTO_IA]
    if api_key_gemini.strip() and HAS_GEMINI:
        try: return genai.Client(api_key=api_key_gemini).models.generate_content(model='gemini-2.5-flash', contents=prompt).text
        except: pass
    try: 
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=300)
        response = client.chat(model=modelo, messages=[{"role": "user", "content": prompt}], options={"temperature": 0.2, "num_ctx": 8192, "repeat_penalty": 1.2})
        return response["message"]["content"].split("</think>")[-1].strip()
    except Exception as e: return f"[ERROR LOCAL] {e}"

# --- GENERADORES DE ARCHIVOS ---
def generar_pdf_pro(titulo, contenido):
    pdf = FPDF()
    pdf.add_page()
    pdf.set_font("Helvetica", "B", 24)
    pdf.cell(0, 40, txt="", ln=True)
    pdf.multi_cell(0, 10, txt=titulo, align='C')
    pdf.set_font("Helvetica", "I", 12)
    pdf.cell(0, 10, txt="Generado por Xavilu Network Pro", align='C', ln=True)
    pdf.add_page()
    pdf.set_font("Helvetica", size=11)
    pdf.multi_cell(0, 8, txt=contenido.encode('latin-1', 'replace').decode('latin-1'))
    return bytes(pdf.output())

def generar_word(texto):
    doc = Document()
    doc.add_heading("Documento Generado por Xavilu", 0)
    for p in texto.split("\n"):
        if p.strip(): 
            if p.startswith("- ") or p.startswith("* "): doc.add_paragraph(p[2:], style='List Bullet')
            else: doc.add_paragraph(p.strip())
    out = BytesIO(); doc.save(out); out.seek(0)
    return out.getvalue()

def generar_excel(texto_csv):
    wb = openpyxl.Workbook()
    ws = wb.active
    reader = csv.reader(texto_csv.replace("```csv", "").replace("```", "").strip().split('\n'))
    for row_idx, row in enumerate(reader, 1):
        for col_idx, val in enumerate(row, 1):
            cell = ws.cell(row=row_idx, column=col_idx, value=val)
            if row_idx == 1: cell.font = Font(bold=True)
    out = BytesIO(); wb.save(out); out.seek(0)
    return out.getvalue()

def generar_pptx(titulo, texto):
    prs = Presentation()
    slide = prs.slides.add_slide(prs.slide_layouts[0])
    slide.shapes.title.text = titulo
    for bloque in [b.strip() for b in texto.split("\n\n") if b.strip()]:
        slide = prs.slides.add_slide(prs.slide_layouts[1])
        lineas = bloque.split("\n")
        slide.shapes.title.text = lineas[0][:50]
        if len(slide.placeholders) > 1: slide.placeholders[1].text = "\n".join(lineas[1:]) if len(lineas) > 1 else ""
    out = BytesIO(); prs.save(out); out.seek(0)
    return out.getvalue()

# --- SÍNTESIS DE VOZ ---
async def _ejecutar_sintesis_edge(texto_guion, voz_mod, voz_exp=None):
    audio_completo = BytesIO()
    if not voz_exp: 
        try:
            async for chunk in edge_tts.Communicate(texto_guion, voz_mod).stream():
                if chunk["type"] == "audio": audio_completo.write(chunk["data"])
        except: pass
    else: 
        for linea in texto_guion.split("\n"):
            if not linea.strip() or ":" not in linea: continue
            rol, _, texto = linea.partition(":")
            voice_actual = voz_mod if "mod" in rol.lower() else voz_exp
            try:
                async for chunk in edge_tts.Communicate(texto.strip(), voice_actual).stream():
                    if chunk["type"] == "audio": audio_completo.write(chunk["data"])
            except: continue
    audio_completo.seek(0)
    return audio_completo.getvalue()

def generar_audio_sincrono(texto_guion, voz_mod, voz_exp=None):
    loop = asyncio.new_event_loop(); asyncio.set_event_loop(loop)
    try: return loop.run_until_complete(_ejecutar_sintesis_edge(texto_guion, voz_mod, voz_exp))
    except: return None
    finally: loop.close()

# =========================================================
# INTERFAZ PRINCIPAL - HUD XAVILU
# =========================================================
st.markdown("""
    <div style="text-align: center; margin-bottom: 20px;">
        <h1 style="margin-bottom: 0;">XAVILU NETWORK</h1>
        <p style="color: #00f0ff; font-family: 'Orbitron'; letter-spacing: 2px; font-size: 12px; text-shadow: none; background: none; border: none; padding: 0;">
            <span style="color: #ff003c;">●</span> RED XAVILU ONLINE
        </p>
    </div>
""", unsafe_allow_html=True)

with st.sidebar:
    if os.path.exists("logo.png"):
        with open("logo.png", "rb") as image_file:
            encoded_string = base64.b64encode(image_file.read()).decode()
        st.markdown(f'<div style="display: flex; justify-content: center; margin-bottom: 20px;"><img src="data:image/png;base64,{encoded_string}" style="width: 150px; height: 150px; border-radius: 50%; border: 3px solid #d4af37; box-shadow: 0 0 20px rgba(212, 175, 55, 0.8); object-fit: cover;"></div>', unsafe_allow_html=True)
    else:
        st.markdown("<h1 style='text-align: center; color: #ff003c;'>X</h1>", unsafe_allow_html=True)

    st.markdown("<h3 style='text-align:center;'>🪐 PANEL DE CONTROL</h3>", unsafe_allow_html=True)
    modelo_seleccionado = st.selectbox("Motor de IA", ["llama3.2:latest", "llama3.2-vision:latest", "gemma2:2b", "deepseek-r1:1.5b"])
    api_key_gemini = st.text_input("Gemini API Key", type="password")
    st.markdown("---")
    st.markdown("<p style='font-size:12px; color:#00f0ff; text-align:center;'>SYSTEM STATUS: OPTIMAL</p>", unsafe_allow_html=True)

# --- 13 PESTAÑAS ---
tabs = st.tabs([
    "🌐 GLOBAL", "📁 DATOS", "🎙️ VOZ", "💬 CHAT", "🎨 VISUAL", 
    "📊 OFFICE", "📄 PDF", "🕸️ MAPAS", "🎬 VIDEO", "💻 CÓDIGO",
    "🎥 YOUTUBE", "📈 GRÁFICOS", "🌍 TRADUCTOR"
])

# --- TAB 1: INTERNET ---
with tabs[0]:
    st.markdown("### 🌐 Información Global")
    query_busqueda = st.text_input("Comando de búsqueda en la red:", key="q_internet")
    if st.button("🔍 Rastrear Web", key="b_internet") and query_busqueda:
        with st.spinner("Rastreando..."): st.session_state.resultados_busqueda = buscar_en_internet(query_busqueda)
    if st.session_state.resultados_busqueda:
        for res in st.session_state.resultados_busqueda:
            with st.expander(f"🌐 {res['titulo']}"):
                st.write(res['resumen'])
    st.markdown("---")
    prompt_usuario = st.text_area("Instrucción para el núcleo neuronal:", height=80)
    if st.button("⚡ Ejecutar Análisis") and prompt_usuario:
        with st.spinner("Redactando..."):
            contexto = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in st.session_state.resultados_busqueda])
            st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, f"{prompt_usuario}\n\nContexto:\n{contexto}", api_key_gemini)
            st.markdown(st.session_state.analisis_ia)

# --- TAB 2: ARCHIVOS (EL DESTRIPADOR CON INDICACIONES) ---
with tabs[1]:
    st.markdown("### 📁 Análisis Forense de Archivos")
    archivo = st.file_uploader("Sube cualquier archivo (PDF, Word, Excel, PPT, TXT, Imágenes)", type=["pdf", "txt", "docx", "xlsx", "pptx", "png", "jpg", "jpeg"])
    
    if archivo:
        if archivo.size > MAX_FILE_MB * 1024 * 1024:
            st.error(f"El archivo supera el límite de {MAX_FILE_MB} MB.")
        else:
            with st.spinner("Extrayendo información del archivo..."):
                st.session_state.texto_archivo = leer_archivo_universal(archivo, modelo_seleccionado)[:MAX_TEXTO_IA]
            
            st.success("Archivo asimilado. ¿Cómo quieres destriparlo?")
            
            # MENÚ DE ANÁLISIS
            tipo_analisis = st.radio("Selecciona el protocolo de análisis:", [
                "Análisis Forense (Destripar a fondo)", 
                "Resumen Ejecutivo Rápido", 
                "Extraer Datos, Cifras y Nombres", 
                "Generar Examen (Preguntas y Respuestas)"
            ])

            # NUEVA CAJA DE INDICACIONES CONCRETAS
            indicacion_extra = st.text_input("🎯 Indícale a la IA algo concreto (Opcional):", 
                                             placeholder="Ej: Sácame solo las fechas de vencimiento, o céntrate en la parte financiera...")
            
            if st.button("🔎 Ejecutar Análisis Neuronal"):
                with st.spinner("Procesando documento..."):
                    # 1. Generamos el prompt según el radio button
                    if "Forense" in tipo_analisis:
                        prompt_base = (
                            f"Actúa como un analista de inteligencia experto. Tu objetivo es destripar este documento por completo.\n"
                            f"Extrae y analiza: ideas principales, nombres propios, fechas, datos numéricos, contradicciones, y conclusiones no explícitas.\n\n"
                            f"CONTENIDO DEL DOCUMENTO:\n{st.session_state.texto_archivo}"
                        )
                    elif "Rápido" in tipo_analisis:
                        prompt_base = f"Haz un resumen ejecutivo conciso de qué trata este documento, cuál es su propósito y sus conclusiones principales:\n\n{st.session_state.texto_archivo}"
                    elif "Datos" in tipo_analisis:
                        prompt_base = f"Extrae de forma estructurada todos los datos clave, nombres propios, fechas, tablas, estadísticas y cifras numéricas que aparezcan en este texto:\n\n{st.session_state.texto_archivo}"
                    else:
                        prompt_base = f"Basándote en este texto, genera un cuestionario de 10 preguntas con 4 opciones cada una (tipo test) y sus respuestas correctas al final:\n\n{st.session_state.texto_archivo}"

                    # 2. Si el usuario ha escrito una indicación extra, se la inyectamos al prompt
                    if indicacion_extra.strip():
                        prompt_final = f"{prompt_base}\n\nINSTRUCCIÓN ESPECÍFICA ADICIONAL DEL USUARIO (Prioriza esto):\n{indicacion_extra}"
                    else:
                        prompt_final = prompt_base

                    # 3. Llamamos a la IA
                    st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt_final, api_key_gemini)
                
                st.text_area("Resultado del Análisis:", value=st.session_state.analisis_ia, height=400)
                st.download_button("📥 Descargar Análisis (TXT)", st.session_state.analisis_ia.encode('utf-8'), "analisis_documento.txt", "text/plain")

# --- TAB 3: PODCAST ---
with tabs[2]:
    st.markdown("### 🎙️ Sintetizador de Voz")
    VOCES = {"es-ES-ElviraNeural": "Elvira (F)", "es-MX-JorgeNeural": "Jorge (M)"}
    v_mod = st.selectbox("Modulador (Moderador)", list(VOCES.keys()), format_func=lambda x: VOCES[x], index=0)
    v_exp = st.selectbox("Experto (Conocimiento)", list(VOCES.keys()), format_func=lambda x: VOCES[x], index=1)
    
    opcion = st.radio("Fuente de datos:", ["Red Global", "Archivo Local"], horizontal=True)
    if opcion == "Red Global":
        tema = st.text_input("Tema de investigación:")
        if st.button("🛰️ Obtener Contexto") and tema:
            st.session_state.contexto_podcast = "\n".join([r['resumen'] for r in buscar_en_internet(tema)])
    else:
        f_pod = st.file_uploader("Sube archivo", key="f_pod")
        if f_pod and st.button("Obtener Contexto"):
            st.session_state.contexto_podcast = leer_archivo_universal(f_pod, modelo_seleccionado)

    if st.session_state.contexto_podcast:
        if st.button("📝 Generar Guion de Voz"):
            st.session_state.guion_podcast = llamar_ia(modelo_seleccionado, f"Guion en ESPAÑOL entre 'Moderador:' y 'Experto:'. Contexto: {st.session_state.contexto_podcast[:3000]}", api_key_gemini)
        if st.session_state.guion_podcast:
            st.text_area("Guion:", value=st.session_state.guion_podcast, height=200, key="ta_guion")
            if st.button("🔊 Sintetizar Audio"):
                st.session_state.audio_podcast = generar_audio_sincrono(st.session_state.ta_guion, v_mod, v_exp)
        if st.session_state.audio_podcast:
            st.audio(st.session_state.audio_podcast, format="audio/mp3")
            st.download_button("📥 Descargar MP3", st.session_state.audio_podcast, "podcast.mp3", "audio/mpeg")

# --- TAB 4: CHAT ---
with tabs[3]:
    st.markdown("### 💬 Interfaz de Comunicación Directa")
    for msg in st.session_state.chat_history:
        st.chat_message(msg["role"]).write(msg["content"])
    if prompt := st.chat_input("Escribe tu mensaje al núcleo neuronal..."):
        st.session_state.chat_history.append({"role": "user", "content": prompt})
        st.chat_message("user").write(prompt)
        with st.spinner("Procesando respuesta..."):
            historial = "\n".join([f"{m['role']}: {m['content']}" for m in st.session_state.chat_history[-6:]])
            res = llamar_ia(modelo_seleccionado, historial, api_key_gemini)
            st.session_state.chat_history.append({"role": "assistant", "content": res})
            st.chat_message("assistant").write(res)

# --- TAB 5: IMÁGENES ---
with tabs[4]:
    st.markdown("### 🎨 Generador de Recreaciones Visuales")
    prompt_img = st.text_input("Describe el escenario a generar:")
    if st.button("Generar Visualización") and prompt_img:
        with st.spinner("Renderizando..."):
            st.session_state.img_url = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(prompt_img)}?width=1024&height=1024&nologo=true"
    if st.session_state.img_url:
        st.image(st.session_state.img_url, use_column_width=True)
        st.markdown(f"[Descargar Visualización Original]({st.session_state.img_url})")

# --- TAB 6: OFFICE ---
with tabs[5]:
    st.markdown("### 📊 Creación de Documentos Office")
    prompt_office = st.text_area("Define la estructura del documento:", height=80)
    formato = st.selectbox("Formato de exportación:", ["Word (.docx)", "Excel (.xlsx)", "PowerPoint (.pptx)"])
        
    if st.button("Compilar Documento") and prompt_office:
        with st.spinner(f"Generando {formato}..."):
            if "Excel" in formato:
                prompt_ia = f"Genera una tabla con los datos para: '{prompt_office}'. Devuelve ÚNICAMENTE el texto en formato CSV."
                csv_text = llamar_ia(modelo_seleccionado, prompt_ia, api_key_gemini)
                st.session_state.office_file_bytes = generar_excel(csv_text)
                st.session_state.office_file_ext = "xlsx"
                st.session_state.office_file_mime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
            else:
                prompt_ia = f"Genera el contenido completo para: '{prompt_office}'."
                texto_generado = llamar_ia(modelo_seleccionado, prompt_ia, api_key_gemini)
                if "Word" in formato:
                    st.session_state.office_file_bytes = generar_word(texto_generado)
                    st.session_state.office_file_ext = "docx"
                    st.session_state.office_file_mime = "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
                elif "PowerPoint" in formato:
                    st.session_state.office_file_bytes = generar_pptx(prompt_office, texto_generado)
                    st.session_state.office_file_ext = "pptx"
                    st.session_state.office_file_mime = "application/vnd.openxmlformats-officedocument.presentationml.presentation"
            st.success("¡Documento Compilado!")
            
    if st.session_state.office_file_bytes:
        st.download_button(f"📥 Descargar {st.session_state.office_file_ext.upper()}", st.session_state.office_file_bytes, f"doc.{st.session_state.office_file_ext}", st.session_state.office_file_mime)

# --- TAB 7: PDF PRO ---
with tabs[6]:
    st.markdown("### 📄 Protocolo de Creación de PDF")
    titulo_pdf = st.text_input("Título del documento:")
    contenido_pdf = st.text_area("Datos a incluir:", height=100)
    
    if st.button("Generar PDF") and titulo_pdf and contenido_pdf:
        with st.spinner("Maquetando..."):
            prompt_pdf = f"Escribe un documento estructurado sobre: {contenido_pdf}."
            texto_final = llamar_ia(modelo_seleccionado, prompt_pdf, api_key_gemini)
            st.session_state.pdf_bytes = generar_pdf_pro(titulo_pdf, texto_final)
            st.success("¡PDF generado!")
            st.text_area("Vista previa:", value=texto_final, height=150, disabled=True)

    if st.session_state.pdf_bytes:
        st.download_button("📥 Descargar PDF", st.session_state.pdf_bytes, f"{titulo_pdf}.pdf", "application/pdf")

# --- TAB 8: MAPAS ---
with tabs[7]:
    st.markdown("### 🕸️ Cartografía de Datos")
    tema_mapa = st.text_input("Nodo Central:", placeholder="Ej: Inteligencia Artificial...")
    tipo_mapa = st.radio("Protocolo de mapa:", ["Mapa Mental (Jerárquico)", "Mapa Conceptual (Flujo)"], horizontal=True)
    
    if st.button("Trazar Mapa") and tema_mapa:
        with st.spinner("Calculando nodos..."):
            if "Mental" in tipo_mapa:
                prompt_mapa = f"Genera el código Mermaid para un MAPA MENTAL sobre '{tema_mapa}'. Usa 'mindmap' y sangrías. Devuelve SOLO el código."
            else:
                prompt_mapa = f"Genera el código Mermaid para un MAPA CONCEPTUAL sobre '{tema_mapa}'. Usa 'graph TD'. Devuelve SOLO el código."
            st.session_state.mermaid_code = llamar_ia(modelo_seleccionado, prompt_mapa, api_key_gemini).replace("```mermaid", "").replace("```", "").strip()

    if st.session_state.mermaid_code:
        components.html(
            f"""
            <div class="mermaid" style="background-color: white; padding: 20px; border-radius: 10px;">
            {st.session_state.mermaid_code}
            </div>
            <script type="module">
              import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
              mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
            </script>
            """, height=600,
        )
        st.download_button("📥 Descargar Código", st.session_state.mermaid_code.encode('utf-8'), "mapa.mmd", "text/plain")

# --- TAB 9: VÍDEO ---
with tabs[8]:
    st.markdown("### 🎬 Simulador de Escenas")
    tema_video = st.text_input("Concepto de la escena:")
    if st.button("Generar Storyboard") and tema_video:
        with st.spinner("Escribiendo..."):
            st.session_state.video_script = llamar_ia(modelo_seleccionado, f"Divide '{tema_video}' en 3 escenas: 'ESCENA X: [Descripción visual en INGLÉS]' y 'Narrador: [Texto en ESPAÑOL]'.", api_key_gemini)
    if st.session_state.video_script:
        st.text_area("Guion de Escenas:", value=st.session_state.video_script, height=200)
        if st.button("🖼️ Renderizar Fotogramas"):
            escenas = re.findall(r'ESCENA \d+:\s*(.*)', st.session_state.video_script)
            for i, escena in enumerate(escenas, 1):
                url_img = f"https://image.pollinations.ai/prompt/{urllib.parse.quote(escena.strip())}?width=800&height=450&nologo=true"
                st.image(url_img, caption=f"Fotograma {i}")

# --- TAB 10: APP BUILDER ---
with tabs[9]:
    st.markdown("### 💻 Asistente de Código (Fábrica de Apps)")
    req_app = st.text_area("Especificaciones de la nueva aplicación:", placeholder="Ej: Una app para calcular propinas...", height=80)
    
    if st.button("Generar Código Base") and req_app:
        with st.spinner("Compilando..."):
            prompt_code = f"Eres un desarrollador experto en Python y Streamlit. Crea el código completo para: '{req_app}'. REGLAS: Usa 'import streamlit as st'. Devuelve ÚNICAMENTE el código Python sin bloques markdown."
            codigo = llamar_ia(modelo_seleccionado, prompt_code, api_key_gemini)
            st.session_state.app_code = codigo.replace("```python", "").replace("```", "").strip()
            
    if st.session_state.app_code:
        st.success("Código fuente compilado.")
        st.code(st.session_state.app_code, language="python")
        st.download_button("📥 Descargar Código (.py)", st.session_state.app_code.encode('utf-8'), "mi_nueva_app.py", "text/plain")

# --- TAB 11: YOUTUBE ---
with tabs[10]:
    st.markdown("### 🎥 Analizador de Vídeos de YouTube")
    url_yt = st.text_input("Pega la URL del vídeo (Ej: https://www.youtube.com/watch?v=...):")
    
    if st.button("Extraer y Resumir") and url_yt:
        with st.spinner("Buscando transcripción y resumiendo..."):
            try:
                video_id = url_yt.split("v=")[-1].split("&")[0] if "v=" in url_yt else url_yt.split("/")[-1]
                transcript = YouTubeTranscriptApi.get_transcript(video_id, languages=['es', 'en'])
                texto_completo = " ".join([t['text'] for t in transcript])
                
                prompt_yt = f"Haz un resumen ejecutivo muy detallado, con los puntos clave y conclusiones, del siguiente texto extraído de un vídeo:\n\n{texto_completo[:5000]}"
                st.session_state.yt_resumen = llamar_ia(modelo_seleccionado, prompt_yt, api_key_gemini)
                st.success("¡Análisis completado!")
            except Exception as e:
                st.error(f"Error al extraer el vídeo: {e}")

    if st.session_state.yt_resumen:
        st.text_area("Resumen Ejecutivo:", value=st.session_state.yt_resumen, height=300)
        st.download_button("📥 Descargar Resumen (TXT)", st.session_state.yt_resumen.encode('utf-8'), "resumen_youtube.txt", "text/plain")

# --- TAB 12: GRÁFICOS ---
with tabs[11]:
    st.markdown("### 📈 Generador de Gráficos Inteligente")
    archivo_datos = st.file_uploader("Sube un archivo CSV o Excel:", type=["csv", "xlsx"])
    
    if archivo_datos:
        df = pd.read_csv(archivo_datos) if archivo_datos.name.endswith('.csv') else pd.read_excel(archivo_datos)
        st.dataframe(df.head(5))
        
        instruccion_grafico = st.text_input("¿Qué gráfico quieres? (Ej: Compara las ventas por mes):")
        if st.button("Generar Gráfico") and instruccion_grafico:
            with st.spinner("La IA está analizando tus datos..."):
                cols = df.columns.tolist()
                prompt_graf = f"Tengo un DataFrame con estas columnas: {cols}. El usuario quiere: '{instruccion_grafico}'. Dime qué gráfico usar (bar, line, pie, scatter) y qué columna va en X y cuál en Y. Devuelve SOLO un formato estricto: tipo_grafico|x_col|y_col. Si no aplica Y, pon 'none'."
                respuesta_ia = llamar_ia(modelo_seleccionado, prompt_graf, api_key_gemini)
                
                try:
                    partes = respuesta_ia.strip().split('|')
                    tipo = partes[0].strip().lower()
                    x_col = partes[1].strip()
                    y_col = partes[2].strip()
                    
                    if "bar" in tipo: fig = px.bar(df, x=x_col, y=y_col, title=instruccion_grafico)
                    elif "line" in tipo: fig = px.line(df, x=x_col, y=y_col, title=instruccion_grafico)
                    elif "pie" in tipo: fig = px.pie(df, names=x_col, values=y_col, title=instruccion_grafico)
                    else: fig = px.scatter(df, x=x_col, y=y_col, title=instruccion_grafico)
                    
                    fig.update_layout(template="plotly_dark")
                    st.plotly_chart(fig, use_container_width=True)
                except:
                    st.warning("La IA no pudo determinar las columnas. Prueba a ser más específico.")

# --- TAB 13: TRADUCTOR ---
with tabs[12]:
    st.markdown("### 🌍 Traductor y Locutor Multilingüe")
    texto_trad = st.text_area("Texto a traducir:", height=100)
    
    IDIOMAS = {
        "Inglés": "en-US-JennyNeural",
        "Francés": "fr-FR-DeniseNeural",
        "Alemán": "de-DE-KatjaNeural",
        "Italiano": "it-IT-ElsaNeural",
        "Portugués": "pt-BR-FranciscaNeural"
    }
    idioma_dest = st.selectbox("Idioma de destino:", list(IDIOMAS.keys()))
    
    if st.button("Traducir y Generar Audio") and texto_trad:
        with st.spinner(f"Traduciendo al {idioma_dest}..."):
            prompt_trad = f"Traduce el siguiente texto al {idioma_dest}. Devuelve ÚNICAMENTE la traducción, sin comillas ni explicaciones:\n\n{texto_trad}"
            traduccion = llamar_ia(modelo_seleccionado, prompt_trad, api_key_gemini)
            st.success("Traducción completada:")
            st.info(traduccion)
            
            with st.spinner("Sintetizando voz nativa..."):
                voz_sel = IDIOMAS[idioma_dest]
                st.session_state.traduccion_audio = generar_audio_sincrono(traduccion, voz_sel)
                
    if st.session_state.traduccion_audio:
        st.audio(st.session_state.traduccion_audio, format="audio/mp3")
        st.download_button("📥 Descargar Audio", st.session_state.traduccion_audio, f"audio_{idioma_dest}.mp3", "audio/mpeg")
