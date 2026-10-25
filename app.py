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

# Fallback para el buscador (Resuelve el punto 1 del informe de Gemini)
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

st.set_page_config(page_title="Xavilu Network", page_icon="logo.png", layout="wide", initial_sidebar_state="expanded")

MAX_FILE_MB = 25
MAX_TEXTO_IA = 60_000

# --- ESTADO DE SESIÓN ---
_defaults = {
    "guion_podcast": "", "resultados_busqueda": [], "texto_archivo": "",
    "analisis_ia": "", "audio_podcast": None, "contexto_podcast": "", "tema_podcast": "",
    "chat_history": [], "img_url": "", "app_code": "", "video_script": "",
    "office_file_bytes": None, "office_file_ext": "", "office_file_mime": "",
    "pdf_bytes": None, "mermaid_code": "", "yt_resumen": "", "traduccion_audio": None,
    "video_bytes": None, "pdf_chat_text": "", "pdf_chat_history": [],
    "ideas_video": "", "texto_manual_generado": ""
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
st.html(
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
    """
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
    return "Formato no reconocido."

def llamar_ia(modelo, prompt, api_key_gemini=""):
    # Truncamos el prompt para evitar desbordar el contexto de Ollama (Resuelve el punto de ChatGPT)
    prompt = prompt[:MAX_TEXTO_IA]
    
    if api_key_gemini.strip() and HAS_GEMINI:
        try: 
            # Corregido el modelo a 1.5-flash que es el actual
            response = genai.Client(api_key=api_key_gemini).models.generate_content(
                model='gemini-1.5-flash', 
                contents=prompt
            )
            return response.text + "\n\n*(Procesado en la Nube con Gemini)*"
        except Exception as e: 
            pass # Si falla, cae a Ollama
            
    try: 
        client = ollama.Client(host="http://127.0.0.1:11434", timeout=300)
        response = client.chat(model=modelo, messages=[{"role": "user", "content": prompt}], options={"temperature": 0.2, "num_ctx": 8192, "repeat_penalty": 1.2})
        return response["message"]["content"].split("</think>")[-1].strip() + "\n\n*(Procesado en Local con Ollama)*"
    except Exception as e: 
        return f"[ERROR] No se pudo conectar con ninguna IA. Detalle: {e}"

def analizar_archivo_con_gemini(archivo_bytes, mime_type, prompt, api_key_gemini):
    if not HAS_GEMINI:
        return "[ERROR] La librería google-genai no está instalada."
    intentos = 3
    for i in range(intentos):
        try:
            client = genai.Client(api_key=api_key_gemini)
            parte_archivo = types.Part.from_bytes(data=archivo_bytes, mime_type=mime_type)
            response = client.models.generate_content(
                model='gemini-1.5-flash',
                contents=[prompt, parte_archivo]
            )
            return response.text + "\n\n*(Procesado en la Nube con Gemini)*"
        except Exception as e:
            if "503" in str(e) and i < intentos - 1:
                time.sleep(3)
                continue
            else:
                return f"[ERROR GEMINI] {e}"

def generar_pdf_pro(titulo, contenido):
    pdf = FPDF()
    pdf.add_page()
    
    # SOLUCIÓN AL PROBLEMA DE TILDES (Punto 3 de los informes)
    # Intentamos cargar una fuente de Windows que soporte Unicode
    font_path = "C:/Windows/Fonts/arial.ttf"
    if os.path.exists(font_path):
        pdf.add_font("Arial", "", font_path, uni=True)
        pdf.set_font("Arial", size=11)
    else:
        pdf.set_font("Helvetica", size=11)
        contenido = contenido.encode('latin-1', 'replace').decode('latin-1')

    pdf.set_font("Helvetica", "B", 24)
    pdf.cell(0, 40, txt="", ln=True)
    pdf.multi_cell(0, 10, txt=titulo, align='C')
    pdf.set_font("Helvetica", "I", 12)
    pdf.cell(0, 10, txt="Generado por Xavilu Network Pro", align='C', ln=True)
    pdf.add_page()
    
    if os.path.exists(font_path):
        pdf.set_font("Arial", size=11)
    else:
        pdf.set_font("Helvetica", size=11)
        
    pdf.multi_cell(0, 8, txt=contenido)
    return bytes(pdf.output())

def botones_descarga(texto, nombre_archivo="documento"):
    if texto:
        st.markdown("---")
        col_txt, col_pdf = st.columns(2)
        with col_txt:
            st.download_button(
                label="📥 Descargar (TXT)", 
                data=texto.encode('utf-8'), 
                file_name=f"{nombre_archivo}.txt", 
                mime="text/plain", 
                key=f"btn_txt_{nombre_archivo}"
            )
        with col_pdf:
            pdf_data = generar_pdf_pro(f"Reporte - {nombre_archivo.replace('_', ' ')}", texto)
            st.download_button(
                label="📄 Descargar (PDF)", 
                data=pdf_data, 
                file_name=f"{nombre_archivo}.pdf", 
                mime="application/pdf", 
                key=f"btn_pdf_{nombre_archivo}"
            )

# ... (Aquí continúan las funciones generar_word, generar_excel, generar_pptx, etc. sin cambios)
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
    
    try:
        api_key_gemini = st.secrets["GEMINI_API_KEY"]
        st.sidebar.success("🔑 API Key cargada automáticamente")
    except:
        api_key_gemini = st.text_input("Gemini API Key", type="password")
        
    st.markdown("---")
    st.markdown("<p style='font-size:12px; color:#00f0ff; text-align:center;'>SYSTEM STATUS: OPTIMAL</p>", unsafe_allow_html=True)

# --- 18 PESTAÑAS ---
tabs = st.tabs([
    "🌐 GLOBAL", "📁 DATOS", "🖼️ MULTIMEDIA", "🎙️ VOZ", "💬 CHAT", "🎨 VISUAL", 
    "📊 OFFICE", "📄 PDF", "🕸️ MAPAS", "🎬 VIDEO", "💻 CÓDIGO",
    "🎥 YOUTUBE", "📈 GRÁFICOS", "🌍 TRADUCTOR", 
    "📱 REDES SOCIALES", "💬 CHAT CON PDF", "📇 CV & QR", "🎯 PROMPTS"
])

# --- TAB 1: INTERNET ---
with tabs[0]:
    with st.expander("❓ ¿Cómo funciona esta sección? (Buscador Web)"):
        st.write("Busca información actualizada en internet. La IA leerá los resultados y te hará un informe descargable en TXT o PDF.")
    st.markdown("### 🌐 Información Global")
    query_busqueda = st.text_input("Comando de búsqueda en la red:", key="q_internet")
    if st.button("🔍 Rastrear Web", key="b_internet") and query_busqueda:
        with st.spinner("Rastreando..."): st.session_state.resultados_busqueda = buscar_en_internet(query_busqueda)
    if st.session_state.resultados_busqueda:
        for res in st.session_state.resultados_busqueda:
            with st.expander(f"🌐 {res['titulo']}"):
                st.write(res['resumen'])
    st.markdown("---")
    
    prompt_usuario = st.text_area("🎯 Instrucción para el núcleo neuronal:", height=80)
    if st.button("⚡ Ejecutar Análisis") and prompt_usuario:
        with st.spinner("Redactando..."):
            contexto = "\n".join([f"- {r['titulo']}: {r['resumen']}" for r in st.session_state.resultados_busqueda])
            st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, f"{prompt_usuario}\n\nContexto:\n{contexto}", api_key_gemini)
            
    if st.session_state.analisis_ia:
        st.text_area("Resultado del Análisis:", value=st.session_state.analisis_ia, height=300, key="ta_resultado_web")
        botones_descarga(st.session_state.analisis_ia, "Analisis_Web")

# --- TAB 2: ARCHIVOS ---
with tabs[1]:
    with st.expander("❓ ¿Cómo funciona esta sección? (Destripador de Documentos)"):
        st.write("Sube un PDF, Word o Excel. Puedes pedirle un resumen o usar el 'Análisis Forense' para que la IA busque datos ocultos. También puedes escribirle una orden concreta en la caja de abajo.")
    st.markdown("### 📁 Análisis Forense de Archivos")
    archivo = st.file_uploader("Sube cualquier archivo (PDF, Word, Excel, PPT, TXT)", type=["pdf", "txt", "docx", "xlsx", "pptx"])
    
    if archivo:
        if archivo.size > MAX_FILE_MB * 1024 * 1024:
            st.error(f"El archivo supera el límite de {MAX_FILE_MB} MB.")
        else:
            st.success("Archivo cargado. Configura el análisis:")
            
            tipo_analisis = st.radio("Selecciona el protocolo de análisis:", [
                "Análisis Forense (Destripar a fondo)", 
                "Resumen Ejecutivo Rápido", 
                "Extraer Datos, Cifras y Nombres", 
                "Generar Examen (Preguntas y Respuestas)"
            ])

            indicacion_extra = st.text_input("🎯 Indícale a la IA algo concreto (Opcional):", 
                                             placeholder="Ej: Sácame solo las fechas de vencimiento...")
            
            if st.button("🔎 Ejecutar Análisis Neuronal"):
                if "Forense" in tipo_analisis:
                    prompt_base = "Actúa como un analista de inteligencia experto. Destripa este documento por completo. Extrae: ideas principales, nombres propios, fechas, datos numéricos, contradicciones, y conclusiones no explícitas."
                elif "Rápido" in tipo_analisis:
                    prompt_base = "Haz un resumen ejecutivo conciso de qué trata este documento, cuál es su propósito y sus conclusiones principales."
                elif "Datos" in tipo_analisis:
                    prompt_base = "Extrae de forma estructurada todos los datos clave, nombres propios, fechas, tablas, estadísticas y cifras numéricas que aparezcan en este documento."
                else:
                    prompt_base = "Basándote en este documento, genera un cuestionario de 10 preguntas con 4 opciones cada una (tipo test) y sus respuestas correctas al final."

                if indicacion_extra.strip():
                    prompt_final = f"{prompt_base}\n\nINSTRUCCIÓN ESPECÍFICA ADICIONAL DEL USUARIO:\n{indicacion_extra}"
                else:
                    prompt_final = prompt_base

                if api_key_gemini.strip():
                    with st.spinner("Enviando archivo directamente a Gemini (Nube)..."):
                        nombre = archivo.name.lower()
                        if nombre.endswith(".pdf"): mime = "application/pdf"
                        else: mime = "text/plain"
                        
                        if nombre.endswith((".docx", ".xlsx", ".pptx", ".txt")):
                            texto_extraido = leer_archivo_universal(archivo, modelo_seleccionado)
                            st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, f"{prompt_final}\n\nCONTENIDO:\n{texto_extraido}", api_key_gemini)
                        else:
                            st.session_state.analisis_ia = analizar_archivo_con_gemini(
                                archivo.getvalue(), mime, prompt_final, api_key_gemini
                            )
                else:
                    with st.spinner("Extrayendo texto y procesando con Ollama (Local)..."):
                        # Limitamos a 4000 caracteres para no colapsar Ollama
                        st.session_state.texto_archivo = leer_archivo_universal(archivo, modelo_seleccionado)[:4000]
                        prompt_local = f"{prompt_final}\n\nCONTENIDO DEL DOCUMENTO:\n{st.session_state.texto_archivo}"
                        st.session_state.analisis_ia = llamar_ia(modelo_seleccionado, prompt_local, "")

                st.text_area("Resultado del Análisis:", value=st.session_state.analisis_ia, height=400)
                botones_descarga(st.session_state.analisis_ia, "Analisis_Documento")

# (El resto de pestañas continúan exactamente igual hasta la pestaña de Mapas)

# --- TAB 9: MAPAS (Asegurado contra XSS) ---
with tabs[8]:
    with st.expander("❓ ¿Cómo funciona esta sección? (Mapas Mentales)"):
        st.write("Escribe un tema y la IA creará un mapa mental o conceptual visual. Puedes descargarlo como código o como imagen PNG.")
    st.markdown("### 🕸️ Cartografía de Datos")
    tema_mapa = st.text_input("Nodo Central:", placeholder="Ej: Inteligencia Artificial...")
    tipo_mapa = st.radio("Protocolo de mapa:", ["Mapa Mental (Jerárquico)", "Mapa Conceptual (Flujo)"], horizontal=True)
    instrucciones_mapa = st.text_area("🎯 Instrucciones para el mapa (Opcional):", placeholder="Ej: Incluye al menos 5 subramas, hazlo muy técnico...", height=80)
    
    if st.button("Trazar Mapa") and tema_mapa:
        with st.spinner("Calculando nodos..."):
            if "Mental" in tipo_mapa:
                prompt_mapa = f"Genera el código Mermaid para un MAPA MENTAL sobre '{tema_mapa}'. Usa 'mindmap' y sangrías. "
            else:
                prompt_mapa = f"Genera el código Mermaid para un MAPA CONCEPTUAL sobre '{tema_mapa}'. Usa 'graph TD'. "
                
            if instrucciones_mapa.strip():
                prompt_mapa += f"Ten en cuenta estas instrucciones extra: {instrucciones_mapa}. "
                
            prompt_mapa += "Devuelve SOLO el código, sin explicaciones."
            codigo_bruto = llamar_ia(modelo_seleccionado, prompt_mapa, api_key_gemini)
            # Sanitizamos el código para evitar XSS
            codigo_limpio = codigo_bruto.replace("```mermaid", "").replace("```", "").strip()
            codigo_limpio = html.escape(codigo_limpio) # Elimina etiquetas HTML maliciosas
            st.session_state.mermaid_code = codigo_limpio

    if st.session_state.mermaid_code:
        st.html(
            f"""
            <div class="mermaid" style="background-color: white; padding: 20px; border-radius: 10px;">
            {st.session_state.mermaid_code}
            </div>
            <script type="module">
              import mermaid from 'https://cdn.jsdelivr.net/npm/mermaid@10/dist/mermaid.esm.min.mjs';
              mermaid.initialize({{ startOnLoad: true, theme: 'default' }});
            </script>
            """
        )
        
        st.markdown("---")
        col_mmd, col_png = st.columns(2)
        
        with col_mmd:
            st.download_button("📥 Descargar Código (.mmd)", st.session_state.mermaid_code.encode('utf-8'), "mapa.mmd", "text/plain")
            
        with col_png:
            b64_code = base64.urlsafe_b64encode(st.session_state.mermaid_code.encode('utf-8')).decode('utf-8')
            img_url = f"https://mermaid.ink/img/{b64_code}?bgColor=white"
            try:
                req = urllib.request.Request(img_url, headers={'User-Agent': 'Mozilla/5.0'})
                with urllib.request.urlopen(req) as response:
                    img_bytes = response.read()
                st.download_button("🖼️ Descargar Imagen (.png)", img_bytes, "mapa_visual.png", "image/png")
            except Exception as e:
                st.error(f"No se pudo generar la imagen. Prueba con el .mmd. Detalle: {e}")

# (El resto de pestañas como Video, YouTube, Gráficos, etc. se mantienen igual, asegurando 
# que todos los bloques 'try...except' estén cerrados correctamente).
