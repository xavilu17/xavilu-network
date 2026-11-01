"""
Xavilu Network Pro - app.py
App Streamlit completa con:
 - Menú de 9 categorías
 - Fallback de IAs: Gemini, Groq, SiliconFlow, Mistral, OpenRouter, DeepSeek, Together, Cerebras, Ollama
 - Generador de letras + reproductor (edge-tts + pydub + imageio-ffmpeg), Modo Dúo, 8 estilos
 - Limpiador de texto para voz
 - Documentos PDF con portada personalizada
 - Gestor de Vecinos (PRL) con calculadora de mayorías, Alex IA Legal y generador de actas
"""

import os
import re
import io
import json
import time
import asyncio
import tempfile
import base64
from datetime import datetime
from typing import Optional, List

import streamlit as st

st.set_page_config(page_title="Xavilu Network Pro", page_icon="🎸", layout="wide")

import requests

# ---------------------------------------------------------------------------
# Imports opcionales (manejados con try/except para que la app no rompa)
# ---------------------------------------------------------------------------
try:
    from pydub import AudioSegment
    from pydub.generators import Sine, WhiteNoise
    import imageio_ffmpeg
    try:
        AudioSegment.converter = imageio_ffmpeg.get_ffmpeg_exe()
    except Exception:
        pass
    PYDUB_OK = True
except Exception:
    PYDUB_OK = False

try:
    import edge_tts
    EDGE_TTS_OK = True
except Exception:
    EDGE_TTS_OK = False

try:
    from reportlab.lib.pagesizes import A4
    from reportlab.pdfgen import canvas as pdfcanvas
    from reportlab.lib.utils import ImageReader
    from reportlab.pdfbase import pdfmetrics
    from PIL import Image
    REPORTLAB_OK = True
except Exception:
    REPORTLAB_OK = False


# ---------------------------------------------------------------------------
# Constantes
# ---------------------------------------------------------------------------
CATEGORIAS = [
    ("XaviWork", "💼"),
    ("Chat", "💬"),
    ("Documentos", "📄"),
    ("Web", "🌐"),
    ("Multimedia", "🎬"),
    ("Estudio", "📚"),
    ("Viajes", "✈️"),
    ("Música", "🎵"),
    ("Gestor de Vecinos", "🏘️"),
]

ESTILOS_MUSICALES = [
    "Balada", "Rock", "Pop", "Reggaetón",
    "Cumbia", "Jazz", "Electrónica", "Clásica",
]

VOCES_ES = {
    "Álvaro (ES, hombre)": "es-ES-AlvaroNeural",
    "Elvira (ES, mujer)": "es-ES-ElviraNeural",
    "Jorge (MX, hombre)": "es-MX-JorgeNeural",
    "Dalia (MX, mujer)": "es-MX-DaliaNeural",
    "Tomás (AR, hombre)": "es-AR-TomasNeural",
    "Elena (AR, mujer)": "es-AR-ElenaNeural",
    "Gonzalo (CO, hombre)": "es-CO-GonzaloNeural",
    "Salomé (CO, mujer)": "es-CO-SalomeNeural",
}

ORDEN_IA_DEFAULT = [
    "gemini", "groq", "siliconflow", "mistral", "openrouter",
    "deepseek", "together", "cerebras", "ollama",
]


# ---------------------------------------------------------------------------
# Helpers generales
# ---------------------------------------------------------------------------
def get_api_key(name: str) -> str:
    """Lee de st.secrets primero y luego de variables de entorno."""
    try:
        if name in st.secrets:
            return str(st.secrets[name])
    except Exception:
        pass
    return os.environ.get(name, "")


def init_session():
    defaults = {
        "pagina": "menu",
        "chat_historial": [],
        "provider_preferido": "auto",
    }
    for k, v in defaults.items():
        if k not in st.session_state:
            st.session_state[k] = v


def limpiar_texto_para_voz(texto: str) -> str:
    """
    Limpia texto para TTS: quita asteriscos, corchetes, paréntesis,
    y palabras tipo 'Coro:', 'Verso 1:', 'Título:', etc.
    """
    if not texto:
        return ""
    t = texto

    # Quitar etiquetas tipo [Coro], [Verso 1], (Puente) conservando nada de su contenido
    t = re.sub(r"\[[^\]]*\]", "", t)
    t = re.sub(r"\([^\)]*\)", "", t)

    # Quitar etiquetas en línea tipo "Coro:", "Verso 1:", "Título:", "Estribillo:", etc.
    etiquetas = (
        r"coro|estribillo|verso|puente|pre-?coro|intro|outro|final|"
        r"t[íi]tulo|title|chorus|verse|bridge|pre-?chorus|intro|outro|"
        r"narrador|voz\s*1|voz\s*2|hombre|mujer"
    )
    t = re.sub(rf"(?im)^\s*(?:{etiquetas})\s*\d*\s*:?\s*$", "", t)
    t = re.sub(rf"(?i)\b(?:{etiquetas})\s*\d*\s*:\s*", "", t)

    # Quitar símbolos markdown y de formato
    t = re.sub(r"[*_#`~>|]", "", t)

    # Normalizar espacios
    t = re.sub(r"[ \t]+", " ", t)
    t = re.sub(r"\n{3,}", "\n\n", t)
    return t.strip()


# ---------------------------------------------------------------------------
# IA: proveedores y fallback
# ---------------------------------------------------------------------------
OPENAI_COMPATIBLE = {
    "groq":        {"url": "https://api.groq.com/openai/v1/chat/completions",
                    "model": "llama-3.1-8b-instant",          "env": "GROQ_API_KEY"},
    "siliconflow": {"url": "https://api.siliconflow.cn/v1/chat/completions",
                    "model": "Qwen/Qwen2.5-7B-Instruct",      "env": "SILICONFLOW_API_KEY"},
    "mistral":     {"url": "https://api.mistral.ai/v1/chat/completions",
                    "model": "mistral-small-latest",          "env": "MISTRAL_API_KEY"},
    "openrouter":  {"url": "https://openrouter.ai/api/v1/chat/completions",
                    "model": "meta-llama/llama-3.1-8b-instruct:free",
                    "env": "OPENROUTER_API_KEY"},
    "deepseek":    {"url": "https://api.deepseek.com/v1/chat/completions",
                    "model": "deepseek-chat",                 "env": "DEEPSEEK_API_KEY"},
    "together":    {"url": "https://api.together.xyz/v1/chat/completions",
                    "model": "meta-llama/Llama-3.3-70B-Instruct-Turbo",
                    "env": "TOGETHER_API_KEY"},
    "cerebras":    {"url": "https://api.cerebras.ai/v1/chat/completions",
                    "model": "llama3.1-8b",                   "env": "CEREBRAS_API_KEY"},
}

TIEMPO_MAX = 60


def _call_gemini(prompt: str, system: str, max_tokens: int) -> str:
    key = get_api_key("GEMINI_API_KEY")
    if not key:
        raise RuntimeError("GEMINI_API_KEY no configurada")
    url = f"https://generativelanguage.googleapis.com/v1beta/models/gemini-1.5-flash:generateContent?key={key}"
    body = {
        "contents": [{"role": "user", "parts": [{"text": (system + "\n\n" + prompt).strip()}]}],
        "generationConfig": {"maxOutputTokens": max_tokens, "temperature": 0.8},
    }
    r = requests.post(url, json=body, timeout=TIEMPO_MAX)
    r.raise_for_status()
    data = r.json()
    return data["candidates"][0]["content"]["parts"][0]["text"]


def _call_openai_compatible(provider: str, prompt: str, system: str, max_tokens: int) -> str:
    cfg = OPENAI_COMPATIBLE[provider]
    key = get_api_key(cfg["env"])
    if not key:
        raise RuntimeError(f"{cfg['env']} no configurada")
    headers = {
        "Authorization": f"Bearer {key}",
        "Content-Type": "application/json",
    }
    if provider == "openrouter":
        headers["HTTP-Referer"] = "https://xavilu.network.pro"
        headers["X-Title"] = "Xavilu Network Pro"
    messages = []
    if system:
        messages.append({"role": "system", "content": system})
    messages.append({"role": "user", "content": prompt})
    body = {"model": cfg["model"], "messages": messages, "max_tokens": max_tokens, "temperature": 0.8}
    r = requests.post(cfg["url"], headers=headers, json=body, timeout=TIEMPO_MAX)
    r.raise_for_status()
    data = r.json()
    return data["choices"][0]["message"]["content"]


def _call_ollama(prompt: str, system: str, max_tokens: int) -> str:
    url = get_api_key("OLLAMA_URL") or "http://localhost:11434"
    model = get_api_key("OLLAMA_MODEL") or "llama3.2"
    full = (system + "\n\n" + prompt).strip() if system else prompt
    body = {"model": model, "prompt": full, "stream": False,
            "options": {"num_predict": max_tokens}}
    r = requests.post(f"{url}/api/generate", json=body, timeout=TIEMPO_MAX)
    r.raise_for_status()
    return r.json().get("response", "")


def _proveedores_disponibles() -> List[str]:
    disp = []
    if get_api_key("GEMINI_API_KEY"):
        disp.append("gemini")
    for p, cfg in OPENAI_COMPATIBLE.items():
        if get_api_key(cfg["env"]):
            disp.append(p)
    # Ollama se asume disponible localmente
    disp.append("ollama")
    return disp


def call_ai(prompt: str, system: str = "", max_tokens: int = 2048) -> str:
    """
    Llama a la IA con fallback automático entre proveedores.
    Devuelve el texto generado o lanza RuntimeError si ninguno responde.
    """
    orden = ORDEN_IA_DEFAULT.copy()
    pref = st.session_state.get("provider_preferido", "auto")
    if pref != "auto" and pref in orden:
        orden.remove(pref)
        orden.insert(0, pref)

    errores = []
    for p in orden:
        try:
            if p == "gemini":
                return _call_gemini(prompt, system, max_tokens)
            elif p == "ollama":
                return _call_ollama(prompt, system, max_tokens)
            elif p in OPENAI_COMPATIBLE:
                return _call_openai_compatible(p, prompt, system, max_tokens)
        except Exception as e:
            errores.append(f"{p}: {e}")
            continue
    raise RuntimeError("Ningún proveedor de IA respondió.\n" + "\n".join(errores))


def barra_proveedores():
    """Muestra un selector de proveedor preferido (o auto)."""
    disp = _proveedores_disponibles()
    opciones = ["auto"] + disp
    idx = opciones.index(st.session_state.get("provider_preferido", "auto")) \
        if st.session_state.get("provider_preferido", "auto") in opciones else 0
    sel = st.selectbox(
        "Proveedor de IA", opciones, index=idx,
        help="'auto' prueba en orden hasta que uno responda. Los no configurados se saltan.",
        key="provider_selector_global",
    )
    st.session_state["provider_preferido"] = sel


# ---------------------------------------------------------------------------
# Música: base musical + TTS + mezcla
# ---------------------------------------------------------------------------
BPM_POR_ESTILO = {
    "Balada": 65, "Rock": 125, "Pop": 110, "Reggaetón": 95,
    "Cumbia": 92, "Jazz": 85, "Electrónica": 128, "Clásica": 72,
}


def _acorde_seg(freqs, dur_ms: int, amp_gain: float = -14) -> AudioSegment:
    seg = AudioSegment.silent(duration=dur_ms, frame_rate=44100)
    for f in freqs:
        tone = Sine(f).to_audio_segment(duration=dur_ms)
        tone = tone.fade_in(60).fade_out(120) + amp_gain
        seg = seg.overlay(tone)
    return seg


def generar_base_musical(estilo: str, duracion_ms: int) -> AudioSegment:
    """Genera una base musical sintética acorde al estilo."""
    bpm = BPM_POR_ESTILO.get(estilo, 100)
    beat_ms = int(60000 / bpm)
    chord_len = beat_ms * 4

    progresiones = {
        "Balada":     [[261.63, 329.63, 392.00], [392.00, 493.88, 587.33], [220.00, 261.63, 329.63], [174.61, 220.00, 261.63]],
        "Rock":       [[196.00, 246.94, 293.66], [261.63, 329.63, 392.00], [220.00, 261.63, 329.63], [174.61, 220.00, 261.63]],
        "Pop":        [[261.63, 329.63, 392.00], [349.23, 440.00, 523.25], [392.00, 493.88, 587.33], [261.63, 329.63, 392.00]],
        "Reggaetón":  [[220.00, 261.63, 329.63], [174.61, 220.00, 261.63], [261.63, 329.63, 392.00], [196.00, 246.94, 293.66]],
        "Cumbia":     [[261.63, 329.63, 392.00], [349.23, 440.00, 523.25], [261.63, 329.63, 392.00], [174.61, 220.00, 261.63]],
        "Jazz":       [[261.63, 311.13, 392.00], [293.66, 349.23, 440.00], [246.94, 293.66, 369.99], [220.00, 261.63, 329.63]],
        "Electrónica":[[261.63, 392.00, 523.25], [196.00, 293.66, 392.00], [329.63, 493.88, 659.25], [220.00, 329.63, 440.00]],
        "Clásica":    [[261.63, 329.63, 392.00], [392.00, 493.88, 587.33], [220.00, 261.63, 329.63], [174.61, 220.00, 261.63]],
    }
    prog = progresiones.get(estilo, progresiones["Balada"])

    audio = AudioSegment.silent(duration=duracion_ms, frame_rate=44100)
    t = 0
    i = 0
    while t < duracion_ms:
        chord = prog[i % len(prog)]
        seg = _acorde_seg(chord, chord_len, amp_gain=-16)
        audio = audio.overlay(seg, position=t)
        t += chord_len
        i += 1

    # Bombo sintético en negras 1 y 3 (para estilos con ritmo)
    if estilo in ("Rock", "Pop", "Reggaetón", "Cumbia", "Electrónica"):
        kick = Sine(55).to_audio_segment(duration=130).fade_out(110) - 6
        hihat = WhiteNoise().to_audio_segment(duration=40) - 26
        pos = 0
        beat_idx = 0
        while pos < duracion_ms:
            if beat_idx % 4 in (0, 2):
                audio = audio.overlay(kick, position=pos)
            if beat_idx % 2 == 1 and estilo in ("Rock", "Electrónica", "Pop"):
                audio = audio.overlay(hihat, position=pos)
            pos += beat_ms
            beat_idx += 1

    return audio[:duracion_ms]


async def _tts_async(texto: str, voz: str, out_path: str, rate: str = "+0%"):
    communicate = edge_tts.Communicate(texto, voz, rate=rate)
    await communicate.save(out_path)


def tts_a_archivo(texto: str, voz: str, out_path: str, rate: str = "+0%"):
    asyncio.run(_tts_async(texto, voz, out_path, rate))


def generar_cancion(letra: str, estilo: str, voz_1: str, voz_2: Optional[str],
                    modo_duo: bool, rate: str = "+0%", progreso=None):
    """
    Devuelve (AudioSegment_final, lista_lineas).
    - Limpia el texto con limpiar_texto_para_voz()
    - Genera audio TTS por línea
    - Alterna voces si modo_dúo
    - Mezcla con base musical del estilo
    """
    if not PYDUB_OK:
        raise RuntimeError("pydub/imageio-ffmpeg no disponibles")
    if not EDGE_TTS_OK:
        raise RuntimeError("edge-tts no disponible")

    texto_limpio = limpiar_texto_para_voz(letra)
    lineas = [ln.strip() for ln in texto_limpio.split("\n") if ln.strip()]
    if not lineas:
        raise RuntimeError("La letra está vacía después de limpiarla.")

    tmpdir = tempfile.mkdtemp(prefix="xavilu_")
    clips = []
    total = len(lineas)
    for idx, linea in enumerate(lineas):
        voz = voz_1 if (not modo_duo or idx % 2 == 0) else (voz_2 or voz_1)
        out = os.path.join(tmpdir, f"l{idx:03d}.mp3")
        try:
            tts_a_archivo(linea, voz, out, rate=rate)
            clips.append(AudioSegment.from_file(out, format="mp3"))
        except Exception as e:
            raise RuntimeError(f"Error TTS en línea {idx+1}: {e}")
        if progreso:
            progreso.progress(min(1.0, (idx + 1) / (total + 1)),
                              text=f"Sintetizando línea {idx+1}/{total}...")

    # Voz completa con pausas entre líneas
    pausa = AudioSegment.silent(duration=350, frame_rate=44100)
    voz_total = AudioSegment.silent(duration=0, frame_rate=44100)
    for c in clips:
        voz_total += c + pausa

    if progreso:
        progreso.progress(1.0, text="Generando base musical...")

    # Base musical con la duración de la voz + pequeño margen
    dur_ms = len(voz_total) + 1500
    base = generar_base_musical(estilo, dur_ms)

    # Mezcla: base más baja, voz principal
    base = base - 6
    final = base.overlay(voz_total, position=800)

    if progreso:
        progreso.progress(1.0, text="Mezcla lista ✅")

    return final, lineas


def audio_segment_a_bytes(seg: AudioSegment, fmt: str = "mp3") -> bytes:
    buf = io.BytesIO()
    seg.export(buf, format=fmt)
    return buf.getvalue()


# ---------------------------------------------------------------------------
# PDF
# ---------------------------------------------------------------------------
def crear_pdf(titulo: str, contenido: str, imagen_bytes: Optional[bytes] = None) -> bytes:
    if not REPORTLAB_OK:
        raise RuntimeError("reportlab/Pillow no disponibles")
    buf = io.BytesIO()
    c = pdfcanvas.Canvas(buf, pagesize=A4)
    W, H = A4

    if imagen_bytes:
        try:
            img = Image.open(io.BytesIO(imagen_bytes)).convert("RGB")
            img_buf = io.BytesIO()
            img.save(img_buf, format="JPEG", quality=88)
            img_buf.seek(0)
            c.drawImage(ImageReader(img_buf), 0, 0, width=W, height=H,
                        preserveAspectRatio=False, mask=None)
            c.showPage()
        except Exception:
            pass

    # Página de contenido con ajuste de líneas
    c.setFont("Helvetica-Bold", 16)
    c.drawString(50, H - 60, titulo[:90])
    c.setFont("Helvetica", 11)
    text_obj = c.beginText(50, H - 90)
    max_chars = 95
    y_min = 60
    for raw in contenido.split("\n"):
        if raw == "":
            text_obj.textLine("")
            continue
        words = raw.split(" ")
        line = ""
        for w in words:
            if len(line) + len(w) + 1 > max_chars:
                text_obj.textLine(line)
                line = w
                if text_obj.getY() < y_min:
                    c.drawText(text_obj)
                    c.showPage()
                    c.setFont("Helvetica", 11)
                    text_obj = c.beginText(50, H - 60)
            else:
                line = (line + " " + w).strip() if line else w
        if line:
            text_obj.textLine(line)
        if text_obj.getY() < y_min:
            c.drawText(text_obj)
            c.showPage()
            c.setFont("Helvetica", 11)
            text_obj = c.beginText(50, H - 60)
    c.drawText(text_obj)
    c.save()
    buf.seek(0)
    return buf.getvalue()


def generar_acta_pdf(titulo: str, fecha: str, asistentes: str,
                     orden_dia: str, acuerdos: str, firmante: str) -> bytes:
    contenido = (
        f"Fecha: {fecha}\n"
        f"Asistentes: {asistentes}\n\n"
        "ORDEN DEL DÍA\n"
        f"{orden_dia}\n\n"
        "ACUERDOS ADOPTADOS\n"
        f"{acuerdos}\n\n"
        f"Firmado: {firmante}\n"
    )
    return crear_pdf(titulo or "Acta de la Comunidad", contenido)


# ---------------------------------------------------------------------------
# UI: cabecera y menú
# ---------------------------------------------------------------------------
def render_header():
    st.markdown(
        "<h1 style='text-align:center;margin-bottom:0'>🎸 Xavilu Network Pro</h1>"
        "<p style='text-align:center;color:#888;margin-top:0'>Tu suite multi-IA todo en uno</p>",
        unsafe_allow_html=True,
    )
    if st.session_state.pagina != "menu":
        if st.button("⬅️ Volver al menú", key="btn_back"):
            st.session_state.pagina = "menu"
            st.rerun()
    st.divider()


def render_menu():
    st.subheader("Menú principal")
    idx = 0
    for fila_inicio in range(0, len(CATEGORIAS), 4):
        fila = CATEGORIAS[fila_inicio:fila_inicio + 4]
        cols = st.columns(4)
        for j, (nombre, emoji) in enumerate(fila):
            if cols[j].button(f"{emoji}\n\n**{nombre}**",
                              use_container_width=True,
                              key=f"menu_btn_{idx}"):
                st.session_state.pagina = nombre
                st.rerun()
            idx += 1
    st.divider()
    barra_proveedores()
    disp = _proveedores_disponibles()
    st.caption(f"Proveedores con API key detectada: {', '.join(disp) if disp else 'ninguno'}")


# ---------------------------------------------------------------------------
# Sección: XaviWork
# ---------------------------------------------------------------------------
def render_xaviwork():
    st.subheader("💼 XaviWork — Productividad laboral")
    tab1, tab2, tab3 = st.tabs(["Redactar email", "Resumir texto", "Planificar tareas"])
    with tab1:
        ctx = st.text_area("Contexto / destinatario", height=80,
                           placeholder="Ej: jefe, para pedir vacaciones la semana del 20...")
        tono = st.selectbox("Tono", ["Formal", "Cercano", "Directo", "Diplomático"])
        if st.button("Redactar email", key="xw_email"):
            if not ctx.strip():
                st.warning("Escribe un contexto.")
            else:
                with st.spinner("Redactando..."):
                    try:
                        out = call_ai(
                            f"Redacta un email en español, tono {tono}. Contexto: {ctx}",
                            system="Eres un asistente experto en comunicación profesional. Devuelve solo el email listo para enviar, con asunto.",
                        )
                        st.text_area("Resultado", out, height=300)
                        st.download_button("Descargar .txt", out, file_name="email.txt")
                    except Exception as e:
                        st.error(f"Error IA: {e}")

    with tab2:
        texto = st.text_area("Texto a resumir", height=200)
        n = st.slider("Puntos clave", 3, 10, 5)
        if st.button("Resumir", key="xw_res"):
            if not texto.strip():
                st.warning("Pega un texto.")
            else:
                with st.spinner("Resumiendo..."):
                    try:
                        out = call_ai(
                            f"Resume el siguiente texto en {n} puntos clave, en español:\n\n{texto}",
                            system="Eres un analista que sintetiza información con claridad.",
                        )
                        st.markdown(out)
                    except Exception as e:
                        st.error(f"Error IA: {e}")

    with tab3:
        objetivos = st.text_area("Objetivos / notas del día", height=150)
        if st.button("Planificar", key="xw_plan"):
            if not objetivos.strip():
                st.warning("Escribe objetivos.")
            else:
                with st.spinner("Planificando..."):
                    try:
                        out = call_ai(
                            f"Convierte esto en un plan de tareas priorizado (matriz Eisenhower):\n\n{objetivos}",
                            system="Eres un coach de productividad.",
                        )
                        st.markdown(out)
                    except Exception as e:
                        st.error(f"Error IA: {e}")


# ---------------------------------------------------------------------------
# Sección: Chat
# ---------------------------------------------------------------------------
def render_chat():
    st.subheader("💬 Chat multi-IA")
    if st.button("🧹 Limpiar historial", key="chat_clear"):
        st.session_state.chat_historial = []
        st.rerun()

    for msg in st.session_state.chat_historial:
        with st.chat_message(msg["role"]):
            st.markdown(msg["content"])

    prompt = st.chat_input("Escribe tu mensaje...")
    if prompt:
        st.session_state.chat_historial.append({"role": "user", "content": prompt})
        with st.chat_message("user"):
            st.markdown(prompt)
        with st.chat_message("assistant"):
            with st.spinner("Pensando..."):
                try:
                    contexto = "\n".join(
                        f"{m['role']}: {m['content']}"
                        for m in st.session_state.chat_historial[-10:]
                    )
                    out = call_ai(contexto, system="Eres un asistente útil que responde en español.")
                    st.markdown(out)
                    st.session_state.chat_historial.append({"role": "assistant", "content": out})
                except Exception as e:
                    st.error(f"Error IA: {e}")


# ---------------------------------------------------------------------------
# Sección: Documentos
# ---------------------------------------------------------------------------
def render_documentos():
    st.subheader("📄 Documentos — Generar PDF")
    tab1, tab2 = st.tabs(["PDF libre", "Acta de comunidad"])

    with tab1:
        titulo = st.text_input("Título del documento", value="Mi documento")
        contenido = st.text_area("Contenido", height=250,
                                 placeholder="Escribe o pega aquí el contenido...")
        portada = st.file_uploader("Imagen de portada (opcional)",
                                   type=["png", "jpg", "jpeg"], key="doc_cover")
        col1, col2 = st.columns(2)
        with col1:
            if st.button("Generar con IA a partir de idea", key="doc_ai_gen"):
                idea = st.session_state.get("doc_idea", "")
                if not idea.strip():
                    st.info("Escribe una idea en el contenido y pulsa de nuevo.")
                else:
                    with st.spinner("Generando contenido..."):
                        try:
                            out = call_ai(f"Redacta un documento titulado '{titulo}' sobre: {idea}",
                                          system="Eres un redactor profesional en español.")
                            st.session_state["doc_generado"] = out
                        except Exception as e:
                            st.error(f"Error IA: {e}")
        with col2:
            if st.button("📄 Generar PDF", key="doc_pdf", type="primary"):
                texto_final = contenido or st.session_state.get("doc_generado", "")
                if not texto_final.strip():
                    st.warning("Añade contenido.")
                else:
                    try:
                        img_bytes = portada.read() if portada else None
                        pdf = crear_pdf(titulo, texto_final, img_bytes)
                        st.success("PDF listo.")
                        st.download_button("⬇️ Descargar PDF", pdf,
                                           file_name=f"{titulo[:40] or 'documento'}.pdf",
                                           mime="application/pdf")
                    except Exception as e:
                        st.error(f"Error al crear PDF: {e}")

        # Guardamos el contenido del textarea como idea para el botón IA
        st.session_state["doc_idea"] = contenido

    with tab2:
        st.caption("Genera un acta formal con encabezado, asistentes, acuerdos y firma.")
        t_acta = st.text_input("Título del acta", value="Acta de Junta de Vecinos")
        fecha = st.date_input("Fecha", datetime.now())
        asistentes = st.text_area("Asistentes (uno por línea)", height=80)
        orden = st.text_area("Orden del día", height=100)
        acuerdos = st.text_area("Acuerdos adoptados", height=120)
        firmante = st.text_input("Firmante (Presidente/a)", value="")
        if st.button("📝 Generar Acta PDF", key="acta_pdf", type="primary"):
            if not acuerdos.strip():
                st.warning("Añade acuerdos.")
            else:
                try:
                    pdf = generar_acta_pdf(
                        t_acta, fecha.strftime("%d/%m/%Y"), asistentes,
                        orden, acuerdos, firmante or "La Presidencia",
                    )
                    st.success("Acta generada.")
                    st.download_button("⬇️ Descargar Acta", pdf,
                                       file_name="acta.pdf", mime="application/pdf")
                except Exception as e:
                    st.error(f"Error: {e}")


# ---------------------------------------------------------------------------
# Sección: Web
# ---------------------------------------------------------------------------
def render_web():
    st.subheader("🌐 Web — Resumen y análisis de URLs")
    url = st.text_input("URL a analizar", placeholder="https://...")
    tarea = st.selectbox("¿Qué hago con ella?",
                         ["Resumir", "Extraer puntos clave", "Analizar tono", "Traducir al español"])
    if st.button("Procesar", key="web_go"):
        if not url.strip().startswith("http"):
            st.warning("Introduce una URL válida (http/https).")
            return
        with st.spinner("Descargando..."):
            try:
                r = requests.get(url, timeout=20, headers={"User-Agent": "Mozilla/5.0 XaviluBot"})
                html = r.text
                # Limpieza rápida
                texto = re.sub(r"<script.*?</script>", " ", html, flags=re.S | re.I)
                texto = re.sub(r"<style.*?</style>", " ", texto, flags=re.S | re.I)
                texto = re.sub(r"<[^>]+>", " ", texto)
                texto = re.sub(r"\s+", " ", texto)[:8000]
            except Exception as e:
                st.error(f"Error descargando: {e}")
                return
        with st.spinner("Analizando..."):
            try:
                out = call_ai(
                    f"Tarea: {tarea}.\n\nContenido de la página:\n{texto}",
                    system="Eres un analista web que responde en español.",
                )
                st.markdown(out)
            except Exception as e:
                st.error(f"Error IA: {e}")


# ---------------------------------------------------------------------------
# Sección: Multimedia
# ---------------------------------------------------------------------------
def render_multimedia():
    st.subheader("🎬 Multimedia — Herramientas de imagen")
    if not REPORTLAB_OK:
        st.warning("Pillow no está instalado. Instala `pillow` para usar esta sección.")
        return
    img_file = st.file_uploader("Sube una imagen", type=["png", "jpg", "jpeg", "webp"])
    if not img_file:
        return
    img = Image.open(img_file)
    st.image(img, caption="Original", use_container_width=True)
    op = st.selectbox("Operación", ["Redimensionar", "Escala de grises", "Rotar 90º", "Blanco y negro (binarizar)"])
    if op == "Redimensionar":
        w = st.number_input("Ancho (px)", 50, 4000, img.width)
        h = st.number_input("Alto (px)", 50, 4000, img.height)
        out_img = img.resize((int(w), int(h)))
    elif op == "Escala de grises":
        out_img = img.convert("L")
    elif op == "Rotar 90º":
        out_img = img.rotate(-90, expand=True)
    else:
        out_img = img.convert("L").point(lambda x: 0 if x < 128 else 255, '1').convert("RGB")

    buf = io.BytesIO()
    fmt = "PNG" if img.format == "PNG" else "JPEG"
    out_img.save(buf, format=fmt)
    st.image(out_img, caption="Resultado", use_container_width=True)
    st.download_button("⬇️ Descargar imagen",
                       buf.getvalue(),
                       file_name=f"resultado.{fmt.lower()}",
                       mime=f"image/{fmt.lower()}")


# ---------------------------------------------------------------------------
# Sección: Estudio
# ---------------------------------------------------------------------------
def render_estudio():
    st.subheader("📚 Estudio — Ayudas al aprendizaje")
    tema = st.text_input("Tema / apuntes", placeholder="Ej: Revolución Francesa, fotosíntesis...")
    modo = st.selectbox("Herramienta", [
        "Resumen didáctico", "Preguntas tipo test (10)", "Flashcards (10)",
        "Explicación para principiantes", "Mapa mental (texto)"
    ])
    if st.button("Generar", key="est_go"):
        if not tema.strip():
            st.warning("Escribe un tema.")
            return
        with st.spinner("Estudiando..."):
            try:
                prompt = f"Tema: {tema}\nHerramienta: {modo}. Responde en español con formato claro."
                out = call_ai(prompt, system="Eres un profesor experto que enseña paso a paso.")
                st.markdown(out)
                st.download_button("⬇️ Descargar .txt", out, file_name="estudio.txt")
            except Exception as e:
                st.error(f"Error IA: {e}")


# ---------------------------------------------------------------------------
# Sección: Viajes
# ---------------------------------------------------------------------------
def render_viajes():
    st.subheader("✈️ Viajes — Planificador")
    c1, c2 = st.columns(2)
    with c1:
        destino = st.text_input("Destino", placeholder="Ej: Lisboa")
        dias = st.number_input("Días", 1, 30, 3)
    with c2:
        presupuesto = st.selectbox("Presupuesto", ["Bajo", "Medio", "Alto"])
        interes = st.text_input("Intereses", placeholder="Ej: gastronomía, arte, naturaleza")
    if st.button("🧭 Planificar viaje", key="via_go"):
        if not destino.strip():
            st.warning("Indica destino.")
            return
        with st.spinner("Organizando itinerario..."):
            try:
                out = call_ai(
                    f"Planifica un viaje a {destino} de {dias} días, presupuesto {presupuesto}, "
                    f"intereses: {interes or 'generales'}. Devuelve itinerario día a día con horarios, "
                    "consejos y presupuesto estimado. En español.",
                    system="Eres un agente de viajes experto.",
                )
                st.markdown(out)
                st.download_button("⬇️ Descargar itinerario", out, file_name="viaje.txt")
            except Exception as e:
                st.error(f"Error IA: {e}")


# ---------------------------------------------------------------------------
# Sección: Música (generador de letras + reproductor con Modo Dúo)
# ---------------------------------------------------------------------------
def render_musica():
    st.subheader("🎵 Música — Letras + Reproductor (Voz + Base)")
    tab_gen, tab_player = st.tabs(["Generar letra", "Reproductor / Mezcla"])

    # --- Generar letra ---
    with tab_gen:
        c1, c2, c3 = st.columns(3)
        with c1:
            tema = st.text_input("Tema de la canción", placeholder="Ej: un verano inolvidable")
        with c2:
            genero = st.selectbox("Género lírico", ESTILOS_MUSICALES)
        with c3:
            idioma = st.selectbox("Idioma", ["Español", "Inglés", "Spanglish"])
        estructura = st.multiselect(
            "Estructura",
            ["Intro", "Verso 1", "Pre-Coro", "Coro", "Verso 2", "Puente", "Outro"],
            default=["Verso 1", "Coro", "Verso 2", "Coro"],
        )
        if st.button("✍️ Generar letra", key="mus_letra"):
            if not tema.strip():
                st.warning("Indica un tema.")
            else:
                with st.spinner("Componiendo..."):
                    try:
                        prompt = (
                            f"Escribe la letra completa de una canción en {idioma} sobre '{tema}', "
                            f"estilo {genero}. Estructura: {', '.join(estructura)}. "
                            "Separa secciones con etiquetas tipo [Coro], [Verso 1]. "
                            "Versos de 4 líneas, rima natural, sin exagerar metáforas."
                        )
                        out = call_ai(prompt, system="Eres un compositor profesional.")
                        st.session_state["letra_generada"] = out
                    except Exception as e:
                        st.error(f"Error IA: {e}")
        if st.session_state.get("letra_generada"):
            st.text_area("Letra generada (editable)", st.session_state["letra_generada"],
                         height=320, key="letra_edit")
            # mantenemos edición sincronizada
            st.session_state["letra_generada"] = st.session_state["letra_edit"]

    # --- Reproductor ---
    with tab_player:
        st.caption("Mezcla la letra (voz) con una base musical generada. Soporta Modo Dúo.")
        letra = st.text_area("Letra a cantar",
                             value=st.session_state.get("letra_generada", ""),
                             height=260, key="letra_player")
        c1, c2 = st.columns(2)
        with c1:
            estilo_mus = st.selectbox("Estilo de base musical", ESTILOS_MUSICALES, key="est_mus")
            voz_a_nombre = st.selectbox("Voz principal", list(VOCES_ES.keys()), index=0)
        with c2:
            modo_duo = st.checkbox("Modo Dúo (2 voces alternando por línea)", value=False)
            voz_b_nombre = st.selectbox("Segunda voz (solo dúo)", list(VOCES_ES.keys()), index=1)

        rate = st.slider("Velocidad voz", -30, 30, 0, step=5,
                         help="Negativo = más lento; positivo = más rápido")
        rate_str = f"{rate:+d}%"

        if not PYDUB_OK:
            st.error("Faltan dependencias: `pip install pydub imageio-ffmpeg`")
        if not EDGE_TTS_OK:
            st.error("Falta `pip install edge-tts`")

        if st.button("🎧 Generar canción", key="mus_play", type="primary"):
            if not letra.strip():
                st.warning("La letra está vacía.")
            elif not (PYDUB_OK and EDGE_TTS_OK):
                st.error("Instala las dependencias que faltan.")
            else:
                progreso = st.progress(0.0, text="Preparando...")
                try:
                    final, lineas = generar_cancion(
                        letra=letra,
                        estilo=estilo_mus,
                        voz_1=VOCES_ES[voz_a_nombre],
                        voz_2=VOCES_ES[voz_b_nombre] if modo_duo else None,
                        modo_duo=modo_duo,
                        rate=rate_str,
                        progreso=progreso,
                    )
                    st.success(f"¡Listo! {len(lineas)} líneas mezcladas · Estilo {estilo_mus}"
                               + (" · Modo Dúo 🎭" if modo_duo else ""))
                    audio_bytes = audio_segment_a_bytes(final, fmt="mp3")
                    st.audio(audio_bytes, format="audio/mp3")
                    st.download_button("⬇️ Descargar MP3", audio_bytes,
                                       file_name="xavilu_cancion.mp3", mime="audio/mpeg")
                    with st.expander("Ver texto limpio enviado al TTS"):
                        st.text(limpiar_texto_para_voz(letra))
                except Exception as e:
                    st.error(f"Error generando canción: {e}")


# ---------------------------------------------------------------------------
# Sección: Gestor de Vecinos (PRL)
# ---------------------------------------------------------------------------
def render_gestor_vecinos():
    st.subheader("🏘️ Gestor de Vecinos (PRL)")
    tab1, tab2, tab3 = st.tabs(["🧮 Calculadora de mayorías", "⚖️ Alex IA Legal", "📝 Generar Acta"])

    # --- Calculadora de mayorías ---
    with tab1:
        st.caption("Calcula si una votación alcanza el umbral requerido (LPH simplificada).")
        total = st.number_input("Total de vecinos (coeficientes de participación)", 1, 500, 10)
        c1, c2, c3 = st.columns(3)
        with c1:
            a_favor = st.number_input("A favor", 0, int(total), 0)
        with c2:
            en_contra = st.number_input("En contra", 0, int(total), 0)
        with c3:
            abstencion = st.number_input("Abstención", 0, int(total), 0)
        umbral = st.selectbox("Mayoría exigida", [
            "Simple (más a favor que en contra)",
            "Absoluta (más de la mitad del total)",
            "3/5 (60%)",
            "2/3 (66,6%)",
            "Unanimidad",
        ])
        if st.button("Calcular", key="prl_calc"):
            emitidos = a_favor + en_contra + abstencion
            if emitidos > total:
                st.error("Los votos emitidos superan el total de vecinos.")
            else:
                cumple = False
                detalle = ""
                if umbral.startswith("Simple"):
                    cumple = a_favor > en_contra
                    detalle = f"A favor ({a_favor}) vs en contra ({en_contra})"
                elif umbral.startswith("Absoluta"):
                    umbral_val = total / 2
                    cumple = a_favor > umbral_val
                    detalle = f"A favor {a_favor} · necesarios > {umbral_val:.1f} del total ({total})"
                elif umbral.startswith("3/5"):
                    umbral_val = total * 0.6
                    cumple = a_favor >= umbral_val
                    detalle = f"A favor {a_favor} · necesarios ≥ {umbral_val:.1f}"
                elif umbral.startswith("2/3"):
                    umbral_val = total * (2 / 3)
                    cumple = a_favor >= umbral_val
                    detalle = f"A favor {a_favor} · necesarios ≥ {umbral_val:.1f}"
                else:
                    cumple = a_favor == total
                    detalle = f"A favor {a_favor} · necesarios = {total}"

                if cumple:
                    st.success(f"✅ Se alcanza la mayoría ({umbral}).\n\n{detalle}")
                else:
                    st.error(f"❌ No se alcanza la mayoría ({umbral}).\n\n{detalle}")
                st.info(f"Votos emitidos: {emitidos} / {total}")

    # --- Alex IA Legal ---
    with tab2:
        st.caption("Consulta dudas de la Ley de Propiedad Horizontal. Alex IA responde con orientación.")
        pregunta = st.text_area("Tu consulta legal",
                                placeholder="Ej: ¿Puedo instalar un aire acondicionado en la fachada?",
                                height=120)
        # --- Línea corregida (antes daba SyntaxError si faltaba cerrar paréntesis) ---
        vecino_afectado = st.text_input("¿A quién afecta? (Ej: A mi vivienda, a todo el bloque...):", value="A mi vivienda")
        if st.button("Consultar a Alex IA", key="alex_go"):
            if not pregunta.strip():
                st.warning("Escribe una pregunta.")
            else:
                with st.spinner("Consultando legislación..."):
                    try:
                        system = (
                            "Eres Alex, un asistente legal especializado en la Ley de Propiedad "
                            "Horizontal (LPH) española y en la prevención de riesgos en comunidades "
                            "de vecinos. Responde en español, con lenguaje claro, citando los "
                            "artículos relevantes cuando aplique. Añade siempre una nota final "
                            "aclarando que tu respuesta es orientativa y no sustituye a un abogado."
                        )
                        prompt = (
                            f"Consulta: {pregunta}\n"
                            f"Ámbito de afección: {vecino_afectado}\n"
                            "Proporciona: 1) respuesta directa, 2) fundamento legal, "
                            "3) qué mayoría se requiere si hay votación, 4) recomendaciones prácticas."
                        )
                        out = call_ai(prompt, system=system)
                        st.markdown(out)
                    except Exception as e:
                        st.error(f"Error IA: {e}")

    # --- Generar acta ---
    with tab3:
        st.caption("Genera un acta formal en PDF a partir de los datos de la reunión.")
        t = st.text_input("Título del acta", value="Acta de Junta Ordinaria")
        fecha = st.date_input("Fecha de la reunión", datetime.now())
        asistentes = st.text_area("Asistentes (uno por línea)", height=80)
        orden = st.text_area("Orden del día", height=100)
        acuerdos = st.text_area("Acuerdos adoptados", height=120)
        firmante = st.text_input("Presidente/a firmante", value="")
        if st.button("📝 Generar Acta PDF", key="acta_vecinos_pdf", type="primary"):
            if not acuerdos.strip():
                st.warning("Faltan acuerdos.")
            else:
                try:
                    pdf = generar_acta_pdf(t, fecha.strftime("%d/%m/%Y"),
                                           asistentes, orden, acuerdos,
                                           firmante or "La Presidencia")
                    st.success("Acta lista.")
                    st.download_button("⬇️ Descargar Acta PDF", pdf,
                                       file_name="acta_vecinos.pdf",
                                       mime="application/pdf")
                except Exception as e:
                    st.error(f"Error al crear PDF: {e}")


# ---------------------------------------------------------------------------
# Enrutador principal
# ---------------------------------------------------------------------------
def main():
    init_session()
    render_header()

    pagina = st.session_state.pagina
    if pagina == "menu":
        render_menu()
    elif pagina == "XaviWork":
        render_xaviwork()
    elif pagina == "Chat":
        render_chat()
    elif pagina == "Documentos":
        render_documentos()
    elif pagina == "Web":
        render_web()
    elif pagina == "Multimedia":
        render_multimedia()
    elif pagina == "Estudio":
        render_estudio()
    elif pagina == "Viajes":
        render_viajes()
    elif pagina == "Música":
        render_musica()
    elif pagina == "Gestor de Vecinos":
        render_gestor_vecinos()
    else:
        st.error(f"Página desconocida: {pagina}")
        st.session_state.pagina = "menu"
        st.rerun()


if __name__ == "__main__":
    main()
