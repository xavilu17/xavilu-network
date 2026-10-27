import streamlit as st
from groq import Groq

# --- CONFIGURACIÓN DE LA PÁGINA ---
st.set_page_config(page_title="CreatorFlow Demo", page_icon="✨", layout="wide")

# --- CONEXIÓN CON GROQ ---
try:
    client = Groq(api_key=st.secrets["GROQ_API_KEY"])
except Exception as e:
    st.error("⚠️ Falta la API Key de Groq. Configúrala en los Secrets de Streamlit.")
    st.stop()

# --- DATOS FALSOS (MOCK) ---
comentarios = [
    {"red": "Instagram", "usuario": "@maria_fit", "texto": "¡Increíble el video de hoy! ¿Qué rutina de ejercicios sigues?"},
    {"red": "X", "usuario": "@tech_lover", "texto": "Oye, ¿el nuevo producto del que hablaste sale a la venta pronto?"},
    {"red": "Instagram", "usuario": "@carlos_style", "texto": "Me encanta tu ropa, ¿dónde compras esas camisetas?"}
]

# --- INTERFAZ GRÁFICA ---
st.title("✨ CreatorFlow - Panel de Influencer")
st.markdown("Bienvenido a tu centro de control. Gestiona tus redes sin estrés.")
st.divider()

col1, col2 = st.columns([1, 1])

with col1:
    st.subheader("📥 Bandeja de Entrada")
    st.info("Estos son comentarios simulados de tus redes.")
    opciones = [f"[{c['red']}] {c['usuario']}: {c['texto']}" for c in comentarios]
    seleccion = st.radio("Selecciona un mensaje para responder:", opciones)

with col2:
    st.subheader("🤖 Asistente de IA")
    if seleccion:
        index = opciones.index(seleccion)
        comentario_elegido = comentarios[index]["texto"]
        red_social = comentarios[index]["red"]
        
        st.write(f"**Respondiendo a:** *{comentario_elegido}*")
        
        if st.button("Generar respuesta con IA", type="primary"):
            with st.spinner("Pensando la mejor respuesta..."):
                try:
                    chat_completion = client.chat.completions.create(
                        messages=[
                            {
                                "role": "system",
                                "content": f"Eres un asistente de redes sociales para un influencer. Responde de forma amigable, breve, usando emojis y adaptándote a la red social ({red_social})."
                            },
                            {
                                "role": "user",
                                "content": f"Genera una respuesta para este comentario: {comentario_elegido}"
                            }
                        ],
                        model="openai/gpt-oss-20b",  # <--- MODELO ACTUALIZADO Y COMPATIBLE
                    )
                    
                    respuesta = chat_completion.choices[0].message.content
                    st.success("¡Respuesta generada!")
                    st.text_area("Revisa y edita tu respuesta:", value=respuesta, height=150)
                    st.button("Publicar en " + red_social)
                    
                except Exception as e:
                    st.error(f"Error al conectar con la IA: {e}")
