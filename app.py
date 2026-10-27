import streamlit as st
import pandas as pd
import plotly.graph_objects as go
import io
import urllib.parse
import docx
from docx import Document
from docx.shared import Inches, Pt, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_TABLE_ALIGNMENT
from docx.oxml import OxmlElement, parse_xml
from docx.oxml.ns import qn, nsdecls

# -----------------------------------------------------------------------------
# CONFIGURACIÓN DE LA PÁGINA
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Programa de Entrenamiento Personalizado | Lic. Carlos Pacheco",
    page_icon="🫀",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Base de datos simulada de códigos de acceso de pacientes (Latinoamérica)
VALID_CODES = {
    "PACHECO-2026-VIP": {"name": "Carlos Rodríguez", "country": "🇻🇪 Venezuela", "condition": "Hipertensión Controlada"},
    "COLOMBIA-01": {"name": "Ana María Gómez", "country": "🇨🇴 Colombia", "condition": "Sedentarismo / Prevención"},
    "CHILE-05": {"name": "Jorge Valenzuela", "country": "🇨🇱 Chile", "condition": "Rigidez Arterial Moderada"},
    "MEXICO-99": {"name": "Sofía Hernández", "country": "🇲🇽 México", "condition": "Resistencia a la Insulina / HTA"}
}

# Estilos CSS personalizados para la plataforma médica
st.markdown("""
    <style>
    .main-header {
        background: linear-gradient(135deg, #0A2540 0%, #1A365D 100%);
        color: white;
        padding: 2rem;
        border-radius: 12px;
        text-align: center;
        margin-bottom: 2rem;
        box-shadow: 0 4px 6px rgba(0, 0, 0, 0.1);
    }
    .slogan-tag {
        color: #00E6C3;
        font-weight: 600;
        font-size: 1.1rem;
        margin-top: 0.5rem;
    }
    .callout-card {
        background-color: #F8F9FA;
        border-left: 5px solid #0A2540;
        padding: 1.2rem;
        border-radius: 6px;
        margin: 1rem 0;
    }
    .video-box {
        background-color: #F1F3F5;
        padding: 1.5rem;
        border-radius: 10px;
        border: 1px solid #E9ECEF;
        margin-bottom: 1rem;
    }
    .stButton>button {
        background-color: #0A2540;
        color: white;
        border-radius: 8px;
        font-weight: bold;
        transition: all 0.3s ease;
    }
    .stButton>button:hover {
        background-color: #00E6C3;
        color: #0A2540;
    }
    </style>
""", unsafe_allow_html=True)

# -----------------------------------------------------------------------------
# CONTROL DE ACCESO POR CÓDIGO ÚNICO (Seguridad de Pacientes)
# -----------------------------------------------------------------------------
if "authenticated" not in st.session_state:
    st.session_state.authenticated = False
if "patient_data" not in st.session_state:
    st.session_state.patient_data = {}

if not st.session_state.authenticated:
    st.markdown("""
        <div class="main-header">
            <h1 style="margin:0; font-size:2.2rem;">Portal de Acceso Exclusivo</h1>
            <h3 style="margin:0; font-weight:400; color:#E2E8F0; font-size:1.2rem;">Lic. Carlos Pacheco | Fisiología & Salud Vascular</h3>
            <p class="slogan-tag">Ingrese su Código de Acceso Personalizado para visualizar su programa clínico</p>
        </div>
    """, unsafe_allow_html=True)

    col1, col2, col3 = st.columns([1, 2, 1])
    with col2:
        st.markdown("### 🔐 Autenticación de Paciente")
        input_code = st.text_input("Código de Acceso Único", type="password", placeholder="Ej: PACHECO-2026-VIP")
        
        if st.button("Acceder al Programa", use_container_width=True):
            if input_code in VALID_CODES:
                st.session_state.authenticated = True
                st.session_state.patient_data = VALID_CODES[input_code]
                st.rerun()
            else:
                st.error("❌ Código de acceso inválido o no registrado. Contacte al soporte por WhatsApp.")
                st.markdown("[📲 Solicitar código por WhatsApp](https://wa.me/584120582064?text=Hola%20Lic.%20Carlos%20Pacheco,%20deseo%20solicitar%20mi%20código%20de%20acceso%20para%20el%20programa.)")
        
        st.markdown("---")
        st.info("💡 **¿No tienes un código?** Realiza tu evaluación inicial con el especialista para obtener tu acceso exclusivo desde cualquier país de Latinoamérica.")
    st.stop()

# -----------------------------------------------------------------------------
# INTERFAZ PRINCIPAL (Paciente Autenticado)
# -----------------------------------------------------------------------------
patient = st.session_state.patient_data

with st.sidebar:
    try:
        st.image("assets/logo1.png", use_container_width=True)
    except Exception:
        st.write("🫀 **Lic. Carlos Pacheco**")
    
    st.markdown("---")
    st.markdown(f"### 👤 Paciente: {patient['name']}")
    st.markdown(f"🌍 **Ubicación:** {patient['country']}")
    st.markdown(f"📋 **Diagnóstico Base:** {patient['condition']}")
    st.markdown("---")
    if st.button("🚪 Cerrar Sesión", use_container_width=True):
        st.session_state.authenticated = False
        st.rerun()

# Encabezado institucional
st.markdown(f"""
    <div class="main-header" style="padding: 1.5rem;">
        <h1 style="margin:0; font-size:2.2rem;">Programa Clínico de Entrenamiento Personalizado</h1>
        <h3 style="margin:0; font-weight:400; color:#E2E8F0; font-size:1.1rem;">Bienvenido(a), {patient['name']} | Monitoreo Hemodinámico Transfronterizo</h3>
    </div>
""", unsafe_allow_html=True)

# Pestañas de navegación del programa
tab1, tab2, tab3, tab4 = st.tabs([
    "📊 1. Encuesta de Valoración de Riesgo", 
    "🎥 2. Videos Demostrativos & Rutinas", 
    "📄 3. Documentos & Guías PDF",
    "📱 4. Soporte & Teleconsulta"
])

# -----------------------------------------------------------------------------
# TAB 1: ENCUESTA DE VALORACIÓN DE RIESGO
# -----------------------------------------------------------------------------
with tab1:
    st.header("Evaluación Hemodinámica y Criterios de Esfuerzo")
    st.write("Complete este formulario de control semanal para ajustar de forma segura su pauta de entrenamiento.")

    with st.form("risk_survey"):
        col_s1, col_s2 = st.columns(2)
        with col_s1:
            val_edad = st.number_input("Edad Actual", min_value=18, max_value=100, value=50)
            val_pas = st.slider("Presión Arterial Sistólica Actual (mmHg)", 90, 200, 130)
            val_pad = st.slider("Presión Arterial Diastólica Actual (mmHg)", 60, 120, 85)
        with col_s2:
            val_fatiga = st.selectbox("Nivel de fatiga percibida post-entrenamiento", ["Bajo (Recuperación rápida)", "Moderado (Cansancio normal)", "Alto (Fatiga persistente / Mareos)"])
            val_sueño = st.slider("Horas de sueño promedio por noche", 3, 10, 7)
            val_sintomas = st.multiselect("¿Ha presentado alguno de estos síntomas recientemente?", ["Palpitaciones", "Dolor de cabeza matutino", "Zumbido en los oídos", "Ninguno"])

        submit_survey = st.form_submit_button("Calcular Estratificación de Riesgo Actual")

        if submit_survey:
            score = 10
            if val_pas > 135: score += 30
            if val_pas > 150: score += 20
            if "Ninguno" not in val_sintomas: score += 25
            if val_fatiga == "Alto (Fatiga persistente / Mareos)": score += 20
            score = min(score, 100)

            if score < 40:
                st.success(f"🟢 **Riesgo Bajo (Puntaje: {score}/100):** Su endotelio responde de forma excelente. Continúe con su fase actual de ejercicios aeróbicos dosificados.")
            elif score < 70:
                st.warning(f"🟡 **Riesgo Moderado (Puntaje: {score}/100):** Se detecta ligera resistencia vascular. Reduzca la intensidad de impacto y mantenga control estricto de frecuencia cardíaca.")
            else:
                st.error(f"🔴 **Riesgo Alto (Puntaje: {score}/100):** Alerta hemodinámica. Suspenda entrenamientos de fuerza intensa y contacte de inmediato al Lic. Carlos Pacheco.")

# -----------------------------------------------------------------------------
# TAB 2: VIDEOS DEMOSTRATIVOS & RUTINAS
# -----------------------------------------------------------------------------
with tab2:
    st.header("Biblioteca de Videos Demostrativos Clínicos")
    st.write("Videos instructivos grabados por el especialista para garantizar la técnica correcta sin picos de presión arterial.")

    v_col1, v_col2 = st.columns(2)

    with v_col1:
        st.markdown("""
            <div class="video-box">
                <h4 style="color:#0A2540;">Módulo 1: Activación Endotelial y Respiración Diafragmática</h4>
                <p>Técnica fundamental para estimular la liberación de óxido nítrico antes de cualquier actividad física.</p>
            </div>
        """, unsafe_allow_html=True)
        # Video embebido de demostración (YouTube estándar de referencia médica)
        st.video("https://www.youtube.com/watch?v=5qap5aO4i9A")

    with v_col2:
        st.markdown("""
            <div class="video-box">
                <h4 style="color:#0A2540;">Módulo 2: Control del Estrés de Cizallamiento en HTA</h4>
                <p>Cómo realizar caminatas aeróbicas y ejercicios de bajo impacto protegiendo las arterias carótideas.</p>
            </div>
        """, unsafe_allow_html=True)
        st.video("https://www.youtube.com/watch?v=1XbIuWfOq_M")

# -----------------------------------------------------------------------------
# TAB 3: DOCUMENTOS Y GUÍAS PDF ESPECIALES
# -----------------------------------------------------------------------------
with tab3:
    st.header("Documentación y Protocolos Especiales en PDF")
    st.write("Descargue su guía clínica personalizada elaborada para su condición vascular específica.")

    def create_sample_pdf(title, subtitle):
        buffer = io.BytesIO()
        doc = Document()
        doc.add_heading(title, 0)
        doc.add_paragraph(f"Paciente: {patient['name']} | País: {patient['country']}")
        doc.add_paragraph(f"Condición: {subtitle}")
        doc.add_heading("Instrucciones Clínicas del Lic. Carlos Pacheco:", level=1)
        doc.add_paragraph("1. Realizar monitoreo de presión arterial antes y después de la sesión.")
        doc.add_paragraph("2. Mantener hidratación óptima (agua a temperatura ambiente).")
        doc.add_paragraph("3. Evitar apneas (bloqueos respiratorios) durante los movimientos.")
        doc.save(buffer)
        buffer.seek(0)
        return buffer

    pdf_file = create_sample_pdf("Guía Oficial de Protección Endotelial", patient['condition'])

    st.markdown("""
        <div class="callout-card">
            <h4>📄 Protocolo Clínico Asignado a su Perfil</h4>
            <p>Este documento contiene las pautas exactas de frecuencia, intensidad y tiempo de esfuerzo físico seguro adaptado a su diagnóstico.</p>
        </div>
    """, unsafe_allow_html=True)

    st.download_button(
        label="📥 Descargar Guía Clínica PDF Personalizada",
        data=pdf_file,
        file_name=f"Guia_Clinica_{patient['name'].replace(' ', '_')}.docx",
        mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document"
    )

# -----------------------------------------------------------------------------
# TAB 4: SOPORTE Y TELECONSULTA INTERNACIONAL
# -----------------------------------------------------------------------------
with tab4:
    st.header("Soporte Directo con el Especialista")
    st.write("¿Tiene dudas sobre su rutina o sus valores de presión arterial? Conéctese directamente vía WhatsApp desde cualquier lugar de Latinoamérica.")

    msg_wa = urllib.parse.quote(f"Hola Lic. Carlos Pacheco, le escribo desde {patient['country']}. Soy el paciente {patient['name']} (Código: {input_code}) y tengo una consulta sobre mi programa de entrenamiento.")
    wa_link = f"https://wa.me/584120582064?text={msg_wa}"

    st.markdown(f"""
        <div style="background-color: #E8F5E9; padding: 2rem; border-radius: 10px; text-align: center; border: 1px solid #C8E6C9;">
            <h3 style="color: #2E7D32; margin-top: 0;">💬 Canal Directo de WhatsApp (Soporte Médico)</h3>
            <p style="color: #388E3C; font-size: 1.1rem;">Atención prioritaria para pacientes inscritos en el programa transfronterizo.</p>
            <a href="{wa_link}" target="_blank">
                <button style="background-color: #25D366; color: white; padding: 12px 24px; border: none; border-radius: 8px; font-weight: bold; font-size: 1.1rem; cursor: pointer;">
                    📱 Iniciar Chat con el Lic. Carlos Pacheco
                </button>
            </a>
        </div>
    """, unsafe_allow_html=True)

# Pie institucional
st.markdown("---")
st.markdown("<p style='text-align:center; color:#6C757D;'>© 2026 Lic. Carlos Pacheco - Todos los derechos reservados. Plataforma de Salud Vascular & Protección Endotelial.</p>", unsafe_allow_html=True)
