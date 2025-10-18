import streamlit as st
import pandas as pd
import smtplib
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
from datetime import datetime, timedelta
import time
import sqlite3
import json
import os

# Configuración de página con tema oscuro
st.set_page_config(
    page_title="Sistema de Correos Académicos", 
    page_icon="📧",
    layout="wide",
    initial_sidebar_state="collapsed"
)

# CSS para tema oscuro completo
st.markdown("""
<style>
    /* Fondo principal */
    .main {
        background-color: #0e0e0e;
        color: #ffffff;
    }
    
    /* Contenedor principal */
    .stApp {
        background-color: #0e0e0e;
    }
    
    /* Texto general */
    body, p, span, div, label {
        color: #ffffff !important;
    }
    
    /* Títulos */
    h1, h2, h3, h4, h5, h6 {
        color: #ffffff !important;
        font-weight: 500;
    }
    
    /* Botones */
    .stButton>button {
        background-color: #1db954;
        color: #ffffff;
        border: none;
        padding: 10px 24px;
        border-radius: 8px;
        font-weight: 500;
        transition: all 0.3s;
    }
    
    .stButton>button:hover {
        background-color: #1ed760;
        box-shadow: 0 4px 12px rgba(29, 185, 84, 0.4);
    }
    
    /* Inputs */
    .stTextInput>div>div>input,
    .stTextArea>div>div>textarea,
    .stSelectbox>div>div>select,
    .stDateInput>div>div>input,
    .stTimeInput>div>div>input,
    .stNumberInput>div>div>input {
        background-color: #1a1a1a;
        color: #ffffff;
        border: 1px solid #333333;
        border-radius: 4px;
    }
    
    /* Expanders */
    .streamlit-expanderHeader {
        background-color: #1a1a1a;
        color: #ffffff;
        border-radius: 4px;
    }
    
    /* Métricas */
    .stMetric {
        background-color: #1a1a1a;
        padding: 12px;
        border-radius: 8px;
        border: 1px solid #333333;
    }
    
    .stMetric label {
        color: #b3b3b3 !important;
    }
    
    .stMetric .metric-value {
        color: #1db954 !important;
    }
    
    /* Alertas */
    .stAlert {
        background-color: #1a1a1a;
        border-radius: 8px;
        border-left: 4px solid #1db954;
    }
    
    /* Success */
    .stSuccess {
        background-color: #1a3a1a;
        color: #4ade80;
        border-left: 4px solid #1db954;
    }
    
    /* Error */
    .stError {
        background-color: #3a1a1a;
        color: #f87171;
        border-left: 4px solid #ef4444;
    }
    
    /* Warning */
    .stWarning {
        background-color: #3a3a1a;
        color: #fbbf24;
        border-left: 4px solid #f59e0b;
    }
    
    /* Info */
    .stInfo {
        background-color: #1a2a3a;
        color: #60a5fa;
        border-left: 4px solid #3b82f6;
    }
    
    /* Dataframe */
    .stDataFrame {
        background-color: #1a1a1a;
        color: #ffffff;
    }
    
    /* Tabs */
    .stTabs [data-baseweb="tab-list"] {
        background-color: #1a1a1a;
        border-radius: 8px;
    }
    
    .stTabs [data-baseweb="tab"] {
        color: #b3b3b3;
        background-color: transparent;
    }
    
    .stTabs [aria-selected="true"] {
        color: #1db954;
        border-bottom-color: #1db954;
    }
    
    /* File uploader */
    .stFileUploader {
        background-color: #1a1a1a;
        border: 2px dashed #333333;
        border-radius: 8px;
    }
    
    /* Divider */
    hr {
        border-color: #333333;
    }
    
    /* Caption */
    .caption {
        color: #b3b3b3 !important;
    }
    
    /* Links */
    a {
        color: #1db954;
    }
    
    a:hover {
        color: #1ed760;
    }
    
    /* Sidebar */
    .css-1d391kg {
        background-color: #1a1a1a;
    }
</style>
""", unsafe_allow_html=True)

# =====================================================
# BASE DE DATOS SQLITE PERSISTENTE
# =====================================================

DB_FILE = "sistema_correos.db"

def init_database():
    """Inicializar base de datos SQLite"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    
    # Tabla de cuentas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS cuentas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            nombre TEXT UNIQUE NOT NULL,
            email TEXT NOT NULL,
            password TEXT NOT NULL,
            ultima_uso TIMESTAMP DEFAULT CURRENT_TIMESTAMP
        )
    ''')
    
    # Tabla de plantillas
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS plantillas (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            institucion TEXT NOT NULL,
            tipo TEXT NOT NULL,
            nombre TEXT NOT NULL,
            asunto TEXT NOT NULL,
            mensaje TEXT NOT NULL,
            fecha_modificacion TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            UNIQUE(institucion, tipo)
        )
    ''')
    
    # Tabla de historial
    cursor.execute('''
        CREATE TABLE IF NOT EXISTS historial (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            timestamp TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
            asunto TEXT,
            destinatario TEXT,
            estado TEXT
        )
    ''')
    
    conn.commit()
    conn.close()

def guardar_cuenta_db(nombre, email, password):
    """Guardar cuenta en base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    try:
        cursor.execute('''
            INSERT OR REPLACE INTO cuentas (nombre, email, password, ultima_uso)
            VALUES (?, ?, ?, CURRENT_TIMESTAMP)
        ''', (nombre, email, password))
        conn.commit()
        return True
    except Exception as e:
        st.error(f"Error al guardar cuenta: {e}")
        return False
    finally:
        conn.close()

def cargar_cuentas_db():
    """Cargar todas las cuentas de la base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('SELECT nombre, email, password, ultima_uso FROM cuentas ORDER BY ultima_uso DESC')
    cuentas = {}
    for row in cursor.fetchall():
        cuentas[row[0]] = {
            'email': row[1],
            'password': row[2],
            'ultima_uso': row[3]
        }
    conn.close()
    return cuentas

def eliminar_cuenta_db(nombre):
    """Eliminar cuenta de la base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM cuentas WHERE nombre = ?', (nombre,))
    conn.commit()
    conn.close()

def guardar_plantilla_db(institucion, tipo, nombre, asunto, mensaje):
    """Guardar plantilla en base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT OR REPLACE INTO plantillas (institucion, tipo, nombre, asunto, mensaje, fecha_modificacion)
        VALUES (?, ?, ?, ?, ?, CURRENT_TIMESTAMP)
    ''', (institucion, tipo, nombre, asunto, mensaje))
    conn.commit()
    conn.close()

def cargar_plantilla_db(institucion, tipo):
    """Cargar plantilla de la base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT nombre, asunto, mensaje FROM plantillas 
        WHERE institucion = ? AND tipo = ?
    ''', (institucion, tipo))
    row = cursor.fetchone()
    conn.close()
    if row:
        return {'nombre': row[0], 'asunto': row[1], 'mensaje': row[2]}
    return None

def guardar_historial_db(asunto, destinatario, estado):
    """Guardar en historial"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        INSERT INTO historial (timestamp, asunto, destinatario, estado)
        VALUES (CURRENT_TIMESTAMP, ?, ?, ?)
    ''', (asunto, destinatario, estado))
    conn.commit()
    conn.close()

def cargar_historial_db(limite=100):
    """Cargar historial de base de datos"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('''
        SELECT timestamp, asunto, destinatario, estado 
        FROM historial 
        ORDER BY timestamp DESC 
        LIMIT ?
    ''', (limite,))
    historial = []
    for row in cursor.fetchall():
        historial.append({
            'timestamp': row[0],
            'asunto': row[1],
            'destinatario': row[2],
            'estado': row[3]
        })
    conn.close()
    return historial

def limpiar_historial_db():
    """Limpiar historial"""
    conn = sqlite3.connect(DB_FILE)
    cursor = conn.cursor()
    cursor.execute('DELETE FROM historial')
    conn.commit()
    conn.close()

# Inicializar base de datos al inicio
init_database()

# Título principal
st.title("📧 Sistema de Correos Académicos")
st.caption("UVEG & NovaUniversitas | Tema Oscuro")

# =====================================================
# GESTIÓN DE CREDENCIALES CON BASE DE DATOS
# =====================================================

st.subheader("⚙️ Configuración de Cuenta")

cuentas = cargar_cuentas_db()

col1, col2 = st.columns([2, 1])

with col1:
    if cuentas:
        st.write("**Cuentas guardadas:**")
        
        cuenta_seleccionada = st.selectbox(
            "Seleccionar cuenta:",
            options=["Nueva cuenta..."] + list(cuentas.keys()),
            key="selector_cuenta"
        )
        
        if cuenta_seleccionada != "Nueva cuenta...":
            cuenta_data = cuentas[cuenta_seleccionada]
            st.session_state.email_actual = cuenta_data['email']
            st.session_state.password_actual = cuenta_data['password']
            st.session_state.cuenta_activa = cuenta_seleccionada
            
            st.success(f"✓ Usando cuenta: {cuenta_data['email']}")
            st.caption(f"Último uso: {cuenta_data['ultima_uso']}")
            
            if st.button("🗑️ Eliminar esta cuenta", key="eliminar_cuenta"):
                eliminar_cuenta_db(cuenta_seleccionada)
                if 'email_actual' in st.session_state:
                    del st.session_state.email_actual
                if 'password_actual' in st.session_state:
                    del st.session_state.password_actual
                if 'cuenta_activa' in st.session_state:
                    del st.session_state.cuenta_activa
                st.rerun()
        else:
            with st.form("nueva_cuenta_form"):
                st.write("**Agregar nueva cuenta:**")
                
                col_email, col_pass = st.columns(2)
                
                with col_email:
                    nuevo_email = st.text_input(
                        "Correo electrónico:",
                        placeholder="tu_correo@gmail.com"
                    )
                
                with col_pass:
                    nuevo_password = st.text_input(
                        "Contraseña de aplicación:",
                        type="password",
                        placeholder="abcd efgh ijkl mnop"
                    )
                
                nombre_cuenta = st.text_input(
                    "Nombre para esta cuenta:",
                    placeholder="ej: UVEG Principal, Nova Coordinación"
                )
                
                submitted = st.form_submit_button("✓ Agregar cuenta")
                
                if submitted:
                    if nuevo_email and nuevo_password and nombre_cuenta:
                        with st.spinner("Verificando..."):
                            try:
                                server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10)
                                server.login(nuevo_email, nuevo_password)
                                server.quit()
                                
                                guardar_cuenta_db(nombre_cuenta, nuevo_email, nuevo_password)
                                st.session_state.email_actual = nuevo_email
                                st.session_state.password_actual = nuevo_password
                                st.session_state.cuenta_activa = nombre_cuenta
                                st.success("✓ Cuenta verificada y guardada")
                                st.balloons()
                                time.sleep(1)
                                st.rerun()
                            except Exception as e:
                                st.error(f"✗ Error: {str(e)}")
                    else:
                        st.error("✗ Completa todos los campos")
    else:
        with st.form("primera_cuenta_form"):
            st.write("**Configurar primera cuenta:**")
            
            col_email, col_pass = st.columns(2)
            
            with col_email:
                primer_email = st.text_input(
                    "Correo electrónico:",
                    placeholder="tu_correo@gmail.com"
                )
            
            with col_pass:
                primer_password = st.text_input(
                    "Contraseña de aplicación:",
                    type="password",
                    placeholder="abcd efgh ijkl mnop"
                )
            
            nombre_cuenta = st.text_input(
                "Nombre para esta cuenta:",
                placeholder="ej: UVEG Principal"
            )
            
            submitted = st.form_submit_button("✓ Configurar")
            
            if submitted:
                if primer_email and primer_password and nombre_cuenta:
                    with st.spinner("Verificando..."):
                        try:
                            server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10)
                            server.login(primer_email, primer_password)
                            server.quit()
                            
                            guardar_cuenta_db(nombre_cuenta, primer_email, primer_password)
                            st.session_state.email_actual = primer_email
                            st.session_state.password_actual = primer_password
                            st.session_state.cuenta_activa = nombre_cuenta
                            st.success("✓ Cuenta verificada y guardada")
                            st.balloons()
                            time.sleep(1)
                            st.rerun()
                        except Exception as e:
                            st.error(f"✗ Error: {str(e)}")
                else:
                    st.error("✗ Completa todos los campos")

with col2:
    with st.expander("ℹ️ Ayuda"):
        st.markdown("""
        **Contraseña de aplicación:**
        
        1. Ve a tu cuenta de Google
        2. Seguridad → 2FA
        3. Contraseñas de aplicaciones
        4. Genera nueva
        5. Copia 16 caracteres
        """)

if 'email_actual' not in st.session_state or 'password_actual' not in st.session_state:
    st.info("⚙️ Configura una cuenta para continuar")
    st.stop()

REMITENTE = st.session_state.email_actual
CLAVE_APP = st.session_state.password_actual

# =====================================================
# TABS PRINCIPALES
# =====================================================

tab1, tab2, tab3 = st.tabs([
    "🎯 Sistema Inteligente (UVEG/Nova)", 
    "📝 Sistema Prácticas",
    "🎓 Sistema Bienvenida Nova"
])

# =====================================================
# TAB 1: SISTEMA INTELIGENTE CON BD
# =====================================================

with tab1:
    st.header("🎯 Sistema Inteligente de Seguimiento")
    
    if 'analisis_generado_tab1' not in st.session_state:
        st.session_state.analisis_generado_tab1 = False
    if 'datos_estudiantes_tab1' not in st.session_state:
        st.session_state.datos_estudiantes_tab1 = None

    CONFIGURACIONES_INSTITUCIONES = {
        "uveg": {
            "nombre": "UVEG",
            "columnas_actividades": [
                "Paquete SCORM:R1. Conversiones entre sistemas numéricos (Real)",
                "Paquete SCORM:R2. Operaciones aritméticas con sistema binario, octal y hexadecimal (Real)",
                "Tarea:R3. Operaciones con conjuntos y su representación (Real)",
                "Tarea:R4. Proposiciones lógicas (Real)",
                "Paquete SCORM:R5. Operadores lógicos y tablas de verdad (Real)",
                "Paquete SCORM:R6. Relaciones y operaciones con relaciones (Real)",
                "Tarea:R7. Propiedades de las relaciones: representación gráfica (Real)"
            ],
            "nombres_actividades": [
                "R1. Conversiones entre sistemas numéricos",
                "R2. Operaciones aritméticas con sistema binario, octal y hexadecimal",
                "R3. Operaciones con conjuntos y su representación",
                "R4. Proposiciones lógicas",
                "R5. Operadores lógicos y tablas de verdad",
                "R6. Relaciones y operaciones con relaciones",
                "R7. Propiedades de las relaciones: representación gráfica"
            ],
            "columnas_requeridas": ["Nombre", "Apellido(s)", "Correo Personal", "Dirección Email"],
            "modulo_default": "Matemáticas Discretas"
        },
        "novauniversitas": {
            "nombre": "NovaUniversitas",
            "columnas_actividades": [
                "Examen:Examen desafío 1 (Real)",
                "Examen:Examen desafío 2 (Real)",
                "Tarea:Evaluación desafío 3 (Real)",
                "Tarea:Examen desafío 4 (Real)",
                "Examen:Evaluación desafío 5 (Real)",
                "Examen:Evaluación desafío 6 (Real)",
                "Foro:todo el foro Foro desafío 7 (Real)"
            ],
            "nombres_actividades": [
                "Desafío 1. Examen desafío 1",
                "Desafío 2. Examen desafío 2",
                "Desafío 3. Evaluación desafío 3",
                "Desafío 4. Examen desafío 4",
                "Desafío 5. Evaluación desafío 5",
                "Desafío 6. Evaluación desafío 6",
                "Desafío 7. Foro desafío 7"
            ],
            "columnas_requeridas": ["Nombre", "Correo Personal", "Dirección Email"],
            "modulo_default": "Matemáticas Discretas"
        }
    }

    PLANTILLAS_BASE = {
        "uveg": {
            "bienvenida": {
                "nombre": "Bienvenida al Módulo",
                "asunto": "Bienvenida al módulo {modulo} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y seré tu asesor virtual en el módulo "{modulo}" de la UVEG. Te doy la bienvenida al curso y quiero compartirte algunas recomendaciones para organizar tu avance.

Como sugerencia, te he marcado como meta entregar las primeras tres actividades antes del próximo {fecha_meta} al mediodía. Sin embargo, si comprendes bien los temas, puedes avanzar a tu propio ritmo, ya que el módulo tiene una duración total de 24 días naturales.

El propósito de este mensaje es conocer si tienes algún inconveniente en este momento, como falta de tiempo, dificultades con las actividades, problemas de acceso a un dispositivo o Internet, o si simplemente tu estrategia de estudio no sigue el ritmo sugerido. Además, quiero que sepas que estaré pendiente de tu desarrollo académico y disponible para cualquier duda que tengas.

A partir del martes, recibirás un correo con información sobre tu avance y para mantenernos en contacto. No es mi intención abrumarte, sino brindarte el soporte necesario para que tengas éxito en el curso.

Importante: Te invito a revisar la sección "AVISOS" en tu aula virtual para estar al tanto de cualquier información relevante.

Por último tenemos una cita este {fecha_sesion} a las {hora_sesion} horas, sesión síncrona para resolver dudas. El enlace lo encuentras dentro de tu aula virtual.

Te pido me confirmes de recibido este correo.

¡Mucho éxito en el módulo! Estoy aquí para ayudarte.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_sin_acceso_semana_1": {
                "nombre": "Seguimiento - Sin Acceso (Semana 1)",
                "asunto": "Seguimiento {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" de la UVEG, en este momento estamos en la {semana_actual} semana del módulo, solo hemos avanzado algunos retos, con lo cual no presentas un gran atraso, te invito a que inicies tus actividades dentro de la plataforma (https://campus.uveg.edu.mx) y si tienes alguna duda con toda confianza me puedes contactar ya sea por este medio, por el mensajero de la plataforma o xxx por whatsapp.

Agradecería pudieras comentarme la situación por la cual no has accedido al módulo, espero todo se encuentre bien.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_sin_acceso_semana_2": {
                "nombre": "Seguimiento - Sin Acceso (Semana 2)",
                "asunto": "Seguimiento {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" de la UVEG. He tratado de comunicarme anteriormente contigo sin éxito.

En este momento nos encontramos en la {semana_actual} semana del módulo y observo que aún no has iniciado ninguna actividad en la plataforma. Aunque todavía hay tiempo para ponerte al día, es importante que comiences cuanto antes para evitar un atraso mayor.

Te invito cordialmente a que accedas a tu aula virtual (https://campus.uveg.edu.mx) y empieces con los retos programados. Recuerda que estoy aquí para apoyarte con cualquier duda o dificultad que tengas, ya sea por este medio, por el mensajero de la plataforma o xxx por whatsapp.

Me preocupa no haber tenido noticias tuyas. Por favor, hazme saber si existe alguna situación que esté impidiendo tu participación en el módulo. Tu éxito académico es importante para mí.

Quedo al pendiente de tu respuesta.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_sin_acceso_semana_3": {
                "nombre": "Seguimiento - Sin Acceso (Semana 3)",
                "asunto": "Seguimiento URGENTE {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" de la UVEG. Espero puedas culminar tus actividades ya que el módulo está próximo a cerrar. No pude contactarte anteriormente, espero que todo vaya bien.

Nos encontramos en la {semana_actual} semana del módulo y lamentablemente no he detectado actividad alguna de tu parte en la plataforma. El módulo cerrará la próxima semana y es crucial que inicies tus entregas lo antes posible para poder acreditar.

Esta es tu última oportunidad para ponerte al corriente. Te exhorto a que ingreses de inmediato a tu aula virtual (https://campus.uveg.edu.mx) y comiences con los retos. Aunque el tiempo es limitado, aún es posible completar el módulo si actúas con rapidez.

Si enfrentas alguna dificultad técnica, personal o académica, por favor comunícate conmigo de inmediato por cualquier medio: este correo, mensajero de la plataforma o xxx por whatsapp. Estoy dispuesto a brindarte todo el apoyo necesario.

Espero sinceramente tener noticias tuyas y poder ayudarte a concluir exitosamente este módulo.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_atraso_semana_1": {
                "nombre": "Seguimiento - Tareas Pendientes (Semana 1)",
                "asunto": "Seguimiento de avance {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" en la UVEG, en este momento estamos en la {semana_actual} semana del módulo. Me pongo en contacto, para saber si tienes algún inconveniente que no te este permitiendo avanzar al ritmo que les he marcado, ya que veo que tienes algunos retos pendientes y en tu avance semanal se ve reflejado.

Retos pendientes:
{actividades_faltantes}

Recuerda que esta es solo una sugerencia, si consideras que no tienes inconveniente en avanzar a tu ritmo, haz caso omiso a este y los siguientes mensajes.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_atraso_semana_2": {
                "nombre": "Seguimiento - Tareas Pendientes (Semana 2)",
                "asunto": "Seguimiento de avance {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" en la UVEG. Nos encontramos en la {semana_actual} semana del módulo y me comunico nuevamente contigo para dar seguimiento a tu progreso académico.

He revisado tu avance y observo que aún tienes retos pendientes que deberías haber completado hasta esta semana. Estos son los retos que tienes pendientes hasta esta semana:

Retos pendientes:
{actividades_faltantes}

Es importante que consideres ponerte al día con estas actividades para mantener un buen ritmo y evitar acumulación de trabajo al final del módulo. Si estás enfrentando alguna dificultad específica con los contenidos o tienes algún impedimento para avanzar, por favor comunícamelo para poder apoyarte.

Recuerda que estas recomendaciones buscan ayudarte a tener éxito en el módulo. Si prefieres manejar tu propio ritmo de estudio, puedes hacer caso omiso a estos mensajes.

Estoy disponible para resolver cualquier duda que tengas.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_atraso_semana_3": {
                "nombre": "Seguimiento - Tareas Pendientes (Semana 3)",
                "asunto": "Seguimiento URGENTE de avance {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" en la UVEG. Estamos en la {semana_actual} semana y debo informarte que el módulo cierra la próxima semana.

He realizado una revisión de tu expediente y me preocupa observar que tienes varios retos sin completar. Estos son TODOS tus retos pendientes:

Retos pendientes:
{actividades_faltantes}

Es fundamental que te enfoques en completar estas actividades durante los próximos días para poder acreditar el módulo. El tiempo se está agotando y necesitas actuar con urgencia.

Te ofrezco mi apoyo total para ayudarte a concluir. Si tienes dudas sobre algún tema, dificultades técnicas o necesitas orientación para organizar tu tiempo, no dudes en contactarme de inmediato por este medio, el mensajero de la plataforma o xxx por whatsapp.

Esta es una situación crítica pero aún estás a tiempo de recuperarte. Te insto a que dediques el máximo esfuerzo posible en estos días finales.

Quedo al pendiente y disponible para apoyarte en todo lo que necesites.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_sin_acceso_semana_4": {
                "nombre": "Seguimiento - Sin Acceso (Semana 4)",
                "asunto": "ÚLTIMA OPORTUNIDAD {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" de la UVEG. Este es mi último intento de comunicación contigo.

Nos encontramos en la {semana_actual} y ÚLTIMA semana del módulo. Lamentablemente, no he detectado ninguna actividad de tu parte durante todo el módulo y el cierre es INMINENTE.

Esta es tu ÚLTIMA OPORTUNIDAD para acreditar el módulo. Si no entregas tus actividades en los próximos días, reprobarás el curso.

Te URJO a que ingreses de inmediato a tu aula virtual (https://campus.uveg.edu.mx) y comiences con los retos. El tiempo se ha agotado y necesitas actuar AHORA MISMO.

Si existe alguna razón de fuerza mayor que te ha impedido participar, comunícate conmigo URGENTEMENTE por cualquier medio: este correo, mensajero de la plataforma o xxx por whatsapp.

Este es el cierre del módulo. Por favor, hazme saber tu situación lo antes posible.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "seguimiento_atraso_semana_4": {
                "nombre": "Seguimiento - Tareas Pendientes (Semana 4)",
                "asunto": "CIERRE DEL MÓDULO {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.

Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" en la UVEG. Este mensaje es CRÍTICO: estamos en la {semana_actual} y ÚLTIMA semana del módulo, y el CIERRE ES INMINENTE.

He revisado tu expediente y tienes los siguientes retos pendientes que DEBES completar URGENTEMENTE para poder acreditar:

Retos pendientes:
{actividades_faltantes}

QUEDAN MUY POCOS DÍAS para el cierre del módulo. Si no completas estos retos INMEDIATAMENTE, reprobarás el curso.

Esta es tu ÚLTIMA OPORTUNIDAD. Te exhorto a que dediques TODO tu tiempo disponible en las próximas horas/días para completar las actividades pendientes.

Estoy disponible las 24 horas para resolver tus dudas de manera urgente. Contáctame de inmediato por este medio, el mensajero de la plataforma o xxx por whatsapp si necesitas ayuda.

El tiempo se acabó. Actúa AHORA.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            "felicitacion": {
                "nombre": "Felicitación por Desempeño",
                "asunto": "Felicitaciones por tu desempeño - {modulo} - UVEG",
                "mensaje": """Un gusto saludarte {nombre}.

Por medio del presente, permíteme felicitarte por tu alto desempeño durante esta semana, con esto demuestras tu compromiso para con tu carrera y la resiliencia del día a día.

Retos completados:
{actividades_completadas}

Continua así y no olvides revisar el tablero de avisos.

Dr. Juan Manuel Martínez Zaragoza"""
            }
        },
        "novauniversitas": {
            "bienvenida": {
                "nombre": "Bienvenida al Módulo",
                "asunto": "Bienvenida al módulo {modulo} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Mi nombre es Juan Manuel y seré tu asesor virtual en el módulo "{modulo}" de NovaUniversitas. Te doy la más cordial bienvenida al curso y quiero compartirte algunas recomendaciones importantes para tu éxito académico.

Como guía inicial, te sugiero completar los primeros tres desafíos antes del {fecha_meta}. Sin embargo, tienes la flexibilidad de avanzar a tu propio ritmo, considerando que el módulo tiene una duración de 24 días naturales.

Este mensaje tiene como propósito identificar si enfrentas algún inconveniente que pueda afectar tu rendimiento académico, tales como:
- Limitaciones de tiempo
- Dificultades técnicas con los desafíos
- Problemas de conectividad o acceso a dispositivos
- Necesidad de ajustar tu estrategia de estudio

Estaré monitoreando constantemente tu progreso académico y me encuentro disponible para resolver cualquier duda o inquietud que puedas tener.

A partir de la próxima semana, recibirás comunicaciones periódicas sobre tu avance académico para mantener un seguimiento personalizado de tu aprendizaje.

Importante: Te recomiendo revisar regularmente la sección de "ANUNCIOS" en tu campus virtual para mantenerte informado sobre comunicaciones relevantes.

Tenemos programada una sesión síncrona el {fecha_sesion} a las {hora_sesion} horas para resolución de dudas. Podrás acceder a través del enlace disponible en tu aula virtual.

Te solicito confirmar la recepción de este correo.

¡Te deseo mucho éxito en tu trayectoria académica!

Dr. Juan Manuel Martínez Zaragoza
Asesor Virtual - NovaUniversitas"""
            },
            "seguimiento_sin_acceso": {
                "nombre": "Seguimiento - Sin Acceso",
                "asunto": "Seguimiento académico {modulo} - Semana {semana} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Mi nombre es Juan Manuel, tu asesor virtual del módulo "{modulo}" en NovaUniversitas. Me dirijo a ti en relación a tu progreso académico en la {semana_actual} semana del módulo.

He observado que aún no has iniciado actividades en la plataforma educativa. Aunque esto no representa un atraso significativo en este momento, es importante que comiences con los desafíos programados para mantener un ritmo adecuado de aprendizaje.

Te invito cordialmente a:
- Acceder a tu campus virtual de NovaUniversitas
- Revisar los desafíos disponibles
- Contactarme ante cualquier duda o dificultad

Estoy disponible para brindarte apoyo a través de:
- Este correo electrónico
- Mensajería interna del campus virtual
- WhatsApp: xxx (agrega tu número)

Me interesa conocer si existe alguna situación particular que esté impidiendo tu participación en el módulo. Tu bienestar y éxito académico son mi prioridad.

Quedo atento a tu pronta respuesta.

Saludos cordiales,

Dr. Juan Manuel Martínez Zaragoza
Asesor Virtual - NovaUniversitas"""
            },
            "seguimiento_atraso": {
                "nombre": "Seguimiento - Desafíos Pendientes",
                "asunto": "Seguimiento de progreso {modulo} - Semana {semana} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Mi nombre es Juan Manuel, tu asesor virtual del módulo "{modulo}" en NovaUniversitas. Me pongo en contacto contigo para hacer un seguimiento de tu progreso académico en la {semana_actual} semana del módulo.

He revisado tu expediente académico y he identificado algunos desafíos pendientes que requieren tu atención para mantener el ritmo de aprendizaje sugerido:

Desafíos pendientes por completar:
{actividades_faltantes}

Es importante mencionar que estas recomendaciones de ritmo están diseñadas para optimizar tu experiencia de aprendizaje. Si consideras que puedes manejar un ritmo diferente y no requieres este seguimiento, puedes hacer caso omiso a estas comunicaciones.

Sin embargo, si necesitas apoyo o tienes alguna dificultad específica, estoy aquí para ayudarte. Podemos trabajar juntos en una estrategia personalizada que se adapte a tus necesidades.

Opciones de contacto:
- Responder a este correo
- Mensajería del campus virtual
- WhatsApp: xxx (agrega tu número)

Tu éxito académico es importante para mí.

Saludos cordiales,

Dr. Juan Manuel Martínez Zaragoza
Asesor Virtual - NovaUniversitas"""
            },
            "felicitacion": {
                "nombre": "Reconocimiento Académico",
                "asunto": "Felicitaciones por tu excelente desempeño - {modulo} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Es un placer dirigirme a ti para reconocer tu destacado desempeño académico en el módulo "{modulo}" de NovaUniversitas durante esta semana.

Tu dedicación y compromiso con tu formación profesional son evidentes a través de los resultados obtenidos:

Desafíos completados exitosamente:
{actividades_completadas}

Este nivel de excelencia académica refleja tu seriedad y determinación hacia tus objetivos educativos. Tu constancia y esfuerzo son cualidades que sin duda te llevarán al éxito profesional.

Te motivo a continuar con esta actitud ejemplar y te recordamos revisar periódicamente el tablero de anuncios en tu campus virtual para mantenerte informado sobre novedades importantes.

Sigue adelante con esa misma dedicación. ¡Tu futuro profesional se construye con cada logro como este!

Felicitaciones nuevamente por tu excelente trabajo.

Saludos cordiales,

Dr. Juan Manuel Martínez Zaragoza
Asesor Virtual - NovaUniversitas"""
            }
        }
    }

    def convertir_a_numerico(valor):
        if pd.isna(valor):
            return 0
        if isinstance(valor, str):
            valor_limpio = valor.strip()
            if not valor_limpio:
                return 0
            try:
                return float(valor_limpio)
            except ValueError:
                return 0
        try:
            return float(valor)
        except (ValueError, TypeError):
            return 0

    def contar_actividades_completadas(row, columnas_actividades):
        completadas = 0
        for col in columnas_actividades:
            if col in row.index:
                valor_numerico = convertir_a_numerico(row[col])
                if valor_numerico > 0:
                    completadas += 1
        return completadas

    def obtener_actividades_completadas(row, columnas_actividades, nombres_actividades):
        actividades_completadas = []
        for i, col in enumerate(columnas_actividades):
            if col in row.index:
                valor_numerico = convertir_a_numerico(row[col])
                if valor_numerico > 0:
                    actividades_completadas.append(nombres_actividades[i])
        return actividades_completadas

    def obtener_actividades_faltantes(row, columnas_actividades, nombres_actividades, actividades_requeridas):
        actividades_faltantes = []
        for i in range(actividades_requeridas):
            col = columnas_actividades[i]
            nombre = nombres_actividades[i]
            if col in row.index:
                valor_numerico = convertir_a_numerico(row[col])
                if valor_numerico <= 0:
                    actividades_faltantes.append(nombre)
            else:
                actividades_faltantes.append(nombre)
        return actividades_faltantes

    def validar_email(email):
        if not email or pd.isna(email):
            return False
        email_str = str(email).strip()
        return "@" in email_str and "." in email_str.split("@")[-1]

    def obtener_emails_validos(estudiante):
        emails_validos = []
        if 'Correo Personal' in estudiante.index and validar_email(estudiante['Correo Personal']):
            emails_validos.append(str(estudiante['Correo Personal']).strip())
        if 'Dirección Email' in estudiante.index and validar_email(estudiante['Dirección Email']):
            email_institucional = str(estudiante['Dirección Email']).strip()
            emails_validos.append(email_institucional)
        return emails_validos

    def obtener_nombre_completo(estudiante, institucion):
        nombre = str(estudiante.get('Nombre', 'Apreciable estudiante'))
        if institucion == "uveg":
            apellidos = str(estudiante.get('Apellido(s)', ''))
            return nombre, apellidos
        else:
            return nombre, ""

    def formatear_fecha(fecha_str):
        try:
            fecha = datetime.strptime(fecha_str, "%Y-%m-%d")
            dias = ["lunes", "martes", "miércoles", "jueves", "viernes", "sábado", "domingo"]
            meses = ["enero", "febrero", "marzo", "abril", "mayo", "junio",
                    "julio", "agosto", "septiembre", "octubre", "noviembre", "diciembre"]
            dia_semana = dias[fecha.weekday()]
            dia = fecha.day
            mes = meses[fecha.month - 1]
            return f"{dia_semana} {dia} de {mes}"
        except:
            return fecha_str

    def enviar_correo_con_adjuntos(destinatario, asunto, cuerpo, archivos_adjuntos=None):
        try:
            msg = MIMEMultipart()
            msg["Subject"] = asunto
            msg["From"] = REMITENTE
            msg["To"] = destinatario
            msg.attach(MIMEText(cuerpo, 'plain', 'utf-8'))
            
            if archivos_adjuntos:
                for archivo in archivos_adjuntos:
                    try:
                        archivo.seek(0)
                        mime_type, _ = mimetypes.guess_type(archivo.name)
                        if mime_type is None:
                            mime_type = 'application/octet-stream'
                        maintype, subtype = mime_type.split('/', 1)
                        part = MIMEBase(maintype, subtype)
                        part.set_payload(archivo.read())
                        encoders.encode_base64(part)
                        part.add_header('Content-Disposition', f'attachment; filename="{archivo.name}"')
                        msg.attach(part)
                    except Exception as e:
                        st.warning(f"No se pudo adjuntar {archivo.name}: {str(e)}")
            
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(REMITENTE, CLAVE_APP)
                smtp.send_message(msg)
            
            guardar_historial_db(asunto, destinatario, 'Enviado')
            return True, "Enviado correctamente"
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            guardar_historial_db(asunto, destinatario, error_msg)
            return False, error_msg

    def enviar_a_todos_los_emails(emails_validos, asunto, cuerpo, nombre_estudiante, archivos_adjuntos=None):
        resultados = []
        emails_enviados = []
        
        for email in emails_validos:
            if email and email not in emails_enviados:
                exito, mensaje = enviar_correo_con_adjuntos(email, asunto, cuerpo, archivos_adjuntos)
                resultados.append({'email': email, 'exito': exito, 'mensaje': mensaje})
                emails_enviados.append(email)
                
                if exito:
                    st.success(f"✓ {email}")
                else:
                    st.error(f"✗ {email}: {mensaje}")
                
                time.sleep(0.8)
        
        exitosos = sum(1 for r in resultados if r['exito'])
        return exitosos, len(resultados), emails_enviados

    def obtener_plantilla(institucion, tipo, semana=None):
        """Obtiene la plantilla correcta según institución, tipo y semana"""
        # Para UVEG, si se proporciona semana, usar la plantilla específica
        if institucion == "uveg" and semana and tipo in ["seguimiento_sin_acceso", "seguimiento_atraso"]:
            tipo_con_semana = f"{tipo}_semana_{semana}"
            plantilla_db = cargar_plantilla_db(institucion, tipo_con_semana)
            if plantilla_db:
                return plantilla_db
            elif tipo_con_semana in PLANTILLAS_BASE[institucion]:
                return PLANTILLAS_BASE[institucion][tipo_con_semana].copy()
        
        # Caso por defecto
        plantilla_db = cargar_plantilla_db(institucion, tipo)
        if plantilla_db:
            return plantilla_db
        else:
            return PLANTILLAS_BASE[institucion][tipo].copy()

    st.divider()
    
    col1, col2 = st.columns([3, 1])
    
    with col1:
        institucion_seleccionada = st.selectbox(
            "Institución:",
            options=["uveg", "novauniversitas"],
            format_func=lambda x: CONFIGURACIONES_INSTITUCIONES[x]["nombre"],
            key="institucion_tab1"
        )
        config_institucion = CONFIGURACIONES_INSTITUCIONES[institucion_seleccionada]
    
    with col2:
        archivos_adjuntos = st.file_uploader(
            "Adjuntos", 
            accept_multiple_files=True,
            key="archivos_tab1"
        )

    st.divider()
    st.subheader("📝 Plantillas")
    
    # Selector de tipo de plantilla
    opciones_plantilla = ["bienvenida", "seguimiento_sin_acceso", "seguimiento_atraso", "felicitacion"]
    
    # Para mostrar los nombres, necesitamos manejar las plantillas base de forma diferente
    def obtener_nombre_plantilla_base(tipo):
        if institucion_seleccionada == "uveg":
            if tipo == "seguimiento_sin_acceso":
                return "Seguimiento - Sin Acceso"
            elif tipo == "seguimiento_atraso":
                return "Seguimiento - Tareas Pendientes"
        # Buscar en plantillas base
        if tipo in PLANTILLAS_BASE[institucion_seleccionada]:
            return PLANTILLAS_BASE[institucion_seleccionada][tipo]["nombre"]
        return tipo
    
    tipo_plantilla_editar = st.selectbox(
        "Tipo de Plantilla:",
        options=opciones_plantilla,
        format_func=obtener_nombre_plantilla_base,
        key="tipo_plantilla_tab1"
    )
    
    # Si es UVEG y es un tipo que tiene variaciones por semana, mostrar selector de semana
    semana_plantilla = 1
    if institucion_seleccionada == "uveg" and tipo_plantilla_editar in ["seguimiento_sin_acceso", "seguimiento_atraso"]:
        semana_plantilla = st.radio(
            "Semana:",
            options=[1, 2, 3],
            format_func=lambda x: f"Semana {x}",
            horizontal=True,
            key="semana_plantilla_tab1"
        )
        tipo_plantilla_con_semana = f"{tipo_plantilla_editar}_semana_{semana_plantilla}"
        plantilla_actual = obtener_plantilla(institucion_seleccionada, tipo_plantilla_editar, semana_plantilla)
    else:
        plantilla_actual = obtener_plantilla(institucion_seleccionada, tipo_plantilla_editar)
    
    col_asunto, col_reset = st.columns([4, 1])
    
    with col_asunto:
        nuevo_asunto = st.text_input(
            "Asunto:",
            value=plantilla_actual["asunto"],
            key="asunto_edit_tab1"
        )
    
    with col_reset:
        if st.button("↺", key="restaurar_tab1", help="Restaurar"):
            if institucion_seleccionada == "uveg" and tipo_plantilla_editar in ["seguimiento_sin_acceso", "seguimiento_atraso"]:
                tipo_con_semana = f"{tipo_plantilla_editar}_semana_{semana_plantilla}"
                if tipo_con_semana in PLANTILLAS_BASE[institucion_seleccionada]:
                    plantilla_actual = PLANTILLAS_BASE[institucion_seleccionada][tipo_con_semana].copy()
            else:
                plantilla_actual = PLANTILLAS_BASE[institucion_seleccionada][tipo_plantilla_editar].copy()
            st.rerun()
    
    nuevo_mensaje = st.text_area(
        "Mensaje:",
        value=plantilla_actual["mensaje"],
        height=350,
        key="mensaje_edit_tab1"
    )
    
    # Determinar el tipo de plantilla para guardar
    tipo_para_guardar = tipo_plantilla_editar
    if institucion_seleccionada == "uveg" and tipo_plantilla_editar in ["seguimiento_sin_acceso", "seguimiento_atraso"]:
        tipo_para_guardar = f"{tipo_plantilla_editar}_semana_{semana_plantilla}"
    
    if nuevo_asunto != plantilla_actual["asunto"] or nuevo_mensaje != plantilla_actual["mensaje"]:
        guardar_plantilla_db(
            institucion_seleccionada, 
            tipo_para_guardar, 
            plantilla_actual["nombre"],
            nuevo_asunto,
            nuevo_mensaje
        )
        st.caption("✓ Guardado")

    st.divider()
    archivo_excel = st.file_uploader(
        "📄 Archivo Excel", 
        type=["xlsx", "xls"],
        key="archivo_excel_tab1"
    )

    if archivo_excel:
        try:
            df = pd.read_excel(archivo_excel)
            st.success(f"✓ {len(df)} estudiantes")
            
            if institucion_seleccionada == "novauniversitas":
                columnas_necesarias = config_institucion['columnas_requeridas'] + config_institucion['columnas_actividades']
                columnas_existentes = [col for col in columnas_necesarias if col in df.columns]
                if len(columnas_existentes) >= len(config_institucion['columnas_requeridas']):
                    df = df[columnas_existentes]
            
            columnas_faltantes = [col for col in config_institucion['columnas_requeridas'] if col not in df.columns]
            
            if columnas_faltantes:
                st.error(f"✗ Columnas faltantes: {', '.join(columnas_faltantes)}")
            else:
                with st.expander("Vista previa"):
                    st.dataframe(df[config_institucion['columnas_requeridas']].head(5), use_container_width=True)
                
                total_estudiantes = len(df)
                emails_validos = sum(1 for _, fila in df.iterrows() if obtener_emails_validos(fila))
                st.info(f"📊 {total_estudiantes} estudiantes | {emails_validos} con emails")
                
                st.subheader("⚙️ Configuración")
                
                col1, col2, col3 = st.columns(3)
                
                with col1:
                    tipo_envio = st.selectbox(
                        "Tipo:",
                        options=["automatico", "bienvenida"],
                        format_func=lambda x: "Análisis" if x == "automatico" else "Bienvenida",
                        key="tipo_envio_tab1"
                    )
                
                with col2:
                    if tipo_envio == "automatico":
                        semana_calificada = st.selectbox(
                            "Semana calificada:",
                            [1, 2, 3],
                            key="semana_tab1"
                        )
                        semana_mensaje = semana_calificada + 1
                        actividades_requeridas = {1: 3, 2: 5, 3: 7}[semana_calificada]
                        st.caption(f"Mensaje: Semana {semana_mensaje}")
                    else:
                        semana_calificada = 1
                        semana_mensaje = 1
                        actividades_requeridas = 0
                
                with col3:
                    modulo_manual = st.text_input("Módulo:", 
                                                value=config_institucion['modulo_default'],
                                                key="modulo_tab1")
                
                # Solo mostrar campos de sesión si NO es semana 4 (cuando semana_calificada = 3)
                mostrar_sesion = not (tipo_envio == "automatico" and semana_mensaje == 4)
                
                if mostrar_sesion:
                    col4, col5 = st.columns(2)
                    
                    with col4:
                        fecha_meta = st.date_input("Fecha meta:", 
                                                 value=datetime.now() + timedelta(days=7),
                                                 key="fecha_meta_tab1")
                        fecha_sesion = st.date_input("Fecha sesión:", 
                                                   value=datetime.now() + timedelta(days=3),
                                                   key="fecha_sesion_tab1")
                    
                    with col5:
                        hora_sesion = st.time_input("Hora:", 
                                                  value=datetime.strptime("20:30", "%H:%M").time(),
                                                  key="hora_sesion_tab1")
                else:
                    # Solo mostrar fecha meta para semana 4
                    fecha_meta = st.date_input("Fecha meta:", 
                                             value=datetime.now() + timedelta(days=7),
                                             key="fecha_meta_tab1")
                    st.info("ℹ️ En la semana 4 no hay sesiones síncronas programadas.")
                
                if st.button("🚀 Procesar", type="primary", key="procesar_tab1"):
                    
                    # Crear variables_extra básicas
                    variables_extra = {
                        'modulo': modulo_manual,
                        'semana': f"semana {semana_mensaje}",
                        'semana_actual': f"semana {semana_mensaje}",
                        'institucion': config_institucion['nombre'],
                        'fecha_meta': formatear_fecha(fecha_meta.strftime("%Y-%m-%d"))
                    }
                    
                    # Solo agregar variables de sesión si NO es semana 4
                    if mostrar_sesion:
                        variables_extra['fecha_sesion'] = formatear_fecha(fecha_sesion.strftime("%Y-%m-%d"))
                        variables_extra['hora_sesion'] = hora_sesion.strftime("%H:%M")
                    
                    if tipo_envio == "automatico":
                        df['actividades_completadas'] = df.apply(
                            lambda row: contar_actividades_completadas(row, config_institucion['columnas_actividades']), 
                            axis=1
                        )
                        
                        estudiantes_completos = df[df['actividades_completadas'] >= actividades_requeridas]
                        estudiantes_incompletos = df[(df['actividades_completadas'] > 0) & (df['actividades_completadas'] < actividades_requeridas)]
                        estudiantes_sin_entregas = df[df['actividades_completadas'] == 0]
                        
                        st.session_state.datos_estudiantes_tab1 = {
                            'completos': estudiantes_completos,
                            'incompletos': estudiantes_incompletos,
                            'sin_entregas': estudiantes_sin_entregas,
                            'bienvenida': pd.DataFrame(),
                            'semana': semana_mensaje,
                            'semana_calificada': semana_calificada,
                            'actividades_requeridas': actividades_requeridas,
                            'institucion': institucion_seleccionada,
                            'config_institucion': config_institucion,
                            'variables_extra': variables_extra,
                            'tipo_envio': 'automatico'
                        }
                        
                        col1, col2, col3 = st.columns(3)
                        col1.metric("✓ Completos", len(estudiantes_completos))
                        col2.metric("⚠ Incompletos", len(estudiantes_incompletos))
                        col3.metric("✗ Sin Entregas", len(estudiantes_sin_entregas))
                        
                        st.success(f"✓ Análisis completado")
                    
                    else:
                        st.session_state.datos_estudiantes_tab1 = {
                            'completos': pd.DataFrame(),
                            'incompletos': pd.DataFrame(),
                            'sin_entregas': pd.DataFrame(),
                            'bienvenida': df,
                            'semana': semana_mensaje,
                            'actividades_requeridas': actividades_requeridas,
                            'institucion': institucion_seleccionada,
                            'config_institucion': config_institucion,
                            'variables_extra': variables_extra,
                            'tipo_envio': 'bienvenida'
                        }
                        
                        st.success(f"✓ Listo: {len(df)} estudiantes")
                    
                    st.session_state.analisis_generado_tab1 = True
                    
        except Exception as e:
            st.error(f"✗ Error: {str(e)}")

    if st.session_state.analisis_generado_tab1 and st.session_state.datos_estudiantes_tab1:
        datos = st.session_state.datos_estudiantes_tab1
        estudiantes_completos = datos['completos']
        estudiantes_incompletos = datos['incompletos']
        estudiantes_sin_entregas = datos['sin_entregas']
        estudiantes_bienvenida = datos['bienvenida']
        config_institucion = datos['config_institucion']
        variables_extra = datos['variables_extra']
        tipo_envio = datos['tipo_envio']
        
        st.divider()
        st.subheader("👁️ Vista Previa y Edición de Mensajes")
        
        # Determinar qué tipos de mensaje mostrar según el tipo de envío
        if tipo_envio == "automatico":
            tipos_disponibles = []
            if len(estudiantes_completos) > 0:
                tipos_disponibles.append(("Felicitaciones", "felicitacion", estudiantes_completos.iloc[0]))
            if len(estudiantes_incompletos) > 0:
                tipos_disponibles.append(("Recordatorios", "seguimiento_atraso", estudiantes_incompletos.iloc[0]))
            if len(estudiantes_sin_entregas) > 0:
                tipos_disponibles.append(("Alertas", "seguimiento_sin_acceso", estudiantes_sin_entregas.iloc[0]))
            
            if tipos_disponibles:
                tipo_mensaje_preview = st.selectbox(
                    "Selecciona el tipo de mensaje a previsualizar/editar:",
                    options=[t[0] for t in tipos_disponibles],
                    key="tipo_mensaje_preview"
                )
                
                # Encontrar el tipo seleccionado
                tipo_plantilla_key = None
                estudiante_ejemplo = None
                for nombre, key, est in tipos_disponibles:
                    if nombre == tipo_mensaje_preview:
                        tipo_plantilla_key = key
                        estudiante_ejemplo = est
                        break
                
                if tipo_plantilla_key and estudiante_ejemplo is not None:
                    # Obtener la semana correcta para la plantilla
                    semana_plantilla = datos.get('semana', 1)
                    
                    # Obtener plantilla actual
                    if tipo_plantilla_key in ["seguimiento_atraso", "seguimiento_sin_acceso"] and datos['institucion'] == "uveg":
                        plantilla = obtener_plantilla(datos['institucion'], tipo_plantilla_key, semana_plantilla)
                    else:
                        plantilla = obtener_plantilla(datos['institucion'], tipo_plantilla_key)
                    
                    # Generar ejemplo del mensaje
                    nombre, apellidos = obtener_nombre_completo(estudiante_ejemplo, datos['institucion'])
                    
                    if tipo_plantilla_key == "felicitacion":
                        actividades_completadas = obtener_actividades_completadas(
                            estudiante_ejemplo,
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades']
                        )
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_completadas)])
                        mensaje_ejemplo = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_completadas=actividades_lista,
                            **variables_extra
                        )
                    elif tipo_plantilla_key == "seguimiento_atraso":
                        num_actividades_a_revisar = len(config_institucion['columnas_actividades']) if semana_plantilla == 4 else datos['actividades_requeridas']
                        actividades_faltantes = obtener_actividades_faltantes(
                            estudiante_ejemplo,
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades'],
                            num_actividades_a_revisar
                        )
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_faltantes)])
                        mensaje_ejemplo = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_faltantes=actividades_lista,
                            **variables_extra
                        )
                    else:  # seguimiento_sin_acceso
                        mensaje_ejemplo = plantilla['mensaje'].format(
                            nombre=nombre,
                            **variables_extra
                        )
                    
                    asunto_ejemplo = plantilla['asunto'].format(**variables_extra)
                    
                    st.info(f"📌 Ejemplo para: **{nombre} {apellidos}**")
                    
                    col_asunto_prev, col_save_prev = st.columns([5, 1])
                    
                    with col_asunto_prev:
                        nuevo_asunto_prev = st.text_input(
                            "Asunto del mensaje:",
                            value=asunto_ejemplo,
                            key="asunto_preview_edit"
                        )
                    
                    with col_save_prev:
                        st.write("")  # Espaciado
                        st.write("")  # Espaciado
                        if st.button("💾", key="guardar_cambios_prev", help="Guardar cambios en BD"):
                            # Determinar el tipo correcto para guardar
                            if tipo_plantilla_key in ["seguimiento_atraso", "seguimiento_sin_acceso"] and datos['institucion'] == "uveg":
                                tipo_guardar = f"{tipo_plantilla_key}_semana_{semana_plantilla}"
                            else:
                                tipo_guardar = tipo_plantilla_key
                            
                            # Guardar en BD
                            guardar_plantilla_db(
                                datos['institucion'],
                                tipo_guardar,
                                plantilla['nombre'],
                                nuevo_asunto_prev,
                                st.session_state.get('mensaje_preview_edit_value', plantilla['mensaje'])
                            )
                            st.success("✅ Cambios guardados en la base de datos")
                            time.sleep(1)
                            st.rerun()
                    
                    # Editor de mensaje
                    nuevo_mensaje_prev = st.text_area(
                        "Mensaje (puedes incluir variables como {nombre}, {actividades_completadas}, {actividades_faltantes}, etc.):",
                        value=plantilla['mensaje'],
                        height=400,
                        key="mensaje_preview_edit",
                        help="Este es el mensaje base. Las variables entre {} serán reemplazadas automáticamente para cada estudiante."
                    )
                    
                    # Guardar el valor en session_state para poder accederlo al guardar
                    st.session_state['mensaje_preview_edit_value'] = nuevo_mensaje_prev
                    
                    # Mostrar cómo se verá el mensaje final
                    with st.expander("👁️ Vista previa del mensaje final (con variables reemplazadas)", expanded=False):
                        if tipo_plantilla_key == "felicitacion":
                            mensaje_final_preview = nuevo_mensaje_prev.format(
                                nombre=nombre,
                                actividades_completadas=actividades_lista,
                                **variables_extra
                            )
                        elif tipo_plantilla_key == "seguimiento_atraso":
                            mensaje_final_preview = nuevo_mensaje_prev.format(
                                nombre=nombre,
                                actividades_faltantes=actividades_lista,
                                **variables_extra
                            )
                        else:
                            mensaje_final_preview = nuevo_mensaje_prev.format(
                                nombre=nombre,
                                **variables_extra
                            )
                        
                        st.text_area("", value=mensaje_final_preview, height=300, disabled=True, key="final_preview_display")
                    
                    if nuevo_asunto_prev != asunto_ejemplo or nuevo_mensaje_prev != plantilla['mensaje']:
                        st.warning("⚠️ Has realizado cambios. Haz clic en 💾 para guardarlos en la base de datos.")
        
        else:  # Bienvenida
            if len(estudiantes_bienvenida) > 0:
                plantilla = obtener_plantilla(datos['institucion'], 'bienvenida')
                estudiante_ejemplo = estudiantes_bienvenida.iloc[0]
                nombre, apellidos = obtener_nombre_completo(estudiante_ejemplo, datos['institucion'])
                
                mensaje_ejemplo = plantilla['mensaje'].format(
                    nombre=nombre,
                    **variables_extra
                )
                asunto_ejemplo = plantilla['asunto'].format(**variables_extra)
                
                st.info(f"📌 Ejemplo para: **{nombre} {apellidos}**")
                
                col_asunto_bien, col_save_bien = st.columns([5, 1])
                
                with col_asunto_bien:
                    nuevo_asunto_bien = st.text_input(
                        "Asunto del mensaje:",
                        value=asunto_ejemplo,
                        key="asunto_bienvenida_edit"
                    )
                
                with col_save_bien:
                    st.write("")
                    st.write("")
                    if st.button("💾", key="guardar_bien", help="Guardar en BD"):
                        guardar_plantilla_db(
                            datos['institucion'],
                            'bienvenida',
                            plantilla['nombre'],
                            nuevo_asunto_bien,
                            st.session_state.get('mensaje_bienvenida_value', plantilla['mensaje'])
                        )
                        st.success("✅ Guardado")
                        time.sleep(1)
                        st.rerun()
                
                nuevo_mensaje_bien = st.text_area(
                    "Mensaje:",
                    value=plantilla['mensaje'],
                    height=400,
                    key="mensaje_bienvenida_edit"
                )
                
                st.session_state['mensaje_bienvenida_value'] = nuevo_mensaje_bien
                
                with st.expander("👁️ Vista previa final", expanded=False):
                    mensaje_final = nuevo_mensaje_bien.format(nombre=nombre, **variables_extra)
                    st.text_area("", value=mensaje_final, height=300, disabled=True, key="bien_final_preview")
                
                if nuevo_asunto_bien != asunto_ejemplo or nuevo_mensaje_bien != plantilla['mensaje']:
                    st.warning("⚠️ Cambios pendientes. Haz clic en 💾 para guardar.")
        
        st.divider()
        st.subheader("📧 Envío")
        
        if tipo_envio == "automatico":
            col1, col2, col3 = st.columns(3)
            
            with col1:
                if st.button("✓ Felicitaciones", disabled=len(estudiantes_completos)==0, key="felicit_tab1"):
                    for _, estudiante in estudiantes_completos.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        
                        actividades_completadas = obtener_actividades_completadas(
                            estudiante, 
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades']
                        )
                        
                        plantilla = obtener_plantilla(datos['institucion'], 'felicitacion')
                        asunto = plantilla['asunto'].format(**variables_extra)
                        
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_completadas)])
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_completadas=actividades_lista,
                            **variables_extra
                        )
                        
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            enviar_a_todos_los_emails(emails_validos, asunto, mensaje, nombre_completo, archivos_adjuntos)
                        
                        time.sleep(1)
                    st.success("✓ Enviados")
            
            with col2:
                if st.button("⚠ Recordatorios", disabled=len(estudiantes_incompletos)==0, key="recordat_tab1"):
                    for _, estudiante in estudiantes_incompletos.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        
                        # Para semana 3, mostrar TODAS las actividades pendientes
                        num_actividades_a_revisar = len(config_institucion['columnas_actividades']) if datos['semana'] == 3 else datos['actividades_requeridas']
                        
                        actividades_faltantes = obtener_actividades_faltantes(
                            estudiante,
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades'],
                            num_actividades_a_revisar
                        )
                        
                        plantilla = obtener_plantilla(datos['institucion'], 'seguimiento_atraso', datos['semana'])
                        asunto = plantilla['asunto'].format(**variables_extra)
                        
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_faltantes)])
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_faltantes=actividades_lista,
                            **variables_extra
                        )
                        
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            enviar_a_todos_los_emails(emails_validos, asunto, mensaje, nombre_completo, archivos_adjuntos)
                        
                        time.sleep(1)
                    st.success("✓ Enviados")
            
            with col3:
                if st.button("✗ Alertas", disabled=len(estudiantes_sin_entregas)==0, key="alertas_tab1"):
                    for _, estudiante in estudiantes_sin_entregas.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        
                        plantilla = obtener_plantilla(datos['institucion'], 'seguimiento_sin_acceso', datos['semana'])
                        asunto = plantilla['asunto'].format(**variables_extra)
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            **variables_extra
                        )
                        
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            enviar_a_todos_los_emails(emails_validos, asunto, mensaje, nombre_completo, archivos_adjuntos)
                        
                        time.sleep(1)
                    st.success("✓ Enviados")
            
            st.divider()
            total_estudiantes = len(estudiantes_completos) + len(estudiantes_incompletos) + len(estudiantes_sin_entregas)
            
            if total_estudiantes > 0:
                if total_estudiantes > 110:
                    st.warning(f"⚠ {total_estudiantes} correos. Pausas automáticas.")
                
                if st.button(f"🚀 Enviar Todos ({total_estudiantes})", type="primary", key="masivo_tab1"):
                    st.balloons()
                    
                    total_exitosos = 0
                    total_procesados = 0
                    contador_lote = 0
                    
                    for categoria, estudiantes_cat, tipo_plantilla in [
                        ("Felicitaciones", estudiantes_completos, "felicitacion"),
                        ("Recordatorios", estudiantes_incompletos, "seguimiento_atraso"),
                        ("Alertas", estudiantes_sin_entregas, "seguimiento_sin_acceso")
                    ]:
                        
                        if len(estudiantes_cat) > 0:
                            st.write(f"**{categoria}...**")
                            
                            for _, estudiante in estudiantes_cat.iterrows():
                                nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                                nombre_completo = f"{nombre} {apellidos}".strip()
                                
                                plantilla = obtener_plantilla(datos['institucion'], tipo_plantilla, datos['semana'])
                                asunto = plantilla['asunto'].format(**variables_extra)
                                
                                if tipo_plantilla == "felicitacion":
                                    actividades_completadas = obtener_actividades_completadas(
                                        estudiante, 
                                        config_institucion['columnas_actividades'],
                                        config_institucion['nombres_actividades']
                                    )
                                    actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_completadas)])
                                    mensaje = plantilla['mensaje'].format(
                                        nombre=nombre,
                                        actividades_completadas=actividades_lista,
                                        **variables_extra
                                    )
                                elif tipo_plantilla == "seguimiento_atraso":
                                    # Para semana 3, mostrar TODAS las actividades pendientes
                                    num_actividades_a_revisar = len(config_institucion['columnas_actividades']) if datos['semana'] == 3 else datos['actividades_requeridas']
                                    
                                    actividades_faltantes = obtener_actividades_faltantes(
                                        estudiante,
                                        config_institucion['columnas_actividades'],
                                        config_institucion['nombres_actividades'],
                                        num_actividades_a_revisar
                                    )
                                    actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_faltantes)])
                                    mensaje = plantilla['mensaje'].format(
                                        nombre=nombre,
                                        actividades_faltantes=actividades_lista,
                                        **variables_extra
                                    )
                                else:
                                    mensaje = plantilla['mensaje'].format(
                                        nombre=nombre,
                                        **variables_extra
                                    )
                                
                                emails_validos = obtener_emails_validos(estudiante)
                                if emails_validos:
                                    exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                                        emails_validos, asunto, mensaje, nombre_completo, archivos_adjuntos
                                    )
                                    total_exitosos += exitosos
                                
                                total_procesados += 1
                                contador_lote += 1
                                
                                if contador_lote % 50 == 0:
                                    st.info(f"Pausa... ({contador_lote})")
                                    time.sleep(10)
                                else:
                                    time.sleep(1)
                    
                    st.success(f"✓ Completado: {total_exitosos}/{total_procesados}")
        
        else:
            if st.button(f"🎓 Enviar Bienvenida ({len(estudiantes_bienvenida)})", type="primary", key="bienvenida_tab1"):
                
                total_exitosos = 0
                total_procesados = 0
                contador_lote = 0
                
                plantilla = obtener_plantilla(datos['institucion'], 'bienvenida')
                
                for _, estudiante in estudiantes_bienvenida.iterrows():
                    nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                    nombre_completo = f"{nombre} {apellidos}".strip()
                    
                    asunto = plantilla['asunto'].format(**variables_extra)
                    mensaje = plantilla['mensaje'].format(
                        nombre=nombre,
                        **variables_extra
                    )
                    
                    emails_validos = obtener_emails_validos(estudiante)
                    if emails_validos:
                        exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                            emails_validos, asunto, mensaje, nombre_completo, archivos_adjuntos
                        )
                        total_exitosos += exitosos
                    
                    total_procesados += 1
                    contador_lote += 1
                    
                    if contador_lote % 50 == 0:
                        st.info(f"Pausa... ({contador_lote})")
                        time.sleep(10)
                    else:
                        time.sleep(1)
                
                st.success(f"✓ Enviados: {total_exitosos}/{total_procesados}")

    historial = cargar_historial_db(100)
    if historial:
        st.divider()
        with st.expander(f"📋 Historial ({len(historial)})"):
            df_historial = pd.DataFrame(historial)
            
            exitosos = len(df_historial[df_historial['estado'] == 'Enviado'])
            fallidos = len(df_historial) - exitosos
            
            col1, col2, col3 = st.columns(3)
            col1.metric("Total", len(df_historial))
            col2.metric("✓", exitosos)
            col3.metric("✗", fallidos)
            
            if st.button("🗑️", key="limpiar_tab1"):
                limpiar_historial_db()
                st.rerun()
            
            st.dataframe(
                df_historial[['timestamp', 'destinatario', 'asunto', 'estado']],
                use_container_width=True,
                height=200
            )

# =====================================================
# TAB 2: SISTEMA DE PRÁCTICAS
# =====================================================

with tab2:
    st.header("📝 Sistema de Envío de Correos - Prácticas Profesionales")
    
    @st.cache_data
    def load_excel_data(file):
        """Cargar datos del archivo Excel"""
        try:
            df = pd.read_excel(file, sheet_name='Calificaciones')
            return df
        except Exception as e:
            st.error(f"Error al cargar el archivo: {str(e)}")
            return None

    PLANTILLAS_PRACTICAS = {
        "Bienvenida": {
            "asunto": "Bienvenida y primer reto - Estancias Profesionales UVEG",
            "contenido": """Apreciable {nombre_completo}:

Mi nombre es {nombre_asesor}, y seré tu asesor durante el desarrollo de las Estancias Profesionales en la {universidad}. Te doy la más cordial bienvenida al curso y aprovecho para compartirte información importante.

Primer Reto (Reto 1 - Carta de Autorización):  
He programado como fecha de entrega el {fecha_cierre_actividad} a medio día.  
Es importante que todos los datos solicitados estén correctos y completos, ya que la evaluación será binaria (100 o 0 puntos).  
Si no acreditas el Reto 1, no podrás continuar con los siguientes retos.

Periodo de prácticas profesionales:  
- Inicio: {fecha_inicio_practicas}  
- Término: {fecha_fin_practicas}

Importante: Revisa con atención la rúbrica del Reto 1. Es fundamental que cumplas todos los criterios exactamente como se indican.

También te pido que revises la sección de "Avisos" en tu aula virtual para mantenerte al tanto de cualquier novedad.

Quedo a tu disposición para cualquier duda. Recuerda que no estás solo/a, estoy aquí para ayudarte durante todo el proceso.

Te agradeceré que me confirmes la recepción de este correo.

Atentamente,  
{nombre_completo_asesor}  
Asesor Virtual  
{nombre_universidad}"""
        },
        
        "Sesión Síncrona": {
            "asunto": "Invitación - Sesión síncrona de las prácticas profesionales",
            "contenido": """Buen día, {nombre}:

El presente tiene el objetivo de invitarte a la: Sesión síncrona de las prácticas profesionales

{info_google_meet}

Te esperamos puntualmente.

Saludos cordiales,
{nombre_asesor}"""
        },
        
        "Envío de Grabación": {
            "asunto": "Grabación de sesión síncrona - Reto {numero_reto}",
            "contenido": """Buen día, {nombre}:

Envío la grabación de la sesión síncrona, correspondiente al reto {numero_reto}:

Enlace: {enlace_grabacion}

Agradeceré me contestes de recibido este mensaje.

Quedo al pendiente.

Saludos,
{nombre_asesor}"""
        },
        
        "Libre": {
            "asunto": "",
            "contenido": """Apreciable {nombre}:

[Escribe aquí tu mensaje personalizado]

Saludos cordiales,
{nombre_asesor}"""
        }
    }

    def enviar_correo_tab2(smtp_server, smtp_port, email_usuario, email_password, destinatario, asunto, contenido, archivos_adjuntos=None, reintentos=3):
        """Función para enviar correo electrónico con reintentos automáticos"""
        for intento in range(reintentos):
            server = None
            try:
                msg = MIMEMultipart()
                msg['From'] = email_usuario
                msg['To'] = destinatario
                msg['Subject'] = asunto
                
                msg.attach(MIMEText(contenido, 'plain', 'utf-8'))
                
                if archivos_adjuntos:
                    for archivo in archivos_adjuntos:
                        # Reiniciar posición del archivo para cada intento
                        archivo.seek(0)
                        part = MIMEBase('application', 'octet-stream')
                        part.set_payload(archivo.read())
                        encoders.encode_base64(part)
                        part.add_header(
                            'Content-Disposition',
                            f'attachment; filename= {archivo.name}'
                        )
                        msg.attach(part)
                
                # Crear nueva conexión con timeout más largo
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                server.starttls()
                server.login(email_usuario, email_password)
                
                text = msg.as_string()
                server.sendmail(email_usuario, destinatario, text)
                server.quit()
                
                guardar_historial_db(asunto, destinatario, 'Enviado')
                return True, "Correo enviado exitosamente"
            
            except Exception as e:
                # Cerrar servidor si está abierto
                if server:
                    try:
                        server.quit()
                    except:
                        pass
                
                if intento < reintentos - 1:
                    # Esperar antes de reintentar (tiempo progresivo)
                    tiempo_espera = 3 * (intento + 1)
                    time.sleep(tiempo_espera)
                    continue
                else:
                    error_msg = f"Error al enviar correo (después de {reintentos} intentos): {str(e)}"
                    guardar_historial_db(asunto, destinatario, error_msg)
                    return False, error_msg
        
        return False, "Error desconocido"

    with st.sidebar:
        st.header("⚙️ Configuración de Correo")

        smtp_server = st.text_input("Servidor SMTP", value="smtp.gmail.com")
        smtp_port = st.number_input("Puerto SMTP", value=587, min_value=1, max_value=65535)
        email_usuario = st.text_input("Email de usuario", value=REMITENTE)
        email_password = st.text_input("Contraseña", type="password", value=CLAVE_APP,
                                      help="Para Gmail, usa una contraseña de aplicación")

        st.header("👨‍🏫 Información del Asesor")
        nombre_asesor = st.text_input("Nombre del Asesor", value="Juan Manuel Martinez Zaragoza")
        nombre_completo_asesor = st.text_input("Nombre Completo del Asesor", value="Juan Manuel Martinez Zaragoza")
        universidad = st.text_input("Universidad", value="UVEG")
        nombre_universidad = st.text_input("Nombre Completo Universidad", value="Universidad Virtual del Estado de Guanajuato")

    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("📁 Cargar Archivo Excel")
        uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'], key="upload_tab2")
        
        if uploaded_file is not None:
            df = load_excel_data(uploaded_file)
            if df is not None:
                st.success(f"Archivo cargado: {len(df)} registros encontrados")
                
                st.subheader("Vista previa de datos:")
                st.dataframe(df[['Nombre', 'Apellido(s)', 'Correo Personal', 'Dirección Email']].head())
                
                st.subheader("Seleccionar destinatarios:")
                enviar_a = st.selectbox("Enviar correos a:", 
                                       ["Correo Personal", "Dirección Email", "Ambos"], key="enviar_a_tab2")

    with col2:
        st.subheader("✉️ Composición del Correo")
        
        plantilla_seleccionada = st.selectbox("Selecciona una plantilla:", 
                                            ["Bienvenida", "Sesión Síncrona", "Envío de Grabación", "Libre"], key="plantilla_tab2")
        
        if plantilla_seleccionada == "Bienvenida":
            col_fecha1, col_fecha2 = st.columns(2)
            with col_fecha1:
                fecha_cierre_actividad = st.date_input("Fecha de cierre de actividad", key="fecha_cierre_tab2")
                fecha_inicio_practicas = st.date_input("Fecha de inicio de prácticas", key="fecha_inicio_tab2")
            with col_fecha2:
                fecha_fin_practicas = st.date_input("Fecha de fin de prácticas", key="fecha_fin_tab2")
        
        elif plantilla_seleccionada == "Sesión Síncrona":
            st.write("**Información de Google Meet:**")
            info_google_meet = st.text_area(
                "Pega aquí la información completa de Google Meet:",
                value="""Martes, 6 de mayo · 8:30 – 9:30pm Zona horaria: America/Mexico_City 
Información para unirse con Google Meet 
Enlace de la videollamada: https://meet.google.com/ibb-zfcq-fps""",
                height=100,
                key="google_meet_tab2"
            )
        
        elif plantilla_seleccionada == "Envío de Grabación":
            numero_reto = st.text_input("Número de reto", value="1", key="numero_reto_tab2")
            enlace_grabacion = st.text_input("Enlace de grabación", 
                                           value="https://drive.google.com/file/d/1xPRTqb_hE0eXi6VOCA2Kg9ydSMj1qn8k/view?usp=sharing",
                                           key="enlace_grab_tab2")
        
        asunto_default = PLANTILLAS_PRACTICAS[plantilla_seleccionada]["asunto"]
        asunto = st.text_input("Asunto del correo:", value=asunto_default, key="asunto_tab2")
        
        contenido_default = PLANTILLAS_PRACTICAS[plantilla_seleccionada]["contenido"]
        contenido = st.text_area("Contenido del correo:", value=contenido_default, height=300, key="contenido_tab2")
        
        st.subheader("📎 Archivos Adjuntos")
        archivos_adjuntos = st.file_uploader("Selecciona archivos para adjuntar", 
                                           accept_multiple_files=True, key="archivos_tab2")
        
        if archivos_adjuntos:
            st.write("Archivos seleccionados:")
            for archivo in archivos_adjuntos:
                st.write(f"- {archivo.name} ({archivo.size} bytes)")

    if uploaded_file is not None and df is not None:
        st.divider()
        st.header("👁️ Vista Previa del Correo")
        
        indice_preview = st.selectbox("Selecciona un destinatario para vista previa:", 
                                    range(len(df)), 
                                    format_func=lambda x: f"{df.iloc[x]['Nombre']} {df.iloc[x]['Apellido(s)']}",
                                    key="preview_tab2")
        
        destinatario = df.iloc[indice_preview]
        nombre_completo = f"{destinatario['Nombre']} {destinatario['Apellido(s)']}"
        
        contenido_personalizado = contenido.format(
            nombre=destinatario['Nombre'],
            nombre_completo=nombre_completo,
            nombre_asesor=nombre_asesor,
            nombre_completo_asesor=nombre_completo_asesor,
            universidad=universidad,
            nombre_universidad=nombre_universidad,
            fecha_cierre_actividad=fecha_cierre_actividad if 'fecha_cierre_actividad' in locals() else "fecha_cierre_actividad",
            fecha_inicio_practicas=fecha_inicio_practicas if 'fecha_inicio_practicas' in locals() else "fecha_inicio_practicas",
            fecha_fin_practicas=fecha_fin_practicas if 'fecha_fin_practicas' in locals() else "fecha_fin_practicas",
            info_google_meet=info_google_meet if 'info_google_meet' in locals() else "info_google_meet",
            numero_reto=numero_reto if 'numero_reto' in locals() else "numero_reto",
            enlace_grabacion=enlace_grabacion if 'enlace_grabacion' in locals() else "enlace_grabacion"
        )
        
        st.subheader(f"📧 Para: {nombre_completo}")
        st.write(f"**Asunto:** {asunto}")
        st.write("**Contenido:**")
        st.text_area("", value=contenido_personalizado, height=200, disabled=True, key="preview_content_tab2")

    if uploaded_file is not None and df is not None:
        st.divider()
        
        col_envio1, col_envio2, col_envio3 = st.columns([1, 1, 1])
        
        with col_envio2:
            if st.button("🚀 Enviar Correos", type="primary", use_container_width=True, key="enviar_tab2"):
                if not email_usuario or not email_password:
                    st.error("Por favor, configura tu email y contraseña en la barra lateral")
                else:
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    enviados = 0
                    errores = 0
                    contador_lote = 0
                    
                    # Información sobre pausas automáticas
                    if len(df) > 50:
                        st.info(f"⏸️ Se realizarán pausas automáticas cada 50 correos para evitar bloqueos. Total a enviar: {len(df)}")
                    
                    for i, row in df.iterrows():
                        progress = (i + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Enviando correo {i+1} de {len(df)} | ✅ Enviados: {enviados} | ❌ Errores: {errores}")
                        
                        nombre_completo = f"{row['Nombre']} {row['Apellido(s)']}"
                        
                        contenido_personalizado = contenido.format(
                            nombre=row['Nombre'],
                            nombre_completo=nombre_completo,
                            nombre_asesor=nombre_asesor,
                            nombre_completo_asesor=nombre_completo_asesor,
                            universidad=universidad,
                            nombre_universidad=nombre_universidad,
                            fecha_cierre_actividad=fecha_cierre_actividad if 'fecha_cierre_actividad' in locals() else "fecha_cierre_actividad",
                            fecha_inicio_practicas=fecha_inicio_practicas if 'fecha_inicio_practicas' in locals() else "fecha_inicio_practicas",
                            fecha_fin_practicas=fecha_fin_practicas if 'fecha_fin_practicas' in locals() else "fecha_fin_practicas",
                            info_google_meet=info_google_meet if 'info_google_meet' in locals() else "info_google_meet",
                            numero_reto=numero_reto if 'numero_reto' in locals() else "numero_reto",
                            enlace_grabacion=enlace_grabacion if 'enlace_grabacion' in locals() else "enlace_grabacion"
                        )
                        
                        destinatarios = []
                        if enviar_a == "Correo Personal":
                            destinatarios.append(row['Correo Personal'])
                        elif enviar_a == "Dirección Email":
                            destinatarios.append(row['Dirección Email'])
                        else:
                            destinatarios.extend([row['Correo Personal'], row['Dirección Email']])
                        
                        for destinatario in destinatarios:
                            if pd.notna(destinatario) and destinatario.strip():
                                exito, mensaje = enviar_correo_tab2(
                                    smtp_server, smtp_port, email_usuario, email_password,
                                    destinatario, asunto, contenido_personalizado, archivos_adjuntos
                                )
                                
                                if exito:
                                    enviados += 1
                                else:
                                    errores += 1
                                    st.error(f"Error enviando a {destinatario}: {mensaje}")
                                
                                # Pausa entre correos individuales
                                time.sleep(2)
                        
                        contador_lote += 1
                        
                        # Pausa cada 50 correos para evitar bloqueos
                        if contador_lote % 50 == 0 and contador_lote < len(df):
                            st.warning(f"⏸️ Pausa preventiva después de {contador_lote} correos procesados...")
                            time.sleep(15)  # Pausa de 15 segundos cada 50 correos
                            st.info("▶️ Continuando envío...")
                    
                    progress_bar.progress(1.0)
                    status_text.text("¡Envío completado!")
                    
                    col_result1, col_result2 = st.columns(2)
                    with col_result1:
                        st.success(f"✅ Correos enviados exitosamente: {enviados}")
                    with col_result2:
                        if errores > 0:
                            st.error(f"❌ Errores encontrados: {errores}")
                        else:
                            st.success("🎉 ¡Todos los correos se enviaron sin errores!")
                    
                    st.balloons()

# =====================================================
# TAB 3: SISTEMA BIENVENIDA NOVAUNIVERSITAS
# =====================================================

with tab3:
    st.header("🎓 Sistema de Bienvenida NovaUniversitas")
    
    def generar_mensaje_personalizado(nombre, correo_institucional, contrasena):
        """Genera el mensaje personalizado para cada estudiante"""
        mensaje = f"""Apreciable {nombre},

¡Te damos la más cordial bienvenida a **NovaUniversitas**! Nos complace que formes parte de nuestra comunidad académica en esta nueva etapa de aprendizaje.

Para acceder a tu **cuenta de correo institucional**:

**Correo institucional:** {correo_institucional}
**Contraseña temporal:** {contrasena}

**Si estás recursando, mantienes tu contraseña que has modificado.**

Te recomendamos acceder cuanto antes y cambiar tu contraseña por una más segura.

**IMPORTANTE**: Una vez dentro de tu correo, busca el mensaje que contiene el usuario y contraseña de la plataforma virtual:
https://virtual.novauniversitas.edu.mx

**Tutorial para inicio de sesiones:**
https://youtu.be/75ib7aN0Tvw?feature=shared

Desde **Coordinación Virtual**, estamos aquí para apoyarte en todo este proceso. Si tienes dudas o necesitas ayuda con el acceso, no dudes en contactarnos en mesadeayuda@virtual.novauniversitas.edu.mx

Te dejo un enlace a los lineamientos que debes seguir con el uso del correo institucional.

**Por favor, confirme de recibido.**

¡Te deseamos mucho éxito en esta nueva etapa!

Atentamente,
**M.D. Juan Manuel Martínez Zaragoza**
Coordinación Virtual
**www.virtual.novauniversitas.edu.mx**"""
        
        return mensaje

    def enviar_correo_tab3(smtp_server, smtp_port, email_usuario, email_password, 
                      destinatario, asunto, mensaje, archivos_adjuntos=None, reintentos=3):
        """Envía un correo electrónico usando SMTP con archivos adjuntos opcionales y reintentos automáticos"""
        for intento in range(reintentos):
            server = None
            try:
                msg = MIMEMultipart()
                msg['From'] = email_usuario
                msg['To'] = destinatario
                msg['Subject'] = asunto
                
                msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))
                
                if archivos_adjuntos:
                    for archivo in archivos_adjuntos:
                        if archivo is not None:
                            try:
                                # Reiniciar posición del archivo para cada intento
                                archivo.seek(0)
                                part = MIMEBase('application', 'octet-stream')
                                part.set_payload(archivo.getvalue())
                                encoders.encode_base64(part)
                                part.add_header(
                                    'Content-Disposition',
                                    f'attachment; filename= {archivo.name}'
                                )
                                msg.attach(part)
                            except Exception as e:
                                if intento == reintentos - 1:
                                    return False, f"Error al adjuntar archivo {archivo.name}: {str(e)}"
                
                # Crear nueva conexión con timeout más largo
                server = smtplib.SMTP(smtp_server, smtp_port, timeout=30)
                server.starttls()
                server.login(email_usuario, email_password)
                
                text = msg.as_string()
                server.sendmail(email_usuario, destinatario, text)
                server.quit()
                
                guardar_historial_db(asunto, destinatario, 'Enviado')
                return True, "Correo enviado exitosamente"
            
            except Exception as e:
                # Cerrar servidor si está abierto
                if server:
                    try:
                        server.quit()
                    except:
                        pass
                
                if intento < reintentos - 1:
                    # Esperar antes de reintentar (tiempo progresivo)
                    tiempo_espera = 3 * (intento + 1)
                    time.sleep(tiempo_espera)
                    continue
                else:
                    error_msg = f"Error al enviar correo (después de {reintentos} intentos): {str(e)}"
                    guardar_historial_db(asunto, destinatario, error_msg)
                    return False, error_msg
        
        return False, "Error desconocido"

    with st.sidebar:
        st.header("⚙️ Configuración de Correo")

        smtp_server_tab3 = st.text_input("Servidor SMTP", value="smtp.gmail.com", key="smtp_server_tab3")
        smtp_port_tab3 = st.number_input("Puerto SMTP", value=587, min_value=1, max_value=65535, key="smtp_port_tab3")
        email_usuario_tab3 = st.text_input("Correo del remitente", value=REMITENTE, key="email_usuario_tab3")
        email_password_tab3 = st.text_input("Contraseña del remitente", type="password", value=CLAVE_APP,
                                          help="Para Gmail, usa una contraseña de aplicación", key="email_password_tab3")

        st.divider()
        st.markdown("### 📋 Instrucciones")
        st.markdown("""
        1. **Configura** los parámetros SMTP
        2. **Carga** el archivo Excel con los datos
        3. **Selecciona** la hoja correspondiente
        4. **Prueba** el envío con un estudiante
        5. **Ejecuta** el envío masivo
        """)

    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📁 Cargar Archivo Excel")
        
        uploaded_file_tab3 = st.file_uploader(
            "Selecciona el archivo Excel con los datos de estudiantes",
            type=['xlsx', 'xls'],
            help="El archivo debe contener las columnas: NP, Grupo, Matrícula, Nombre, Email_personal, Correo Institucional, Contraseña, Cuatrimestre, Carrera",
            key="uploaded_file_tab3"
        )
        
        if uploaded_file_tab3 is not None:
            try:
                excel_file = pd.ExcelFile(uploaded_file_tab3)
                sheet_names = excel_file.sheet_names
                
                selected_sheet = st.selectbox(
                    "Selecciona la hoja a procesar:",
                    sheet_names,
                    help="Elige la hoja que contiene los datos de estudiantes",
                    key="selected_sheet_tab3"
                )
                
                df_tab3 = pd.read_excel(uploaded_file_tab3, sheet_name=selected_sheet)
                
                required_columns = ['Nombre', 'Email_personal', 'Correo Institucional', 'Contraseña']
                missing_columns = [col for col in required_columns if col not in df_tab3.columns]
                
                if missing_columns:
                    st.error(f"❌ Faltan las siguientes columnas: {', '.join(missing_columns)}")
                else:
                    df_tab3 = df_tab3.dropna(subset=['Nombre', 'Email_personal', 'Correo Institucional'])
                    
                    st.success(f"✅ Archivo cargado exitosamente: {len(df_tab3)} estudiantes encontrados")
                    
                    st.subheader("👀 Vista Previa de Datos")
                    st.dataframe(df_tab3[['Nombre', 'Email_personal', 'Correo Institucional', 'Contraseña']].head(10))
                    
                    st.info(f"📊 **Estadísticas:**\n- Total de estudiantes: {len(df_tab3)}\n- Columnas disponibles: {len(df_tab3.columns)}")
                    
            except Exception as e:
                st.error(f"❌ Error al leer el archivo: {str(e)}")

    with col2:
        st.subheader("📧 Configuración de Envío")
        
        if uploaded_file_tab3 is not None and 'df_tab3' in locals() and not df_tab3.empty:
            
            asunto_tab3 = st.text_input("Asunto del correo", 
                                  value="Bienvenida a NovaUniversitas - Credenciales de Acceso",
                                  key="asunto_tab3")
            
            st.subheader("📎 Archivos Adjuntos")
            
            num_archivos = st.selectbox(
                "¿Cuántos archivos deseas adjuntar?",
                options=[0, 1, 2, 3, 4, 5],
                help="Selecciona la cantidad de archivos que deseas adjuntar a todos los correos",
                key="num_archivos_tab3"
            )
            
            archivos_adjuntos_tab3 = []
            if num_archivos > 0:
                st.info(f"📁 Selecciona {num_archivos} archivo(s) para adjuntar:")
                
                for i in range(num_archivos):
                    archivo = st.file_uploader(
                        f"Archivo {i+1}:",
                        key=f"archivo_tab3_{i}",
                        help="Formatos soportados: PDF, DOC, DOCX, XLS, XLSX, TXT, JPG, PNG, etc."
                    )
                    if archivo is not None:
                        archivos_adjuntos_tab3.append(archivo)
                        st.success(f"✅ {archivo.name} ({archivo.size / 1024:.1f} KB)")
            
            if archivos_adjuntos_tab3:
                st.markdown("### 📋 Resumen de Archivos Adjuntos:")
                total_size = sum(archivo.size for archivo in archivos_adjuntos_tab3)
                st.write(f"**Total de archivos:** {len(archivos_adjuntos_tab3)}")
                st.write(f"**Tamaño total:** {total_size / 1024:.1f} KB")
                
                for i, archivo in enumerate(archivos_adjuntos_tab3, 1):
                    st.write(f"{i}. {archivo.name} ({archivo.size / 1024:.1f} KB)")
                
                if total_size > 10 * 1024 * 1024:
                    st.warning("⚠️ El tamaño total de archivos es mayor a 10 MB. Algunos servidores de correo pueden rechazar el mensaje.")
            
            envio_option = st.radio(
                "Selecciona el tipo de envío:",
                ["Vista previa del mensaje", "Envío de prueba", "Envío masivo"],
                key="envio_option_tab3"
            )
            
            if envio_option == "Vista previa del mensaje":
                st.subheader("📄 Vista Previa del Mensaje")
                if not df_tab3.empty:
                    nombre_ejemplo = df_tab3.iloc[0]['Nombre']
                    correo_ejemplo = df_tab3.iloc[0]['Correo Institucional']
                    contrasena_ejemplo = df_tab3.iloc[0]['Contraseña'] if 'Contraseña' in df_tab3.columns else "0125070109"
                    
                    mensaje_ejemplo = generar_mensaje_personalizado(nombre_ejemplo, correo_ejemplo, contrasena_ejemplo)
                    
                    st.text_area("Mensaje que se enviará:", mensaje_ejemplo, height=400, key="mensaje_preview_tab3")
                    st.info(f"📌 Ejemplo generado para: **{nombre_ejemplo}**")
                    
                    if archivos_adjuntos_tab3:
                        st.markdown("### 📎 Archivos que se adjuntarán:")
                        for archivo in archivos_adjuntos_tab3:
                            st.write(f"• {archivo.name}")
            
            elif envio_option == "Envío de prueba":
                st.subheader("🧪 Envío de Prueba")
                
                estudiante_prueba = st.selectbox(
                    "Selecciona un estudiante para prueba:",
                    df_tab3['Nombre'].tolist(),
                    key="estudiante_prueba_tab3"
                )
                
                if st.button("📤 Enviar Correo de Prueba", type="primary", key="enviar_prueba_tab3"):
                    if not all([smtp_server_tab3, smtp_port_tab3, email_usuario_tab3, email_password_tab3]):
                        st.error("❌ Por favor, completa toda la configuración SMTP")
                    else:
                        estudiante_data = df_tab3[df_tab3['Nombre'] == estudiante_prueba].iloc[0]
                        
                        mensaje = generar_mensaje_personalizado(
                            estudiante_data['Nombre'],
                            estudiante_data['Correo Institucional'],
                            estudiante_data['Contraseña'] if 'Contraseña' in df_tab3.columns else "0125070109"
                        )
                        
                        with st.spinner("Enviando correo de prueba..."):
                            exito, resultado = enviar_correo_tab3(
                                smtp_server_tab3, smtp_port_tab3, email_usuario_tab3, email_password_tab3,
                                estudiante_data['Email_personal'], asunto_tab3, mensaje, archivos_adjuntos_tab3
                            )
                        
                        if exito:
                            st.success(f"✅ {resultado}")
                            st.balloons()
                        else:
                            st.error(f"❌ {resultado}")
            
            elif envio_option == "Envío masivo":
                st.subheader("📬 Envío Masivo")
                
                st.warning("⚠️ **Advertencia**: Esta acción enviará correos a todos los estudiantes en el archivo.")
                
                if archivos_adjuntos_tab3:
                    st.info(f"📎 Se adjuntarán {len(archivos_adjuntos_tab3)} archivo(s) a cada correo")
                
                delay_between_emails = st.slider(
                    "Retraso entre correos (segundos):",
                    min_value=1, max_value=10, value=2,
                    help="Retraso para evitar ser marcado como spam",
                    key="delay_tab3"
                )
                
                if st.button("📤 Iniciar Envío Masivo", type="primary", key="enviar_masivo_tab3"):
                    if not all([smtp_server_tab3, smtp_port_tab3, email_usuario_tab3, email_password_tab3]):
                        st.error("❌ Por favor, completa toda la configuración SMTP")
                    else:
                        enviados = 0
                        errores = 0
                        contador_lote = 0
                        
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        log_container = st.empty()
                        logs = []
                        
                        # Información sobre pausas automáticas
                        if len(df_tab3) > 50:
                            st.info(f"⏸️ Se realizarán pausas automáticas cada 50 correos para evitar bloqueos. Total a enviar: {len(df_tab3)}")
                        
                        for index, row in df_tab3.iterrows():
                            try:
                                mensaje = generar_mensaje_personalizado(
                                    row['Nombre'],
                                    row['Correo Institucional'],
                                    row['Contraseña'] if 'Contraseña' in df_tab3.columns else "0125070109"
                                )
                                
                                exito, resultado = enviar_correo_tab3(
                                    smtp_server_tab3, smtp_port_tab3, email_usuario_tab3, email_password_tab3,
                                    row['Email_personal'], asunto_tab3, mensaje, archivos_adjuntos_tab3
                                )
                                
                                if exito:
                                    enviados += 1
                                    logs.append(f"✅ {row['Nombre']} - {resultado}")
                                else:
                                    errores += 1
                                    logs.append(f"❌ {row['Nombre']} - {resultado}")
                                
                                progress = (index + 1) / len(df_tab3)
                                progress_bar.progress(progress)
                                status_text.text(f"Procesando: {index + 1}/{len(df_tab3)} | ✅ Enviados: {enviados} | ❌ Errores: {errores}")
                                
                                with log_container.container():
                                    st.text_area("Registro de envíos:", "\n".join(logs[-10:]), height=200, key=f"log_area_tab3_{index}")
                                
                                contador_lote += 1
                                
                                # Pausa cada 50 correos para evitar bloqueos
                                if contador_lote % 50 == 0 and contador_lote < len(df_tab3):
                                    st.warning(f"⏸️ Pausa preventiva después de {contador_lote} correos procesados...")
                                    time.sleep(15)  # Pausa de 15 segundos cada 50 correos
                                    st.info("▶️ Continuando envío...")
                                    time.sleep(delay_between_emails)
                                else:
                                    time.sleep(delay_between_emails)
                                
                            except Exception as e:
                                errores += 1
                                logs.append(f"❌ {row['Nombre']} - Error: {str(e)}")
                        
                        st.success(f"🎉 **Envío completado!**\n- Enviados: {enviados}\n- Errores: {errores}")
                        
                        if enviados > 0:
                            st.balloons()
                        
                        log_text = "\n".join(logs)
                        st.download_button(
                            label="📥 Descargar Log Completo",
                            data=log_text,
                            file_name="log_envio_correos.txt",
                            mime="text/plain",
                            key="download_log_tab3"
                        )

st.divider()
st.caption("Sistema de Correos v4.0 Dark | Base de datos persistente")
