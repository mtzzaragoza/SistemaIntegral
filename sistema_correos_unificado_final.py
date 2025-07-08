import streamlit as st
import pandas as pd
import smtplib
from email.message import EmailMessage
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email import encoders
import mimetypes
import re
from datetime import datetime, timedelta
import time

# Configuración de la página
st.set_page_config(
    page_title="Sistema Integral de Correos Académicos", 
    page_icon="📧",
    layout="wide"
)

# Título principal
st.title("📧 Sistema Integral de Correos Académicos")
st.markdown("**UVEG & NovaUniversitas - Plataforma Unificada**")

# Crear pestañas principales
tab1, tab2, tab3 = st.tabs([
    "🎯 Sistema Inteligente (UVEG/Nova)", 
    "📧 Sistema de Envío Masivo (Prácticas)",
    "🎓 Sistema Bienvenida (NovaUniversitas)"
])

# =====================================================
# TAB 1: SISTEMA INTELIGENTE (UVEG & NOVAUNIVERSITAS)
# =====================================================

with tab1:
    st.header("🎯 Sistema Inteligente de Seguimiento Académico")
    st.markdown("*Análisis automático y envío personalizado por progreso académico*")
    
    # Configuración del remitente (se obtiene dinámicamente)
    def obtener_credenciales():
        """Función para obtener credenciales de forma segura"""
        st.subheader("🔐 Configuración de Credenciales de Correo")
        
        with st.form("credenciales_form"):
            col1, col2 = st.columns(2)
            
            with col1:
                remitente = st.text_input(
                    "📧 Correo electrónico:",
                    placeholder="tu_correo@gmail.com",
                    help="Debe ser una cuenta de Gmail con 2FA activado"
                )
            
            with col2:
                clave_app = st.text_input(
                    "🔑 Contraseña de aplicación:",
                    type="password",
                    placeholder="abcd efgh ijkl mnop",
                    help="Contraseña de 16 caracteres generada por Gmail"
                )
            
            # Información de ayuda
            with st.expander("ℹ️ ¿Cómo generar una contraseña de aplicación?"):
                st.markdown("""
                ### Pasos para generar contraseña de aplicación:
                
                1. **Activar 2FA**: Ve a [myaccount.google.com](https://myaccount.google.com) → Seguridad → Verificación en 2 pasos
                2. **Generar contraseña**: En la misma sección → Contraseñas de aplicación
                3. **Seleccionar**: App = "Correo", Dispositivo = "Otro (personalizado)"
                4. **Nombre**: "Sistema Correos Académicos"
                5. **Copiar**: La contraseña de 16 caracteres que aparece
                
                ⚠️ **Importante**: Nunca uses tu contraseña normal de Gmail
                """)
            
            submitted = st.form_submit_button("🔓 Configurar Credenciales", type="primary")
            
            if submitted:
                if remitente and clave_app:
                    # Validar formato de email
                    if "@" in remitente and "." in remitente:
                        # Validar que sea Gmail (opcional, puedes quitar esta validación)
                        if "gmail.com" in remitente or "gmail" in remitente.lower():
                            return remitente, clave_app
                        else:
                            st.warning("⚠️ Este sistema está optimizado para Gmail. ¿Estás seguro que es correcto?")
                            return remitente, clave_app
                    else:
                        st.error("❌ Formato de email inválido")
                        return None, None
                else:
                    st.error("❌ Por favor completa ambos campos")
                    return None, None
        
        return None, None

    # Obtener credenciales al inicio
    if 'credenciales_configuradas_tab1' not in st.session_state:
        st.session_state.credenciales_configuradas_tab1 = False
        st.session_state.remitente_tab1 = None
        st.session_state.clave_app_tab1 = None

    # Mostrar formulario de credenciales si no están configuradas
    if not st.session_state.credenciales_configuradas_tab1:
        remitente, clave_app = obtener_credenciales()
        
        if remitente and clave_app:
            # Probar las credenciales
            with st.spinner("🧪 Verificando credenciales..."):
                try:
                    # Intento de conexión para validar
                    server = smtplib.SMTP_SSL("smtp.gmail.com", 465, timeout=10)
                    server.login(remitente, clave_app)
                    server.quit()
                    
                    # Si llegamos aquí, las credenciales son correctas
                    st.session_state.credenciales_configuradas_tab1 = True
                    st.session_state.remitente_tab1 = remitente
                    st.session_state.clave_app_tab1 = clave_app
                    st.success("✅ ¡Credenciales verificadas correctamente!")
                    st.balloons()
                    time.sleep(1)
                    st.rerun()
                    
                except Exception as e:
                    st.error(f"❌ Error de autenticación: {str(e)}")
                    st.markdown("""
                    **Posibles soluciones:**
                    - Verifica que tengas 2FA activado en Gmail
                    - Asegúrate de usar una contraseña de aplicación (no tu contraseña normal)
                    - Revisa que no haya espacios extra en la contraseña
                    - Intenta generar una nueva contraseña de aplicación
                    """)
        
        # Detener la ejecución hasta que se configuren las credenciales
        st.stop()

    else:
        # Mostrar credenciales configuradas y opción para cambiar
        col1, col2 = st.columns([3, 1])
        
        with col1:
            st.success(f"✅ **Credenciales configuradas** - Remitente: {st.session_state.remitente_tab1}")
        
        with col2:
            if st.button("🔄 Cambiar Credenciales", key="cambiar_cred_tab1"):
                st.session_state.credenciales_configuradas_tab1 = False
                st.session_state.remitente_tab1 = None
                st.session_state.clave_app_tab1 = None
                st.rerun()

    # Usar las credenciales guardadas
    REMITENTE = st.session_state.remitente_tab1
    CLAVE_APP = st.session_state.clave_app_tab1

    # Inicializar estado de sesión para tab1
    if 'historial_envios_tab1' not in st.session_state:
        st.session_state.historial_envios_tab1 = []
    if 'analisis_generado_tab1' not in st.session_state:
        st.session_state.analisis_generado_tab1 = False
    if 'datos_estudiantes_tab1' not in st.session_state:
        st.session_state.datos_estudiantes_tab1 = None
    if 'plantillas_editadas_tab1' not in st.session_state:
        st.session_state.plantillas_editadas_tab1 = {}

    # Configuraciones por institución
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

    # Plantillas base por institución
    PLANTILLAS_BASE = {
        "uveg": {
            "bienvenida": {
                "nombre": "🎓 Bienvenida al Módulo - UVEG",
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
            
            "seguimiento_sin_acceso": {
                "nombre": "⚠️ Seguimiento - Sin Acceso - UVEG",
                "asunto": "Seguimiento {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.
					
Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" de la UVEG, en este momento estamos iniciando la {semana} semana del módulo, solo hemos avanzado algunos retos, con lo cual no presentas un gran atraso, te invito a que inicies tus actividades dentro de la plataforma (https://campus.uveg.edu.mx) y si tienes alguna duda con toda confianza me puedes contactar ya sea por este medio, por el mensajero de la plataforma o xxx por whatsapp.

Agradecería pudieras comentarme la situación por la cual no has accedido al módulo, espero todo se encuentre bien.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            
            "seguimiento_atraso": {
                "nombre": "📋 Seguimiento - Tareas Pendientes - UVEG",
                "asunto": "Seguimiento de avance {modulo} - Semana {semana} - UVEG",
                "mensaje": """Buen día {nombre}.
					
Mi nombre es Juan Manuel y soy tu asesor virtual en el módulo de "{modulo}" en la UVEG, en este momento estamos iniciando la {semana} semana del módulo. Me pongo en contacto, para saber si tienes algún inconveniente que no te este permitiendo avanzar al ritmo que les he marcado, ya que veo que tienes algunos retos pendientes y en tu avance semanal se ve reflejado.

Retos pendientes:
{actividades_faltantes}

Recuerda que esta es solo una sugerencia, si consideras que no tienes inconveniente en avanzar a tu ritmo, haz caso omiso a este y los siguientes mensajes.

Dr. Juan Manuel Martínez Zaragoza"""
            },
            
            "felicitacion": {
                "nombre": "🎉 Felicitación por Desempeño - UVEG",
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
                "nombre": "🎓 Bienvenida al Módulo - NovaUniversitas",
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
                "nombre": "⚠️ Seguimiento Académico - Sin Acceso - NovaUniversitas",
                "asunto": "Seguimiento académico {modulo} - Semana {semana} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Mi nombre es Juan Manuel, tu asesor virtual del módulo "{modulo}" en NovaUniversitas. Me dirijo a ti en relación a tu progreso académico en la {semana} semana del módulo.

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
                "nombre": "📋 Seguimiento Académico - Desafíos Pendientes - NovaUniversitas",
                "asunto": "Seguimiento de progreso {modulo} - Semana {semana} - NovaUniversitas",
                "mensaje": """Apreciable {nombre},

Mi nombre es Juan Manuel, tu asesor virtual del módulo "{modulo}" en NovaUniversitas. Me pongo en contacto contigo para hacer un seguimiento de tu progreso académico en la {semana} semana del módulo.

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
                "nombre": "🎉 Reconocimiento Académico - NovaUniversitas",
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

    # Funciones auxiliares para Tab 1
    def convertir_a_numerico(valor):
        """Convertir valor a numérico, manejando diferentes tipos de datos"""
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
        """Contar cuántas actividades ha completado un estudiante"""
        completadas = 0
        for col in columnas_actividades:
            if col in row.index:
                valor_numerico = convertir_a_numerico(row[col])
                if valor_numerico > 0:
                    completadas += 1
        return completadas

    def obtener_actividades_completadas(row, columnas_actividades, nombres_actividades):
        """Obtener lista de actividades completadas por un estudiante"""
        actividades_completadas = []
        for i, col in enumerate(columnas_actividades):
            if col in row.index:
                valor_numerico = convertir_a_numerico(row[col])
                if valor_numerico > 0:
                    actividades_completadas.append(nombres_actividades[i])
        return actividades_completadas

    def obtener_actividades_faltantes(row, columnas_actividades, nombres_actividades, actividades_requeridas):
        """Obtener lista de actividades faltantes según la semana"""
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
        """Validar que el email tenga formato correcto"""
        if not email or pd.isna(email):
            return False
        email_str = str(email).strip()
        return "@" in email_str and "." in email_str.split("@")[-1]

    def obtener_emails_validos(estudiante):
        """Obtener lista de emails válidos de un estudiante"""
        emails_validos = []
        
        # Verificar correo personal
        if 'Correo Personal' in estudiante.index and validar_email(estudiante['Correo Personal']):
            emails_validos.append(str(estudiante['Correo Personal']).strip())
        
        # Verificar dirección email
        if 'Dirección Email' in estudiante.index and validar_email(estudiante['Dirección Email']):
            email_institucional = str(estudiante['Dirección Email']).strip()
            emails_validos.append(email_institucional)
        
        return emails_validos

    def obtener_nombre_completo(estudiante, institucion):
        """Obtener nombre completo del estudiante según la institución"""
        nombre = str(estudiante.get('Nombre', 'Apreciable estudiante'))
        
        if institucion == "uveg":
            apellidos = str(estudiante.get('Apellido(s)', ''))
            return nombre, apellidos
        else:  # novauniversitas
            return nombre, ""

    def formatear_fecha(fecha_str):
        """Convierte string de fecha a formato legible"""
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

    def enviar_correo_tab1(destinatario, asunto, cuerpo):
        """Enviar correo usando Gmail"""
        try:
            msg = EmailMessage()
            msg["Subject"] = asunto
            msg["From"] = REMITENTE
            msg["To"] = destinatario
            msg.set_content(cuerpo)
            
            with smtplib.SMTP_SSL("smtp.gmail.com", 465) as smtp:
                smtp.login(REMITENTE, CLAVE_APP)
                smtp.send_message(msg)
            
            # Registrar en historial
            st.session_state.historial_envios_tab1.append({
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'asunto': asunto,
                'destinatario': destinatario,
                'estado': 'Enviado'
            })
            
            return True, "Enviado correctamente"
        except Exception as e:
            error_msg = f"Error: {str(e)}"
            st.session_state.historial_envios_tab1.append({
                'timestamp': datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                'asunto': asunto,
                'destinatario': destinatario,
                'estado': error_msg
            })
            return False, error_msg

    def enviar_a_todos_los_emails(emails_validos, asunto, cuerpo, nombre_estudiante):
        """Enviar correo a todas las direcciones válidas del estudiante"""
        resultados = []
        emails_enviados = []
        
        for email in emails_validos:
            if email and email not in emails_enviados:
                exito, mensaje = enviar_correo_tab1(email, asunto, cuerpo)
                resultados.append({'email': email, 'exito': exito, 'mensaje': mensaje})
                emails_enviados.append(email)
                
                if exito:
                    st.success(f"  ✅ Enviado a {email}")
                else:
                    st.error(f"  ❌ Error en {email}: {mensaje}")
                
                time.sleep(0.5)
        
        exitosos = sum(1 for r in resultados if r['exito'])
        return exitosos, len(resultados), emails_enviados

    def obtener_plantilla_editada(institucion, tipo_plantilla):
        """Obtener plantilla editada o base si no existe"""
        key = f"{institucion}_{tipo_plantilla}"
        if key in st.session_state.plantillas_editadas_tab1:
            return st.session_state.plantillas_editadas_tab1[key]
        else:
            return PLANTILLAS_BASE[institucion][tipo_plantilla].copy()

    def guardar_plantilla_editada(institucion, tipo_plantilla, plantilla_data):
        """Guardar plantilla editada en session state"""
        key = f"{institucion}_{tipo_plantilla}"
        st.session_state.plantillas_editadas_tab1[key] = plantilla_data

    # Interfaz principal Tab 1
    col1, col2 = st.columns([2, 1])

    with col1:
        st.subheader("🏫 Selección de Institución")
        institucion_seleccionada = st.selectbox(
            "Selecciona la institución:",
            options=["uveg", "novauniversitas"],
            format_func=lambda x: CONFIGURACIONES_INSTITUCIONES[x]["nombre"],
            help="Cada institución tiene diferentes columnas de actividades y mensajes específicos",
            key="institucion_tab1"
        )
        
        config_institucion = CONFIGURACIONES_INSTITUCIONES[institucion_seleccionada]
        st.info(f"📊 Institución: **{config_institucion['nombre']}** - {len(config_institucion['columnas_actividades'])} actividades")

    with col2:
        st.subheader("📎 Archivos Adjuntos")
        archivos_adjuntos = st.file_uploader(
            "Sube archivos adjuntos (opcional)", 
            accept_multiple_files=True,
            help="Guías, cronogramas, etc.",
            key="archivos_tab1"
        )

    # Editor de plantillas
    st.markdown("---")
    st.subheader("📝 Editor de Plantillas")

    with st.expander("✏️ Personalizar Plantillas de Correo", expanded=False):
        col_select, col_edit = st.columns([1, 2])
        
        with col_select:
            tipo_plantilla_editar = st.selectbox(
                "Selecciona plantilla a editar:",
                options=["bienvenida", "seguimiento_sin_acceso", "seguimiento_atraso", "felicitacion"],
                format_func=lambda x: PLANTILLAS_BASE[institucion_seleccionada][x]["nombre"],
                key="tipo_plantilla_tab1"
            )
            
            plantilla_actual = obtener_plantilla_editada(institucion_seleccionada, tipo_plantilla_editar)
            
            st.write(f"**Editando**: {plantilla_actual['nombre']}")
        
        with col_edit:
            # Editor de asunto
            nuevo_asunto = st.text_input(
                "Asunto del correo:",
                value=plantilla_actual["asunto"],
                help="Puedes usar variables como {nombre}, {modulo}, {semana}, etc.",
                key="asunto_tab1"
            )
            
            # Editor de mensaje
            nuevo_mensaje = st.text_area(
                "Mensaje del correo:",
                value=plantilla_actual["mensaje"],
                height=300,
                help="Puedes personalizar el mensaje, agregar tu WhatsApp, etc.",
                key="mensaje_tab1"
            )
            
            # Botones de acción
            col_save, col_reset = st.columns(2)
            
            with col_save:
                if st.button("💾 Guardar Cambios", type="primary", key="guardar_tab1"):
                    plantilla_editada = {
                        "nombre": plantilla_actual["nombre"],
                        "asunto": nuevo_asunto,
                        "mensaje": nuevo_mensaje
                    }
                    guardar_plantilla_editada(institucion_seleccionada, tipo_plantilla_editar, plantilla_editada)
                    st.success("✅ Plantilla guardada correctamente")
            
            with col_reset:
                if st.button("🔄 Restaurar Original", key="restaurar_tab1"):
                    key = f"{institucion_seleccionada}_{tipo_plantilla_editar}"
                    if key in st.session_state.plantillas_editadas_tab1:
                        del st.session_state.plantillas_editadas_tab1[key]
                    st.success("✅ Plantilla restaurada al original")
                    st.rerun()

    # Subir archivo Excel
    st.markdown("---")
    st.subheader("📁 Cargar Archivo Excel de Estudiantes")
    archivo_excel = st.file_uploader(
        f"Sube el archivo Excel de {config_institucion['nombre']}", 
        type=["xlsx", "xls"],
        help=f"Columnas requeridas: {', '.join(config_institucion['columnas_requeridas'])}",
        key="archivo_excel_tab1"
    )

    # Procesar archivo Excel para Tab 1
    if archivo_excel:
        try:
            df = pd.read_excel(archivo_excel)
            st.success("✅ Archivo Excel cargado correctamente")
            
            # Para NovaUniversitas, filtrar solo las columnas necesarias
            if institucion_seleccionada == "novauniversitas":
                columnas_necesarias = config_institucion['columnas_requeridas'] + config_institucion['columnas_actividades']
                columnas_existentes = [col for col in columnas_necesarias if col in df.columns]
                
                if len(columnas_existentes) >= len(config_institucion['columnas_requeridas']):
                    df = df[columnas_existentes]
                    st.info(f"📋 Filtradas {len(columnas_existentes)} columnas relevantes de {len(df.columns)} totales")
            
            # Verificar columnas requeridas
            columnas_faltantes = [col for col in config_institucion['columnas_requeridas'] if col not in df.columns]
            
            if columnas_faltantes:
                st.error(f"❌ Columnas faltantes: {', '.join(columnas_faltantes)}")
                st.info("Columnas disponibles:")
                st.write(list(df.columns))
            else:
                st.success("✅ Todas las columnas necesarias detectadas")
                
                # Vista previa de datos
                st.subheader("👀 Vista Previa de Estudiantes")
                datos_preview = df[config_institucion['columnas_requeridas']].head(10)
                st.dataframe(datos_preview, use_container_width=True)
                
                # Validar emails
                total_estudiantes = len(df)
                emails_validos = sum(1 for _, fila in df.iterrows() 
                                   if obtener_emails_validos(fila))
                
                st.info(f"📊 Total estudiantes: {total_estudiantes} | Con emails válidos: {emails_validos}")
                
                # Selección de tipo de envío y semana
                st.subheader("📧 Configuración de Envío")
                
                col_tipo, col_semana = st.columns([1, 1])
                
                with col_tipo:
                    tipo_envio = st.selectbox(
                        "Tipo de envío:",
                        options=["automatico", "bienvenida"],
                        format_func=lambda x: "🔍 Análisis Automático" if x == "automatico" else "🎓 Bienvenida Manual",
                        key="tipo_envio_tab1"
                    )
                
                with col_semana:
                    if tipo_envio == "automatico":
                        semana = st.selectbox(
                            "Semana para análisis:",
                            [1, 2, 3],
                            format_func=lambda x: f"Semana {x}",
                            key="semana_tab1"
                        )
                        actividades_requeridas = {1: 3, 2: 5, 3: 7}[semana]
                        st.write(f"**Actividades requeridas:** {actividades_requeridas}")
                    else:
                        st.write("**Envío de bienvenida** a todos los estudiantes")
                        semana = 1
                        actividades_requeridas = 0
                
                # Configuración adicional
                col_config1, col_config2 = st.columns(2)
                
                with col_config1:
                    modulo_manual = st.text_input("Nombre del módulo:", 
                                                value=config_institucion['modulo_default'],
                                                key="modulo_tab1")
                    
                    fecha_meta = st.date_input("Fecha meta:", 
                                             value=datetime.now() + timedelta(days=7),
                                             key="fecha_meta_tab1")
                
                with col_config2:
                    fecha_sesion = st.date_input("Fecha sesión síncrona:", 
                                               value=datetime.now() + timedelta(days=3),
                                               key="fecha_sesion_tab1")
                    
                    hora_sesion = st.time_input("Hora sesión:", 
                                              value=datetime.strptime("20:30", "%H:%M").time(),
                                              key="hora_sesion_tab1")
                
                # Botón para procesar según el tipo de envío
                if st.button("🚀 Procesar Envío", type="primary", key="procesar_tab1"):
                    
                    # Variables para personalización
                    variables_extra = {
                        'modulo': modulo_manual,
                        'semana': f"semana {semana}",
                        'institucion': config_institucion['nombre'],
                        'fecha_meta': formatear_fecha(fecha_meta.strftime("%Y-%m-%d")),
                        'fecha_sesion': formatear_fecha(fecha_sesion.strftime("%Y-%m-%d")),
                        'hora_sesion': hora_sesion.strftime("%H:%M")
                    }
                    
                    if tipo_envio == "automatico":
                        # Análisis automático
                        df['actividades_completadas'] = df.apply(
                            lambda row: contar_actividades_completadas(row, config_institucion['columnas_actividades']), 
                            axis=1
                        )
                        
                        # Clasificar estudiantes automáticamente
                        estudiantes_completos = df[df['actividades_completadas'] >= actividades_requeridas]
                        estudiantes_incompletos = df[(df['actividades_completadas'] > 0) & (df['actividades_completadas'] < actividades_requeridas)]
                        estudiantes_sin_entregas = df[df['actividades_completadas'] == 0]
                        
                        # Guardar datos para envío
                        st.session_state.datos_estudiantes_tab1 = {
                            'completos': estudiantes_completos,
                            'incompletos': estudiantes_incompletos,
                            'sin_entregas': estudiantes_sin_entregas,
                            'bienvenida': pd.DataFrame(),  # Vacío para automático
                            'semana': semana,
                            'actividades_requeridas': actividades_requeridas,
                            'institucion': institucion_seleccionada,
                            'config_institucion': config_institucion,
                            'variables_extra': variables_extra,
                            'tipo_envio': 'automatico'
                        }
                        
                        # Mostrar métricas del análisis
                        col1, col2, col3 = st.columns(3)
                        col1.metric("✅ Completos", len(estudiantes_completos))
                        col2.metric("⚠️ Incompletos", len(estudiantes_incompletos))
                        col3.metric("❌ Sin Entregas", len(estudiantes_sin_entregas))
                        
                        st.success("🎯 **Análisis automático completado - Plantillas asignadas:**")
                        st.write(f"• **{len(estudiantes_completos)}** estudiantes → 🎉 Felicitación")
                        st.write(f"• **{len(estudiantes_incompletos)}** estudiantes → 📋 Seguimiento - Tareas Pendientes")
                        st.write(f"• **{len(estudiantes_sin_entregas)}** estudiantes → ⚠️ Seguimiento - Sin Acceso")
                    
                    else:
                        # Envío manual de bienvenida
                        st.session_state.datos_estudiantes_tab1 = {
                            'completos': pd.DataFrame(),
                            'incompletos': pd.DataFrame(),
                            'sin_entregas': pd.DataFrame(),
                            'bienvenida': df,  # Todos los estudiantes
                            'semana': semana,
                            'actividades_requeridas': actividades_requeridas,
                            'institucion': institucion_seleccionada,
                            'config_institucion': config_institucion,
                            'variables_extra': variables_extra,
                            'tipo_envio': 'bienvenida'
                        }
                        
                        st.success("🎓 **Listo para envío de bienvenida:**")
                        st.write(f"• **{len(df)}** estudiantes → 🎓 Plantilla de Bienvenida")
                    
                    st.session_state.analisis_generado_tab1 = True
                    
        except Exception as e:
            st.error(f"❌ Error al procesar el archivo Excel: {str(e)}")

    # Mostrar sección de envío si el análisis está generado
    if st.session_state.analisis_generado_tab1 and st.session_state.datos_estudiantes_tab1:
        datos = st.session_state.datos_estudiantes_tab1
        estudiantes_completos = datos['completos']
        estudiantes_incompletos = datos['incompletos']
        estudiantes_sin_entregas = datos['sin_entregas']
        estudiantes_bienvenida = datos['bienvenida']
        config_institucion = datos['config_institucion']
        variables_extra = datos['variables_extra']
        tipo_envio = datos['tipo_envio']
        
        st.markdown("---")
        st.subheader("📧 Envío de Correos")
        
        if tipo_envio == "automatico":
            # Botones de envío por categoría (análisis automático)
            col_env1, col_env2, col_env3 = st.columns(3)
            
            with col_env1:
                if st.button("🎉 Enviar Felicitaciones", 
                           type="primary", 
                           disabled=len(estudiantes_completos)==0,
                           key="felicit_tab1"):
                    st.write("### 📤 Enviando Felicitaciones...")
                    for _, estudiante in estudiantes_completos.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        st.write(f"Enviando a: {nombre_completo}")
                        
                        # Obtener actividades completadas
                        actividades_completadas = obtener_actividades_completadas(
                            estudiante, 
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades']
                        )
                        
                        # Usar plantilla editada
                        plantilla = obtener_plantilla_editada(datos['institucion'], 'felicitacion')
                        asunto = plantilla['asunto'].format(**variables_extra)
                        
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_completadas)])
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_completadas=actividades_lista,
                            **variables_extra
                        )
                        
                        # Enviar correos
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                                emails_validos, asunto, mensaje, nombre_completo
                            )
                            st.info(f"📧 Enviado a {len(emails_enviados)} direcciones")
                        else:
                            st.warning(f"⚠️ Sin emails válidos para {nombre_completo}")
                        
                        time.sleep(1)
                    st.success("🎉 Felicitaciones enviadas!")
            
            with col_env2:
                if st.button("📋 Enviar Recordatorios", 
                           type="secondary", 
                           disabled=len(estudiantes_incompletos)==0,
                           key="recordat_tab1"):
                    st.write("### 📤 Enviando Recordatorios...")
                    for _, estudiante in estudiantes_incompletos.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        st.write(f"Enviando a: {nombre_completo}")
                        
                        # Obtener actividades faltantes
                        actividades_faltantes = obtener_actividades_faltantes(
                            estudiante,
                            config_institucion['columnas_actividades'],
                            config_institucion['nombres_actividades'],
                            datos['actividades_requeridas']
                        )
                        
                        # Usar plantilla editada
                        plantilla = obtener_plantilla_editada(datos['institucion'], 'seguimiento_atraso')
                        asunto = plantilla['asunto'].format(**variables_extra)
                        
                        actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_faltantes)])
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            actividades_faltantes=actividades_lista,
                            **variables_extra
                        )
                        
                        # Enviar correos
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                                emails_validos, asunto, mensaje, nombre_completo
                            )
                            st.info(f"📧 Enviado a {len(emails_enviados)} direcciones")
                        else:
                            st.warning(f"⚠️ Sin emails válidos para {nombre_completo}")
                        
                        time.sleep(1)
                    st.success("📋 Recordatorios enviados!")
            
            with col_env3:
                if st.button("⚠️ Enviar Alertas", 
                           type="secondary", 
                           disabled=len(estudiantes_sin_entregas)==0,
                           key="alertas_tab1"):
                    st.write("### 📤 Enviando Alertas...")
                    for _, estudiante in estudiantes_sin_entregas.iterrows():
                        nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                        nombre_completo = f"{nombre} {apellidos}".strip()
                        st.write(f"Enviando a: {nombre_completo}")
                        
                        # Usar plantilla editada
                        plantilla = obtener_plantilla_editada(datos['institucion'], 'seguimiento_sin_acceso')
                        asunto = plantilla['asunto'].format(**variables_extra)
                        mensaje = plantilla['mensaje'].format(
                            nombre=nombre,
                            **variables_extra
                        )
                        
                        # Enviar correos
                        emails_validos = obtener_emails_validos(estudiante)
                        if emails_validos:
                            exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                                emails_validos, asunto, mensaje, nombre_completo
                            )
                            st.info(f"📧 Enviado a {len(emails_enviados)} direcciones")
                        else:
                            st.warning(f"⚠️ Sin emails válidos para {nombre_completo}")
                        
                        time.sleep(1)
                    st.success("⚠️ Alertas enviadas!")
            
            # Botón de envío masivo para análisis automático
            st.markdown("---")
            total_estudiantes = len(estudiantes_completos) + len(estudiantes_incompletos) + len(estudiantes_sin_entregas)
            
            if total_estudiantes > 0:
                if st.button(f"🚀 Enviar Todos los Correos Automáticos ({total_estudiantes} estudiantes)", type="primary", key="masivo_tab1"):
                    st.balloons()
                    st.write("### 📤 Iniciando Envío Masivo Inteligente...")
                    
                    total_exitosos = 0
                    total_procesados = 0
                    
                    # Procesar cada categoría con su plantilla correspondiente
                    for categoria, estudiantes_cat, tipo_plantilla in [
                        ("Felicitaciones", estudiantes_completos, "felicitacion"),
                        ("Recordatorios", estudiantes_incompletos, "seguimiento_atraso"),
                        ("Alertas", estudiantes_sin_entregas, "seguimiento_sin_acceso")
                    ]:
                        
                        if len(estudiantes_cat) > 0:
                            st.write(f"**📧 Enviando {categoria}...**")
                            
                            for _, estudiante in estudiantes_cat.iterrows():
                                nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                                nombre_completo = f"{nombre} {apellidos}".strip()
                                st.write(f"→ {nombre_completo}")
                                
                                # Usar plantilla editada
                                plantilla = obtener_plantilla_editada(datos['institucion'], tipo_plantilla)
                                asunto = plantilla['asunto'].format(**variables_extra)
                                
                                # Personalizar mensaje según tipo
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
                                    actividades_faltantes = obtener_actividades_faltantes(
                                        estudiante,
                                        config_institucion['columnas_actividades'],
                                        config_institucion['nombres_actividades'],
                                        datos['actividades_requeridas']
                                    )
                                    actividades_lista = "\n".join([f"{i+1}. {act}" for i, act in enumerate(actividades_faltantes)])
                                    mensaje = plantilla['mensaje'].format(
                                        nombre=nombre,
                                        actividades_faltantes=actividades_lista,
                                        **variables_extra
                                    )
                                else:  # seguimiento_sin_acceso
                                    mensaje = plantilla['mensaje'].format(
                                        nombre=nombre,
                                        **variables_extra
                                    )
                                
                                # Enviar correos
                                emails_validos = obtener_emails_validos(estudiante)
                                if emails_validos:
                                    exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                                        emails_validos, asunto, mensaje, nombre_completo
                                    )
                                    total_exitosos += exitosos
                                    st.info(f"  📧 Enviado a {len(emails_enviados)} direcciones")
                                else:
                                    st.warning(f"  ⚠️ Sin emails válidos")
                                
                                total_procesados += 1
                                time.sleep(1)
                    
                    st.success(f"🎉 **¡Proceso Completado!** {total_exitosos} correos enviados de {total_procesados} estudiantes")
        
        else:
            # Envío manual de bienvenida
            st.write("### 🎓 Envío de Bienvenida")
            
            if st.button(f"🎓 Enviar Bienvenida a Todos ({len(estudiantes_bienvenida)} estudiantes)", type="primary", key="bienvenida_tab1"):
                st.write("### 📤 Enviando Correos de Bienvenida...")
                
                total_exitosos = 0
                total_procesados = 0
                
                # Usar plantilla editada
                plantilla = obtener_plantilla_editada(datos['institucion'], 'bienvenida')
                
                for _, estudiante in estudiantes_bienvenida.iterrows():
                    nombre, apellidos = obtener_nombre_completo(estudiante, datos['institucion'])
                    nombre_completo = f"{nombre} {apellidos}".strip()
                    st.write(f"→ {nombre_completo}")
                    
                    asunto = plantilla['asunto'].format(**variables_extra)
                    mensaje = plantilla['mensaje'].format(
                        nombre=nombre,
                        **variables_extra
                    )
                    
                    # Enviar correos
                    emails_validos = obtener_emails_validos(estudiante)
                    if emails_validos:
                        exitosos, total, emails_enviados = enviar_a_todos_los_emails(
                            emails_validos, asunto, mensaje, nombre_completo
                        )
                        total_exitosos += exitosos
                        st.info(f"  📧 Enviado a {len(emails_enviados)} direcciones")
                    else:
                        st.warning(f"  ⚠️ Sin emails válidos")
                    
                    total_procesados += 1
                    time.sleep(1)
                
                st.success(f"🎉 **¡Bienvenidas enviadas!** {total_exitosos} correos enviados de {total_procesados} estudiantes")

    # Mostrar historial de envíos Tab 1
    if st.session_state.historial_envios_tab1:
        st.markdown("---")
        st.subheader("📋 Historial de Envíos")
        
        df_historial = pd.DataFrame(st.session_state.historial_envios_tab1)
        col1, col2, col3, col4 = st.columns(4)
        
        total = len(df_historial)
        exitosos = len(df_historial[df_historial['estado'] == 'Enviado'])
        
        col1.metric("Total", total)
        col2.metric("✅ Exitosos", exitosos)
        col3.metric("❌ Fallidos", total - exitosos)
        
        if col4.button("🗑️ Limpiar Historial", key="limpiar_tab1"):
            st.session_state.historial_envios_tab1 = []
            st.rerun()
        
        st.dataframe(
            df_historial[['timestamp', 'destinatario', 'asunto', 'estado']].sort_values('timestamp', ascending=False),
            use_container_width=True
        )

# =====================================================
# TAB 2: SISTEMA DE ENVÍO MASIVO (PRÁCTICAS) - INTEGRACIÓN DEL email_app.py
# =====================================================

with tab2:
    st.header("📧 Sistema de Envío de Correos - Prácticas Profesionales")
    st.markdown("*Envío masivo personalizado para prácticas y estancias profesionales*")
    
    # Función para procesar archivos Excel
    @st.cache_data
    def load_excel_data(file):
        """Cargar datos del archivo Excel"""
        try:
            df = pd.read_excel(file, sheet_name='Calificaciones')
            return df
        except Exception as e:
            st.error(f"Error al cargar el archivo: {str(e)}")
            return None

    # Plantillas de correo para prácticas
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

    def enviar_correo_tab2(smtp_server, smtp_port, email_usuario, email_password, destinatario, asunto, contenido, archivos_adjuntos=None):
        """Función para enviar correo electrónico"""
        try:
            # Crear mensaje
            msg = MIMEMultipart()
            msg['From'] = email_usuario
            msg['To'] = destinatario
            msg['Subject'] = asunto
            
            # Agregar cuerpo del mensaje
            msg.attach(MIMEText(contenido, 'plain', 'utf-8'))
            
            # Agregar archivos adjuntos
            if archivos_adjuntos:
                for archivo in archivos_adjuntos:
                    part = MIMEBase('application', 'octet-stream')
                    part.set_payload(archivo.read())
                    encoders.encode_base64(part)
                    part.add_header(
                        'Content-Disposition',
                        f'attachment; filename= {archivo.name}'
                    )
                    msg.attach(part)
            
            # Conectar al servidor SMTP
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email_usuario, email_password)
            
            # Enviar correo
            text = msg.as_string()
            server.sendmail(email_usuario, destinatario, text)
            server.quit()
            
            return True, "Correo enviado exitosamente"
        
        except Exception as e:
            return False, f"Error al enviar correo: {str(e)}"

    # Sidebar para configuración Tab 2
    st.sidebar.header("⚙️ Configuración de Correo")

    # Configuración SMTP
    smtp_server = st.sidebar.text_input("Servidor SMTP", value="smtp.gmail.com")
    smtp_port = st.sidebar.number_input("Puerto SMTP", value=587, min_value=1, max_value=65535)
    email_usuario = st.sidebar.text_input("Email de usuario", placeholder="tu_email@gmail.com")
    email_password = st.sidebar.text_input("Contraseña", type="password", 
                                          help="Para Gmail, usa una contraseña de aplicación")

    # Información del asesor
    st.sidebar.header("👨‍🏫 Información del Asesor")
    nombre_asesor = st.sidebar.text_input("Nombre del Asesor", value="Juan Manuel Martinez Zaragoza")
    nombre_completo_asesor = st.sidebar.text_input("Nombre Completo del Asesor", value="Juan Manuel Martinez Zaragoza")
    universidad = st.sidebar.text_input("Universidad", value="UVEG")
    nombre_universidad = st.sidebar.text_input("Nombre Completo Universidad", value="Universidad Virtual del Estado de Guanajuato")

    # Área principal Tab 2
    col1, col2 = st.columns([1, 2])

    with col1:
        st.subheader("📁 Cargar Archivo Excel")
        uploaded_file = st.file_uploader("Selecciona el archivo Excel", type=['xlsx', 'xls'], key="upload_tab2")
        
        if uploaded_file is not None:
            df = load_excel_data(uploaded_file)
            if df is not None:
                st.success(f"Archivo cargado: {len(df)} registros encontrados")
                
                # Mostrar vista previa
                st.subheader("Vista previa de datos:")
                st.dataframe(df[['Nombre', 'Apellido(s)', 'Correo Personal', 'Dirección Email']].head())
                
                # Selector de destinatarios
                st.subheader("Seleccionar destinatarios:")
                enviar_a = st.selectbox("Enviar correos a:", 
                                       ["Correo Personal", "Dirección Email", "Ambos"], key="enviar_a_tab2")

    with col2:
        st.subheader("✉️ Composición del Correo")
        
        # Seleccionar plantilla
        plantilla_seleccionada = st.selectbox("Selecciona una plantilla:", 
                                            ["Bienvenida", "Sesión Síncrona", "Envío de Grabación", "Libre"], key="plantilla_tab2")
        
        # Campos adicionales según la plantilla
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
                help="Copia y pega directamente la información que genera Google Meet al crear una reunión",
                key="google_meet_tab2"
            )
        
        elif plantilla_seleccionada == "Envío de Grabación":
            numero_reto = st.text_input("Número de reto", value="1", key="numero_reto_tab2")
            enlace_grabacion = st.text_input("Enlace de grabación", 
                                           value="https://drive.google.com/file/d/1xPRTqb_hE0eXi6VOCA2Kg9ydSMj1qn8k/view?usp=sharing",
                                           key="enlace_grab_tab2")
        
        # Asunto del correo
        asunto_default = PLANTILLAS_PRACTICAS[plantilla_seleccionada]["asunto"]
        asunto = st.text_input("Asunto del correo:", value=asunto_default, key="asunto_tab2")
        
        # Contenido del correo
        contenido_default = PLANTILLAS_PRACTICAS[plantilla_seleccionada]["contenido"]
        contenido = st.text_area("Contenido del correo:", value=contenido_default, height=300, key="contenido_tab2")
        
        # Archivos adjuntos
        st.subheader("📎 Archivos Adjuntos")
        archivos_adjuntos = st.file_uploader("Selecciona archivos para adjuntar", 
                                           accept_multiple_files=True, key="archivos_tab2")
        
        if archivos_adjuntos:
            st.write("Archivos seleccionados:")
            for archivo in archivos_adjuntos:
                st.write(f"- {archivo.name} ({archivo.size} bytes)")

    # Vista previa del correo Tab 2
    if uploaded_file is not None and df is not None:
        st.markdown("---")
        st.header("👁️ Vista Previa del Correo")
        
        # Seleccionar un registro para vista previa
        indice_preview = st.selectbox("Selecciona un destinatario para vista previa:", 
                                    range(len(df)), 
                                    format_func=lambda x: f"{df.iloc[x]['Nombre']} {df.iloc[x]['Apellido(s)']}",
                                    key="preview_tab2")
        
        # Obtener datos del destinatario
        destinatario = df.iloc[indice_preview]
        nombre_completo = f"{destinatario['Nombre']} {destinatario['Apellido(s)']}"
        
        # Personalizar contenido
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
        
        # Mostrar vista previa
        st.subheader(f"📧 Para: {nombre_completo}")
        st.write(f"**Asunto:** {asunto}")
        st.write("**Contenido:**")
        st.text_area("", value=contenido_personalizado, height=200, disabled=True, key="preview_content_tab2")

    # Botón para enviar correos Tab 2
    if uploaded_file is not None and df is not None:
        st.markdown("---")
        
        col_envio1, col_envio2, col_envio3 = st.columns([1, 1, 1])
        
        with col_envio2:
            if st.button("🚀 Enviar Correos", type="primary", use_container_width=True, key="enviar_tab2"):
                if not email_usuario or not email_password:
                    st.error("Por favor, configura tu email y contraseña en la barra lateral")
                else:
                    # Barra de progreso
                    progress_bar = st.progress(0)
                    status_text = st.empty()
                    
                    enviados = 0
                    errores = 0
                    
                    for i, row in df.iterrows():
                        # Actualizar progreso
                        progress = (i + 1) / len(df)
                        progress_bar.progress(progress)
                        status_text.text(f"Enviando correo {i+1} de {len(df)}")
                        
                        # Obtener datos del destinatario
                        nombre_completo = f"{row['Nombre']} {row['Apellido(s)']}"
                        
                        # Personalizar contenido
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
                        
                        # Determinar destinatarios
                        destinatarios = []
                        if enviar_a == "Correo Personal":
                            destinatarios.append(row['Correo Personal'])
                        elif enviar_a == "Dirección Email":
                            destinatarios.append(row['Dirección Email'])
                        else:  # Ambos
                            destinatarios.extend([row['Correo Personal'], row['Dirección Email']])
                        
                        # Enviar correos
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
                    
                    # Mostrar resultados finales
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

# =====================================================
# TAB 3: SISTEMA BIENVENIDA NOVAUNIVERSITAS
# =====================================================

with tab3:
    st.header("🎓 Sistema de Bienvenida NovaUniversitas")
    st.markdown("*Envío masivo de credenciales y bienvenida institucional*")
    
    # Función para generar el correo personalizado
    def generar_mensaje_personalizado(nombre, correo_institucional, contrasena):
        """
        Genera el mensaje personalizado para cada estudiante
        """
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

    # Función para enviar correo con archivos adjuntos
    def enviar_correo_tab3(smtp_server, smtp_port, email_usuario, email_password, 
                      destinatario, asunto, mensaje, archivos_adjuntos=None):
        """
        Envía un correo electrónico usando SMTP con archivos adjuntos opcionales
        """
        try:
            # Crear el mensaje
            msg = MIMEMultipart()
            msg['From'] = email_usuario
            msg['To'] = destinatario
            msg['Subject'] = asunto
            
            # Adjuntar el mensaje
            msg.attach(MIMEText(mensaje, 'plain', 'utf-8'))
            
            # Adjuntar archivos si existen
            if archivos_adjuntos:
                for archivo in archivos_adjuntos:
                    if archivo is not None:
                        try:
                            # Crear el adjunto
                            part = MIMEBase('application', 'octet-stream')
                            part.set_payload(archivo.getvalue())
                            encoders.encode_base64(part)
                            part.add_header(
                                'Content-Disposition',
                                f'attachment; filename= {archivo.name}'
                            )
                            msg.attach(part)
                        except Exception as e:
                            return False, f"Error al adjuntar archivo {archivo.name}: {str(e)}"
            
            # Conectar al servidor SMTP
            server = smtplib.SMTP(smtp_server, smtp_port)
            server.starttls()
            server.login(email_usuario, email_password)
            
            # Enviar el correo
            text = msg.as_string()
            server.sendmail(email_usuario, destinatario, text)
            server.quit()
            
            return True, "Correo enviado exitosamente"
        
        except Exception as e:
            return False, f"Error al enviar correo: {str(e)}"

    # Sidebar para configuración SMTP Tab 3
    with st.sidebar:
        st.header("⚙️ Configuración de Correo")
        st.markdown("Configura los parámetros SMTP para el envío de correos")

        smtp_server_tab3 = st.text_input("Servidor SMTP", value="smtp.gmail.com", key="smtp_server_tab3")
        smtp_port_tab3 = st.number_input("Puerto SMTP", value=587, min_value=1, max_value=65535, key="smtp_port_tab3")
        email_usuario_tab3 = st.text_input("Correo del remitente", placeholder="tu_correo@gmail.com", key="email_usuario_tab3")
        email_password_tab3 = st.text_input("Contraseña del remitente", type="password", 
                                          help="Para Gmail, usa una contraseña de aplicación", key="email_password_tab3")

        st.markdown("---")
        st.markdown("### 📋 Instrucciones")
        st.markdown("""
        1. **Configura** los parámetros SMTP
        2. **Carga** el archivo Excel con los datos
        3. **Selecciona** la hoja correspondiente
        4. **Prueba** el envío con un estudiante
        5. **Ejecuta** el envío masivo
        """)

    # Columnas principales Tab 3
    col1, col2 = st.columns([1, 1])

    with col1:
        st.subheader("📁 Cargar Archivo Excel")
        
        # Cargar archivo
        uploaded_file_tab3 = st.file_uploader(
            "Selecciona el archivo Excel con los datos de estudiantes",
            type=['xlsx', 'xls'],
            help="El archivo debe contener las columnas: NP, Grupo, Matrícula, Nombre, Email_personal, Correo Institucional, Contraseña, Cuatrimestre, Carrera",
            key="uploaded_file_tab3"
        )
        
        if uploaded_file_tab3 is not None:
            try:
                # Leer todas las hojas del archivo
                excel_file = pd.ExcelFile(uploaded_file_tab3)
                sheet_names = excel_file.sheet_names
                
                # Selector de hoja
                selected_sheet = st.selectbox(
                    "Selecciona la hoja a procesar:",
                    sheet_names,
                    help="Elige la hoja que contiene los datos de estudiantes",
                    key="selected_sheet_tab3"
                )
                
                # Leer la hoja seleccionada
                df_tab3 = pd.read_excel(uploaded_file_tab3, sheet_name=selected_sheet)
                
                # Verificar columnas necesarias
                required_columns = ['Nombre', 'Email_personal', 'Correo Institucional', 'Contraseña']
                missing_columns = [col for col in required_columns if col not in df_tab3.columns]
                
                if missing_columns:
                    st.error(f"❌ Faltan las siguientes columnas: {', '.join(missing_columns)}")
                else:
                    # Limpiar datos
                    df_tab3 = df_tab3.dropna(subset=['Nombre', 'Email_personal', 'Correo Institucional'])
                    
                    st.success(f"✅ Archivo cargado exitosamente: {len(df_tab3)} estudiantes encontrados")
                    
                    # Mostrar vista previa
                    st.subheader("👀 Vista Previa de Datos")
                    st.dataframe(df_tab3[['Nombre', 'Email_personal', 'Correo Institucional', 'Contraseña']].head(10))
                    
                    # Estadísticas
                    st.info(f"📊 **Estadísticas:**\n- Total de estudiantes: {len(df_tab3)}\n- Columnas disponibles: {len(df_tab3.columns)}")
                    
            except Exception as e:
                st.error(f"❌ Error al leer el archivo: {str(e)}")

    with col2:
        st.subheader("📧 Configuración de Envío")
        
        if uploaded_file_tab3 is not None and 'df_tab3' in locals() and not df_tab3.empty:
            
            # Asunto del correo
            asunto_tab3 = st.text_input("Asunto del correo", 
                                  value="Bienvenida a NovaUniversitas - Credenciales de Acceso",
                                  key="asunto_tab3")
            
            # Sección de archivos adjuntos
            st.subheader("📎 Archivos Adjuntos")
            
            # Selector de cantidad de archivos
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
            
            # Mostrar resumen de archivos adjuntos
            if archivos_adjuntos_tab3:
                st.markdown("### 📋 Resumen de Archivos Adjuntos:")
                total_size = sum(archivo.size for archivo in archivos_adjuntos_tab3)
                st.write(f"**Total de archivos:** {len(archivos_adjuntos_tab3)}")
                st.write(f"**Tamaño total:** {total_size / 1024:.1f} KB")
                
                # Mostrar lista de archivos
                for i, archivo in enumerate(archivos_adjuntos_tab3, 1):
                    st.write(f"{i}. {archivo.name} ({archivo.size / 1024:.1f} KB)")
                
                # Advertencia sobre el tamaño
                if total_size > 10 * 1024 * 1024:  # 10 MB
                    st.warning("⚠️ El tamaño total de archivos es mayor a 10 MB. Algunos servidores de correo pueden rechazar el mensaje.")
            
            # Opciones de envío
            envio_option = st.radio(
                "Selecciona el tipo de envío:",
                ["Vista previa del mensaje", "Envío de prueba", "Envío masivo"],
                key="envio_option_tab3"
            )
            
            if envio_option == "Vista previa del mensaje":
                st.subheader("📝 Vista Previa del Mensaje")
                if not df_tab3.empty:
                    # Usar el primer estudiante como ejemplo
                    nombre_ejemplo = df_tab3.iloc[0]['Nombre']
                    correo_ejemplo = df_tab3.iloc[0]['Correo Institucional']
                    contrasena_ejemplo = df_tab3.iloc[0]['Contraseña'] if 'Contraseña' in df_tab3.columns else "0125070109"
                    
                    mensaje_ejemplo = generar_mensaje_personalizado(nombre_ejemplo, correo_ejemplo, contrasena_ejemplo)
                    
                    st.text_area("Mensaje que se enviará:", mensaje_ejemplo, height=400, key="mensaje_preview_tab3")
                    st.info(f"📌 Ejemplo generado para: **{nombre_ejemplo}**")
                    
                    # Mostrar archivos que se adjuntarán
                    if archivos_adjuntos_tab3:
                        st.markdown("### 📎 Archivos que se adjuntarán:")
                        for archivo in archivos_adjuntos_tab3:
                            st.write(f"• {archivo.name}")
            
            elif envio_option == "Envío de prueba":
                st.subheader("🧪 Envío de Prueba")
                
                # Seleccionar estudiante para prueba
                estudiante_prueba = st.selectbox(
                    "Selecciona un estudiante para prueba:",
                    df_tab3['Nombre'].tolist(),
                    key="estudiante_prueba_tab3"
                )
                
                if st.button("📤 Enviar Correo de Prueba", type="primary", key="enviar_prueba_tab3"):
                    if not all([smtp_server_tab3, smtp_port_tab3, email_usuario_tab3, email_password_tab3]):
                        st.error("❌ Por favor, completa toda la configuración SMTP")
                    else:
                        # Obtener datos del estudiante seleccionado
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
                
                # Mostrar información sobre archivos adjuntos
                if archivos_adjuntos_tab3:
                    st.info(f"📎 Se adjuntarán {len(archivos_adjuntos_tab3)} archivo(s) a cada correo")
                
                # Opciones adicionales
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
                        # Inicializar contadores
                        enviados = 0
                        errores = 0
                        
                        # Crear contenedor para el progreso
                        progress_bar = st.progress(0)
                        status_text = st.empty()
                        
                        # Contenedor para logs
                        log_container = st.empty()
                        logs = []
                        
                        for index, row in df_tab3.iterrows():
                            try:
                                # Generar mensaje personalizado
                                mensaje = generar_mensaje_personalizado(
                                    row['Nombre'],
                                    row['Correo Institucional'],
                                    row['Contraseña'] if 'Contraseña' in df_tab3.columns else "0125070109"
                                )
                                
                                # Enviar correo
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
                                
                                # Actualizar progreso
                                progress = (index + 1) / len(df_tab3)
                                progress_bar.progress(progress)
                                status_text.text(f"Procesando: {index + 1}/{len(df_tab3)} - Enviados: {enviados} - Errores: {errores}")
                                
                                # Mostrar logs
                                with log_container.container():
                                    st.text_area("Registro de envíos:", "\n".join(logs[-10:]), height=200, key=f"log_area_tab3_{index}")
                                
                                # Retraso entre correos
                                time.sleep(delay_between_emails)
                                
                            except Exception as e:
                                errores += 1
                                logs.append(f"❌ {row['Nombre']} - Error: {str(e)}")
                        
                        # Resultado final
                        st.success(f"🎉 **Envío completado!**\n- Enviados: {enviados}\n- Errores: {errores}")
                        
                        if enviados > 0:
                            st.balloons()
                        
                        # Descargar log completo
                        log_text = "\n".join(logs)
                        st.download_button(
                            label="📥 Descargar Log Completo",
                            data=log_text,
                            file_name="log_envio_correos.txt",
                            mime="text/plain",
                            key="download_log_tab3"
                        )

# =====================================================
# FOOTER E INFORMACIÓN ADICIONAL
# =====================================================

st.markdown("---")

# Información adicional en un expander
with st.expander("ℹ️ Instrucciones y Ayuda del Sistema Unificado"):
    
    col_tab1_help, col_tab2_help, col_tab3_help = st.columns(3)
    
    with col_tab1_help:
        st.markdown("### 🎯 Sistema Inteligente")
        st.markdown("""
        **Características principales:**
        - Análisis automático por progreso académico
        - Plantillas editables y personalizables
        - Soporte para UVEG y NovaUniversitas
        - Envío inteligente por categorías
        - Editor de mensajes integrado
        
        **Flujo de trabajo:**
        1. Configurar credenciales Gmail
        2. Seleccionar institución
        3. Personalizar plantillas (opcional)
        4. Cargar archivo Excel
        5. Elegir análisis automático o bienvenida
        6. Procesar y enviar correos
        
        **Variables disponibles:**
        - `{nombre}` - Nombre del estudiante
        - `{modulo}` - Nombre del módulo
        - `{semana}` - Semana actual
        - `{actividades_completadas}` - Lista actividades hechas
        - `{actividades_faltantes}` - Lista actividades pendientes
        """)
    
    with col_tab2_help:
        st.markdown("### 📧 Sistema Prácticas")
        st.markdown("""
        **Características principales:**
        - Plantillas predefinidas para prácticas
        - Personalización de fechas y datos
        - Soporte para archivos adjuntos
        - Vista previa de correos
        - Envío a múltiples direcciones
        
        **Flujo de trabajo:**
        1. Configurar SMTP en sidebar
        2. Cargar archivo Excel de estudiantes
        3. Seleccionar plantilla de correo
        4. Personalizar campos específicos
        5. Vista previa del mensaje
        6. Ejecutar envío masivo
        
        **Plantillas disponibles:**
        - Bienvenida y primer reto
        - Sesión síncrona
        - Envío de grabación
        - Mensaje libre personalizable
        """)
    
    with col_tab3_help:
        st.markdown("### 🎓 Sistema Bienvenida")
        st.markdown("""
        **Características principales:**
        - Envío masivo de credenciales
        - Soporte para archivos adjuntos
        - Mensaje de bienvenida personalizado
        - Vista previa y envío de prueba
        - Log detallado de envíos
        
        **Flujo de trabajo:**
        1. Configurar SMTP en sidebar
        2. Cargar archivo Excel con credenciales
        3. Seleccionar hoja de datos
        4. Configurar archivos adjuntos
        5. Hacer prueba de envío
        6. Ejecutar envío masivo
        
        **Columnas requeridas en Excel:**
        - `Nombre` - Nombre del estudiante
        - `Email_personal` - Email personal
        - `Correo Institucional` - Email institucional
        - `Contraseña` - Contraseña temporal
        """)

    st.markdown("### 🔧 Configuración Técnica")
    st.markdown("""
    **Para Gmail:**
    - Servidor SMTP: `smtp.gmail.com`
    - Puerto: `587` (Tabs 2 y 3) o `465` (Tab 1)
    - Requiere autenticación de 2 factores
    - Usa contraseña de aplicación (no tu contraseña normal)
    
    **Archivos adjuntos:**
    - Máximo 5 archivos por correo
    - Tamaño total recomendado: menor a 10 MB
    - Formatos: PDF, DOC, DOCX, XLS, XLSX, TXT, JPG, PNG, ZIP
    
    **Diferencias entre pestañas:**
    - **Pestaña 1**: Análisis académico inteligente con plantillas dinámicas
    - **Pestaña 2**: Envío masivo para prácticas profesionales con plantillas específicas
    - **Pestaña 3**: Envío de credenciales de bienvenida con archivos adjuntos
    """)

# Instrucciones específicas
st.markdown("### 📝 Instrucciones de uso:")

st.markdown("""
#### 🎯 **Pestaña 1 - Sistema Inteligente:**
1. **Configuración:** Completa credenciales Gmail y verifica la conexión
2. **Institución:** Selecciona UVEG o NovaUniversitas según corresponda
3. **Plantillas:** Personaliza mensajes según tus necesidades
4. **Archivo Excel:** Carga archivo con columnas de estudiantes y actividades
5. **Análisis:** Elige análisis automático por progreso o envío de bienvenida
6. **Envío:** Procesa y envía correos segmentados por categoría

#### 📧 **Pestaña 2 - Sistema Prácticas:**
1. **Configuración:** Completa datos SMTP y del asesor en la barra lateral
2. **Archivo Excel:** Carga archivo con hoja 'Calificaciones'
3. **Plantilla:** Selecciona tipo de mensaje (Bienvenida, Sesión, Grabación, Libre)
4. **Personalización:** Completa fechas y datos específicos según plantilla
5. **Vista previa:** Revisa cómo se verá el correo antes de enviarlo
6. **Envío:** Ejecuta envío masivo con barra de progreso

#### 🎓 **Pestaña 3 - Sistema Bienvenida:**
1. **Configuración:** Configura parámetros SMTP en la barra lateral
2. **Archivo Excel:** Carga archivo con credenciales de estudiantes
3. **Hoja:** Selecciona la hoja correcta del archivo Excel
4. **Adjuntos:** Configura archivos para adjuntar (opcional)
5. **Prueba:** Realiza envío de prueba a un estudiante
6. **Masivo:** Ejecuta envío masivo con log detallado

### 🛡️ Seguridad:
- Las credenciales no se almacenan permanentemente
- Se recomienda usar cuentas de correo dedicadas para envíos masivos
- Verifica siempre el contenido antes de enviar
- Usa contraseñas de aplicación, nunca tu contraseña principal
""")

# Footer
st.markdown("---")
st.markdown(
    """
    <div style='text-align: center; color: #666;'>
        <p><strong>Sistema Integral de Correos Académicos</strong></p>
        <p>UVEG & NovaUniversitas - Plataforma Unificada</p>
        <p>Desarrollado para Coordinación Virtual y Asesores Académicos</p>
        <p>Versión 3.0 - Integración Completa de Sistemas</p>
    </div>
    """,
    unsafe_allow_html=True
)
