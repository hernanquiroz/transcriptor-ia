
import os
import tempfile
import html
from datetime import datetime
from io import BytesIO

import av
import streamlit as st
from faster_whisper import WhisperModel
from docx import Document
from docx.shared import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    PageBreak,
)

# ==========================================================
# CONFIGURACIÓN GENERAL
# ==========================================================

st.set_page_config(
    page_title="Transcriptor IA",
    page_icon="🎙️",
    layout="centered",
    initial_sidebar_state="collapsed",
)

NOMBRE_APP = "Transcriptor IA"
CREDITO = "Diseño por IA, Ing. Hernàn Quiroz"
DURACION_MAXIMA_SEGUNDOS = 90 * 60
TAMANO_MAXIMO_MB = 500

# ==========================================================
# ESTILOS DE LA INTERFAZ
# ==========================================================

st.markdown(
    """
    <style>
    .main .block-container {
        max-width: 900px;
        padding-top: 2rem;
        padding-bottom: 3rem;
    }

    .titulo-app {
        text-align: center;
        font-size: 2.5rem;
        font-weight: 800;
        color: #3984D6;
        margin-bottom: 0.2rem;
    }

    .subtitulo-app {
        text-align: center;
        color: #64748b;
        font-size: 1rem;
        margin-bottom: 0.5rem;
    }

    .firma-app {
        text-align: center;
        font-size: 0.95rem;
        color: #64748b;
        border-bottom: 1px solid #dbe2ea;
        padding-bottom: 1.2rem;
        margin-bottom: 1.5rem;
    }

    .nota-app {
        background-color: rgba(128, 128, 128, 0.08);
        padding: 14px;
        border-radius: 10px;
        margin: 10px 0;
    }

    div.stButton > button[kind="primary"] {
        width: 100%;
        min-height: 3rem;
        font-weight: 700;
    }
    </style>
    """,
    unsafe_allow_html=True,
)

# ==========================================================
# ENCABEZADO Y FIRMA
# ==========================================================

st.markdown(
    f'<div class="titulo-app">🎙️ {NOMBRE_APP}</div>',
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="subtitulo-app">'
    "Convierte tus audios y videos en texto con inteligencia artificial"
    "</div>",
    unsafe_allow_html=True,
)

st.markdown(
    f'<div class="firma-app">{html.escape(CREDITO)}</div>',
    unsafe_allow_html=True,
)

# ==========================================================
# MODELO DE INTELIGENCIA ARTIFICIAL
# ==========================================================

@st.cache_resource(show_spinner=False)
def cargar_modelo():
    """
    Carga Whisper una sola vez por instancia de la aplicación.
    CPU e INT8 reducen el consumo de memoria.
    """
    return WhisperModel(
        "small",
        device="cpu",
        compute_type="int8",
    )


# ==========================================================
# FUNCIONES AUXILIARES
# ==========================================================

def obtener_duracion(ruta_archivo):
    """
    Obtiene la duración mediante PyAV.
    Si no está disponible la duración global, intenta
    calcularla a partir de las pistas del archivo.
    """
    with av.open(ruta_archivo) as contenedor:
        if contenedor.duration is not None:
            return float(contenedor.duration / av.time_base)

        duraciones = []

        for flujo in contenedor.streams:
            if flujo.duration is not None and flujo.time_base is not None:
                duraciones.append(
                    float(flujo.duration * flujo.time_base)
                )

        if duraciones:
            return max(duraciones)

    raise ValueError(
        "No fue posible determinar la duración del archivo."
    )


def formato_tiempo(segundos):
    segundos = max(0, int(segundos))
    horas = segundos // 3600
    minutos = (segundos % 3600) // 60
    segundos_restantes = segundos % 60

    if horas:
        return f"{horas:02d}:{minutos:02d}:{segundos_restantes:02d}"

    return f"{minutos:02d}:{segundos_restantes:02d}"


def transcribir_audio(
    ruta_archivo,
    idioma,
    incluir_marcas_tiempo,
):
    modelo = cargar_modelo()

    idioma_modelo = None if idioma == "Detectar automáticamente" else idioma

    segmentos, informacion = modelo.transcribe(
        ruta_archivo,
        language=idioma_modelo,
        beam_size=5,
        vad_filter=True,
        condition_on_previous_text=True,
    )

    resultado = []
    lineas = []

    for segmento in segmentos:
        texto = segmento.text.strip()

        if not texto:
            continue

        inicio = formato_tiempo(segmento.start)
        fin = formato_tiempo(segmento.end)

        resultado.append(
            {
                "inicio": inicio,
                "fin": fin,
                "texto": texto,
            }
        )

        if incluir_marcas_tiempo:
            lineas.append(f"[{inicio} - {fin}] {texto}")
        else:
            lineas.append(texto)

    return "\n".join(lineas), resultado, informacion


# ==========================================================
# EXPORTACIÓN A WORD
# ==========================================================

def generar_word(texto, nombre_archivo):
    documento = Document()

    seccion = documento.sections[0]
    seccion.top_margin = Inches(0.7)
    seccion.bottom_margin = Inches(0.7)
    seccion.left_margin = Inches(0.8)
    seccion.right_margin = Inches(0.8)

    titulo = documento.add_heading(NOMBRE_APP, level=0)
    titulo.alignment = 1

    firma = documento.add_paragraph(CREDITO)
    firma.alignment = 1

    documento.add_paragraph(
        f"Archivo original: {nombre_archivo}"
    )
    documento.add_paragraph(
        "Fecha de generación: "
        + datetime.now().strftime("%d/%m/%Y %H:%M")
    )

    documento.add_heading("Transcripción", level=1)

    for linea in texto.splitlines():
        documento.add_paragraph(linea)

    estilo_normal = documento.styles["Normal"]
    estilo_normal.font.name = "Arial"
    estilo_normal.font.size = Pt(10)

    memoria = BytesIO()
    documento.save(memoria)
    memoria.seek(0)

    return memoria.getvalue()


# ==========================================================
# EXPORTACIÓN A PDF
# ==========================================================

def generar_pdf(texto, nombre_archivo):
    memoria = BytesIO()

    documento = SimpleDocTemplate(
        memoria,
        pagesize=letter,
        rightMargin=0.65 * inch,
        leftMargin=0.65 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
        title=NOMBRE_APP,
        author=CREDITO,
    )

    estilos = getSampleStyleSheet()

    estilo_titulo = ParagraphStyle(
        "TituloApp",
        parent=estilos["Title"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#3984D6"),
        fontName="Helvetica-Bold",
        fontSize=20,
        leading=25,
        spaceAfter=8,
    )

    estilo_firma = ParagraphStyle(
        "FirmaApp",
        parent=estilos["Normal"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#64748b"),
        fontSize=10,
        leading=14,
        spaceAfter=18,
    )

    estilo_cuerpo = ParagraphStyle(
        "TextoTranscripcion",
        parent=estilos["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=14,
        spaceAfter=7,
        wordWrap="LTR",
    )

    contenido = [
        Paragraph(html.escape(NOMBRE_APP), estilo_titulo),
        Paragraph(html.escape(CREDITO), estilo_firma),
        Paragraph(
            "<b>Archivo original:</b> "
            + html.escape(nombre_archivo),
            estilos["Normal"],
        ),
        Paragraph(
            "<b>Fecha:</b> "
            + datetime.now().strftime("%d/%m/%Y %H:%M"),
            estilos["Normal"],
        ),
        Spacer(1, 16),
        Paragraph("Transcripción", estilos["Heading1"]),
        Spacer(1, 8),
    ]

    for linea in texto.splitlines():
        if linea.strip():
            contenido.append(
                Paragraph(
                    html.escape(linea),
                    estilo_cuerpo,
                )
            )
        else:
            contenido.append(Spacer(1, 6))

    documento.build(contenido)
    memoria.seek(0)

    return memoria.getvalue()


# ==========================================================
# INTERFAZ PRINCIPAL
# ==========================================================

st.markdown("### 📁 1. Selecciona tu archivo")

st.info(
    "Puedes cargar archivos de audio o video de hasta "
    f"{TAMANO_MAXIMO_MB} MB y con una duración máxima de 90 minutos."
)

archivo = st.file_uploader(
    "Sube un archivo de audio o video",
    type=[
        "mp3",
        "wav",
        "m4a",
        "aac",
        "flac",
        "ogg",
        "opus",
        "wma",
        "mp4",
        "mov",
        "mkv",
        "avi",
        "webm",
        "mpeg",
        "mpg",
    ],
    help="El archivo se procesará en el servidor de la aplicación.",
)

st.markdown("### ⚙️ 2. Configura la transcripción")

idioma = st.selectbox(
    "Idioma del audio",
    options=[
        "Español",
        "Detectar automáticamente",
        "Inglés",
        "Francés",
        "Portugués",
        "Italiano",
        "Alemán",
    ],
    index=0,
)

idiomas_modelo = {
    "Español": "es",
    "Inglés": "en",
    "Francés": "fr",
    "Portugués": "pt",
    "Italiano": "it",
    "Alemán": "de",
}

incluir_marcas_tiempo = st.checkbox(
    "Incluir marcas de tiempo en la transcripción",
    value=False,
)

st.caption(
    "El modelo small ofrece un equilibrio entre calidad y consumo "
    "de recursos. La velocidad depende de la carga del servidor."
)

# ==========================================================
# PROCESAMIENTO
# ==========================================================

if archivo is not None:
    tamano_mb = archivo.size / (1024 * 1024)

    st.markdown("### 📋 Información del archivo")

    st.write(f"**Nombre:** {archivo.name}")
    st.write(f"**Tamaño:** {tamano_mb:.2f} MB")

    if tamano_mb > TAMANO_MAXIMO_MB:
        st.error(
            f"El archivo supera el límite de {TAMANO_MAXIMO_MB} MB."
        )
        st.stop()

    # Guardar temporalmente el archivo para procesarlo.
    ruta_temporal = None

    try:
        extension = os.path.splitext(archivo.name)[1]

        with tempfile.NamedTemporaryFile(
            delete=False,
            suffix=extension,
        ) as temporal:
            temporal.write(archivo.getvalue())
            ruta_temporal = temporal.name

        duracion = obtener_duracion(ruta_temporal)

        st.write(
            f"**Duración:** {formato_tiempo(duracion)}"
        )

        if duracion <= 0:
            st.error("El archivo tiene una duración no válida.")
            st.stop()

        if duracion > DURACION_MAXIMA_SEGUNDOS:
            st.error(
                "El archivo supera la duración máxima de 90 minutos."
            )
            st.stop()

    except Exception as error:
        st.error(
            "No se pudo leer el archivo. Comprueba que sea un audio "
            "o video válido y que contenga una pista de audio."
        )
        st.caption(f"Detalle técnico: {error}")
        st.stop()

    if st.button(
        "🚀 Transcribir con inteligencia artificial",
        type="primary",
    ):
        try:
            with st.spinner(
                "La IA está procesando el archivo. "
                "Esto puede tardar varios minutos..."
            ):
                texto, segmentos, informacion = transcribir_audio(
                    ruta_temporal,
                    idiomas_modelo.get(idioma, idioma),
                    incluir_marcas_tiempo,
                )

            if not texto.strip():
                st.warning(
                    "No se detectó contenido hablado. "
                    "Prueba con un archivo de mejor calidad."
                )
            else:
                st.session_state["transcripcion"] = texto
                st.session_state["nombre_original"] = archivo.name
                st.session_state["segmentos"] = segmentos

                st.success("¡Transcripción completada!")

                idioma_detectado = getattr(
                    informacion,
                    "language",
                    None,
                )

                if idioma_detectado:
                    st.caption(
                        f"Idioma detectado por Whisper: {idioma_detectado}"
                    )

        except Exception as error:
            st.error(
                "Ocurrió un problema durante la transcripción."
            )
            st.caption(f"Detalle técnico: {error}")

        finally:
            if ruta_temporal and os.path.exists(ruta_temporal):
                try:
                    os.remove(ruta_temporal)
                except OSError:
                    pass

# ==========================================================
# RESULTADOS Y DESCARGAS
# ==========================================================

if "transcripcion" in st.session_state:
    st.markdown("---")
    st.markdown("### 📝 3. Revisa y edita el resultado")

    texto_editado = st.text_area(
        "Puedes corregir el texto antes de descargarlo",
        value=st.session_state["transcripcion"],
        height=350,
        key="texto_editado",
    )

    st.caption(
        f"Palabras aproximadas: {len(texto_editado.split())}"
    )

    nombre_original = st.session_state.get(
        "nombre_original",
        "audio",
    )

    nombre_base = os.path.splitext(nombre_original)[0]
    nombre_seguro = "".join(
        caracter
        if caracter.isalnum() or caracter in ("-", "_")
        else "_"
        for caracter in nombre_base
    ).strip("_") or "transcripcion"

    st.markdown("### 📥 4. Descarga tu transcripción")

    columna_word, columna_pdf = st.columns(2)

    with columna_word:
        contenido_word = generar_word(
            texto_editado,
            nombre_original,
        )

        st.download_button(
            label="📘 Descargar Word",
            data=contenido_word,
            file_name=f"{nombre_seguro}.docx",
            mime=(
                "application/vnd.openxmlformats-officedocument."
                "wordprocessingml.document"
            ),
            use_container_width=True,
        )

    with columna_pdf:
        contenido_pdf = generar_pdf(
            texto_editado,
            nombre_original,
        )

        st.download_button(
            label="📕 Descargar PDF",
            data=contenido_pdf,
            file_name=f"{nombre_seguro}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    if st.button("🗑️ Limpiar resultado"):
        for clave in (
            "transcripcion",
            "nombre_original",
            "segmentos",
            "texto_editado",
        ):
            st.session_state.pop(clave, None)

        st.rerun()

# ==========================================================
# PIE DE PÁGINA
# ==========================================================

st.markdown("---")

st.markdown(
    f"""
    <div style="text-align:center; color:#64748b; font-size:0.85rem;">
        {html.escape(NOMBRE_APP)}<br>
        {html.escape(CREDITO)}<br><br>
        Herramienta de transcripción asistida por IA
    </div>
    """,
    unsafe_allow_html=True,
)

