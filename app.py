import io
import os
import tempfile
from pathlib import Path
from datetime import datetime

import av
import streamlit as st
from faster_whisper import WhisperModel
from docx import Document
from docx.shared import Inches, Pt
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer
from xml.sax.saxutils import escape

# ----------------------------
# Configuración de la aplicación
# ----------------------------
MAX_DURATION_SECONDS = 90 * 60
MAX_FILE_BYTES = 500 * 1024 * 1024  # 500 MB; el alojamiento puede imponer un límite menor.
SUPPORTED_EXTENSIONS = {
    "mp3", "wav", "m4a", "aac", "ogg", "flac",
    "mp4", "mov", "mkv", "webm", "mpeg", "mpga"
}

st.set_page_config(
    page_title="Transcriptor IA",
    page_icon="🎙️",
    layout="wide",
    initial_sidebar_state="collapsed",
)

# HTML/CSS personalizado; los controles siguen siendo widgets nativos de Streamlit
# para que la carga, el procesamiento y las descargas se conecten correctamente a Python.
st.markdown(
    """
    <style>
    :root { --ink:#172033; --muted:#667085; --accent:#1769aa; --line:#e3e8ef; }
    .stApp { background: linear-gradient(180deg, #f5f8fc 0%, #ffffff 360px); }
    .block-container { max-width: 1050px; padding-top: 2rem; padding-bottom: 3rem; }
    .hero {
      padding: 1.8rem 2rem; border-radius: 22px;
      background: linear-gradient(125deg, #102a43, #1769aa);
      color: white; margin-bottom: 1.25rem;
      box-shadow: 0 12px 32px rgba(16,42,67,.16);
    }
    .hero h1 { color:white; margin:0 0 .45rem 0; font-size:2.15rem; }
    .hero p { color:#e7f1fb; margin:0; font-size:1rem; line-height:1.55; }
    .pill {
      display:inline-block; padding:.32rem .65rem; margin:.8rem .35rem 0 0;
      border:1px solid rgba(255,255,255,.32); border-radius:999px;
      color:white; font-size:.82rem;
    }
    .section-card {
      background:white; border:1px solid var(--line); border-radius:18px;
      padding:1rem 1.2rem; margin:.75rem 0;
    }
    .small-muted { color:#667085; font-size:.9rem; }
    div[data-testid="stFileUploader"] {
      background:white; border:1px dashed #aab8c8; border-radius:14px; padding:.5rem;
    }
    div.stButton > button[kind="primary"] {
      background:#1769aa; border:0; border-radius:10px; min-height:3rem;
      font-weight:650;
    }
    div.stDownloadButton > button { border-radius:10px; min-height:2.8rem; }
    </style>
    <div class="hero">
      <h1>🎙️ Transcriptor IA</h1>
      <p>Convierte audio y video en texto editable. Revisa la transcripción y exporta a Word o PDF.</p>
      <span class="pill">Hasta 90 minutos</span>
      <span class="pill">Español y otros idiomas</span>
      <span class="pill">Word + PDF</span>
    </div>
    """,
    unsafe_allow_html=True,
)

st.markdown(
    '<div class="section-card"><b>Cómo funciona</b><br>'
    '<span class="small-muted">1. Sube el archivo · 2. Selecciona el idioma · '
    '3. Transcribe · 4. Corrige el texto · 5. Descarga el documento.</span></div>',
    unsafe_allow_html=True,
)

with st.expander("Formatos, privacidad y límites", expanded=False):
    st.write(
        "Formatos admitidos: MP3, WAV, M4A, AAC, OGG, FLAC, MP4, MOV, MKV, "
        "WEBM, MPEG y MPGA. El soporte final depende del códec del archivo."
    )
    st.write(
        "La duración máxima configurada es 90 minutos y el límite de carga de esta "
        "aplicación es 500 MB. Streamlit Cloud puede imponer límites de recursos "
        "o de subida adicionales."
    )
    st.write(
        "El archivo se procesa en el servidor donde está alojada la aplicación. "
        "No subas grabaciones confidenciales si no tienes autorización para alojarlas "
        "en ese servicio. El archivo temporal se elimina al terminar el procesamiento."
    )
    st.write(
        "La transcripción automática puede cometer errores, especialmente con ruido, "
        "varios hablantes, nombres propios o voces lejanas. Revisa el texto antes de usarlo."
    )

@st.cache_resource(show_spinner=False)
def load_model():
    # Modelo multilingüe. "small" es un equilibrio razonable para una instancia CPU.
    # En un servidor con más recursos se puede cambiar a "medium".
    return WhisperModel("small", device="cpu", compute_type="int8")

def get_media_duration(path: str):
    """Devuelve duración aproximada en segundos y comprueba que exista una pista de audio."""
    with av.open(path) as container:
        audio_streams = [stream for stream in container.streams if stream.type == "audio"]
        if not audio_streams:
            raise ValueError("El archivo no contiene una pista de audio reconocible.")
        if container.duration is not None:
            duration = float(container.duration / av.time_base)
        else:
            stream = audio_streams[0]
            if stream.duration is None or stream.time_base is None:
                raise ValueError("No fue posible determinar la duración del archivo.")
            duration = float(stream.duration * stream.time_base)
        return duration

def format_time(seconds: float) -> str:
    total = max(0, int(seconds))
    hours, remainder = divmod(total, 3600)
    minutes, secs = divmod(remainder, 60)
    if hours:
        return f"{hours:02d}:{minutes:02d}:{secs:02d}"
    return f"{minutes:02d}:{secs:02d}"

def build_transcript_docx(title: str, transcript: str, source_name: str, language: str):
    output = io.BytesIO()
    doc = Document()
    section = doc.sections[0]
    section.top_margin = Inches(0.75)
    section.bottom_margin = Inches(0.75)
    section.left_margin = Inches(0.85)
    section.right_margin = Inches(0.85)

    heading = doc.add_heading(title.strip() or "Transcripción", 0)
    heading.alignment = 0
    p = doc.add_paragraph()
    p.add_run("Archivo: ").bold = True
    p.add_run(source_name)
    p = doc.add_paragraph()
    p.add_run("Idioma detectado: ").bold = True
    p.add_run(language)
    p = doc.add_paragraph()
    p.add_run("Generado: ").bold = True
    p.add_run(datetime.now().strftime("%Y-%m-%d %H:%M"))
    doc.add_paragraph("")

    for line in transcript.splitlines():
        if line.strip():
            para = doc.add_paragraph(line.strip())
            para.paragraph_format.space_after = Pt(6)

    doc.save(output)
    return output.getvalue()

def build_transcript_pdf(title: str, transcript: str, source_name: str, language: str):
    output = io.BytesIO()
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(
        name="TranscriptTitle",
        parent=styles["Title"],
        alignment=TA_CENTER,
        textColor=colors.HexColor("#17324d"),
        fontName="Helvetica-Bold",
        fontSize=19,
        leading=23,
        spaceAfter=14,
    ))
    styles.add(ParagraphStyle(
        name="TranscriptBody",
        parent=styles["BodyText"],
        fontName="Helvetica",
        fontSize=9.5,
        leading=13,
        spaceAfter=7,
        wordWrap="LTR",
    ))
    styles.add(ParagraphStyle(
        name="Meta",
        parent=styles["BodyText"],
        fontSize=8.5,
        leading=11,
        textColor=colors.HexColor("#566273"),
        spaceAfter=4,
    ))

    story = [
        Paragraph(escape(title.strip() or "Transcripción"), styles["TranscriptTitle"]),
        Paragraph(f"<b>Archivo:</b> {escape(source_name)}", styles["Meta"]),
        Paragraph(f"<b>Idioma detectado:</b> {escape(language)}", styles["Meta"]),
        Paragraph(f"<b>Generado:</b> {datetime.now().strftime('%Y-%m-%d %H:%M')}", styles["Meta"]),
        Spacer(1, 12),
    ]
    for line in transcript.splitlines():
        if line.strip():
            story.append(Paragraph(escape(line.strip()), styles["TranscriptBody"]))

    pdf = SimpleDocTemplate(
        output,
        pagesize=A4,
        rightMargin=2 * cm,
        leftMargin=2 * cm,
        topMargin=1.8 * cm,
        bottomMargin=1.8 * cm,
        title=title.strip() or "Transcripción",
        author="Transcriptor IA",
    )
    pdf.build(story)
    return output.getvalue()

# Retener el resultado al cambiar widgets o descargar.
if "transcript_text" not in st.session_state:
    st.session_state.transcript_text = ""
if "source_name" not in st.session_state:
    st.session_state.source_name = ""
if "detected_language" not in st.session_state:
    st.session_state.detected_language = "No disponible"
if "document_title" not in st.session_state:
    st.session_state.document_title = "Transcripción"

col_left, col_right = st.columns([1.25, 0.75], gap="large")
with col_left:
    uploaded = st.file_uploader(
        "Selecciona un archivo de audio o video",
        type=sorted(SUPPORTED_EXTENSIONS),
        help="Duración máxima: 90 minutos. El tamaño máximo configurado es 500 MB.",
    )
with col_right:
    document_title = st.text_input(
        "Título del documento",
        value=st.session_state.document_title,
        max_chars=150,
    )
    language_label = st.selectbox(
        "Idioma del audio",
        [
            "Detectar automáticamente",
            "Español",
            "Inglés",
            "Francés",
            "Portugués",
            "Italiano",
            "Alemán",
        ],
        index=0,
    )
    timestamps = st.checkbox("Incluir marcas de tiempo", value=True)

if uploaded is not None:
    st.caption(f"Archivo seleccionado: **{uploaded.name}** · {uploaded.size / (1024 * 1024):.1f} MB")
    if uploaded.size > MAX_FILE_BYTES:
        st.error("El archivo supera el límite configurado de 500 MB. Reduce el tamaño del archivo e inténtalo de nuevo.")
    else:
        st.caption("La transcripción puede tardar bastante en grabaciones largas. Mantén abierta esta página.")

    transcribe_clicked = st.button(
        "▶  Transcribir archivo",
        type="primary",
        use_container_width=True,
        disabled=uploaded.size > MAX_FILE_BYTES,
    )

    if transcribe_clicked:
        suffix = Path(uploaded.name).suffix.lower()
        temp_path = None
        try:
            with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as temp_file:
                temp_file.write(uploaded.getvalue())
                temp_path = temp_file.name

            with st.spinner("Comprobando el archivo y su duración..."):
                duration = get_media_duration(temp_path)

            if duration <= 0:
                st.error("No se pudo determinar una duración válida para el archivo.")
            elif duration > MAX_DURATION_SECONDS:
                st.error(
                    f"El archivo dura {format_time(duration)}. "
                    "El máximo permitido es 01:30:00."
                )
            else:
                st.success(f"Duración detectada: {format_time(duration)}")
                language_codes = {
                    "Español": "es",
                    "Inglés": "en",
                    "Francés": "fr",
                    "Portugués": "pt",
                    "Italiano": "it",
                    "Alemán": "de",
                }
                language_code = language_codes.get(language_label)

                with st.spinner("Cargando el modelo de reconocimiento de voz..."):
                    model = load_model()

                progress = st.progress(0, text="Preparando la transcripción...")
                status = st.empty()
                segments, info = model.transcribe(
                    temp_path,
                    language=language_code,
                    beam_size=3,
                    vad_filter=True,
                    condition_on_previous_text=True,
                )

                lines = []
                last_progress = -1
                for segment in segments:
                    segment_text = segment.text.strip()
                    if segment_text:
                        if timestamps:
                            lines.append(f"[{format_time(segment.start)}] {segment_text}")
                        else:
                            lines.append(segment_text)

                    current = min(99, int((segment.end / duration) * 100)) if duration else 0
                    if current >= last_progress + 2:
                        progress.progress(current, text=f"Transcribiendo: {current}%")
                        status.caption(f"Procesado aproximadamente hasta {format_time(segment.end)} de {format_time(duration)}.")
                        last_progress = current

                transcript = "\n".join(lines).strip()
                progress.progress(100, text="Transcripción terminada")
                status.empty()

                if not transcript:
                    st.warning("No se detectó voz. Prueba con un archivo más claro o verifica que contenga audio.")
                else:
                    st.session_state.transcript_text = transcript
                    st.session_state.source_name = uploaded.name
                    st.session_state.detected_language = info.language or "No identificado"
                    st.session_state.document_title = document_title or "Transcripción"
                    st.success("Transcripción completada. Revisa el texto antes de descargar.")
        except Exception as exc:
            st.error(
                "No fue posible completar la transcripción. Comprueba que el archivo "
                "esté dañado o protegido y que el servidor tenga recursos suficientes. "
                f"Detalle técnico: {exc}"
            )
        finally:
            if temp_path and os.path.exists(temp_path):
                try:
                    os.remove(temp_path)
                except OSError:
                    pass

if st.session_state.transcript_text:
    st.markdown("---")
    st.subheader("📝 Revisa y edita la transcripción")
    st.caption(
        f"Idioma detectado por el modelo: {st.session_state.detected_language} · "
        f"Archivo: {st.session_state.source_name}"
    )
    edited_text = st.text_area(
        "Texto transcrito",
        value=st.session_state.transcript_text,
        height=420,
        key="edited_transcript",
        help="Puedes corregir nombres, puntuación y errores antes de exportar.",
    )
    title_for_export = st.text_input(
        "Título para los documentos",
        value=st.session_state.document_title,
        key="title_for_export",
    )

    word_data = build_transcript_docx(
        title_for_export, edited_text, st.session_state.source_name,
        st.session_state.detected_language
    )
    pdf_data = build_transcript_pdf(
        title_for_export, edited_text, st.session_state.source_name,
        st.session_state.detected_language
    )

    export_col1, export_col2 = st.columns(2)
    safe_base = "".join(
        c if c.isalnum() or c in ("-", "_") else "_"
        for c in (title_for_export.strip() or "transcripcion")
    ).strip("_")[:80] or "transcripcion"

    with export_col1:
        st.download_button(
            "⬇ Descargar Word (.docx)",
            data=word_data,
            file_name=f"{safe_base}.docx",
            mime="application/vnd.openxmlformats-officedocument.wordprocessingml.document",
            use_container_width=True,
        )
    with export_col2:
        st.download_button(
            "⬇ Descargar PDF (.pdf)",
            data=pdf_data,
            file_name=f"{safe_base}.pdf",
            mime="application/pdf",
            use_container_width=True,
        )

    if st.button("Limpiar transcripción", use_container_width=True):
        for key in ["transcript_text", "source_name", "detected_language", "edited_transcript"]:
            st.session_state.pop(key, None)
        st.rerun()

st.markdown("---")
st.markdown(
    '<p class="small-muted" style="text-align:center">'
    'Transcriptor IA · Reconocimiento automático de voz · Revise el texto antes de usarlo oficialmente.</p>',
    unsafe_allow_html=True,
)
