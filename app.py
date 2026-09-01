import io
import os
import re
import tempfile
from pathlib import Path
from datetime import date

import streamlit as st
from pypdf import PdfReader, PdfWriter
from docx import Document
from PIL import Image
from reportlab.lib.pagesizes import A4
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.enums import TA_LEFT, TA_CENTER
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, PageBreak
from reportlab.lib.units import cm

from google import genai


# ============================================================
# CONFIGURACIÓN
# ============================================================

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "PLANTILLA OFICIAL GLOSAS.docx"
NORM_DIR = APP_DIR / "normativa"

DEFAULT_MODEL = "gemini-2.5-flash"

st.set_page_config(
    page_title="Contestador Inteligente de Glosas",
    page_icon="📄",
    layout="wide",
)

st.title("📄 Contestador Inteligente de Glosas")
st.caption(
    "Lectura de glosa + análisis con Gemini + revisión humana + "
    "respuesta institucional + PDF consolidado"
)


# ============================================================
# FUNCIONES DE TEXTO / ARCHIVOS
# ============================================================

def clean(value):
    return "" if value is None else str(value).strip()


def pdf_text(data: bytes) -> str:
    reader = PdfReader(io.BytesIO(data))
    pages = []
    for page in reader.pages:
        try:
            pages.append(page.extract_text() or "")
        except Exception:
            pages.append("")
    return "\n".join(pages).strip()


def docx_text(data: bytes) -> str:
    doc = Document(io.BytesIO(data))
    out = []

    for p in doc.paragraphs:
        text = clean(p.text)
        if text:
            out.append(text)

    for table in doc.tables:
        for row in table.rows:
            cells = [clean(c.text) for c in row.cells]
            if any(cells):
                out.append(" | ".join(cells))

    return "\n".join(out).strip()


def image_text_note(data: bytes, filename: str) -> str:
    # La imagen se conserva como soporte. Para no depender de OCR adicional,
    # Gemini recibe la imagen directamente cuando se hace el análisis.
    try:
        img = Image.open(io.BytesIO(data))
        return f"[Imagen de soporte: {filename} - {img.width}x{img.height}px]"
    except Exception:
        return f"[Imagen de soporte: {filename}]"


def extract_uploaded_file(uploaded):
    data = uploaded.getvalue()
    name = uploaded.name.lower()

    if name.endswith(".pdf"):
        return pdf_text(data), data, "application/pdf"

    if name.endswith(".docx"):
        return docx_text(data), data, (
            "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
        )

    if name.endswith((".png", ".jpg", ".jpeg")):
        return image_text_note(data, uploaded.name), data, uploaded.type or "image/jpeg"

    return "", data, uploaded.type or "application/octet-stream"


def read_normative_folder():
    if not NORM_DIR.exists():
        return ""

    chunks = []

    for path in sorted(NORM_DIR.rglob("*")):
        if not path.is_file():
            continue

        try:
            suffix = path.suffix.lower()
            data = path.read_bytes()

            if suffix == ".pdf":
                text = pdf_text(data)
            elif suffix == ".docx":
                text = docx_text(data)
            elif suffix in {".txt", ".md"}:
                text = data.decode("utf-8", errors="ignore")
            else:
                continue

            text = clean(text)
            if text:
                chunks.append(f"\n===== NORMA: {path.name} =====\n{text}")

        except Exception:
            continue

    return "\n".join(chunks).strip()


# ============================================================
# GEMINI
# ============================================================

def get_api_key():
    # 1. Streamlit Secrets
    try:
        key = st.secrets.get("GEMINI_API_KEY", "")
        if key:
            return clean(key)
    except Exception:
        pass

    # 2. Variable de entorno
    key = os.getenv("GEMINI_API_KEY", "")
    if key:
        return clean(key)

    # 3. Clave introducida en pantalla
    return clean(st.session_state.get("gemini_key", ""))


def get_client():
    key = get_api_key()
    if not key:
        return None
    return genai.Client(api_key=key)


def gemini_generate(prompt, model=DEFAULT_MODEL):
    client = get_client()
    if client is None:
        raise RuntimeError(
            "No hay una clave Gemini configurada. "
            "Introduce tu Gemini API Key en la barra lateral."
        )

    response = client.models.generate_content(
        model=model,
        contents=prompt,
    )

    text = getattr(response, "text", None)
    if not text:
        raise RuntimeError("Gemini no devolvió texto.")
    return text.strip()


# ============================================================
# PROMPTS
# ============================================================

SYSTEM_RULES = """
Eres el motor de análisis de una aplicación institucional para contestar glosas.

REGLAS OBLIGATORIAS:
1. No inventes códigos de glosa.
2. No inventes hechos, soportes, valores, fechas, servicios ni diagnósticos.
3. Usa exclusivamente la información entregada en la glosa, los soportes y la normativa disponible.
4. Si un dato no aparece, escribe claramente: "No consta en la documentación suministrada".
5. No afirmes que un soporte existe si no fue suministrado.
6. Distingue hechos documentados de inferencias.
7. La decisión final siempre queda para revisión humana.
8. La respuesta debe ser técnica, clara, prudente y apta para revisión institucional.
9. Si hay una contradicción entre documentos, señálala en vez de resolverla inventando información.
"""


def build_analysis_prompt(glosa_text, support_text, norm_text):
    return f"""
{SYSTEM_RULES}

Analiza la siguiente glosa.

===== GLOSA =====
{glosa_text[:90000]}

===== SOPORTES ADICIONALES =====
{support_text[:90000] if support_text else "No se suministraron soportes adicionales."}

===== NORMATIVA DISPONIBLE =====
{norm_text[:120000] if norm_text else "No hay normativa local cargada."}

Entrega el análisis con esta estructura exacta:

1. RESUMEN DE LA GLOSA
2. CÓDIGO DE GLOSA
3. VALOR OBJETADO
4. HECHOS Y DATOS DOCUMENTADOS
5. HALLAZGOS
6. SOPORTES QUE RESPALDAN LA RESPUESTA
7. NORMATIVA APLICABLE EN LOS DOCUMENTOS SUMINISTRADOS
8. CONCLUSIÓN TÉCNICA
9. RECOMENDACIÓN PARA REVISIÓN HUMANA

Cuando un campo no esté disponible, indícalo expresamente.
"""


def build_response_prompt(analysis, human_review, glosa_text, norm_text):
    return f"""
{SYSTEM_RULES}

Con base en el análisis y en la revisión humana, redacta una respuesta institucional
a la glosa.

===== ANÁLISIS =====
{analysis[:90000]}

===== REVISIÓN HUMANA =====
{human_review[:30000] if human_review else "No se agregó revisión humana."}

===== GLOSA ORIGINAL =====
{glosa_text[:60000]}

===== NORMATIVA =====
{norm_text[:80000] if norm_text else "No hay normativa local disponible."}

Redacta una respuesta formal, concreta y sustentada.

Debe contener:
- Asunto
- Identificación de la glosa (solo si consta)
- Respuesta institucional
- Sustento técnico
- Sustento normativo (solo si consta en la documentación)
- Relación de soportes
- Solicitud/conclusión final

NO inventes números, códigos, valores, fechas o artículos.
Si algo no consta, dilo expresamente.
"""


# ============================================================
# DOCUMENTO Y PDF
# ============================================================

def add_response_to_template(response_text, analysis_text):
    """
    Usa la plantilla institucional como base.
    No modifica ni borra el contenido existente de la plantilla.
    Agrega la respuesta al final para evitar destruir ejemplos o formato
    institucional que la organización haya dejado en el DOCX.
    """
    if TEMPLATE_PATH.exists():
        doc = Document(str(TEMPLATE_PATH))
    else:
        doc = Document()
        doc.add_heading("Respuesta institucional a glosa", level=1)

    doc.add_paragraph("")
    doc.add_heading("RESPUESTA GENERADA PARA REVISIÓN", level=1)

    for block in response_text.split("\n"):
        line = block.strip()
        if line:
            doc.add_paragraph(line)

    doc.add_page_break()
    doc.add_heading("ANÁLISIS TÉCNICO PARA REVISIÓN HUMANA", level=1)

    for block in analysis_text.split("\n"):
        line = block.strip()
        if line:
            doc.add_paragraph(line)

    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".docx")
    tmp.close()
    doc.save(tmp.name)

    return Path(tmp.name)


def make_response_pdf(response_text, analysis_text):
    tmp = tempfile.NamedTemporaryFile(delete=False, suffix=".pdf")
    tmp.close()

    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="InstitutionTitle",
            parent=styles["Title"],
            alignment=TA_CENTER,
            fontSize=15,
            leading=18,
            spaceAfter=12,
        )
    )
    styles.add(
        ParagraphStyle(
            name="BodySafe",
            parent=styles["BodyText"],
            fontSize=9.5,
            leading=13,
            alignment=TA_LEFT,
            spaceAfter=5,
        )
    )

    doc = SimpleDocTemplate(
        tmp.name,
        pagesize=A4,
        rightMargin=1.7 * cm,
        leftMargin=1.7 * cm,
        topMargin=1.7 * cm,
        bottomMargin=1.7 * cm,
    )

    story = [
        Paragraph("RESPUESTA INSTITUCIONAL A GLOSA", styles["InstitutionTitle"]),
        Paragraph(f"Fecha: {date.today().isoformat()}", styles["BodySafe"]),
        Spacer(1, 8),
    ]

    def safe_paragraph(text):
        text = clean(text)
        if not text:
            return None
        text = (
            text.replace("&", "&amp;")
            .replace("<", "&lt;")
            .replace(">", "&gt;")
        )
        return Paragraph(text, styles["BodySafe"])

    for line in response_text.splitlines():
        p = safe_paragraph(line)
        if p:
            story.append(p)

    story.append(PageBreak())
    story.append(
        Paragraph("ANÁLISIS TÉCNICO PARA REVISIÓN HUMANA", styles["InstitutionTitle"])
    )

    for line in analysis_text.splitlines():
        p = safe_paragraph(line)
        if p:
            story.append(p)

    doc.build(story)
    return Path(tmp.name)


def merge_pdfs(pdf_bytes_list):
    writer = PdfWriter()

    for data in pdf_bytes_list:
        reader = PdfReader(io.BytesIO(data))
        for page in reader.pages:
            writer.add_page(page)

    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()


# ============================================================
# BARRA LATERAL
# ============================================================

with st.sidebar:
    st.header("Configuración")

    entered_key = st.text_input(
        "Clave de Gemini",
        value="",
        type="password",
        help="Pega aquí la Gemini API Key que creaste en Google AI Studio.",
    )

    if entered_key:
        st.session_state["gemini_key"] = entered_key

    model = st.selectbox(
        "Modelo Gemini",
        [
            "gemini-2.5-flash",
            "gemini-2.5-flash-lite",
            "gemini-3.7-flash",
        ],
        index=0,
    )

    st.divider()

    st.subheader("Reglas de la aplicación")
    st.write("• No detecta ni inventa códigos de glosa.")
    st.write("• Los códigos y valores solo se toman de la documentación.")
    st.write("• La normativa se usa únicamente si está cargada.")
    st.write("• La respuesta generada queda para revisión humana.")
    st.write("• La plantilla institucional se conserva como base.")

    if get_api_key():
        st.success("Clave Gemini configurada.")
    else:
        st.warning("Falta configurar la clave Gemini.")


# ============================================================
# CARGA DE DOCUMENTOS
# ============================================================

st.header("1. Cargar glosa")

glosa_file = st.file_uploader(
    "Documento de glosa (PDF)",
    type=["pdf"],
    accept_multiple_files=False,
)

st.header("2. Cargar soportes adicionales (opcional y múltiples)")

support_files = st.file_uploader(
    "Puedes cargar PDF, DOCX o imágenes.",
    type=["pdf", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

norm_text = read_normative_folder()

if norm_text:
    st.success("Normativa local encontrada y disponible para el análisis.")
else:
    st.info(
        "No se encontró contenido legible en la carpeta 'normativa'. "
        "La aplicación puede funcionar sin ella, pero no inventará sustento normativo."
    )


# ============================================================
# FLUJO
# ============================================================

st.header("3. Flujo de trabajo")

if glosa_file is None:
    st.info("Carga primero el PDF de la glosa.")
    st.stop()

glosa_text, glosa_bytes, glosa_mime = extract_uploaded_file(glosa_file)

if not glosa_text:
    st.error(
        "No fue posible extraer texto de la glosa. "
        "Verifica que el PDF contenga texto seleccionable."
    )
    st.stop()

support_text_parts = []
support_payloads = []

for uploaded in support_files or []:
    text, data, mime = extract_uploaded_file(uploaded)

    if text:
        support_text_parts.append(
            f"===== SOPORTE: {uploaded.name} =====\n{text}"
        )

    # Las imágenes se mantienen como datos para mostrar que fueron cargadas.
    support_payloads.append(
        {
            "name": uploaded.name,
            "mime": mime,
            "data": data,
        }
    )

support_text = "\n\n".join(support_text_parts)


st.subheader("3.1 Datos extraídos de la glosa")
with st.expander("Ver texto de la glosa", expanded=False):
    st.text_area(
        "Texto",
        glosa_text,
        height=280,
        label_visibility="collapsed",
    )


if st.button("🔎 Analizar glosa con Gemini", type="primary"):
    try:
        with st.spinner("Gemini está analizando la glosa y los soportes..."):
            prompt = build_analysis_prompt(
                glosa_text,
                support_text,
                norm_text,
            )
            analysis = gemini_generate(prompt, model=model)

        st.session_state["analysis"] = analysis
        st.session_state["glosa_text"] = glosa_text
        st.session_state["support_text"] = support_text
        st.success("Análisis generado correctamente.")

    except Exception as exc:
        st.error(f"No fue posible ejecutar Gemini: {exc}")


analysis = st.session_state.get("analysis", "")

if analysis:
    st.subheader("4. Análisis de Gemini")
    st.markdown(analysis)

    st.subheader("5. Revisión humana")
    human_review = st.text_area(
        "Agrega aquí las correcciones, aclaraciones o decisión del revisor.",
        height=220,
        placeholder=(
            "Ejemplo: confirmar valor, aclarar soporte, corregir interpretación, "
            "indicar si se acepta o se objeta la glosa..."
        ),
    )

    if st.button("✍️ Generar respuesta institucional"):
        try:
            with st.spinner("Redactando respuesta institucional..."):
                response = gemini_generate(
                    build_response_prompt(
                        analysis,
                        human_review,
                        glosa_text,
                        norm_text,
                    ),
                    model=model,
                )

            st.session_state["response"] = response
            st.success("Respuesta institucional generada.")

        except Exception as exc:
            st.error(f"No fue posible generar la respuesta: {exc}")


response = st.session_state.get("response", "")

if response:
    st.subheader("6. Respuesta institucional")
    st.markdown(response)

    st.subheader("7. Generar documentos")

    col1, col2 = st.columns(2)

    with col1:
        try:
            docx_path = add_response_to_template(response, analysis)
            docx_data = docx_path.read_bytes()

            st.download_button(
                "⬇️ Descargar DOCX",
                data=docx_data,
                file_name="respuesta_glosa_GEMINI.docx",
                mime=(
                    "application/vnd.openxmlformats-officedocument."
                    "wordprocessingml.document"
                ),
            )
        except Exception as exc:
            st.error(f"No fue posible crear el DOCX: {exc}")

    with col2:
        try:
            response_pdf_path = make_response_pdf(response, analysis)
            response_pdf = response_pdf_path.read_bytes()

            consolidated = merge_pdfs(
                [glosa_bytes, response_pdf]
            )

            st.download_button(
                "⬇️ Descargar PDF consolidado",
                data=consolidated,
                file_name="respuesta_glosa_GEMINI_CONSOLIDADO.pdf",
                mime="application/pdf",
            )
        except Exception as exc:
            st.error(f"No fue posible crear el PDF consolidado: {exc}")


st.divider()
st.caption(
    "La aplicación usa Gemini mediante la API de Google. "
    "La revisión humana sigue siendo necesaria antes de emitir una respuesta oficial."
)
