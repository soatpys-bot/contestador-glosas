import io, json, os, re, shutil, subprocess, tempfile
from datetime import date
from pathlib import Path

import streamlit as st
from pypdf import PdfReader, PdfWriter
from docx import Document
from PIL import Image
from google import genai
from google.genai import types

APP_DIR = Path(__file__).resolve().parent
TEMPLATE_PATH = APP_DIR / "PLANTILLA OFICIAL GLOSAS.docx"
NORM_DIR = APP_DIR / "normativa"

CODE_CONTEXT = {
    "RE9601": "Opción disponible para revisión humana.",
    "RE9602": "Opción predeterminada para revisión humana.",
    "RE9701": "Opción disponible para revisión humana.",
    "RE9801": "Opción disponible para revisión humana.",
    "RE9901": "Opción disponible para revisión humana.",
}

st.set_page_config(page_title="Contestador Inteligente de Glosas", page_icon="📄", layout="wide")
st.title("📄 Contestador Inteligente de Glosas")
st.caption("Gemini dentro de la aplicación · revisión humana · plantilla institucional · PDF consolidado")


def clean(value):
    return "" if value is None else str(value).strip()


def pdf_text(data):
    reader = PdfReader(io.BytesIO(data))
    return "\n".join((page.extract_text() or "") for page in reader.pages).strip()


def docx_text(data):
    doc = Document(io.BytesIO(data))
    out = [p.text for p in doc.paragraphs if p.text.strip()]
    for table in doc.tables:
        for row in table.rows:
            out.append(" | ".join(cell.text.strip() for cell in row.cells))
    return "\n".join(out).strip()


def read_any(name, data):
    ext = Path(name).suffix.lower()
    if ext == ".pdf":
        return pdf_text(data)
    if ext == ".docx":
        return docx_text(data)
    if ext == ".txt":
        return data.decode("utf-8", errors="ignore")
    return ""


@st.cache_data(show_spinner=False)
def load_normativa():
    chunks = []
    if not NORM_DIR.exists():
        return ""
    for path in sorted(NORM_DIR.rglob("*")):
        if not path.is_file() or path.suffix.lower() not in {".pdf", ".docx", ".txt"}:
            continue
        try:
            chunks.append(f"===== {path.name} =====\n{read_any(path.name, path.read_bytes())[:140000]}")
        except Exception as exc:
            chunks.append(f"===== {path.name} =====\n[No se pudo leer: {exc}]")
    return "\n\n".join(chunks)


def get_key():
    try:
        key = clean(st.secrets.get("GEMINI_API_KEY", ""))
        if key:
            return key
    except Exception:
        pass
    return clean(os.getenv("GEMINI_API_KEY", ""))


def clear_cell(cell, value, font_size=None):
    """Replace all visible cell text while retaining the template's cell formatting."""
    value = clean(value)
    if not cell.paragraphs:
        p = cell.add_paragraph()
    else:
        p = cell.paragraphs[0]
    for run in p.runs:
        run.text = ""
    if p.runs:
        p.runs[0].text = value
        if font_size:
            for run in p.runs:
                run.font.size = run.font.size or None
    else:
        run = p.add_run(value)
        if font_size:
            from docx.shared import Pt
            run.font.size = Pt(font_size)
    for extra in list(cell.paragraphs[1:]):
        parent = extra._element.getparent()
        if parent is not None:
            parent.remove(extra._element)


def replace_paragraph_text(doc, old, new):
    for paragraph in doc.paragraphs:
        if paragraph.text.strip() == old:
            for run in paragraph.runs:
                run.text = ""
            if paragraph.runs:
                paragraph.runs[0].text = clean(new)
            else:
                paragraph.add_run(clean(new))
            return True
    return False


def visible_doc_text(doc):
    parts = [p.text for p in doc.paragraphs]
    for table in doc.tables:
        for row in table.rows:
            parts.append(" | ".join(cell.text for cell in row.cells))
    return "\n".join(parts)


def fill_template(meta, justification, out_docx):
    if not TEMPLATE_PATH.exists():
        raise FileNotFoundError("No se encontró PLANTILLA OFICIAL GLOSAS.docx.")

    required = [
        "entidad", "factura", "valor_glosado", "codigo_respuesta",
        "valor_aceptado", "valor_objetado", "fecha"
    ]
    missing = [key for key in required if not clean(meta.get(key))]
    if not clean(justification):
        missing.append("justificacion")
    if missing:
        raise ValueError("Faltan campos obligatorios: " + ", ".join(missing))

    doc = Document(str(TEMPLATE_PATH))
    if len(doc.tables) != 3:
        raise ValueError("La plantilla oficial esperada debe contener exactamente 3 tablas.")

    # La plantilla suministrada contiene marcadores, no datos reales.
    t0, t1, t2 = doc.tables
    clear_cell(t0.cell(0, 2), f"FECHA\n{meta['fecha']}")
    clear_cell(t1.cell(0, 1), meta["factura"])
    clear_cell(t1.cell(1, 1), meta["valor_glosado"])
    clear_cell(t2.cell(1, 0), meta["codigo_respuesta"])
    clear_cell(t2.cell(1, 1), meta["valor_aceptado"])
    clear_cell(t2.cell(1, 2), meta["valor_objetado"])
    clear_cell(t2.cell(1, 3), justification)

    if not replace_paragraph_text(doc, "{{eps o aseguradora}}", meta["entidad"]):
        raise ValueError("No se encontró el marcador de entidad {{eps o aseguradora}} en la plantilla.")

    doc.save(str(out_docx))

    # Bloqueo de seguridad: los marcadores y valores de ejemplo no pueden pasar al PDF.
    check = Document(str(out_docx))
    visible = visible_doc_text(check)
    forbidden = [
        "{{eps o aseguradora}}", "{{fecha}}", "{{factura}}", "{{valor_glosado}}",
        "{{codigo_respuesta}}", "{{valor_aceptado}}", "{{valor_objetado}}", "{{justificacion_ia}}",
        "PREVISORA SOAT", "FEDV348712", "398.000", "RE998", "398000"
    ]
    leaked = [item for item in forbidden if item in visible]
    if leaked:
        raise ValueError("Se detectaron marcadores/datos de ejemplo en el documento final: " + ", ".join(leaked))


def libreoffice_convert(input_path, output_dir):
    last_error = ""
    for binary in ("libreoffice", "soffice"):
        if not shutil.which(binary):
            continue
        try:
            result = subprocess.run(
                [binary, "--headless", "--convert-to", "pdf", "--outdir", str(output_dir), str(input_path)],
                stdout=subprocess.PIPE, stderr=subprocess.PIPE, text=True, timeout=180,
            )
            pdf_path = Path(output_dir) / (Path(input_path).stem + ".pdf")
            if result.returncode == 0 and pdf_path.exists():
                return pdf_path
            last_error = result.stderr or result.stdout or "Conversión sin archivo de salida."
        except Exception as exc:
            last_error = str(exc)
    raise RuntimeError("No se pudo convertir la plantilla a PDF. Verifica que LibreOffice esté en packages.txt. " + clean(last_error))


def file_to_pdf(name, data, temp_dir):
    ext = Path(name).suffix.lower()
    target = Path(temp_dir) / (Path(name).stem + ".pdf")
    if ext == ".pdf":
        target.write_bytes(data)
        return target
    if ext == ".docx":
        source = Path(temp_dir) / (Path(name).stem + ".docx")
        source.write_bytes(data)
        return libreoffice_convert(source, temp_dir)
    if ext in {".png", ".jpg", ".jpeg"}:
        with Image.open(io.BytesIO(data)) as image:
            image.convert("RGB").save(target, "PDF", resolution=150.0)
        return target
    raise ValueError(f"Formato no soportado: {name}")


def merge_pdfs(paths):
    writer = PdfWriter()
    for path in paths:
        reader = PdfReader(str(path))
        for page in reader.pages:
            writer.add_page(page)
    output = io.BytesIO()
    writer.write(output)
    return output.getvalue()

def support_parts(supports):
    text_chunks = []
    media_parts = []
    for uploaded in supports or []:
        data = uploaded.getvalue()
        ext = Path(uploaded.name).suffix.lower()
        if ext in {".pdf", ".docx", ".txt"}:
            try:
                text_chunks.append(f"===== {uploaded.name} =====\n{read_any(uploaded.name, data)[:70000]}")
            except Exception as exc:
                text_chunks.append(f"===== {uploaded.name} =====\n[No se pudo leer: {exc}]")
        elif ext in {".png", ".jpg", ".jpeg"}:
            mime = "image/png" if ext == ".png" else "image/jpeg"
            media_parts.append(types.Part.from_bytes(data=data, mime_type=mime))
    return "\n\n".join(text_chunks), media_parts


def gemini_analyze(api_key, model, glosa_bytes, glosa_text, support_text, support_media, norm_text, user_entity):
    client = genai.Client(api_key=api_key)
    schema = {
        "type": "object",
        "properties": {
            "entidad": {"type": "string"},
            "factura": {"type": "string"},
            "valor_glosado": {"type": "string"},
            "argumentacion": {"type": "string"},
            "soportes_a_revisar": {"type": "array", "items": {"type": "string"}},
            "observaciones": {"type": "string"},
        },
        "required": ["entidad", "factura", "valor_glosado", "argumentacion", "soportes_a_revisar", "observaciones"],
    }
    prompt = f"""
Eres el motor de análisis de un contestador institucional de glosas médicas en Colombia.

REGLAS OBLIGATORIAS:
1. No inventes hechos clínicos, diagnósticos, procedimientos, valores, fechas, autorizaciones, soportes ni números de factura.
2. No inventes ni detectes códigos de glosa. El código de respuesta se seleccionará manualmente en la aplicación y queda RE9602 por defecto.
3. La plantilla institucional contiene marcadores de ejemplo. Jamás copies sus valores de ejemplo como datos del caso actual.
4. Produce UNA SOLA argumentación general para la factura, incluso si el documento contiene varias glosas.
5. Basa la argumentación en el PDF de la glosa, los soportes aportados y la normativa suministrada. Si un dato no aparece, déjalo vacío o indícalo en observaciones.
6. Si no hay soportes adicionales, trabaja únicamente con lo que realmente aparece en la glosa y la normativa.
7. No incluyas el significado del código de respuesta dentro de la argumentación.
8. La redacción debe ser formal, técnica, clara y prudente. No afirmes como hecho algo que no esté sustentado.
9. Identifica la entidad solo si existe evidencia en el documento (texto o imagen/membrete). Si no hay certeza, usa el dato del usuario.
10. La respuesta será revisada por una persona antes de radicarse.

ENTIDAD DADA POR EL USUARIO (puede estar vacía): {user_entity}

TEXTO EXTRAÍDO DE LA GLOSA:
{glosa_text[:90000]}

TEXTO DE SOPORTES:
{support_text[:90000]}

NORMATIVA DISPONIBLE:
{norm_text[:120000]}

Devuelve exclusivamente JSON según el esquema solicitado.
"""
    contents = [types.Part.from_bytes(data=glosa_bytes, mime_type="application/pdf"), prompt]
    contents.extend(support_media)
    response = client.models.generate_content(
        model=model,
        contents=contents,
        config=types.GenerateContentConfig(
            temperature=0.1,
            response_mime_type="application/json",
            response_schema=schema,
        ),
    )
    raw = clean(response.text)
    if not raw:
        raise RuntimeError("Gemini no devolvió una respuesta.")
    try:
        return json.loads(raw)
    except json.JSONDecodeError as exc:
        raise RuntimeError("Gemini devolvió un formato que no pudo interpretarse como JSON.") from exc


with st.sidebar:
    st.header("Configuración")
    key_manual = st.text_input("Clave Gemini (no la compartas)", type="password")
    model = st.selectbox("Modelo Gemini", ["gemini-2.5-flash-lite", "gemini-2.5-flash"], index=0)
    st.divider()
    st.markdown("### Código de respuesta")
    st.write("RE9602 queda seleccionado por defecto. Puedes cambiarlo antes de generar.")
    st.caption("El contexto de los códigos solo se muestra en pantalla y nunca se imprime en el PDF.")

api_key = clean(key_manual) or get_key()
if api_key:
    st.success("Gemini configurado.")
else:
    st.warning("Escribe la clave Gemini en la barra lateral o configura GEMINI_API_KEY en Streamlit Secrets.")

norm_text = load_normativa()
if norm_text:
    st.sidebar.success("Normativa cargada dentro de la aplicación.")
else:
    st.sidebar.error("No se encontró la carpeta normativa/ con la Resolución 2284 de 2023.")

st.subheader("1. Cargar glosa")
glosa = st.file_uploader("Documento de glosa (PDF)", type=["pdf"], accept_multiple_files=False)

st.subheader("2. Cargar soportes adicionales (opcional y múltiples)")
supports = st.file_uploader(
    "PDF, DOCX o imágenes. Se incorporarán al PDF final como anexos.",
    type=["pdf", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
)

if glosa:
    glosa_bytes = glosa.getvalue()
    try:
        glosa_text = pdf_text(glosa_bytes)
    except Exception as exc:
        st.error(f"No se pudo leer el PDF de la glosa: {exc}")
        st.stop()

    st.subheader("3. Análisis con Gemini")
    entity_hint = st.text_input("Entidad a la que se responderá (solo si no aparece claramente en la glosa)")
    with st.expander("Texto extraído de la glosa", expanded=False):
        st.text_area("Texto", glosa_text[:30000], height=260, disabled=True)
    if supports:
        with st.expander(f"Soportes cargados: {len(supports)}", expanded=False):
            for item in supports:
                st.write("•", item.name)

    if st.button("🔎 Analizar glosa con Gemini", type="primary", use_container_width=True):
        if not api_key:
            st.error("Falta la clave Gemini.")
            st.stop()
        with st.spinner("Gemini está analizando la glosa, soportes y normativa..."):
            try:
                support_text, media = support_parts(supports)
                result = gemini_analyze(api_key, model, glosa_bytes, glosa_text, support_text, media, norm_text, entity_hint)
                st.session_state["analysis"] = result
                st.session_state["support_files"] = [(item.name, item.getvalue(), item.type) for item in supports or []]
                st.success("Análisis terminado. Revisa los datos antes de generar el PDF.")
            except Exception as exc:
                st.error(f"Error al consultar Gemini: {exc}")

if "analysis" in st.session_state:
    analysis = st.session_state["analysis"]
    st.subheader("4. Revisión humana obligatoria")
    st.caption("Aquí puedes corregir cualquier dato antes de producir el documento. La aplicación no radica nada automáticamente.")

    col1, col2 = st.columns(2)
    entidad = col1.text_input("Entidad", value=clean(analysis.get("entidad")), key="final_entidad")
    factura = col2.text_input("N. de factura", value=clean(analysis.get("factura")), key="final_factura")
    valor_glosado = col1.text_input("Valor glosado", value=clean(analysis.get("valor_glosado")), key="final_valor_glosado")
    codigo = col2.selectbox("Código de respuesta", list(CODE_CONTEXT.keys()), index=list(CODE_CONTEXT).index("RE9602"), key="final_codigo")
    st.info(f"Contexto de {codigo}: {CODE_CONTEXT[codigo]} · este texto NO aparece en el PDF.")

    aceptado = st.text_input("Valor aceptado", value="0", key="final_aceptado")
    objetado = st.text_input("Valor objetado", value=valor_glosado, key="final_objetado")
    justificacion = st.text_area(
        "JUSTIFICACIÓN — una sola respuesta general para la factura",
        value=clean(analysis.get("argumentacion")), height=380, key="final_just",
    )
    st.text_area("Observaciones internas (NO se imprime en el PDF)", value=clean(analysis.get("observaciones")), height=120, key="final_obs")

    st.markdown("**Soportes que Gemini recomienda revisar/anexar**")
    st.text_area("Lista de soportes", value="\n".join(analysis.get("soportes_a_revisar", [])), height=120, key="final_support_list")

    st.subheader("5. Generar PDF institucional + anexos")
    st.caption(f"La fecha se coloca automáticamente: {date.today().strftime('%d/%m/%Y')}")
    st.warning("La plantilla oficial se conserva. Sus valores de ejemplo no se usan: antes de generar se reemplazan todos los campos variables y se bloquea el PDF si queda algún marcador o dato de ejemplo conocido.")

    if st.button("📄 Generar PDF final", type="primary", use_container_width=True):
        meta = {
            "fecha": date.today().strftime("%d/%m/%Y"),
            "entidad": entidad,
            "factura": factura,
            "valor_glosado": valor_glosado,
            "codigo_respuesta": codigo,
            "valor_aceptado": aceptado,
            "valor_objetado": objetado,
        }
        with st.spinner("Generando documento institucional y anexando soportes..."):
            try:
                with tempfile.TemporaryDirectory() as temp:
                    temp = Path(temp)
                    out_docx = temp / "Respuesta_institucional.docx"
                    fill_template(meta, justificacion, out_docx)
                    institutional_pdf = libreoffice_convert(out_docx, temp)
                    pdf_paths = [institutional_pdf]
                    for name, data, _mime in st.session_state.get("support_files", []):
                        pdf_paths.append(file_to_pdf(name, data, temp))
                    final_pdf = merge_pdfs(pdf_paths)
                st.session_state["final_pdf"] = final_pdf
                st.session_state["final_filename"] = f"Respuesta_Glosas_{factura or 'sin_factura'}.pdf"
                st.success("PDF final generado correctamente.")
            except Exception as exc:
                st.error(f"No se pudo generar el PDF final: {exc}")

    if st.session_state.get("final_pdf"):
        st.download_button(
            "⬇️ Descargar PDF final",
            data=st.session_state["final_pdf"],
            file_name=st.session_state["final_filename"],
            mime="application/pdf",
            use_container_width=True,
        )
