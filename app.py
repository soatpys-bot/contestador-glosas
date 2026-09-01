import base64
import copy
import io
import json
import os
import re
import shutil
import subprocess
import tempfile
from datetime import date
from pathlib import Path

import streamlit as st
from pypdf import PdfReader, PdfWriter
from docx import Document
from docx.shared import Pt

try:
    import fitz  # PyMuPDF
except Exception:
    fitz = None

try:
    from PIL import Image
except Exception:
    Image = None

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="Contestador Inteligente de Glosas", page_icon="📄", layout="wide")
BASE_DIR = Path(__file__).parent
NORM_PATH = BASE_DIR / "normativa" / "Resolucion_2284_2023.pdf"
TEMPLATE_PATH = BASE_DIR / "PLANTILLA OFICIAL GLOSAS.docx"

DEFAULT_CODE = "RE9602"
CODE_OPTIONS = {
    "RE9601": "Devolución injustificada al 100%.",
    "RE9602": "Glosa injustificada al 100%.",
    "RE9701": "Glosa/devolución totalmente aceptada.",
    "RE9801": "Glosa parcialmente aceptada y subsanada parcialmente.",
    "RE9901": "Glosa no aceptada y subsanada en su totalidad.",
}

# La descripción de los códigos es solo para consulta en pantalla. Nunca se imprime en el PDF.


def pdf_text(data: bytes) -> str:
    return "\n".join((p.extract_text() or "") for p in PdfReader(io.BytesIO(data)).pages)


def pdf_pages_as_images(data: bytes, max_pages: int = 5):
    if fitz is None:
        return []
    doc = fitz.open(stream=data, filetype="pdf")
    out = []
    try:
        for i in range(min(len(doc), max_pages)):
            pix = doc.load_page(i).get_pixmap(matrix=fitz.Matrix(1.6, 1.6), alpha=False)
            out.append(pix.tobytes("png"))
    finally:
        doc.close()
    return out


def money_number(value) -> float:
    s = re.sub(r"[^\d,\.\-]", "", str(value or "").strip())
    if not s:
        return 0.0
    if "." in s and "," in s:
        s = s.replace(".", "").replace(",", ".") if s.rfind(",") > s.rfind(".") else s.replace(",", "")
    elif "," in s:
        parts = s.split(",")
        s = "".join(parts) if len(parts[-1]) == 3 else s.replace(",", ".")
    elif "." in s:
        parts = s.split(".")
        s = "".join(parts) if len(parts[-1]) == 3 else s
    try:
        return float(s)
    except Exception:
        return 0.0


def money_display(value) -> str:
    return f"{money_number(value):,.0f}".replace(",", ".")


def first_match(patterns, text):
    for pattern in patterns:
        m = re.search(pattern, text or "", re.I | re.M)
        if m:
            return m.group(1).strip()
    return ""


def extract_basic(text):
    # Solo extrae factura/valor. Deliberadamente NO detecta códigos de glosa.
    factura = first_match([
        r"(?:N[.°ºo]?\s*(?:DE\s*)?FACTURA|N[ÚU]MERO\s+DE\s+FACTURA|FACTURA)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-_\/]+)",
        r"(?:FACTURA\s+DE\s+VENTA)\s*[:#\-]?\s*([A-Z0-9][A-Z0-9\-_\/]+)",
    ], text)
    valor = first_match([
        r"(?:VALOR\s+(?:TOTAL\s+)?(?:DE\s+)?GLOSA|VALOR\s+GLOSADO|TOTAL\s+GLOSADO|VALOR\s+OBJETADO)\s*[:$]?\s*\$?\s*([\d\.,]+)",
        r"(?:GLOSA|OBJECI[ÓO]N)\s*(?:TOTAL)?\s*[:$]?\s*\$?\s*([\d\.,]+)",
    ], text)
    return {"factura": factura, "valor_glosado": valor}


@st.cache_data(show_spinner=False)
def load_normativa():
    return pdf_text(NORM_PATH.read_bytes()) if NORM_PATH.exists() else ""


def norm_snippets(norm_text, query, max_chars=14000):
    if not norm_text:
        return ""
    terms = [t.lower() for t in re.findall(r"[a-záéíóúñ0-9]{4,}", query or "") if t.lower() not in {"para", "como", "esta", "este", "glosa", "valor", "factura"}]
    lines = norm_text.splitlines()
    hits = []
    for i, line in enumerate(lines):
        score = sum(1 for t in terms if t in line.lower())
        if score:
            hits.extend(lines[max(0, i-3):min(len(lines), i+8)])
            if len("\n".join(hits)) >= max_chars:
                break
    return "\n".join(dict.fromkeys(hits))[:max_chars]


def ai_analyze(glosa_text, support_text, norm_text, api_key, model, images):
    empty = {
        "entidad": "",
        "factura": "",
        "valor_glosado": "",
        "observaciones": [],
        "resumen": "",
        "justificacion": "",
        "soportes_revisar": "",
        "fundamento_normativo": "",
    }
    if not api_key or OpenAI is None:
        return empty

    client = OpenAI(api_key=api_key)
    norm = norm_snippets(norm_text, glosa_text[:9000])
    system = """Eres un asistente especializado en auditoría de cuentas médicas en Colombia. Analiza una glosa y sus soportes. NO inventes hechos, diagnósticos, procedimientos, resultados, fechas, valores, soportes ni evidencias. NO detectes, asignas ni inventes códigos de glosa o códigos de respuesta. La respuesta debe ser una sola argumentación general que abarque todas las glosas/objeciones reales encontradas en la factura. Extrae la entidad responsable de pago si aparece claramente en texto, membrete o logo. Si no hay certeza, devuelve entidad vacía. Extrae número de factura y valor glosado solo si están sustentados. Identifica todas las observaciones reales y sus valores cuando existan. Puedes construir argumentos jurídicos, técnicos y administrativos con la normativa suministrada aun sin soporte adicional, pero debes distinguir lo que está demostrado de lo que debe revisarse. Si se suministran soportes, úsalos solo cuando su contenido demuestre algo. Nunca afirmes que existe un soporte que no fue suministrado."""
    prompt = f"""TEXTO DE LA GLOSA:\n{glosa_text[:50000]}\n\nSOPORTES ADICIONALES:\n{support_text[:50000]}\n\nEXTRACTO DE NORMATIVA:\n{norm}\n\nDevuelve SOLO JSON válido con esta estructura exacta:\n{{\"entidad\":\"\",\"factura\":\"\",\"valor_glosado\":\"\",\"observaciones\":[{{\"descripcion\":\"\",\"valor\":\"\"}}],\"resumen\":\"\",\"justificacion\":\"\",\"soportes_revisar\":\"\",\"fundamento_normativo\":\"\"}}\nNo incluyas ningún código."""
    content = [{"type": "text", "text": prompt}]
    for img in images[:5]:
        content.append({"type": "image_url", "image_url": {"url": "data:image/png;base64," + base64.b64encode(img).decode()}})
    try:
        r = client.chat.completions.create(
            model=model,
            messages=[{"role": "system", "content": system}, {"role": "user", "content": content}],
            temperature=0.1,
            response_format={"type": "json_object"},
        )
        data = json.loads(r.choices[0].message.content)
        return data if isinstance(data, dict) else empty
    except Exception as exc:
        empty["soportes_revisar"] = f"No fue posible completar el análisis automático: {exc}"
        return empty


def read_docx_text(data):
    try:
        d = Document(io.BytesIO(data))
        parts = [p.text for p in d.paragraphs]
        for t in d.tables:
            for row in t.rows:
                parts.append(" | ".join(c.text for c in row.cells))
        return "\n".join(parts)
    except Exception:
        return ""


def support_to_pdf(data: bytes, name: str):
    low = name.lower()
    if low.endswith(".pdf"):
        return data
    if low.endswith(".docx"):
        soffice = shutil.which("soffice") or shutil.which("libreoffice")
        if not soffice:
            raise RuntimeError("No se encontró LibreOffice para convertir un DOCX a PDF.")
        with tempfile.TemporaryDirectory() as tmp:
            p = Path(tmp)
            src = p / Path(name).name
            src.write_bytes(data)
            r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(p), str(src)], capture_output=True, text=True, timeout=120)
            if r.returncode != 0:
                raise RuntimeError(r.stderr or r.stdout or "Falló la conversión de DOCX a PDF.")
            out = p / (src.stem + ".pdf")
            if not out.exists():
                raise RuntimeError("No se generó el PDF del soporte DOCX.")
            return out.read_bytes()
    if low.endswith((".png", ".jpg", ".jpeg")):
        if Image is None:
            raise RuntimeError("No se encontró Pillow para convertir imágenes a PDF.")
        img = Image.open(io.BytesIO(data)).convert("RGB")
        out = io.BytesIO()
        img.save(out, format="PDF")
        return out.getvalue()
    raise RuntimeError(f"Formato de soporte no compatible: {name}")


def merge_pdfs(main_pdf, attachments):
    writer = PdfWriter()
    for data in [main_pdf] + attachments:
        for page in PdfReader(io.BytesIO(data)).pages:
            writer.add_page(page)
    out = io.BytesIO()
    writer.write(out)
    return out.getvalue()


def all_paragraphs(parent):
    if hasattr(parent, "paragraphs"):
        for p in parent.paragraphs:
            yield p
    if hasattr(parent, "tables"):
        for table in parent.tables:
            for row in table.rows:
                for cell in row.cells:
                    yield from all_paragraphs(cell)


def replace_placeholder_everywhere(doc, placeholder, value):
    value = str(value or "")
    for p in all_paragraphs(doc):
        if placeholder in p.text:
            # Conservar el formato general del documento; el placeholder se sustituye en el párrafo.
            for r in p.runs:
                r.text = r.text.replace(placeholder, value)
            if placeholder in p.text:
                p.text = p.text.replace(placeholder, value)


def set_cell_text(cell, text, size=8.5):
    if not cell.paragraphs:
        cell.add_paragraph()
    p = cell.paragraphs[0]
    p.text = str(text or "")
    for r in p.runs:
        r.font.size = Pt(size)


def locate_response_table(doc):
    for table in doc.tables:
        header = " ".join(c.text.upper() for row in table.rows[:2] for c in row.cells)
        if "VALOR ACEPTADO" in header and "VALOR OBJETADO" in header and ("COD. RESP" in header or "COD. RESPUESTA" in header):
            return table
    return None


def prepare_template(meta, row):
    if not TEMPLATE_PATH.exists():
        raise RuntimeError("No existe PLANTILLA OFICIAL GLOSAS.docx en el proyecto.")
    doc = Document(str(TEMPLATE_PATH))

    # La plantilla oficial contiene estos marcadores y se conserva su estructura.
    replace_placeholder_everywhere(doc, "{{fecha}}", meta["fecha"])
    replace_placeholder_everywhere(doc, "{{factura}}", meta["factura"])
    replace_placeholder_everywhere(doc, "{{valor_glosado}}", meta["valor_glosado"])
    replace_placeholder_everywhere(doc, "{{eps o aseguradora}}", meta["entidad"])
    replace_placeholder_everywhere(doc, "{{codigo_respuesta}}", row["codigo"])
    replace_placeholder_everywhere(doc, "{{valor_aceptado}}", row["valor_aceptado"])
    replace_placeholder_everywhere(doc, "{{valor_objetado}}", row["valor_objetado"])
    replace_placeholder_everywhere(doc, "{{justificacion_ia}}", row["justificacion"])

    table = locate_response_table(doc)
    if table is not None and len(table.rows) >= 2:
        # Si el texto del placeholder quedó fragmentado por Word, escribimos directamente las celdas.
        set_cell_text(table.rows[1].cells[0], row["codigo"])
        set_cell_text(table.rows[1].cells[1], row["valor_aceptado"])
        set_cell_text(table.rows[1].cells[2], row["valor_objetado"])
        set_cell_text(table.rows[1].cells[3], row["justificacion"], 8)
    return doc


def docx_to_pdf(docx_bytes):
    soffice = shutil.which("soffice") or shutil.which("libreoffice")
    if not soffice:
        raise RuntimeError("No se encontró LibreOffice. Verifica packages.txt en Streamlit Cloud.")
    with tempfile.TemporaryDirectory() as tmp:
        p = Path(tmp)
        d = p / "respuesta_institucional.docx"
        d.write_bytes(docx_bytes)
        r = subprocess.run([soffice, "--headless", "--convert-to", "pdf", "--outdir", str(p), str(d)], capture_output=True, text=True, timeout=120)
        if r.returncode != 0:
            raise RuntimeError(r.stderr or r.stdout or "Falló la conversión a PDF.")
        out = p / "respuesta_institucional.pdf"
        if not out.exists():
            raise RuntimeError("No se generó el PDF institucional.")
        return out.read_bytes()


def render_pdf_preview(pdf_bytes):
    if fitz is None:
        return []
    doc = fitz.open(stream=pdf_bytes, filetype="pdf")
    imgs = []
    try:
        for i in range(min(len(doc), 4)):
            pix = doc.load_page(i).get_pixmap(matrix=fitz.Matrix(1.35, 1.35), alpha=False)
            imgs.append(pix.tobytes("png"))
    finally:
        doc.close()
    return imgs


st.title("📄 Contestador Inteligente de Glosas")
st.caption("Lectura de glosa + revisión humana + plantilla institucional + PDF consolidado")

with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input("Clave de IA", type="password", value=st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", "")))
    model = st.selectbox("Modelo", ["gpt-4.1-mini", "gpt-4.1"], index=0)
    st.divider()
    st.markdown("**Reglas de la aplicación**")
    st.write("• No detecta ni inventa códigos de glosa.")
    st.write("• RE9602 queda seleccionado por defecto.")
    st.write("• La descripción del código solo se muestra en pantalla.")
    st.write("• La fecha se genera con el día de elaboración.")
    st.write("• La respuesta puede generarse aunque no haya soportes adicionales.")

st.subheader("1. Cargar glosa")
glosa = st.file_uploader("Documento de glosa (PDF)", type=["pdf"], key="glosa")

st.subheader("2. Cargar soportes adicionales (opcional y múltiples)")
supports = st.file_uploader(
    "Puedes cargar PDF, DOCX o imágenes. Se incorporarán al PDF final como anexos.",
    type=["pdf", "docx", "png", "jpg", "jpeg"],
    accept_multiple_files=True,
    key="supports",
)

if glosa:
    glosa_bytes = glosa.getvalue()
    glosa_text = pdf_text(glosa_bytes)
    basic = extract_basic(glosa_text)

    support_text_parts = []
    support_pdf_bytes = []
    support_errors = []
    for f in supports or []:
        data = f.getvalue()
        low = f.name.lower()
        if low.endswith(".pdf"):
            support_text_parts.append(f"--- {f.name} ---\n{pdf_text(data)}")
        elif low.endswith(".docx"):
            support_text_parts.append(f"--- {f.name} ---\n{read_docx_text(data)}")
        else:
            support_text_parts.append(f"--- {f.name} ---\n[Imagen aportada como soporte visual]")
        try:
            support_pdf_bytes.append(support_to_pdf(data, f.name))
        except Exception as exc:
            support_errors.append(f"{f.name}: {exc}")

    support_text = "\n".join(support_text_parts)

    if st.button("🔎 Analizar glosa", type="primary"):
        with st.spinner("Analizando entidad, factura, valores, observaciones y normativa..."):
            analysis = ai_analyze(glosa_text, support_text, load_normativa(), api_key, model, pdf_pages_as_images(glosa_bytes))
            st.session_state.analysis = analysis
            st.session_state.meta = {
                "fecha": date.today().strftime("%d/%m/%Y"),
                "entidad": analysis.get("entidad") or "",
                "factura": analysis.get("factura") or basic.get("factura") or "",
                "valor_glosado": analysis.get("valor_glosado") or basic.get("valor_glosado") or "",
            }
            total = money_number(st.session_state.meta["valor_glosado"])
            st.session_state.row = {
                "codigo": DEFAULT_CODE,
                "valor_aceptado": money_display(0),
                "valor_objetado": money_display(total) if total else "",
                "justificacion": analysis.get("justificacion") or "",
            }
            st.session_state.preview_pdf = None

    if "analysis" in st.session_state:
        meta = st.session_state.meta
        row = st.session_state.row
        analysis = st.session_state.analysis

        st.subheader("3. Revisar información identificada")
        c1, c2 = st.columns(2)
        meta["entidad"] = c1.text_input("Entidad a la que se responderá", meta["entidad"], help="Se intenta leer del texto, membrete o logo. Si no hay certeza, escríbela aquí.")
        meta["factura"] = c2.text_input("N. de factura", meta["factura"])
        c3, c4 = st.columns(2)
        c3.text_input("Fecha de respuesta", meta["fecha"], disabled=True)
        meta["valor_glosado"] = c4.text_input("Valor glosado / total de la objeción", meta["valor_glosado"])
        if not meta["entidad"]:
            st.warning("No se identificó con certeza la entidad. Debes escribirla antes de generar el PDF.")
        if not meta["factura"]:
            st.warning("No se identificó con certeza el número de factura. Debes escribirlo antes de generar el PDF.")
        if not meta["valor_glosado"]:
            st.warning("No se identificó con certeza el valor glosado. Revisa el documento y escríbelo antes de generar el PDF.")

        st.subheader("4. Código de respuesta y valores")
        codes = list(CODE_OPTIONS)
        row["codigo"] = st.selectbox(
            "Código específico de respuesta",
            codes,
            index=codes.index(row["codigo"]) if row["codigo"] in codes else codes.index(DEFAULT_CODE),
        )
        st.info(f"**{row['codigo']}** — {CODE_OPTIONS[row['codigo']]}")
        st.caption("Esta descripción es solo de consulta y NO aparecerá en el PDF institucional.")
        total = money_number(meta["valor_glosado"])
        c1, c2 = st.columns(2)
        row["valor_aceptado"] = c1.text_input("Valor aceptado", row["valor_aceptado"])
        row["valor_objetado"] = c2.text_input("Valor objetado", row["valor_objetado"] or (money_display(total) if total else ""))

        st.subheader("5. Argumentación general de todas las glosas")
        st.caption("La aplicación prepara una sola respuesta general que integra las observaciones encontradas. Puedes editarla antes de generar el PDF.")
        row["justificacion"] = st.text_area("Justificación que irá al documento", row["justificacion"], height=300)

        st.subheader("6. Observaciones encontradas")
        obs = analysis.get("observaciones") or []
        if obs:
            for i, o in enumerate(obs, 1):
                desc = o.get("descripcion", "") if isinstance(o, dict) else str(o)
                val = o.get("valor", "") if isinstance(o, dict) else ""
                st.write(f"**{i}.** {desc}" + (f" — Valor: {val}" if val else ""))
        else:
            st.info("No se identificaron observaciones estructuradas. Revisa el texto original de la glosa.")

        if analysis.get("fundamento_normativo"):
            with st.expander("Ver fundamento normativo usado como apoyo"):
                st.write(analysis["fundamento_normativo"])
        if analysis.get("soportes_revisar"):
            with st.expander("Ver advertencias / soportes por revisar"):
                st.write(analysis["soportes_revisar"])

        st.subheader("7. Soportes anexos")
        if supports:
            for f in supports:
                st.write("✅ " + f.name)
            st.caption("Los soportes se anexarán después de la respuesta institucional, sin alterar su contenido.")
        else:
            st.info("No se cargaron soportes adicionales. La respuesta puede generarse igualmente.")
        for err in support_errors:
            st.warning(err)

        st.subheader("8. Revisar todo antes del PDF")
        st.write("Antes de generar el PDF puedes revisar y modificar todos los campos editables. La fecha es automática.")
        if meta["entidad"] and meta["factura"] and meta["valor_glosado"] and row["justificacion"].strip():
            if st.button("👁️ Generar previsualización institucional", type="secondary"):
                try:
                    doc = prepare_template(meta, row)
                    b = io.BytesIO()
                    doc.save(b)
                    preview_pdf = docx_to_pdf(b.getvalue())
                    st.session_state.preview_pdf = preview_pdf
                    st.success("Previsualización lista. Revisa visualmente el documento antes de descargarlo.")
                except Exception as exc:
                    st.error(f"No fue posible preparar la previsualización: {exc}")
        else:
            st.warning("Completa entidad, factura, valor glosado y justificación para habilitar la previsualización.")

        if st.session_state.get("preview_pdf"):
            st.markdown("### Previsualización")
            for img in render_pdf_preview(st.session_state.preview_pdf):
                st.image(img, use_container_width=True)

            if st.button("📄 GENERAR PDF FINAL", type="primary"):
                try:
                    doc = prepare_template(meta, row)
                    b = io.BytesIO()
                    doc.save(b)
                    main_pdf = docx_to_pdf(b.getvalue())
                    final_pdf = merge_pdfs(main_pdf, support_pdf_bytes)
                    safe_invoice = re.sub(r"[^A-Za-z0-9_-]+", "_", meta["factura"])
                    st.download_button(
                        "⬇️ Descargar PDF final",
                        final_pdf,
                        f"Respuesta_Glosas_{safe_invoice or 'factura'}.pdf",
                        "application/pdf",
                    )
                    st.success("PDF final generado correctamente, con la respuesta institucional y los soportes anexos.")
                except Exception as exc:
                    st.error(f"No fue posible generar el PDF final: {exc}")

        with st.expander("Ver texto original de la glosa"):
            st.text_area("Contenido extraído", glosa_text, height=280)
else:
    st.markdown("""
### Flujo de trabajo
1. Carga la glosa en PDF.
2. Carga soportes adicionales si los tienes; pueden ser varios.
3. Pulsa **Analizar glosa**.
4. Revisa entidad, factura, valor y fecha.
5. Selecciona el código de respuesta; **RE9602** aparece por defecto.
6. Consulta el significado del código en pantalla; esa descripción no se imprime.
7. Revisa/edita la única argumentación general que integrará todas las glosas.
8. Genera la **previsualización institucional**.
9. Si todo está correcto, genera el **PDF final** con los soportes anexos.

La aplicación **no detecta ni inventa códigos de glosa** y no inventa hechos o soportes.
""")
