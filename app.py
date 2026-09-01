import io, os, re, json, textwrap
from datetime import date
from pathlib import Path

import streamlit as st
from pypdf import PdfReader
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import cm
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle, PageBreak

try:
    from openai import OpenAI
except Exception:
    OpenAI = None

st.set_page_config(page_title="Contestador Inteligente de Glosas", page_icon="📄", layout="wide")

APP_TITLE = "Contestador Inteligente de Glosas"
NORM_PATH = Path(__file__).parent / "normativa" / "Resolucion_2284_2023.pdf"

st.title("📄 " + APP_TITLE)
st.caption("Versión web — Resolución 2284 de 2023 + Manual Único + plantilla institucional")

with st.sidebar:
    st.header("Configuración")
    api_key = st.text_input(
        "Clave de IA (opcional en pruebas)",
        type="password",
        value=st.secrets.get("OPENAI_API_KEY", os.getenv("OPENAI_API_KEY", "")),
        help="Para producción, guárdala como secret y no en el código."
    )
    model = st.selectbox("Modelo", ["gpt-4.1-mini", "gpt-4.1"], index=0)
    st.divider()
    st.markdown("**Regla institucional**")
    st.write("La IA argumenta, pero no inventa hechos clínicos, diagnósticos, resultados, fechas ni soportes.")

def pdf_text(file_bytes):
    reader = PdfReader(io.BytesIO(file_bytes))
    return "\n".join((p.extract_text() or "") for p in reader.pages)

@st.cache_data
def load_normativa():
    if not NORM_PATH.exists():
        return ""
    return pdf_text(NORM_PATH.read_bytes())

def extract_basic(text):
    def first(pattern):
        m = re.search(pattern, text, re.I | re.M)
        return m.group(1).strip() if m else ""
    return {
        "factura": first(r"(?:factura|n[úu]mero de factura)\s*[:#]?\s*([A-Z0-9-]+)"),
        "siniestro": first(r"siniestro\s*[:#]?\s*([A-Z0-9-]+)"),
        "radicado": first(r"radicado\s*[:#]?\s*([A-Z0-9-]+)"),
    }

def extract_candidates(text):
    lines = [x.strip() for x in text.splitlines() if x.strip()]
    out, seen = [], set()
    for i, line in enumerate(lines):
        for m in re.finditer(r"\b(\d{5})\b", line):
            code = m.group(1)
            ctx = " ".join(lines[max(0, i-2):min(len(lines), i+5)])
            key = (code, ctx)
            if key not in seen:
                seen.add(key)
                out.append({"procedimiento": code, "contexto": ctx})
    return out[:100]

def norm_snippets(norm_text, query, max_chars=7000):
    terms = [t.lower() for t in re.findall(r"[a-záéíóúñ0-9]{4,}", query) if t.lower() not in {"para","como","esta","este","glosa"}]
    lines = norm_text.splitlines()
    hits = []
    for i, line in enumerate(lines):
        low = line.lower()
        if sum(1 for t in terms if t in low) >= 1:
            hits.extend(lines[max(0,i-2):min(len(lines),i+4)])
            if len("\n".join(hits)) >= max_chars:
                break
    return "\n".join(dict.fromkeys(hits))[:max_chars]

def ai_argument(glosa_text, norm_text, meta, code, context):
    if not api_key or OpenAI is None:
        return (
            "Borrador pendiente de IA: controvertir la glosa con fundamento en la "
            "causal identificada y en la normativa aplicable, sin incorporar hechos "
            "clínicos no contenidos en los soportes disponibles."
        )
    client = OpenAI(api_key=api_key)
    evidence = norm_snippets(norm_text, context + " " + code)
    prompt = f"""
Eres auditor de cuentas médicas de una IPS en Colombia. Debes contestar una glosa.
OBJETIVO: defender la factura y argumentar la reconsideración de la glosa.

REGLA ABSOLUTA: NO inventes hechos clínicos, diagnósticos, resultados, fechas,
procedimientos realizados, órdenes médicas ni soportes. Si un hecho no aparece
en la glosa o en la información suministrada, no lo afirmes como hecho.
SÍ puedes formular argumentos jurídicos, técnicos y administrativos basados en
la normativa suministrada y en la naturaleza de la objeción.

Normativa base:
{evidence}

Datos:
{json.dumps(meta, ensure_ascii=False)}

Código/procedimiento detectado: {code}
Contexto de la glosa:
{context}

Texto completo de la glosa (para contexto):
{glosa_text[:18000]}

Devuelve SOLO JSON con:
codigo_respuesta_sugerido, valor_aceptado_sugerido, valor_objetado_sugerido,
justificacion, soporte_a_revisar.
La justificación debe ser un texto formal, firme y técnico para contestar la glosa.
No digas que "no hay evidencia" como motivo para no contestar; contesta con el
mejor argumento normativo disponible y, si corresponde, deja el soporte como
pendiente de revisión.
"""
    resp = client.chat.completions.create(
        model=model,
        messages=[{"role":"user","content":prompt}],
        temperature=0.1,
        response_format={"type":"json_object"}
    )
    return json.loads(resp.choices[0].message.content)

def make_pdf(meta, rows):
    buf = io.BytesIO()
    doc = SimpleDocTemplate(
        buf, pagesize=letter,
        rightMargin=1.5*cm, leftMargin=1.5*cm,
        topMargin=1.3*cm, bottomMargin=1.3*cm
    )
    styles = getSampleStyleSheet()
    styles.add(ParagraphStyle(name="Small", parent=styles["BodyText"], fontSize=8.5, leading=10))
    styles.add(ParagraphStyle(name="Body", parent=styles["BodyText"], fontSize=9.5, leading=12))
    styles.add(ParagraphStyle(name="Title2", parent=styles["Heading1"], alignment=TA_CENTER, fontSize=12))
    story = []

    header = Table([["RESPUESTA A GLOSAS", f"FECHA\n{date.today().strftime('%d/%m/%Y')}"]], colWidths=[12*cm, 5*cm])
    header.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.7,colors.black),
        ("ALIGN",(0,0),(0,0),"CENTER"),
        ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
        ("FONTNAME",(0,0),(0,0),"Helvetica-Bold"),
        ("FONTNAME",(1,0),(1,0),"Helvetica"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("TOPPADDING",(0,0),(-1,-1),8),("BOTTOMPADDING",(0,0),(-1,-1),8)
    ]))
    story += [header, Spacer(1, 10)]
    info = Table([
        ["N. DE FACTURA", meta.get("factura","")],
        ["VALOR GLOSADO", meta.get("valor_glosado","")]
    ], colWidths=[5*cm, 12*cm])
    info.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.7,colors.black),
        ("FONTNAME",(0,0),(0,-1),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),9),
        ("TOPPADDING",(0,0),(-1,-1),7),("BOTTOMPADDING",(0,0),(-1,-1),7)
    ]))
    story += [info, Spacer(1, 12)]
    story += [Paragraph("Señores", styles["Body"]), Paragraph(meta.get("aseguradora",""), styles["Body"]),
              Spacer(1,8), Paragraph("<b>Asunto:</b> Respuesta a glosas totales / parciales", styles["Body"]),
              Spacer(1,8), Paragraph("Respetados señores:", styles["Body"]),
              Spacer(1,6),
              Paragraph("En atención a la objeción formulada por su entidad, nos permitimos dar respuesta a la glosa correspondiente a la factura que se relaciona a continuación, anexando la documentación soporte a que haya lugar, con el fin de que se proceda con la revisión y el reconocimiento del pago respectivo.", styles["Body"]),
              Spacer(1,12), Paragraph("<b>Detalle de la respuesta</b>", styles["Body"]), Spacer(1,6)]
    data = [["COD. RESP","VALOR ACEPTADO","VALOR OBJETADO","JUSTIFICACIÓN"]]
    for r in rows:
        data.append([
            Paragraph(str(r.get("codigo_respuesta","")), styles["Small"]),
            Paragraph(str(r.get("valor_aceptado","")), styles["Small"]),
            Paragraph(str(r.get("valor_objetado","")), styles["Small"]),
            Paragraph(str(r.get("justificacion","")), styles["Small"])
        ])
    tbl = Table(data, colWidths=[2.2*cm,3*cm,3*cm,8.8*cm], repeatRows=1)
    tbl.setStyle(TableStyle([
        ("GRID",(0,0),(-1,-1),0.5,colors.black),
        ("BACKGROUND",(0,0),(-1,0),colors.lightgrey),
        ("FONTNAME",(0,0),(-1,0),"Helvetica-Bold"),
        ("FONTSIZE",(0,0),(-1,-1),7.5),
        ("VALIGN",(0,0),(-1,-1),"TOP"),
        ("TOPPADDING",(0,0),(-1,-1),5),("BOTTOMPADDING",(0,0),(-1,-1),5)
    ]))
    story += [tbl, Spacer(1,12),
              Paragraph("Agradecemos la atención prestada a la presente y quedamos atentos a cualquier información adicional que se requiera para el trámite correspondiente.", styles["Body"]),
              Spacer(1,14), Paragraph("Atentamente,", styles["Body"]), Spacer(1,28),
              Paragraph("<b>ROSA ROBLES GALLEGO</b>", styles["Body"]),
              Paragraph("Auditora de Cuentas Médicas", styles["Body"])]
    doc.build(story)
    return buf.getvalue()

pdf = st.file_uploader("1. Sube el PDF de la glosa", type=["pdf"])
if pdf:
    glosa_bytes = pdf.getvalue()
    glosa_text = pdf_text(glosa_bytes)
    meta = extract_basic(glosa_text)

    c1,c2,c3 = st.columns(3)
    meta["factura"] = c1.text_input("Factura", meta["factura"])
    meta["siniestro"] = c2.text_input("Siniestro", meta["siniestro"])
    meta["radicado"] = c3.text_input("Radicado", meta["radicado"])

    meta["aseguradora"] = st.text_input("EPS / Aseguradora", meta.get("aseguradora",""))
    meta["valor_glosado"] = st.text_input("Valor glosado", meta.get("valor_glosado",""))

    candidates = extract_candidates(glosa_text)
    st.subheader(f"2. Glosas/procedimientos detectados: {len(candidates)}")

    if "results" not in st.session_state:
        st.session_state.results = []

    if st.button("🚀 Generar contestación", type="primary"):
        norm_text = load_normativa()
        results = []
        for item in candidates:
            r = ai_argument(glosa_text, norm_text, meta, item["procedimiento"], item["contexto"])
            if isinstance(r, str):
                r = {
                    "codigo_respuesta_sugerido": "",
                    "valor_aceptado_sugerido": "",
                    "valor_objetado_sugerido": "",
                    "justificacion": r,
                    "soporte_a_revisar": "Revisión institucional"
                }
            r["procedimiento"] = item["procedimiento"]
            results.append(r)
        st.session_state.results = results

    if st.session_state.results:
        st.subheader("3. Revisar contestaciones")
        edited = []
        for i, r in enumerate(st.session_state.results):
            with st.expander(f"Glosa {i+1} — procedimiento {r.get('procedimiento','')}", expanded=True):
                r["codigo_respuesta_sugerido"] = st.text_input("Código de respuesta", r.get("codigo_respuesta_sugerido",""), key=f"cr{i}")
                r["valor_aceptado_sugerido"] = st.text_input("Valor aceptado", r.get("valor_aceptado_sugerido",""), key=f"va{i}")
                r["valor_objetado_sugerido"] = st.text_input("Valor objetado", r.get("valor_objetado_sugerido",""), key=f"vo{i}")
                r["justificacion"] = st.text_area("Justificación", r.get("justificacion",""), height=180, key=f"ju{i}")
                r["soporte_a_revisar"] = st.text_input("Soporte a revisar posteriormente", r.get("soporte_a_revisar",""), key=f"sp{i}")
                edited.append(r)
        st.session_state.results = edited

        if st.button("📄 Generar PDF institucional"):
            out = make_pdf(meta, edited)
            st.download_button("Descargar respuesta PDF", out,
                               file_name=f"Respuesta_Glosas_{meta.get('factura','sin_factura')}.pdf",
                               mime="application/pdf")

        st.download_button("Descargar datos JSON",
                           json.dumps({"meta":meta,"glosas":edited}, ensure_ascii=False, indent=2),
                           file_name="respuesta_glosas.json",
                           mime="application/json")
else:
    st.markdown("""
    ### Cómo funciona
    1. Carga el PDF recibido.
    2. La aplicación identifica los registros.
    3. El motor consulta la normativa incorporada.
    4. Genera una contestación argumentada.
    5. Tú revisas y editas.
    6. Descargas el PDF.
    """)
