# Contestador Inteligente de Glosas — Gemini (versión estable)

## Archivos del proyecto
- `app.py`
- `requirements.txt`
- `packages.txt`
- `PLANTILLA OFICIAL GLOSAS.docx`
- `normativa/Resolucion_2284_2023.pdf`

## Configuración en Streamlit Cloud
1. Sube/reemplaza los archivos del repositorio manteniendo `app.py` como archivo principal.
2. En **Manage app > Settings > Secrets** puedes agregar:
   `GEMINI_API_KEY = "TU_CLAVE"`
3. También puedes pegar temporalmente la clave en **Configuración > Clave Gemini** dentro de la aplicación.
4. Reinicia la aplicación después de actualizar los archivos.

## Corrección principal
Esta versión elimina `PdfMerger`, que no está disponible con `pypdf 5.x`, y usa `PdfWriter` para consolidar el PDF. Esto evita el `ImportError` que estaba deteniendo la aplicación al iniciar.

## Dependencias
`requirements.txt` fija `pypdf` en la rama 5.x y no requiere `reportlab`.

## Flujo
1. Cargar PDF de glosa.
2. Cargar soportes opcionales (PDF, DOCX, PNG/JPG).
3. Analizar con Gemini.
4. Revisar y corregir datos.
5. Generar respuesta usando la plantilla institucional.
6. Convertir la plantilla a PDF y anexar soportes.
7. Descargar PDF final.
