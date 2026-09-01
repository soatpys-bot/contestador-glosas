# Contestador Inteligente de Glosas

Paquete limpio para reemplazar la versión anterior.

## Archivos que deben quedar en el repositorio
- `app.py`
- `requirements.txt`
- `packages.txt`
- `PLANTILLA OFICIAL GLOSAS.docx`
- `normativa/Resolucion_2284_2023.pdf`

## Configuración en Streamlit
1. Conecta este repositorio y la rama `main`.
2. Archivo principal: `app.py`.
3. En Secrets agrega `OPENAI_API_KEY = "tu_clave"`.
4. Despliega/reinicia la aplicación.

## Comportamiento acordado
- Mantiene la plantilla institucional cargada, sin rediseñarla.
- Fecha automática del día de generación.
- Entidad dinámica; intenta leer texto, membrete y logo. Si no hay certeza, permite escribirla.
- Factura y valor glosado dinámicos.
- No detecta ni inventa códigos de glosa.
- Código de respuesta `RE9602` por defecto, modificable.
- Muestra el contexto del código en pantalla, pero no lo imprime en el PDF.
- Valor aceptado `0` por defecto.
- Valor objetado igual al total de la objeción por defecto, modificable.
- Una sola argumentación general para todas las glosas de la factura.
- Permite revisar y modificar todo antes de generar el PDF.
- Permite múltiples soportes y los incorpora al PDF final.
- Puede generar respuesta aun sin soportes adicionales.
