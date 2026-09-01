# Contestador Inteligente de Glosas — Gemini

Versión preparada para trabajar con la plantilla institucional `PLANTILLA OFICIAL GLOSAS.docx`.

## Importante sobre la plantilla
La plantilla contiene **datos de ejemplo** (por ejemplo, entidad, factura, valores y código). Esos datos sirven únicamente como modelo visual y **no se usan como datos reales**. Antes de generar el PDF, la aplicación reemplaza los campos variables con la información revisada por el usuario y bloquea la generación si faltan campos obligatorios o si detecta que un dato de ejemplo quedó en el documento.

## Configuración
En Streamlit Secrets:

`GEMINI_API_KEY = "TU_CLAVE"`

También existe una casilla temporal en la barra lateral. No subir la clave a GitHub.

## Flujo
1. Cargar glosa PDF.
2. Cargar múltiples soportes opcionales.
3. Gemini analiza la glosa y la información disponible.
4. Se muestra una sola argumentación general para la factura.
5. El usuario revisa y modifica todos los campos.
6. RE9602 queda por defecto para revisión; el código seleccionado no imprime su significado en el PDF.
7. Aceptado = 0 por defecto.
8. Objetado = valor glosado por defecto.
9. Se genera el documento usando la plantilla institucional suministrada.
10. Los soportes PDF/DOCX/PNG/JPG se convierten y se anexan al PDF final en un solo archivo.

## Normativa
Colocar la Resolución 2284 de 2023 y otros documentos normativos dentro de `normativa/`. La carpeta se conserva separada de los archivos de la aplicación.
