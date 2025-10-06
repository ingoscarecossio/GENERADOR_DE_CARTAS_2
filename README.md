
# Generador de Cartas (v6 SUPREMO)

## Novedades
- **Reglas por YAML**: elige plantilla/tabla por grupo y define exportación a PDF, watermark y nombre de archivo por regla.
- **Placeholders derivados (Jinja2)**: p. ej., `SALUDO: "{{PREFIJO}} {{NOMBRE_DIRECTIVO}}"` sin tocar la plantilla.
- **Pie de página automático**: texto institucional + "Página X de Y" y logo opcional.
- **Firma digital de PDFs** (opcional, con PFX).
- **Validador de calidad**: faltantes por columna, duplicados por ACTOR, rango de fechas por ACTOR.

## Ejecutar
```bash
pip install -r requirements.txt
streamlit run app.py
```
