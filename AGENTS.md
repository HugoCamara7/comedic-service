# AGENTS.md — Guia para agentes/desarrolladores que continuen este proyecto

## Que es este proyecto

Aplicacion Streamlit para automatizar la generacion de carnets de sanidad:
carga de personas desde Excel, validacion, seleccion, vista previa y
generacion masiva de carnets en Word (editable) y PDF (listo para
imprimir), con historial local en SQLite.

Este es el **MVP (Fase 1)**. Ver "Pendientes / Fase 2" al final.

## Como ejecutar

```bash
pip install -r requirements.txt
# LibreOffice debe estar instalado en el sistema (ver README.md)
streamlit run app.py
```

## Como correr las pruebas

```bash
pytest tests/ -v
```

Todas las pruebas deben pasar antes de considerar terminado un cambio.
Ademas de pytest, para cualquier cambio que afecte la generacion de
documentos, **revisa visualmente** el resultado:

```bash
python3 - <<'EOF'
from pathlib import Path
from modules import template_manager, word_generator, pdf_generator
from modules.config import AppConfig

tpl = template_manager.ensure_default_template()
people = [{"dni": "12345678", "nombre_completo": "Prueba Visual", "numero_carnet": "C-1",
           "fecha_nacimiento": "01/01/1990", "fecha_emision": "01/01/2026",
           "fecha_vencimiento": "01/01/2027", "empresa": "X", "puesto": "Y", "fotografia": ""}]
cfg = AppConfig()
out = Path("/tmp/preview.docx")
word_generator.generate_word_document(people, tpl, None, False, cfg.print, out)
pdf = pdf_generator.convert_docx_to_pdf(out, Path("/tmp"))
print(pdf)
EOF
pdftoppm -jpeg -r 120 /tmp/preview.pdf /tmp/preview
# luego revisa /tmp/preview-1.jpg
```

## Decisiones de arquitectura importantes (no revertir sin razon)

1. **PDF = conversion de LibreOffice del mismo .docx final**, no ReportLab.
   Esto garantiza que el PDF impreso coincide exactamente con el Word
   editable, porque provienen del mismo archivo fuente. Ver
   `modules/pdf_generator.py` para el razonamiento completo.

2. **Cada carnet se renderiza individualmente con docxtpl** (un
   `DocxTemplate` por persona, contexto propio) y **luego se compone en la
   hoja maestra copiando XML con python-docx** (`modules/word_generator.py`).
   Se descarto usar `docxtpl.Subdoc` con datos distintos por celda porque
   docxtpl renderiza todo `document.xml` en una sola pasada de Jinja2 con un
   contexto compartido: subdocumentos con placeholders identicos (`{{ dni }}`)
   en distintas celdas terminarian todos mostrando el mismo valor. Renderizar
   cada carnet por separado y copiar el resultado ya resuelto evita ese
   problema de scope.

3. **La numeracion de carnet y la fecha de vencimiento nunca se inventan.**
   Si el Excel no trae esos valores y la autogeneracion/autocalculo estan
   desactivados en Configuracion, el registro queda "observado" con un
   motivo claro. No agregar logica que asuma un formato de numeracion o un
   periodo de vigencia por defecto sin que el usuario lo haya configurado
   explicitamente.

4. **DNI, Nombres, Apellidos, Fecha de nacimiento y Fecha de emision son
   siempre obligatorios.** Empresa y Puesto son siempre opcionales. No
   cambiar esto sin que el usuario lo pida explicitamente.

5. **No agregar Codigo QR, campo Resultado, campo Observaciones, Supabase,
   Power Automate, Google Sheets, GitHub Actions ni despliegue en
   produccion** hasta que el usuario lo solicite (fuera de alcance de la
   Fase 1 por decision explicita).

## Estructura del codigo

```
app.py                      # Pantalla "Inicio" (punto de entrada de Streamlit)
pages/
  1_Nueva_generacion.py     # Descarga Excel estandar, sube Excel/plantilla/fotos, valida
  2_Validacion.py           # Muestra correctos/observados, descarga reporte de errores
  3_Seleccion.py            # Selecciona que personas procesar (busqueda/filtros)
  4_Vista_previa.py         # Renderiza y muestra cada hoja antes de generar
  5_Generacion.py           # Genera Word + PDF finales, guarda en historial
  6_Historial.py            # Lista generaciones pasadas, permite reimprimir
  7_Configuracion.py        # Edita impresion, numeracion, vigencia, validacion
modules/
  excel_reader.py           # Plantilla Excel estandar + carga/normalizacion
  config.py                 # AppConfig persistida en data/config.json
  validations.py            # Reglas de validacion por registro
  assignment.py             # Resuelve numero_carnet autogenerado
  template_manager.py       # Plantilla Word por defecto + gestion de plantillas
  print_layout.py           # Construye la hoja maestra (grilla) segun Configuracion
  word_generator.py         # Renderiza carnets individuales y compone las hojas
  pdf_generator.py          # Conversion a PDF via LibreOffice
  history.py                # SQLite: generaciones y personas procesadas
  session.py                # Carpeta de trabajo por sesion de Streamlit (fotos, salidas)
tests/                      # pytest — ejecutar antes de dar por terminado un cambio
templates/                  # Plantillas Word (la por defecto se genera automaticamente)
data/                       # config.json + history.db (no se versiona, ver .gitignore)
generated/                  # Archivos generados (sessions/, history/) (no se versiona)
```

## Convenciones

- Todo el codigo de negocio (validaciones, generacion, historial) vive en
  `modules/`, no en `pages/*.py`. Las paginas solo orquestan UI + llamadas a
  `modules/`.
- Los textos de la interfaz y los mensajes de error estan en espanol,
  orientados a que el usuario final entienda **como corregir** el problema,
  no solo que hubo un error.
- No usar datos personales reales en pruebas ni ejemplos: usar
  DNIs/nombres ficticios (ver `tests/` para el patron).
- Los archivos subidos por el usuario (Excel, plantilla, fotos, ZIP) siempre
  se validan (extension, tamano) antes de procesarse. Ver
  `modules/excel_reader.py`, `modules/template_manager.py` y
  `modules/session.py` para los limites actuales.

## Pendientes / Fase 2 (explicitamente fuera de alcance de esta version)

- Codigo QR en el carnet.
- Campos "Resultado" y "Observaciones" del examen medico.
- Integracion con Supabase (base de datos remota / autenticacion).
- Integracion con Power Automate / Google Sheets.
- CI con GitHub Actions.
- Despliegue en produccion (Docker, servidor, dominio).
- Autenticacion de usuarios (la arquitectura actual separa sesiones por
  `session_id` de Streamlit, pero no hay login).
- Mejoras posibles no solicitadas todavia: soporte para mas de una foto por
  plantilla, edicion visual (drag & drop) de la posicion de los campos en la
  plantilla, exportacion del historial completo a Excel.
