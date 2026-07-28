# Carnets de Sanidad — Generador automatizado

Aplicacion web (Streamlit) para automatizar la generacion de carnets de
sanidad: carga una lista de personas desde Excel, valida los datos, coloca
la informacion en una plantilla Word y genera documentos Word (editables) y
PDF (listos para imprimir), con varios carnets distribuidos por hoja.

> **Version:** MVP / Fase 1. Ver `AGENTS.md` para el detalle de que queda
> fuera de alcance por ahora (Codigo QR, campos Resultado/Observaciones,
> integraciones externas, despliegue en produccion, etc.)

## Requisitos del sistema

- Python 3.10 o superior.
- **LibreOffice** instalado en el sistema (se usa para convertir el Word
  final a PDF manteniendo exactamente el mismo layout). No se requiere ni se
  asume que Microsoft Word este instalado.

  - Ubuntu/Debian: `sudo apt-get install libreoffice`
    (o, para una instalacion minima: `sudo apt-get install libreoffice-writer`)
  - macOS: `brew install --cask libreoffice`
  - Windows: descargar el instalador desde https://www.libreoffice.org/download/

  Verifica la instalacion con:
  ```bash
  soffice --version
  ```

## Instalacion

```bash
git clone <este-repositorio>
cd carnets-sanidad
python3 -m venv .venv
source .venv/bin/activate   # En Windows: .venv\Scripts\activate
pip install -r requirements.txt
```

## Ejecucion

```bash
streamlit run app.py
```

Esto abre la aplicacion en el navegador (por defecto en
`http://localhost:8501`).

## Uso paso a paso

1. **Inicio** — resumen general y ultimas generaciones.
2. **Nueva generacion**:
   - Descarga el **Excel estandar** (boton en la parte superior). Trae los
     encabezados exactos requeridos y una hoja de instrucciones.
   - Completa el Excel con tus datos (no cambies los encabezados).
   - Elige una **plantilla Word** de carnet: puedes usar la plantilla por
     defecto incluida, o subir la tuya propia (ver "Crear tu propia
     plantilla" mas abajo).
   - Si tu plantilla usa el campo `{{ foto }}`, sube las fotografias
     (individualmente o en un `.zip`). El nombre del archivo debe coincidir
     exactamente con la columna "Fotografia" del Excel.
   - Sube tu Excel completado y presiona "Analizar y validar registros".
3. **Validacion** — revisa los registros correctos y los observados (con el
   motivo exacto de cada error). Puedes descargar un Excel con el detalle de
   los observados.
4. **Seleccion** — elige que personas procesar (busqueda por DNI/nombre,
   filtros por empresa/puesto, seleccionar todos/quitar todos).
5. **Vista previa** — revisa como quedara cada hoja impresa antes de
   generar los archivos finales.
6. **Generacion** — genera el Word editable y el PDF listo para imprimir.
   Puedes descargarlos por separado o juntos en un `.zip`. La generacion
   queda registrada en el Historial.
7. **Historial** — consulta generaciones anteriores y vuelve a descargar
   (reimprimir) sus archivos.
8. **Configuracion** — ajusta cuantos carnets van por hoja, tamano de
   papel, margenes, separacion, calibracion de impresion, formato de
   numeracion automatica y periodo de vigencia.

## Estructura del Excel

La aplicacion espera exactamente estas columnas (ver la hoja "Personas" del
Excel estandar descargable):

| Columna | Obligatoria | Notas |
|---|---|---|
| DNI | Si | Formato configurable (por defecto 8 digitos) |
| Nombres | Si | |
| Apellido paterno | Si | |
| Apellido materno | Si | |
| Fecha de nacimiento | Si | Formato sugerido `AAAA-MM-DD` |
| Numero de carnet | No* | *Obligatoria salvo que actives autogeneracion en Configuracion |
| Fecha de emision | Si | |
| Fecha de vencimiento | No* | *Obligatoria salvo que actives autocalculo en Configuracion |
| Empresa | No | |
| Puesto | No | |
| Fotografia | No** | **Obligatoria solo si la plantilla Word usa `{{ foto }}` |

No debe haber columnas de "Resultado", "Observaciones" ni "Codigo QR": no
forman parte del alcance de esta version.

## Crear tu propia plantilla Word

La plantilla es un `.docx` normal, del tamano de un carnet (por ejemplo,
tarjeta 85.6 x 54 mm), con estos campos escritos literalmente donde quieras
que aparezca cada dato:

```
{{ numero_carnet }}
{{ nombres_completos }}
{{ dni }}
{{ fecha_nacimiento }}
{{ fecha_emision }}
{{ fecha_vencimiento }}
{{ empresa }}
{{ puesto }}
{{ foto }}
```

Recomendaciones:
- Usa una tabla sin bordes para posicionar los elementos (no uses espacios,
  tabulaciones ni saltos de linea manuales para "acomodar" el texto).
- Escribe cada campo como texto normal (sin dividirlo en varias fuentes o
  formatos dentro del mismo campo), para que la aplicacion pueda ubicarlo y
  reemplazarlo correctamente.
- Ajusta el tamano de carnet configurado en **Configuracion → Distribucion e
  impresion** (Ancho/Alto de carnet) para que coincida con el tamano real de
  tu plantilla.

## Impresion

El PDF generado es el documento recomendado para imprimir, porque se genera
convirtiendo el mismo `.docx` final: las posiciones en el PDF coinciden
exactamente con las del Word editable. Antes de imprimir:

- Verifica en Configuracion que el tamano de papel y la orientacion
  coincidan con tu impresora.
- Imprime siempre al **100% de escala** (sin "ajustar a pagina"), para que
  las medidas en milimetros configuradas se respeten.
- Usa el desplazamiento horizontal/vertical de calibracion si notas que la
  impresion queda corrida respecto del papel/troquel que usas.

## Pruebas

```bash
pytest tests/ -v
```

Cubren: Excel correcto, columnas faltantes, DNI vacio/duplicado/previamente
procesado, fechas invalidas, vencimiento anterior a emision, numero de
carnet repetido, fotografia faltante, nombres muy largos, generacion con una
persona, con varias personas, con la ultima hoja incompleta, generacion de
Word, generacion de PDF, y el historial local.

## Datos y privacidad

- No se incluyen DNIs ni datos reales en el repositorio (ver `.gitignore`).
- Los archivos generados y el historial (`data/`, `generated/`) tampoco se
  versionan: cada instalacion mantiene los suyos localmente.
- Los archivos subidos (Excel, plantillas, fotos, ZIP) se validan por
  extension y tamano antes de procesarse.

## Limitaciones conocidas de esta version (Fase 1)

- El reemplazo de campos en plantillas Word subidas por el usuario funciona
  de forma confiable cuando cada campo `{{ campo }}` esta escrito como texto
  simple y continuo. Si Word fragmenta el texto en varias "runs" (por
  ejemplo, por autocorreccion), el campo podria no reemplazarse; vuelve a
  escribirlo de corrido si eso ocurre.
- El tamano del carnet dentro de la plantilla debe coincidir razonablemente
  con el "Ancho/Alto de carnet" configurado, para que la hoja se vea
  ordenada.
- No hay autenticacion de usuarios en esta version (ver `AGENTS.md`).
