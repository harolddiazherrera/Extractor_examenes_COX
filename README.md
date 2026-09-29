# Extractor de exámenes ocupacionales

Herramienta para extraer a Excel los datos principales (nombre, RUT, cargo,
código de verificación del QR, examen, fechas y resultado) de informes de
exámenes ocupacionales en PDF. Reconoce los formatos de **CMT** (Centro
Médico del Trabajador) y **WORKMED**; cualquier otro queda marcado como
*FORMATO NO RECONOCIDO*.

Todo el procesamiento es local: no usa internet para leer los PDF ni
inteligencia artificial. Los archivos nunca se envían a un servidor.

## Contenido del repositorio

| Archivo | Uso |
|---|---|
| `index.html` | Versión web. Se abre en el navegador (GitHub Pages, carpeta compartida o doble clic). Entrega el Excel de datos y el informe de incidencias. |
| `extractor_examenes_gui.py` | Versión de escritorio con ventana (Tkinter). Permite elegir carpeta de origen, subcarpetas y carpeta de destino. |
| `extraer_examenes.py` | Versión de línea de comandos: `python extraer_examenes.py <carpeta> [salida.xlsx]`. |
| `requirements.txt` | Librerías Python necesarias para las versiones de escritorio. |

## Versión web

La pantalla se divide en tres columnas: a la izquierda la lista de funciones, al centro la
configuración de la función elegida y sus resultados, y a la derecha el progreso (porcentaje,
archivos procesados, tiempo transcurrido y restante) con los botones **Pausar**, **Detener**
y **Reiniciar**, más el registro de eventos con hora. Los controles de ejecución funcionan
igual en todas las funciones; al detener se conserva lo procesado hasta ese momento y el
Excel se marca como corrida parcial.

1. Abre `index.html` (o la URL publicada).
2. Elige la función **Extraer exámenes a Excel** en el panel izquierdo.
3. Arrastra los PDF o una carpeta completa. La casilla *Incluir subcarpetas* controla si se procesan las carpetas internas.
4. Presiona **Procesar**.
5. Descarga:
   - **Excel de datos**: una fila por examen.
   - **Informe de incidencias**: PDF no reconocidos o con campos faltantes, junto con el texto leído de cada página. Este archivo es el insumo para agregar formatos nuevos al programa.

La página carga dos librerías públicas desde CDN (`pdf.js` y `SheetJS`), por lo que requiere conexión a internet solo para abrirse; los PDF no salen del computador.

## Versión de escritorio

```
pip install -r requirements.txt
python extractor_examenes_gui.py
```

## Cómo agregar una función nueva (versión web)

Las funciones se definen en el arreglo `FUNCIONES` de `index.html`. Cada una es un objeto
con `id`, `nombre`, `descripcion`, `montar(cont)` (dibuja su configuración en el panel
central), `reiniciar()` y su lógica de ejecución, que usa los objetos comunes `ejecucion`
(pausar/detener), `progreso` y `evento` para integrarse con la interfaz. Al agregarla al
arreglo aparece automáticamente en el panel izquierdo.

## Cómo agregar un formato nuevo

1. Procesa los PDF con la versión web y descarga el informe de incidencias.
2. La hoja *Incidencias* contiene el texto de cada página no reconocida. Ese texto contiene datos personales: trátalo con la misma reserva que el PDF original.
3. Con ese texto se ajustan las reglas de extracción (funciones `parsearCMT` / `parsearWORKMED` en `index.html`, o `parsear_cmt` / `parsear_workmed` en los `.py`) y se agrega un parser nuevo.
4. Actualiza la constante `VERSION` al pie de `index.html` y en la hoja *Resumen* quedará registrado con qué versión se generó cada Excel.

## Columnas del Excel de datos

Archivo PDF · Centro médico · Nombre completo · RUT · Cargo · Código de
verificación (QR) · N° de reserva · Examen · Fecha de evaluación ·
Vigencia hasta · Resultado · Sucursal / lugar · Estado extracción

## Historial de versiones

- **1.0 (2026-09-29)**: formatos CMT y WORKMED; versión web, escritorio y línea de comandos.
- **1.1 (2026-09-29)**: versión web con panel de funciones a la izquierda, configuración y resultados al centro, progreso con tiempo estimado y botones Pausar / Detener / Reiniciar, y registro de eventos con hora. Reglas de extracción sin cambios.
