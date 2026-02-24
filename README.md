# MetraCAD

Utilidad AutoLISP para medir longitudes de entidades en AutoCAD, agrupadas por tipo y capa.

**Autor:** Jean Carlos | **Licencia:** MIT | **Requiere:** AutoCAD 2010+ con Visual LISP (no compatible con LT)

## Instalación

En AutoCAD ejecute `APPLOAD`, busque `metracad.lsp` y cárguelo. Para que se cargue automáticamente en cada sesión, agréguelo en la sección **Contenido de arranque** del mismo diálogo.

## Comandos

| Comando | Qué hace |
| ------- | -------- |
| `METRACAD` | Mide longitudes de forma interactiva |
| `METRACADEXPORT` | Exporta el historial de la sesión a CSV |
| `METRACADCLIP` | Copia el último reporte al portapapeles |

## Uso de METRACAD

Al ejecutarlo, el comando pregunta tres cosas:

1. **Precisión** `[2/3/4]` — cantidad de decimales en el reporte
2. **Unidad** `[m/cm/mm/ft]` — etiqueta que aparece en el resultado (no convierte valores, debe coincidir con las unidades del DWG)
3. **Modo de selección:**

| Modo | Comportamiento |
| --- | --- |
| `Selection` | Usted selecciona los objetos manualmente |
| `Layer` | Selecciona todo lo de la misma capa que el objeto indicado |
| `Color` | Selecciona todo lo del mismo color que el objeto indicado |

Al finalizar muestra un reporte por tipo de entidad y por capa, y ofrece copiarlo al portapapeles.

> Las entidades dentro de bloques no se miden. Use `EXPLODE` primero.

## Entidades soportadas

`LINE` · `LWPOLYLINE` · `POLYLINE` · `ARC` · `CIRCLE` · `SPLINE` · `ELLIPSE`

## Exportación CSV

El archivo generado por `METRACADEXPORT` tiene dos secciones:

- **Resumen por sesión** — una fila por cada vez que ejecutó `METRACAD`, con fecha, archivo, modo, filtro aplicado (ej. `Capa: TUBERIAS`), total de entidades, longitud total y unidad.
- **Detalle por capa** — una fila por cada capa, con su cantidad de entidades y longitud acumulada.

> El historial se borra al cerrar AutoCAD. Exporte antes de cerrar.

## Historial de versiones

### `0.0.1` Versión inicial

### `0.0.2`

- Corrección de variables locales.
- Manejo de errores por entidad.
- Reemplazo de `(exit)` por flag `abort`.
- Timestamp único por sesión.
- Comando `METRACADCLIP`, portapapeles al finalizar.
- CSV con columna Filtro y detalle por capa.
