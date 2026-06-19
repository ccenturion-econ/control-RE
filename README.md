# control-RE

Aplicacion de escritorio para calcular horas extra a partir del PDF del reloj marcador.

## Descarga para usuarios

Entrar a la seccion Releases del repositorio y descargar la version correspondiente:

- Mac Apple Silicon: `GeneradorHorasExtra-mac-arm64.zip`
- Windows: `GeneradorHorasExtra-windows.zip`

## Uso

1. Abrir la aplicacion.
2. Seleccionar el PDF de asistencia.
3. Elegir donde guardar el Excel.
4. Revisar o ajustar los parametros visibles.
   En `Fechas de lluvia`, se pueden ingresar varias fechas separadas por comas.
5. Presionar `Generar Excel`.

El Excel abre en `Vista principal`, con las columnas de uso habitual y el resumen semanal en la parte superior. La hoja `Detalle` conserva todos los cálculos, controles y columnas para revisión.

## Parametros principales

- Hora oficial de entrada: hora desde la cual empieza a contar la jornada. Los minutos marcados antes de esta hora no suman para horas extra.
- Llegada tarde desde: primer minuto que cuenta como llegada tarde.
- Salida anticipada antes de: hora antes de la cual se registra uso de cupo por salida anticipada.
- Entrada planificada: hora de entrada que se usa para proyectar dias futuros en meses parciales.
- Fechas de lluvia: acepta `YYYY-MM-DD` y `DD/MM/YYYY`, separadas por comas. En esas fechas se agregan 30 minutos a la tolerancia normal de llegada.

Los siguientes valores se muestran en el Excel generado; salvo el tope mensual, son fijos en la aplicacion:

- Horas regulares: 8.
- Tope diario: 3 horas.
- Tope semanal: 8 horas.
- Tope mensual: editable por el usuario, con 25 horas como valor predeterminado.
- Usos de cupo permitidos: 3.

## Reglas incluidas

- Entrada/salida se muestran con segundos.
- Los calculos se hacen truncando cada marca al minuto.
- Los minutos previos a la hora oficial de entrada no cuentan para horas extra.
- La salida anticipada cuenta como uso de cupo informativo.
- Desde la cuarta llegada tarde, ese dia no suma horas extra.
- En un dia de lluvia, la llegada tarde comienza 30 minutos despues del umbral normal. Por ejemplo, si el umbral es 08:16, comienza a las 08:46.
- La lluvia no modifica la salida anticipada ni impide remunerar horas extra cuando la llegada esta dentro de la tolerancia extendida.

## Linea de comandos

```bash
python src/generate_billable_hours_from_pdf.py asistencia.pdf \
  --rain-dates "2026-05-04,12/05/2026"
```
