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
4. Revisar o ajustar los parametros.
5. Presionar `Generar Excel`.

## Parametros principales

- Hora oficial de entrada: hora desde la cual empieza a contar la jornada. Los minutos marcados antes de esta hora no suman para horas extra.
- Llegada tarde desde: primer minuto que cuenta como llegada tarde.
- Salida anticipada antes de: hora antes de la cual se registra uso de cupo por salida anticipada.
- Horas regulares: horas que deben cumplirse antes de empezar a contar horas extra.
- Topes diario, semanal y mensual.

## Reglas incluidas

- Entrada/salida se muestran con segundos.
- Los calculos se hacen truncando cada marca al minuto.
- Los minutos previos a la hora oficial de entrada no cuentan para horas extra.
- La salida anticipada cuenta como uso de cupo informativo.
- Desde la cuarta llegada tarde, ese dia no suma horas extra.
