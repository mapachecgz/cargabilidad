# Ejemplos de configuracion de recursos

Estos archivos usan una persona ficticia y son listas JSON completas con la misma estructura que `config/recursos.json`. El generador no los carga automaticamente: siempre lee `recursos.json` del directorio de configuracion seleccionado.

| Archivo | Ejemplo |
|---|---|
| `recursos_vacaciones.json` | Un intervalo de vacaciones. |
| `recursos_incapacidades.json` | Un intervalo de incapacidad. |
| `recursos_ausencias.json` | Dos ausencias generales con motivo: permiso y licencia. |
| `recursos_varios_periodos.json` | Vacaciones e incapacidades como listas, combinadas con una ausencia general. |

## Aplicar a una persona real

1. Abre `config/recursos.json` y encuentra a la persona por documento o nombre.
2. Copia solo el campo que necesitas: `vacaciones`, `incapacidades` o `ausencias`.
3. Ajusta las fechas y el motivo. Conserva su documento, nombre, correo, equipo y demas datos reales.
4. Mantiene `activo: true` si la persona sigue vinculada pero temporalmente ausente. `activo: false` significa inactividad y deja capacidad cero en todos los periodos calculados.
5. Ejecuta `ejecutar.bat` para regenerar los reportes.

Las fechas son inclusivas y usan YYYY-MM-DD. Para un dia usa la misma fecha en inicio y fin. Una ausencia se configura por dias completos; no hay campo de horas parciales. Ausencias generales requieren inicio y fin; el motivo identifica la causa. Vacaciones/incapacidades sin periodo se expresan como `{"inicio": null, "fin": null}` o `[]`. Para varios periodos usa listas. No declares solo una fecha ni un fin anterior al inicio.

Los descuentos solo cuentan dias de lunes a viernes dentro de cada sprint, excluyen festivos y no duplican dias coincidentes. Los casos abiertos generan alertas cuando la ausencia incluye la fecha de corte o coincide con su sprint, incluso sin horas estimadas. Closed/Removed/Resolved no son abiertos con la configuracion actual; los estados finales se leen de parametros.json.

Estos ejemplos no sustituyen los otros cuatro JSON requeridos. Para una configuracion de prueba separada, prepara una carpeta con recursos.json, sprints.json, parametros.json, tiempos.json y festivos.json, y usa `--config` o el cuarto argumento del BAT.
