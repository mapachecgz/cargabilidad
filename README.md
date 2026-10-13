# Cargabilidad del equipo

Proyecto local que transforma un export CSV de Azure DevOps y configuraciones JSON en un Excel y un HTML interactivo. No consulta Azure DevOps ni modifica tareas. Story Points se interpreta como horas estimadas; no mide tiempo ejecutado ni productividad.


## Tabla de contenido

- [Instalacion y ejecucion](#instalacion-y-ejecucion)
- [Archivos del proyecto](#archivos-del-proyecto)
- [CSV de entrada](#csv-de-entrada)
- [Configuraciones JSON: campos](#configuraciones-json-campos)
  - [recursos.json](#recursosjson)
  - [Ejemplos de vacaciones, incapacidades y ausencias](#ejemplos-de-vacaciones-incapacidades-y-ausencias)
  - [sprints.json](#sprintsjson)
  - [parametros.json](#parametrosjson)
  - [tiempos.json](#tiemposjson)
  - [festivos.json](#festivosjson)
  - [tareas_administrativas.json](#tareas_administrativasjson)
- [Formulas e interpretacion](#formulas-e-interpretacion)
- [Alertas de disponibilidad y tareas activas](#alertas-de-disponibilidad-y-tareas-activas)
- [HTML: secciones y campos](#html-secciones-y-campos)
- [Excel: hojas y campos](#excel-hojas-y-campos)
  - [Tablero](#tablero)
  - [Detalle](#detalle)
  - [Alertas](#alertas)
  - [Base talentos](#base-talentos)
  - [Config recursos](#config-recursos)
  - [Config sprints](#config-sprints)
  - [Config festivos](#config-festivos)
  - [Config tiempos](#config-tiempos)
  - [Carga recursos](#carga-recursos)
  - [Historico recursos](#historico-recursos)
  - [Matriz por estado](#matriz-por-estado)
  - [Metricas sprint](#metricas-sprint)
  - [Detalle celulas EVC](#detalle-celulas-evc)
  - [Celulas y EVC](#celulas-y-evc)
  - [Equipos internos](#equipos-internos)
  - [Sprint vigente](#sprint-vigente)
  - [Calidad datos](#calidad-datos)
  - [Guia de lectura](#guia-de-lectura)
  - [Pendientes anteriores](#pendientes-anteriores)
  - [Disponibilidad y tareas](#disponibilidad-y-tareas)
  - [Tareas administrativas](#tareas-administrativas)
  - [Config tareas administrativas](#config-tareas-administrativas)
- [Verificacion y problemas frecuentes](#verificacion-y-problemas-frecuentes)

## Instalacion y ejecucion

Requiere Windows y Python 3 con acceso desde `python` o `py -3`.

```powershell
python -m pip install -r requirements.txt
.\ejecutar.bat
```

Al ejecutar el BAT, ingresa las rutas **sin comillas**. Puedes dejarlas vacias para usar `entrada` y `salida` dentro del proyecto. La entrada es un directorio que debe contener exactamente un CSV, no la ruta de un archivo. La salida es un directorio; se crea si no existe. La fecha de corte usa `YYYY-MM-DD`; vacia usa la fecha actual. Las rutas relativas se interpretan desde la carpeta del proyecto, aunque abras el BAT desde otra ubicacion. Las rutas con espacios se admiten.

Tambien acepta argumentos para ejecutar sin preguntas:

```powershell
.\ejecutar.bat "C:\Datos\Export Azure" "C:\Reportes\Equipo" "2026-10-08"
.\ejecutar.bat "C:\Datos\Export Azure" "C:\Reportes\Equipo" "2026-10-08" "C:\Datos\Configuracion"
```

| Argumento del BAT | Descripcion | Predeterminado |
|---|---|---|
| 1: entrada | Directorio con exactamente un CSV. | `entrada` del proyecto; se pregunta si no se proporciona. |
| 2: salida | Directorio donde guardar Excel y HTML. | `salida` del proyecto; se pregunta si no se proporciona. |
| 3: fecha | Fecha de corte del analisis. | Hoy; se pregunta si no se proporciona. |
| 4: configuracion | Directorio con los cinco JSON obligatorios y tareas_administrativas.json opcional. | `config` del proyecto. |

El BAT valida que exista Python y que este instalado openpyxl. Si falta una dependencia, muestra el comando de instalacion. Devuelve codigo 0 si termina correctamente y un codigo distinto de cero si falla. En ejecucion interactiva deja visible el resultado hasta pulsar una tecla.

Ejecucion directa equivalente:

```powershell
python src/generar_reporte.py --entrada "C:\Datos\Export Azure" --directorio-salida "C:\Reportes\Equipo" --config config --fecha-corte 2026-10-08
```

| Opcion Python | Descripcion |
|---|---|
| `--entrada` / `--base` | Directorio del CSV; por defecto `entrada` del proyecto. |
| `--directorio-salida` | Directorio del Excel y HTML. |
| `--salida` | Alternativa compatible: recibe una ruta de Excel, pero utiliza su directorio; asigna el nombre automaticamente. No combinar con `--directorio-salida`. Por defecto usa `salida`. |
| `--config` | Directorio de configuracion (cinco JSON obligatorios y tareas_administrativas.json opcional); por defecto `config`. |
| `--fecha-corte` | Fecha ISO para sprint vigente, ausencias y pendientes anteriores; por defecto hoy. |

Los dos archivos comparten nombre: `cargabilidad_YYYY-MM-DD_HH-mm-ss.xlsx` y `cargabilidad_YYYY-MM-DD_HH-mm-ss.html`, con fecha y hora de **generacion en Bogota (UTC-5)**. La fecha de corte puede ser diferente y se muestra dentro del reporte. Se conserva un unico HTML: tras guardar el nuevo, se retiran los anteriores `cargabilidad*.html` identificados como reportes de esta herramienta en ese mismo directorio. Se conservan los Excel anteriores y los HTML ajenos a la herramienta. Dos ejecuciones en el mismo segundo usan el mismo nombre.

## Archivos del proyecto

| Ruta | Funcion |
|---|---|
| `ejecutar.bat` | Ejecutor Windows interactivo o por argumentos; permite cambiar rutas y fecha sin editar codigo. |
| `requirements.txt` | Dependencia del proyecto: openpyxl 3.1.5 para crear y leer Excel. |
| `README.md` | Instrucciones, formulas, limites y diccionario de campos. |
| `src/generar_reporte.py` | Lee CSV/JSON, vincula responsables, valida datos, calcula capacidad, genera alertas y coordina la exportacion. |
| `src/administrativos.py` | Valida reglas, normaliza Title, clasifica tareas administrativas y compara Story Points con referencias. |
| `src/metricas.py` | Agrega horas y conteos de recursos identificados; construye hojas de metricas y datos del HTML. |
| `src/vistas.py` | Organiza subtotales y detalles desplegables del Excel, jerarquias organizativas y matriz por estado. |
| `src/dashboard.py` | Inserta datos en la plantilla UTF-8, escapa el JSON, fecha el HTML y conserva una sola version. |
| `src/dashboard.html` | Plantilla HTML, estilos y JavaScript de filtros, tarjetas, tablas y comparaciones. `__DATA__` se sustituye al generar. |
| `config/recursos.json` | Personas reconocidas, jornada, dedicacion, estado y ausencias. |
| `config/sprints.json` | Calendario y estado de cada sprint. |
| `config/tareas_administrativas.json` | Reglas opcionales de titulos, exclusiones y tiempos para warnings de estimacion administrativa. |
| `config/parametros.json` | Estados finales y tipos de elemento que cuentan como carga. |
| `config/tiempos.json` | Descuentos por festivos y actividades por sprint. |
| `config/ejemplos/*.json` y `config/ejemplos/README.md` | Ejemplos ficticios de ausencias y guia para aplicarlos; no se leen automaticamente. |
| `config/festivos.json` | Fechas festivas que se descuentan si coinciden con el sprint. |
| `entrada/*.csv` | Un unico export de Azure DevOps, con cualquier nombre. |
| `salida/*.xlsx` | Excel generado con fecha/hora; conserva versiones anteriores. |
| `salida/*.html` | Reporte HTML autocontenido mas reciente, con fecha/hora. |
| `InformacionBase/` | Archivos originales de referencia, incluido el CSV; el generador no lee esta carpeta. |
| `Estilos/index.html` y `Estilos/styles.css` | Referencias visuales de la paleta y presentacion; no son los reportes finales. |
| `tests/test_administrativos.py` | Pruebas de categorias, tildes, umbrales, exclusiones, alcance por estado y validacion del archivo administrativo. |
| `tests/test_reglas.py` | Pruebas de reglas, calendario, ausencias, festivos, capacidad e identidad. |
| `tests/test_metricas.py` | Pruebas de agregaciones, controles HTML, serializacion y conservacion de un unico HTML. |
| `tests/test_exportacion.py` | Pruebas integradas de Excel/HTML, alertas y BAT con rutas con espacios. |
| `src/__pycache__/` y `tests/__pycache__/` | Cache de Python; no son configuraciones ni reportes. |

## CSV de entrada

Admite UTF-8 (con o sin BOM) y cp1252, con separador coma o punto y coma. No admite IDs repetidos. Debe tener registros y las columnas exigidas por el generador. `Title` tambien se utiliza para el detalle de casos.

| Columna | Uso |
|---|---|
| `ID` | Identificador unico del elemento. |
| `Parent` | ID del padre; puede estar vacio. Se valida jerarquia, padres no exportados y ciclos. |
| `Work Item Type` | Tipo: Task, Habilitador, User Story, Issue, etc.; determina si aporta a carga segun parametros. |
| `Title` | Titulo del elemento. |
| `Assigned To` | Responsable, normalmente `Nombre <correo>`; se vincula por correo y luego nombre/alias normalizados. |
| `State` | Estado real; es pendiente si no pertenece a `estados_finales`. |
| `Iteration Level 3` | Iteracion; se extrae el numero al final para identificar el sprint. |
| `Story Points` | Horas estimadas para las metricas. Vacio o invalido no se convierte en horas ejecutadas; genera faltante cuando corresponde. |
| `Original Estimate` | Estimacion original informativa; no sustituye Story Points. |
| `Area Level 4` | Celula de Azure. |
| `Area Level 3` | EVC de Azure. |
| Otras columnas | Se conservan en Detalle, por ejemplo Created Date, Closed Date, Iteration Level 1, Value Area y Related Link Count. No determinan la capacidad. |

## Configuraciones JSON: campos

Usar UTF-8 y fechas `YYYY-MM-DD`. `true`/`false` y `null` son valores JSON, sin comillas. Las configuraciones dentro del Excel son copias informativas: para cambiar reglas, editar estos JSON y regenerar.

### recursos.json

Lista de objetos, uno por persona.

| Campo | Tipo | Descripcion |
|---|---|---|
| `documento` | Texto | Identificador del talento para asociar capacidad y exportacion. |
| `nombre` | Texto | Nombre canonico mostrado en los reportes. |
| `correo` | Texto | Correo para vincular Assigned To; puede estar vacio. |
| `alias` | Lista de textos | Nombres o correos alternativos para reconocer a la misma persona. |
| `activo` | Booleano | `false` da capacidad cero y genera alerta si tiene elementos abiertos. No elimina su trabajo historico. |
| `horas_dia` | Numero | Jornada diaria antes de multiplicar por dedicacion, normalmente 9. |
| `dedicacion` | Numero | Fraccion de jornada para el equipo: 1 = 100 %, 0.5 = 50 %. |
| `ausencias` | Lista de objetos | Ausencias generales, cada una con `inicio`, `fin` y `motivo`. |
| `vacaciones` | Objeto o lista | Periodos de vacaciones con `inicio` y `fin`; admite varios intervalos. |
| `incapacidades` | Objeto o lista | Periodos de incapacidad con `inicio` y `fin`; admite varios intervalos. |
| `datos_base` | Objeto | Datos organizativos originales; se copian a Base talentos y se usan en agrupaciones. |
| `datos_base.Doc` | Numero/texto | Documento original de referencia. |
| `datos_base.Talento` | Texto | Nombre original de referencia; el nombre canonico es `nombre`. |
| `datos_base.Celula interna BD` | Texto | Celula interna usada en filtros y Equipos internos; distinta de la celula de Azure. |
| `datos_base.Líder` | Texto | Lider del recurso. |
| `datos_base.Coordinador` | Texto | Coordinador del recurso. |
| `datos_base.EVC` | Texto | EVC interno de referencia; distinto del EVC de Azure. |
| `inicio` / `fin` de un intervalo | Fecha o null | Limites inclusivos. En vacaciones/incapacidades, ambos null significan sin periodo; solo una fecha o fin anterior a inicio genera error. |
| `motivo` de una ausencia | Texto | Razon, por ejemplo permiso. En vacaciones/incapacidades el generador identifica el tipo automaticamente. |

Ejemplo de un recurso:

```json
{
  "documento": "123", "nombre": "Ana Perez", "correo": "ana@example.com",
  "alias": ["Ana P."], "activo": true, "horas_dia": 9, "dedicacion": 1,
  "ausencias": [{"inicio": "2026-10-14", "fin": "2026-10-14", "motivo": "Permiso"}],
  "vacaciones": {"inicio": "2026-10-07", "fin": "2026-10-09"},
  "incapacidades": {"inicio": null, "fin": null},
  "datos_base": {"Doc": "123", "Talento": "Ana Perez", "Celula interna BD": "Equipo", "Líder": "Lider", "Coordinador": "Coordinador", "EVC": "EVC interno"}
}
```

### Ejemplos de vacaciones, incapacidades y ausencias

Los ejemplos estan en [config/ejemplos](config/ejemplos/README.md). Son personas ficticias; el generador no los lee automaticamente ni modifica la configuracion real.

| Archivo | Contenido |
|---|---|
| [recursos_vacaciones.json](config/ejemplos/recursos_vacaciones.json) | Un periodo de vacaciones. |
| [recursos_incapacidades.json](config/ejemplos/recursos_incapacidades.json) | Un periodo de incapacidad. |
| [recursos_ausencias.json](config/ejemplos/recursos_ausencias.json) | Permiso y licencia con motivos. |
| [recursos_varios_periodos.json](config/ejemplos/recursos_varios_periodos.json) | Varias vacaciones/incapacidades y una ausencia general. |

Copia solo los campos de ausencia al recurso real en `config/recursos.json`, cambia las fechas y regenera. Conserva `activo: true` para vacaciones o ausencia temporal; no reemplaces toda tu lista de recursos por una persona ficticia. Los periodos son inclusivos, por dias completos. No se duplican descuentos si coinciden entre si o con festivos.

### sprints.json

Lista de periodos.

| Campo | Tipo | Descripcion |
|---|---|---|
| `sprint` | Texto | Identificador unico, por ejemplo `"255"`. |
| `inicio` / `fin` | Fecha | Fechas inclusivas del periodo. Se cuentan los dias reales; no se fija 10 dias habiles en codigo. |
| `activo` | Booleano | Permite capacidad si es true. Un sprint desactivado tiene capacidad cero. No significa que el sprint sea vigente: eso depende de las fechas y la fecha de corte. |
| `origen` | Texto opcional | `inferido` para calendario extrapolado; si no existe, se muestra configurado. |
| `patron_dias` | Numero opcional | Duracion usada en la inferencia, actualmente 14 dias calendario. |

Si el CSV trae un sprint numerico sin configurar, se extrapola desde el calendario conocido solo cuando todos los periodos conocidos tienen duracion y separacion uniformes. Se guardan los periodos inferidos en este JSON. Un calendario irregular genera error. Un elemento sin identificador no recibe fechas inventadas.

### parametros.json

| Campo | Tipo | Descripcion |
|---|---|---|
| `estados_finales` | Lista de textos | Estados que no se consideran abiertos: actualmente Closed, Removed y Resolved. Coincidencia exacta. |
| `tipos_carga` | Lista de textos | Tipos principales que aportan horas de carga; actualmente Task. |
| `tipos_ha` | Lista de textos | Otros tipos que aportan carga; actualmente Habilitador y User Story. |
| `nota` | Texto | Explicacion informativa; no modifica la formula. |

### tiempos.json

| Campo | Tipo | Descripcion |
|---|---|---|
| `festivo_horas` | Numero | Descuento por festivo habil dentro del sprint, multiplicado por dedicacion; actualmente 9 h. |
| `actividades_por_sprint` | Objeto numerico | Descuentos fijos por persona activa y sprint. Se suman todos sus valores. |
| `actividades_por_sprint.daily` | Numero | 6 h por sprint. |
| `actividades_por_sprint.planning` | Numero | 2 h por sprint. |
| `actividades_por_sprint.retro` | Numero | 1 h por sprint. |
| `actividades_por_sprint.reuniones_y_alineaciones` | Numero | 4 h por sprint. |
| `referencia_limites` | Objeto | Referencias informativas conservadas en el archivo. Actualmente el codigo usa directamente las proporciones 64/64, 55/64 y 46/64; cambiar este objeto no cambia esos calculos. |
| `referencia_limites.horas_totales` | Numero | Base informativa: 64. |
| `referencia_limites.exceso_desde` | Numero | Referencia informativa para exceso: 64. |
| `referencia_limites.punto_medio` | Numero | Referencia informativa del punto medio: 55. |
| `referencia_limites.capacidad` | Numero | Referencia informativa de capacidad objetivo: 46. |

`referencia_limites` no fija horas para todos los sprints y el codigo no lo consulta. Los limites se recalculan por recurso y periodo a partir de su capacidad neta: exceso = 100 %, punto medio = 55/64 (85,94 %), capacidad objetivo = 46/64 (71,875 %). Con capacidad neta de 64 h, los limites son 64, 55 y 46 h; con 77 h, son 77, 66,17 y 55,34 h. Cambiar o retirar ese bloque informativo no cambia el resultado actual.

Los valores de `actividades_por_sprint` y `festivo_horas` si afectan el calculo: son la configuracion comun usada para todos los sprints, no valores individualizados por sprint. La capacidad cambia por calendario, festivos, jornada, dedicacion y ausencias.

Los descuentos de tiempo deben ser numeros no negativos. Las 13 h de actividades son fijas; no se multiplican por dedicacion.

### festivos.json

| Campo | Tipo | Descripcion |
|---|---|---|
| `fechas` | Lista de fechas | Festivos configurados. Solo se descuentan fechas de lunes a viernes dentro del sprint. Se eliminan duplicados; no se descuenta un festivo fuera del periodo. |

### tareas_administrativas.json

Archivo opcional en el directorio `--config`. Si falta, se desactiva solo la clasificacion administrativa; el resto del reporte sigue funcionando. Si existe pero tiene reglas invalidas, la generacion falla con un mensaje que identifica la configuracion.

Identifica tareas por el texto de `Title` y compara sus Story Points con una referencia editable. La referencia inicial es **1 h**, con warning cuando las horas son **mayores que 2 h**. Son referencias de revision de alcance, no tiempos garantizados ni medidas de ejecucion. Se escogieron porque las categorias detectadas en este CSV tienen mediana de Story Points de 1 h. No se usaron tiempos genericos de internet. Las estadisticas iniciales quedan documentadas en el archivo y no se recalculan automaticamente al cambiar el CSV.

| Campo | Tipo | Descripcion |
|---|---|---|
| `version` | Numero | Version documental del esquema: 1. |
| `activo` | Booleano | Habilita o desactiva toda esta validacion. |
| `tipos_elemento` | Lista | Tipos que se clasifican; inicialmente Task, Habilitador y User Story. Coincidencia exacta. |
| `solo_pendientes` | Booleano | false incluye casos cerrados para comparar el historico; true excluye los estados finales de parametros.json. |
| `origen_referencias` | Texto | Explicacion informativa del origen y limites de los tiempos iniciales. |
| `excluir_patrones` | Lista | Expresiones regulares globales que envian tareas candidatas a REVISION_MANUAL. |
| `reglas` | Lista | Categorias configurables. Se pueden agregar, modificar o desactivar sin cambiar codigo. |
| `reglas[].id` | Texto | Identificador unico de la regla. |
| `reglas[].nombre` | Texto | Nombre de categoria mostrado en HTML/Excel. |
| `reglas[].activo` | Booleano | Habilita esa categoria. Gestion de usuarios comienza desactivada por alcance ambiguo. |
| `reglas[].patrones` | Lista | Expresiones regulares de coincidencia sobre el titulo normalizado; basta que coincida una. |
| `reglas[].excluir_patrones` | Lista | Exclusiones adicionales de esa regla. |
| `reglas[].horas_referencia` | Numero | Tiempo orientativo de la operacion simple; positivo y finito. |
| `reglas[].horas_maximas_sin_warning` | Numero | Umbral permitido, igual o superior a la referencia. WARNING solo si Story Points lo supera estrictamente. |
| `reglas[].ejemplos_titulos` | Lista | Ejemplos informativos; no intervienen en la coincidencia. |
| `reglas[].supuesto` | Texto | Alcance esperado, informativo. |
| `reglas[].estadisticas_csv_inicial` | Objeto | Evidencia de la muestra original, informativa. Una tarea puede coincidir con varias categorias. |
| `estadisticas_csv_inicial.coincidencias` | Numero | Casos que coinciden con esa regla activa en la muestra inicial, incluidos los enviados a revision manual. |
| `estadisticas_csv_inicial.con_story_points` | Numero | Coincidencias con estimacion numerica registrada. |
| `estadisticas_csv_inicial.mediana_story_points_horas` | Numero o null | Mediana de esas estimaciones; null si no hay muestra. No es tiempo ejecutado. |

Categorias iniciales: creacion de usuarios, eliminacion/inactivacion de usuarios, ejecucion de scripts preparados, asignacion/ajuste de permisos, cambio/restablecimiento de contrasena y desbloqueo. La categoria general Gestion de usuarios es un ejemplo desactivado.

La normalizacion ignora mayusculas, tildes y separadores: `CREACION`, `creacion` y la variante con tilde se comparan igual; guiones y otros signos se convierten en espacios. Escribir patrones en minusculas y sin tildes. En JSON hay que duplicar las barras de expresiones regulares: `\\b` significa limite de palabra, `\\s+` espacios y `usuarios?` acepta usuario/usuarios. Tambien se admite una frase literal como `"creacion de usuarios"`, aunque sera menos flexible que la expresion de los ejemplos.

Ejemplo de regla que puedes agregar dentro de `reglas`:

```json
{
  "id": "consulta_simple",
  "nombre": "Consulta SQL simple",
  "activo": true,
  "patrones": [
    "\\bconsulta\\s+sql\\s+simple\\b"
  ],
  "excluir_patrones": [
    "\\boptimizacion\\b"
  ],
  "horas_referencia": 0.5,
  "horas_maximas_sin_warning": 1
}
```

Despues de modificar el archivo, regenera con el BAT. Cambiar la referencia no cambia las horas originales del CSV ni las metricas de capacidad; solo la comparacion y los warnings.

| Resultado | Regla |
|---|---|
| `WARNING_ESTIMACION_ALTA` | Una categoria simple y horas mayores que su umbral; agrega ESTIMACION_ADMIN_ALTA en Alertas. |
| `DENTRO_UMBRAL` | Una categoria simple y horas menores o iguales al umbral; cero sigue siendo una estimacion explicita. |
| `SIN_ESTIMACION` | Una categoria simple, pero Story Points vacio o invalido; no se interpreta como cero horas. |
| `REVISION_MANUAL` | Varias categorias o exclusion de complejidad; no se aplica un tiempo simple ni se genera automaticamente un warning de horas altas. |

Las exclusiones iniciales cubren trabajo masivo, migraciones, optimizacion, automatizacion, desarrollo, restauracion/backup, auditorias/revisiones, creacion de otros objetos y varios ambientes. Una tarea como crear usuario y asignar permisos requiere revision manual. El titulo no describe completamente cantidad de usuarios, aprobaciones, pruebas ni complejidad: ajustar reglas y tiempos con el equipo. No se identifica toda tarea administrativa posible.

Los resultados incluyen personas no identificadas para auditar estimaciones. Estas filas no se agregan a los totales de negocio de recursos reconocidos. El filtro Activos/Inactivos del HTML excluye recursos cuyo estado no se conoce.

## Formulas e interpretacion

`Capacidad por persona = max(0, dias lunes a viernes x horas_dia x dedicacion - festivos - ausencias - actividades)`.

Los dias incluyen inicio y fin. Un sprint de 14 dias calendario tiene normalmente 10 dias de lunes a viernes. Festivos = numero de festivos habiles del periodo x `festivo_horas` x dedicacion. Ausencias = dias habiles ausentes no festivos x `horas_dia` x dedicacion; intervalos coincidentes no duplican descuentos. Recurso o sprint inactivo aporta cero.

Ejemplo sin festivos ni ausencias: `10 x 9 x 1 - (6 + 2 + 1 + 4) = 77 h por persona`. Para 27 recursos con esa configuracion: `27 x 77 = 2.079 h`. Incluye personas sin horas registradas, segun los filtros seleccionados.

- Asignadas: Story Points de los tipos de carga, en cualquier estado, incluidas cerradas.
- Cerradas: estimaciones de elementos Closed, no horas ejecutadas.
- Pendientes: estimaciones de elementos que no tienen un estado final configurado.
- Otros finales: estimaciones de Removed/Resolved u otros estados finales distintos de Closed.
- Todos los tipos: estimaciones de todos los tipos, aunque no aporten a carga.
- Carga: horas comparadas / capacidad; 100 % equivale a la capacidad calculada.
- Diferencia: capacidad menos asignadas, incluidas cerradas. No demuestra tiempo sin trabajo ni productividad.
- Limites por sprint: exceso desde capacidad x 64/64; punto medio x 55/64; objetivo x 46/64.

La capacidad historica aplica la configuracion **actual** de activo, jornada, dedicacion y ausencias. No existe historial de cambios por persona, asi que no confirma la capacidad real del pasado. En sprints finalizados, capacidad no significa horas libres hoy. Un CSV incompleto puede mostrar cero horas aunque haya existido trabajo. Sprints sin fechas tienen capacidad desconocida; las comparaciones con capacidad excluyen sus horas cuando corresponde y lo indican.

## Alertas de disponibilidad y tareas activas

Se considera abierto cualquier estado fuera de `estados_finales`, no solo el literal Active. Las alertas se evaluan por elemento, de cualquier tipo, con o sin horas, para personas reconocidas. Elementos cerrados no generan alerta de disponibilidad.

| Alerta | Regla | Accion |
|---|---|---|
| ESTIMACION_ADMIN_ALTA | Una categoria administrativa simple supera el umbral de horas configurado. | Revisar alcance, procedimiento y Story Points; no implica sobreestimacion confirmada. |
| RECURSO_NO_DISPONIBLE | Persona inactiva o con vacaciones, incapacidad o ausencia que incluye la fecha de corte, con un elemento abierto. Muestra el motivo y las fechas. | Revisar el estado del caso, reasignar o reprogramar. |
| AUSENCIA_EN_SPRINT | Persona activa con elemento abierto cuyo sprint coincide total o parcialmente con un intervalo de ausencia. | Revisar asignacion; no hay fechas de ejecucion por tarea y la coincidencia no prueba que se trabajo durante la ausencia. |
| PENDIENTE_SPRINT_FINALIZADO | Elemento abierto en un sprint que termino antes de la fecha de corte. | Cerrar si corresponde, reprogramar o reasignar. |
| SOBRECARGA | Horas pendientes positivas alcanzan la capacidad o existe capacidad cero. | Revisar compromisos y configuracion. |
| SIN_ESTIMACION | Elemento pendiente de carga sin Story Points validos. | Completar estimacion. |
| RECURSO_NO_CATALOGADO / RECURSO_AMBIGUO | No se identifica de forma unica al responsable. | Corregir nombre, correo o alias. |
| SIN_ASIGNAR | Caso abierto sin responsable. | Asignar responsable. |
| SPRINT_SIN_CONFIGURACION / SPRINT_INACTIVO | Sprint sin fechas conocidas o desactivado. | Revisar calendario. |
| PARENT_NO_EXPORTADO / CICLO_PARENT | Padre fuera del CSV o jerarquia circular. | Revisar export; un padre ausente no se inventa ni excluye trabajo valido. |
| VALOR_INVALIDO | Estimacion negativa o no numerica. | Corregir dato de origen. |

Un mismo caso puede generar dos alertas de disponibilidad, por fecha de corte y por coincidencia con el sprint. El HTML distingue **casos unicos** y **numero de alertas**. No se considera que un recurso historicamente haya estado inactivo solo porque hoy lo este; se informa segun la configuracion disponible.

## HTML: secciones y campos

El HTML es autocontenido y funciona al abrirlo en un navegador, sin servidor ni conexion externa. Usa UTF-8 para conservar tildes y separadores. Regenerar actualiza los datos; editar JSON no cambia un HTML ya creado.

| Seccion / campo | Descripcion |
|---|---|
| Fecha de corte / vista inicial | Fecha del analisis. Inicia en un sprint vigente con horas o, en su ausencia, en el ultimo sprint del CSV con horas. |
| Filtro Sprint | Un sprint o todos los del CSV; no incluye automaticamente capacidad de futuros sin registros. |
| Celula interna / Disponibilidad / Recurso / Buscar | Seleccion por equipo, activo/inactivo, persona y coincidencia de nombre. Activo/inactivo no describe vacaciones. |
| Tarjetas Recursos, Asignadas, Cerradas, Pendientes, Disponibles | Conteos y horas de la seleccion; Disponibles es capacidad calculada del periodo, no horas libres actuales. |
| Trabajo pendiente y capacidad del periodo seleccionado | Por recurso: pendientes/capacidad, porcentaje, horas asignadas, cerradas y pendientes, y diferencia capacidad-asignadas. Aclara alcance historico y estimaciones. |
| Carga y capacidad por sprint | Compara cada sprint del CSV; aplica filtros de recursos, pero no el filtro superior de sprint. Selector Asignadas/Solo pendientes, porcentaje etiquetado, capacidad, desglose y diferencia usando asignadas incluidas cerradas. |
| Recursos con horas / con capacidad por sprint | Personas con horas positivas del indicador elegido y personas con capacidad positiva; pueden ser distintas. |
| Analizar este sprint | Selecciona el periodo y lleva a la tabla de recursos. |
| Horas asignadas frente a disponibilidad | Por persona: carga asignada incluida la cerrada frente a capacidad del periodo. No prueba exceso de pendientes. |
| Carga frente a capacidad por celula | Cantidad de recursos, horas seleccionadas, capacidad y porcentaje por celula interna. Selector Asignadas/Pendientes; boton para inspeccionar sus recursos. |
| Casos abiertos de sprints anteriores | Sprints terminados antes del corte. Independiente del filtro superior de sprint; respeta filtros de recursos. Selector todos/ultimo finalizado. |
| Resumen de casos abiertos por sprint | Sprint, fin, casos abiertos, recursos, horas registradas y casos sin estimacion. |
| Detalle de casos abiertos | Recurso, sprint, fin, dias desde fin, ID, tipo, titulo, estado y horas. Incluye todos los tipos. |
| Tareas administrativas y estimaciones | Casos detectados por Title, warnings, revision manual y sin estimacion; filtro por resultado. Tabla con recurso, sprint, ID, tipo, titulo, estado, categoria, Story Points, referencia, umbral, veces referencia, resultado y detalle. Aplica filtros de sprint y recursos. |
| Alertas de recursos no disponibles con casos abiertos | Casos unicos, recursos y numero de alertas; tabla por recurso, sprint, ID, tipo, titulo, estado, horas, tipo de alerta y motivo/alcance. Aplica filtro de sprint y recursos. |
| Alertas de carga y disponibilidad | Recurso, sprint, aviso y accion sugerida. Incluye sobrecarga, inactividad, falta de estimacion y ausencias. |
| Recursos identificados | Recurso, celula, activo, horas asignadas/cerradas/pendientes, capacidad, porcentaje de carga pendiente y pendientes sin estimacion. |
| Como leer los indicadores | Alcance del CSV, tipos/estados, capacidad, dedicacion y limites del historico. |

Los botones Limpiar filtros y Mostrar todas las celulas restablecen filtros de recursos. Una comparacion de pendientes al 0 % no implica ausencia de trabajo: el desglose mantiene visibles las horas cerradas. Los estados excluidos del cuadro de casos abiertos se leen de la configuracion al generar.

## Excel: hojas y campos

Usar los controles **+** para desplegar detalles. No sumar subtotales y sus hijos: representan los mismos datos. Las graficas estan en el HTML; el Excel usa tablas y jerarquias. Los totales de negocio solo incluyen personas identificadas, activas e inactivas. Detalle y Alertas conservan registros no vinculados para diagnostico.

Las columnas originales del CSV en Detalle conservan su significado y estan descritas en la seccion CSV. Las siguientes tablas describen las columnas de todas las hojas generadas; las columnas `SP <numero>` varian segun los sprints del CSV.

### Tablero

Indicadores y contribuciones desplegables por recurso, sprint y tipo.

| Campo | Descripcion |
|---|---|
| Indicador / Recurso | Nombre del indicador o de su contribuyente desplegado. |
| Valor | Valor del indicador o aporte de la fila. |
| Periodo | TOTAL, periodo o nivel de detalle; identifica subtotales. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |

### Detalle

Todas las columnas del CSV mas vinculacion y datos calculados. Conserva elementos no reconocidos.

| Campo | Descripcion |
|---|---|
| ID | Identificador del elemento en Azure. |
| Parent | ID del padre exportado. |
| Work Item Type | Tipo original del elemento. |
| Title | Titulo original. |
| Assigned To | Responsable original del CSV. |
| State | Estado original del caso. |
| Iteration Level 1 | Iteracion superior original; informativa. |
| Created Date | Fecha de creacion original; informativa. |
| Iteration Level 3 | Iteracion de la que se obtiene el numero de sprint. |
| Closed Date | Fecha original de cierre; no determina horas ejecutadas. |
| Story Points | Horas estimadas originales. |
| Area Level 4 | Celula original de Azure. |
| Original Estimate | Estimacion original informativa. |
| Value Area | Clasificacion original del valor. |
| Related Link Count | Conteo de enlaces original. |
| Area Level 3 | EVC original de Azure. |
| Celula Azure | Celula de la estructura de Azure, Area Level 4. |
| EVC Azure | EVC de Azure, Area Level 3. |
| Nombre asignado | Nombre separado de Assigned To. |
| Correo asignado | Correo extraido de Assigned To. |
| Documento talento | Documento de la persona vinculada; vacio si no se reconoce. |
| Talento vinculado | Nombre canonico de la persona identificada. |
| Pendiente | Si cuando el estado no es final; No cuando es final. |
| Sprint ID | Numero al final de Iteration Level 3; vacio si no se obtiene. |
| Ancestros exportados | Cadena de padres incluidos en el CSV; no inventa padres faltantes. |
| Horas originales numericas | Original Estimate convertido a numero o vacio si invalido. |
| Horas Story Points numericas | Story Points convertido a numero o vacio si invalido. |

### Alertas

Todas las validaciones por elemento y alertas agregadas; pueden incluir casos no vinculados.

| Campo | Descripcion |
|---|---|
| Tipo | Codigo de alerta o tipo de elemento, segun la hoja. |
| ID | Identificador del elemento en Azure. |
| Parent | ID del padre exportado. |
| Tipo elemento | Tipo del caso que genera la alerta. |
| Responsable | Assigned To original o recurso para alertas agregadas. |
| Estado | Estado del caso en el CSV. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Celula Azure | Celula de la estructura de Azure, Area Level 4. |
| EVC Azure | EVC de Azure, Area Level 3. |
| Detalle | Explicacion concreta de la alerta. |

### Base talentos

Copia de los campos originales de datos_base.

| Campo | Descripcion |
|---|---|
| Doc | Documento original de datos_base. |
| Talento | Nombre del talento; en Base talentos conserva el dato original. |
| Celula interna BD | Celula interna de configuracion. |
| Líder | Lider original de datos_base. |
| Coordinador | Coordinador de configuracion. |
| EVC | EVC interno de referencia. |

### Config recursos

Copia informativa de la configuracion de personas.

| Campo | Descripcion |
|---|---|
| Documento | Identificador de la persona. |
| Talento | Nombre del talento; en Base talentos conserva el dato original. |
| Correo | Correo configurado. |
| Alias | Nombres/correos alternativos separados por coma. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Horas dia | Jornada diaria configurada. |
| Dedicacion al equipo | Fraccion de jornada mostrada como porcentaje. |
| Ausencias | Intervalos y motivos generales serializados como JSON. |
| Vacaciones | Intervalos de vacaciones serializados como JSON. |
| Incapacidades | Intervalos de incapacidad serializados como JSON. |

### Config sprints

Copia del calendario usado.

| Campo | Descripcion |
|---|---|
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Inicio | Fecha inicial inclusiva del sprint. |
| Fin | Fecha final inclusiva del sprint. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Origen | Configurado o inferido. |

### Config festivos

Fechas configuradas.

| Campo | Descripcion |
|---|---|
| Fecha | Fecha festiva configurada. |

### Config tiempos

Valores de descuentos usados.

| Campo | Descripcion |
|---|---|
| Grupo | Descuentos o actividades por sprint. |
| Concepto | Concepto de descuento. |
| Horas | Horas del concepto o estimacion del caso, segun contexto. |

### Carga recursos

Totales por persona con detalle por sprint; periodos con registros identificados del CSV.

| Campo | Descripcion |
|---|---|
| Documento | Identificador de la persona. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas otros finales | Estimaciones de tipos de carga en finales distintos de Closed. |
| Disponibles en sprints del CSV | Capacidad acumulada solo en periodos del CSV con fechas conocidas; en hijos, capacidad del sprint. |
| Carga pendiente con fechas | Pendientes de periodos con fechas / capacidad de esos mismos periodos. |
| Pendientes sin estimacion | Cantidad de casos pendientes de carga sin estimacion valida. |
| Sprints sin fechas | Cantidad de sprints con carga asignada sin calendario; en hijos, 1 si no hay fechas. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |

### Historico recursos

Horas historicas por persona, desplegables por periodo.

| Campo | Descripcion |
|---|---|
| Documento | Identificador de la persona. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas otros finales | Estimaciones de tipos de carga en finales distintos de Closed. |
| Elementos sin estimacion | Cantidad de casos de carga sin horas validas, de cualquier estado. |
| Horas todos los tipos | Story Points sumado sin restringir tipo ni estado. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |

### Matriz por estado

Todos los tipos: total y categorias por estado real, con recursos debajo. No es la matriz de tipos de carga.

| Campo | Descripcion |
|---|---|
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Celula interna | Celula interna del recurso. |
| Coordinador | Coordinador de configuracion. |
| Lider | Lider del recurso. |
| SP 234 | Horas de todos los tipos para el sprint 234, en la categoria de estado de la fila. |
| SP 236 | Horas de todos los tipos para el sprint 236, en la categoria de estado de la fila. |
| SP 237 | Horas de todos los tipos para el sprint 237, en la categoria de estado de la fila. |
| SP 238 | Horas de todos los tipos para el sprint 238, en la categoria de estado de la fila. |
| SP 239 | Horas de todos los tipos para el sprint 239, en la categoria de estado de la fila. |
| SP 240 | Horas de todos los tipos para el sprint 240, en la categoria de estado de la fila. |
| SP 241 | Horas de todos los tipos para el sprint 241, en la categoria de estado de la fila. |
| SP 242 | Horas de todos los tipos para el sprint 242, en la categoria de estado de la fila. |
| SP 243 | Horas de todos los tipos para el sprint 243, en la categoria de estado de la fila. |
| SP 244 | Horas de todos los tipos para el sprint 244, en la categoria de estado de la fila. |
| SP 245 | Horas de todos los tipos para el sprint 245, en la categoria de estado de la fila. |
| SP 246 | Horas de todos los tipos para el sprint 246, en la categoria de estado de la fila. |
| SP 247 | Horas de todos los tipos para el sprint 247, en la categoria de estado de la fila. |
| SP 248 | Horas de todos los tipos para el sprint 248, en la categoria de estado de la fila. |
| SP 249 | Horas de todos los tipos para el sprint 249, en la categoria de estado de la fila. |
| SP 250 | Horas de todos los tipos para el sprint 250, en la categoria de estado de la fila. |
| SP 251 | Horas de todos los tipos para el sprint 251, en la categoria de estado de la fila. |
| SP 252 | Horas de todos los tipos para el sprint 252, en la categoria de estado de la fila. |
| SP 253 | Horas de todos los tipos para el sprint 253, en la categoria de estado de la fila. |
| Total | Suma de las columnas de sprint de esa fila; no sumar categoria y sus hijos. |

### Metricas sprint

Totales del equipo por sprint y aportes por recurso.

| Campo | Descripcion |
|---|---|
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Elementos de carga | Numero de elementos de los tipos configurados de carga. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas otros finales | Estimaciones de tipos de carga en finales distintos de Closed. |
| Elementos sin estimacion | Cantidad de casos de carga sin horas validas, de cualquier estado. |
| Horas disponibles | Capacidad calculada del periodo; no horas libres actuales. |
| Carga pendiente | Horas pendientes / capacidad del sprint. |
| Carga asignada | Horas asignadas / capacidad del sprint. |
| Activos con exceso de pendientes | Cantidad de activos con pendientes positivos que alcanzan capacidad; incluye capacidad cero. |
| Activos sin horas registradas | Activos con capacidad positiva y cero horas asignadas de carga. |
| Inactivos con pendientes | Inactivos con horas pendientes positivas; las alertas por caso tambien incluyen los sin horas. |
| Porcentaje horas cerradas | Cerradas / asignadas; vacio si no hay asignadas. |
| Pendientes sin estimacion | Cantidad de casos pendientes de carga sin estimacion valida. |
| Datos del CSV | Indica si existen elementos de carga, no si el sprint carecio realmente de trabajo. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |

### Detalle celulas EVC

Jerarquia EVC de Azure -> celula -> sprint -> recurso.

| Campo | Descripcion |
|---|---|
| EVC / celula / sprint / recurso | Etiqueta del nivel de agrupacion; expandir para ver sus hijos. |
| Registros Azure | Numero de IDs vinculados de todos los tipos; no numero de personas. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas todos los tipos | Story Points sumado sin restringir tipo ni estado. |
| Nivel | EVC, Celula, Sprint o Recurso. |

### Celulas y EVC

Totales jerarquicos organizativos del CSV.

| Campo | Descripcion |
|---|---|
| EVC / celula / sprint / recurso | Etiqueta del nivel de agrupacion; expandir para ver sus hijos. |
| Celula Azure | Celula de la estructura de Azure, Area Level 4. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Registros Azure | Numero de IDs vinculados de todos los tipos; no numero de personas. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas todos los tipos | Story Points sumado sin restringir tipo ni estado. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |

### Equipos internos

Celula interna -> recurso -> sprint.

| Campo | Descripcion |
|---|---|
| Celula / recurso / sprint | Etiqueta jerarquica del equipo, persona o periodo. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Activo / cantidad | En subtotal de celula, cantidad de personas; en recurso/sprint, 1 activo o 0 inactivo. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Disponibles | Capacidad calculada de los periodos de la fila. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |

### Sprint vigente

Periodos activos que contienen la fecha de corte; puede no tener trabajo registrado.

| Campo | Descripcion |
|---|---|
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Horas asignadas | Suma de estimaciones de tipos de carga en cualquier estado, incluidas cerradas. |
| Horas cerradas | Estimaciones de tipos de carga en estado Closed. |
| Horas pendientes | Estimaciones de tipos de carga fuera de estados finales. |
| Horas disponibles | Capacidad calculada del periodo; no horas libres actuales. |
| Estado pendientes | SIN FECHAS, SIN CAPACIDAD, EXCESO, ESTIMACION INCOMPLETA, SOBRE PUNTO MEDIO, SOBRE CAPACIDAD OBJETIVO o BAJO CAPACIDAD OBJETIVO. |
| Pendientes sin estimacion | Cantidad de casos pendientes de carga sin estimacion valida. |

### Calidad datos

Conteos de validaciones de recursos identificados.

| Campo | Descripcion |
|---|---|
| Tipo alerta | Codigo de validacion. |
| Eventos de recursos identificados | Cantidad de alertas vinculadas o agregadas, no casos unicos. |

### Guia de lectura

Glosario dentro del libro.

| Campo | Descripcion |
|---|---|
| Concepto | Concepto de descuento. |
| Descripcion | Explicacion del concepto. |

### Pendientes anteriores

Casos abiertos de todos los tipos en sprints finalizados antes del corte.

| Campo | Descripcion |
|---|---|
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| Fin sprint | Fecha final del sprint anterior. |
| Dias desde fin | Dias calendario entre fin del sprint y fecha de corte. |
| ID | Identificador del elemento en Azure. |
| Tipo | Codigo de alerta o tipo de elemento, segun la hoja. |
| Titulo | Titulo del caso. |
| Estado | Estado del caso en el CSV. |
| Horas registradas | Story Points del caso; vacio si no tiene estimacion. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Celula interna | Celula interna del recurso. |
| Celula Azure | Celula de la estructura de Azure, Area Level 4. |
| EVC Azure | EVC de Azure, Area Level 3. |

### Disponibilidad y tareas

Alertas de recursos no disponibles con elementos abiertos, incluso sin horas; una fila por alerta.

| Campo | Descripcion |
|---|---|
| Recurso | Nombre de persona, categoria o etiqueta de subtotal segun el nivel. |
| Activo | Estado configurado; falso implica capacidad cero. |
| Celula interna | Celula interna del recurso. |
| Sprint | Identificador del periodo; puede contener TOTAL RECURSO, TODOS u otra etiqueta en filas de subtotal. |
| ID | Identificador del elemento en Azure. |
| Tipo elemento | Tipo del caso que genera la alerta. |
| Titulo | Titulo del caso. |
| Estado | Estado del caso en el CSV. |
| Horas registradas | Story Points del caso; vacio si no tiene estimacion. |
| Tipo alerta | Codigo de validacion. |
| Motivo y alcance | Inactividad o intervalo de ausencia; distingue fecha de corte y coincidencia con sprint. |

### Tareas administrativas

Una fila por elemento clasificado por Title, incluso cerrado o de persona no identificada. Las horas originales se conservan. Casos con varias categorias o exclusiones quedan para revision manual sin referencia numerica.

| Campo | Descripcion |
|---|---|
| Recurso | Nombre canonico si se reconoce; nombre original o Sin asignar en otro caso. |
| Identificado | Indica si se vinculo con recursos.json. |
| Activo | Estado del recurso reconocido; vacio si no se conoce. |
| Celula interna | Equipo configurado o No identificado. |
| Sprint | Identificador del periodo. |
| ID | ID del elemento. |
| Tipo elemento | Tipo original, limitado por tipos_elemento. |
| Titulo | Title original del CSV. |
| Estado | State original. |
| Horas registradas | Story Points numerico; vacio si no hay estimacion valida. |
| Reglas | IDs de las reglas que coincidieron. |
| Categoria | Nombres de las categorias coincidentes. |
| Horas referencia | Tiempo orientativo; vacio para revision manual. |
| Umbral warning | Maximo permitido sin warning; se alerta cuando las horas lo superan. |
| Diferencia referencia | Story Points menos referencia; puede ser negativa. No es diferencia respecto al umbral. |
| Veces referencia | Story Points / referencia; vacio si no aplica. |
| Resultado | WARNING_ESTIMACION_ALTA, DENTRO_UMBRAL, SIN_ESTIMACION o REVISION_MANUAL. |
| Detalle | Explicacion y accion de revision. |

### Config tareas administrativas

Copia informativa de las reglas, incluidas las desactivadas. Para cambiar los tiempos, editar JSON y regenerar.

| Campo | Descripcion |
|---|---|
| ID regla | Identificador unico. |
| Categoria | Nombre de regla. |
| Activa | Habilitacion de la regla individual; la configuracion global puede desactivar todas. |
| Patrones | Lista JSON de expresiones regulares de coincidencia. |
| Exclusiones | Lista JSON de exclusiones de esa regla; las globales estan en el archivo de configuracion. |
| Horas referencia | Tiempo inicial orientativo. |
| Umbral warning | Horas maximas sin warning. |

## Verificacion y problemas frecuentes

```powershell
python -m unittest discover -s tests -v
```

Las pruebas cubren clasificacion administrativa por titulo, umbrales, tareas complejas y estimaciones faltantes, ademas de formulas, exportacion, estados finales, ausencia de horas, alertas de vacaciones/incapacidades/ausencias/inactividad, rutas con espacios del BAT y nombres sincronizados de Excel/HTML.

- Cero o varios CSV: dejar exactamente uno en el directorio de entrada.
- Recurso no reconocido: revisar correo, nombre y alias. Una coincidencia ambigua no se asigna automaticamente.
- Ausencia incompleta/invertida: completar ambas fechas y revisar su orden.
- Calendario irregular: corregir fechas antes de inferir nuevos sprints.
- Error al guardar: comprobar permisos de salida y cerrar el Excel si se repite el mismo nombre.
- Cambios de configuracion sin efecto: regenerar y abrir el HTML mas reciente.
- Capacidad historica alta frente a pocas horas: revisar cobertura del CSV y configuracion actual; no concluir baja productividad.
