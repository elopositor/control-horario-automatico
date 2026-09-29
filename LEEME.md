# Sincronización automática de fichajes → «Copia de horario 2026.xlsx»

Vuelca los fichajes registrados en el portal de RRHH dentro del libro de horario del escritorio,
sin intervención manual.

## Llevarlo a otro equipo

**Copia esta carpeta entera** donde quieras (no tiene por qué ser la misma ruta) y ejecuta:

```powershell
powershell -ExecutionPolicy Bypass -File .\Instalar.ps1
```

El instalador se encarga de todo:

1. Comprueba PowerShell, Excel y que no falte ningún fichero del paquete.
2. **Busca el libro de horario**; si no lo encuentra, abre un diálogo para que lo señales.
3. Verifica que el libro tiene la estructura esperada (hojas semanales con «Entrada» en `B11`
   y «Salida» en `B16`) — así no se configura por error un libro que no es.
4. Comprueba que se llega a la intranet de fichajes. Si no hay red corporativa, avisa pero
   continúa: funcionará al volver a la oficina.
5. Guarda la ruta en `configuracion.json`.
6. Registra las dos tareas programadas y crea los accesos directos.

**La ruta se pregunta una sola vez.** No se puede pedir en cada ejecución porque las tareas
programadas corren desatendidas y no habría nadie para responder.

Para cambiar de libro más adelante: `.\Instalar.ps1 -RutaLibro "D:\otra\ruta.xlsx"`.
Para quitarlo del equipo: `.\Desinstalar.ps1` (borra tareas y accesos directos; **no** toca el
libro ni las copias de seguridad).

### Dónde tiene que estar el Excel

**Donde quieras.** No hay ninguna ruta obligatoria: lo que manda es `configuracion.json`.

Si está en el escritorio con su nombre habitual, se detecta solo. Si no, el instalador lo
pregunta. Y si algún día lo mueves, la siguiente ejecución **lo vuelve a buscar** por el
escritorio (incluido el de OneDrive) y `Documentos`, y reescribe la configuración sola.

La carpeta de copias de seguridad **se deriva de dónde esté el libro**, así que si lo mueves,
las copias lo siguen sin tocar nada.

### Requisitos en el equipo destino

| | |
|---|---|
| Windows con PowerShell 5.1 | viene de serie |
| Microsoft Excel | necesario, se automatiza por COM |
| Sesión de dominio corporativo | la intranet usa autenticación integrada |
| Permisos de administrador | **no** hacen falta: las tareas se registran para tu usuario |

No hay que instalar **nada** más: ni módulos, ni complementos de Office.

## Qué hace

1. Entra en la URL de tu portal de fichajes con **autenticación integrada de
   Windows** (no hace falta contraseña: el servidor acepta tu sesión de dominio).
2. Lee la tabla de fichajes de la semana en curso y, mediante *postback*, de las semanas
   anteriores que se le pidan.
3. Localiza en el libro la hoja de cada semana y escribe las marcas en la columna C.

## Reglas de escritura

| Situación | Comportamiento |
|---|---|
| Fichaje real en la web | Se escribe |
| Salida aún no fichada (día en curso) | **No se toca** — tu previsión se respeta |
| Días futuros con previsión | **No se tocan** |
| Valor distinto al fichaje real | Se **corrige**, también en semanas pasadas |
| Cualquier celda con fórmula | Nunca se toca |

Nunca se borra nada.

### Por qué se revisan 4 semanas

Las tareas corren con `-SemanasAtras 3 -Forzar`, es decir: repasan la semana en curso y las
**tres anteriores**, corrigiendo lo que no coincida con el fichaje oficial.

Es por las **solicitudes de incidencia**: cuando un fichaje sale mal y se pide su corrección,
esta tarda unos días en aparecer en el sistema. Con una ventana de una sola semana, y
limitándose a rellenar huecos, la corrección nunca llegaría al libro.

Como el saldo se encadena de una semana a la siguiente, al corregir un fichaje antiguo se
recalcula **`C5` de todas las semanas posteriores** de la ventana, en orden cronológico.

Si prefieres volver al comportamiento conservador, quita `-Forzar` de las dos tareas
programadas: entonces las semanas pasadas solo rellenarían celdas vacías.

## Mapa del libro

Una hoja por semana. Las **únicas** celdas manuales son de la columna C:

```
base = 10 + 8*d          d = 0 (lunes) … 4 (viernes)

base+1  Entrada           base+2  Salida desayuno    base+3  Entrada desayuno
base+4  Salida comer      base+5  Entrada comer      base+6  Salida
```

La hoja se identifica por **su nombre** (`24-08-26 a …`), no por las etiquetas internas:
varias hojas arrastran erratas de copiar-pegar (las de agosto siguen diciendo «julio» en
`I1` y «Lunes 27» en `B10`). Los datos de esas hojas sí son correctos.

### Cómo se interpretan las marcas

Los fichajes van **en pares** entrada/salida:

- **Número par de marcas** → la jornada está cerrada y la última es la salida del día.
- **Número impar** → sigues dentro, y la última marca es una **vuelta de pausa**, no la salida.
  Tomarla por salida machacaría la hora prevista con la del café.

Las parejas intermedias son pausas, y se clasifican así:

| Pausa | Va a |
|---|---|
| Hasta 60 min (`-MaxMinutosDescanso`), a la hora que sea | Salida/Entrada **desayuno**, el café |
| Más de 60 min y empieza a las 12:00 o después (`-HoraMinimaComida`) | Salida/Entrada **comida** |
| Más de 60 min pero **antes** de esa hora | No se registra, solo se avisa |

El café va por duración y no por hora del día porque una comida temprana (12:30 a 13:50) se
clasificaría mal con un corte horario.

La comida, en cambio, necesita además la hora: una ausencia larga de primera hora (una gestión,
el médico) no es la comida. El 24-09-2026 hubo una pausa de 08:25 a 13:13 y, al tomarse por
comida, machacó la de verdad, que fue de 15:06 a 16:40. Esa pausa temprana no se anota en
ninguna parte -sólo caben un café y una comida por día-, así que si hay que descontarla se hace
a mano.

Si un día aparecen dos pausas del mismo tipo se avisa y se queda la primera.

El descuento lo calcula el propio libro: los primeros 15 minutos de pausa (`K3`) son de
cortesía y no restan; solo cuenta el exceso. Un café de 30 minutos retrasa la salida 15.

## Cuándo se ejecuta

Dos tareas programadas:

| Tarea | Cuándo | Qué hace |
|---|---|---|
| **Resumen horario** | al iniciar sesión (+3 min) | sincroniza **y** avisa en pantalla |
| **Sincronizar horario** | 09:30, 12:00 y 15:20 | solo sincroniza, en silencio |

Los horarios están elegidos para caer dentro de la jornada: a las 18:45 el equipo ya suele
estar apagado y esa pasada no llegaba a ejecutarse nunca.

Con `StartWhenAvailable`, si el equipo estaba apagado a la hora prevista, la ejecución se
recupera al encenderlo. Si no hay red, no se ejecuta y se reintenta en la siguiente pasada.

**Si sales después de las 15:20**, esa salida no se recoge ese mismo día: entra sola a la
mañana siguiente, en la pasada del inicio de sesión. No se pierde nada.

## El aviso de la mañana

`Mostrar-Resumen.ps1` lanza una notificación de Windows con el saldo de horas y la hora de
salida. **Solo lee el libro; nunca escribe.**

```powershell
& "$env:USERPROFILE\Scripts\Horario\Mostrar-Resumen.ps1" -Consola          # por pantalla
& "$env:USERPROFILE\Scripts\Horario\Mostrar-Resumen.ps1" -Fecha 2026-08-28 # otro día
```

### A mano, cuando quieras

Accesos directos en el escritorio:

| Acceso directo | Qué hace |
|---|---|
| **Mis horas** | Actualiza los fichajes y muestra cómo vas. Para mirar a media tarde |
| **Nueva semana** | Crea la hoja de la siguiente semana, con las fechas ya puestas |
| **Actualizar horario (portapapeles)** | Para cuando estás fuera y has copiado la tabla |

«Nueva semana» calcula sola cuál toca (la siguiente a la última que tenga hoja). Para una
semana concreta: `.\Nueva-Semana.ps1 -Fecha 2026-09-21` (vale cualquier día de esa semana).

No hace falta para la semana en curso —las tareas ya la crean cada lunes—, sirve para
adelantarse.

## El botón dentro del Excel

El libro es **`Copia de horario 2026.xlsm`** (con macros) y lleva un botón **«Nueva semana»**
en la hoja de la semana, arriba a la derecha. Hace lo mismo que el acceso directo, pero sin
salir de Excel. Como las hojas nuevas se crean copiando la anterior, **el botón se hereda solo**.

El código está en el módulo `ModHorario` y su copia de referencia en `ModHorario.bas`, dentro
de la carpeta de scripts. Para reinstalarlo tras un cambio:

```powershell
$wb.VBProject.VBComponents.Remove($wb.VBProject.VBComponents.Item('ModHorario'))
$wb.VBProject.VBComponents.Import('...\ModHorario.bas')
```

Requiere que el Centro de confianza de Excel tenga marcado *«Confiar en el acceso al modelo de
objetos de proyectos de VBA»* (`AccessVBOM = 1`).

**La lógica y el mensaje van separados**: `CrearSemanaMsg()` hace el trabajo y **devuelve** el
texto; `CrearSemanaSiguiente()` es la que enseña el `MsgBox`. Así se puede probar la macro por
automatización — un `MsgBox` dejaría el proceso colgado esperando un clic.

El `.bas` se guarda en **Windows-1252** y sin acentos: VBA no importa bien UTF-8.

«Mis horas» sincroniza primero, así que si ya has fichado la salida la cuenta. Equivale a:

```powershell
& "$env:USERPROFILE\Scripts\Horario\Ver-MisHoras.ps1"
& "$env:USERPROFILE\Scripts\Horario\Ver-MisHoras.ps1" -SinSincronizar   # solo lee el Excel
```

### De dónde salen las horas

El objetivo diario `C(base)` **ya lleva descontado el saldo acumulado**: con +1:25 a favor, el
objetivo de hoy no es 7:00 sino 5:35. De ahí:

```
saldo previo   = C4 - C(base)                                (negativo => debes horas)
salida mínima  = D(base+1) + G(base+2) + G(base+4) + C(base) → deja el saldo a cero
jornada plena  = D(base+1) + G(base+2) + G(base+4) + C4      → mantiene el saldo
```

No se usa la fórmula `M(base)` («Hora óptima de salida») del libro porque devuelve `?` siempre
que no se sale a comer, que es el caso habitual.

### Qué mensaje da cada día

Como la preferencia es **conservar margen y gastarlo, si acaso, el viernes**:

- **Lunes a jueves** — se destaca la jornada completa: *«Sal a las 15:59 y mantienes tu margen
  / No bajes de las 14:34 o te quedarás en negativo»*. La hora mínima es exactamente el punto
  donde el saldo cruza a cero: por debajo se entra en negativo, y entre esa hora y la jornada
  completa se gasta parte del margen sin llegar a deber horas.

Mientras la jornada siga abierta se añade una línea con **cómo quedaría el día si se sale a la
hora prevista** que haya escrita: *«Si sales a las 14:43 tu balance de horas queda en −0:06»*.
Es `salida prevista − entrada − pausas − objetivo`. Desaparece en cuanto el día está cerrado,
porque entonces ya se informa del resultado real.
- **Viernes** — se destaca la mínima: *«Puedes salir ya a las 15:01 y gastar el margen / A las
  15:10 lo acumulas para la semana que viene»*.
- **Debiendo horas** — se invierte: la mínima pasa a ser más tardía que la jornada normal.
- **Jornada ya cerrada** — informa de cómo quedó el día.
- **Sábado y domingo** — no avisa.

### Detalles de implementación

El toast se construye con **XML propio**, no con `GetTemplateContent`: `ToastText04` solo trae
3 nodos de texto (título + 2 líneas) y aquí hacen falta más. Si la notificación fallara, el
respaldo es `WScript.Shell.Popup` **con cierre automático a los 45 s** — un `MessageBox`
normal dejaría la tarea programada colgada esperando un clic.

Windows solo muestra **dos** elementos `<text>` y colapsa el resto en «+N de notificaciones».
Por eso el título va en un `<text>` y **todo el cuerpo en un único `<text>`** con saltos
`&#10;` y `hint-maxLines='5'`.

Pero `hint-maxLines` cuenta **renglones dibujados, no frases**: una línea larga ocupa dos y
agota el cupo, cortando el mensaje a media palabra. De ahí que el resumen se genere en **dos
versiones**:

| | |
|---|---|
| `$lineas` | Texto completo, frases enteras → **consola** |
| `$cortas` | Condensado, ~30 caracteres por línea → **notificación** |

Si se toca el texto del aviso, hay que mantener `$cortas` por debajo de unos 36 caracteres por
línea o volverá a truncarse.

El toast lleva `Tag` y `Group` fijos (`horario`), así que cada aviso **sustituye al anterior**
en lugar de irse acumulando en el centro de notificaciones.

### Cuánto dura en pantalla

Windows **no admite un número de segundos a medida**: solo `short` (~7 s) y `long` (~25 s).
Está puesto en `duration="long"`.

Para que se quede fijo **hasta que lo cierres**, usa `-Persistente`, que cambia a
`scenario="reminder"` y añade un botón «Cerrar» (obligatorio en ese modo). Para aplicarlo al
aviso de la mañana hay que añadir ese parámetro a la segunda acción de la tarea
**«Resumen horario»**.

Al margen del script, Windows tiene su propio ajuste en *Configuración → Accesibilidad →
Efectos visuales → Descartar notificaciones después de*, con valores de 5 s a 15 min. Ese
afecta a todas las aplicaciones, pero permite un control más fino.

⚠️ En PowerShell `[int]` **redondea**, no trunca: `[int]8.98` da `9`. Para convertir fracciones
de día a horas hay que usar `[Math]::Floor()` o `.Hours`, o todo sale desplazado una hora.

## Uso manual

```powershell
$H = "$env:USERPROFILE\Scripts\Horario\Sync-Horario.ps1"

& $H -ModoPrueba                      # muestra qué haría, sin escribir
& $H                                  # semana actual + anterior
& $H -SemanasAtras 4                  # revisa un mes hacia atrás
& $H -CrearHojaSiFalta                # crea la hoja de la semana si no existe
& $H -DesdePortapapeles               # toma los datos copiados del navegador
& $H -Forzar                          # corrige también semanas pasadas, no solo huecos
```

`-SemanasAtras` existe porque la propia intranet advierte que los fichajes pueden tardar
hasta **72 h** en consolidarse: repasar la semana anterior recoge los que llegaron tarde.

## Jornada de verano y de invierno

| Periodo | Jornada |
|---|---|
| **15 de junio – 15 de septiembre** (ambos incluidos) | **7:00** |
| Resto del año | **7:43** |

Las fechas y las duraciones están en `configuracion.json`, así que se cambian sin tocar código:

```json
"JornadaVerano": "7:00",  "JornadaInvierno": "7:43",
"VeranoDesde": "15-06",   "VeranoHasta": "15-09"
```

Importa porque `C4` (horas/día) alimenta el objetivo diario. Al crear una hoja copiando la
anterior se **heredaría la jornada vieja**, y en la semana del cambio eso descuadraría los
objetivos de toda la hoja. Por eso:

- **Al crear una hoja**, `C4` se pone según sus fechas, no según lo heredado.
- **En hojas que ya existen**, si `C4` no corresponde a sus fechas **se avisa pero no se toca**:
  ahí mandan tus valores.
- **La semana del cambio** (el 15-09-2026 cae en martes, así que la del 14 tiene días de los dos
  regímenes) se pone la jornada de la mayoría de sus días y **se avisa para que la revises**.

La jornada de una semana es siempre la de la **mayoría de sus días laborables**, nunca la del
lunes. Mirando sólo el lunes, la semana del 14-09-2026 salía como de verano (7:00) y el aviso
daba por incorrecta una hoja que estaba bien puesta a 7:43.

## La hoja de cada semana se crea sola

Las tareas llevan `-CrearHojaSiFalta`, así que **cada lunes se crea sola la hoja de la semana**.
Sin eso, el script avisaría y no escribiría nada hasta que la crearas a mano.

La hoja nueva se copia de la de la semana anterior —hereda formato, fórmulas y los parámetros
de cabecera— y después:

- Se le pone el nombre `dd-MM-aa a dd-MM-aa`.
- Se **vacían los fichajes** de los cinco días. Nace en blanco, sin previsiones: las pones tú
  si quieres.
- Se reetiquetan los días conservando **tu forma de escribirlos** (`Miercoles` sin tilde,
  jueves y viernes en minúscula) y el número a dos dígitos.
- Se actualiza `I1` con el mes. Si la semana cruza de mes, pone el del lunes.
- `C5` (renta semanal) **se corrige sola** con la renta real del viernes anterior.

Lo único que puede requerir un vistazo es `C3` (horas semanales), que se hereda tal cual.

## Vacaciones, festivos y días sin fichar

La web pone el **motivo** de la ausencia en el *tooltip* de la celda `Diario:` — «Vacaciones»,
«Local» (festivo local), «Permiso»… — en el atributo `title` del `<td>`. **Ese motivo es
la señal fiable**, y muchas veces viene **sin horas**:

```
31-ago   title="Vacaciones"      7:00      <- con horas
01-sep   title="Permiso"   (vacío)   <- permiso, sin horas
02-sep   title="Local"           (vacío)   <- festivo local, sin horas
```

| En la web | Interpretación | Qué hace |
|---|---|---|
| **Con motivo** en el `title` | Ausencia justificada | Computa. Usa las horas de `Diario:` y, si no las trae, **la jornada oficial de esa fecha** |
| Sin motivo pero con horas | Ausencia sin etiquetar | Computa esas horas |
| Sin motivo y sin horas | Olvido de fichar | **No toca nada**, solo avisa |

Un día **con motivo se computa aunque sea futuro**: es una ausencia ya confirmada por la
empresa, no una suposición. Sin motivo, solo se miran días ya pasados.

No es cosmético. Sin fichajes, las «Horas trabajadas» del día quedan en `?`, y como el objetivo
de cada día se calcula a partir de la renta del anterior, ese `?` **se propaga en cadena y
descuadra el resto de la semana**. Computando la jornada, el día sale neutro y el saldo sigue
igual que antes.

### Se usan las horas de la web, no las del libro

Un día de vacaciones **sí trae valor en `Diario:`** aunque no tenga marcas, y ese valor es el
que hay que usar: en junio de 2026 fue **7:43**, no las 7:00 de la jornada de verano.

El libro maneja las dos jornadas —`C4` vale 7:43 hasta el 12-06-26 y 7:00 a partir del 15— así
que tomar una constante crearía un desfase de 43 min por día. Se toma siempre lo que diga la
web, que es lo que cuenta la empresa.

### Vacaciones frente a olvido de fichar

La distinción sale de la propia tabla:

| En la web | Interpretación | Qué hace |
|---|---|---|
| Sin marcas, **con** horas en `Diario:` | Ausencia justificada | Computa esas horas |
| Sin marcas y **sin** horas en `Diario:` | Olvido de fichar | **No toca nada**, solo avisa |

Además, la celda de `Diario:` lleva un **tooltip** con el motivo («Vacaciones», «Festivo»…).
Se recoge del atributo `title` del `<td>` y se usa en el aviso y en el comentario de la celda.
Los `title` decorativos de sábado y domingo se ignoran.

⚠️ El tooltip **se pierde al pegar como texto plano**: por el portapapeles en modo texto el día
se computa igual, pero sin el motivo.

### Salvaguardas

| Salvaguarda | Por qué |
|---|---|
| Exige horas en `Diario:` | Sin eso no se escribe nada: no es una ausencia justificada |
| `-DiasMargenAusencia` (1 por defecto) | Un fichaje reciente puede estar sin consolidar (la web avisa de hasta 72 h). Súbelo si salen falsas ausencias |
| Solo días **anteriores** a hoy | Hoy no se toca nunca |
| Se avisa en el log de **cada** día computado | Para poder revisarlo |
| **Comentario en la celda** con el motivo | Para distinguirlo de un fichaje real |
| Las pausas de ese día se vacían | Sin trabajar no hubo café, y si quedara previsto descontaría tiempo |
| `-SinAusencias` | Desactiva el comportamiento por completo |

Es **autocorrectivo**: si el fichaje aparece más tarde, al estar dentro de la ventana de 4
semanas se sobrescribe con los datos reales.

## Renta semanal (C5)

Lo que antes se copiaba a mano: **`C5` de cada semana = `M46` (renta del viernes) de la
semana anterior**. Alimenta el objetivo del lunes (`C10 = C4 - C5`) y, en cadena, el de toda
la semana. Se actualiza sola en cada sincronización.

Se **abstiene y avisa** en los casos en que el dato no sería fiable:

| Situación | Qué hace |
|---|---|
| No existe la hoja de la semana inmediatamente anterior (vacaciones, huecos) | No toca `C5` |
| El viernes anterior aún no tiene salida fichada | No toca `C5` |
| `M46` no está calculada (devuelve `?`) | No toca `C5` |
| La semana anterior cerró **debiendo** (`L46 = "-"`) | Pone **0:00** y avisa |

Lo de la deuda es el criterio que ya venías siguiendo: la única semana que cerró en negativo
(12-01-26, −0:21) arrancó la siguiente desde 0:00, sin arrastrarla.

⚠️ **Nunca se escribe un valor negativo en `C5`.** El libro usa el sistema de fechas 1900
(`Date1904 = False`) y `C5` tiene formato `h:mm` sin corchetes: un negativo se vería como
`########` y arrastraría el error a los objetivos de toda la semana. Hay una salvaguarda
explícita además del caso `L46 = "-"`.

## Seguridad ante Excel abierto

- Libro abierto **con cambios sin guardar** → no se toca nada, se reintenta en la siguiente pasada.
- Libro abierto y guardado → se escribe sobre esa misma sesión de Excel.
- Libro cerrado → se abre en segundo plano.

## Copias de seguridad

Antes de cada guardado **con cambios** se copia el libro a
`Escritorio\_backups_horario\`, conservando las **30** últimas.

## Registro

- `sync-horario.log` — histórico completo.
- `ultimo-resultado.txt` — resumen de la última ejecución.

La pantalla muestra **menos** que el fichero: las líneas de traza que solo sirven para
diagnosticar (la cabecera `===== Sync-Horario …` y `Conectando a la intranet …`) se escriben
con `Write-Log -SoloLog`, así que quedan registradas pero no salen por consola. Si en algún
momento hay que depurar un fallo, la traza completa sigue en el log.

## Corregir etiquetas internas

Al crear una semana copiando la anterior, las cabeceras se quedan con los valores viejos.
`Corregir-Etiquetas.ps1` las realinea con la fecha real de la hoja:

```powershell
& "$env:USERPROFILE\Scripts\Horario\Corregir-Etiquetas.ps1" -ModoPrueba   # ver qué cambiaría
& "$env:USERPROFILE\Scripts\Horario\Corregir-Etiquetas.ps1"               # aplicar
```

Ajusta `B10/B18/B26/B34/B42` (nombre del día + número) e `I1` (mes). **No toca los fichajes
ni ninguna fórmula.** Conserva el nombre del día tal y como esté escrito, incluida la mezcla
de mayúsculas del libro y el «Miercoles» sin tilde; solo recalcula el número, en dos dígitos.

`G1` («SEMANA n») se deja intacto: su criterio de numeración no está definido — casi todas
las hojas dicen «SEMANA 2» sin relación con la semana real.

Ejecutado el 26/08/2026 sobre 18 hojas: mes equivocado en 16 y, además, día equivocado en las
dos primeras de agosto, que arrastraban las fechas de julio.

## Acceso desde fuera de la oficina

**No es automatizable, y no hace falta que lo sea.**

El portal externo para empleados exige: DNI → **CAPTCHA** (identificar imágenes
o continuar un patrón) → usuario y contraseña. El CAPTCHA está puesto justamente para impedir
el acceso programático, así que esa vía queda descartada por diseño.

El CAPTCHA es **hCaptcha** (`hcaptcha.com/1/api.js`): necesita un navegador con JavaScript y
el token que genera **caduca en un par de minutos**, así que no sirve resolverlo una vez y
reutilizarlo durante el día.

**No hay VPN corporativa**, así que la vía elegida es el modo portapapeles. La automatización
del login queda descartada de forma definitiva.

### Modo portapapeles — asistido, sin credenciales

Entras al portal externo a mano como siempre (DNI → hCaptcha → usuario y contraseña) y, al
ver tus fichajes:

1. Selecciona la tabla de fichajes en el navegador y pulsa **Ctrl+C**.
2. Doble clic en **«Actualizar horario (portapapeles)»** en el escritorio.

El acceso directo lanza `Actualizar-DesdePortapapeles.ps1`, que comprueba que hay algo copiado,
sincroniza y deja el resultado en pantalla hasta que pulses una tecla. Equivale a:

```powershell
& "$env:USERPROFILE\Scripts\Horario\Sync-Horario.ps1" -DesdePortapapeles
```

El script lee el **HTML** del portapapeles, que conserva la estructura de la tabla, y lo pasa
por el mismo parser que la intranet: mismas reglas, mismas protecciones, misma copia de
seguridad. Deduce sola la semana a partir de los días copiados (la tabla no incluye el año)
eligiendo la más cercana a hoy.

Las horas se buscan sobre el **texto** de cada celda, no sobre un `<span>` literal: al copiar
desde el navegador el marcado cambia (spans con `style`, clases…) y un patrón rígido devuelve
cero marcas aunque la tabla se reconozca.

### Si se pega como texto plano

Funciona también, pero es más delicado: al perderse las columnas, los siete días quedan en una
línea y las horas caen sueltas debajo **sin decir a qué día pertenece cada una**.

Se reconstruye aprovechando que las marcas de un mismo día van en orden creciente: **cuando una
hora es menor que la anterior, empieza un día nuevo**. El reparto se **verifica** después
contra la fila `Diario:`, que trae un valor por cada día con actividad. Si el número de grupos
no coincide con el de días con datos, se aborta y se pide copiar la tabla con formato, en vez
de escribir un reparto que podría estar desplazado.

No guarda ninguna credencial y no interviene en el login. Necesita `-STA` para poder leer el
portapapeles: el acceso directo ya lo pasa.

### Y si no haces nada, tampoco pasa nada

Los días que no se sincronicen se recuperan solos al volver a la oficina, porque el script
repasa siempre la semana anterior. Con `-SemanasAtras 4` cubre un mes.

## Sobre el cambio mensual de contraseña

**No hay nada que gestionar.** El script no guarda ni usa contraseña: se autentica con
Negotiate/Kerberos, es decir, con el testigo de tu sesión de Windows ya iniciada. Cambiar la
contraseña cada mes no lo afecta en absoluto y no requiere tocar ningún archivo.

Por eso se eliminó `Guardar-Credenciales.ps1`: guardar una contraseña que caduca cada mes,
para un portal que además pide CAPTCHA, no aportaba nada y solo añadía un secreto en disco.
