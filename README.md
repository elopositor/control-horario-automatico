# Sincronizador de fichajes → Excel

Automatiza el control horario personal: lee los fichajes del portal de RRHH corporativo y los
vuelca en un libro de Excel, sin intervención manual.

Nació de una tarea diaria y tediosa —copiar a mano las horas de entrada y salida del portal
web al Excel donde se lleva el cómputo— y acabó cubriendo todo el ciclo: pausas, vacaciones,
festivos, cambios de jornada y creación automática de las hojas semanales.

> **PowerShell 5.1 · Excel COM · VBA · Windows Task Scheduler · Notificaciones nativas**

---

## El problema

La empresa registra los fichajes en un portal web. El control personal de horas se lleva en un
Excel con una hoja por semana y fórmulas que calculan el saldo acumulado. Cada día había que:

1. Entrar al portal, mirar las marcas de entrada y salida.
2. Teclearlas en la hoja de la semana.
3. Cada lunes, duplicar la hoja anterior y corregir a mano las cinco fechas, el mes y la jornada.
4. Al final de la semana, arrastrar el saldo a la hoja siguiente.

## La solución

| | |
|---|---|
| **Al iniciar sesión** | Sincroniza y lanza una notificación con el saldo y la hora de salida |
| **09:30 · 12:00 · 15:20** | Sincroniza en silencio |
| **Cada lunes** | Crea la hoja de la semana con fechas, mes y jornada correctos |

Y tres accesos directos para uso manual: consultar el estado, crear la semana siguiente y
volcar los fichajes desde el portapapeles cuando se trabaja fuera de la red corporativa.

---

## Lo interesante del problema

Lo que parecía un *scraper* de cuatro líneas resultó tener bastantes aristas. Algunas:

### Los fichajes van en pares

Un número **impar** de marcas significa que la jornada sigue abierta: la última es una vuelta
de pausa, no la salida del día. Tomarla por salida machacaba la previsión del usuario — y pasó
en producción: un café a las 11:13-11:43 escribió «11:43» como hora de salida.

### Distinguir el café de la comida

Por **duración**, no por la hora: una comida temprana (12:30 a 13:50) se clasificaría mal con
un corte horario. Los primeros 15 minutos de pausa son de cortesía y no descuentan; solo cuenta
el exceso.

### Vacaciones sin horas

Un día de ausencia no tiene marcas, pero el portal deja el **motivo en el tooltip** de la celda
de balance (`title` del `<td>`, en el tag y no en el contenido). Ese motivo es la señal fiable:

| En el portal | Interpretación |
|---|---|
| Con motivo (`Vacaciones`, `Festivo`…) | Ausencia justificada → se computa |
| Sin motivo pero con horas | Ausencia sin etiquetar → se computa |
| Sin motivo y sin horas | Olvido de fichar → **no se toca nada** |

Importa porque, sin fichajes, las horas del día quedan en `?` y ese `?` **se propaga en cadena**
por los objetivos de los días siguientes, descuadrando la semana entera.

### Dos jornadas al año

7:00 del 15 de junio al 15 de septiembre; 7:43 el resto. Las hojas nuevas heredaban la jornada
de la semana anterior, así que el cambio de temporada las habría descuadrado en silencio. La
semana que cae a caballo recibe la jornada mayoritaria y un aviso para revisarla.

### Sin contraseñas

El portal interno acepta **autenticación integrada de Windows**, así que no hay credenciales que
guardar ni que rotar. El portal externo, en cambio, exige CAPTCHA: en vez de intentar sortearlo
—que no procede—, se resuelve leyendo la tabla que el usuario copia al portapapeles.

---

## Arquitectura

```
Instalar.ps1              Comprueba requisitos, configura y registra las tareas
Desinstalar.ps1           Quita tareas y accesos directos
Configuracion.ps1         Configuración compartida (rutas, URL, jornadas)

Sync-Horario.ps1          Núcleo: descarga, interpreta y escribe en el libro
Mostrar-Resumen.ps1       Notificación con el saldo y la hora de salida
Nueva-Semana.ps1          Crea la hoja de una semana concreta
Ver-MisHoras.ps1          Consulta bajo demanda
Actualizar-DesdePortapapeles.ps1   Vuelco manual fuera de la red corporativa

ModHorario.bas            Macro VBA: botón «Nueva semana» dentro del propio Excel
```

## Requisitos

- Windows con PowerShell 5.1 (de serie)
- Microsoft Excel (se automatiza por COM)
- Sesión de dominio corporativo, si el portal usa autenticación integrada

No requiere permisos de administrador ni instalar dependencias.

## Instalación

```powershell
powershell -ExecutionPolicy Bypass -File .\Instalar.ps1
```

Pregunta una sola vez por el libro de Excel y la URL del portal, y lo guarda en
`configuracion.json` (que **no** se versiona). Hay una plantilla en
`configuracion.ejemplo.json`.

## Documentación

[`LEEME.md`](LEEME.md) recoge el detalle completo: mapa del libro, reglas de escritura,
tratamiento de ausencias, gotchas de Excel COM y decisiones de diseño.

---

## Notas

Proyecto personal, desarrollado para uso propio. El código está anonimizado: la dirección del
portal de fichajes y las rutas locales viven en la configuración, fuera del repositorio, y no
se publica ningún dato de horarios reales.

## Licencia

MIT — ver [`LICENSE`](LICENSE).
