<#
.SYNOPSIS
    Muestra un aviso con el estado de horas del dia y la hora minima de salida.

.DESCRIPTION
    Lee la hoja de la semana en curso de "Copia de horario 2026.xlsx" y calcula:

      * Saldo acumulado con el que empieza el dia.
      * Hora minima de salida para no deber horas (consume el saldo a favor).
      * Hora de jornada completa (mantiene el saldo intacto).

    Se apoya en el propio modelo del libro. La clave es que el objetivo diario C(base) ya
    lleva descontado el saldo acumulado: si arrastras +1:25 a favor, el objetivo de hoy no
    es 7:00 sino 5:35. De ahi:

        saldo previo  = C4 - C(base)                      (negativo => debes horas)
        salida minima = D(base+1) + G(base+2) + G(base+4) + C(base)
        jornada plena = D(base+1) + G(base+2) + G(base+4) + C4

    No se usa la formula M(base) ("Hora optima de salida") del libro porque devuelve "?"
    siempre que no se sale a comer, que es el caso habitual.

    Este script SOLO LEE el libro. Nunca escribe.

.PARAMETER Consola
    Escribe el resumen por pantalla en vez de lanzar la notificacion.
#>
[CmdletBinding()]
param(
    [switch]   $Consola,
    # El aviso se queda en pantalla hasta que lo cierras, en vez de ocultarse solo.
    [switch]   $Persistente,
    [datetime] $Fecha = (Get-Date).Date,   # solo para comprobar otros dias
    # Si no se indica, sale de configuracion.json (lo deja escrito Instalar.ps1).
    [string]   $Libro
)

$ErrorActionPreference = 'Stop'
$ci = [Globalization.CultureInfo]::GetCultureInfo('es-ES')

. (Join-Path $PSScriptRoot 'Configuracion.ps1')
if (-not $Libro) { $Libro = (Get-ConfigHorario).RutaLibro }

function Get-LunesDe {
    param([datetime]$Fecha)
    $offset = (([int]$Fecha.DayOfWeek) + 6) % 7
    return $Fecha.Date.AddDays(-$offset)
}

function Get-LunesDeNombre {
    param([string]$Nombre)
    $m = [regex]::Match($Nombre.Trim(), '^(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?')
    if (-not $m.Success) { return $null }
    $d = [int]$m.Groups[1].Value; $mes = [int]$m.Groups[2].Value
    if ($mes -lt 1 -or $mes -gt 12 -or $d -lt 1 -or $d -gt 31) { return $null }
    if ($m.Groups[3].Success) {
        $a = [int]$m.Groups[3].Value
        $anio = if ($a -lt 100) { 2000 + $a } else { $a }
        try { return (Get-Date -Year $anio -Month $mes -Day $d).Date } catch { return $null }
    }
    foreach ($anio in 2025, 2026, 2027) {
        try { $f = (Get-Date -Year $anio -Month $mes -Day $d).Date } catch { continue }
        if ($f.DayOfWeek -eq [DayOfWeek]::Monday) { return $f }
    }
    return $null
}

# Ojo: [int] en PowerShell REDONDEA (8.98 -> 9), no trunca. Usar siempre Floor/.Hours
# o las horas salen desplazadas.

function Format-Horas {
    <# 0.2326 -> "5:35" ; admite negativos #>
    param([double]$Fraccion)
    $signo = if ($Fraccion -lt 0) { '-' } else { '' }
    $t = [timespan]::FromDays([Math]::Abs($Fraccion))
    return '{0}{1}:{2:00}' -f $signo, [Math]::Floor($t.TotalHours), $t.Minutes
}

function Format-Hora {
    <# 0.6069 -> "14:34" #>
    param([double]$Fraccion)
    $t = [timespan]::FromDays($Fraccion)
    return '{0:00}:{1:00}' -f $t.Hours, $t.Minutes
}

function Show-Aviso {
    param([string]$Titulo, [string[]]$Lineas, [string[]]$Cortas)

    if ($Consola) {
        Write-Host ''
        Write-Host "  $Titulo" -ForegroundColor Cyan
        $Lineas | ForEach-Object { Write-Host "  $_" }
        Write-Host ''
        return
    }

    # En la notificacion manda la version corta; la larga no cabe y se corta a media frase.
    $Lineas = if ($Cortas -and $Cortas.Count -gt 0) { $Cortas } else { $Lineas }

    # Notificacion nativa de Windows. Se construye el XML a mano en vez de usar una plantilla:
    # ToastText04 solo trae 3 nodos de texto (titulo + 2 lineas) y aqui hacen falta mas.
    try {
        [void][Windows.UI.Notifications.ToastNotificationManager, Windows.UI.Notifications, ContentType = WindowsRuntime]
        [void][Windows.Data.Xml.Dom.XmlDocument, Windows.Data.Xml.Dom, ContentType = WindowsRuntime]

        $esc = { param($s) [Security.SecurityElement]::Escape($s) }

        # Windows solo muestra dos elementos <text> y colapsa el resto en "+N notificaciones".
        # Por eso TODO el cuerpo va en UN unico <text> con saltos (&#10;) y hint-maxLines="5",
        # que si se despliega entero.
        $cuerpo = (($Lineas | Select-Object -First 5) | ForEach-Object { & $esc $_ }) -join '&#10;'

        # duration="long" lo mantiene ~25 s en pantalla en vez de los ~7 de por defecto.
        # scenario="reminder" lo deja fijo hasta que se cierra, pero exige un boton de accion.
        if ($Persistente) {
            $atributos = "scenario='reminder'"
            $acciones  = "<actions><action activationType='system' arguments='dismiss' content='Cerrar'/></actions>"
        } else {
            $atributos = "duration='long'"
            $acciones  = ''
        }

        $xml = [Windows.Data.Xml.Dom.XmlDocument]::new()
        $xml.LoadXml("<toast $atributos><visual><binding template='ToastGeneric'><text>$(& $esc $Titulo)</text><text hint-maxLines='5'>$cuerpo</text></binding></visual>$acciones</toast>")

        $toast = [Windows.UI.Notifications.ToastNotification]::new($xml)
        # Tag y Group fijos: cada aviso SUSTITUYE al anterior en vez de ir acumulandose.
        $toast.Tag   = 'horario'
        $toast.Group = 'horario'

        $appId = '{1AC14E77-02E7-4E5D-B744-2EB1AE5198B7}\WindowsPowerShell\v1.0\powershell.exe'
        [Windows.UI.Notifications.ToastNotificationManager]::CreateToastNotifier($appId).Show($toast)
    }
    catch {
        # Respaldo que NUNCA bloquea: se cierra solo a los 45 s. Un MessageBox normal dejaria
        # la tarea programada colgada esperando un clic.
        try {
            $ws = New-Object -ComObject WScript.Shell
            [void]$ws.Popup(($Lineas -join "`r`n"), 45, $Titulo, 64)
            [void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)
        } catch {
            Write-Host "$Titulo`r`n$($Lineas -join "`r`n")"
        }
    }
}

# ----------------------------------------------------------------------- main

$hoy = $Fecha.Date
if ($hoy.DayOfWeek -in [DayOfWeek]::Saturday, [DayOfWeek]::Sunday) { return }
if (-not (Test-Path $Libro)) { return }

function Get-Resumen {
    <#
        Recoge los datos del libro y devuelve @{ Titulo = ...; Lineas = ... }.
        Va en una funcion propia para que Excel quede CERRADO antes de mostrar nada: si el
        aviso se quedase bloqueado, un Excel abierto aqui sobreviviria como proceso huerfano.
        Los "return" salen de la funcion, no del script, asi que el finally siempre corre.
    #>
    param([datetime]$hoy)

    $excel = $null; $wb = $null; $yaAbierto = $false
    try {
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($lb in $excel.Workbooks) { if ($lb.FullName -eq $Libro) { $wb = $lb; $yaAbierto = $true; break } }
        if (-not $yaAbierto) { $excel = $null }
    } catch { $excel = $null }

    if (-not $yaAbierto) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Libro, $false, $true)   # solo lectura
    }

    $lunes = Get-LunesDe $hoy
    $hoja  = $null
    foreach ($h in $wb.Sheets) { if ((Get-LunesDeNombre $h.Name) -eq $lunes) { $hoja = $h; break } }
    if (-not $hoja) {
        return @{ Titulo = "Horario - $($hoy.ToString('dddd d', $ci))"
                  Lineas = @('No existe todavía la hoja de esta semana.') }
    }

    $d    = ([int]$hoy.DayOfWeek + 6) % 7
    $base = 10 + 8 * $d

    $jornada  = [double]$hoja.Range('C4').Value2
    $objetivo = $hoja.Cells.Item($base, 3).Value2
    $entrada  = $hoja.Cells.Item($base + 1, 4).Value2   # D = entrada efectiva
    $descDes  = $hoja.Cells.Item($base + 2, 7).Value2
    $descCom  = $hoja.Cells.Item($base + 4, 7).Value2
    $salida   = $hoja.Cells.Item($base + 6, 4).Value2   # D = salida efectiva
    $salidaC  = $hoja.Cells.Item($base + 6, 3).Value2   # C = lo que hay escrito

    $titulo = "Horario - $((Get-Culture).TextInfo.ToTitleCase($hoy.ToString('dddd d', $ci)))"

    if ($null -eq $entrada -or "$entrada" -eq '') {
        return @{ Titulo = $titulo; Lineas = @('Aún no hay entrada fichada hoy.') }
    }
    if ($objetivo -isnot [double]) {
        return @{ Titulo = $titulo
                  Lineas = @("Entrada $(Format-Hora ([double]$entrada))",
                             'El objetivo del día no está calculado: faltan datos de días anteriores.') }
    }

    $descuentos = 0.0
    if ($descDes -is [double]) { $descuentos += $descDes }
    if ($descCom -is [double]) { $descuentos += $descCom }

    $salidaMin  = [double]$entrada + $descuentos + [double]$objetivo
    $salidaPlena= [double]$entrada + $descuentos + $jornada
    $saldo      = $jornada - [double]$objetivo     # + a favor, - deuda

    # Se preparan DOS versiones del mismo mensaje:
    #   $lineas -> texto completo, para la consola, que no tiene limite de espacio.
    #   $cortas -> version condensada para la notificacion de Windows. Su "hint-maxLines"
    #              cuenta lineas VISUALES, no elementos: una frase larga ocupa dos renglones
    #              y agota el cupo, cortando el mensaje a mitad. Hay que dejarlas en ~36
    #              caracteres para que cada una ocupe un solo renglon.
    $lineas = @()
    $cortas = @()

    if ($saldo -ge 0) {
        $lineas += "Saldo: +$(Format-Horas $saldo) a favor  |  entrada $(Format-Hora ([double]$entrada))"
        $cortas += "+$(Format-Horas $saldo) a favor · entrada $(Format-Hora ([double]$entrada))"
    } else {
        $lineas += "Saldo: debes $(Format-Horas ([Math]::Abs($saldo)))  |  entrada $(Format-Hora ([double]$entrada))"
        $cortas += "Debes $(Format-Horas ([Math]::Abs($saldo))) · entrada $(Format-Hora ([double]$entrada))"
    }

    # Si alguna pausa se ha pasado de la cortesia, se dice cuanto retrasa la salida.
    # El umbral de medio minuto evita que aparezca "0:00" por residuos de coma flotante:
    # una pausa de exactamente 15 min deja en G un valor del orden de 1e-17, no cero.
    $minimoVisible = 0.5 / 1440.0
    if ($descuentos -gt $minimoVisible) {
        $desde = $hoja.Cells.Item($base + 2, 3).Value2
        $hasta = $hoja.Cells.Item($base + 3, 3).Value2
        if ($descDes -is [double] -and $descDes -gt $minimoVisible -and $desde -is [double] -and $hasta -is [double]) {
            $lineas += "Café $(Format-Hora ([double]$desde))-$(Format-Hora ([double]$hasta)): retrasa la salida $(Format-Horas $descuentos)"
            $cortas += "Café $(Format-Hora ([double]$desde))-$(Format-Hora ([double]$hasta)) → +$(Format-Horas $descuentos)"
        } else {
            $lineas += "Pausas: retrasan la salida $(Format-Horas $descuentos)"
            $cortas += "Pausas → +$(Format-Horas $descuentos)"
        }
    }

    # Para un dia que no es hoy, se considera cerrado a todos los efectos.
    $ahora = if ($hoy -eq (Get-Date).Date) { ((Get-Date) - (Get-Date).Date).TotalDays } else { 1.0 }

    if ($salidaC -is [double] -and $salida -is [double] -and $ahora -gt [double]$salida) {
        # El dia ya esta cerrado: se informa de como quedo.
        $trabajadas = [double]$salida - [double]$entrada - $descuentos
        $renta = $trabajadas - [double]$objetivo
        $lineas += "Salida $(Format-Hora ([double]$salida))  |  trabajadas $(Format-Horas $trabajadas)"
        $cortas += "Salida $(Format-Hora ([double]$salida)) · $(Format-Horas $trabajadas) trabajadas"
        if ($renta -ge 0) {
            $lineas += "Día cerrado con +$(Format-Horas $renta) sobre el objetivo."
            $cortas += "Día cerrado: +$(Format-Horas $renta)"
        } else {
            $lineas += "Día cerrado con -$(Format-Horas ([Math]::Abs($renta))) sobre el objetivo."
            $cortas += "Día cerrado: -$(Format-Horas ([Math]::Abs($renta)))"
        }
    }
    else {
        # Jornada abierta. Como quedaria el dia si se sale a la hora que hay escrita como
        # prevision: es la referencia mas util, porque suele ser la que se piensa cumplir.
        $prevista = if ($salida -is [double]) { $salida } else { $salidaC }
        $lineaPrevista = $null
        $cortaPrevista = $null
        if ($prevista -is [double]) {
            $balance = ([double]$prevista - [double]$entrada - $descuentos) - [double]$objetivo
            $signo   = if ($balance -lt 0) { '-' } else { '+' }
            $lineaPrevista = "Si sales a las $(Format-Hora ([double]$prevista)) tu balance de horas queda en $signo$(Format-Horas ([Math]::Abs($balance)))"
            $cortaPrevista = "Si sales $(Format-Hora ([double]$prevista)) → $signo$(Format-Horas ([Math]::Abs($balance)))"
        }

        if ($saldo -lt 0) {
            # Se deben horas: para saldar hay que quedarse MAS que la jornada normal.
            $lineas += "Sal a las $(Format-Hora $salidaMin) y saldas la deuda"
            if ($lineaPrevista) { $lineas += $lineaPrevista }
            $lineas += "Saliendo a las $(Format-Hora $salidaPlena) seguirías debiendo $(Format-Horas ([Math]::Abs($saldo)))"

            $cortas += "$(Format-Hora $salidaMin) saldas la deuda"
            if ($cortaPrevista) { $cortas += $cortaPrevista }
        }
        elseif ($hoy.DayOfWeek -eq [DayOfWeek]::Friday) {
            # El viernes es el dia de gastar el margen acumulado, si se quiere.
            $lineas += "Puedes salir ya a las $(Format-Hora $salidaMin) y gastar el margen"
            if ($lineaPrevista) { $lineas += $lineaPrevista }
            $lineas += "A las $(Format-Hora $salidaPlena) lo acumulas para la semana que viene"

            $cortas += "$(Format-Hora $salidaMin) gastas · $(Format-Hora $salidaPlena) acumulas"
            if ($cortaPrevista) { $cortas += $cortaPrevista }
        }
        else {
            # De lunes a jueves interesa conservar el margen, no consumirlo.
            $lineas += "Sal a las $(Format-Hora $salidaPlena) y mantienes tu margen"
            if ($lineaPrevista) { $lineas += $lineaPrevista }
            $lineas += "No bajes de las $(Format-Hora $salidaMin) o te quedarás en negativo"

            $cortas += "$(Format-Hora $salidaPlena) mantienes · $(Format-Hora $salidaMin) mínimo"
            if ($cortaPrevista) { $cortas += $cortaPrevista }
        }
    }

        return @{ Titulo = $titulo; Lineas = $lineas; Cortas = $cortas }
    }
    catch {
        Write-Host "Fallo al preparar el resumen: $($_.Exception.Message)" -ForegroundColor Red
        return $null
    }
    finally {
        try {
            if ($wb -and -not $yaAbierto)    { $wb.Close($false) }
            if ($excel -and -not $yaAbierto) { $excel.Quit() }
        } catch { }
        # Hay que soltar TODAS las referencias, tambien la de la hoja: cualquier RCW vivo
        # mantiene el proceso EXCEL.EXE en memoria aunque se haya llamado a Quit().
        foreach ($o in $hoja, $wb, $excel) {
            if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
        }
        $hoja = $null; $wb = $null; $excel = $null
    }
}

# Excel ya esta cerrado en este punto: el aviso se muestra sobre datos en memoria.
$resumen = Get-Resumen -hoy $hoy

# La recoleccion va AQUI, no dentro de la funcion: alli sus variables locales siguen vivas
# y el GC no puede soltar los RCW (uno por cada hoja recorrida), que dejarian EXCEL.EXE
# en memoria pese al Quit().
[GC]::Collect(); [GC]::WaitForPendingFinalizers()
[GC]::Collect(); [GC]::WaitForPendingFinalizers()

if ($resumen) { Show-Aviso -Titulo $resumen.Titulo -Lineas $resumen.Lineas -Cortas $resumen.Cortas }
