<#
.SYNOPSIS
    Crea la hoja de una semana con las fechas ya puestas, sin tener que corregir nada a mano.

.DESCRIPTION
    Copia la hoja de la semana anterior --hereda formato, formulas y parametros de cabecera--
    y despues deja lista la nueva:

      * Nombre de la hoja "dd-MM-aa a dd-MM-aa".
      * Fichajes de los cinco dias VACIOS (no arrastra los de la semana anterior).
      * Etiquetas de los dias con su fecha correcta, respetando como los escribe el libro.
      * Mes de la cabecera segun la MAYORIA de dias laborables.
      * Jornada (C4) segun la epoca: 7:00 en verano, 7:43 el resto del anio.
      * "Renta semana" (C5) tomada del viernes de la semana anterior.

    Por defecto crea la SIGUIENTE semana que falte. Con -Fecha se puede pedir otra cualquiera.

    Nota: no hace falta usarlo para la semana en curso -- las tareas programadas ya la crean
    solas cada lunes. Esto sirve para adelantarse.

.PARAMETER Fecha
    Cualquier dia de la semana que se quiere crear. Por defecto, la siguiente pendiente.
#>
[CmdletBinding()]
param(
    [datetime] $Fecha = [datetime]::MinValue
)

$RaizScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$Host.UI.RawUI.WindowTitle = 'Nueva semana en el horario'

. (Join-Path $RaizScript 'Configuracion.ps1')

function Get-LunesDe { param([datetime]$F) $F.Date.AddDays(-((([int]$F.DayOfWeek) + 6) % 7)) }

Write-Host ''
Write-Host '  Crear la hoja de una semana' -ForegroundColor Cyan
Write-Host '  ---------------------------' -ForegroundColor Cyan
Write-Host ''

try {
    $libro = (Get-ConfigHorario).RutaLibro

    # Localizar la ultima semana que ya tiene hoja, para saber cual toca.
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($libro, $false, $true)

    $ultimo = [datetime]::MinValue
    foreach ($s in $wb.Sheets) {
        $m = [regex]::Match($s.Name.Trim(), '^(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?')
        if (-not $m.Success) { continue }
        $anio = $null
        if ($m.Groups[3].Success) {
            $a = [int]$m.Groups[3].Value
            $anio = if ($a -lt 100) { 2000 + $a } else { $a }
        } else {
            foreach ($y in 2025, 2026, 2027) {
                try { if ((Get-Date -Year $y -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[1].Value)).DayOfWeek -eq [DayOfWeek]::Monday) { $anio = $y; break } } catch { }
            }
        }
        if (-not $anio) { continue }
        try { $f = (Get-Date -Year $anio -Month ([int]$m.Groups[2].Value) -Day ([int]$m.Groups[1].Value)).Date } catch { continue }
        if ($f -gt $ultimo) { $ultimo = $f }
    }

    $wb.Close($false); $xl.Quit()
    foreach ($o in $wb, $xl) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } } }
    $wb = $null; $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()

    if ($Fecha -eq [datetime]::MinValue) {
        if ($ultimo -eq [datetime]::MinValue) { throw 'No se ha encontrado ninguna hoja semanal en el libro.' }
        $destino = $ultimo.AddDays(7)
        Write-Host "  Ultima semana con hoja : $($ultimo.ToString('dd-MM-yyyy'))" -ForegroundColor Gray
    } else {
        $destino = Get-LunesDe $Fecha
    }

    $ci = [Globalization.CultureInfo]::GetCultureInfo('es-ES')
    Write-Host "  Se va a crear la semana: $($destino.ToString('dd-MM-yyyy')) a $($destino.AddDays(4).ToString('dd-MM-yyyy'))" -ForegroundColor Gray
    Write-Host ''

    & (Join-Path $RaizScript 'Sync-Horario.ps1') -CrearSemana $destino -SemanasAtras 1 -Forzar
}
catch {
    Write-Host ''
    Write-Host "  Ha fallado: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Write-Host '  Pulsa una tecla para cerrar...' -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
