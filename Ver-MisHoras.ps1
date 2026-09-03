<#
.SYNOPSIS
    Consulta bajo demanda: actualiza los fichajes y muestra como vas de horas.

.DESCRIPTION
    Pensado para el acceso directo del escritorio, para mirarlo a media tarde y decidir a que
    hora salir. Primero sincroniza (asi cuenta tu salida si ya la has fichado) y despues
    muestra el resumen por pantalla, dejandolo visible hasta que pulses una tecla.

.PARAMETER SinSincronizar
    Salta la consulta a la intranet y solo lee el Excel. Mas rapido, pero con los datos de la
    ultima sincronizacion.
#>
[CmdletBinding()]
param(
    [switch] $SinSincronizar
)

$RaizScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$Host.UI.RawUI.WindowTitle = 'Mis horas'

if (-not $SinSincronizar) {
    Write-Host ''
    Write-Host '  Actualizando fichajes...' -ForegroundColor DarkGray
    try {
        # -CrearHojaSiFalta va aqui tambien: sin el, abrir "Mis horas" un lunes solo avisaba de
        # que faltaba la hoja de la semana, en vez de crearla.
        & (Join-Path $RaizScript 'Sync-Horario.ps1') -SemanasAtras 3 -Forzar -CrearHojaSiFalta | Out-Null
    } catch {
        Write-Host "  No se pudo actualizar: $($_.Exception.Message)" -ForegroundColor Yellow
        Write-Host '  Se muestra el estado con los ultimos datos guardados.' -ForegroundColor Yellow
    }
}

try {
    & (Join-Path $RaizScript 'Mostrar-Resumen.ps1') -Consola
} catch {
    Write-Host ''
    Write-Host "  Ha fallado: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host '  Pulsa una tecla para cerrar...' -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
