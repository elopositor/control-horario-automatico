<#
.SYNOPSIS
    Envoltorio interactivo para actualizar el horario con la tabla copiada del navegador.

.DESCRIPTION
    Pensado para lanzarse desde el acceso directo del escritorio, cuando estas fuera de la
    oficina y has entrado al portal externo a mano. Comprueba que hay algo copiado, lanza la
    sincronizacion y deja el resultado en pantalla hasta que pulses una tecla.
#>
[CmdletBinding()]
param()

$RaizScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$Sync       = Join-Path $RaizScript 'Sync-Horario.ps1'

$Host.UI.RawUI.WindowTitle = 'Actualizar horario desde el portapapeles'

Write-Host ''
Write-Host '  Actualizar horario desde el portapapeles' -ForegroundColor Cyan
Write-Host '  ----------------------------------------' -ForegroundColor Cyan
Write-Host ''
Write-Host '  Antes de continuar, en el navegador:' -ForegroundColor Gray
Write-Host '    1. Entra en Mis Fichajes del portal de empleados.' -ForegroundColor Gray
Write-Host '    2. Selecciona la tabla de fichajes y pulsa Ctrl+C.' -ForegroundColor Gray
Write-Host ''

try {
    Add-Type -AssemblyName System.Windows.Forms
    $hayHtml  = $false
    $hayTexto = $false
    try { $hayHtml  = -not [string]::IsNullOrWhiteSpace([Windows.Forms.Clipboard]::GetText([Windows.Forms.TextDataFormat]::Html)) } catch { }
    try { $hayTexto = -not [string]::IsNullOrWhiteSpace([Windows.Forms.Clipboard]::GetText()) } catch { }

    if (-not $hayHtml -and -not $hayTexto) {
        Write-Host '  No hay nada copiado.' -ForegroundColor Yellow
        Write-Host '  Copia la tabla de fichajes (Ctrl+C) y vuelve a abrir este acceso directo.' -ForegroundColor Yellow
    }
    else {
        if (-not $hayHtml) {
            Write-Host '  Aviso: lo copiado no conserva el formato de tabla.' -ForegroundColor Yellow
            Write-Host '  Se intentara leer como texto; revisa el resultado.' -ForegroundColor Yellow
            Write-Host ''
        }
        & $Sync -DesdePortapapeles -CrearHojaSiFalta
    }
}
catch {
    Write-Host ''
    Write-Host "  Ha fallado: $($_.Exception.Message)" -ForegroundColor Red
}

Write-Host ''
Write-Host '  Pulsa una tecla para cerrar...' -ForegroundColor DarkGray
$null = $Host.UI.RawUI.ReadKey('NoEcho,IncludeKeyDown')
