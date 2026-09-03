<#
.SYNOPSIS
    Quita la sincronizacion del horario de este equipo.

.DESCRIPTION
    Elimina las dos tareas programadas y los accesos directos del escritorio.

    NO toca el libro de horario, ni sus copias de seguridad, ni la carpeta de scripts:
    para deshacerlo del todo, borra la carpeta a mano despues de ejecutar esto.
#>
[CmdletBinding()]
param()

$ErrorActionPreference = 'Continue'

Write-Host ''
Write-Host '  Desinstalando la sincronizacion de horario' -ForegroundColor Cyan
Write-Host ''

foreach ($t in 'Sincronizar horario','Resumen horario') {
    try {
        Unregister-ScheduledTask -TaskName $t -Confirm:$false -ErrorAction Stop
        Write-Host "  [OK] Tarea eliminada: $t" -ForegroundColor Green
    } catch {
        Write-Host "  [ ] No existia la tarea: $t" -ForegroundColor DarkGray
    }
}

$esc = [Environment]::GetFolderPath('Desktop')
foreach ($l in 'Mis horas.lnk','Actualizar horario (portapapeles).lnk') {
    $p = Join-Path $esc $l
    if (Test-Path -LiteralPath $p) {
        Remove-Item -LiteralPath $p -Force
        Write-Host "  [OK] Acceso directo eliminado: $l" -ForegroundColor Green
    } else {
        Write-Host "  [ ] No existia: $l" -ForegroundColor DarkGray
    }
}

Write-Host ''
Write-Host '  El libro de horario y sus copias de seguridad NO se han tocado.' -ForegroundColor Yellow
Write-Host '  Para eliminarlo todo, borra tambien esta carpeta de scripts.' -ForegroundColor Yellow
Write-Host ''
