<#
.SYNOPSIS
    Instala la sincronizacion del horario en este equipo.

.DESCRIPTION
    Pensado para llevar la carpeta entera a otro ordenador y dejarlo funcionando:

        1. Localiza el libro de horario, o lo pregunta si no lo encuentra.
        2. Comprueba que el libro tiene la estructura esperada.
        3. Comprueba que se llega a la intranet de fichajes.
        4. Guarda la ruta en 'configuracion.json'.
        5. Registra las dos tareas programadas.
        6. Crea los accesos directos en el escritorio.

    La ruta se pregunta UNA sola vez: a partir de ahi queda en configuracion.json. Para
    cambiarla, volver a ejecutar este script o pasarle -RutaLibro.

    No necesita permisos de administrador: las tareas se registran para el usuario actual.

.PARAMETER RutaLibro
    Ruta del libro. Si se omite, se busca y, si no aparece, se pregunta.

.PARAMETER SinTareas
    Solo configura la ruta y los accesos directos, sin registrar las tareas programadas.
#>
[CmdletBinding()]
param(
    [string] $RutaLibro,
    [switch] $SinTareas
)

$ErrorActionPreference = 'Stop'
$dir = $PSScriptRoot

function Escribir { param([string]$T, [string]$C = 'Gray') Write-Host $T -ForegroundColor $C }

Escribir ''
Escribir '  Instalacion de la sincronizacion de horario' 'Cyan'
Escribir '  ==========================================' 'Cyan'
Escribir ''

# ---------------------------------------------------------------- 1. requisitos

$fallos = @()
if ($PSVersionTable.PSVersion.Major -lt 5) { $fallos += "Hace falta PowerShell 5 o superior (hay $($PSVersionTable.PSVersion))." }
try { $null = New-Object -ComObject Excel.Application; }
catch { $fallos += 'No se encuentra Microsoft Excel instalado.' }
Get-Process EXCEL -ErrorAction SilentlyContinue | Where-Object { $_.MainWindowHandle -eq 0 } | ForEach-Object { $_.Kill() }

foreach ($f in 'Sync-Horario.ps1','Mostrar-Resumen.ps1','Configuracion.ps1','Ver-MisHoras.ps1','Actualizar-DesdePortapapeles.ps1') {
    if (-not (Test-Path (Join-Path $dir $f))) { $fallos += "Falta el fichero $f en la carpeta." }
}
if ($fallos) {
    $fallos | ForEach-Object { Escribir "  [X] $_" 'Red' }
    Escribir ''
    return
}
Escribir '  [OK] PowerShell, Excel y ficheros del paquete.' 'Green'

# ------------------------------------------------------------------- 2. libro

. (Join-Path $dir 'Configuracion.ps1')

if (-not $RutaLibro) { $RutaLibro = Find-LibroHorario }

if ($RutaLibro) {
    Escribir "  [OK] Libro localizado: $RutaLibro" 'Green'
} else {
    Escribir '  [ ] No se encuentra el libro de horario.' 'Yellow'
    # Dialogo grafico si se puede; si no, por teclado.
    try {
        Add-Type -AssemblyName System.Windows.Forms
        $dlg = New-Object Windows.Forms.OpenFileDialog
        $dlg.Title  = 'Selecciona el libro de horario'
        $dlg.Filter = 'Libros de Excel (*.xlsx;*.xlsm)|*.xlsx;*.xlsm'
        $dlg.InitialDirectory = [Environment]::GetFolderPath('Desktop')
        if ($dlg.ShowDialog() -eq [Windows.Forms.DialogResult]::OK) { $RutaLibro = $dlg.FileName }
    } catch { }

    if (-not $RutaLibro) { $RutaLibro = (Read-Host '  Escribe la ruta completa del libro').Trim('"') }
}

if (-not $RutaLibro -or -not (Test-Path -LiteralPath $RutaLibro)) {
    Escribir '  [X] Sin un libro valido no se puede continuar.' 'Red'; Escribir ''
    return
}
$RutaLibro = (Resolve-Path -LiteralPath $RutaLibro).Path

# ------------------------------------------------- 3. estructura del libro

Escribir '  [ ] Comprobando la estructura del libro...' 'Gray'
$xl = $null; $wb = $null; $hojasOk = 0
try {
    $xl = New-Object -ComObject Excel.Application
    $xl.Visible = $false; $xl.DisplayAlerts = $false
    $wb = $xl.Workbooks.Open($RutaLibro, $false, $true)
    foreach ($s in $wb.Sheets) {
        if ($s.Name.Trim() -match '^\d{1,2}-\d{1,2}') {
            # Una hoja de semana valida tiene "Entrada" en B11 y "Salida" en B16.
            if ("$($s.Range('B11').Text)" -match 'Entrada' -and "$($s.Range('B16').Text)" -match 'Salida') { $hojasOk++ }
        }
    }
} catch {
    Escribir "  [X] No se pudo abrir el libro: $($_.Exception.Message)" 'Red'; return
} finally {
    try { if ($wb) { $wb.Close($false) }; if ($xl) { $xl.Quit() } } catch { }
    foreach ($o in $wb, $xl) { if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } } }
    $wb = $null; $xl = $null
    [GC]::Collect(); [GC]::WaitForPendingFinalizers()
}

if ($hojasOk -eq 0) {
    Escribir '  [X] Ese libro no tiene hojas semanales con el formato esperado' 'Red'
    Escribir '      (una hoja por semana, "Entrada" en B11 y "Salida" en B16).' 'Red'
    return
}
Escribir "  [OK] Estructura correcta: $hojasOk hoja(s) semanal(es)." 'Green'

# ---------------------------------------------------------------- 4. intranet

# La direccion del portal de fichajes es interna de cada organizacion: no esta en el codigo,
# se guarda en configuracion.json y se pregunta la primera vez.
$url = (Get-ConfigHorarioSegura).UrlIntranet
if (-not $url) {
    Escribir ''
    Escribir '  Falta la direccion del portal de fichajes.' 'Yellow'
    Escribir "  Ejemplo: $script:UrlIntranetEjemplo" 'DarkGray'
    $url = (Read-Host '  Pega aqui la URL de tu portal de fichajes').Trim()
}

try {
    $ProgressPreference = 'SilentlyContinue'
    $r = Invoke-WebRequest -Uri $url -UseDefaultCredentials -UseBasicParsing -TimeoutSec 20
    if ($r.Content -match 'tabla_fichajes') { Escribir '  [OK] Intranet de fichajes accesible.' 'Green' }
    else { Escribir '  [!] La intranet responde pero sin la tabla de fichajes.' 'Yellow' }
} catch {
    Escribir '  [!] No se llega ahora a la intranet (sin red corporativa o VPN).' 'Yellow'
    Escribir '      La instalacion sigue; funcionara cuando estes en la oficina.' 'Yellow'
}

# ------------------------------------------------------------ 5. configuracion

Set-ConfigHorario -RutaLibro $RutaLibro -UrlIntranet $url
Escribir "  [OK] Ruta guardada en configuracion.json" 'Green'

# ------------------------------------------------------------------ 6. tareas

$pshell = "$env:SystemRoot\System32\WindowsPowerShell\v1.0\powershell.exe"
$user   = "$env:USERDOMAIN\$env:USERNAME"
$argSync = "-NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$dir\Sync-Horario.ps1`" -SemanasAtras 3 -Forzar -CrearHojaSiFalta"

if (-not $SinTareas) {
    $ajustes = New-ScheduledTaskSettingsSet -StartWhenAvailable -RunOnlyIfNetworkAvailable -DontStopOnIdleEnd `
        -ExecutionTimeLimit (New-TimeSpan -Minutes 15) -MultipleInstances IgnoreNew -Hidden
    $principal = New-ScheduledTaskPrincipal -UserId $user -LogonType Interactive -RunLevel Limited

    $a1 = New-ScheduledTaskAction -Execute $pshell -Argument $argSync
    $t1 = @(
        (New-ScheduledTaskTrigger -Daily -At '09:30'),
        (New-ScheduledTaskTrigger -Daily -At '12:00'),
        (New-ScheduledTaskTrigger -Daily -At '15:20')
    )
    try { Unregister-ScheduledTask -TaskName 'Sincronizar horario' -Confirm:$false -ErrorAction Stop } catch { }
    Register-ScheduledTask -TaskName 'Sincronizar horario' -Action $a1 -Trigger $t1 -Settings $ajustes `
        -Principal $principal -Description 'Vuelca los fichajes de la intranet corporativa en el libro de horario.' | Out-Null

    $a2 = @(
        (New-ScheduledTaskAction -Execute $pshell -Argument $argSync),
        (New-ScheduledTaskAction -Execute $pshell -Argument "-STA -NoProfile -NonInteractive -WindowStyle Hidden -ExecutionPolicy Bypass -File `"$dir\Mostrar-Resumen.ps1`"")
    )
    $t2 = New-ScheduledTaskTrigger -AtLogOn -User $user
    $t2.Delay = 'PT3M'
    try { Unregister-ScheduledTask -TaskName 'Resumen horario' -Confirm:$false -ErrorAction Stop } catch { }
    Register-ScheduledTask -TaskName 'Resumen horario' -Action $a2 -Trigger $t2 -Settings $ajustes `
        -Principal $principal -Description 'Al iniciar sesion: sincroniza y avisa del saldo de horas.' | Out-Null

    Escribir '  [OK] Tareas programadas: 09:30, 12:00, 15:20 y al iniciar sesion.' 'Green'
} else {
    Escribir '  [ ] Tareas programadas omitidas (-SinTareas).' 'Yellow'
}

# -------------------------------------------------------- 7. accesos directos

$esc = [Environment]::GetFolderPath('Desktop')
$ws  = New-Object -ComObject WScript.Shell

$l1 = $ws.CreateShortcut((Join-Path $esc 'Mis horas.lnk'))
$l1.TargetPath = $pshell
$l1.Arguments  = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$dir\Ver-MisHoras.ps1`""
$l1.WorkingDirectory = $dir
$l1.Description = 'Actualiza los fichajes y muestra el saldo de horas'
$l1.IconLocation = "$env:SystemRoot\System32\shell32.dll,13"
$l1.Save()

$l3 = $ws.CreateShortcut((Join-Path $esc 'Nueva semana.lnk'))
$l3.TargetPath = $pshell
$l3.Arguments  = "-NoProfile -ExecutionPolicy Bypass -File `"$dir\Nueva-Semana.ps1`""
$l3.WorkingDirectory = $dir
$l3.Description = 'Crea la hoja de la siguiente semana con las fechas ya puestas'
$l3.IconLocation = "$env:SystemRoot\System32\shell32.dll,171"
$l3.Save()

$l2 = $ws.CreateShortcut((Join-Path $esc 'Actualizar horario (portapapeles).lnk'))
$l2.TargetPath = $pshell
$l2.Arguments  = "-STA -NoProfile -ExecutionPolicy Bypass -File `"$dir\Actualizar-DesdePortapapeles.ps1`""
$l2.WorkingDirectory = $dir
$l2.Description = 'Actualiza el horario con la tabla copiada del portal externo'
$xlExe = @("${env:ProgramFiles}\Microsoft Office\root\Office16\EXCEL.EXE",
           "${env:ProgramFiles(x86)}\Microsoft Office\root\Office16\EXCEL.EXE") |
         Where-Object { Test-Path $_ } | Select-Object -First 1
if ($xlExe) { $l2.IconLocation = "$xlExe,0" }
$l2.Save()
[void][Runtime.InteropServices.Marshal]::ReleaseComObject($ws)

Escribir '  [OK] Accesos directos creados en el escritorio.' 'Green'

Escribir ''
Escribir '  Listo.' 'Cyan'
Escribir "  Libro    : $RutaLibro"
Escribir "  Scripts  : $dir"
Escribir "  Copias   : $(Join-Path (Split-Path -Parent $RutaLibro) '_backups_horario')"
Escribir ''
Escribir '  Comprueba que todo va con:  .\Ver-MisHoras.ps1' 'Gray'
Escribir ''
