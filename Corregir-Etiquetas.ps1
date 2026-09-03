<#
.SYNOPSIS
    Corrige las etiquetas internas desactualizadas de las hojas de "Copia de horario 2026.xlsx".

.DESCRIPTION
    Al crear cada semana copiando la anterior, las etiquetas de cabecera se quedan con los
    valores de la hoja origen. Este script las realinea con la fecha real de la hoja:

      * B10 / B18 / B26 / B34 / B42  -> nombre del dia + numero correcto
      * I1                            -> mes correcto

    La fecha de referencia se toma del NOMBRE de la hoja (primer tramo "dd-mm[-aa]"), que es
    la unica fuente fiable del libro. Si el nombre no lleva anio, se deduce probando cual hace
    que ese dia caiga realmente en lunes.

    NO se tocan los fichajes (columna C) ni ninguna formula. G1 ("SEMANA n") se deja como esta
    porque su criterio de numeracion no esta definido.

.PARAMETER ModoPrueba
    Muestra los cambios sin escribir.
#>
[CmdletBinding()]
param(
    [switch] $ModoPrueba,
    # Si no se indican, salen de configuracion.json (lo deja escrito Instalar.ps1).
    [string] $Libro,
    [string] $CarpetaBak
)

$ErrorActionPreference = 'Stop'
$ci = [Globalization.CultureInfo]::GetCultureInfo('es-ES')

. (Join-Path $PSScriptRoot 'Configuracion.ps1')
if (-not $Libro)      { $Libro      = (Get-ConfigHorario).RutaLibro }
if (-not $CarpetaBak) { $CarpetaBak = Join-Path (Split-Path -Parent $Libro) '_backups_horario' }

function Get-LunesDeNombre {
    <# Devuelve el lunes de la hoja, o $null si el nombre no es interpretable. #>
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
    # Sin anio en el nombre: elegir el que haga que la fecha caiga en lunes.
    foreach ($anio in 2025, 2026, 2027) {
        try { $f = (Get-Date -Year $anio -Month $mes -Day $d).Date } catch { continue }
        if ($f.DayOfWeek -eq [DayOfWeek]::Monday) { return $f }
    }
    return $null
}

if (-not (Test-Path $Libro)) { throw "No existe el libro: $Libro" }

$excel = $null; $wb = $null; $yaAbierto = $false
$cambios = New-Object System.Collections.Generic.List[string]

try {
    # Reutilizar la sesion de Excel si el libro ya esta abierto.
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($lb in $excel.Workbooks) { if ($lb.FullName -eq $Libro) { $wb = $lb; $yaAbierto = $true; break } }
        if (-not $yaAbierto) { $excel = $null }
    } catch { $excel = $null }

    if ($yaAbierto -and -not $wb.Saved) {
        Write-Host 'El libro esta abierto con cambios sin guardar. Guardalos y vuelve a lanzarlo.' -ForegroundColor Yellow
        return
    }
    if (-not $yaAbierto) {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false; $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Libro)
    }

    foreach ($hoja in $wb.Sheets) {
        $lunes = Get-LunesDeNombre $hoja.Name
        if ($null -eq $lunes) { Write-Host "  omitida '$($hoja.Name)' (nombre no interpretable)" -ForegroundColor DarkGray; continue }
        if ($lunes.DayOfWeek -ne [DayOfWeek]::Monday) {
            Write-Host "  omitida '$($hoja.Name)': $($lunes.ToString('dd-MM-yyyy')) no es lunes." -ForegroundColor Yellow; continue
        }

        $deHoja = @()

        for ($d = 0; $d -lt 5; $d++) {
            $fila   = 10 + 8 * $d
            $celda  = $hoja.Cells.Item($fila, 2)
            if ("$($celda.Formula)" -like '=*') { continue }   # nunca sobrescribir formulas

            $actual = "$($celda.Text)"
            $fecha  = $lunes.AddDays($d)

            # Conservar el nombre del dia tal y como esta escrito en el libro (respeta su
            # mezcla de mayusculas y "Miercoles" sin tilde); solo se recalcula el numero.
            $nombreDia = ($actual -replace '\d', '').Trim()
            if ([string]::IsNullOrWhiteSpace($nombreDia)) {
                $nombreDia = $fecha.ToString('dddd', $ci)
                $nombreDia = $nombreDia.Substring(0,1).ToUpper() + $nombreDia.Substring(1)
            }
            $nuevo = '{0} {1:00}' -f $nombreDia, $fecha.Day

            if ($actual -ne $nuevo) {
                $deHoja += "    B$fila : '$actual' -> '$nuevo'"
                if (-not $ModoPrueba) { $celda.Value2 = $nuevo }
            }
        }

        # I1 = mes. En semanas a caballo se acepta el del lunes o el del viernes.
        $celdaMes = $hoja.Range('I1')
        if ("$($celdaMes.Formula)" -notlike '=*') {
            $actualMes = "$($celdaMes.Text)".Trim()
            $mesLunes  = $lunes.ToString('MMMM', $ci)
            $mesVier   = $lunes.AddDays(4).ToString('MMMM', $ci)
            if ($actualMes.ToLower() -notin @($mesLunes, $mesVier)) {
                $deHoja += "    I1  : '$actualMes' -> '$mesLunes'"
                if (-not $ModoPrueba) { $celdaMes.Value2 = $mesLunes }
            }
        }

        if ($deHoja.Count) {
            $cambios.Add("  '$($hoja.Name)'  (lunes $($lunes.ToString('dd-MM-yyyy')))")
            $deHoja | ForEach-Object { $cambios.Add($_) }
        }
    }

    if ($cambios.Count -eq 0) {
        Write-Host 'Todas las etiquetas estaban correctas.' -ForegroundColor Green
        return
    }

    $cambios | ForEach-Object { Write-Host $_ }
    $nHojas = ($cambios | Where-Object { $_ -match "^  '" }).Count
    Write-Host ''

    if ($ModoPrueba) {
        Write-Host "MODO PRUEBA: $nHojas hoja(s) por corregir. No se ha escrito nada." -ForegroundColor Cyan
    } else {
        if (-not (Test-Path $CarpetaBak)) { New-Item -ItemType Directory -Path $CarpetaBak -Force | Out-Null }
        $destino = Join-Path $CarpetaBak ("{0}_etiquetas_{1}.xlsx" -f `
            [IO.Path]::GetFileNameWithoutExtension($Libro), (Get-Date -Format 'yyyyMMdd-HHmm'))
        Copy-Item -Path $Libro -Destination $destino -Force   # el disco aun tiene el estado previo
        $wb.Save()
        Write-Host "$nHojas hoja(s) corregidas. Copia previa: $(Split-Path -Leaf $destino)" -ForegroundColor Green
    }
}
finally {
    try {
        if ($wb -and -not $yaAbierto)    { $wb.Close($false) }
        if ($excel -and -not $yaAbierto) { $excel.Quit() }
    } catch { }
    foreach ($o in $wb, $excel) { if ($o) { try { [Runtime.InteropServices.Marshal]::ReleaseComObject($o) | Out-Null } catch { } } }
}
