<#
.SYNOPSIS
    Configuracion compartida por todos los scripts del horario.

.DESCRIPTION
    Se carga con dot-source desde los demas scripts:

        . (Join-Path $PSScriptRoot 'Configuracion.ps1')

    Guarda la ruta del libro en 'configuracion.json', junto a los scripts, para que la
    instalacion sea portable: al llevar la carpeta a otro equipo solo hay que volver a
    ejecutar Instalar.ps1.

    NO se pide la ruta en cada ejecucion: las tareas programadas corren desatendidas y no
    pueden responder a una pregunta. Solo se pregunta al instalar, o si el libro configurado
    ha desaparecido Y hay alguien delante.
#>

$script:FicheroConfig = Join-Path $PSScriptRoot 'configuracion.json'

function Find-LibroHorario {
    <# Busca el libro por los sitios habituales. Devuelve la ruta o $null. #>
    $candidatos = @(
        (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Copia de horario 2026.xlsm'),
        (Join-Path ([Environment]::GetFolderPath('Desktop')) 'Copia de horario 2026.xlsx'),
        (Join-Path "$env:USERPROFILE\Desktop" 'Copia de horario 2026.xlsx'),
        (Join-Path "$env:USERPROFILE\OneDrive\Escritorio" 'Copia de horario 2026.xlsx'),
        (Join-Path "$env:USERPROFILE\Documents" 'Copia de horario 2026.xlsx')
    )
    foreach ($c in $candidatos) { if (Test-Path -LiteralPath $c) { return (Resolve-Path -LiteralPath $c).Path } }

    # Ultimo recurso: cualquier libro de horario del escritorio, con macros o sin ellas.
    $esc = [Environment]::GetFolderPath('Desktop')
    if (Test-Path -LiteralPath $esc) {
        $h = Get-ChildItem -LiteralPath $esc -File -ErrorAction SilentlyContinue |
             Where-Object { $_.Name -like '*horario*' -and $_.Extension -in '.xlsm', '.xlsx' } |
             Sort-Object LastWriteTime -Descending | Select-Object -First 1
        if ($h) { return $h.FullName }
    }
    return $null
}

# Jornada oficial. En verano (15-jun a 15-sep, ambos incluidos) son 7:00; el resto del anio,
# 7:43. Se guarda en configuracion.json para poder cambiarlo sin tocar el codigo si algun anio
# se mueven las fechas o la duracion.
$script:JornadaPorDefecto = @{
    JornadaVerano   = '7:00'
    JornadaInvierno = '7:43'
    VeranoDesde     = '15-06'
    VeranoHasta     = '15-09'
}

# La direccion del portal de fichajes NO va en el codigo: es interna de cada organizacion y
# este repositorio es publico. Vive en configuracion.json, que esta fuera del control de
# versiones. Instalar.ps1 la pregunta la primera vez.
$script:UrlIntranetEjemplo = 'https://intranet.empresa.local/RRHH/Fichajes.aspx?Tab=4'

function Set-ConfigHorario {
    param([Parameter(Mandatory)][string]$RutaLibro, [int]$MaxBackups = 30, [hashtable]$Jornada,
          [string]$UrlIntranet)

    $j = @{} + $script:JornadaPorDefecto
    if ($Jornada) { foreach ($k in $Jornada.Keys) { $j[$k] = $Jornada[$k] } }

    # Si no se pasa, conservar la que ya hubiera guardada.
    if (-not $UrlIntranet -and (Test-Path -LiteralPath $script:FicheroConfig)) {
        try { $UrlIntranet = ([IO.File]::ReadAllText($script:FicheroConfig) | ConvertFrom-Json).UrlIntranet } catch { }
    }

    [pscustomobject]@{
        RutaLibro       = $RutaLibro
        UrlIntranet     = $UrlIntranet
        MaxBackups      = $MaxBackups
        JornadaVerano   = $j.JornadaVerano
        JornadaInvierno = $j.JornadaInvierno
        VeranoDesde     = $j.VeranoDesde
        VeranoHasta     = $j.VeranoHasta
        Guardado        = (Get-Date).ToString('s')
    } | ConvertTo-Json | ForEach-Object { [IO.File]::WriteAllText($script:FicheroConfig, $_, (New-Object Text.UTF8Encoding($false))) }
}

function ConvertTo-FraccionHoras {
    param([string]$HHmm)   # "7:43"
    $p = $HHmm.Split(':')
    return ([double]$p[0] * 60 + [double]$p[1]) / 1440.0
}

function Get-JornadaOficial {
    <#
        Devuelve la jornada que corresponde a una fecha, como fraccion de dia.
        Verano: del 15 de junio al 15 de septiembre, ambos incluidos.
    #>
    param([Parameter(Mandatory)][datetime]$Fecha, [hashtable]$Config)

    $c = if ($Config) { $Config } else { Get-ConfigHorario }

    $d1 = $c.VeranoDesde.Split('-'); $d2 = $c.VeranoHasta.Split('-')
    $ini = Get-Date -Year $Fecha.Year -Month ([int]$d1[1]) -Day ([int]$d1[0]) -Hour 0 -Minute 0 -Second 0
    $fin = Get-Date -Year $Fecha.Year -Month ([int]$d2[1]) -Day ([int]$d2[0]) -Hour 0 -Minute 0 -Second 0

    # Ojo: PowerShell 5.1 NO admite "if" como expresion dentro de una llamada; parsea pero
    # falla en ejecucion. Hay que pasar por una variable.
    $esVerano = ($Fecha.Date -ge $ini.Date -and $Fecha.Date -le $fin.Date)
    $texto = if ($esVerano) { $c.JornadaVerano } else { $c.JornadaInvierno }
    return (ConvertTo-FraccionHoras $texto)
}

function Get-ConfigHorario {
    <#
        Devuelve @{ RutaLibro; CarpetaBackups; MaxBackups }.
        La carpeta de copias se deriva SIEMPRE de la ubicacion del libro, para que al mover
        el libro las copias lo sigan sin tener que reconfigurar nada.
    #>
    $ruta = $null
    $max  = 30
    $url  = $null
    $j    = @{} + $script:JornadaPorDefecto

    if (Test-Path -LiteralPath $script:FicheroConfig) {
        try {
            $c = [IO.File]::ReadAllText($script:FicheroConfig) | ConvertFrom-Json
            if ($c.RutaLibro)  { $ruta = $c.RutaLibro }
            if ($c.MaxBackups) { $max  = [int]$c.MaxBackups }
            if ($c.UrlIntranet) { $url = $c.UrlIntranet }
            foreach ($k in @($script:JornadaPorDefecto.Keys)) {
                if ($c.PSObject.Properties[$k] -and $c.$k) { $j[$k] = $c.$k }
            }
        } catch { }
    }

    # Sin configuracion, o apuntando a un libro que ya no esta: intentar localizarlo.
    if (-not $ruta -or -not (Test-Path -LiteralPath $ruta)) {
        $encontrado = Find-LibroHorario
        if ($encontrado) {
            $ruta = $encontrado
            Set-ConfigHorario -RutaLibro $ruta -MaxBackups $max   # se recuerda para la proxima
        }
    }

    if (-not $ruta -or -not (Test-Path -LiteralPath $ruta)) {
        throw "No se encuentra el libro de horario. Ejecuta Instalar.ps1 para indicar donde esta."
    }

    return @{
        RutaLibro       = $ruta
        UrlIntranet     = $url
        CarpetaBackups  = (Join-Path (Split-Path -Parent $ruta) '_backups_horario')
        MaxBackups      = $max
        JornadaVerano   = $j.JornadaVerano
        JornadaInvierno = $j.JornadaInvierno
        VeranoDesde     = $j.VeranoDesde
        VeranoHasta     = $j.VeranoHasta
    }
}

function Get-ConfigHorarioSegura {
    <# Igual que Get-ConfigHorario, pero devuelve @{} en vez de lanzar cuando aun no hay nada
       configurado. La usa el instalador, que precisamente corre antes de que exista config. #>
    try { return Get-ConfigHorario } catch { }
    if (Test-Path -LiteralPath $script:FicheroConfig) {
        try { 
            $c = [IO.File]::ReadAllText($script:FicheroConfig) | ConvertFrom-Json
            return @{ RutaLibro = $c.RutaLibro; UrlIntranet = $c.UrlIntranet }
        } catch { }
    }
    return @{ RutaLibro = $null; UrlIntranet = $null }
}