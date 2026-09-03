<#
.SYNOPSIS
    Sincroniza los fichajes del portal de RRHH con un libro de Excel de control horario.

.DESCRIPTION
    Lee los fichajes desde la intranet (la URL configurada en configuracion.json) usando
    la autenticacion integrada de Windows -- no necesita contrasena -- y los escribe en las
    celdas de entrada manual del libro de horario.

    Reglas de escritura acordadas:
      * Solo se escriben fichajes REALES presentes en la web.
      * Nunca se borra nada: las previsiones de dias futuros y la salida aun no fichada de
        hoy se quedan intactas.
      * Semana en curso  -> corrige valores existentes si difieren del fichaje real.
      * Semanas anteriores -> solo rellena celdas vacias (respeta ajustes manuales ya hechos).

    Mapa del libro (una hoja por semana, columna C = unicas celdas manuales):
        base = 10 + 8*d   con d = 0(lunes) .. 4(viernes)
        base+1 Entrada | base+2 Salida desayuno | base+3 Entrada desayuno
        base+4 Salida comer | base+5 Entrada comer | base+6 Salida

.PARAMETER SemanasAtras
    Cuantas semanas anteriores revisar ademas de la actual. Por defecto 1, porque la propia
    intranet advierte que los fichajes pueden tardar hasta 72 h en consolidarse.

.PARAMETER ModoPrueba
    Muestra lo que haria sin tocar el Excel.

.PARAMETER CrearHojaSiFalta
    Si no existe la hoja de la semana, la crea copiando la ultima y limpiando los fichajes.
    Desactivado por defecto: revisa a mano "Horas semanales" (C3) y "Renta semana" (C5).
#>
[CmdletBinding()]
param(
    [int]    $SemanasAtras     = 1,
    [switch] $ModoPrueba,
    [switch] $CrearHojaSiFalta,
    # Crear ademas la hoja de una semana concreta (cualquier dia de esa semana vale), aunque
    # sea futura. Para adelantarse a la semana que viene sin esperar al lunes.
    [datetime] $CrearSemana = [datetime]::MinValue,
    # Toma los fichajes del portapapeles en vez de la intranet: para cuando estas fuera de la
    # oficina y has entrado al portal externo a mano (DNI + hCaptcha + usuario/contrasena).
    # Selecciona la tabla de fichajes en el navegador, Ctrl+C, y lanza el script con este flag.
    [switch] $DesdePortapapeles,
    # Corrige tambien semanas pasadas en vez de limitarse a rellenar celdas vacias.
    [switch] $Forzar,
    # No computar como jornada los dias laborables ya pasados sin ningun fichaje.
    [switch] $SinAusencias,
    # Dias que han de haber pasado para dar un dia por ausencia. Con 1, ayer ya cuenta.
    # Subelo si tus fichajes tardan en consolidarse y salen falsas ausencias.
    [int]    $DiasMargenAusencia = 1,
    # Si no se indica, sale de configuracion.json (lo deja escrito Instalar.ps1).
    [string] $Libro,
    # Direccion del portal de fichajes. No va en el codigo (es interna de cada organizacion):
    # sale de configuracion.json, que Instalar.ps1 rellena la primera vez.
    [string] $UrlIntranet,
    [int]    $MaxBackups       = 30,
    # Una pausa de hasta estos minutos es el descanso (el cafe, normalmente sobre las 11h);
    # mas larga que eso, se considera la comida. Se distingue por DURACION y no por la hora
    # del dia, que clasificaria mal una comida temprana.
    [int]    $MaxMinutosDescanso = 60
)

$ErrorActionPreference = 'Stop'
$ProgressPreference    = 'SilentlyContinue'

$RaizScript = Split-Path -Parent $MyInvocation.MyCommand.Path
$LogFile    = Join-Path $RaizScript 'sync-horario.log'
$AvisoFile  = Join-Path $RaizScript 'ultimo-resultado.txt'

. (Join-Path $RaizScript 'Configuracion.ps1')
$script:Cfg = Get-ConfigHorario
if (-not $Libro) {
    $Libro = $script:Cfg.RutaLibro
    if (-not $PSBoundParameters.ContainsKey('MaxBackups')) { $MaxBackups = $script:Cfg.MaxBackups }
}
if (-not $UrlIntranet) { $UrlIntranet = $script:Cfg.UrlIntranet }

# Las copias viven junto al libro: si se mueve el libro, se mueven con el.
$CarpetaBak = Join-Path (Split-Path -Parent $Libro) '_backups_horario'

$script:Resumen  = New-Object System.Collections.Generic.List[string]
$script:Avisos   = New-Object System.Collections.Generic.List[string]
$script:HojasAvisadas = New-Object System.Collections.Generic.HashSet[string]
$script:Cambios  = 0

function Write-Log {
    <#
        -SoloLog deja constancia en el fichero pero no lo muestra por pantalla: para las
        lineas de traza que sirven al diagnosticar un fallo pero no dicen nada al usuario.
    #>
    param(
        [string] $Mensaje,
        [ValidateSet('INFO','AVISO','ERROR','OK')][string] $Nivel = 'INFO',
        [switch] $SoloLog
    )
    $linea = "{0}  [{1,-5}] {2}" -f (Get-Date -Format 'yyyy-MM-dd HH:mm:ss'), $Nivel, $Mensaje
    Add-Content -Path $LogFile -Value $linea -Encoding utf8

    if (-not $SoloLog) {
        switch ($Nivel) {
            'ERROR' { Write-Host $linea -ForegroundColor Red }
            'AVISO' { Write-Host $linea -ForegroundColor Yellow }
            'OK'    { Write-Host $linea -ForegroundColor Green }
            default { Write-Host $linea }
        }
    }
    if ($Nivel -in 'AVISO','ERROR') { $script:Avisos.Add($Mensaje) }
}

# ---------------------------------------------------------------- utilidades

$MesesEs = @{
    'ene'=1; 'feb'=2; 'mar'=3; 'abr'=4; 'may'=5; 'jun'=6
    'jul'=7; 'ago'=8; 'sep'=9; 'oct'=10; 'nov'=11; 'dic'=12
}

function Get-LunesDe {
    param([datetime]$Fecha)
    # DayOfWeek: domingo = 0. Queremos lunes como inicio de semana.
    $offset = (([int]$Fecha.DayOfWeek) + 6) % 7
    return $Fecha.Date.AddDays(-$offset)
}

function ConvertTo-FraccionDia {
    param([string]$Hora)   # "09:17"
    $p = $Hora.Split(':')
    return ([double]$p[0] * 60 + [double]$p[1]) / 1440.0
}

$script:SetProp = [System.Reflection.BindingFlags]::SetProperty

function Set-Valor {
    <#
        Escribe en una celda SIEMPRE por InvokeMember, nunca con "$celda.Value2 = ...".
        El adaptador COM de PowerShell cachea el tipo del setter en su primer uso: si en una
        misma ejecucion se escribe primero texto (las etiquetas de una hoja recien creada) y
        luego un numero (las horas), la segunda escritura lanza InvalidCastException.
    #>
    param($Celda, $Valor)
    [void]$Celda.GetType().InvokeMember('Value2', $script:SetProp, $null, $Celda, @($Valor))
}

function Get-CampoOculto {
    param([string]$Html, [string]$Nombre)
    $m = [regex]::Match($Html, "(?is)<input[^>]*name=""$([regex]::Escape($Nombre))""[^>]*value=""([^""]*)""")
    if ($m.Success) { return $m.Groups[1].Value }
    $m = [regex]::Match($Html, "(?is)<input[^>]*value=""([^""]*)""[^>]*name=""$([regex]::Escape($Nombre))""")
    if ($m.Success) { return $m.Groups[1].Value }
    return ''
}

# ------------------------------------------------------- lectura de fichajes

function ConvertFrom-TablaFichajes {
    <#
        Devuelve un hashtable: [datetime]::Date -> objeto con
            .Marcas  string[] de horas "HH:mm" fichadas
            .Diario  horas que la WEB computa a ese dia (fraccion de dia), o $null
            .Nota    texto del tooltip de esa celda ("Vacaciones", ...), o $null

        La fila "Diario:" es la clave para tratar las ausencias: un dia de vacaciones NO tiene
        marcas pero SI trae ahi las horas que la empresa da por trabajadas (p. ej. 7:43, que no
        tiene por que coincidir con la jornada del libro). Un dia sin marcas Y sin "Diario:" es
        otra cosa -- un olvido de fichar, por ejemplo -- y no se computa.
    #>
    param([string]$Html, [datetime]$LunesSemana)

    $resultado = @{}

    # Las comillas del atributo se aceptan dobles, simples o ausentes: al copiar desde el
    # navegador el HTML no siempre conserva el formato original.
    $filaFechas = [regex]::Match($Html, '(?is)<tr[^>]*class=["'']?fechas\b[^>]*>(.*?)</tr>')
    $filaHoras  = [regex]::Match($Html, '(?is)<tr[^>]*class=["'']?horas\b[^>]*>(.*?)</tr>')
    if (-not $filaFechas.Success -or -not $filaHoras.Success) {
        throw 'No se encontro la tabla de fichajes en el HTML (cambio la pagina?).'
    }

    $tdFechas = [regex]::Matches($filaFechas.Groups[1].Value, '(?is)<td[^>]*>(.*?)</td>')
    $tdHoras  = [regex]::Matches($filaHoras.Groups[1].Value,  '(?is)<td[^>]*>(.*?)</td>')

    # Fila "Diario:" (class="balance"): mismas posiciones de columna que las anteriores.
    $filaBalance = [regex]::Match($Html, '(?is)<tr[^>]*class=["'']?balance\b[^>]*>(.*?)</tr>')
    # Se captura la celda ENTERA, con su tag de apertura: el tooltip ("Vacaciones") viaja en
    # el atributo title del <td>, no en su contenido.
    $tdBalance = if ($filaBalance.Success) {
        [regex]::Matches($filaBalance.Groups[1].Value, '(?is)<td[^>]*>.*?</td>')
    } else { @() }

    # La primera celda de cada fila es la etiqueta vacia de la izquierda.
    for ($i = 1; $i -lt $tdFechas.Count; $i++) {
        $txtFecha = ($tdFechas[$i].Groups[1].Value -replace '<[^>]+>', '').Trim()
        $mf = [regex]::Match($txtFecha, '^(\d{1,2})-([a-zA-Zéó]{3})')
        if (-not $mf.Success) { continue }

        $dia = [int]$mf.Groups[1].Value
        $abr = $mf.Groups[2].Value.ToLower() -replace 'é','e' -replace 'ó','o'
        if (-not $MesesEs.ContainsKey($abr)) { continue }

        # La web solo da "dd-mmm" sin anio: se deduce del lunes de la semana consultada,
        # corrigiendo el salto cuando la semana cruza el fin de anio (dic -> ene).
        $fecha = Get-Date -Year $LunesSemana.Year -Month $MesesEs[$abr] -Day $dia -Hour 0 -Minute 0 -Second 0
        if     ($fecha -lt $LunesSemana.AddDays(-2)) { $fecha = $fecha.AddYears(1) }
        elseif ($fecha -gt $LunesSemana.AddDays(8))  { $fecha = $fecha.AddYears(-1) }

        # Las horas se buscan sobre el TEXTO de la celda, no sobre un <span> literal: al copiar
        # desde el navegador el marcado cambia (spans con style, clases, etc.) y un patron
        # como '<span>09:17</span>' deja de encontrar nada, con lo que el dia sale vacio.
        $marcas = @()
        if ($i -lt $tdHoras.Count) {
            $textoCelda = $tdHoras[$i].Groups[1].Value -replace '<[^>]+>', ' '
            $marcas = [regex]::Matches($textoCelda, '\b([01]\d|2[0-3]):([0-5]\d)\b') |
                      ForEach-Object { $_.Value }
        }

        # Horas que la web computa al dia, y el tooltip si lo lleva ("Vacaciones", ...).
        $diario = $null; $nota = $null
        if ($i -lt $tdBalance.Count) {
            $bruto = $tdBalance[$i].Value
            $mt = [regex]::Match($bruto, '(?i)title\s*=\s*["'']([^"'']+)["'']')
            # Se ignoran los title decorativos que la propia tabla pone en fin de semana.
            if ($mt.Success -and $mt.Groups[1].Value.Trim() -notmatch '(?i)^(s.bado|domingo)$') {
                $nota = $mt.Groups[1].Value.Trim()
            }

            $md = [regex]::Match(($bruto -replace '<[^>]+>', ' '), '\b(\d{1,2}):([0-5]\d)\b')
            if ($md.Success) {
                $diario = ([double]$md.Groups[1].Value * 60 + [double]$md.Groups[2].Value) / 1440.0
            }
        }

        $resultado[$fecha.Date] = [pscustomobject]@{
            Marcas = @($marcas)
            Diario = $diario
            Nota   = $nota
        }
    }
    return $resultado
}

function Get-FichajesIntranet {
    param([datetime[]]$LunesAConsultar)

    if (-not $UrlIntranet) {
        throw 'No hay configurada la direccion del portal de fichajes. Ejecuta Instalar.ps1 para indicarla.'
    }
    Write-Log "Conectando a la intranet ($($LunesAConsultar.Count) semana(s))..." -SoloLog
    $sesion = $null
    $primera = Invoke-WebRequest -Uri $UrlIntranet -UseDefaultCredentials -UseBasicParsing `
                                 -TimeoutSec 30 -SessionVariable sesion

    if ($primera.Content -notmatch 'tabla_fichajes') {
        throw 'La intranet respondio pero sin la tabla de fichajes (sesion caducada?).'
    }

    $tarjeta = [regex]::Match($primera.Content, 'id="num_tarjeta">([^<]*)<').Groups[1].Value
    $balance = [regex]::Match($primera.Content, 'id="lblBalance">([^<]*)<').Groups[1].Value
    if ($tarjeta) { Write-Log "$tarjeta | $balance (balance OFICIAL de la web; tu saldo del libro va aparte)" }

    $todos = @{}
    $htmlActual = $primera.Content

    foreach ($lunes in $LunesAConsultar) {
        $clave = $lunes.ToString('dd-MM-yyyy')
        $seleccionada = [regex]::Match($htmlActual, '(?is)<option selected="selected" value="([^"]+)"').Groups[1].Value

        if ($seleccionada -ne $clave) {
            if ($htmlActual -notmatch [regex]::Escape("value=""$clave""")) {
                Write-Log "La semana del $clave no esta disponible en el desplegable; se omite." 'AVISO'
                continue
            }
            $body = @{
                '__EVENTTARGET'        = 'ddl_semanas'
                '__EVENTARGUMENT'      = ''
                '__LASTFOCUS'          = ''
                '__VIEWSTATE'          = (Get-CampoOculto $htmlActual '__VIEWSTATE')
                '__VIEWSTATEGENERATOR' = (Get-CampoOculto $htmlActual '__VIEWSTATEGENERATOR')
                'ddl_semanas'          = $clave
            }
            $resp = Invoke-WebRequest -Uri $UrlIntranet -Method POST -Body $body -UseDefaultCredentials `
                                      -UseBasicParsing -TimeoutSec 30 -WebSession $sesion
            $htmlActual = $resp.Content
        }

        foreach ($kv in (ConvertFrom-TablaFichajes -Html $htmlActual -LunesSemana $lunes).GetEnumerator()) {
            $todos[$kv.Key] = $kv.Value
        }
    }
    return $todos
}

function Get-FichajesPortapapeles {
    <#
        Lee la tabla de fichajes copiada del navegador. Prioriza el formato HTML del
        portapapeles, que conserva las clases "fechas"/"horas" y permite reutilizar el mismo
        parser que la intranet. Si solo hay texto plano, cae a una lectura por lineas.
    #>
    Add-Type -AssemblyName System.Windows.Forms

    $html = ''
    try { $html = [Windows.Forms.Clipboard]::GetText([Windows.Forms.TextDataFormat]::Html) } catch { }

    if ($html -match 'class=["'']?fechas\b' -and $html -match 'class=["'']?horas\b') {
        Write-Log 'Portapapeles: detectada la tabla en formato HTML.'
        # El lunes no viene en el fragmento copiado: se deduce del primer dia listado,
        # eligiendo el anio que lo deje mas cerca de hoy.
        $m = [regex]::Match($html, '>\s*(\d{1,2})-([a-zA-Zéó]{3})\s*<')
        if (-not $m.Success) { throw 'No se pudo leer ninguna fecha de la tabla copiada.' }

        $abr = $m.Groups[2].Value.ToLower() -replace 'é','e' -replace 'ó','o'
        if (-not $MesesEs.ContainsKey($abr)) { throw "Mes no reconocido: '$abr'." }

        $hoy = (Get-Date).Date
        $mejor = $null
        foreach ($a in ($hoy.Year - 1), $hoy.Year, ($hoy.Year + 1)) {
            try { $f = Get-Date -Year $a -Month $MesesEs[$abr] -Day ([int]$m.Groups[1].Value) -Hour 0 -Minute 0 -Second 0 } catch { continue }
            if ($null -eq $mejor -or [Math]::Abs(($f - $hoy).TotalDays) -lt [Math]::Abs(($mejor - $hoy).TotalDays)) { $mejor = $f.Date }
        }
        $lunes = Get-LunesDe $mejor
        Write-Log "Portapapeles: semana del $($lunes.ToString('dd-MM-yyyy'))."
        return (ConvertFrom-TablaFichajes -Html $html -LunesSemana $lunes)
    }

    $texto = ''
    try { $texto = Get-Clipboard -Raw } catch { }
    if ([string]::IsNullOrWhiteSpace($texto)) {
        throw 'El portapapeles esta vacio. Selecciona la tabla de fichajes en el navegador, pulsa Ctrl+C y vuelve a lanzarlo.'
    }

    Write-Log 'Portapapeles: texto plano, sin formato de tabla.' -SoloLog
    return (ConvertFrom-TextoFichajes -Texto $texto)
}

function ConvertFrom-TextoFichajes {
    <#
        Interpreta la tabla de fichajes pegada como TEXTO PLANO.

        Es mucho mas dificil que el HTML porque se pierde la estructura de columnas: los siete
        dias van en una sola linea y despues las horas caen sueltas, sin indicar a que dia
        pertenece cada una:

            lunes   martes  miercoles ...
            24-ago  25-ago  26-ago  ...
            09:17
            16:26
            08:22            <- ya es del martes, pero nada lo dice
            ...
            Diario: 7:09    7:56    7:16 ...

        Se reconstruye asi: dentro de un mismo dia las marcas van en orden creciente, de modo
        que **cuando una hora es menor que la anterior empieza un dia nuevo**. Y el reparto se
        VERIFICA contra la fila "Diario:", que tiene un valor por cada dia con actividad: si el
        numero de grupos no coincide con el de dias con datos, se aborta en vez de escribir un
        reparto que podria estar desplazado.
    #>
    param([string]$Texto)

    $lineas = $Texto -split "`r?`n"
    $hoy    = (Get-Date).Date

    # 1. La fila de fechas es la unica con cinco o mas "dd-mmm" (la cabecera "Del X al Y" tiene dos).
    $idxFechas = -1; $tokens = $null
    for ($i = 0; $i -lt $lineas.Count; $i++) {
        $m = [regex]::Matches($lineas[$i], '\b(\d{1,2})-([a-zA-Zéó]{3})\b')
        if ($m.Count -ge 5) { $idxFechas = $i; $tokens = $m; break }
    }
    if ($idxFechas -lt 0) {
        throw 'No se reconoce la tabla de fichajes en lo copiado. Selecciona la tabla entera, incluida la fila de fechas.'
    }

    $fechas = @()
    foreach ($t in $tokens) {
        $abr = $t.Groups[2].Value.ToLower() -replace 'é','e' -replace 'ó','o'
        if (-not $MesesEs.ContainsKey($abr)) { continue }
        $f = Get-Date -Year $hoy.Year -Month $MesesEs[$abr] -Day ([int]$t.Groups[1].Value) -Hour 0 -Minute 0 -Second 0
        if (($f - $hoy).TotalDays -gt 180) { $f = $f.AddYears(-1) }
        if (($hoy - $f).TotalDays -gt 180) { $f = $f.AddYears(1) }
        $fechas += $f.Date
    }

    # 2. La fila "Diario:" cierra el bloque de horas y dice que dias tienen actividad.
    $idxDiario = -1
    for ($i = $idxFechas + 1; $i -lt $lineas.Count; $i++) {
        if ($lineas[$i] -match '^\s*Diario') { $idxDiario = $i; break }
    }

    $diasConDatos = @()
    $diarios = @{}
    if ($idxDiario -ge 0) {
        $campos = $lineas[$idxDiario] -split "`t"
        for ($d = 0; $d -lt $fechas.Count; $d++) {
            if ($campos.Count -le ($d + 1)) { continue }
            $v = $campos[$d + 1].Trim()
            if ($v -eq '') { continue }
            $diasConDatos += $d
            $mv = [regex]::Match($v, '\b(\d{1,2}):([0-5]\d)\b')
            if ($mv.Success) { $diarios[$d] = ([double]$mv.Groups[1].Value * 60 + [double]$mv.Groups[2].Value) / 1440.0 }
        }
    }

    # 3. Horas entre la fila de fechas y la de "Diario" (asi se excluye el "Balance de horas").
    $fin = if ($idxDiario -ge 0) { $idxDiario } else { $lineas.Count }
    $horas = @()
    for ($i = $idxFechas + 1; $i -lt $fin; $i++) {
        foreach ($h in [regex]::Matches($lineas[$i], '\b([01]\d|2[0-3]):([0-5]\d)\b')) { $horas += $h.Value }
    }
    # 4. Agrupar: una hora menor que la anterior abre un dia nuevo.
    $grupos = @(); $actual = @(); $previo = -1
    foreach ($h in $horas) {
        $p = $h.Split(':'); $min = [int]$p[0] * 60 + [int]$p[1]
        if ($min -lt $previo) { $grupos += ,$actual; $actual = @() }
        $actual += $h; $previo = $min
    }
    if ($actual.Count -gt 0) { $grupos += ,$actual }

    $resultado = @{}

    # 5. Repartir, contrastando contra "Diario:".
    if ($diasConDatos.Count -eq 0) {
        if ($grupos.Count -eq 0) { throw 'No se ha encontrado ninguna hora de fichaje en lo copiado.' }
        Write-Log 'Lo copiado no trae la fila "Diario:", asi que el reparto por dias no se ha podido verificar. Revisa el resultado.' 'AVISO'
        $diasConDatos = 0..($grupos.Count - 1)
    }
    elseif ($grupos.Count -eq 0) {
        # Dias con horas computadas pero SIN ninguna marca: semana completa de vacaciones.
        Write-Log "Lo copiado no tiene ninguna marca de fichaje, pero si horas en 'Diario:': se tratara como dias de ausencia."
    }
    elseif ($grupos.Count -ne $diasConDatos.Count) {
        throw ("No se puede repartir con seguridad lo copiado: salen $($grupos.Count) grupo(s) de marcas " +
               "pero la fila 'Diario:' indica $($diasConDatos.Count) dia(s) con horas. " +
               "Copia la tabla con formato (seleccionandola en el navegador) en vez de como texto suelto.")
    }

    for ($k = 0; $k -lt $diasConDatos.Count; $k++) {
        $d = $diasConDatos[$k]
        if ($d -ge $fechas.Count) { continue }
        $resultado[$fechas[$d]] = [pscustomobject]@{
            Marcas = if ($k -lt $grupos.Count) { @($grupos[$k]) } else { @() }
            Diario = if ($diarios.ContainsKey($d)) { $diarios[$d] } else { $null }
            Nota   = $null   # el texto del tooltip se pierde al pegar como texto plano
        }
    }

    Write-Log "Portapapeles (texto): $($resultado.Count) dia(s) reconocido(s) y verificados contra la fila 'Diario:'."
    return $resultado
}

# ------------------------------------------------------------ acceso a Excel

function Get-LunesDeHoja {
    <# Extrae el lunes a partir del nombre de la hoja: "24-08-26 a 28-08-26" -> 24/08/2026.
       Los nombres del libro tienen erratas en el segundo tramo, pero el primero es fiable. #>
    param([string]$Nombre, [int]$AnioPorDefecto)
    $m = [regex]::Match($Nombre.Trim(), '^(\d{1,2})-(\d{1,2})(?:-(\d{2,4}))?')
    if (-not $m.Success) { return $null }
    $d = [int]$m.Groups[1].Value
    $mes = [int]$m.Groups[2].Value
    if ($mes -lt 1 -or $mes -gt 12 -or $d -lt 1 -or $d -gt 31) { return $null }
    $anio = $AnioPorDefecto
    if ($m.Groups[3].Success) {
        $a = [int]$m.Groups[3].Value
        $anio = if ($a -lt 100) { 2000 + $a } else { $a }
    }
    try { return (Get-Date -Year $anio -Month $mes -Day $d -Hour 0 -Minute 0 -Second 0).Date }
    catch { return $null }
}

function Find-HojaSemana {
    param($Workbook, [datetime]$Lunes)
    foreach ($hoja in $Workbook.Sheets) {
        $l = Get-LunesDeHoja -Nombre $hoja.Name -AnioPorDefecto $Lunes.Year
        if ($null -ne $l -and $l -eq $Lunes.Date) {
            # El nombre de la hoja es la fuente fiable. Las etiquetas B10..B42 arrastran
            # erratas de hojas copiadas, asi que su discrepancia se anota pero no bloquea ni
            # cuenta como aviso.
            # El numero se compara COMO NUMERO: el libro lo escribe a dos digitos ("Lunes 03")
            # y una comparacion de texto da falsos positivos.
            $b10 = "$($hoja.Range('B10').Text)"
            $num = [regex]::Match($b10, '(\d{1,2})\s*$').Groups[1].Value
            if ("$num" -eq '' -or [int]$num -ne $Lunes.Day) {
                # Una misma hoja se busca varias veces por ejecucion: avisar solo la primera.
                if (-not $script:HojasAvisadas.Contains($hoja.Name)) {
                    [void]$script:HojasAvisadas.Add($hoja.Name)
                    Write-Log "Nota: la hoja '$($hoja.Name)' tiene la etiqueta B10='$b10' sin actualizar. Los datos se escriben igual (manda el nombre de la hoja)."
                }
            }
            return $hoja
        }
    }
    return $null
}

function New-HojaSemana {
    <#
        Crea la hoja de una semana copiando la anterior: hereda formato, formulas y los
        parametros de cabecera (jornada, horarios limite, tiempos de pausa).

        Deja los fichajes en blanco y reetiqueta los dias. "Renta semana" (C5) no se toca
        aqui: la pone Update-RentaSemana con la renta real del viernes anterior.
    #>
    param($Workbook, [datetime]$Lunes)

    $anterior = Find-HojaSemana -Workbook $Workbook -Lunes $Lunes.AddDays(-7)
    if (-not $anterior) {
        Write-Log "No se puede crear la hoja del $($Lunes.ToString('dd-MM-yyyy')): falta la de la semana anterior ($($Lunes.AddDays(-7).ToString('dd-MM-yyyy')))." 'AVISO'
        return $null
    }

    $anterior.Copy([System.Reflection.Missing]::Value, $anterior)
    $nueva = $Workbook.Sheets.Item($anterior.Index + 1)
    $nueva.Name = "{0} a {1}" -f $Lunes.ToString('dd-MM-yy'), $Lunes.AddDays(4).ToString('dd-MM-yy')

    $ci = [Globalization.CultureInfo]::GetCultureInfo('es-ES')

    for ($d = 0; $d -lt 5; $d++) {
        $base  = 10 + 8 * $d
        $fecha = $Lunes.AddDays($d)

        # Vaciar las seis celdas manuales del dia (entrada, pausas y salida).
        foreach ($off in 1, 2, 3, 4, 5, 6) { $nueva.Cells.Item($base + $off, 3).ClearContents() | Out-Null }

        # Reetiquetar conservando como escribe el libro el nombre del dia (mezcla de
        # mayusculas propia, "Miercoles" sin tilde) y su numero a dos digitos.
        $celda  = $nueva.Cells.Item($base, 2)
        $actual = "$($celda.Text)"
        $nombreDia = ($actual -replace '\d', '').Trim()
        if ([string]::IsNullOrWhiteSpace($nombreDia)) {
            $nombreDia = $fecha.ToString('dddd', $ci)
            $nombreDia = $nombreDia.Substring(0, 1).ToUpper() + $nombreDia.Substring(1)
        }
        Set-Valor -Celda $celda -Valor ('{0} {1:00}' -f $nombreDia, $fecha.Day)
    }

    # Mes de la cabecera: el de la MAYORIA de los dias laborables, no el del lunes. La semana
    # del 31-08-26 al 04-09-26, por ejemplo, es de septiembre aunque empiece en agosto.
    if ("$($nueva.Range('I1').Formula)" -notlike '=*') {
        $meses = @{}
        for ($d = 0; $d -lt 5; $d++) {
            $m = $Lunes.AddDays($d).ToString('MMMM', $ci)
            if (-not $meses.ContainsKey($m)) { $meses[$m] = 0 }
            $meses[$m]++
        }
        $mesMayoritario = ($meses.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
        Set-Valor -Celda $nueva.Range('I1') -Valor $mesMayoritario
    }

    # Jornada oficial: 7:00 en verano (15-jun a 15-sep) y 7:43 el resto del anio. Se pone aqui
    # porque al copiar la semana anterior se heredaria la jornada vieja, y en la semana del
    # cambio eso descuadraria toda la hoja.
    $cuenta = @{}
    for ($d = 0; $d -lt 5; $d++) {
        $j = Get-JornadaOficial -Fecha $Lunes.AddDays($d) -Config $script:Cfg
        if (-not $cuenta.ContainsKey($j)) { $cuenta[$j] = 0 }
        $cuenta[$j]++
    }
    $jornada = ($cuenta.GetEnumerator() | Sort-Object Value -Descending | Select-Object -First 1).Key
    $tj = [timespan]::FromDays([double]$jornada)
    $textoJ = '{0}:{1:00}' -f [Math]::Floor($tj.TotalHours), $tj.Minutes

    $actualJ = $nueva.Range('C4').Value2
    if ($actualJ -isnot [double] -or [Math]::Abs([double]$actualJ - [double]$jornada) -ge (0.5/1440.0)) {
        Set-Valor -Celda $nueva.Range('C4') -Valor ([double]$jornada)
        Write-Log "Jornada de la hoja nueva ajustada a $textoJ (la heredada no correspondia a estas fechas)."
    }
    if ($cuenta.Count -gt 1) {
        Write-Log "La semana del $($Lunes.ToString('dd-MM-yyyy')) cae sobre el cambio de jornada de verano: se ha puesto $textoJ, que es la de la mayoria de sus dias. REVISALA." 'AVISO'
    }

    Write-Log "Creada la hoja '$($nueva.Name)' a partir de '$($anterior.Name)'. Revisa C3 (horas semanales) si cambia." 'AVISO'
    return $nueva
}

function Set-CeldaHora {
    <# Escribe una marca si procede. Devuelve $true si cambio algo. #>
    param($Hoja, [int]$Fila, [string]$Marca, [string]$Etiqueta, [bool]$SoloSiVacia)

    $celda  = $Hoja.Cells.Item($Fila, 3)
    $actual = $celda.Value2
    $nuevo  = ConvertTo-FraccionDia $Marca

    $vacia = ($null -eq $actual) -or ("$actual" -eq '')
    if (-not $vacia -and $SoloSiVacia) { return $false }

    if (-not $vacia) {
        # Tolerancia de 30 s para no reescribir por redondeo del serial de Excel.
        if ([math]::Abs([double]$actual - $nuevo) -lt (0.5 / 1440.0)) { return $false }
        $antes = [timespan]::FromDays([double]$actual).ToString('hh\:mm')
        $script:Resumen.Add("      ~ $Etiqueta : $antes -> $Marca  (corregido)")
    } else {
        $script:Resumen.Add("      + $Etiqueta : $Marca")
    }

    if (-not $ModoPrueba) { Set-Valor -Celda $celda -Valor $nuevo }
    return $true
}

function Set-Ausencia {
    <#
        Computa como jornada completa un dia laborable ya pasado del que la web no trae ningun
        fichaje: vacaciones, festivo, permiso o baja.

        El motivo no es cosmetico. Sin fichajes, "Horas trabajadas" (M) del dia queda en "?",
        y como el objetivo de cada dia se calcula a partir de la renta del anterior, ese "?" se
        propaga en cadena y descuadra el resto de la semana. Escribiendo una entrada y una
        salida separadas por la jornada, el dia sale neutro y la cadena sigue.

        Se marca con un comentario en la celda del dia para que se distinga de un fichaje real.
    #>
    param($Hoja, [int]$Base, [datetime]$Fecha, [bool]$SoloSiVacia, $HorasWeb, [string]$Nota, [string]$Origen = 'segun la web')

    $entradaActual = $Hoja.Cells.Item($Base + 1, 3).Value2
    $vacia = ($null -eq $entradaActual) -or ("$entradaActual" -eq '')
    if (-not $vacia -and $SoloSiVacia) { return $false }

    $inicio = $Hoja.Range('C6').Value2    # Hora de Inicio
    if ($inicio -isnot [double]) { return $false }

    # Se computan las horas que da la PROPIA WEB para ese dia, no la jornada del libro: en
    # vacaciones la empresa abona la jornada oficial de la epoca (7:43 en junio, por ejemplo),
    # que no tiene por que ser la del libro (7:00). Usar la del libro crearia un deficit falso.
    if ($HorasWeb -isnot [double] -or [double]$HorasWeb -le 0) { return $false }
    $jornada = [double]$HorasWeb

    $salida = [double]$inicio + $jornada

    # Si ya estaba puesto exactamente asi, no hay nada que hacer.
    if (-not $vacia) {
        $salidaActual = $Hoja.Cells.Item($Base + 6, 3).Value2
        if ($salidaActual -is [double] -and
            [Math]::Abs([double]$entradaActual - [double]$inicio) -lt (0.5/1440.0) -and
            [Math]::Abs([double]$salidaActual  - $salida)          -lt (0.5/1440.0)) { return $false }
    }

    if (-not $ModoPrueba) {
        Set-Valor -Celda $Hoja.Cells.Item($Base + 1, 3) -Valor ([double]$inicio)
        Set-Valor -Celda $Hoja.Cells.Item($Base + 6, 3) -Valor $salida
        # Sin trabajar no hubo pausas: si quedaran previstas, descontarian tiempo.
        foreach ($off in 2, 3, 4, 5) { $Hoja.Cells.Item($Base + $off, 3).ClearContents() | Out-Null }

        # Comentario en la cabecera del dia, para que se vea que no es un fichaje real.
        $motivo = if ($Nota) { $Nota } else { 'ausencia, festivo o vacaciones' }
        $celdaDia = $Hoja.Cells.Item($Base, 2)
        try { if ($celdaDia.Comment) { $celdaDia.Comment.Delete() } } catch { }
        try { [void]$celdaDia.AddComment("$motivo. Sin marcas de fichaje: se computan las horas que da la web para este dia.") } catch { }
    }

    $t = [timespan]::FromDays($jornada)
    $etiqueta = if ($Nota) { $Nota } else { 'Ausencia' }
    $script:Resumen.Add("   $($Fecha.ToString('ddd dd-MM'))  [sin marcas]")
    $script:Resumen.Add(("      = {0}: computadas {1}:{2:00} ({3}), de {4} a {5}" -f `
        $etiqueta, [Math]::Floor($t.TotalHours), $t.Minutes, $Origen,
        [timespan]::FromDays([double]$inicio).ToString('hh\:mm'),
        [timespan]::FromDays($salida).ToString('hh\:mm')))
    return $true
}

function Update-Semana {
    param($Hoja, [datetime]$Lunes, [hashtable]$Fichajes, [bool]$EsSemanaActual)

    $soloVacias = -not $EsSemanaActual
    $tocado = $false

    for ($d = 0; $d -lt 5; $d++) {
        $fecha = $Lunes.AddDays($d).Date
        if (-not $Fichajes.ContainsKey($fecha)) { continue }

        $dato   = $Fichajes[$fecha]
        $marcas = @($dato.Marcas)
        $base   = 10 + 8 * $d

        if ($marcas.Count -eq 0) {
            if ($SinAusencias) { continue }

            # La web pone el motivo de la ausencia en el tooltip de la celda "Diario:"
            # ("Vacaciones", "Local" para un festivo local, "Permiso"...). Ese motivo es
            # la senal fiable: significa ausencia JUSTIFICADA, y muchas veces viene sin horas.
            $motivo = "$($dato.Nota)".Trim()

            if ($motivo -eq '') {
                # Sin motivo declarado solo cabe el olvido de fichar, y unicamente se mira en
                # dias ya pasados (uno reciente puede estar sin consolidar: hasta 72 h).
                $limite = (Get-Date).Date.AddDays(-$DiasMargenAusencia)
                if ($fecha -gt $limite) { continue }
                if ($dato.Diario -isnot [double]) {
                    Write-Log "$($fecha.ToString('dd-MM-yyyy')) no tiene marcas, ni motivo, ni horas en 'Diario:'. No se computa: parece un olvido de fichar." 'AVISO'
                    continue
                }
            }
            # Con motivo declarado se computa aunque el dia sea futuro: es una ausencia ya
            # confirmada por la empresa, no una suposicion.

            # Horas: las que de la web si las trae; si no (lo normal en permisos y festivos),
            # la jornada oficial que corresponda a esa fecha.
            $horas  = $dato.Diario
            $origen = 'segun la web'
            if ($horas -isnot [double] -or [double]$horas -le 0) {
                $horas  = Get-JornadaOficial -Fecha $fecha -Config $script:Cfg
                $origen = 'jornada oficial de esa fecha'
            }

            if (Set-Ausencia -Hoja $Hoja -Base $base -Fecha $fecha -SoloSiVacia $soloVacias -HorasWeb $horas -Nota $motivo -Origen $origen) {
                $q = [timespan]::FromDays([double]$horas)
                $m = if ($motivo -ne '') { $motivo } else { 'sin marcas de fichaje' }
                Write-Log ("{0}: {1}. Se computan {2}:{3:00} ({4})." -f `
                    $fecha.ToString('dd-MM-yyyy'), $m, [Math]::Floor($q.TotalHours), $q.Minutes, $origen) 'AVISO'
                $tocado = $true
            }
            continue
        }

        $etiquetaDia = $fecha.ToString('ddd dd-MM')

        $lineasAntes = $script:Resumen.Count
        $cambioReal  = $false   # solo cuenta la escritura efectiva, no las lineas informativas

        # Los fichajes van en pares entrada/salida. Si el numero de marcas es IMPAR la jornada
        # sigue abierta y la ultima marca es una vuelta de pausa (p.ej. el cafe), NO la salida
        # del dia: tomarla por salida machacaria la hora prevista con la del cafe.
        $jornadaCerrada = ($marcas.Count % 2 -eq 0)

        if (Set-CeldaHora -Hoja $Hoja -Fila ($base + 1) -Marca $marcas[0] -Etiqueta 'Entrada' -SoloSiVacia $soloVacias) { $cambioReal = $true }

        if ($jornadaCerrada) {
            if (Set-CeldaHora -Hoja $Hoja -Fila ($base + 6) -Marca $marcas[-1] -Etiqueta 'Salida' -SoloSiVacia $soloVacias) { $cambioReal = $true }
        } else {
            # Solo hay entrada: se informa de la salida prevista que se esta respetando.
            $prevista = $Hoja.Cells.Item($base + 6, 3).Value2
            if ($null -ne $prevista -and "$prevista" -ne '') {
                $txt = [timespan]::FromDays([double]$prevista).ToString('hh\:mm')
                $script:Resumen.Add("      . Salida  : aun sin fichar, se respeta lo que haya ($txt)")
            } else {
                $script:Resumen.Add("      . Salida  : aun sin fichar, y no hay ninguna prevista")
            }
        }

        # Pausas: los pares (salida, entrada) que quedan entre la entrada del dia y, si la
        # jornada esta cerrada, la salida final. Con jornada abierta la ultima marca cierra
        # la ultima pausa.
        $limite = if ($jornadaCerrada) { $marcas.Count - 2 } else { $marcas.Count - 1 }
        $tiposUsados = @()

        for ($k = 1; $k -lt $limite; $k += 2) {
            # Es la duracion, no la hora, la que dice si fue el cafe o la comida.
            $duracion = ((ConvertTo-FraccionDia $marcas[$k + 1]) - (ConvertTo-FraccionDia $marcas[$k])) * 1440

            if ($duracion -le $MaxMinutosDescanso) { $filaSal = $base + 2; $tipo = 'desayuno' }
            else                                   { $filaSal = $base + 4; $tipo = 'comida'   }

            if ($tiposUsados -contains $tipo) {
                Write-Log "$etiquetaDia tiene mas de una pausa de $tipo ($($marcas -join ', ')); solo cabe una en la hoja, revisala." 'AVISO'
                continue
            }
            $tiposUsados += $tipo

            if (Set-CeldaHora -Hoja $Hoja -Fila $filaSal       -Marca $marcas[$k]     -Etiqueta "Salida $tipo"  -SoloSiVacia $soloVacias) { $cambioReal = $true }
            if (Set-CeldaHora -Hoja $Hoja -Fila ($filaSal + 1) -Marca $marcas[$k + 1] -Etiqueta "Entrada $tipo" -SoloSiVacia $soloVacias) { $cambioReal = $true }
        }

        if ($script:Resumen.Count -gt $lineasAntes) {
            $script:Resumen.Insert($lineasAntes, "   $etiquetaDia  [$($marcas -join '  ')]")
        }
        if ($cambioReal) { $tocado = $true }
    }
    return $tocado
}

function Update-RentaSemana {
    <#
        Vuelca la renta acumulada del viernes de la semana anterior (M46, con su signo en L46)
        a "Renta semana" (C5) de la semana en curso. Es lo que hasta ahora se copiaba a mano.

        C5 alimenta el objetivo del lunes (C10 = C4 - C5) y, en cadena, el de toda la semana.

        Se abstiene -- y avisa -- en los casos en que el dato no seria fiable:
          * no existe la hoja de la semana inmediatamente anterior (vacaciones, huecos),
          * el viernes de esa semana no tiene salida fichada: la renta aun no es firme,
          * la renta no esta calculada (M46 devuelve "?").
    #>
    param($Workbook, [datetime]$Lunes)

    $hojaActual = Find-HojaSemana -Workbook $Workbook -Lunes $Lunes
    if (-not $hojaActual) { return $false }

    $hojaPrevia = Find-HojaSemana -Workbook $Workbook -Lunes $Lunes.AddDays(-7)
    if (-not $hojaPrevia) {
        Write-Log "No existe la hoja de la semana anterior al $($Lunes.ToString('dd-MM-yyyy')); 'Renta semana' (C5) se deja como esta." 'AVISO'
        return $false
    }

    # Fila 46 = renta del viernes (base 42 + 4); C48 = su salida.
    $salidaViernes = $hojaPrevia.Range('C48').Value2
    if ($null -eq $salidaViernes -or "$salidaViernes" -eq '') {
        Write-Log "El viernes de '$($hojaPrevia.Name)' no tiene salida; 'Renta semana' se deja como esta." 'AVISO'
        return $false
    }

    $renta = $hojaPrevia.Range('M46').Value2
    if ($renta -isnot [double]) {
        Write-Log "La renta del viernes de '$($hojaPrevia.Name)' no esta calculada (M46='$($hojaPrevia.Range('M46').Text)')." 'AVISO'
        return $false
    }

    $nuevo = [double]$renta
    if ("$($hojaPrevia.Range('L46').Text)".Trim() -eq '-') {
        # El libro nunca ha arrastrado deuda de una semana a otra: se empieza de cero.
        Write-Log "'$($hojaPrevia.Name)' cerro DEBIENDO $($hojaPrevia.Range('M46').Text); se pone 'Renta semana' a 0:00, como venias haciendo. Revisalo si esa semana quieres recuperarla." 'AVISO'
        $nuevo = 0.0
    }

    # Salvaguarda: en C5 NUNCA puede acabar un valor negativo. Con el sistema de fechas 1900
    # Excel no sabe representar tiempos negativos y la celda se veria como "########",
    # arrastrando el error a C10 y a los objetivos de toda la semana.
    if ($nuevo -lt 0) {
        Write-Log "Se ha evitado escribir una 'Renta semana' negativa ($nuevo); se deja en 0:00." 'AVISO'
        $nuevo = 0.0
    }

    $actual = $hojaActual.Range('C5').Value2
    if ($actual -is [double] -and [Math]::Abs([double]$actual - $nuevo) -lt (0.5 / 1440.0)) { return $false }

    $antes   = if ($actual -is [double]) { [timespan]::FromDays([double]$actual).ToString('hh\:mm') } else { '(vacio)' }
    $despues = [timespan]::FromDays($nuevo).ToString('hh\:mm')

    $script:Resumen.Add("Hoja '$($hojaActual.Name)':")
    $script:Resumen.Add("      ~ Renta semana (C5): $antes -> $despues  (del viernes de '$($hojaPrevia.Name)')")
    if (-not $ModoPrueba) { Set-Valor -Celda $hojaActual.Range('C5') -Valor $nuevo }
    return $true
}

# --------------------------------------------------------------- copia de seguridad

function Backup-Libro {
    if ($ModoPrueba) { return }
    if (-not (Test-Path $CarpetaBak)) { New-Item -ItemType Directory -Path $CarpetaBak -Force | Out-Null }

    $destino = Join-Path $CarpetaBak ("{0}_{1}.xlsx" -f `
        [IO.Path]::GetFileNameWithoutExtension($Libro), (Get-Date -Format 'yyyyMMdd-HHmm'))
    Copy-Item -Path $Libro -Destination $destino -Force
    Write-Log "Copia de seguridad: $(Split-Path -Leaf $destino)"

    Get-ChildItem $CarpetaBak -Filter '*.xlsx' |
        Sort-Object LastWriteTime -Descending |
        Select-Object -Skip $MaxBackups |
        ForEach-Object { Remove-Item $_.FullName -Force -Confirm:$false }
}

# ----------------------------------------------------------------------- main

$excel   = $null
$wb      = $null
$yaAbierto = $false

try {
    Write-Log "===== Sync-Horario (SemanasAtras=$SemanasAtras, ModoPrueba=$ModoPrueba) =====" -SoloLog

    if (-not (Test-Path $Libro)) { throw "No existe el libro: $Libro" }

    # --- 1. fichajes -------------------------------------------------------
    $hoy   = (Get-Date).Date
    $lunes = Get-LunesDe $hoy

    $sinDatos = $false

    if ($DesdePortapapeles) {
        $fichajes = Get-FichajesPortapapeles
        # Las semanas a repasar salen de lo que se haya copiado, no del calendario.
        $semanas  = @($fichajes.Keys | ForEach-Object { Get-LunesDe $_ } | Sort-Object -Unique)
    } else {
        $semanas = @()
        for ($i = $SemanasAtras; $i -ge 0; $i--) { $semanas += $lunes.AddDays(-7 * $i) }
        try {
            $fichajes = Get-FichajesIntranet -LunesAConsultar $semanas
        } catch {
            # Sin red no hay fichajes que volcar, pero SI se puede dejar creada la hoja de la
            # semana. Antes se abortaba aqui, y un lunes sin conexion se quedaba sin hoja.
            $fichajes = @{}
            $sinDatos = $true
            Write-Log "No se ha podido leer la intranet ($($_.Exception.Message)). Se continua solo para asegurar la hoja de la semana." 'AVISO'
        }
    }

    # Cuentan tanto los dias con marcas como los que solo traen horas o motivo en "Diario:".
    $conMarcas = ($fichajes.GetEnumerator() | Where-Object {
        $_.Value.Marcas.Count -gt 0 -or $_.Value.Diario -is [double] -or "$($_.Value.Nota)".Trim() -ne ''
    }).Count
    Write-Log "Recibidos de la web $conMarcas dia(s) con datos, en $($semanas.Count) semana(s)." -SoloLog
    if ($conMarcas -eq 0) {
        $sinDatos = $true
        Write-Log 'Sin fichajes que sincronizar.' 'OK'
    }

    # --- 2. Excel: reutilizar sesion viva si la hay -------------------------
    try {
        $excel = [Runtime.InteropServices.Marshal]::GetActiveObject('Excel.Application')
        foreach ($libroAbierto in $excel.Workbooks) {
            if ($libroAbierto.FullName -eq $Libro) { $wb = $libroAbierto; $yaAbierto = $true; break }
        }
        if (-not $yaAbierto) { $excel = $null }
    } catch { $excel = $null }

    if ($yaAbierto) {
        if (-not $wb.Saved) {
            Write-Log 'El libro esta abierto con cambios SIN GUARDAR. No se toca; se reintentara en la proxima pasada.' 'AVISO'
            return
        }
        Write-Log 'El libro ya esta abierto en Excel; se escribe sobre esa sesion.'
    } else {
        $excel = New-Object -ComObject Excel.Application
        $excel.Visible = $false
        $excel.DisplayAlerts = $false
        $wb = $excel.Workbooks.Open($Libro)
    }

    # --- 3. escribir -------------------------------------------------------
    $huboCambios = $false
    # Semanas cuya hoja hay que asegurar aunque no haya datos que volcar: la actual (si no, un
    # lunes sin conexion se quedaria sin hoja) y la que se haya pedido con -CrearSemana.
    $aCrear = @()
    if ($CrearHojaSiFalta) { $aCrear += $lunes }
    if ($CrearSemana -ne [datetime]::MinValue) { $aCrear += (Get-LunesDe $CrearSemana) }

    foreach ($lc in ($aCrear | Sort-Object -Unique)) {
        if ($ModoPrueba) { continue }
        if (Find-HojaSemana -Workbook $wb -Lunes $lc) { continue }
        if (New-HojaSemana -Workbook $wb -Lunes $lc) {
            $huboCambios = $true
            if (Update-RentaSemana -Workbook $wb -Lunes $lc) { $huboCambios = $true }
        }
    }

    foreach ($l in $semanas) {
        if ($sinDatos) { break }
        $hoja = Find-HojaSemana -Workbook $wb -Lunes $l
        if (-not $hoja) {
            if ($CrearHojaSiFalta -and $l -eq $lunes -and -not $ModoPrueba) {
                $hoja = New-HojaSemana -Workbook $wb -Lunes $l
            }
            if (-not $hoja) {
                Write-Log "No hay hoja para la semana del $($l.ToString('dd-MM-yyyy')). Creala y se rellenara sola." 'AVISO'
                continue
            }
        }
        # Aviso (sin tocar nada) si la jornada de una hoja ya existente no corresponde a sus
        # fechas: pasa en la semana del cambio de horario de verano.
        $jEsperada = Get-JornadaOficial -Fecha $l -Config $script:Cfg
        $jHoja     = $hoja.Range('C4').Value2
        if ($jHoja -is [double] -and [Math]::Abs([double]$jHoja - [double]$jEsperada) -ge (0.5/1440.0)) {
            $te = [timespan]::FromDays([double]$jEsperada); $th = [timespan]::FromDays([double]$jHoja)
            Write-Log ("'{0}' tiene jornada {1}:{2:00} y para esas fechas corresponde {3}:{4:00}. No se toca, pero revisalo: descuadra los objetivos de la semana." -f `
                $hoja.Name, [Math]::Floor($th.TotalHours), $th.Minutes, [Math]::Floor($te.TotalHours), $te.Minutes) 'AVISO'
        }

        $script:Resumen.Add("Hoja '$($hoja.Name)':")
        $antes = $script:Resumen.Count
        if (Update-Semana -Hoja $hoja -Lunes $l -Fichajes $fichajes -EsSemanaActual (($l -eq $lunes) -or $Forzar)) {
            $huboCambios = $true
        }
        if ($script:Resumen.Count -eq $antes) { $script:Resumen.RemoveAt($antes - 1) }
    }

    # "Renta semana" (C5), de la mas antigua a la mas reciente: si una incidencia corrige un
    # fichaje viejo, el saldo cambia en cadena y hay que reencadenar todas las semanas
    # posteriores, no solo la actual.
    foreach ($l in ($semanas | Sort-Object)) {
        if (Update-RentaSemana -Workbook $wb -Lunes $l) { $huboCambios = $true }
    }

    # --- 4. guardar --------------------------------------------------------
    if ($huboCambios -and -not $ModoPrueba) {
        # El archivo en disco aun tiene el contenido previo (los cambios estan en memoria),
        # asi que copiarlo justo antes de Save() captura exactamente el estado anterior.
        Backup-Libro
        $wb.Save()
        Write-Log 'Libro guardado.' 'OK'
    } elseif ($ModoPrueba) {
        Write-Log 'MODO PRUEBA: no se ha escrito nada.' 'OK'
    } else {
        Write-Log 'Todo estaba ya al dia; sin cambios.' 'OK'
    }

    if ($script:Resumen.Count -gt 0) {
        Write-Log "Detalle:`r`n`r`n$($script:Resumen -join "`r`n")"
    }
}
catch {
    Write-Log "FALLO: $($_.Exception.Message)" 'ERROR'
    Write-Log "  en: $($_.InvocationInfo.PositionMessage -replace '\r?\n', ' ')" 'ERROR'
}
finally {
    try {
        if ($wb -and -not $yaAbierto) { $wb.Close($false) }
        if ($excel -and -not $yaAbierto) { $excel.Quit() }
    } catch { }
    # Cualquier referencia COM viva mantiene EXCEL.EXE en memoria pese al Quit(), asi que se
    # sueltan todas y se fuerza la recoleccion.
    foreach ($o in $hoja, $wb, $excel) {
        if ($o) { try { [void][Runtime.InteropServices.Marshal]::ReleaseComObject($o) } catch { } }
    }
    $hoja = $null; $wb = $null; $excel = $null
    if (-not $yaAbierto) {
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
        [GC]::Collect(); [GC]::WaitForPendingFinalizers()
    }

    $estado = if ($script:Avisos.Count -gt 0) { "CON AVISOS`r`n - " + ($script:Avisos -join "`r`n - ") } else { 'OK' }
    @(
        "Ultima sincronizacion: $(Get-Date -Format 'dd/MM/yyyy HH:mm')"
        "Estado: $estado"
        ''
        ($script:Resumen -join "`r`n")
    ) -join "`r`n" | Out-File -FilePath $AvisoFile -Encoding utf8
}
