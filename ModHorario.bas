Attribute VB_Name = "ModHorario"
Option Explicit

' ===========================================================================
'  Crear la hoja de una semana nueva
'  ---------------------------------------------------------------------
'  Copia la hoja de la ultima semana y la deja lista:
'    - nombre "dd-mm-aa a dd-mm-aa"
'    - fichajes de los cinco dias VACIOS
'    - etiquetas de dia con la fecha correcta (respeta como estan escritas)
'    - mes de la cabecera por MAYORIA de dias laborables (semanas a caballo)
'    - jornada: 7:00 del 15-jun al 15-sep, 7:43 el resto del anio
'    - "Renta semana" heredada del viernes anterior (nunca negativa)
'
'  Cada dia ocupa 8 filas: base = 10 + 8*d, con d = 0 (lunes) .. 4 (viernes)
'    base+1 Entrada | base+2/3 pausa desayuno | base+4/5 pausa comida | base+6 Salida
' ===========================================================================

Private Const JORNADA_VERANO As String = "7:00"
Private Const JORNADA_RESTO  As String = "7:43"

Private avisoJornada As String

' --- Boton: crea la siguiente semana que falte -----------------------------
'  La logica vive en CrearSemanaMsg, que DEVUELVE el resultado en vez de mostrarlo. Asi se
'  puede probar por automatizacion: un MsgBox dejaria el proceso colgado esperando un clic.
Public Sub CrearSemanaSiguiente()
    MsgBox CrearSemanaMsg(0), vbInformation, "Nueva semana"
End Sub

Public Sub CrearSemanaDe(ByVal lunes As Date)
    MsgBox CrearSemanaMsg(lunes), vbInformation, "Nueva semana"
End Sub

' --- Crea la hoja y devuelve el mensaje de resultado -----------------------
'  Con lunes = 0 crea la siguiente semana a la ultima que tenga hoja.
Public Function CrearSemanaMsg(ByVal lunes As Date) As String
    Dim origen As Worksheet, nueva As Worksheet, ws As Worksheet
    Dim nombreNuevo As String, ultima As Date, f As Date

    If lunes = 0 Then
        ultima = 0
        For Each ws In ThisWorkbook.Worksheets
            f = LunesDeHoja(ws.Name)
            If f > ultima Then ultima = f
        Next ws
        If ultima = 0 Then
            CrearSemanaMsg = "No se ha encontrado ninguna hoja de semana en el libro."
            Exit Function
        End If
        lunes = ultima + 7
    End If

    lunes = LunesDe(lunes)
    nombreNuevo = Format(lunes, "dd-mm-yy") & " a " & Format(lunes + 4, "dd-mm-yy")

    If ExisteHoja(nombreNuevo) Then
        CrearSemanaMsg = "La hoja """ & nombreNuevo & """ ya existe."
        Exit Function
    End If

    Set origen = HojaDeSemana(lunes - 7)
    If origen Is Nothing Then
        CrearSemanaMsg = "No se encuentra la hoja de la semana anterior (" & _
                         Format(lunes - 7, "dd-mm-yyyy") & "), que es la que se copia."
        Exit Function
    End If

    ' Limpiar el aviso: es variable de modulo y, si no, se arrastraria a la siguiente llamada.
    avisoJornada = ""

    Application.ScreenUpdating = False
    origen.Copy After:=origen
    Set nueva = ThisWorkbook.Sheets(origen.Index + 1)
    nueva.Name = nombreNuevo

    PrepararHoja nueva, origen, lunes
    Application.ScreenUpdating = True

    nueva.Activate
    nueva.Range("C11").Select

    CrearSemanaMsg = "Creada la hoja """ & nombreNuevo & """." & vbCrLf & vbCrLf & _
                     "Jornada: " & nueva.Range("C4").Text & _
                     "    Renta de la semana: " & nueva.Range("C5").Text & avisoJornada
End Function

' --- Deja la hoja nueva con sus fechas y sin fichajes ----------------------
Private Sub PrepararHoja(nueva As Worksheet, origen As Worksheet, ByVal lunes As Date)
    Dim d As Integer, base As Integer, off As Integer
    Dim nombreDia As String, fecha As Date
    Dim cuenta As Object, clave As Variant, mejor As Variant, maxN As Long

    Set cuenta = CreateObject("Scripting.Dictionary")

    For d = 0 To 4
        base = 10 + 8 * d
        fecha = lunes + d

        ' Vaciar las seis celdas manuales del dia.
        For off = 1 To 6
            nueva.Cells(base + off, 3).ClearContents
        Next off

        ' Conservar como esta escrito el nombre del dia y recalcular solo el numero.
        nombreDia = Trim(SoloLetras(CStr(nueva.Cells(base, 2).Value)))
        If Len(nombreDia) = 0 Then nombreDia = Format(fecha, "dddd")
        nueva.Cells(base, 2).Value = nombreDia & " " & Format(Day(fecha), "00")

        ' Contar meses de los dias laborables, para la cabecera.
        clave = Format(fecha, "mmmm")
        If cuenta.Exists(clave) Then
            cuenta(clave) = cuenta(clave) + 1
        Else
            cuenta.Add clave, 1
        End If
    Next d

    ' Mes de la cabecera: el de la mayoria de los dias, no el del lunes.
    maxN = 0
    For Each clave In cuenta.Keys
        If cuenta(clave) > maxN Then
            maxN = cuenta(clave)
            mejor = clave
        End If
    Next clave
    If Not IsEmpty(mejor) Then
        If Left(nueva.Range("I1").Formula, 1) <> "=" Then nueva.Range("I1").Value = mejor
    End If

    ' Jornada: la de la MAYORIA de los dias laborables, no la del lunes. La semana del cambio
    ' (el 15-09-2026 cae en martes) tiene dias de los dos regimenes.
    Dim nVerano As Integer, nResto As Integer
    nVerano = 0: nResto = 0
    For d = 0 To 4
        If EsVerano(lunes + d) Then nVerano = nVerano + 1 Else nResto = nResto + 1
    Next d
    If nVerano >= nResto Then
        nueva.Range("C4").Value = CDate(JORNADA_VERANO)
    Else
        nueva.Range("C4").Value = CDate(JORNADA_RESTO)
    End If
    If nVerano > 0 And nResto > 0 Then
        avisoJornada = "  OJO: esta semana cae sobre el cambio de jornada de verano." & vbCrLf & _
                       "  Se ha puesto la de la mayoria de sus dias; revisala."
    End If

    ' "Renta semana": la del viernes de la hoja de origen (fila 46), nunca negativa,
    ' porque Excel no sabe representar tiempos negativos y mostraria "########".
    Dim renta As Double
    renta = 0
    If IsNumeric(origen.Range("M46").Value) Then
        If Trim(CStr(origen.Range("L46").Value)) <> "-" Then renta = origen.Range("M46").Value
    End If
    nueva.Range("C5").Value = renta
End Sub

' --- Auxiliares ------------------------------------------------------------

' Verano: del 15 de junio al 15 de septiembre, ambos incluidos.
Private Function EsVerano(ByVal f As Date) As Boolean
    EsVerano = (f >= DateSerial(Year(f), 6, 15) And f <= DateSerial(Year(f), 9, 15))
End Function

Private Function LunesDe(ByVal f As Date) As Date
    LunesDe = f - ((Weekday(f, vbMonday) - 1))
End Function

' Lee el lunes del nombre de la hoja ("24-08-26 a 28-08-26"). 0 si no se puede.
Private Function LunesDeHoja(ByVal nombre As String) As Date
    Dim p() As String, d As Integer, m As Integer, a As Integer
    On Error GoTo fallo
    p = Split(Trim(nombre), "-")
    If UBound(p) < 2 Then GoTo fallo
    d = CInt(p(0))
    m = CInt(p(1))
    a = CInt(Left(Trim(p(2)), 2))
    If m < 1 Or m > 12 Or d < 1 Or d > 31 Then GoTo fallo
    LunesDeHoja = DateSerial(2000 + a, m, d)
    Exit Function
fallo:
    LunesDeHoja = 0
End Function

Private Function HojaDeSemana(ByVal lunes As Date) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If LunesDeHoja(ws.Name) = lunes Then
            Set HojaDeSemana = ws
            Exit Function
        End If
    Next ws
End Function

Private Function ExisteHoja(ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If StrComp(ws.Name, nombre, vbTextCompare) = 0 Then
            ExisteHoja = True
            Exit Function
        End If
    Next ws
End Function

Private Function SoloLetras(ByVal s As String) As String
    Dim i As Long, c As String, r As String
    For i = 1 To Len(s)
        c = Mid(s, i, 1)
        If Not (c >= "0" And c <= "9") Then r = r & c
    Next i
    SoloLetras = r
End Function
