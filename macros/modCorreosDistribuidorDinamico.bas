Attribute VB_Name = "modCorreosDistribuidorDinamico"
Option Explicit

' ==========================================================
' CORREOS MASIVOS OUTLOOK - VERSION DINAMICA POR DISTRIBUIDOR
' ==========================================================
' Soluciona:
' 1) Asunto actualizado por distribuidor.
' 2) Caratula y detalle filtrados por el mismo distribuidor.
' 3) Texto configurable para primera/segunda/tercera/cuarta entrega.
' 4) Espacio entre caratula y parrafo, sin nombre fijo de Jose Luis.
' 5) Respeta el mensaje de la fila del correo cuando viene capturado.
' 6) El ID de la hoja Correos puede seguir siendo consecutivo.
'
' Requisitos de hojas de datos:
' - Resumen: contiene RUTA, DISTRIBUIDOR, NOMBRE DISTRIBUIDOR, Total en notas de credito.
' - Detalle de partidas: contiene Cliente y las partidas.
' - Correos: contiene destinatarios y control de envio.
' - Configuracion: contiene textos y modo de envio.
' ==========================================================

' -- Columnas de la hoja Correos ---------------------------
Private Const COL_ID        As String = "A"
Private Const COL_ENVIAR    As String = "B"
Private Const COL_ESTADO    As String = "C"
Private Const COL_PARA_INI  As String = "D"
Private Const COL_PARA_FIN  As String = "F"
Private Const COL_PARA_EXT  As String = "G"
Private Const COL_CC_INI    As String = "H"
Private Const COL_CC_FIN    As String = "Q"
Private Const COL_CC_EXT    As String = "R"
Private Const COL_CCO       As String = "S"
Private Const COL_ASUNTO    As String = "T"
Private Const COL_CUERPO    As String = "U"
Private Const COL_ADJUNTOS  As String = "V"
Private Const COL_CNT_PARA  As String = "W"
Private Const COL_CNT_CC    As String = "X"
Private Const COL_VALIDACION As String = "Y"
Private Const COL_FECHA     As String = "Z"
Private Const COL_LOG       As String = "AA"

' Columna opcional. Si existe un encabezado con DISTRIBUIDOR o CLIENTE en Correos,
' la macro usara ese valor. Si no existe, el ID consecutivo se usa como ordinal
' para tomar el distribuidor 1, 2, 3... desde la hoja Resumen.
Private Const COL_DISTRIBUIDOR_OPCIONAL As String = "AB"

' -- Minimos configurables ---------------------------------
Private Const MIN_PARA As Long = 3
Private Const MIN_CC   As Long = 10

' -- Hojas fuente ------------------------------------------
Private Const HOJA_CORREOS As String = "Correos"
Private Const HOJA_CONFIG As String = "Configuracion"
Private Const HOJA_RESUMEN As String = "Resumen"
Private Const HOJA_DETALLE As String = "Detalle de partidas"

' -- Colores del correo ------------------------------------
Private Const COLOR_ROJO As String = "#C00000"
Private Const COLOR_AZUL As String = "#061320"
Private Const COLOR_NEGRO As String = "#000000"
Private Const COLOR_VERDE As String = "#66FF66"
Private Const COLOR_BLANCO As String = "#FFFFFF"

' ==========================================================
' TEXTOS DEFAULT SEGUROS PARA ACENTOS
' ==========================================================

Private Function DefaultAsuntoBase() As String
    DefaultAsuntoBase = "Notas de cr" & ChrW(233) & "dito Manual"
End Function

Private Function DefaultTexto1() As String
    DefaultTexto1 = "Buen d" & ChrW(237) & "a."
End Function

Private Function DefaultTexto2() As String
    DefaultTexto2 = "Comunico que su {ENTREGA} entrega de nota de cr" & ChrW(233) & "dito manual del mes por CFO ya se encuentra disponible."
End Function

Private Function DefaultTexto3() As String
    DefaultTexto3 = "Cualquier duda o comentario estoy a sus " & ChrW(243) & "rdenes." & vbCrLf & vbCrLf & "Saludos."
End Function

' ==========================================================
' UTILIDADES GENERALES
' ==========================================================

Private Function HojaExiste(ByVal nombre As String) As Boolean
    Dim ws As Worksheet
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    HojaExiste = Not (ws Is Nothing)
End Function

Private Function EsSi(ByVal valor As String) As Boolean
    Dim v As String
    v = UCase$(Trim$(CStr(valor)))
    v = Replace(v, ChrW(205), "I")
    EsSi = (v = "SI" Or v = "S" Or v = "TRUE" Or v = "VERDADERO")
End Function

Private Function TextoClave(ByVal texto As String) As String
    texto = UCase$(Trim$(CStr(texto)))
    texto = Replace(texto, ChrW(193), "A")
    texto = Replace(texto, ChrW(201), "E")
    texto = Replace(texto, ChrW(205), "I")
    texto = Replace(texto, ChrW(211), "O")
    texto = Replace(texto, ChrW(218), "U")
    texto = Replace(texto, ChrW(225), "A")
    texto = Replace(texto, ChrW(233), "E")
    texto = Replace(texto, ChrW(237), "I")
    texto = Replace(texto, ChrW(243), "O")
    texto = Replace(texto, ChrW(250), "U")
    texto = Replace(texto, "Ñ", "N")
    texto = Replace(texto, "ñ", "N")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, "#", "")
    TextoClave = texto
End Function

Private Function LimpiarIdDistribuidor(ByVal valor As Variant) As String
    Dim s As String

    ' IMPORTANTE:
    ' En algunas celdas vacias con formula Excel devuelve 0.
    ' Ese 0 NO es un distribuidor valido; se debe tratar como vacio.
    On Error GoTo SalirVacio

    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then
        LimpiarIdDistribuidor = ""
        Exit Function
    End If

    If IsNumeric(valor) Then
        If CDbl(valor) = 0 Then
            LimpiarIdDistribuidor = ""
            Exit Function
        End If
        s = Format$(CDbl(valor), "0")
    Else
        s = Trim$(CStr(valor))
    End If

    s = Replace(s, Chr(160), "")
    s = Replace(s, " ", "")
    s = Replace(s, ",", "")
    If InStr(1, s, ".", vbTextCompare) > 0 And IsNumeric(s) Then s = Split(s, ".")(0)

    If s = "" Or s = "0" Then
        LimpiarIdDistribuidor = ""
    Else
        LimpiarIdDistribuidor = s
    End If
    Exit Function

SalirVacio:
    LimpiarIdDistribuidor = ""
End Function

Private Function IdsIguales(ByVal a As Variant, ByVal b As Variant) As Boolean
    IdsIguales = (LimpiarIdDistribuidor(a) <> "" And LimpiarIdDistribuidor(a) = LimpiarIdDistribuidor(b))
End Function

Private Function LeerConfig(ByVal clave As String, Optional ByVal defecto As String = "") As String
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long

    If Not HojaExiste(HOJA_CONFIG) Then
        LeerConfig = defecto
        Exit Function
    End If

    Set ws = ThisWorkbook.Worksheets(HOJA_CONFIG)
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To ultimaFila
        If TextoClave(ws.Cells(i, "A").Value) = TextoClave(clave) Then
            LeerConfig = CStr(ws.Cells(i, "B").Value)
            Exit Function
        End If
    Next i

    LeerConfig = defecto
End Function

Private Function LeerConfigTexto(ByVal clave As String, Optional ByVal defecto As String = "") As String
    Dim texto As String
    texto = LeerConfig(clave, defecto)
    texto = Replace(texto, "\n", vbCrLf)
    LeerConfigTexto = texto
End Function

Private Function BuscarFilaConfig(ByVal clave As String) As Long
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long

    If Not HojaExiste(HOJA_CONFIG) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(HOJA_CONFIG)
    ultimaFila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    For i = 1 To ultimaFila
        If TextoClave(ws.Cells(i, "A").Value) = TextoClave(clave) Then
            BuscarFilaConfig = i
            Exit Function
        End If
    Next i
End Function

Private Sub EscribirConfigSiNoExiste(ByVal clave As String, ByVal valor As String)
    Dim ws As Worksheet
    Dim fila As Long

    If Not HojaExiste(HOJA_CONFIG) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(HOJA_CONFIG)
    fila = BuscarFilaConfig(clave)

    If fila = 0 Then
        fila = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row + 1
        ws.Cells(fila, "A").Value = clave
        ws.Cells(fila, "B").Value = valor
    End If
End Sub

Public Sub PrepararConfiguracionCorreos()
    ' Ejecuta esta macro una vez despues de pegar el modulo.
    ' Agrega los campos necesarios sin borrar lo que ya tengas.
    Dim ws As Worksheet

    If Not HojaExiste(HOJA_CONFIG) Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HOJA_CONFIG
        ws.Range("A1").Value = "Configuracion"
    End If

    EscribirConfigSiNoExiste "Modo", "REVISAR"
    EscribirConfigSiNoExiste "Pausa entre envios (segundos)", "2"
    EscribirConfigSiNoExiste "Usar firma de Outlook", "SI"
    EscribirConfigSiNoExiste "Ruta base de adjuntos (opcional)", ""
    EscribirConfigSiNoExiste "Asunto base", DefaultAsuntoBase()
    EscribirConfigSiNoExiste "Usar asunto automatico", "SI"
    EscribirConfigSiNoExiste "Entrega nota de credito", "segunda"
    EscribirConfigSiNoExiste "Texto 1 antes de caratula", DefaultTexto1()
    EscribirConfigSiNoExiste "Texto 2 entre caratula y detalle", DefaultTexto2()
    EscribirConfigSiNoExiste "Texto 3 despues de detalle", DefaultTexto3()
    EscribirConfigSiNoExiste "Nombre remitente cierre", ""
    EscribirConfigSiNoExiste "Filas detalle correo (incluye encabezado)", "0"
    EscribirConfigSiNoExiste "Solo distribuidores con total distinto de cero", "SI"
    EscribirConfigSiNoExiste "Nota variables texto", "Puedes usar {ENTREGA}, {CLIENTE}, {NOMBRE_CLIENTE}, {DISTRIBUIDOR}, {NOMBRE_DISTRIBUIDOR} y {REMITENTE}. Para saltos usa Alt+Enter o \n."

    If HojaExiste(HOJA_CORREOS) Then
        With ThisWorkbook.Worksheets(HOJA_CORREOS)
            If Trim$(CStr(.Range(COL_DISTRIBUIDOR_OPCIONAL & "1").Value)) = "" Then
                .Range(COL_DISTRIBUIDOR_OPCIONAL & "1").Value = "Distribuidor / Cliente opcional"
            End If
        End With
    End If

    With ThisWorkbook.Worksheets(HOJA_CONFIG)
        .Columns("A:B").AutoFit
        .Columns("B:B").ColumnWidth = 90
        .Columns("B:B").WrapText = True
    End With

    MsgBox "Configuracion preparada. Puedes usar el ID consecutivo o capturar el distribuidor en la columna AB.", vbInformation
End Sub

Private Function ObtenerOutlook() As Object
    On Error Resume Next
    Set ObtenerOutlook = GetObject(, "Outlook.Application")
    On Error GoTo 0

    If ObtenerOutlook Is Nothing Then
        Set ObtenerOutlook = CreateObject("Outlook.Application")
    End If
End Function

Private Function NormalizarListaCorreos(ByVal texto As String) As String
    Dim limpio As String
    Dim arr As Variant
    Dim i As Long
    Dim token As String
    Dim key As Variant
    Dim dict As Object
    Dim salida As String

    limpio = Trim$(CStr(texto))
    limpio = Replace(limpio, vbCrLf, ";")
    limpio = Replace(limpio, vbLf, ";")
    limpio = Replace(limpio, " ;", ";")
    limpio = Replace(limpio, "; ", ";")

    Do While InStr(1, limpio, ";;", vbTextCompare) > 0
        limpio = Replace(limpio, ";;", ";")
    Loop

    If Len(limpio) > 0 Then
        If Left$(limpio, 1) = ";" Then limpio = Mid$(limpio, 2)
        If Len(limpio) > 0 And Right$(limpio, 1) = ";" Then limpio = Left$(limpio, Len(limpio) - 1)
    End If

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = 1

    If limpio <> "" Then
        arr = Split(limpio, ";")
        For i = LBound(arr) To UBound(arr)
            token = Trim$(CStr(arr(i)))
            If token <> "" Then
                If Not dict.Exists(LCase$(token)) Then dict.Add LCase$(token), token
            End If
        Next i
    End If

    For Each key In dict.Keys
        If salida <> "" Then salida = salida & ";"
        salida = salida & dict(key)
    Next key

    NormalizarListaCorreos = salida
End Function

Private Function UnirCorreos(ByVal rango As Range, Optional ByVal extra As String = "") As String
    Dim celda As Range
    Dim texto As String

    For Each celda In rango.Cells
        If Trim$(CStr(celda.Value)) <> "" Then
            If texto <> "" Then texto = texto & ";"
            texto = texto & Trim$(CStr(celda.Value))
        End If
    Next celda

    If Trim$(extra) <> "" Then
        If texto <> "" Then texto = texto & ";"
        texto = texto & Trim$(extra)
    End If

    UnirCorreos = NormalizarListaCorreos(texto)
End Function

Private Function ContarCorreos(ByVal lista As String) As Long
    lista = NormalizarListaCorreos(lista)
    If lista = "" Then
        ContarCorreos = 0
    Else
        ContarCorreos = UBound(Split(lista, ";")) - LBound(Split(lista, ";")) + 1
    End If
End Function

Private Sub AgregarAdjuntos(ByVal Mail As Object, ByVal textoAdjuntos As String, ByVal rutaBase As String)
    Dim arr As Variant
    Dim i As Long
    Dim ruta As String
    Dim esAbsoluta As Boolean

    textoAdjuntos = Trim$(textoAdjuntos)
    If textoAdjuntos = "" Then Exit Sub

    If rutaBase <> "" And Right$(rutaBase, 1) <> "\" Then rutaBase = rutaBase & "\"

    arr = Split(textoAdjuntos, "|")

    For i = LBound(arr) To UBound(arr)
        ruta = Trim$(CStr(arr(i)))
        If ruta = "" Then GoTo SiguienteAdj

        esAbsoluta = (InStr(1, ruta, ":\", vbTextCompare) > 0) Or (Left$(ruta, 2) = "\\")
        If Not esAbsoluta And rutaBase <> "" Then ruta = rutaBase & ruta

        If Dir(ruta) <> "" Then
            Mail.Attachments.Add ruta
        Else
            Err.Raise vbObjectError + 1001, "AgregarAdjuntos", "Archivo no encontrado: " & ruta
        End If

SiguienteAdj:
    Next i
End Sub

' ==========================================================
' BUSQUEDA DE COLUMNAS / DATOS
' ==========================================================

Private Function ColumnaRuta(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If h = "RUTA" Then ColumnaRuta = c: Exit Function
    Next c
End Function

Private Function ColumnaDistribuidor(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If InStr(1, h, "DISTRIBUIDOR", vbTextCompare) > 0 And InStr(1, h, "NOMBRE", vbTextCompare) = 0 Then
            ColumnaDistribuidor = c
            Exit Function
        End If
    Next c
End Function

Private Function ColumnaNombreDistribuidor(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If InStr(1, h, "NOMBRE", vbTextCompare) > 0 And InStr(1, h, "DISTRIBUIDOR", vbTextCompare) > 0 Then
            ColumnaNombreDistribuidor = c
            Exit Function
        End If
    Next c
End Function

Private Function ColumnaTotalNotas(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If InStr(1, h, "TOTAL", vbTextCompare) > 0 And _
           (InStr(1, h, "NOTAS", vbTextCompare) > 0 Or InStr(1, h, "CREDITO", vbTextCompare) > 0) Then
            ColumnaTotalNotas = c
            Exit Function
        End If
    Next c
End Function

Private Function ColumnaClienteDetalle(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If h = "CLIENTE" Then ColumnaClienteDetalle = c: Exit Function
    Next c
End Function

Private Function ColumnaReferenciaDetalle(ByVal ws As Worksheet, ByVal filaEnc As Long) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To lastCol
        h = TextoClave(ws.Cells(filaEnc, c).Text)
        If InStr(1, h, "REFERENCIA", vbTextCompare) > 0 Then ColumnaReferenciaDetalle = c: Exit Function
    Next c
End Function

Private Function FilaEncabezadoResumen() As Long
    Dim ws As Worksheet
    Dim r As Long

    If Not HojaExiste(HOJA_RESUMEN) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(HOJA_RESUMEN)

    For r = 1 To 50
        If ColumnaRuta(ws, r) > 0 And _
           ColumnaDistribuidor(ws, r) > 0 And _
           ColumnaNombreDistribuidor(ws, r) > 0 Then
            FilaEncabezadoResumen = r
            Exit Function
        End If
    Next r
End Function

Private Function FilaEncabezadoDetalle() As Long
    Dim ws As Worksheet
    Dim r As Long

    If Not HojaExiste(HOJA_DETALLE) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(HOJA_DETALLE)

    For r = 1 To 30
        If ColumnaClienteDetalle(ws, r) > 0 And ColumnaReferenciaDetalle(ws, r) > 0 Then
            FilaEncabezadoDetalle = r
            Exit Function
        End If
    Next r
End Function

Private Function TextoCeldaSeguro(ByVal celda As Range) As String
    Dim s As String
    s = CStr(celda.Text)

    If Trim$(s) = "" Or InStr(1, s, "###", vbTextCompare) > 0 Then
        If IsError(celda.Value) Then
            s = ""
        ElseIf IsDate(celda.Value) Then
            s = Format$(celda.Value, "dd.mm.yyyy")
        Else
            s = CStr(celda.Value)
        End If
    End If

    TextoCeldaSeguro = s
End Function

Private Function ObtenerDatosResumen(ByVal idDistribuidor As String, _
                                     ByRef ruta As String, _
                                     ByRef distribuidor As String, _
                                     ByRef nombreDistribuidor As String, _
                                     ByRef totalNotas As String) As Boolean
    Dim ws As Worksheet
    Dim filaEnc As Long
    Dim colRuta As Long, colDist As Long, colNombre As Long, colTotal As Long
    Dim ultimaFila As Long
    Dim r As Long

    If Not HojaExiste(HOJA_RESUMEN) Then Exit Function
    Set ws = ThisWorkbook.Worksheets(HOJA_RESUMEN)

    filaEnc = FilaEncabezadoResumen()
    If filaEnc = 0 Then Exit Function

    colRuta = ColumnaRuta(ws, filaEnc)
    colDist = ColumnaDistribuidor(ws, filaEnc)
    colNombre = ColumnaNombreDistribuidor(ws, filaEnc)
    colTotal = ColumnaTotalNotas(ws, filaEnc)

    ultimaFila = ws.Cells(ws.Rows.Count, colDist).End(xlUp).Row

    For r = filaEnc + 1 To ultimaFila
        If IdsIguales(ws.Cells(r, colDist).Value, idDistribuidor) Then
            ruta = TextoCeldaSeguro(ws.Cells(r, colRuta))
            distribuidor = TextoCeldaSeguro(ws.Cells(r, colDist))
            nombreDistribuidor = TextoCeldaSeguro(ws.Cells(r, colNombre))
            If colTotal > 0 Then
                totalNotas = TextoCeldaSeguro(ws.Cells(r, colTotal))
            Else
                totalNotas = ""
            End If
            ObtenerDatosResumen = True
            Exit Function
        End If
    Next r
End Function

Private Function ObtenerNombreDistribuidor(ByVal idDistribuidor As String) As String
    Dim ruta As String, distribuidor As String, nombre As String, total As String
    If ObtenerDatosResumen(idDistribuidor, ruta, distribuidor, nombre, total) Then
        ObtenerNombreDistribuidor = nombre
    End If
End Function

Private Function ObtenerDistribuidorPorOrdinal(ByVal ordinal As Long) As String
    Dim ws As Worksheet
    Dim filaEnc As Long
    Dim colDist As Long, colTotal As Long
    Dim ultimaFila As Long
    Dim r As Long
    Dim contador As Long
    Dim soloConTotal As Boolean
    Dim totalVal As Variant

    If ordinal <= 0 Then Exit Function
    If Not HojaExiste(HOJA_RESUMEN) Then Exit Function

    Set ws = ThisWorkbook.Worksheets(HOJA_RESUMEN)
    filaEnc = FilaEncabezadoResumen()
    If filaEnc = 0 Then Exit Function

    colDist = ColumnaDistribuidor(ws, filaEnc)
    colTotal = ColumnaTotalNotas(ws, filaEnc)
    ultimaFila = ws.Cells(ws.Rows.Count, colDist).End(xlUp).Row
    soloConTotal = EsSi(LeerConfig("Solo distribuidores con total distinto de cero", "SI"))

    For r = filaEnc + 1 To ultimaFila
        If Trim$(CStr(ws.Cells(r, colDist).Value)) <> "" Then
            If soloConTotal And colTotal > 0 Then
                totalVal = ws.Cells(r, colTotal).Value
                If IsNumeric(totalVal) Then
                    If CDbl(totalVal) = 0 Then GoTo SiguienteR
                End If
            End If

            contador = contador + 1
            If contador = ordinal Then
                ObtenerDistribuidorPorOrdinal = LimpiarIdDistribuidor(ws.Cells(r, colDist).Value)
                Exit Function
            End If
        End If
SiguienteR:
    Next r
End Function

Private Function BuscarColDistribuidorEnCorreos(ByVal ws As Worksheet) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = TextoClave(ws.Cells(1, c).Text)
        If h <> "ID" And h <> "" Then
            If (InStr(1, h, "DISTRIBUIDOR", vbTextCompare) > 0 Or InStr(1, h, "CLIENTE", vbTextCompare) > 0) And _
               InStr(1, h, "NOMBRE", vbTextCompare) = 0 Then
                BuscarColDistribuidorEnCorreos = c
                Exit Function
            End If
        End If
    Next c
End Function

Private Function ObtenerDistribuidorParaFilaCorreo(ByVal wsCorreos As Worksheet, ByVal fila As Long) As String
    Dim colOpcional As Long
    Dim valor As String
    Dim ordinal As Long

    ' 1) Si existe una columna opcional Distribuidor/Cliente, se usa primero.
    '    Pero si la celda esta vacia o trae 0 por formula, se ignora.
    colOpcional = BuscarColDistribuidorEnCorreos(wsCorreos)

    If colOpcional > 0 Then
        valor = LimpiarIdDistribuidor(wsCorreos.Cells(fila, colOpcional).Value)
        If valor <> "" And valor <> "0" Then
            ObtenerDistribuidorParaFilaCorreo = valor
            Exit Function
        End If
    End If

    ' 2) Si no hay distribuidor directo, el ID funciona como consecutivo:
    '    ID 1 = primer distribuidor de Resumen, ID 2 = segundo, etc.
    If IsNumeric(wsCorreos.Cells(fila, COL_ID).Value) Then
        ordinal = CLng(wsCorreos.Cells(fila, COL_ID).Value)
        If ordinal > 0 Then
            ObtenerDistribuidorParaFilaCorreo = ObtenerDistribuidorPorOrdinal(ordinal)
        End If
    End If
End Function

' ==========================================================
' HTML
' ==========================================================

Private Function EscaparHTML(ByVal texto As String) As String
    texto = Replace(texto, "&", "&amp;")
    texto = Replace(texto, "<", "&lt;")
    texto = Replace(texto, ">", "&gt;")
    texto = Replace(texto, """", "&quot;")
    EscaparHTML = texto
End Function

Private Function TextoAHTML(ByVal texto As String) As String
    texto = EscaparHTML(texto)
    texto = Replace(texto, vbCrLf, "<br>")
    texto = Replace(texto, vbLf, "<br>")
    TextoAHTML = texto
End Function

Private Function EsColumnaFacturaRelacionada(ByVal encabezado As String) As Boolean
    Dim h As String
    h = TextoClave(encabezado)
    EsColumnaFacturaRelacionada = _
        (InStr(1, h, "FACTURA RELACIONADA", vbTextCompare) > 0) Or _
        (InStr(1, h, "UUID FACTURA", vbTextCompare) > 0)
End Function

Private Function EstiloBaseTabla() As String
    EstiloBaseTabla = "border-collapse:collapse;font-family:Calibri,Arial,sans-serif;font-size:10pt;margin:0 0 16px 0;"
End Function

Private Function ConstruirHTMLCaratula(ByVal idDistribuidor As String) As String
    Dim ruta As String, dist As String, nombre As String, total As String
    Dim headers(1 To 4) As String
    Dim valores(1 To 4) As String
    Dim html As String
    Dim c As Long
    Dim estilo As String

    If Not ObtenerDatosResumen(idDistribuidor, ruta, dist, nombre, total) Then
        Err.Raise vbObjectError + 3101, "ConstruirHTMLCaratula", "No se encontro el distribuidor " & idDistribuidor & " en la hoja Resumen."
    End If

    headers(1) = "RUTA"
    headers(2) = "DISTRIBUIDOR"
    headers(3) = "NOMBRE DISTRIBUIDOR"
    headers(4) = "Total, en notas de cr" & ChrW(233) & "dito"

    valores(1) = ruta
    valores(2) = dist
    valores(3) = nombre
    valores(4) = total

    html = "<table style='" & EstiloBaseTabla() & "'>"
    html = html & "<tr>"
    For c = 1 To 4
        estilo = "border:1px solid #000000;padding:6px 9px;text-align:center;vertical-align:middle;color:#FFFFFF;font-weight:bold;white-space:nowrap;"
        If c <= 2 Then
            estilo = estilo & "background-color:" & COLOR_ROJO & ";"
        Else
            estilo = estilo & "background-color:" & COLOR_AZUL & ";"
        End If
        html = html & "<th style='" & estilo & "'>" & TextoAHTML(headers(c)) & "</th>"
    Next c
    html = html & "</tr><tr>"

    For c = 1 To 4
        estilo = "border:1px solid #000000;padding:6px 9px;vertical-align:middle;background-color:#FFFFFF;color:#000000;white-space:nowrap;"
        If c = 3 Then estilo = estilo & "text-align:left;"
        If c = 4 Then estilo = estilo & "text-align:right;"
        If c <> 3 And c <> 4 Then estilo = estilo & "text-align:center;"
        html = html & "<td style='" & estilo & "'>" & TextoAHTML(valores(c)) & "</td>"
    Next c

    html = html & "</tr></table>"
    ConstruirHTMLCaratula = html
End Function

Private Function ConstruirHTMLDetalle(ByVal idDistribuidor As String, ByVal filasMaxIncluyeEncabezado As Long) As String
    Dim ws As Worksheet
    Dim filaEnc As Long
    Dim colCliente As Long
    Dim startCol As Long, endCol As Long
    Dim lastRow As Long
    Dim r As Long, c As Long
    Dim html As String
    Dim estilo As String
    Dim encabezado As String
    Dim dataCount As Long
    Dim maxDataRows As Long

    If Not HojaExiste(HOJA_DETALLE) Then
        Err.Raise vbObjectError + 3201, "ConstruirHTMLDetalle", "No existe la hoja '" & HOJA_DETALLE & "'."
    End If

    Set ws = ThisWorkbook.Worksheets(HOJA_DETALLE)
    filaEnc = FilaEncabezadoDetalle()
    If filaEnc = 0 Then
        Err.Raise vbObjectError + 3202, "ConstruirHTMLDetalle", "No se encontro el encabezado de detalle. Debe existir una columna Cliente y Referencia."
    End If

    colCliente = ColumnaClienteDetalle(ws, filaEnc)
    startCol = colCliente
    endCol = ws.Cells(filaEnc, ws.Columns.Count).End(xlToLeft).Column
    lastRow = ws.Cells(ws.Rows.Count, colCliente).End(xlUp).Row

    If filasMaxIncluyeEncabezado > 0 Then
        maxDataRows = filasMaxIncluyeEncabezado - 1
        If maxDataRows < 0 Then maxDataRows = 0
    Else
        maxDataRows = 0
    End If

    html = "<table style='" & EstiloBaseTabla() & "'>"

    ' Encabezado
    html = html & "<tr>"
    For c = startCol To endCol
        encabezado = TextoCeldaSeguro(ws.Cells(filaEnc, c))
        estilo = "border:1px solid #000000;padding:6px 9px;text-align:center;vertical-align:middle;color:#FFFFFF;font-weight:bold;white-space:nowrap;"
        If EsColumnaFacturaRelacionada(encabezado) Then
            estilo = estilo & "background-color:" & COLOR_ROJO & ";"
        Else
            estilo = estilo & "background-color:" & COLOR_NEGRO & ";"
        End If
        html = html & "<th style='" & estilo & "'>" & TextoAHTML(encabezado) & "</th>"
    Next c
    html = html & "</tr>"

    ' Filas filtradas por distribuidor/cliente
    For r = filaEnc + 1 To lastRow
        If IdsIguales(ws.Cells(r, colCliente).Value, idDistribuidor) Then
            If maxDataRows > 0 And dataCount >= maxDataRows Then Exit For
            dataCount = dataCount + 1
            html = html & "<tr>"
            For c = startCol To endCol
                estilo = "border:1px solid #000000;padding:6px 9px;vertical-align:middle;background-color:" & COLOR_VERDE & ";color:#000000;white-space:nowrap;"
                If IsNumeric(ws.Cells(r, c).Value) And Trim$(CStr(ws.Cells(r, c).Value)) <> "" Then
                    estilo = estilo & "text-align:right;"
                Else
                    estilo = estilo & "text-align:left;"
                End If
                html = html & "<td style='" & estilo & "'>" & TextoAHTML(TextoCeldaSeguro(ws.Cells(r, c))) & "</td>"
            Next c
            html = html & "</tr>"
        End If
    Next r

    If dataCount = 0 Then
        html = html & "<tr><td colspan='" & (endCol - startCol + 1) & "' style='border:1px solid #000000;padding:7px;background-color:#FFF2CC;color:#000000;'>Sin partidas encontradas para el distribuidor " & TextoAHTML(idDistribuidor) & ".</td></tr>"
    End If

    html = html & "</table>"
    ConstruirHTMLDetalle = html
End Function

Private Function ReemplazarVariablesTexto(ByVal texto As String, ByVal idDistribuidor As String) As String
    Dim nombre As String
    Dim remitente As String
    Dim entrega As String

    nombre = ObtenerNombreDistribuidor(idDistribuidor)
    remitente = LeerConfigTexto("Nombre remitente cierre", "")
    entrega = LeerConfigTexto("Entrega nota de credito", "segunda")

    texto = Replace(texto, "{ENTREGA}", entrega)
    texto = Replace(texto, "{CLIENTE}", idDistribuidor)
    texto = Replace(texto, "{DISTRIBUIDOR}", idDistribuidor)
    texto = Replace(texto, "{NOMBRE_CLIENTE}", nombre)
    texto = Replace(texto, "{NOMBRE_DISTRIBUIDOR}", nombre)
    texto = Replace(texto, "{REMITENTE}", remitente)

    ReemplazarVariablesTexto = texto
End Function

Private Function ConstruirAsuntoAutomatico(ByVal idDistribuidor As String) As String
    Dim asuntoBase As String
    Dim nombre As String

    asuntoBase = Trim$(LeerConfig("Asunto base", DefaultAsuntoBase()))
    nombre = ObtenerNombreDistribuidor(idDistribuidor)

    If nombre <> "" Then
        ConstruirAsuntoAutomatico = asuntoBase & "; " & nombre
    Else
        ConstruirAsuntoAutomatico = asuntoBase & "; " & idDistribuidor
    End If
End Function

Private Function ConstruirHTMLCorreo(ByVal idDistribuidor As String, ByVal wsCorreos As Worksheet, ByVal filaCorreo As Long) As String
    Dim html As String
    Dim filasDetalle As Long
    Dim texto1 As String
    Dim texto2 As String
    Dim texto3 As String

    filasDetalle = CLng(Val(LeerConfig("Filas detalle correo (incluye encabezado)", "0")))

    texto1 = ReemplazarVariablesTexto(LeerConfigTexto("Texto 1 antes de caratula", DefaultTexto1()), idDistribuidor)

    ' Importante: si en Correos columna U hay mensaje, se respeta ese mensaje por fila.
    texto2 = Trim$(CStr(wsCorreos.Cells(filaCorreo, COL_CUERPO).Value))
    If texto2 = "" Then texto2 = LeerConfigTexto("Texto 2 entre caratula y detalle", DefaultTexto2())
    texto2 = ReemplazarVariablesTexto(texto2, idDistribuidor)

    texto3 = ReemplazarVariablesTexto(LeerConfigTexto("Texto 3 despues de detalle", DefaultTexto3()), idDistribuidor)

    html = "<html><body style='font-family:Calibri,Arial,sans-serif;font-size:11pt;color:#000000;margin:0;padding:0;'>"

    If Trim$(texto1) <> "" Then
        html = html & "<div style='margin-bottom:14px;'>" & TextoAHTML(texto1) & "</div>"
    End If

    html = html & ConstruirHTMLCaratula(idDistribuidor)

    If Trim$(texto2) <> "" Then
        html = html & "<div style='margin:24px 0 18px 0;'>" & TextoAHTML(texto2) & "</div>"
    End If

    html = html & ConstruirHTMLDetalle(idDistribuidor, filasDetalle)

    If Trim$(texto3) <> "" Then
        html = html & "<div style='margin-top:24px;'>" & TextoAHTML(texto3) & "</div>"
    End If

    html = html & "</body></html>"
    ConstruirHTMLCorreo = html
End Function

' ==========================================================
' VALIDACION
' ==========================================================

Public Sub ValidarFilasCorreos()
    If Not HojaExiste(HOJA_CORREOS) Then
        MsgBox "No existe la hoja '" & HOJA_CORREOS & "'.", vbExclamation
        Exit Sub
    End If
    If Not HojaExiste(HOJA_RESUMEN) Then
        MsgBox "No existe la hoja '" & HOJA_RESUMEN & "'.", vbExclamation
        Exit Sub
    End If
    If Not HojaExiste(HOJA_DETALLE) Then
        MsgBox "No existe la hoja '" & HOJA_DETALLE & "'.", vbExclamation
        Exit Sub
    End If

    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim paraLista As String, ccLista As String
    Dim cntPara As Long, cntCC As Long
    Dim idDist As String, asunto As String
    Dim ok As Long, errores As Long, omitidos As Long

    Set ws = ThisWorkbook.Worksheets(HOJA_CORREOS)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To ultimaFila
        If Trim$(CStr(ws.Cells(i, COL_ID).Value)) = "" Then GoTo SiguienteV

        paraLista = UnirCorreos(ws.Range(COL_PARA_INI & i & ":" & COL_PARA_FIN & i), ws.Cells(i, COL_PARA_EXT).Value)
        ccLista = UnirCorreos(ws.Range(COL_CC_INI & i & ":" & COL_CC_FIN & i), ws.Cells(i, COL_CC_EXT).Value)
        cntPara = ContarCorreos(paraLista)
        cntCC = ContarCorreos(ccLista)
        ws.Cells(i, COL_CNT_PARA).Value = cntPara
        ws.Cells(i, COL_CNT_CC).Value = cntCC

        idDist = ObtenerDistribuidorParaFilaCorreo(ws, i)
        If idDist <> "" Then
            asunto = ConstruirAsuntoAutomatico(idDist)
            ws.Cells(i, COL_ASUNTO).Value = asunto
        Else
            asunto = ""
        End If

        If Not EsSi(ws.Cells(i, COL_ENVIAR).Value) Then
            ws.Cells(i, COL_VALIDACION).Value = "No enviar"
            ws.Cells(i, COL_ESTADO).Value = "Omitido"
            omitidos = omitidos + 1
        ElseIf idDist = "" Then
            ws.Cells(i, COL_VALIDACION).Value = "No se encontro distribuidor para este ID"
            ws.Cells(i, COL_ESTADO).Value = "Error"
            errores = errores + 1
        ElseIf cntPara < MIN_PARA Then
            ws.Cells(i, COL_VALIDACION).Value = "Faltan minimo " & MIN_PARA & " Para"
            ws.Cells(i, COL_ESTADO).Value = "Error"
            errores = errores + 1
        ElseIf cntCC < MIN_CC Then
            ws.Cells(i, COL_VALIDACION).Value = "Faltan minimo " & MIN_CC & " CC"
            ws.Cells(i, COL_ESTADO).Value = "Error"
            errores = errores + 1
        ElseIf asunto = "" Then
            ws.Cells(i, COL_VALIDACION).Value = "No se pudo generar asunto"
            ws.Cells(i, COL_ESTADO).Value = "Error"
            errores = errores + 1
        Else
            ws.Cells(i, COL_VALIDACION).Value = "OK - Distribuidor " & idDist
            If ws.Cells(i, COL_ESTADO).Value = "" Or ws.Cells(i, COL_ESTADO).Value = "Error" Then ws.Cells(i, COL_ESTADO).Value = "Pendiente"
            ok = ok + 1
        End If

SiguienteV:
    Next i

    Application.ScreenUpdating = True

    MsgBox "Validacion terminada." & vbCrLf & vbCrLf & _
           "OK: " & ok & " filas" & vbCrLf & _
           "Errores: " & errores & " filas" & vbCrLf & _
           "Omitidos: " & omitidos & " filas", vbInformation
End Sub

' ==========================================================
' ENVIO / GENERACION
' ==========================================================

Public Sub GenerarOCrearCorreosMasivos()
    If Not HojaExiste(HOJA_CORREOS) Then
        MsgBox "No existe la hoja '" & HOJA_CORREOS & "'.", vbExclamation
        Exit Sub
    End If
    If Not HojaExiste(HOJA_CONFIG) Then
        MsgBox "No existe la hoja '" & HOJA_CONFIG & "'. Ejecuta primero PrepararConfiguracionCorreos.", vbExclamation
        Exit Sub
    End If
    If Not HojaExiste(HOJA_RESUMEN) Then
        MsgBox "No existe la hoja '" & HOJA_RESUMEN & "'.", vbExclamation
        Exit Sub
    End If
    If Not HojaExiste(HOJA_DETALLE) Then
        MsgBox "No existe la hoja '" & HOJA_DETALLE & "'.", vbExclamation
        Exit Sub
    End If

    Dim modo As String
    Dim pausaSegundos As Double
    Dim usarFirma As Boolean
    Dim rutaBase As String

    modo = UCase$(Trim$(LeerConfig("Modo", "REVISAR")))
    pausaSegundos = Val(LeerConfig("Pausa entre envios (segundos)", "2"))
    usarFirma = EsSi(LeerConfig("Usar firma de Outlook", "SI"))
    rutaBase = Trim$(LeerConfig("Ruta base de adjuntos (opcional)", ""))

    If modo <> "REVISAR" And modo <> "ENVIAR" Then
        MsgBox "Modo invalido: " & modo & ". Usa REVISAR o ENVIAR.", vbExclamation
        Exit Sub
    End If

    If modo = "ENVIAR" Then
        If MsgBox("Estas en modo ENVIAR. Los correos se enviaran directamente. Deseas continuar?", vbQuestion + vbYesNo, "Confirmar envio") = vbNo Then Exit Sub
    End If

    Dim OutlookApp As Object
    Dim Mail As Object
    Dim ws As Worksheet
    Dim ultimaFila As Long
    Dim i As Long
    Dim enviados As Long, erroresCnt As Long, omitidosCnt As Long

    On Error GoTo ErrorGeneral
    Set OutlookApp = ObtenerOutlook()
    On Error GoTo 0

    Set ws = ThisWorkbook.Worksheets(HOJA_CORREOS)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    Application.ScreenUpdating = False

    For i = 2 To ultimaFila
        On Error GoTo ErrorFila

        If Trim$(CStr(ws.Cells(i, COL_ID).Value)) = "" Then GoTo SiguienteFila

        If Not EsSi(ws.Cells(i, COL_ENVIAR).Value) Then
            ws.Cells(i, COL_ESTADO).Value = "Omitido"
            ws.Cells(i, COL_LOG).Value = "Marcado como No enviar."
            omitidosCnt = omitidosCnt + 1
            GoTo SiguienteFila
        End If

        Dim paraLista As String, ccLista As String, ccoLista As String
        Dim asunto As String, htmlCorreo As String, idDist As String

        paraLista = UnirCorreos(ws.Range(COL_PARA_INI & i & ":" & COL_PARA_FIN & i), ws.Cells(i, COL_PARA_EXT).Value)
        ccLista = UnirCorreos(ws.Range(COL_CC_INI & i & ":" & COL_CC_FIN & i), ws.Cells(i, COL_CC_EXT).Value)
        ccoLista = NormalizarListaCorreos(CStr(ws.Cells(i, COL_CCO).Value))

        idDist = ObtenerDistribuidorParaFilaCorreo(ws, i)
        If idDist = "" Then Err.Raise vbObjectError + 2000, , "No se pudo resolver distribuidor para la fila. Usa ID consecutivo valido o captura distribuidor en columna AB."

        If EsSi(LeerConfig("Usar asunto automatico", "SI")) Then
            asunto = ConstruirAsuntoAutomatico(idDist)
        Else
            asunto = Trim$(CStr(ws.Cells(i, COL_ASUNTO).Value))
        End If
        ws.Cells(i, COL_ASUNTO).Value = asunto

        If ContarCorreos(paraLista) < MIN_PARA Then Err.Raise vbObjectError + 2001, , "Faltan minimo " & MIN_PARA & " destinatarios principales."
        If ContarCorreos(ccLista) < MIN_CC Then Err.Raise vbObjectError + 2002, , "Faltan minimo " & MIN_CC & " destinatarios en CC."
        If asunto = "" Then Err.Raise vbObjectError + 2003, , "No se pudo generar asunto."

        htmlCorreo = ConstruirHTMLCorreo(idDist, ws, i)

        Set Mail = OutlookApp.CreateItem(0)

        With Mail
            .To = paraLista
            .CC = ccLista
            .BCC = ccoLista
            .Subject = asunto

            If usarFirma Then
                .Display
                .HTMLBody = htmlCorreo & "<br>" & .HTMLBody
            Else
                .HTMLBody = htmlCorreo
            End If

            AgregarAdjuntos Mail, CStr(ws.Cells(i, COL_ADJUNTOS).Value), rutaBase

            If modo = "ENVIAR" Then
                .Send
                ws.Cells(i, COL_ESTADO).Value = "Enviado"
            Else
                If Not usarFirma Then .Display
                ws.Cells(i, COL_ESTADO).Value = "Generado"
            End If
        End With

        ws.Cells(i, COL_FECHA).Value = Now
        ws.Cells(i, COL_LOG).Value = "OK - Distribuidor " & idDist
        enviados = enviados + 1

        If modo = "ENVIAR" And pausaSegundos > 0 Then Application.Wait Now + (pausaSegundos / 86400#)

SiguienteFila:
        Set Mail = Nothing
        On Error GoTo 0
    Next i

    Application.ScreenUpdating = True

    MsgBox "Proceso terminado en modo " & modo & "." & vbCrLf & vbCrLf & _
           "Procesados: " & enviados & " correos" & vbCrLf & _
           "Errores: " & erroresCnt & " filas" & vbCrLf & _
           "Omitidos: " & omitidosCnt & " filas", vbInformation
    Exit Sub

ErrorFila:
    ws.Cells(i, COL_ESTADO).Value = "Error"
    ws.Cells(i, COL_LOG).Value = "Fila " & i & ": " & Err.Description
    erroresCnt = erroresCnt + 1
    Set Mail = Nothing
    Err.Clear
    Resume SiguienteFila

ErrorGeneral:
    Application.ScreenUpdating = True
    MsgBox "Error critico al inicializar: " & Err.Description, vbCritical
End Sub
