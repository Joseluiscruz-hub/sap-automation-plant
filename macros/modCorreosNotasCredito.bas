Attribute VB_Name = "modCorreosNotasCredito"
Option Explicit

' ==========================================================
' CORREOS MASIVOS OUTLOOK - NOTAS DE CREDITO POR DISTRIBUIDOR
' Autor: Jose Luis Cruz Prieto
' Version 3.0
'
' Cada fila de la hoja Correos genera un correo con:
'   Texto 1 + caratula del distribuidor (Resumen) + Texto 2 +
'   detalle de partidas filtrado por distribuidor + Texto 3 + firma.
'
' Seguridad:
'   - Modo REVISAR (abre cada correo), BORRADOR (guarda en Borradores)
'     o ENVIAR (envia). Al terminar, el modo regresa a REVISAR.
'   - Filas con Estado "Enviado" se omiten (para reenviar, borra el Estado).
'   - El distribuidor es obligatorio por fila y debe existir en Resumen.
'   - Si falta un adjunto o algo falla, el correo se descarta (no queda abierto).
'   - Valida formato de correos y, opcionalmente, dominios permitidos.
'   - Bitacora en la hoja LOG_CORREOS (sin direcciones de correo).
'
' Entradas: PrepararConfiguracionCorreos, ValidarFilasCorreos,
'           GenerarOCrearCorreosMasivos.
' Layout: docs/layout-correos-notas-credito.md
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
Private Const COL_DISTRIBUIDOR As String = "AB"   ' obligatorio (encabezado "Distribuidor")

' -- Minimos configurables ---------------------------------
Private Const MIN_PARA As Long = 1
Private Const MIN_CC   As Long = 1

' -- Hojas --------------------------------------------------
Private Const HOJA_CORREOS As String = "Correos"
Private Const HOJA_CONFIG As String = "Configuracion"
Private Const HOJA_RESUMEN As String = "Resumen"
Private Const HOJA_DETALLE As String = "Detalle de partidas"
Private Const HOJA_LOG As String = "LOG_CORREOS"

Private Const ESTADO_ENVIADO As String = "Enviado"
Private Const OL_DISCARD As Long = 1

' -- Colores del correo ------------------------------------
Private Const COLOR_ROJO As String = "#C00000"
Private Const COLOR_AZUL As String = "#061320"
Private Const COLOR_NEGRO As String = "#000000"
Private Const COLOR_VERDE As String = "#66FF66"
Private Const COLOR_BLANCO As String = "#FFFFFF"

' ==========================================================
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
    DefaultTexto2 = "Comunico que su {ENTREGA} entrega de nota de cr" & ChrW(233) & "dito manual del mes ya se encuentra disponible."
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
    texto = Replace(texto, ChrW(209), "N")
    texto = Replace(texto, ChrW(241), "N")
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
    EscribirConfigSiNoExiste "Nota modo", "REVISAR = abre cada correo; BORRADOR = lo guarda en Borradores; ENVIAR = envia. Al terminar, la macro regresa el modo a REVISAR."
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
    EscribirConfigSiNoExiste "Dominios permitidos (opcional, separados por ;)", ""
    EscribirConfigSiNoExiste "Nota variables texto", "Puedes usar {ENTREGA}, {CLIENTE}, {NOMBRE_CLIENTE}, {DISTRIBUIDOR}, {NOMBRE_DISTRIBUIDOR} y {REMITENTE}. Para saltos usa Alt+Enter o \n."

    If HojaExiste(HOJA_CORREOS) Then
        With ThisWorkbook.Worksheets(HOJA_CORREOS)
            If BuscarColDistribuidorEnCorreos(ThisWorkbook.Worksheets(HOJA_CORREOS)) = 0 Then
                .Range(COL_DISTRIBUIDOR & "1").Value = "Distribuidor"
            End If
        End With
    End If

    With ThisWorkbook.Worksheets(HOJA_CONFIG)
        .Columns("A:B").AutoFit
        .Columns("B:B").ColumnWidth = 90
        .Columns("B:B").WrapText = True
    End With

    MsgBox "Configuracion preparada. Captura el distribuidor de cada fila en la columna " & COL_DISTRIBUIDOR & " (obligatorio).", vbInformation
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

Private Function BuscarColDistribuidorEnCorreos(ByVal ws As Worksheet) As Long
    Dim c As Long, lastCol As Long, h As String
    lastCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        h = TextoClave(ws.Cells(1, c).Text)
        If Left$(h, 12) = "DISTRIBUIDOR" And InStr(1, h, "NOMBRE", vbTextCompare) = 0 Then
            BuscarColDistribuidorEnCorreos = c
            Exit Function
        End If
    Next c
End Function

' El distribuidor es OBLIGATORIO por fila. Ya no se usa el ID como posicion en Resumen:
' si cambiaba el orden de Resumen, el correo de un distribuidor podia llevar datos de otro.
Private Function ObtenerDistribuidorParaFilaCorreo(ByVal wsCorreos As Worksheet, ByVal fila As Long, _
                                                   ByRef detalle As String) As String
    Dim col As Long, valor As String
    Dim ruta As String, dist As String, nombre As String, total As String

    col = BuscarColDistribuidorEnCorreos(wsCorreos)
    If col = 0 Then
        detalle = "Falta la columna 'Distribuidor' en la hoja " & HOJA_CORREOS & " (ejecuta PrepararConfiguracionCorreos)."
        Exit Function
    End If

    valor = LimpiarIdDistribuidor(wsCorreos.Cells(fila, col).Value)
    If valor = "" Then
        detalle = "Captura el distribuidor de la fila en la columna " & ColLetra(col) & "."
        Exit Function
    End If

    If Not ObtenerDatosResumen(valor, ruta, dist, nombre, total) Then
        detalle = "El distribuidor " & valor & " no existe en la hoja " & HOJA_RESUMEN & "."
        Exit Function
    End If

    ObtenerDistribuidorParaFilaCorreo = valor
End Function

Private Function TotalDistribuidorEsCero(ByVal idDistribuidor As String) As Boolean
    Dim ws As Worksheet, filaEnc As Long, colDist As Long, colTotal As Long, r As Long, ultimaFila As Long
    Set ws = ThisWorkbook.Worksheets(HOJA_RESUMEN)
    filaEnc = FilaEncabezadoResumen()
    If filaEnc = 0 Then Exit Function
    colDist = ColumnaDistribuidor(ws, filaEnc)
    colTotal = ColumnaTotalNotas(ws, filaEnc)
    If colTotal = 0 Then Exit Function
    ultimaFila = ws.Cells(ws.Rows.Count, colDist).End(xlUp).Row
    For r = filaEnc + 1 To ultimaFila
        If IdsIguales(ws.Cells(r, colDist).Value, idDistribuidor) Then
            If IsNumeric(ws.Cells(r, colTotal).Value) Then TotalDistribuidorEsCero = (CDbl(ws.Cells(r, colTotal).Value) = 0)
            Exit Function
        End If
    Next r
End Function

Private Function ColLetra(ByVal col As Long) As String
    ColLetra = Split(ThisWorkbook.Worksheets(HOJA_CORREOS).Cells(1, col).Address(True, False), "$")(0)
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
' VALIDACION DE CORREOS
' ==========================================================

Private Function CorreoValido(ByVal s As String) As Boolean
    s = Trim$(s)
    If Len(s) < 6 Or InStr(s, " ") > 0 Then Exit Function
    If Len(s) - Len(Replace(s, "@", "")) <> 1 Then Exit Function
    If Not (s Like "?*@?*.?*") Then Exit Function
    If Left$(Split(s, "@")(1), 1) = "." Or Right$(s, 1) = "." Then Exit Function
    CorreoValido = True
End Function

Private Function DominioPermitido(ByVal s As String, ByVal permitidos As String) As Boolean
    Dim arr As Variant, i As Long, dom As String
    If Trim$(permitidos) = "" Then DominioPermitido = True: Exit Function
    dom = LCase$(Split(s, "@")(1))
    arr = Split(NormalizarListaCorreos(permitidos), ";")
    For i = LBound(arr) To UBound(arr)
        If LCase$(Trim$(Replace(arr(i), "@", ""))) = dom Then DominioPermitido = True: Exit Function
    Next i
End Function

' Devuelve "" si todos son validos; si no, la descripcion del problema (sin la direccion completa).
Private Function RevisarListaCorreos(ByVal lista As String, ByVal etiqueta As String) As String
    Dim arr As Variant, i As Long, permitidos As String
    If lista = "" Then Exit Function
    permitidos = LeerConfig("Dominios permitidos (opcional, separados por ;)", "")
    arr = Split(lista, ";")
    For i = LBound(arr) To UBound(arr)
        If Not CorreoValido(CStr(arr(i))) Then
            RevisarListaCorreos = etiqueta & ": correo con formato invalido (posicion " & (i + 1) & ")."
            Exit Function
        End If
        If Not DominioPermitido(CStr(arr(i)), permitidos) Then
            RevisarListaCorreos = etiqueta & ": dominio no permitido (posicion " & (i + 1) & ")."
            Exit Function
        End If
    Next i
End Function

Private Function EsEnviado(ByVal ws As Worksheet, ByVal fila As Long) As Boolean
    EsEnviado = (StrComp(Trim$(CStr(ws.Cells(fila, COL_ESTADO).Value)), ESTADO_ENVIADO, vbTextCompare) = 0)
End Function

' Revisa una fila. Devuelve "" si esta lista; si no, el motivo.
Private Function RevisarFila(ByVal ws As Worksheet, ByVal i As Long, ByRef idDist As String, _
                             ByRef paraLista As String, ByRef ccLista As String, ByRef ccoLista As String) As String
    Dim det As String

    paraLista = UnirCorreos(ws.Range(COL_PARA_INI & i & ":" & COL_PARA_FIN & i), CStr(ws.Cells(i, COL_PARA_EXT).Value))
    ccLista = UnirCorreos(ws.Range(COL_CC_INI & i & ":" & COL_CC_FIN & i), CStr(ws.Cells(i, COL_CC_EXT).Value))
    ccoLista = NormalizarListaCorreos(CStr(ws.Cells(i, COL_CCO).Value))

    idDist = ObtenerDistribuidorParaFilaCorreo(ws, i, det)
    If idDist = "" Then RevisarFila = det: Exit Function

    If ContarCorreos(paraLista) < MIN_PARA Then RevisarFila = "Faltan minimo " & MIN_PARA & " destinatario(s) Para.": Exit Function
    If ContarCorreos(ccLista) < MIN_CC Then RevisarFila = "Faltan minimo " & MIN_CC & " destinatario(s) CC.": Exit Function

    det = RevisarListaCorreos(paraLista, "Para")
    If det = "" Then det = RevisarListaCorreos(ccLista, "CC")
    If det = "" Then det = RevisarListaCorreos(ccoLista, "CCO")
    If det <> "" Then RevisarFila = det: Exit Function

    If EsSi(LeerConfig("Solo distribuidores con total distinto de cero", "SI")) Then
        If TotalDistribuidorEsCero(idDist) Then RevisarFila = "El distribuidor " & idDist & " tiene total 0 en Resumen.": Exit Function
    End If
End Function

Public Sub ValidarFilasCorreos()
    Dim ws As Worksheet
    Dim ultimaFila As Long, i As Long
    Dim paraLista As String, ccLista As String, ccoLista As String
    Dim idDist As String, det As String
    Dim ok As Long, errores As Long, omitidos As Long, yaEnviados As Long
    Dim prevScreen As Boolean

    If Not HojasRequeridas(False) Then Exit Sub

    Set ws = ThisWorkbook.Worksheets(HOJA_CORREOS)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False
    On Error GoTo ErrHandler

    For i = 2 To ultimaFila
        If Trim$(CStr(ws.Cells(i, COL_ID).Value)) = "" Then GoTo SiguienteV

        det = RevisarFila(ws, i, idDist, paraLista, ccLista, ccoLista)
        ws.Cells(i, COL_CNT_PARA).Value = ContarCorreos(paraLista)
        ws.Cells(i, COL_CNT_CC).Value = ContarCorreos(ccLista)
        If idDist <> "" And EsSi(LeerConfig("Usar asunto automatico", "SI")) Then
            ws.Cells(i, COL_ASUNTO).Value = ConstruirAsuntoAutomatico(idDist)
        End If

        If EsEnviado(ws, i) Then
            ws.Cells(i, COL_VALIDACION).Value = "Ya enviado; no se reenvia (borra el Estado para reenviar)"
            yaEnviados = yaEnviados + 1
        ElseIf Not EsSi(ws.Cells(i, COL_ENVIAR).Value) Then
            ws.Cells(i, COL_VALIDACION).Value = "No enviar"
            ws.Cells(i, COL_ESTADO).Value = "Omitido"
            omitidos = omitidos + 1
        ElseIf det <> "" Then
            ws.Cells(i, COL_VALIDACION).Value = det
            ws.Cells(i, COL_ESTADO).Value = "Error"
            errores = errores + 1
        Else
            ws.Cells(i, COL_VALIDACION).Value = "OK - Distribuidor " & idDist & " (" & ObtenerNombreDistribuidor(idDist) & ")"
            ws.Cells(i, COL_ESTADO).Value = "Pendiente"
            ok = ok + 1
        End If
SiguienteV:
    Next i

    Application.ScreenUpdating = prevScreen
    MsgBox "Validacion terminada." & vbCrLf & vbCrLf & _
           "OK: " & ok & vbCrLf & "Errores: " & errores & vbCrLf & _
           "Omitidos: " & omitidos & vbCrLf & "Ya enviados: " & yaEnviados, vbInformation
    Exit Sub
ErrHandler:
    Application.ScreenUpdating = prevScreen
    MsgBox "Validacion detenida en la fila " & i & ": " & Err.Description, vbCritical
End Sub

' ==========================================================
' ENVIO / GENERACION
' ==========================================================

Public Sub GenerarOCrearCorreosMasivos()
    Dim modo As String, pausaSegundos As Double, usarFirma As Boolean, rutaBase As String
    Dim OutlookApp As Object, Mail As Object, ws As Worksheet
    Dim ultimaFila As Long, i As Long, pendientes As Long
    Dim procesados As Long, erroresCnt As Long, omitidosCnt As Long, yaEnviadosCnt As Long
    Dim paraLista As String, ccLista As String, ccoLista As String
    Dim asunto As String, htmlCorreo As String, idDist As String, det As String, nAdj As Long
    Dim prevScreen As Boolean, errDesc As String

    If Not HojasRequeridas(True) Then Exit Sub

    modo = UCase$(Trim$(LeerConfig("Modo", "REVISAR")))
    pausaSegundos = Val(LeerConfig("Pausa entre envios (segundos)", "2"))
    usarFirma = EsSi(LeerConfig("Usar firma de Outlook", "SI"))
    rutaBase = Trim$(LeerConfig("Ruta base de adjuntos (opcional)", ""))

    If modo <> "REVISAR" And modo <> "BORRADOR" And modo <> "ENVIAR" Then
        MsgBox "Modo invalido: " & modo & ". Usa REVISAR, BORRADOR o ENVIAR.", vbExclamation
        Exit Sub
    End If

    Set ws = ThisWorkbook.Worksheets(HOJA_CORREOS)
    ultimaFila = ws.Cells(ws.Rows.Count, COL_ID).End(xlUp).Row

    For i = 2 To ultimaFila
        If Trim$(CStr(ws.Cells(i, COL_ID).Value)) <> "" Then
            If EsSi(ws.Cells(i, COL_ENVIAR).Value) And Not EsEnviado(ws, i) Then pendientes = pendientes + 1
        End If
    Next i
    If pendientes = 0 Then
        MsgBox "No hay filas pendientes (Enviar? = SI y Estado distinto de Enviado).", vbExclamation
        Exit Sub
    End If

    If modo = "ENVIAR" Then
        If MsgBox("Modo ENVIAR: se enviaran hasta " & pendientes & " correo(s) directamente." & vbCrLf & _
                  "Las filas ya enviadas se omiten. Continuar?", vbQuestion + vbYesNo + vbDefaultButton2, "Confirmar envio") = vbNo Then
            Exit Sub
        End If
    End If

    On Error GoTo ErrorGeneral
    Set OutlookApp = ObtenerOutlook()
    If OutlookApp Is Nothing Then Err.Raise vbObjectError + 1900, , "No se pudo abrir Outlook."

    prevScreen = Application.ScreenUpdating
    Application.ScreenUpdating = False

    For i = 2 To ultimaFila
        On Error GoTo ErrorFila
        Set Mail = Nothing
        idDist = ""
        If Trim$(CStr(ws.Cells(i, COL_ID).Value)) = "" Then GoTo SiguienteFila

        If EsEnviado(ws, i) Then
            yaEnviadosCnt = yaEnviadosCnt + 1
            RegistrarLogCorreo modo, ws, i, "", 0, 0, 0, "Omitido", "Ya enviado anteriormente"
            GoTo SiguienteFila
        End If

        If Not EsSi(ws.Cells(i, COL_ENVIAR).Value) Then
            ws.Cells(i, COL_ESTADO).Value = "Omitido"
            ws.Cells(i, COL_LOG).Value = "Marcado como No enviar."
            omitidosCnt = omitidosCnt + 1
            GoTo SiguienteFila
        End If

        det = RevisarFila(ws, i, idDist, paraLista, ccLista, ccoLista)
        If det <> "" Then Err.Raise vbObjectError + 2000, , det

        If EsSi(LeerConfig("Usar asunto automatico", "SI")) Then
            asunto = ConstruirAsuntoAutomatico(idDist)
        Else
            asunto = Trim$(CStr(ws.Cells(i, COL_ASUNTO).Value))
        End If
        If asunto = "" Then Err.Raise vbObjectError + 2003, , "No se pudo generar asunto."
        ws.Cells(i, COL_ASUNTO).Value = asunto

        htmlCorreo = ConstruirHTMLCorreo(idDist, ws, i)

        Set Mail = OutlookApp.CreateItem(0)
        With Mail
            .To = paraLista
            .CC = ccLista
            .BCC = ccoLista
            .Subject = asunto

            ' Adjuntos ANTES de mostrar: si falta uno, el correo se descarta en ErrorFila
            AgregarAdjuntos Mail, CStr(ws.Cells(i, COL_ADJUNTOS).Value), rutaBase
            nAdj = .Attachments.Count

            If usarFirma Then
                .Display
                .HTMLBody = htmlCorreo & "<br>" & .HTMLBody
            Else
                .HTMLBody = htmlCorreo
            End If

            Select Case modo
                Case "ENVIAR"
                    .Send
                    ws.Cells(i, COL_ESTADO).Value = ESTADO_ENVIADO
                Case "BORRADOR"
                    .Save
                    If usarFirma Then .Close 0   ' olSave: guarda en Borradores y cierra la ventana
                    ws.Cells(i, COL_ESTADO).Value = "Borrador"
                Case Else
                    If Not usarFirma Then .Display
                    ws.Cells(i, COL_ESTADO).Value = "Generado"
            End Select
        End With

        ws.Cells(i, COL_FECHA).Value = Now
        ws.Cells(i, COL_LOG).Value = "OK - Distribuidor " & idDist
        RegistrarLogCorreo modo, ws, i, idDist, ContarCorreos(paraLista), ContarCorreos(ccLista), nAdj, _
                           CStr(ws.Cells(i, COL_ESTADO).Value), "OK"
        procesados = procesados + 1

        If modo = "ENVIAR" And pausaSegundos > 0 Then Application.Wait Now + (pausaSegundos / 86400#)

SiguienteFila:
        Set Mail = Nothing
        On Error GoTo 0
    Next i

    Application.ScreenUpdating = prevScreen
    RestablecerModoRevisar modo
    MsgBox "Proceso terminado en modo " & modo & "." & vbCrLf & vbCrLf & _
           "Procesados: " & procesados & vbCrLf & _
           "Errores: " & erroresCnt & vbCrLf & _
           "Omitidos: " & omitidosCnt & vbCrLf & _
           "Ya enviados (no se reenviaron): " & yaEnviadosCnt & _
           IIf(modo <> "REVISAR", vbCrLf & vbCrLf & "El modo regreso a REVISAR.", ""), _
           IIf(erroresCnt > 0, vbExclamation, vbInformation)
    Exit Sub

ErrorFila:
    errDesc = Err.Description
    DescartarCorreo Mail
    ws.Cells(i, COL_ESTADO).Value = "Error"
    ws.Cells(i, COL_LOG).Value = "Fila " & i & ": " & errDesc & " (correo descartado)"
    RegistrarLogCorreo modo, ws, i, idDist, 0, 0, 0, "Error", errDesc
    erroresCnt = erroresCnt + 1
    Resume SiguienteFila

ErrorGeneral:
    errDesc = Err.Description
    Application.ScreenUpdating = True
    RestablecerModoRevisar modo
    MsgBox "Error critico: " & errDesc, vbCritical
End Sub

' ==========================================================
' APOYO
' ==========================================================

Private Function HojasRequeridas(ByVal conConfig As Boolean) As Boolean
    Dim faltante As String
    If Not HojaExiste(HOJA_CORREOS) Then faltante = HOJA_CORREOS
    If faltante = "" And Not HojaExiste(HOJA_RESUMEN) Then faltante = HOJA_RESUMEN
    If faltante = "" And Not HojaExiste(HOJA_DETALLE) Then faltante = HOJA_DETALLE
    If faltante = "" And conConfig And Not HojaExiste(HOJA_CONFIG) Then faltante = HOJA_CONFIG & " (ejecuta PrepararConfiguracionCorreos)"
    If faltante <> "" Then
        MsgBox "No existe la hoja '" & faltante & "'.", vbExclamation
        Exit Function
    End If
    HojasRequeridas = True
End Function

' Cierra sin guardar un correo a medio armar (olDiscard). Nunca falla.
Private Sub DescartarCorreo(ByVal Mail As Object)
    On Error Resume Next
    If Not Mail Is Nothing Then Mail.Close OL_DISCARD
    On Error GoTo 0
End Sub

Private Sub RestablecerModoRevisar(ByVal modoUsado As String)
    Dim fila As Long
    On Error Resume Next
    fila = BuscarFilaConfig("Modo")
    If fila > 0 Then ThisWorkbook.Worksheets(HOJA_CONFIG).Cells(fila, "B").Value = "REVISAR"
    On Error GoTo 0
End Sub

' Bitacora sin direcciones de correo: solo conteos.
Private Sub RegistrarLogCorreo(ByVal modo As String, ByVal ws As Worksheet, ByVal fila As Long, ByVal idDist As String, _
                               ByVal nPara As Long, ByVal nCC As Long, ByVal nAdj As Long, _
                               ByVal estado As String, ByVal detalle As String)
    Dim lg As Worksheet, nxt As Long

    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets(HOJA_LOG)
    If lg Is Nothing Then
        Set lg = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        lg.Name = HOJA_LOG
    End If
    If Trim$(CStr(lg.Range("A1").Value)) = "" Then
        lg.Range("A1:I1").Value = Array("FECHA_HORA", "MODO", "FILA", "ID", "DISTRIBUIDOR", "N_PARA", "N_CC", "N_ADJUNTOS", "ESTADO / DETALLE")
    End If
    nxt = lg.Cells(lg.Rows.Count, "A").End(xlUp).Row + 1
    lg.Range("A" & nxt & ":I" & nxt).Value = Array(Now, modo, fila, CStr(ws.Cells(fila, COL_ID).Value), idDist, _
                                                  nPara, nCC, nAdj, estado & " - " & Left$(detalle, 250))
    On Error GoTo 0
End Sub
