Attribute VB_Name = "modConvertirSAP_EPP_CargaLimpia"
Option Explicit

'===========================================================
' MACRO: Convierte un libro SAP a archivo plano de carga EPP
' Salida: .xlsx con UNA sola hoja llamada EPP
' Encabezados obligatorios en fila 1:
' Planta, Alma, Material, Texto breve de Material, Talla,
' Ubicación, Umb, Precio variable, Stock
'
' IMPORTANTE:
' Esta macro genera un archivo NUEVO limpio para cargar en la app.
' No subas el .xlsm con portada/dashboard; sube el .xlsx generado.
'===========================================================

Private Const HOJA_TEMPLATE As String = "EPP"
Private Const HOJA_SALIDA As String = "EPP"
Private Const FILA_DATOS_INICIO As Long = 2

Public Sub Convertir_LibroSAP_A_CargaLimpia_EPP()
    On Error GoTo ErrHandler

    Dim wbBase As Workbook, wbSap As Workbook, wbOut As Workbook
    Dim wsSap As Worksheet, wsOut As Worksheet
    Dim rutaSAP As Variant, rutaSalida As String
    Dim filaHeaderSAP As Long
    Dim dictTallas As Object

    Dim colCentro As Long, colAlmacen As Long, colMaterial As Long
    Dim colTexto As Long, colUbicacion As Long, colUMB As Long
    Dim colStock As Long, colPrecio As Long

    Dim lastSap As Long, r As Long, outRow As Long
    Dim material As String, texto As String, talla As String
    Dim faltantes As String
    Dim oldCalc As XlCalculation

    Set wbBase = ThisWorkbook
    Set dictTallas = CrearDiccionarioTallasDesdeTemplate(wbBase)

    rutaSAP = Application.GetOpenFilename( _
        FileFilter:="Archivos Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Selecciona el libro SAP")

    If VarType(rutaSAP) = vbBoolean Then Exit Sub

    oldCalc = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbSap = Workbooks.Open(CStr(rutaSAP), ReadOnly:=True)
    Set wsSap = ObtenerHojaSAP(wbSap)

    filaHeaderSAP = DetectarFilaEncabezadoSAP(wsSap)
    If filaHeaderSAP = 0 Then
        Err.Raise vbObjectError + 100, , "No encontré la fila de encabezados SAP. Debe contener Centro, Almacén, Material, Texto corto, Ubicación, Unidad medida base, Stock total valorado y Precio variable."
    End If

    colCentro = BuscarCol(wsSap, filaHeaderSAP, "Centro")
    colAlmacen = BuscarCol(wsSap, filaHeaderSAP, "Almacén", "Almacen", "Alma")
    colMaterial = BuscarCol(wsSap, filaHeaderSAP, "Material")
    colTexto = BuscarCol(wsSap, filaHeaderSAP, "Texto corto", "Texto breve de Material", "Texto breve de material")
    colUbicacion = BuscarCol(wsSap, filaHeaderSAP, "Ubicación", "Ubicacion")
    colUMB = BuscarCol(wsSap, filaHeaderSAP, "Unidad medida base", "Umb", "UMB")
    colStock = BuscarCol(wsSap, filaHeaderSAP, "Stock total valorado", "Stock")
    colPrecio = BuscarCol(wsSap, filaHeaderSAP, "Precio variable")

    If colCentro = 0 Then faltantes = faltantes & vbCrLf & "- Centro"
    If colAlmacen = 0 Then faltantes = faltantes & vbCrLf & "- Almacén"
    If colMaterial = 0 Then faltantes = faltantes & vbCrLf & "- Material"
    If colTexto = 0 Then faltantes = faltantes & vbCrLf & "- Texto corto"
    If colUbicacion = 0 Then faltantes = faltantes & vbCrLf & "- Ubicación"
    If colUMB = 0 Then faltantes = faltantes & vbCrLf & "- Unidad medida base"
    If colStock = 0 Then faltantes = faltantes & vbCrLf & "- Stock total valorado"
    If colPrecio = 0 Then faltantes = faltantes & vbCrLf & "- Precio variable"

    If Len(faltantes) > 0 Then
        Err.Raise vbObjectError + 101, , "Faltan columnas en el libro SAP:" & faltantes
    End If

    lastSap = wsSap.Cells(wsSap.Rows.Count, colMaterial).End(xlUp).Row
    If lastSap <= filaHeaderSAP Then
        Err.Raise vbObjectError + 102, , "El libro SAP no tiene datos debajo del encabezado."
    End If

    'Crear salida plana: una hoja, encabezado en fila 1, sin portada, sin dashboard.
    Set wbOut = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbOut.Worksheets(1)
    wsOut.Name = HOJA_SALIDA

    EscribirEncabezados wsOut
    PrepararFormatoTexto wsOut

    outRow = FILA_DATOS_INICIO

    For r = filaHeaderSAP + 1 To lastSap
        material = NormalizarMaterial(wsSap.Cells(r, colMaterial).Value)

        If Len(material) > 0 Then
            texto = TextoLimpio(wsSap.Cells(r, colTexto).Value)
            talla = ObtenerTallaFinal(material, texto, dictTallas)

            wsOut.Cells(outRow, "A").Value = ConvertirPlanta(TextoLimpio(wsSap.Cells(r, colCentro).Value))
            wsOut.Cells(outRow, "B").Value = TextoLimpio(wsSap.Cells(r, colAlmacen).Value)
            wsOut.Cells(outRow, "C").Value = material
            wsOut.Cells(outRow, "D").Value = texto
            wsOut.Cells(outRow, "E").Value = talla
            wsOut.Cells(outRow, "F").Value = TextoLimpio(wsSap.Cells(r, colUbicacion).Value)
            wsOut.Cells(outRow, "G").Value = TextoLimpio(wsSap.Cells(r, colUMB).Value)
            wsOut.Cells(outRow, "H").Value = NumeroSeguro(wsSap.Cells(r, colPrecio).Value)
            wsOut.Cells(outRow, "I").Value = NumeroSeguro(wsSap.Cells(r, colStock).Value)

            outRow = outRow + 1
        End If
    Next r

    If outRow = FILA_DATOS_INICIO Then
        Err.Raise vbObjectError + 103, , "No se importó ningún material. Revisa que la columna Material tenga datos."
    End If

    FormatearSalida wsOut, outRow - 1

    rutaSalida = Left$(CStr(rutaSAP), InStrRev(CStr(rutaSAP), "\")) & "EPP_CARGA_LIMPIA_DESDE_SAP.xlsx"
    wbOut.SaveAs Filename:=rutaSalida, FileFormat:=xlOpenXMLWorkbook

    MsgBox "Archivo limpio generado correctamente:" & vbCrLf & rutaSalida & vbCrLf & vbCrLf & _
           "Filas listas para cargar: " & Format(outRow - 2, "#,##0") & vbCrLf & vbCrLf & _
           "Sube este .xlsx generado, no el .xlsm con macro.", vbInformation, "Carga EPP SAP"

SalidaLimpia:
    On Error Resume Next
    If Not wbSap Is Nothing Then wbSap.Close SaveChanges:=False
    If Not wbOut Is Nothing Then wbOut.Close SaveChanges:=True
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error al convertir SAP a carga EPP:" & vbCrLf & Err.Description, vbCritical, "Carga EPP SAP"
    Resume SalidaLimpia
End Sub

Private Sub EscribirEncabezados(ByVal ws As Worksheet)
    ws.Range("A1:I1").Value = Array("Planta", "Alma", "Material", "Texto breve de Material", "Talla", "Ubicación", "Umb", "Precio variable", "Stock")
End Sub

Private Sub PrepararFormatoTexto(ByVal ws As Worksheet)
    ws.Columns("A:G").NumberFormat = "@"
    ws.Columns("H:H").NumberFormat = "#,##0.00"
    ws.Columns("I:I").NumberFormat = "0"
End Sub

Private Sub FormatearSalida(ByVal ws As Worksheet, ByVal lastRow As Long)
    With ws.Range("A1:I1")
        .Font.Bold = True
        .Font.Color = vbWhite
        .Interior.Color = RGB(192, 0, 0)
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Range("A1:I" & lastRow).AutoFilter
    ws.Columns("A:A").ColumnWidth = 16
    ws.Columns("B:B").ColumnWidth = 10
    ws.Columns("C:C").ColumnWidth = 14
    ws.Columns("D:D").ColumnWidth = 42
    ws.Columns("E:E").ColumnWidth = 12
    ws.Columns("F:F").ColumnWidth = 14
    ws.Columns("G:G").ColumnWidth = 8
    ws.Columns("H:H").ColumnWidth = 16
    ws.Columns("I:I").ColumnWidth = 10
    ws.Rows("1:1").RowHeight = 24
    ws.Activate
    ws.Range("A2").Select
    ActiveWindow.FreezePanes = True
End Sub

Private Function ObtenerHojaSAP(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set ObtenerHojaSAP = wb.Worksheets("Hoja1")
    On Error GoTo 0

    If ObtenerHojaSAP Is Nothing Then
        Set ObtenerHojaSAP = wb.Worksheets(1)
    End If
End Function

Private Function DetectarFilaEncabezadoSAP(ByVal ws As Worksheet) As Long
    Dim fila As Long
    For fila = 1 To 10
        If BuscarCol(ws, fila, "Centro") > 0 And _
           BuscarCol(ws, fila, "Material") > 0 And _
           BuscarCol(ws, fila, "Precio variable") > 0 Then
            DetectarFilaEncabezadoSAP = fila
            Exit Function
        End If
    Next fila
    DetectarFilaEncabezadoSAP = 0
End Function

Private Function CrearDiccionarioTallasDesdeTemplate(ByVal wb As Workbook) As Object
    Dim dict As Object
    Dim ws As Worksheet
    Dim lastRow As Long, r As Long
    Dim material As String, talla As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    On Error Resume Next
    Set ws = wb.Worksheets(HOJA_TEMPLATE)
    On Error GoTo 0

    If ws Is Nothing Then
        Set CrearDiccionarioTallasDesdeTemplate = dict
        Exit Function
    End If

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For r = 2 To lastRow
        material = NormalizarMaterial(ws.Cells(r, "C").Value)
        talla = TextoLimpio(ws.Cells(r, "E").Value)

        If UCase$(talla) = "N/A" Or UCase$(talla) = "NA" Or Len(talla) = 0 Then
            talla = ExtraerTalla(TextoLimpio(ws.Cells(r, "D").Value))
        End If

        If Len(material) > 0 And Len(talla) > 0 Then
            dict(material) = talla
        End If
    Next r

    Set CrearDiccionarioTallasDesdeTemplate = dict
End Function

Private Function BuscarCol(ByVal ws As Worksheet, ByVal filaEncabezado As Long, ParamArray nombres()) As Long
    Dim lastCol As Long, c As Long, i As Long
    Dim encabezadoActual As String, encabezadoBuscado As String

    lastCol = ws.Cells(filaEncabezado, ws.Columns.Count).End(xlToLeft).Column

    For c = 1 To lastCol
        encabezadoActual = NormalizarTexto(ws.Cells(filaEncabezado, c).Value)
        For i = LBound(nombres) To UBound(nombres)
            encabezadoBuscado = NormalizarTexto(CStr(nombres(i)))
            If encabezadoActual = encabezadoBuscado Then
                BuscarCol = c
                Exit Function
            End If
        Next i
    Next c

    BuscarCol = 0
End Function

' NOTA: sustituye <CENTRO>, <PLANTA> y <VARIANTE> por los valores reales de tu sistema.
Private Function ConvertirPlanta(ByVal centro As String) As String
    Dim c As String
    c = UCase$(Trim$(centro))

    Select Case c
        Case "<CENTRO>", "<PLANTA>", "<PLANTA>", "<VARIANTE>", "<CENTRO>"
            ConvertirPlanta = "<PLANTA>"
        Case "<PLANTA_2>", "<VARIANTE_2>"
            ConvertirPlanta = "<PLANTA_2>"
        Case Else
            If Len(c) = 0 Then
                ConvertirPlanta = "<PLANTA>"
            Else
                ConvertirPlanta = centro
            End If
    End Select
End Function

Private Function ObtenerTallaFinal(ByVal material As String, ByVal descripcion As String, ByVal dictTallas As Object) As String
    Dim tallaDic As String, tallaTexto As String

    If dictTallas.Exists(material) Then
        tallaDic = TextoLimpio(dictTallas(material))
    End If

    tallaTexto = ExtraerTalla(descripcion)

    If Len(tallaDic) > 0 And UCase$(tallaDic) <> "N/A" And UCase$(tallaDic) <> "NA" Then
        ObtenerTallaFinal = tallaDic
    ElseIf Len(tallaTexto) > 0 Then
        ObtenerTallaFinal = tallaTexto
    Else
        ObtenerTallaFinal = "N/A"
    End If
End Function

Private Function ExtraerTalla(ByVal descripcion As String) As String
    Dim t As String
    Dim re As Object, m As Object
    Dim codigo As String

    t = UCase$(TextoLimpio(descripcion))
    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    'Tallas tipo T-M, T XL, T-XXL, T-7, T8
    re.Pattern = "\bT\s*[-]?\s*(XXXL|XXL|XL|XG|EG|CH|M|G|[0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        codigo = UCase$(m.SubMatches(0))
        If IsNumeric(codigo) Then
            ExtraerTalla = "T-" & codigo
        Else
            ExtraerTalla = "T-" & codigo
        End If
        Exit Function
    End If

    'Números tipo N 6, N-7, N8
    re.Pattern = "\bN\s*[-]?\s*([0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        ExtraerTalla = "N " & m.SubMatches(0)
        Exit Function
    End If

    'Universal / unitalla
    re.Pattern = "\b(UNIV|UNITALLA|UNIVERSAL)\b"
    If re.Test(t) Then
        ExtraerTalla = "UNIV"
        Exit Function
    End If

    ExtraerTalla = "N/A"
End Function

Private Function NormalizarMaterial(ByVal valor As Variant) As String
    If IsError(valor) Or IsEmpty(valor) Then
        NormalizarMaterial = vbNullString
    ElseIf IsNumeric(valor) Then
        NormalizarMaterial = CStr(CLng(CDbl(valor)))
    Else
        NormalizarMaterial = Trim$(CStr(valor))
    End If
End Function

Private Function TextoLimpio(ByVal valor As Variant) As String
    Dim t As String
    If IsError(valor) Or IsEmpty(valor) Then
        TextoLimpio = vbNullString
        Exit Function
    End If

    t = CStr(valor)
    t = Replace(t, vbCr, " ")
    t = Replace(t, vbLf, " ")
    t = Replace(t, vbTab, " ")
    t = Replace(t, Chr(160), " ")

    Do While InStr(t, "  ") > 0
        t = Replace(t, "  ", " ")
    Loop

    TextoLimpio = Trim$(t)
End Function

Private Function NormalizarTexto(ByVal valor As Variant) As String
    Dim t As String

    If IsError(valor) Or IsEmpty(valor) Then
        NormalizarTexto = vbNullString
        Exit Function
    End If

    t = LCase$(TextoLimpio(valor))
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "ü", "u")
    t = Replace(t, "ñ", "n")

    NormalizarTexto = t
End Function

Private Function NumeroSeguro(ByVal valor As Variant) As Double
    If IsError(valor) Or IsEmpty(valor) Or Len(Trim$(CStr(valor))) = 0 Then
        NumeroSeguro = 0
    ElseIf IsNumeric(valor) Then
        NumeroSeguro = CDbl(valor)
    Else
        NumeroSeguro = 0
    End If
End Function
