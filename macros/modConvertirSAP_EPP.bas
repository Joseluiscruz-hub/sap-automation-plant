Attribute VB_Name = "modConvertirSAP_EPP"
Option Explicit

'===========================================================
' MACRO: Convertir libro SAP al template EPP SAP
' Archivo destino: TEMPLATE EPP SAP.xlsm / hoja "EPP"
' Archivo origen: libro SAP con encabezados en fila 2
'===========================================================

Private Const HOJA_TEMPLATE As String = "EPP"
Private Const FILA_ENCABEZADO_TEMPLATE As Long = 1
Private Const FILA_ENCABEZADO_SAP As Long = 2
Private Const PRIMERA_FILA_DATOS_SAP As Long = 3

'IMPORTANTE:
'En tu archivo ejemplo, la columna H trae STOCK y la I trae PRECIO,
'aunque los encabezados dicen H="Precio variable" e I="Stock".
'Así se deja para respetar el comportamiento del template actual.
'Si quieres respetar el encabezado literal, cambia esta constante a False.
Private Const TEMPLATE_STOCK_EN_H As Boolean = True

Public Sub Convertir_LibroSAP_A_TemplateEPP()
    On Error GoTo ErrHandler

    Dim wbTpl As Workbook, wbSap As Workbook
    Dim wsTpl As Worksheet, wsSap As Worksheet
    Dim rutaSAP As Variant
    Dim dictTallas As Object

    Dim colCentro As Long, colAlmacen As Long, colMaterial As Long
    Dim colTexto As Long, colUbicacion As Long, colUMB As Long
    Dim colStock As Long, colPrecio As Long

    Dim lastSap As Long, lastTpl As Long, lastOut As Long
    Dim r As Long, outRow As Long
    Dim material As String, texto As String, talla As String
    Dim centro As String, ubicacion As String, umb As String
    Dim colOrdenStock As String
    Dim faltantes As String

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbTpl = ThisWorkbook
    Set wsTpl = wbTpl.Worksheets(HOJA_TEMPLATE)

    'Guarda catálogo de tallas existente antes de limpiar el template
    Set dictTallas = CrearDiccionarioTallas(wsTpl)

    rutaSAP = Application.GetOpenFilename( _
        FileFilter:="Archivos Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Selecciona el libro SAP a convertir")

    If VarType(rutaSAP) = vbBoolean Then GoTo SalidaLimpia

    Set wbSap = Workbooks.Open(CStr(rutaSAP), ReadOnly:=True)
    Set wsSap = ObtenerHojaSAP(wbSap)

    'Buscar columnas por encabezado en el libro SAP
    colCentro = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Centro")
    colAlmacen = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Almacén", "Almacen", "Alma")
    colMaterial = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Material")
    colTexto = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Texto corto", "Texto breve de Material", "Texto breve de material")
    colUbicacion = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Ubicación", "Ubicacion")
    colUMB = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Unidad medida base", "Umb", "UMB")
    colStock = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Stock total valorado", "Stock")
    colPrecio = BuscarCol(wsSap, FILA_ENCABEZADO_SAP, "Precio variable")

    If colCentro = 0 Then faltantes = faltantes & vbCrLf & "- Centro"
    If colAlmacen = 0 Then faltantes = faltantes & vbCrLf & "- Almacén"
    If colMaterial = 0 Then faltantes = faltantes & vbCrLf & "- Material"
    If colTexto = 0 Then faltantes = faltantes & vbCrLf & "- Texto corto"
    If colUbicacion = 0 Then faltantes = faltantes & vbCrLf & "- Ubicación"
    If colUMB = 0 Then faltantes = faltantes & vbCrLf & "- Unidad medida base"
    If colStock = 0 Then faltantes = faltantes & vbCrLf & "- Stock total valorado"
    If colPrecio = 0 Then faltantes = faltantes & vbCrLf & "- Precio variable"

    If Len(faltantes) > 0 Then
        Err.Raise vbObjectError + 100, , "No se encontraron estas columnas en el libro SAP:" & faltantes
    End If

    lastSap = wsSap.Cells(wsSap.Rows.Count, colMaterial).End(xlUp).Row
    If lastSap < PRIMERA_FILA_DATOS_SAP Then
        Err.Raise vbObjectError + 101, , "El libro SAP no contiene filas de datos debajo de los encabezados."
    End If

    'Limpiar datos actuales del template, conservando encabezados y estilos
    lastTpl = wsTpl.Cells(wsTpl.Rows.Count, "A").End(xlUp).Row
    If lastTpl >= 2 Then wsTpl.Range("A2:M" & lastTpl).ClearContents

    outRow = 2

    For r = PRIMERA_FILA_DATOS_SAP To lastSap
        material = NormalizarMaterial(wsSap.Cells(r, colMaterial).Value)

        If Len(material) > 0 Then
            centro = TextoLimpio(wsSap.Cells(r, colCentro).Value)
            texto = TextoLimpio(wsSap.Cells(r, colTexto).Value)
            ubicacion = TextoLimpio(wsSap.Cells(r, colUbicacion).Value)
            umb = TextoLimpio(wsSap.Cells(r, colUMB).Value)

            If dictTallas.Exists(material) Then
                talla = dictTallas(material)
            Else
                talla = ExtraerTalla(texto)
            End If

            wsTpl.Cells(outRow, "A").Value = ConvertirPlanta(centro)
            wsTpl.Cells(outRow, "B").Value = wsSap.Cells(r, colAlmacen).Value

            If IsNumeric(material) Then
                wsTpl.Cells(outRow, "C").Value = CLng(material)
            Else
                wsTpl.Cells(outRow, "C").Value = material
            End If

            wsTpl.Cells(outRow, "D").Value = texto
            wsTpl.Cells(outRow, "E").Value = talla
            wsTpl.Cells(outRow, "F").Value = ubicacion
            wsTpl.Cells(outRow, "G").Value = umb

            'Respeta tu template actual: H=Stock e I=Precio
            If TEMPLATE_STOCK_EN_H Then
                wsTpl.Cells(outRow, "H").Value = wsSap.Cells(r, colStock).Value
                wsTpl.Cells(outRow, "I").Value = wsSap.Cells(r, colPrecio).Value
            Else
                wsTpl.Cells(outRow, "H").Value = wsSap.Cells(r, colPrecio).Value
                wsTpl.Cells(outRow, "I").Value = wsSap.Cells(r, colStock).Value
            End If

            'Columna auxiliar como en tu template actual
            wsTpl.Cells(outRow, "M").Value = talla

            outRow = outRow + 1
        End If
    Next r

    lastOut = outRow - 1

    If lastOut >= 2 Then
        'Extender formato de la fila 2 a todas las filas importadas
        wsTpl.Range("A2:M2").Copy
        wsTpl.Range("A2:M" & lastOut).PasteSpecial xlPasteFormats
        Application.CutCopyMode = False

        wsTpl.Range("B2:C" & lastOut).NumberFormat = "0"
        wsTpl.Range("H2:I" & lastOut).NumberFormat = "#,##0.00"

        'Ordenar por stock de mayor a menor
        If TEMPLATE_STOCK_EN_H Then
            colOrdenStock = "H"
        Else
            colOrdenStock = "I"
        End If

        With wsTpl.Sort
            .SortFields.Clear
            .SortFields.Add Key:=wsTpl.Range(colOrdenStock & "2:" & colOrdenStock & lastOut), _
                            SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
            .SetRange wsTpl.Range("A1:M" & lastOut)
            .Header = xlYes
            .Apply
        End With

        If wsTpl.AutoFilterMode Then wsTpl.AutoFilterMode = False
        wsTpl.Range("A1:M" & lastOut).AutoFilter

        wsTpl.Columns("A:M").AutoFit
    End If

    MsgBox "Conversión terminada correctamente." & vbCrLf & _
           "Filas importadas: " & Format(lastOut - 1, "#,##0"), vbInformation, "SAP a Template EPP"

SalidaLimpia:
    On Error Resume Next
    If Not wbSap Is Nothing Then wbSap.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error en conversión:" & vbCrLf & Err.Description, vbCritical, "SAP a Template EPP"
    Resume SalidaLimpia
End Sub

Private Function ObtenerHojaSAP(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set ObtenerHojaSAP = wb.Worksheets("Hoja1")
    On Error GoTo 0

    If ObtenerHojaSAP Is Nothing Then
        Set ObtenerHojaSAP = wb.Worksheets(1)
    End If
End Function

Private Function CrearDiccionarioTallas(ByVal ws As Worksheet) As Object
    Dim dict As Object
    Dim lastRow As Long, r As Long
    Dim material As String, talla As String

    Set dict = CreateObject("Scripting.Dictionary")
    dict.CompareMode = vbTextCompare

    lastRow = ws.Cells(ws.Rows.Count, "C").End(xlUp).Row

    For r = 2 To lastRow
        material = NormalizarMaterial(ws.Cells(r, "C").Value)
        talla = TextoLimpio(ws.Cells(r, "E").Value)

        If Len(talla) = 0 Then
            talla = TextoLimpio(ws.Cells(r, "M").Value)
        End If

        If Len(material) > 0 And Len(talla) > 0 Then
            dict(material) = talla
        End If
    Next r

    Set CrearDiccionarioTallas = dict
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
        Case "<CENTRO>", "<PLANTA>", "<PLANTA>", "<CENTRO>"
            ConvertirPlanta = "<PLANTA>"
        Case Else
            If Len(c) = 0 Then
                ConvertirPlanta = "<PLANTA>"
            Else
                ConvertirPlanta = centro
            End If
    End Select
End Function

Private Function ExtraerTalla(ByVal descripcion As String) As String
    Dim t As String
    Dim re As Object, m As Object
    Dim codigo As String

    t = UCase$(TextoLimpio(descripcion))

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    'Ejemplos: T-M, T-XL, T-XXL, T7, T 23
    re.Pattern = "\bT\s*[-]?\s*(XXL|XL|XG|EG|CH|M|G|[0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        codigo = UCase$(m.SubMatches(0))

        If IsNumeric(codigo) Then
            ExtraerTalla = "T" & codigo
        Else
            ExtraerTalla = "T-" & codigo
        End If

        Exit Function
    End If

    'Ejemplos: N 6, N-7, N8
    re.Pattern = "\bN\s*[-]?\s*([0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        ExtraerTalla = "N " & m.SubMatches(0)
        Exit Function
    End If

    'Ejemplos sin T explícita: CH, M, G, XL, XXL
    re.Pattern = "\b(XXL|XL|XG|EG|CH|M|G)\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        codigo = UCase$(m.SubMatches(0))
        ExtraerTalla = "T-" & codigo
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
