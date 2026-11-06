Attribute VB_Name = "modSAP_a_TXT_Kiosko_EPP"
Option Explicit

'======================================================================
' MACRO: Convierte un libro SAP en archivo TXT para Kiosko EPP
' Salida: archivo tabulado como epp-kiosko-v05.txt
' Columnas exactas:
' Planta | Alma | Material | Texto breve de Material | Talla | Ubicación | Umb | Precio variable |  Stock 
'
' Uso recomendado:
' 1) Abre TEMPLATE EPP SAP.xlsm
' 2) Importa este módulo
' 3) Ejecuta: Crear_TXT_Kiosko_EPP_Desde_SAP
' 4) Selecciona el libro SAP
' 5) Guarda el TXT generado
'======================================================================

Private Const HOJA_TEMPLATE As String = "EPP"
Private Const FILA_ENCABEZADO_TEMPLATE As Long = 1
Private Const FILA_ENCABEZADO_SAP As Long = 1
Private Const FILA_ENCABEZADO_SAP_ALT As Long = 2
Private Const PRIMERA_FILA_DATOS_SAP As Long = 3

Public Sub Crear_TXT_Kiosko_EPP_Desde_SAP()
    On Error GoTo ErrHandler

    Dim wbBase As Workbook, wbSap As Workbook, wbTmp As Workbook
    Dim wsTpl As Worksheet, wsSap As Worksheet, wsOut As Worksheet
    Dim rutaSAP As Variant, rutaSalida As Variant
    Dim encabezadoSap As Long, primeraFilaDatos As Long

    Dim dictTalla As Object, dictTexto As Object, dictUbicacion As Object, dictUmb As Object
    Dim dictAlma As Object

    Dim colCentro As Long, colAlmacen As Long, colMaterial As Long
    Dim colTexto As Long, colUbicacion As Long, colUMB As Long
    Dim colStock As Long, colPrecio As Long
    Dim faltantes As String

    Dim lastSap As Long, r As Long, outRow As Long, lastOut As Long
    Dim material As String, texto As String, talla As String, ubicacion As String
    Dim umb As String, alma As String, planta As String
    Dim stockValor As Double, precioValor As Double
    Dim contenido As String

    Dim oldCalc As XlCalculation
    oldCalc = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    Set wbBase = ThisWorkbook

    Set dictTalla = CreateObject("Scripting.Dictionary")
    Set dictTexto = CreateObject("Scripting.Dictionary")
    Set dictUbicacion = CreateObject("Scripting.Dictionary")
    Set dictUmb = CreateObject("Scripting.Dictionary")
    Set dictAlma = CreateObject("Scripting.Dictionary")
    dictTalla.CompareMode = vbTextCompare
    dictTexto.CompareMode = vbTextCompare
    dictUbicacion.CompareMode = vbTextCompare
    dictUmb.CompareMode = vbTextCompare
    dictAlma.CompareMode = vbTextCompare

    'Carga el catálogo del template si existe.
    'Esto mantiene tallas, ubicaciones y textos maestros como en tu archivo epp-kiosko.
    On Error Resume Next
    Set wsTpl = wbBase.Worksheets(HOJA_TEMPLATE)
    On Error GoTo ErrHandler

    If Not wsTpl Is Nothing Then
        CargarCatalogoTemplate wsTpl, dictTalla, dictTexto, dictUbicacion, dictUmb, dictAlma
    End If

    rutaSAP = Application.GetOpenFilename( _
        FileFilter:="Archivos Excel (*.xlsx;*.xlsm;*.xlsb;*.xls),*.xlsx;*.xlsm;*.xlsb;*.xls", _
        Title:="Selecciona el libro SAP de inventario EPP")

    If VarType(rutaSAP) = vbBoolean Then GoTo SalidaLimpia

    Set wbSap = Workbooks.Open(CStr(rutaSAP), ReadOnly:=True)
    Set wsSap = ObtenerHojaSAP(wbSap)

    encabezadoSap = DetectarFilaEncabezadoSAP(wsSap)
    primeraFilaDatos = encabezadoSap + 1

    colCentro = BuscarCol(wsSap, encabezadoSap, "Centro")
    colAlmacen = BuscarCol(wsSap, encabezadoSap, "Almacén", "Almacen", "Alma")
    colMaterial = BuscarCol(wsSap, encabezadoSap, "Material")
    colTexto = BuscarCol(wsSap, encabezadoSap, "Texto corto", "Texto breve de Material", "Texto breve de material")
    colUbicacion = BuscarCol(wsSap, encabezadoSap, "Ubicación", "Ubicacion")
    colUMB = BuscarCol(wsSap, encabezadoSap, "Unidad medida base", "Umb", "UMB")
    colStock = BuscarCol(wsSap, encabezadoSap, "Stock total valorado", "Stock")
    colPrecio = BuscarCol(wsSap, encabezadoSap, "Precio variable")

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
    If lastSap < primeraFilaDatos Then
        Err.Raise vbObjectError + 101, , "El libro SAP no contiene filas de datos debajo de los encabezados."
    End If

    Set wbTmp = Workbooks.Add(xlWBATWorksheet)
    Set wsOut = wbTmp.Worksheets(1)
    wsOut.Name = "TXT_KIOSKO"

    wsOut.Range("A1:I1").Value = Array("Planta", "Alma", "Material", "Texto breve de Material", "Talla", "Ubicación", "Umb", "Precio variable", " Stock ")
    outRow = 2

    For r = primeraFilaDatos To lastSap
        material = NormalizarMaterial(wsSap.Cells(r, colMaterial).Value)

        If Len(material) > 0 Then
            planta = ConvertirPlanta(TextoLimpio(wsSap.Cells(r, colCentro).Value))
            alma = TextoLimpio(wsSap.Cells(r, colAlmacen).Value)
            texto = TextoLimpio(wsSap.Cells(r, colTexto).Value)
            ubicacion = TextoLimpio(wsSap.Cells(r, colUbicacion).Value)
            umb = TextoLimpio(wsSap.Cells(r, colUMB).Value)
            precioValor = NumeroSeguro(wsSap.Cells(r, colPrecio).Value)
            stockValor = NumeroSeguro(wsSap.Cells(r, colStock).Value)

            'Si el material ya existe en el template, usar datos maestros del template.
            If dictTexto.Exists(material) Then texto = dictTexto(material)
            If dictTalla.Exists(material) Then
                talla = dictTalla(material)
            Else
                talla = ExtraerTalla(texto)
            End If
            If dictUbicacion.Exists(material) Then ubicacion = dictUbicacion(material)
            If dictUmb.Exists(material) Then umb = dictUmb(material)
            If dictAlma.Exists(material) Then alma = dictAlma(material)

            wsOut.Cells(outRow, "A").Value = planta
            wsOut.Cells(outRow, "B").Value = alma
            wsOut.Cells(outRow, "C").Value = material
            wsOut.Cells(outRow, "D").Value = texto
            wsOut.Cells(outRow, "E").Value = talla
            wsOut.Cells(outRow, "F").Value = ubicacion
            wsOut.Cells(outRow, "G").Value = umb
            wsOut.Cells(outRow, "H").Value = " " & FormatoEntero(precioValor) & " "
            wsOut.Cells(outRow, "I").Value = CLng(Application.WorksheetFunction.Round(stockValor, 0))

            outRow = outRow + 1
        End If
    Next r

    lastOut = outRow - 1

    If lastOut < 2 Then
        Err.Raise vbObjectError + 102, , "No se generaron registros para el archivo TXT."
    End If

    'Orden como tu archivo ejemplo: mayor stock arriba, ceros al final.
    With wsOut.Sort
        .SortFields.Clear
        .SortFields.Add Key:=wsOut.Range("I2:I" & lastOut), _
                        SortOn:=xlSortOnValues, Order:=xlDescending, DataOption:=xlSortNormal
        .SetRange wsOut.Range("A1:I" & lastOut)
        .Header = xlYes
        .Apply
    End With

    contenido = ConstruirTextoTabulado(wsOut, lastOut)

    rutaSalida = Application.GetSaveAsFilename( _
        InitialFileName:=NombreSalidaSugerido(wbSap.Path), _
        FileFilter:="Archivo TXT (*.txt),*.txt", _
        Title:="Guardar archivo TXT para Kiosko EPP")

    If VarType(rutaSalida) = vbBoolean Then GoTo SalidaLimpia

    GuardarUTF8SinBOM contenido, CStr(rutaSalida)

    MsgBox "Archivo TXT generado correctamente:" & vbCrLf & CStr(rutaSalida) & vbCrLf & vbCrLf & _
           "Registros generados: " & Format(lastOut - 1, "#,##0"), _
           vbInformation, "SAP a TXT Kiosko EPP"

SalidaLimpia:
    On Error Resume Next
    If Not wbTmp Is Nothing Then wbTmp.Close SaveChanges:=False
    If Not wbSap Is Nothing Then wbSap.Close SaveChanges:=False
    Application.DisplayAlerts = True
    Application.Calculation = oldCalc
    Application.EnableEvents = True
    Application.ScreenUpdating = True
    Exit Sub

ErrHandler:
    MsgBox "Error al generar el TXT:" & vbCrLf & Err.Description, vbCritical, "SAP a TXT Kiosko EPP"
    Resume SalidaLimpia
End Sub

Private Sub CargarCatalogoTemplate(ByVal ws As Worksheet, _
                                   ByRef dictTalla As Object, _
                                   ByRef dictTexto As Object, _
                                   ByRef dictUbicacion As Object, _
                                   ByRef dictUmb As Object, _
                                   ByRef dictAlma As Object)
    Dim colMaterial As Long, colTexto As Long, colTalla As Long
    Dim colUbicacion As Long, colUmb As Long, colAlma As Long
    Dim lastRow As Long, r As Long
    Dim material As String, valor As String

    colMaterial = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Material")
    colTexto = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Texto breve de Material", "Texto corto")
    colTalla = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Talla")
    colUbicacion = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Ubicación", "Ubicacion")
    colUmb = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Umb", "UMB", "Unidad medida base")
    colAlma = BuscarCol(ws, FILA_ENCABEZADO_TEMPLATE, "Alma", "Almacén", "Almacen")

    If colMaterial = 0 Then Exit Sub

    lastRow = ws.Cells(ws.Rows.Count, colMaterial).End(xlUp).Row

    For r = 2 To lastRow
        material = NormalizarMaterial(ws.Cells(r, colMaterial).Value)
        If Len(material) > 0 Then
            If colTexto > 0 Then
                valor = TextoLimpio(ws.Cells(r, colTexto).Value)
                If Len(valor) > 0 Then dictTexto(material) = valor
            End If

            If colTalla > 0 Then
                valor = TextoLimpio(ws.Cells(r, colTalla).Value)
                If Len(valor) > 0 Then dictTalla(material) = valor
            End If

            If colUbicacion > 0 Then
                valor = TextoLimpio(ws.Cells(r, colUbicacion).Value)
                If Len(valor) > 0 Then dictUbicacion(material) = valor
            End If

            If colUmb > 0 Then
                valor = TextoLimpio(ws.Cells(r, colUmb).Value)
                If Len(valor) > 0 Then dictUmb(material) = valor
            End If

            If colAlma > 0 Then
                valor = TextoLimpio(ws.Cells(r, colAlma).Value)
                If Len(valor) > 0 Then dictAlma(material) = valor
            End If
        End If
    Next r
End Sub

Private Function ConstruirTextoTabulado(ByVal ws As Worksheet, ByVal lastRow As Long) As String
    Dim r As Long, c As Long
    Dim linea(1 To 9) As String
    Dim partes() As String
    Dim texto As String

    ReDim partes(1 To lastRow)

    For r = 1 To lastRow
        For c = 1 To 9
            linea(c) = TextoPlanoTXT(ws.Cells(r, c).Value)
        Next c
        partes(r) = Join(linea, vbTab)
    Next r

    texto = Join(partes, vbLf)
    ConstruirTextoTabulado = texto
End Function

Private Function TextoPlanoTXT(ByVal valor As Variant) As String
    Dim t As String
    t = TextoLimpio(valor)
    t = Replace(t, """", "")
    TextoPlanoTXT = t
End Function

Private Sub GuardarUTF8SinBOM(ByVal texto As String, ByVal ruta As String)
    Dim stText As Object, stBin As Object

    Set stText = CreateObject("ADODB.Stream")
    stText.Type = 2 'adTypeText
    stText.Charset = "utf-8"
    stText.Open
    stText.WriteText texto

    'ADODB escribe BOM UTF-8. Para quitarlo, se cambia a binario y se saltan 3 bytes.
    stText.Position = 0
    stText.Type = 1 'adTypeBinary
    stText.Position = 3

    Set stBin = CreateObject("ADODB.Stream")
    stBin.Type = 1 'adTypeBinary
    stBin.Open

    stText.CopyTo stBin
    stBin.SaveToFile ruta, 2 'adSaveCreateOverWrite

    stText.Close
    stBin.Close
End Sub

Private Function NombreSalidaSugerido(ByVal carpeta As String) As String
    Dim basePath As String
    If Len(carpeta) = 0 Then
        basePath = ThisWorkbook.Path
    Else
        basePath = carpeta
    End If

    If Len(basePath) = 0 Then basePath = Environ$("USERPROFILE") & "\Desktop"

    NombreSalidaSugerido = basePath & "\epp-kiosko-v" & Format(Now, "yyyymmdd-hhnn") & ".txt"
End Function

Private Function DetectarFilaEncabezadoSAP(ByVal ws As Worksheet) As Long
    If BuscarCol(ws, FILA_ENCABEZADO_SAP_ALT, "Material") > 0 Then
        DetectarFilaEncabezadoSAP = FILA_ENCABEZADO_SAP_ALT
    ElseIf BuscarCol(ws, FILA_ENCABEZADO_SAP, "Material") > 0 Then
        DetectarFilaEncabezadoSAP = FILA_ENCABEZADO_SAP
    Else
        DetectarFilaEncabezadoSAP = FILA_ENCABEZADO_SAP_ALT
    End If
End Function

Private Function ObtenerHojaSAP(ByVal wb As Workbook) As Worksheet
    On Error Resume Next
    Set ObtenerHojaSAP = wb.Worksheets("Hoja1")
    On Error GoTo 0

    If ObtenerHojaSAP Is Nothing Then
        Set ObtenerHojaSAP = wb.Worksheets(1)
    End If
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
                ConvertirPlanta = SinAcentos(TextoLimpio(centro))
            End If
    End Select
End Function

Private Function ExtraerTalla(ByVal descripcion As String) As String
    Dim t As String
    Dim re As Object, m As Object
    Dim sep As String, codigo As String

    t = UCase$(TextoLimpio(descripcion))

    Set re = CreateObject("VBScript.RegExp")
    re.Global = False
    re.IgnoreCase = True

    'Ejemplos: T-M, T-XL, T-XXL, T7, T23, T-10
    re.Pattern = "\bT\s*(-?)\s*(XXL|XL|XG|EG|CH|M|G|[0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        sep = CStr(m.SubMatches(0))
        codigo = UCase$(CStr(m.SubMatches(1)))

        If IsNumeric(codigo) Then
            If sep = "-" Then
                ExtraerTalla = "T-" & codigo
            Else
                ExtraerTalla = "T" & codigo
            End If
        Else
            ExtraerTalla = "T-" & codigo
        End If
        Exit Function
    End If

    'Ejemplos: N 6, N10, N-7
    re.Pattern = "\bN(\s|-)?([0-9]{1,2})\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        sep = CStr(m.SubMatches(0))
        codigo = CStr(m.SubMatches(1))
        If sep = "" Then
            ExtraerTalla = "N" & codigo
        Else
            ExtraerTalla = "N " & codigo
        End If
        Exit Function
    End If

    'Ejemplos sin T explícita: CH, M, G, XL, XXL
    re.Pattern = "\b(XXL|XL|XG|EG|CH|M|G)\b"
    If re.Test(t) Then
        Set m = re.Execute(t)(0)
        codigo = UCase$(CStr(m.SubMatches(0)))
        ExtraerTalla = "T-" & codigo
        Exit Function
    End If

    ExtraerTalla = "N/A"
End Function

Private Function FormatoEntero(ByVal valor As Double) As String
    FormatoEntero = Format$(Application.WorksheetFunction.Round(valor, 0), "#,##0")
End Function

Private Function NumeroSeguro(ByVal valor As Variant) As Double
    Dim t As String
    If IsError(valor) Or IsEmpty(valor) Then
        NumeroSeguro = 0
    ElseIf IsNumeric(valor) Then
        NumeroSeguro = CDbl(valor)
    Else
        t = TextoLimpio(valor)
        t = Replace(t, "$", "")
        t = Replace(t, " ", "")
        t = Replace(t, ",", "")
        If IsNumeric(t) Then NumeroSeguro = CDbl(t) Else NumeroSeguro = 0
    End If
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
    NormalizarTexto = LCase$(SinAcentos(TextoLimpio(valor)))
End Function

Private Function SinAcentos(ByVal texto As String) As String
    Dim t As String
    t = texto
    t = Replace(t, "á", "a")
    t = Replace(t, "é", "e")
    t = Replace(t, "í", "i")
    t = Replace(t, "ó", "o")
    t = Replace(t, "ú", "u")
    t = Replace(t, "Á", "A")
    t = Replace(t, "É", "E")
    t = Replace(t, "Í", "I")
    t = Replace(t, "Ó", "O")
    t = Replace(t, "Ú", "U")
    t = Replace(t, "ü", "u")
    t = Replace(t, "Ü", "U")
    SinAcentos = t
End Function
