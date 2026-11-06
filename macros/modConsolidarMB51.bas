Attribute VB_Name = "Consolidar_MB51"
Option Explicit

' ================================================================
' CONSOLIDADOR Y LIMPIADOR DE EXPORTACIONES SAP MB51
' Entrada principal: Consolidar_MB51
'
' Regla de negativos:
'   "ELIMINAR" = omite filas cuya Cantidad sea menor que cero.
'   "ABSOLUTO" = convierte la Cantidad negativa a positiva.
' Recomendado para anulaciones/reversiones MB51: ELIMINAR.
'
' Mejoras v3:
'   - Bug filaEnc en texto UTF-16 (base 0)
'   - Log de errores por archivo
'   - Encabezados duplicados unicos
'   - Omite filas de titulo repetidas de SAP
'   - Ultima fila robusta al borrar columnas vacias
'   - MsgBox segun MODO_NEGATIVOS
'   - Valida la configuracion antes de procesar
'   - Tolera celdas con errores de Excel (#N/A, #VALUE!, etc.)
'   - Reconoce negativos SAP con signo final o parentesis
'   - Hoja Resumen con estadisticas y detalle de errores
'   - Tiempo transcurrido correcto aunque cruce medianoche
' ================================================================

Private Const MODO_NEGATIVOS As String = "ELIMINAR"
Private Const NOMBRE_SALIDA As String = "Consolidado_MB51_Limpio.xlsx"
Private Const MAX_FILAS_EXCEL As Long = 1048576
Private Const CARPETA_PREDETERMINADA As String = _
    "C:\SAP\Descargas_MB51\"

Public Sub Consolidar_MB51()
    Dim carpeta As String, archivo As String, rutaSalida As String
    Dim wbSalida As Workbook, wsSalida As Worksheet
    Dim encabezados As Object, filas As Collection
    Dim calcAnterior As XlCalculation, pantallaAnterior As Boolean
    Dim eventosAnterior As Boolean, alertasAnterior As Boolean
    Dim totalArchivos As Long, totalLeidas As Long, totalNegativas As Long
    Dim totalVacias As Long, totalErrores As Long
    Dim erroresLog As Collection
    Dim textoNegativos As String, detalleErrores As String
    Dim numeroArchivo As Long, descripcionArchivo As String
    Dim mensajeError As String
    Dim i As Long, t As Double

    On Error GoTo FalloGeneral

    ValidarConfiguracion

    carpeta = ElegirCarpeta(CARPETA_PREDETERMINADA)
    If Len(carpeta) = 0 Then Exit Sub
    If Right$(carpeta, 1) <> "\" Then carpeta = carpeta & "\"
    rutaSalida = carpeta & NOMBRE_SALIDA

    Set encabezados = CreateObject("Scripting.Dictionary")
    encabezados.CompareMode = vbTextCompare
    Set filas = New Collection
    Set erroresLog = New Collection

    pantallaAnterior = Application.ScreenUpdating
    eventosAnterior = Application.EnableEvents
    alertasAnterior = Application.DisplayAlerts
    calcAnterior = Application.Calculation
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando consolidacion MB51..."
    t = Timer

    archivo = Dir$(carpeta & "*.*")
    Do While Len(archivo) > 0
        If EsArchivoExcel(archivo) And LCase$(archivo) <> LCase$(NOMBRE_SALIDA) Then
            totalArchivos = totalArchivos + 1
            Application.StatusBar = "Leyendo " & totalArchivos & ": " & archivo
            On Error GoTo ErrorArchivo
            If EsTextoUTF16(carpeta & archivo) Then
                ProcesarTextoSAP carpeta & archivo, archivo, encabezados, filas, _
                                  totalLeidas, totalNegativas, totalVacias
            Else
                ProcesarLibroExcel carpeta & archivo, archivo, encabezados, filas, _
                                   totalLeidas, totalNegativas, totalVacias
            End If
ContinuarArchivo:
            On Error GoTo FalloGeneral
        End If
        archivo = Dir$()
    Loop

    If totalArchivos = 0 Then
        MsgBox "No se encontraron archivos .XLS, .XLSX o .XLSM en:" & vbCrLf & carpeta, vbExclamation
        GoTo SalidaSegura
    End If
    If filas.Count = 0 Then
        detalleErrores = TextoErrores(erroresLog, 8)
        MsgBox "Se revisaron " & totalArchivos & " archivos, pero no quedaron filas validas." & _
               IIf(Len(detalleErrores) > 0, vbCrLf & vbCrLf & detalleErrores, ""), _
               vbExclamation
        GoTo SalidaSegura
    End If

    Set wbSalida = Workbooks.Add(xlWBATWorksheet)
    Set wsSalida = wbSalida.Worksheets(1)
    wsSalida.Name = "Datos_MB51"
    Application.StatusBar = "Creando tabla consolidada..."
    VolcarConsolidado wsSalida, encabezados, filas
    EliminarColumnasVacias wsSalida
    DarFormatoFinal wsSalida

    CrearHojaResumen wbSalida, carpeta, totalArchivos, totalLeidas, _
                      filas.Count, totalNegativas, totalVacias, erroresLog, _
                      SegundosTranscurridos(t)

    If Len(Dir$(rutaSalida)) > 0 Then Kill rutaSalida
    wbSalida.SaveAs Filename:=rutaSalida, FileFormat:=xlOpenXMLWorkbook

    If UCase$(MODO_NEGATIVOS) = "ABSOLUTO" Then
        textoNegativos = "Filas negativas convertidas a absoluto: "
    Else
        textoNegativos = "Filas negativas eliminadas: "
    End If

    detalleErrores = TextoErrores(erroresLog, 8)

    MsgBox "Consolidacion terminada." & vbCrLf & vbCrLf & _
           "Archivos procesados: " & Format$(totalArchivos, "#,##0") & vbCrLf & _
           "Filas leidas: " & Format$(totalLeidas, "#,##0") & vbCrLf & _
           textoNegativos & Format$(totalNegativas, "#,##0") & vbCrLf & _
           "Filas vacias/omitidas: " & Format$(totalVacias, "#,##0") & vbCrLf & _
           "Archivos con error: " & Format$(totalErrores, "#,##0") & vbCrLf & _
           "Filas guardadas: " & Format$(filas.Count, "#,##0") & vbCrLf & _
           "Segundos: " & Format$(SegundosTranscurridos(t), "0.0") & vbCrLf & vbCrLf & _
           "Resultado:" & vbCrLf & rutaSalida & _
           IIf(Len(detalleErrores) > 0, vbCrLf & vbCrLf & detalleErrores, ""), _
           vbInformation, "MB51 consolidado"

SalidaSegura:
    Application.StatusBar = False
    Application.ScreenUpdating = pantallaAnterior
    Application.EnableEvents = eventosAnterior
    Application.DisplayAlerts = alertasAnterior
    Application.Calculation = calcAnterior
    Exit Sub

ErrorArchivo:
    numeroArchivo = Err.Number
    descripcionArchivo = Err.Description
    totalErrores = totalErrores + 1
    On Error Resume Next
    erroresLog.Add archivo & " -> " & numeroArchivo & ": " & descripcionArchivo
    Err.Clear
    On Error GoTo 0
    Resume ContinuarArchivo

FalloGeneral:
    mensajeError = "No se pudo completar la consolidacion." & vbCrLf & _
                   "Error " & Err.Number & ": " & Err.Description
    On Error Resume Next
    If Not wbSalida Is Nothing Then wbSalida.Close SaveChanges:=False
    On Error GoTo 0
    MsgBox mensajeError, vbCritical
    Resume SalidaSegura
End Sub

Private Sub ValidarConfiguracion()
    Dim modo As String
    modo = UCase$(Trim$(MODO_NEGATIVOS))
    If modo <> "ELIMINAR" And modo <> "ABSOLUTO" Then
        Err.Raise vbObjectError + 100, , _
                  "MODO_NEGATIVOS debe ser ELIMINAR o ABSOLUTO. Valor actual: " & MODO_NEGATIVOS
    End If
End Sub

Private Function SegundosTranscurridos(ByVal inicio As Double) As Double
    Dim fin As Double
    fin = Timer
    If fin < inicio Then fin = fin + 86400#
    SegundosTranscurridos = fin - inicio
End Function

' Compatibilidad con el nombre anterior
Public Sub Consolidar_MB51_2025()
    Consolidar_MB51
End Sub

Private Function ElegirCarpeta(ByVal sugerida As String) As String
    Dim dlg As FileDialog
    Set dlg = Application.FileDialog(msoFileDialogFolderPicker)
    With dlg
        .Title = "Selecciona la carpeta con las descargas MB51"
        .AllowMultiSelect = False
        If Len(Dir$(sugerida, vbDirectory)) > 0 Then .InitialFileName = sugerida
        If .Show = -1 Then ElegirCarpeta = .SelectedItems(1)
    End With
End Function

Private Function EsArchivoExcel(ByVal nombre As String) As Boolean
    Dim ext As String
    ext = LCase$(Mid$(nombre, InStrRev(nombre, ".") + 1))
    EsArchivoExcel = (ext = "xls" Or ext = "xlsx" Or ext = "xlsm") And Left$(nombre, 2) <> "~$"
End Function

Private Function EsTextoUTF16(ByVal ruta As String) As Boolean
    Dim n As Integer, b1 As Byte, b2 As Byte
    On Error GoTo NoEs
    n = FreeFile
    Open ruta For Binary Access Read As #n
    If LOF(n) >= 2 Then
        Get #n, 1, b1
        Get #n, 2, b2
        EsTextoUTF16 = (b1 = &HFF And b2 = &HFE) Or (b1 = &HFE And b2 = &HFF)
    End If
    Close #n
    Exit Function
NoEs:
    On Error Resume Next
    Close #n
End Function

Private Sub ProcesarTextoSAP(ByVal ruta As String, ByVal archivoOrigen As String, _
                             ByRef encabezados As Object, ByRef filas As Collection, _
                             ByRef leidas As Long, ByRef negativas As Long, ByRef vacias As Long)
    Dim stm As Object, contenido As String, lineas() As String, campos() As String
    Dim i As Long, filaEnc As Long, encabezadoLocal() As String

    filaEnc = -1

    Set stm = CreateObject("ADODB.Stream")
    With stm
        .Type = 2
        .Charset = "unicode"
        .Open
        .LoadFromFile ruta
        contenido = .ReadText(-1)
        .Close
    End With
    contenido = Replace(contenido, vbCrLf, vbLf)
    contenido = Replace(contenido, vbCr, vbLf)
    lineas = Split(contenido, vbLf)

    For i = LBound(lineas) To UBound(lineas)
        If InStr(1, lineas(i), "Material", vbTextCompare) > 0 And _
           InStr(1, lineas(i), "Cantidad", vbTextCompare) > 0 Then
            filaEnc = i
            campos = Split(lineas(i), vbTab)
            encabezadoLocal = LimpiarEncabezados(campos)
            RegistrarEncabezados encabezados, encabezadoLocal
            Exit For
        End If
    Next i

    If filaEnc < 0 Then Err.Raise vbObjectError + 101, , "No se encontro el encabezado SAP"

    For i = filaEnc + 1 To UBound(lineas)
        If Len(Trim$(lineas(i))) = 0 Then
            vacias = vacias + 1
        Else
            campos = Split(lineas(i), vbTab)
            AgregarFila campos, encabezadoLocal, archivoOrigen, "Exportacion SAP", _
                        encabezados, filas, leidas, negativas, vacias
        End If
    Next i
End Sub

Private Sub ProcesarLibroExcel(ByVal ruta As String, ByVal archivoOrigen As String, _
                               ByRef encabezados As Object, ByRef filas As Collection, _
                               ByRef leidas As Long, ByRef negativas As Long, ByRef vacias As Long)
    Dim wb As Workbook, ws As Worksheet, datos As Variant
    Dim filaEnc As Long, r As Long, c As Long, encabezadoLocal() As String, campos() As String

    On Error GoTo CerrarConError
    Set wb = Workbooks.Open(Filename:=ruta, ReadOnly:=True, UpdateLinks:=0, AddToMru:=False)

    For Each ws In wb.Worksheets
        If WorksheetFunction.CountA(ws.UsedRange) > 0 Then
            datos = ws.UsedRange.Value2
            filaEnc = BuscarEncabezadoMatriz(datos)
            If filaEnc > 0 Then
                ReDim encabezadoLocal(0 To UBound(datos, 2) - 1)
                For c = 1 To UBound(datos, 2)
                    encabezadoLocal(c - 1) = EncabezadoUnico(CStr(datos(filaEnc, c)), c, encabezadoLocal, c - 2)
                Next c
                RegistrarEncabezados encabezados, encabezadoLocal
                For r = filaEnc + 1 To UBound(datos, 1)
                    ReDim campos(0 To UBound(datos, 2) - 1)
                    For c = 1 To UBound(datos, 2)
                        campos(c - 1) = ValorCeldaComoTexto(datos(r, c))
                    Next c
                    AgregarFila campos, encabezadoLocal, archivoOrigen, ws.Name, _
                                encabezados, filas, leidas, negativas, vacias
                Next r
            End If
        End If
    Next ws

    wb.Close SaveChanges:=False
    Exit Sub

CerrarConError:
    Dim numeroError As Long, descripcionError As String
    numeroError = Err.Number
    descripcionError = Err.Description
    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False
    On Error GoTo 0
    Err.Raise numeroError, , descripcionError
End Sub

Private Function ValorCeldaComoTexto(ByVal v As Variant) As String
    If IsError(v) Then
        ValorCeldaComoTexto = vbNullString
    ElseIf IsNull(v) Then
        ValorCeldaComoTexto = vbNullString
    ElseIf IsEmpty(v) Then
        ValorCeldaComoTexto = vbNullString
    Else
        ValorCeldaComoTexto = CStr(v)
    End If
End Function

Private Function BuscarEncabezadoMatriz(ByVal datos As Variant) As Long
    Dim r As Long, c As Long, texto As String, limite As Long
    limite = WorksheetFunction.Min(30, UBound(datos, 1))
    For r = 1 To limite
        texto = vbNullString
        For c = 1 To UBound(datos, 2)
            texto = texto & "|" & ValorCeldaComoTexto(datos(r, c))
        Next c
        If InStr(1, texto, "Material", vbTextCompare) > 0 And _
           InStr(1, texto, "Cantidad", vbTextCompare) > 0 Then
            BuscarEncabezadoMatriz = r
            Exit Function
        End If
    Next r
End Function

Private Function LimpiarEncabezados(ByRef campos() As String) As String()
    Dim salida() As String, i As Long
    ReDim salida(LBound(campos) To UBound(campos))
    For i = LBound(campos) To UBound(campos)
        salida(i) = EncabezadoUnico(campos(i), i + 1, salida, i - 1)
    Next i
    LimpiarEncabezados = salida
End Function

Private Function EncabezadoUnico(ByVal valor As String, ByVal posicion As Long, _
                                 ByRef yaUsados() As String, ByVal hastaIdx As Long) As String
    Dim baseName As String, candidato As String, n As Long, j As Long, existe As Boolean

    baseName = LimpiarTexto(valor)
    If Len(baseName) = 0 Then baseName = "Columna_" & posicion

    candidato = baseName
    n = 1
    Do
        existe = False
        If hastaIdx >= LBound(yaUsados) Then
            For j = LBound(yaUsados) To hastaIdx
                If StrComp(yaUsados(j), candidato, vbTextCompare) = 0 Then
                    existe = True
                    Exit For
                End If
            Next j
        End If
        If Not existe Then Exit Do
        n = n + 1
        candidato = baseName & "_" & n
    Loop

    EncabezadoUnico = candidato
End Function

Private Sub RegistrarEncabezados(ByRef dic As Object, ByRef encabezadoLocal() As String)
    Dim i As Long, nombre As String
    For i = LBound(encabezadoLocal) To UBound(encabezadoLocal)
        nombre = encabezadoLocal(i)
        If Not dic.Exists(nombre) Then dic.Add nombre, dic.Count + 1
    Next i
    If Not dic.Exists("Archivo origen") Then dic.Add "Archivo origen", dic.Count + 1
    If Not dic.Exists("Hoja origen") Then dic.Add "Hoja origen", dic.Count + 1
End Sub

Private Sub AgregarFila(ByRef campos() As String, ByRef encabezadoLocal() As String, _
                        ByVal archivoOrigen As String, ByVal hojaOrigen As String, _
                        ByRef encabezados As Object, ByRef filas As Collection, _
                        ByRef leidas As Long, ByRef negativas As Long, ByRef vacias As Long)
    Dim registro As Object, i As Long, valor As String, cantidad As Double
    Dim tieneDatos As Boolean, colCantidad As Long, tieneClaveMB51 As Boolean
    Dim nombreNorm As String, materialNorm As String

    Set registro = CreateObject("Scripting.Dictionary")
    registro.CompareMode = vbTextCompare
    colCantidad = 0

    For i = LBound(encabezadoLocal) To UBound(encabezadoLocal)
        If i <= UBound(campos) Then valor = LimpiarTexto(campos(i)) Else valor = vbNullString
        If Len(valor) > 0 Then tieneDatos = True
        registro(encabezadoLocal(i)) = valor

        nombreNorm = NormalizarNombre(encabezadoLocal(i))
        If nombreNorm = "CANTIDAD" Then colCantidad = i + 1

        If nombreNorm = "MATERIAL" Or nombreNorm = "TEXTO BREVE DE MATERIAL" Then
            If Len(valor) > 0 Then
                materialNorm = NormalizarNombre(valor)
                ' Omite cabeceras repetidas de listados SAP y basura de totales
                If materialNorm <> "MATERIAL" And _
                   materialNorm <> "TEXTO BREVE DE MATERIAL" And _
                   materialNorm <> "TEXTO BREVE MATERIAL" Then
                    tieneClaveMB51 = True
                End If
            End If
        End If
    Next i

    If Not tieneDatos Then
        vacias = vacias + 1
        Exit Sub
    End If

    ' Descarta pies de pagina y renglones de totales de la salida dinamica SAP.
    If Not tieneClaveMB51 Then
        vacias = vacias + 1
        Exit Sub
    End If

    leidas = leidas + 1

    If colCantidad > 0 Then
        valor = CStr(registro(encabezadoLocal(colCantidad - 1)))
        If TryNumeroSAP(valor, cantidad) Then
            If cantidad < 0 Then
                If UCase$(MODO_NEGATIVOS) = "ELIMINAR" Then
                    negativas = negativas + 1
                    Exit Sub
                ElseIf UCase$(MODO_NEGATIVOS) = "ABSOLUTO" Then
                    cantidad = Abs(cantidad)
                    negativas = negativas + 1
                End If
            End If
            registro(encabezadoLocal(colCantidad - 1)) = cantidad
        End If
    End If

    ConvertirTipos registro
    registro("Archivo origen") = archivoOrigen
    registro("Hoja origen") = hojaOrigen

    If filas.Count + 2 > MAX_FILAS_EXCEL Then
        Err.Raise vbObjectError + 102, , "Se excedio el limite de filas de Excel"
    End If

    filas.Add registro
End Sub

Private Sub ConvertirTipos(ByRef registro As Object)
    Dim k As Variant, n As Double, d As Date, nombre As String
    For Each k In registro.Keys
        nombre = NormalizarNombre(CStr(k))
        If nombre = "FECHA DOC" Or nombre = "FECHA DOC." Or InStr(nombre, "FECHA") = 1 Then
            If TryFechaSAP(CStr(registro(k)), d) Then registro(k) = d
        ElseIf nombre = "IMPTE ML" Or nombre = "IMPTE.ML" Or nombre = "CANTIDAD" Then
            If TryNumeroSAP(CStr(registro(k)), n) Then registro(k) = n
        End If
    Next k
End Sub

Private Function TryNumeroSAP(ByVal texto As String, ByRef resultado As Double) As Boolean
    Dim s As String, pComa As Long, pPunto As Long, decSep As String
    Dim esNegativo As Boolean
    On Error GoTo NoNumero
    s = Replace(Trim$(texto), " ", "")
    s = Replace(s, ChrW(160), "")
    If Len(s) = 0 Then Exit Function

    ' SAP puede exportar negativos como 123,45-, (123,45) o con el signo normal.
    If Right$(s, 1) = "-" Then
        esNegativo = True
        s = Left$(s, Len(s) - 1)
    ElseIf Left$(s, 1) = "(" And Right$(s, 1) = ")" Then
        esNegativo = True
        s = Mid$(s, 2, Len(s) - 2)
    ElseIf Left$(s, 1) = "-" Then
        esNegativo = True
        s = Mid$(s, 2)
    End If

    ' Quita simbolos comunes sin alterar los separadores numericos.
    s = Replace(s, "$", "")
    s = Replace(s, "MXN", "", 1, -1, vbTextCompare)
    s = Replace(s, "USD", "", 1, -1, vbTextCompare)

    pComa = InStrRev(s, ",")
    pPunto = InStrRev(s, ".")

    If pComa > 0 And pPunto > 0 Then
        If pComa > pPunto Then
            s = Replace(s, ".", "")
            s = Replace(s, ",", ".")
        Else
            s = Replace(s, ",", "")
        End If
    ElseIf pComa > 0 Then
        If Len(s) - pComa = 3 Then
            s = Replace(s, ",", "")
        Else
            s = Replace(s, ",", ".")
        End If
    End If

    decSep = Application.International(xlDecimalSeparator)
    If decSep <> "." Then s = Replace(s, ".", decSep)
    resultado = CDbl(s)
    If esNegativo Then resultado = -Abs(resultado)
    TryNumeroSAP = True
    Exit Function
NoNumero:
    TryNumeroSAP = False
End Function

Private Sub CrearHojaResumen(ByVal wb As Workbook, ByVal carpeta As String, _
                              ByVal archivos As Long, ByVal leidas As Long, _
                              ByVal guardadas As Long, ByVal negativas As Long, _
                              ByVal vacias As Long, ByVal errores As Collection, _
                              ByVal segundos As Double)
    Dim ws As Worksheet, fila As Long, i As Long
    Set ws = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    ws.Name = "Resumen"

    ws.Range("A1:B1").Merge
    ws.Range("A1").Value = "Resumen de consolidacion MB51"
    ws.Range("A3").Value = "Fecha de ejecucion": ws.Range("B3").Value = Now
    ws.Range("A4").Value = "Carpeta origen": ws.Range("B4").Value = carpeta
    ws.Range("A5").Value = "Archivos revisados": ws.Range("B5").Value = archivos
    ws.Range("A6").Value = "Filas leidas": ws.Range("B6").Value = leidas
    ws.Range("A7").Value = "Filas guardadas": ws.Range("B7").Value = guardadas
    ws.Range("A8").Value = IIf(UCase$(MODO_NEGATIVOS) = "ELIMINAR", _
                                "Negativas eliminadas", "Negativas convertidas")
    ws.Range("B8").Value = negativas
    ws.Range("A9").Value = "Filas vacias/omitidas": ws.Range("B9").Value = vacias
    ws.Range("A10").Value = "Archivos con error": ws.Range("B10").Value = errores.Count
    ws.Range("A11").Value = "Segundos": ws.Range("B11").Value = segundos

    fila = 13
    ws.Cells(fila, 1).Value = "Detalle de errores"
    ws.Cells(fila, 1).Font.Bold = True
    If errores.Count = 0 Then
        ws.Cells(fila + 1, 1).Value = "Sin errores"
    Else
        For i = 1 To errores.Count
            ws.Cells(fila + i, 1).Value = errores(i)
        Next i
    End If

    With ws.Range("A1:B1")
        .Font.Bold = True
        .Font.Size = 14
        .Interior.Color = RGB(31, 78, 121)
        .Font.Color = vbWhite
        .HorizontalAlignment = xlCenter
    End With
    ws.Range("A3:A11").Font.Bold = True
    ws.Range("B3").NumberFormat = "dd/mm/yyyy hh:mm:ss"
    ws.Columns("A:B").AutoFit
    If ws.Columns("A").ColumnWidth > 80 Then ws.Columns("A").ColumnWidth = 80
    If ws.Columns("B").ColumnWidth > 80 Then ws.Columns("B").ColumnWidth = 80
    ws.Columns("A:B").WrapText = True
End Sub

Private Function TryFechaSAP(ByVal texto As String, ByRef resultado As Date) As Boolean
    Dim p() As String
    On Error GoTo NoFecha
    texto = Trim$(texto)
    If Len(texto) = 0 Then Exit Function

    p = Split(texto, ".")
    If UBound(p) = 2 Then
        resultado = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
        TryFechaSAP = True
        Exit Function
    End If

    If IsDate(texto) Then
        resultado = CDate(texto)
        TryFechaSAP = True
    End If
    Exit Function
NoFecha:
    TryFechaSAP = False
End Function

Private Function LimpiarTexto(ByVal s As String) As String
    s = Replace(s, ChrW(160), " ")
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")
    s = Replace(s, vbTab, " ")
    LimpiarTexto = WorksheetFunction.Trim(s)
End Function

Private Function NormalizarNombre(ByVal s As String) As String
    NormalizarNombre = UCase$(LimpiarTexto(s))
End Function

Private Sub VolcarConsolidado(ByVal ws As Worksheet, ByRef encabezados As Object, ByRef filas As Collection)
    Dim salida() As Variant, claves As Variant, i As Long, j As Long, reg As Object
    claves = encabezados.Keys
    ReDim salida(1 To filas.Count + 1, 1 To encabezados.Count)

    For j = 0 To UBound(claves)
        salida(1, j + 1) = claves(j)
    Next j

    For i = 1 To filas.Count
        Set reg = filas(i)
        For j = 0 To UBound(claves)
            If reg.Exists(claves(j)) Then salida(i + 1, j + 1) = reg(claves(j))
        Next j
    Next i

    ws.Range(ws.Cells(1, 1), ws.Cells(UBound(salida, 1), UBound(salida, 2))).Value = salida
End Sub

Private Sub EliminarColumnasVacias(ByVal ws As Worksheet)
    Dim c As Long, ultimaCol As Long, ultimaFila As Long
    Dim celda As Range

    Set celda = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultimaCol = celda.Column

    Set celda = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, LookAt:=xlPart, _
                              SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultimaFila = celda.Row

    If ultimaFila < 2 Then Exit Sub

    For c = ultimaCol To 1 Step -1
        If WorksheetFunction.CountA(ws.Range(ws.Cells(2, c), ws.Cells(ultimaFila, c))) = 0 Then
            ws.Columns(c).Delete
        End If
    Next c
End Sub

Private Sub DarFormatoFinal(ByVal ws As Worksheet)
    Dim ultimaFila As Long, ultimaCol As Long, c As Long, nombre As String
    Dim rango As Range, tabla As ListObject
    Dim celda As Range

    Set celda = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultimaFila = celda.Row

    Set celda = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                              LookIn:=xlFormulas, SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    If celda Is Nothing Then Exit Sub
    ultimaCol = celda.Column

    Set rango = ws.Range(ws.Cells(1, 1), ws.Cells(ultimaFila, ultimaCol))

    On Error Resume Next
    ws.ListObjects("tblMB51Consolidado").Delete
    On Error GoTo 0

    Set tabla = ws.ListObjects.Add(xlSrcRange, rango, , xlYes)
    tabla.Name = "tblMB51Consolidado"
    tabla.TableStyle = "TableStyleMedium2"

    ws.Rows(1).Font.Bold = True
    ws.Rows(1).RowHeight = 30
    ws.Rows(1).WrapText = True

    For c = 1 To ultimaCol
        nombre = NormalizarNombre(CStr(ws.Cells(1, c).Value))
        If nombre Like "*FECHA*" Then ws.Columns(c).NumberFormat = "dd/mm/yyyy"
        If nombre = "CANTIDAD" Or nombre Like "CANTIDAD_*" Then
            ws.Columns(c).NumberFormat = "#,##0.###;[Red]-#,##0.###"
        End If
        If nombre = "IMPTE ML" Or nombre = "IMPTE.ML" Or nombre Like "IMPTE*" Then
            ws.Columns(c).NumberFormat = "#,##0.00;[Red]-#,##0.00"
        End If
    Next c

    rango.Columns.AutoFit
    For c = 1 To ultimaCol
        If ws.Columns(c).ColumnWidth > 45 Then ws.Columns(c).ColumnWidth = 45
        If ws.Columns(c).ColumnWidth < 10 Then ws.Columns(c).ColumnWidth = 10
    Next c

    ws.Activate
    ActiveWindow.SplitRow = 1
    ActiveWindow.FreezePanes = True
    ws.Range("A2").Select
End Sub

Private Function TextoErrores(ByVal erroresLog As Collection, ByVal maxItems As Long) As String
    Dim i As Long, n As Long, s As String
    If erroresLog Is Nothing Then Exit Function
    If erroresLog.Count = 0 Then Exit Function

    n = erroresLog.Count
    If n > maxItems Then n = maxItems

    s = "Detalle de errores (max " & maxItems & "):"
    For i = 1 To n
        s = s & vbCrLf & "- " & erroresLog(i)
    Next i
    If erroresLog.Count > maxItems Then
        s = s & vbCrLf & "... y " & (erroresLog.Count - maxItems) & " mas."
    End If
    TextoErrores = s
End Function
