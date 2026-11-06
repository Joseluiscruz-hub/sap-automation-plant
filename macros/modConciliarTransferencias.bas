Attribute VB_Name = "Módulo6"
' =========================================================
' MODULO UNICO LIMPIO - CONCILIACION DE RUTAS
' Autor: generado para integrarse al libro Conciliacion de Rutas.xlsm
' Uso:
'   - ConciliarTransferenciasOptimizado
'   - RestablecerDatos
'   - ValidarDatosAntesProcesar
'   - GenerarEstadisticas
' =========================================================
Option Explicit

Private Const SH_LIQ As String = "Liquidacion de Ruta"
Private Const SH_DET As String = "Detalle Cliente DA - Vendedor"
Private Const SH_RES As String = "Resumen Transferencias"

Private Const FIRST_LIQ_ROW As Long = 8
Private Const FIRST_DET_ROW As Long = 2

Private Const COL_LIQ_FECHA As Long = 1       ' A
Private Const COL_LIQ_SALDO As Long = 10      ' J
Private Const COL_LIQ_OUT_MONTO As Long = 11  ' K
Private Const COL_LIQ_OUT_FECHA As Long = 12  ' L
Private Const COL_LIQ_OUT_DOCS As Long = 13   ' M

Private Const COL_DET_CLIENTE As Long = 1     ' A
Private Const COL_DET_DOCUMENTO As Long = 5   ' E
Private Const COL_DET_FECHA As Long = 8       ' H
Private Const COL_DET_IMPORTE As Long = 11    ' K

Private Const MAX_DP_STATES As Long = 80000

Private Type DocumentoInfo
    Fila As Long
    Cliente As String
    Fecha As Date
    Importe As Currency
    ImporteCents As Double
    Documento As String
    Usado As Boolean
End Type

Private Type LiquidacionInfo
    Fila As Long
    Fecha As Date
    Saldo As Currency
    SaldoCents As Double
    FechaLimite As Date
End Type

Private Type ResultadoConciliacion
    FilaLiquidacion As Long
    Fecha As Date
    MontoTransferir As Currency
    MontoTransferido As Currency
    Documentos As String
    Conciliado As Boolean
    Metodo As String
End Type

Private documentos() As DocumentoInfo
Private liquidaciones() As LiquidacionInfo
Private resultados() As ResultadoConciliacion

Private nDocs As Long
Private nLiq As Long

Private indiceImportes As Object

Private oldScreenUpdating As Boolean
Private oldEnableEvents As Boolean
Private oldDisplayAlerts As Boolean
Private oldCalculation As XlCalculation
Private oldStatusBar As Variant

' ==========================================
' MACRO PRINCIPAL
' ==========================================
Public Sub ConciliarTransferenciasOptimizado()
    Dim tolExacta As Currency
    Dim tolCombinacion As Currency

    Dim tolComboDefault As Currency
    tolComboDefault = ObtenerToleranciaDefaultHoja()
    If tolComboDefault <= 0 Then tolComboDefault = 1

    If Not ObtenerTolerancia("Ingrese la tolerancia para importes exactos:", 0.1, tolExacta) Then Exit Sub
    If Not ObtenerTolerancia("Ingrese la tolerancia para combinaciones:", tolComboDefault, tolCombinacion) Then Exit Sub

    PrepararExcelRapido

    On Error GoTo ErrorHandler

    Application.StatusBar = "Cargando datos de liquidacion y detalle..."

    If Not CargarDatos() Then
        MsgBox "No se pudieron cargar los datos. Revisa nombres de hojas, columnas y formatos.", vbCritical
        GoTo Salida
    End If

    If nLiq = 0 Then
        MsgBox "No hay liquidaciones validas para procesar.", vbExclamation
        GoTo Salida
    End If

    If nDocs = 0 Then
        MsgBox "No hay documentos validos para el cliente indicado en la liquidacion.", vbExclamation
        GoTo Salida
    End If

    CrearIndiceImportes

    Application.StatusBar = "Conciliando importes exactos y combinaciones..."
    EjecutarConciliacion tolExacta, tolCombinacion

    Application.StatusBar = "Aplicando resultados..."
    AplicarResultados

    Application.StatusBar = "Generando resumen..."
    GenerarResumen tolExacta, tolCombinacion

    MsgBox "Conciliacion completada." & vbCrLf & _
           "Liquidaciones procesadas: " & nLiq & vbCrLf & _
           "Documentos evaluados: " & nDocs, vbInformation

Salida:
    RestaurarExcel
    LimpiarMemoria
    Exit Sub

ErrorHandler:
    MsgBox "Error durante la conciliacion: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Salida
End Sub



Public Sub ConciliarTransferenciasOptimizadoV2()
    ConciliarTransferenciasOptimizado
End Sub

Public Sub ConciliarConMismoSignoOptimizado()
    ConciliarTransferenciasOptimizado
End Sub

Public Sub RestablecerDatos()
    On Error GoTo ErrorHandler

    PrepararExcelRapido
    Application.StatusBar = "Restableciendo salidas de conciliacion..."

    Dim wsLiq As Worksheet
    Dim wsDet As Worksheet
    Dim wsRes As Worksheet
    Dim lastLiq As Long
    Dim lastDet As Long
    Dim r As Long

    Set wsLiq = ThisWorkbook.Worksheets(SH_LIQ)
    Set wsDet = ThisWorkbook.Worksheets(SH_DET)

    lastLiq = UltimaFila(wsLiq, COL_LIQ_FECHA)
    If lastLiq >= FIRST_LIQ_ROW Then
        With wsLiq.Range(wsLiq.Cells(FIRST_LIQ_ROW, COL_LIQ_OUT_MONTO), wsLiq.Cells(lastLiq, COL_LIQ_OUT_DOCS))
            .ClearContents
            .Font.Color = vbBlack
        End With
    End If

    lastDet = UltimaFila(wsDet, COL_DET_CLIENTE)
    If lastDet >= FIRST_DET_ROW Then
        For r = FIRST_DET_ROW To lastDet
            If wsDet.Rows(r).Interior.Color = RGB(255, 255, 0) Or wsDet.Rows(r).Interior.Color = RGB(220, 255, 220) Then
                wsDet.Rows(r).Interior.ColorIndex = xlColorIndexNone
            End If
            If r Mod 1000 = 0 Then DoEvents
        Next r
    End If

    If HojaExiste(SH_RES) Then
        Set wsRes = ThisWorkbook.Worksheets(SH_RES)
        wsRes.Delete
    End If

    MsgBox "Restablecimiento completado. Se limpiaron K:M, colores de documentos usados y hoja de resumen.", vbInformation

Salida:
    RestaurarExcel
    LimpiarMemoria
    Exit Sub

ErrorHandler:
    MsgBox "Error durante el restablecimiento: " & Err.Number & " - " & Err.Description, vbCritical
    Resume Salida
End Sub

Public Sub ValidarDatosAntesProcesar()
    ValidarDatosAntesProcesarV2
End Sub

Public Sub GenerarEstadisticas()
    On Error GoTo ErrorHandler

    If Not HojaExiste(SH_RES) Then
        MsgBox "Aun no existe la hoja '" & SH_RES & "'. Ejecuta primero la conciliacion.", vbInformation
        Exit Sub
    End If

    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_RES)

    Dim mensaje As String
    mensaje = "ESTADISTICAS DE CONCILIACION" & vbCrLf & vbCrLf
    mensaje = mensaje & "Conciliados: " & ws.Range("K4").Text & vbCrLf
    mensaje = mensaje & "Pendientes: " & ws.Range("K5").Text & vbCrLf
    mensaje = mensaje & "% conciliacion: " & ws.Range("K6").Text & vbCrLf & vbCrLf
    mensaje = mensaje & "Total a transferir: " & ws.Range("K7").Text & vbCrLf
    mensaje = mensaje & "Total transferido: " & ws.Range("K8").Text & vbCrLf
    mensaje = mensaje & "Diferencia: " & ws.Range("K9").Text

    MsgBox mensaje, vbInformation, "Estadisticas"
    Exit Sub

ErrorHandler:
    MsgBox "No se pudieron generar las estadisticas: " & Err.Description, vbExclamation
End Sub
' ==========================================
' CONFIGURACION DE EXCEL
' ==========================================
Private Sub PrepararExcelRapido()
    oldScreenUpdating = Application.ScreenUpdating
    oldEnableEvents = Application.EnableEvents
    oldDisplayAlerts = Application.DisplayAlerts
    oldCalculation = Application.Calculation
    oldStatusBar = Application.StatusBar

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando conciliacion..."
End Sub

Private Sub RestaurarExcel()
    On Error Resume Next
    Application.ScreenUpdating = oldScreenUpdating
    Application.EnableEvents = oldEnableEvents
    Application.DisplayAlerts = oldDisplayAlerts
    Application.Calculation = oldCalculation
    Application.StatusBar = oldStatusBar
    On Error GoTo 0
End Sub

' ==========================================
' CARGA DE DATOS
' ==========================================
Private Function CargarDatos() As Boolean
    On Error GoTo ErrorHandler

    Dim wsLiq As Worksheet
    Dim wsDet As Worksheet
    Dim arrLiq As Variant
    Dim arrDet As Variant
    Dim lastLiq As Long
    Dim lastDet As Long
    Dim cliente As String
    Dim diasRango As Long
    Dim i As Long

    Set wsLiq = ThisWorkbook.Worksheets(SH_LIQ)
    Set wsDet = ThisWorkbook.Worksheets(SH_DET)

    lastLiq = UltimaFila(wsLiq, COL_LIQ_FECHA)
    lastDet = UltimaFila(wsDet, COL_DET_CLIENTE)

    If lastLiq < FIRST_LIQ_ROW Or lastDet < FIRST_DET_ROW Then
        CargarDatos = False
        Exit Function
    End If

    arrLiq = wsLiq.Range("A1:L" & lastLiq).Value2
    arrDet = wsDet.Range("A1:K" & lastDet).Value2

    cliente = Trim$(CStr(arrLiq(2, 10))) ' J2
    diasRango = CLng(Val(arrLiq(2, 12))) ' L2

    If Len(cliente) = 0 Then
        MsgBox "El cliente en J2 esta vacio.", vbExclamation
        CargarDatos = False
        Exit Function
    End If

    nLiq = 0
    For i = FIRST_LIQ_ROW To UBound(arrLiq, 1)
        If EsFechaExcel(arrLiq(i, COL_LIQ_FECHA)) And EsNumero(arrLiq(i, COL_LIQ_SALDO)) Then
            nLiq = nLiq + 1
        End If
    Next i

    nDocs = 0
    For i = FIRST_DET_ROW To UBound(arrDet, 1)
        If StrComp(Trim$(CStr(arrDet(i, COL_DET_CLIENTE))), cliente, vbTextCompare) = 0 Then
            If EsFechaExcel(arrDet(i, COL_DET_FECHA)) And EsNumero(arrDet(i, COL_DET_IMPORTE)) Then
                nDocs = nDocs + 1
            End If
        End If
    Next i

    If nLiq > 0 Then ReDim liquidaciones(1 To nLiq)
    If nDocs > 0 Then ReDim documentos(1 To nDocs)
    If nLiq > 0 Then ReDim resultados(1 To nLiq)

    nLiq = 0
    For i = FIRST_LIQ_ROW To UBound(arrLiq, 1)
        If EsFechaExcel(arrLiq(i, COL_LIQ_FECHA)) And EsNumero(arrLiq(i, COL_LIQ_SALDO)) Then
            nLiq = nLiq + 1
            With liquidaciones(nLiq)
                .Fila = i
                .Fecha = FechaExcel(arrLiq(i, COL_LIQ_FECHA))
                .Saldo = Moneda2(arrLiq(i, COL_LIQ_SALDO))
                .SaldoCents = ToCents(.Saldo)
                .FechaLimite = DateAdd("d", diasRango, .Fecha)
            End With
        End If
    Next i

    nDocs = 0
    For i = FIRST_DET_ROW To UBound(arrDet, 1)
        If StrComp(Trim$(CStr(arrDet(i, COL_DET_CLIENTE))), cliente, vbTextCompare) = 0 Then
            If EsFechaExcel(arrDet(i, COL_DET_FECHA)) And EsNumero(arrDet(i, COL_DET_IMPORTE)) Then
                nDocs = nDocs + 1
                With documentos(nDocs)
                    .Fila = i
                    .Cliente = Trim$(CStr(arrDet(i, COL_DET_CLIENTE)))
                    .Fecha = FechaExcel(arrDet(i, COL_DET_FECHA))
                    .Importe = Moneda2(arrDet(i, COL_DET_IMPORTE))
                    .ImporteCents = ToCents(.Importe)
                    .Documento = Trim$(CStr(arrDet(i, COL_DET_DOCUMENTO)))
                    .Usado = False
                End With
            End If
        End If
    Next i

    CargarDatos = True
    Exit Function

ErrorHandler:
    CargarDatos = False
End Function

Private Sub CrearIndiceImportes()
    Dim i As Long
    Dim key As String
    Dim col As Collection

    Set indiceImportes = CreateObject("Scripting.Dictionary")

    For i = 1 To nDocs
        key = KeyCents(documentos(i).ImporteCents)

        If indiceImportes.Exists(key) Then
            Set col = indiceImportes(key)
        Else
            Set col = New Collection
            indiceImportes.Add key, col
        End If

        col.Add i
    Next i
End Sub

' ==========================================
' LOGICA DE CONCILIACION
' ==========================================
Private Sub EjecutarConciliacion(ByVal tolExacta As Currency, ByVal tolCombinacion As Currency)
    Dim i As Long
    Dim tolExactCents As Double
    Dim tolCombCents As Double

    tolExactCents = ToCents(tolExacta)
    tolCombCents = ToCents(tolCombinacion)

    ' Fase 1: coincidencia exacta por importe usando indice.
    For i = 1 To nLiq
        InicializarResultado i, resultados(i)

        If BuscarExacta(i, tolExactCents, resultados(i)) Then
            resultados(i).Metodo = "Exacto"
        End If

        If i Mod 50 = 0 Then
            Application.StatusBar = "Fase exacta: " & i & " de " & nLiq
            DoEvents
        End If
    Next i

    ' Fase 2: combinaciones solo para pendientes.
    For i = 1 To nLiq
        If Not resultados(i).Conciliado Then
            If BuscarCombinada(i, tolCombCents, resultados(i)) Then
                resultados(i).Metodo = "Combinacion"
            End If
        End If

        If i Mod 20 = 0 Then
            Application.StatusBar = "Fase combinaciones: " & i & " de " & nLiq
            DoEvents
        End If
    Next i
End Sub

Private Sub InicializarResultado(ByVal idxLiq As Long, ByRef resultado As ResultadoConciliacion)
    With resultado
        .FilaLiquidacion = liquidaciones(idxLiq).Fila
        .Fecha = liquidaciones(idxLiq).Fecha
        .MontoTransferir = liquidaciones(idxLiq).Saldo
        .MontoTransferido = 0
        .Documentos = vbNullString
        .Conciliado = False
        .Metodo = "Pendiente"
    End With
End Sub

Private Function BuscarExacta(ByVal idxLiq As Long, ByVal toleranciaCents As Double, _
                             ByRef resultado As ResultadoConciliacion) As Boolean
    Dim targetCents As Double
    Dim baseCents As Double
    Dim offset As Long
    Dim maxOffset As Long
    Dim key As String
    Dim col As Collection
    Dim item As Variant
    Dim docIdx As Long
    Dim bestIdx As Long
    Dim bestDiff As Double
    Dim diff As Double

    targetCents = -liquidaciones(idxLiq).SaldoCents
    baseCents = Application.WorksheetFunction.Round(targetCents, 0)
    maxOffset = CLng(Application.WorksheetFunction.Round(toleranciaCents, 0))

    bestIdx = 0
    bestDiff = toleranciaCents + 1

    For offset = -maxOffset To maxOffset
        key = KeyCents(baseCents + offset)

        If indiceImportes.Exists(key) Then
            Set col = indiceImportes(key)

            For Each item In col
                docIdx = CLng(item)

                If DocumentoDisponibleEnRango(docIdx, liquidaciones(idxLiq)) Then
                    diff = Abs(documentos(docIdx).ImporteCents - targetCents)

                    If diff <= toleranciaCents And diff < bestDiff Then
                        bestDiff = diff
                        bestIdx = docIdx
                        If bestDiff = 0 Then Exit For
                    End If
                End If
            Next item
        End If

        If bestDiff = 0 Then Exit For
    Next offset

    If bestIdx > 0 Then
        documentos(bestIdx).Usado = True

        With resultado
            .MontoTransferido = documentos(bestIdx).Importe
            .Documentos = documentos(bestIdx).Documento
            .Conciliado = True
        End With

        BuscarExacta = True
    End If
End Function

Private Function BuscarCombinada(ByVal idxLiq As Long, ByVal toleranciaCents As Double, _
                                 ByRef resultado As ResultadoConciliacion) As Boolean
    Dim targetCents As Double
    Dim candidatos() As Long
    Dim nCand As Long
    Dim seleccion() As Long
    Dim selCount As Long
    Dim sumaCents As Double
    Dim i As Long
    Dim docIdx As Long
    Dim docsTexto As String

    targetCents = -liquidaciones(idxLiq).SaldoCents

    RecolectarCandidatos idxLiq, targetCents, toleranciaCents, candidatos, nCand

    If nCand = 0 Then Exit Function

    OrdenarCandidatosPorImporte candidatos, 1, nCand

    If BuscarSubconjuntoDP(candidatos, nCand, targetCents, toleranciaCents, seleccion, selCount, sumaCents) Then
        For i = 1 To selCount
            docIdx = seleccion(i)
            documentos(docIdx).Usado = True

            If Len(docsTexto) > 0 Then docsTexto = docsTexto & ", "
            docsTexto = docsTexto & documentos(docIdx).Documento
        Next i

        With resultado
            .MontoTransferido = FromCents(sumaCents)
            .Documentos = docsTexto
            .Conciliado = True
        End With

        BuscarCombinada = True
    End If
End Function

Private Sub RecolectarCandidatos(ByVal idxLiq As Long, ByVal targetCents As Double, ByVal toleranciaCents As Double, ByRef candidatos() As Long, ByRef nCand As Long)
    Dim i As Long
    Dim cap As Long
    Dim sTarget As Long
    Dim imp As Double

    nCand = 0
    cap = 0
    sTarget = Sgn(targetCents)

    For i = 1 To nDocs
        If DocumentoDisponibleEnRango(i, liquidaciones(idxLiq)) Then
            imp = documentos(i).ImporteCents

            If imp <> 0 Then
                ' Mantiene el sentido contable: saldo negativo se cancela con documentos positivos, y viceversa.
                If sTarget = 0 Or Sgn(imp) = sTarget Then
                    If sTarget > 0 Then
                        If imp <= targetCents + toleranciaCents Then AddIndex candidatos, nCand, cap, i
                    ElseIf sTarget < 0 Then
                        If imp >= targetCents - toleranciaCents Then AddIndex candidatos, nCand, cap, i
                    Else
                        AddIndex candidatos, nCand, cap, i
                    End If
                End If
            End If
        End If
    Next i
End Sub

Private Function BuscarSubconjuntoDP(ByRef candidatos() As Long, ByVal nCand As Long, _
                                     ByVal objetivoCents As Double, ByVal toleranciaCents As Double, _
                                     ByRef seleccion() As Long, ByRef selCount As Long, _
                                     ByRef sumaBestCents As Double) As Boolean
    Dim estados As Object
    Dim keys As Variant
    Dim key As Variant
    Dim newKey As String
    Dim combo As String
    Dim newCombo As String
    Dim newSum As Double
    Dim imp As Double
    Dim i As Long
    Dim docIdx As Long
    Dim diff As Double
    Dim bestDiff As Double
    Dim bestKey As String

    Set estados = CreateObject("Scripting.Dictionary")
    estados.Add "0", vbNullString

    bestDiff = Abs(objetivoCents)
    bestKey = vbNullString

    For i = 1 To nCand
        docIdx = candidatos(i)
        imp = documentos(docIdx).ImporteCents
        keys = estados.Keys

        For Each key In keys
            newSum = CDbl(key) + imp

            If SumaPermitida(newSum, objetivoCents, toleranciaCents) Then
                newKey = KeyCents(newSum)

                If Not estados.Exists(newKey) Then
                    combo = CStr(estados(CStr(key)))

                    If Len(combo) > 0 Then
                        newCombo = combo & "|" & CStr(docIdx)
                    Else
                        newCombo = CStr(docIdx)
                    End If

                    estados.Add newKey, newCombo

                    diff = Abs(newSum - objetivoCents)
                    If diff < bestDiff Then
                        bestDiff = diff
                        bestKey = newKey
                    End If

                    If diff <= toleranciaCents Then
                        DecodificarSeleccion newCombo, seleccion, selCount
                        sumaBestCents = newSum
                        BuscarSubconjuntoDP = True
                        Exit Function
                    End If

                    If estados.Count > MAX_DP_STATES Then
                        ' Evita consumir demasiada memoria en casos enormes.
                        BuscarSubconjuntoDP = BuscarPorHeuristica(candidatos, nCand, objetivoCents, _
                                                                  toleranciaCents, seleccion, selCount, sumaBestCents)
                        Exit Function
                    End If
                End If
            End If
        Next key
    Next i

    If Len(bestKey) > 0 Then
        If Abs(CDbl(bestKey) - objetivoCents) <= toleranciaCents Then
            DecodificarSeleccion CStr(estados(bestKey)), seleccion, selCount
            sumaBestCents = CDbl(bestKey)
            BuscarSubconjuntoDP = True
        End If
    End If
End Function

Private Function BuscarPorHeuristica(ByRef candidatos() As Long, ByVal nCand As Long, _
                                     ByVal objetivoCents As Double, ByVal toleranciaCents As Double, _
                                     ByRef seleccion() As Long, ByRef selCount As Long, _
                                     ByRef sumaTotalCents As Double) As Boolean
    Dim i As Long
    Dim docIdx As Long
    Dim actual As Double
    Dim nuevo As Double
    Dim diffActual As Double
    Dim diffNuevo As Double
    Dim cap As Long

    selCount = 0
    cap = 0
    actual = 0

    For i = 1 To nCand
        docIdx = candidatos(i)
        nuevo = actual + documentos(docIdx).ImporteCents

        If SumaPermitida(nuevo, objetivoCents, toleranciaCents) Then
            diffActual = Abs(actual - objetivoCents)
            diffNuevo = Abs(nuevo - objetivoCents)

            If diffNuevo <= diffActual Then
                actual = nuevo
                AddIndex seleccion, selCount, cap, docIdx

                If Abs(actual - objetivoCents) <= toleranciaCents Then
                    sumaTotalCents = actual
                    BuscarPorHeuristica = True
                    Exit Function
                End If
            End If
        End If
    Next i

    sumaTotalCents = actual
    BuscarPorHeuristica = (Abs(actual - objetivoCents) <= toleranciaCents)
End Function

Private Function DocumentoDisponibleEnRango(ByVal docIdx As Long, ByRef liq As LiquidacionInfo) As Boolean
    With documentos(docIdx)
        DocumentoDisponibleEnRango = (Not .Usado) _
                                     And (.Fecha >= liq.Fecha) _
                                     And (.Fecha <= liq.FechaLimite)
    End With
End Function

Private Function SumaPermitida(ByVal sumaCents As Double, ByVal objetivoCents As Double, _
                               ByVal toleranciaCents As Double) As Boolean
    If objetivoCents >= 0 Then
        SumaPermitida = (sumaCents >= 0 And sumaCents <= objetivoCents + toleranciaCents)
    Else
        SumaPermitida = (sumaCents <= 0 And sumaCents >= objetivoCents - toleranciaCents)
    End If
End Function

' ==========================================
' APLICACION DE RESULTADOS
' ==========================================
Private Sub AplicarResultados()
    Dim wsLiq As Worksheet
    Dim wsDet As Worksheet
    Dim lastLiq As Long
    Dim outArr As Variant
    Dim i As Long
    Dim rowOffset As Long
    Dim totalRows As Long

    Set wsLiq = ThisWorkbook.Worksheets(SH_LIQ)
    Set wsDet = ThisWorkbook.Worksheets(SH_DET)

    lastLiq = UltimaFila(wsLiq, COL_LIQ_FECHA)
    totalRows = lastLiq - FIRST_LIQ_ROW + 1

    If totalRows > 0 Then
        ReDim outArr(1 To totalRows, 1 To 3)

        For i = 1 To totalRows
            outArr(i, 1) = 0
            outArr(i, 2) = vbNullString
            outArr(i, 3) = vbNullString
        Next i

        For i = 1 To nLiq
            rowOffset = resultados(i).FilaLiquidacion - FIRST_LIQ_ROW + 1

            If rowOffset >= 1 And rowOffset <= totalRows Then
                If resultados(i).Conciliado Then
                    outArr(rowOffset, 1) = resultados(i).MontoTransferido
                    outArr(rowOffset, 2) = resultados(i).Fecha
                    outArr(rowOffset, 3) = resultados(i).Documentos
                Else
                    outArr(rowOffset, 1) = 0
                End If
            End If
        Next i

        With wsLiq.Cells(FIRST_LIQ_ROW, COL_LIQ_OUT_MONTO).Resize(totalRows, 3)
            .Value = outArr
            .Columns(1).NumberFormat = "$#,##0.00"
            .Columns(2).NumberFormat = "dd/mm/yyyy"
        End With

        wsLiq.Cells(FIRST_LIQ_ROW, COL_LIQ_OUT_MONTO).Resize(totalRows, 1).Font.Color = RGB(0, 0, 0)

        For i = 1 To nLiq
            If resultados(i).Conciliado Then
                wsLiq.Cells(resultados(i).FilaLiquidacion, COL_LIQ_OUT_MONTO).Font.Color = RGB(255, 0, 0)
            End If
        Next i
    End If

    MarcarDocumentosUsados wsDet
End Sub

Private Sub MarcarDocumentosUsados(ByVal wsDet As Worksheet)
    Dim i As Long
    Dim rng As Range
    Dim lote As Long

    For i = 1 To nDocs
        If documentos(i).Usado Then
            If rng Is Nothing Then
                Set rng = wsDet.Rows(documentos(i).Fila)
            Else
                Set rng = Union(rng, wsDet.Rows(documentos(i).Fila))
            End If

            lote = lote + 1

            If lote >= 100 Then
                rng.Interior.Color = RGB(255, 255, 0)
                Set rng = Nothing
                lote = 0
            End If
        End If
    Next i

    If Not rng Is Nothing Then rng.Interior.Color = RGB(255, 255, 0)
End Sub

' ==========================================
' RESUMEN
' ==========================================
Private Sub GenerarResumen(ByVal tolExacta As Currency, ByVal tolCombinacion As Currency)
    Dim ws As Worksheet
    Dim data As Variant
    Dim i As Long
    Dim conciliados As Long
    Dim pendientes As Long
    Dim totalTransferir As Currency
    Dim totalTransferido As Currency

    Set ws = ObtenerOCrearHoja(SH_RES, SH_DET)
    ws.Cells.Clear

    ws.Range("A1:H1").Value = Array("Fila Liquidacion", "Fecha", "Monto a Transferir", _
                                    "Monto Transferido", "Diferencia", "Documentos Usados", _
                                    "Estado", "Metodo")
    With ws.Range("A1:H1")
        .Font.Bold = True
        .Interior.Color = RGB(200, 200, 200)
    End With

    If nLiq > 0 Then
        ReDim data(1 To nLiq, 1 To 8)

        For i = 1 To nLiq
            With resultados(i)
                data(i, 1) = .FilaLiquidacion
                data(i, 2) = .Fecha
                data(i, 3) = .MontoTransferir
                data(i, 4) = .MontoTransferido
                data(i, 5) = .MontoTransferir + .MontoTransferido
                data(i, 6) = .Documentos
                data(i, 7) = IIf(.Conciliado, "Conciliado", "Pendiente")
                data(i, 8) = .Metodo

                totalTransferir = totalTransferir + .MontoTransferir
                totalTransferido = totalTransferido + .MontoTransferido

                If .Conciliado Then
                    conciliados = conciliados + 1
                Else
                    pendientes = pendientes + 1
                End If
            End With
        Next i

        ws.Range("A2").Resize(nLiq, 8).Value = data
        ws.Range("B2:B" & nLiq + 1).NumberFormat = "dd/mm/yyyy"
        ws.Range("C2:E" & nLiq + 1).NumberFormat = "$#,##0.00"

        For i = 2 To nLiq + 1
            If ws.Cells(i, 7).Value = "Conciliado" Then
                ws.Rows(i).Interior.Color = RGB(220, 255, 220)
            Else
                ws.Rows(i).Interior.Color = RGB(255, 220, 220)
            End If
        Next i
    End If

    ws.Range("J1").Value = "Estadisticas"
    ws.Range("J1").Font.Bold = True

    ws.Range("J2").Value = "Tolerancia exactos"
    ws.Range("K2").Value = tolExacta

    ws.Range("J3").Value = "Tolerancia combinaciones"
    ws.Range("K3").Value = tolCombinacion

    ws.Range("J4").Value = "Conciliados"
    ws.Range("K4").Value = conciliados

    ws.Range("J5").Value = "Pendientes"
    ws.Range("K5").Value = pendientes

    ws.Range("J6").Value = "% conciliacion"
    If nLiq > 0 Then ws.Range("K6").Value = conciliados / nLiq

    ws.Range("J7").Value = "Total a transferir"
    ws.Range("K7").Value = totalTransferir

    ws.Range("J8").Value = "Total transferido"
    ws.Range("K8").Value = totalTransferido

    ws.Range("J9").Value = "Diferencia"
    ws.Range("K9").Value = totalTransferir + totalTransferido

    ws.Range("K2:K3,K7:K9").NumberFormat = "$#,##0.00"
    ws.Range("K6").NumberFormat = "0.00%"
    ws.Columns.AutoFit
End Sub

' ==========================================
' VALIDACION OPCIONAL
' ==========================================
Public Sub ValidarDatosAntesProcesarV2()
    Dim errores As String
    Dim wsLiq As Worksheet
    Dim wsDet As Worksheet

    On Error Resume Next
    Set wsLiq = ThisWorkbook.Worksheets(SH_LIQ)
    Set wsDet = ThisWorkbook.Worksheets(SH_DET)
    On Error GoTo 0

    If wsLiq Is Nothing Then errores = errores & "- Falta la hoja '" & SH_LIQ & "'" & vbCrLf
    If wsDet Is Nothing Then errores = errores & "- Falta la hoja '" & SH_DET & "'" & vbCrLf

    If Not wsLiq Is Nothing Then
        If Len(Trim$(CStr(wsLiq.Range("J2").Value))) = 0 Then errores = errores & "- Falta cliente en J2" & vbCrLf
        If Not IsNumeric(wsLiq.Range("L2").Value) Then errores = errores & "- Dias de rango en L2 no es numerico" & vbCrLf
    End If

    If Len(errores) > 0 Then
        MsgBox "Errores encontrados:" & vbCrLf & errores, vbCritical, "Validacion"
    Else
        MsgBox "Validacion basica correcta.", vbInformation, "Validacion"
    End If
End Sub

' ==========================================
' UTILIDADES
' ==========================================

Private Function ObtenerToleranciaDefaultHoja() As Currency
    On Error GoTo Fallback
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Worksheets(SH_LIQ)
    If EsNumero(ws.Cells(2, 11).Value2) Then
        ObtenerToleranciaDefaultHoja = Moneda2(ws.Cells(2, 11).Value2)
    Else
        ObtenerToleranciaDefaultHoja = 1
    End If
    Exit Function
Fallback:
    ObtenerToleranciaDefaultHoja = 1
End Function

Private Function HojaExiste(ByVal sheetName As String) As Boolean
    On Error Resume Next
    HojaExiste = Not ThisWorkbook.Worksheets(sheetName) Is Nothing
    On Error GoTo 0
End Function

Private Function ObtenerTolerancia(ByVal prompt As String, ByVal valorDefault As Currency, _
                                   ByRef tolerancia As Currency) As Boolean
    Dim userInput As String
    Dim valor As Double

    userInput = Trim$(InputBox(prompt, "Tolerancia", Format$(valorDefault, "0.00")))

    If Len(userInput) = 0 Then
        ObtenerTolerancia = False
        Exit Function
    End If

    If Not TryParseDouble(userInput, valor) Then
        MsgBox "Debe ingresar un valor numerico valido.", vbExclamation
        ObtenerTolerancia = False
        Exit Function
    End If

    If valor < 0 Then
        MsgBox "La tolerancia debe ser cero o positiva.", vbExclamation
        ObtenerTolerancia = False
        Exit Function
    End If

    tolerancia = CCur(Application.WorksheetFunction.Round(valor, 2))
    ObtenerTolerancia = True
End Function

Private Function TryParseDouble(ByVal textValue As String, ByRef result As Double) As Boolean
    Dim decSep As String

    textValue = Trim$(textValue)
    decSep = Application.DecimalSeparator

    If decSep = "," Then
        If InStr(1, textValue, ".", vbBinaryCompare) > 0 And InStr(1, textValue, ",", vbBinaryCompare) = 0 Then
            textValue = Replace(textValue, ".", ",")
        End If
    Else
        If InStr(1, textValue, ",", vbBinaryCompare) > 0 And InStr(1, textValue, ".", vbBinaryCompare) = 0 Then
            textValue = Replace(textValue, ",", ".")
        End If
    End If

    If IsNumeric(textValue) Then
        result = CDbl(textValue)
        TryParseDouble = True
    End If
End Function

Private Function EsNumero(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function
    If Len(Trim$(CStr(value))) = 0 Then Exit Function
    EsNumero = IsNumeric(value)
End Function

Private Function EsFechaExcel(ByVal value As Variant) As Boolean
    If IsError(value) Then Exit Function

    If IsDate(value) Then
        EsFechaExcel = True
    ElseIf IsNumeric(value) Then
        EsFechaExcel = (CDbl(value) > 0)
    End If
End Function

Private Function FechaExcel(ByVal value As Variant) As Date
    If IsDate(value) Then
        FechaExcel = CDate(value)
    Else
        FechaExcel = DateSerial(1899, 12, 30) + CDbl(value)
    End If
End Function

Private Function Moneda2(ByVal value As Variant) As Currency
    Moneda2 = CCur(Application.WorksheetFunction.Round(CDbl(value), 2))
End Function

Private Function ToCents(ByVal amount As Currency) As Double
    ToCents = Application.WorksheetFunction.Round(CDbl(amount) * 100, 0)
End Function

Private Function FromCents(ByVal cents As Double) As Currency
    FromCents = CCur(cents / 100#)
End Function

Private Function KeyCents(ByVal cents As Double) As String
    KeyCents = Format$(Application.WorksheetFunction.Round(cents, 0), "0")
End Function

Private Function UltimaFila(ByVal ws As Worksheet, ByVal col As Long) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, col).End(xlUp).Row
End Function

Private Sub AddIndex(ByRef arr() As Long, ByRef count As Long, ByRef capacity As Long, ByVal value As Long)
    If capacity = 0 Then
        capacity = 32
        ReDim arr(1 To capacity)
    ElseIf count >= capacity Then
        capacity = capacity * 2
        ReDim Preserve arr(1 To capacity)
    End If

    count = count + 1
    arr(count) = value
End Sub

Private Sub DecodificarSeleccion(ByVal combo As String, ByRef seleccion() As Long, ByRef selCount As Long)
    Dim parts As Variant
    Dim i As Long

    If Len(combo) = 0 Then
        selCount = 0
        Exit Sub
    End If

    parts = Split(combo, "|")
    selCount = UBound(parts) - LBound(parts) + 1
    ReDim seleccion(1 To selCount)

    For i = 0 To UBound(parts)
        seleccion(i + 1) = CLng(parts(i))
    Next i
End Sub

Private Sub OrdenarCandidatosPorImporte(ByRef arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim i As Long
    Dim j As Long
    Dim pivot As Double
    Dim temp As Long

    If first >= last Then Exit Sub

    i = first
    j = last
    pivot = Abs(documentos(arr((first + last) \ 2)).ImporteCents)

    Do While i <= j
        Do While Abs(documentos(arr(i)).ImporteCents) > pivot
            i = i + 1
        Loop

        Do While Abs(documentos(arr(j)).ImporteCents) < pivot
            j = j - 1
        Loop

        If i <= j Then
            temp = arr(i)
            arr(i) = arr(j)
            arr(j) = temp
            i = i + 1
            j = j - 1
        End If
    Loop

    If first < j Then OrdenarCandidatosPorImporte arr, first, j
    If i < last Then OrdenarCandidatosPorImporte arr, i, last
End Sub

Private Function ObtenerOCrearHoja(ByVal sheetName As String, ByVal afterSheetName As String) As Worksheet
    On Error Resume Next
    Set ObtenerOCrearHoja = ThisWorkbook.Worksheets(sheetName)
    On Error GoTo 0

    If ObtenerOCrearHoja Is Nothing Then
        Set ObtenerOCrearHoja = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(afterSheetName))
        ObtenerOCrearHoja.Name = sheetName
    End If
End Function

Private Sub LimpiarMemoria()
    Erase documentos
    Erase liquidaciones
    Erase resultados
    Set indiceImportes = Nothing
    nDocs = 0
    nLiq = 0
End Sub
