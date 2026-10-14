Attribute VB_Name = "modConciliacionRobusta"
Option Explicit

' ================================================================
'  CONCILIACION ROBUSTA DE ENVASES Y CAJAS CONTRA SAP MB51
'  Ejecutar: ConciliarTodo
'  - No requiere activar referencias: usa Dictionary con late binding.
'  - Detecta automaticamente la fila de encabezados de cada fuente.
'  - Detecta y avisa duplicados en proveedor y MB51.
'  - Considera movimientos de reversa comunes de SAP.
'  - Genera Conc Envase, Conc Caja, Incidencias y Dashboard.
'  - Registra cada ejecucion en una hoja de auditoria oculta.
' ================================================================

Private Const FILA_DATOS_SALIDA As Long = 5

Private Const SH_MB51 As String = "MB51"
Private Const SH_PROV_ENVASE As String = "Proveedor Envase"
Private Const SH_PROV_CAJA As String = "Proveedor Caja"
Private Function SH_CONFIG() As String
    SH_CONFIG = "Configuraci" & ChrW(243) & "n"
End Function
Private Const SH_CONC_ENVASE As String = "Conc Envase"
Private Const SH_CONC_CAJA As String = "Conc Caja"
Private Const SH_INCIDENCIAS As String = "Incidencias"
Private Const SH_DASHBOARD As String = "Dashboard"
Private Const SH_LOG As String = "Log Conciliacion"

Private Enum eMB51
    mbConteo = 0
    mbPedidos = 1
    mbDocumentos = 2
    mbMateriales = 3
    mbTextos = 4
    mbCantidad = 5
    mbImporte = 6
    mbFecha = 7
    mbMaterialMultiple = 8
End Enum

Public Sub ConciliarTodo()
    Dim t0 As Single
    Dim calcAnterior As XlCalculation
    Dim eventosAnteriores As Boolean
    Dim pantallaAnterior As Boolean
    Dim alertasAnteriores As Boolean

    Dim wsMB51 As Worksheet
    Dim wsProvEnvase As Worksheet
    Dim wsProvCaja As Worksheet
    Dim wsConfig As Worksheet
    Dim wsConcEnvase As Worksheet
    Dim wsConcCaja As Worksheet
    Dim wsIncidencias As Worksheet
    Dim wsDashboard As Worksheet

    Dim idxMB51 As Object
    Dim conteoProveedor As Object
    Dim catalogo As Object
    Dim duplicadosMB51Afectados As Object

    Dim estEnvase As Variant
    Dim estCaja As Variant
    Dim toleranciaCantidad As Double
    Dim toleranciaFecha As Long
    Dim totalRegistros As Long
    Dim totalConciliados As Long
    Dim totalRevision As Long
    Dim filasDuplicadasProveedor As Long
    Dim referenciasDuplicadasProveedor As Long
    Dim segundos As Double
    Dim mensaje As String

    On Error GoTo ControlError

    t0 = Timer
    calcAnterior = Application.Calculation
    eventosAnteriores = Application.EnableEvents
    pantallaAnterior = Application.ScreenUpdating
    alertasAnteriores = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual
    Application.StatusBar = "Preparando conciliaci" & ChrW(243) & "n..."

    Set wsMB51 = BuscarHoja(Array(SH_MB51))
    Set wsProvEnvase = BuscarHoja(Array(SH_PROV_ENVASE, "Proveedor de Envase", "PROVEEDOR ENVASE"))
    Set wsProvCaja = BuscarHoja(Array(SH_PROV_CAJA, "Proveedor de Caja", "PROVEEDOR CAJA"))

    If wsMB51 Is Nothing Then Err.Raise vbObjectError + 1001, , "No se encontr" & ChrW(243) & " la hoja MB51."
    If wsProvEnvase Is Nothing Then Err.Raise vbObjectError + 1002, , "No se encontr" & ChrW(243) & " la hoja Proveedor Envase."
    If wsProvCaja Is Nothing Then Err.Raise vbObjectError + 1003, , "No se encontr" & ChrW(243) & " la hoja Proveedor Caja."

    Set wsConfig = BuscarHoja(Array(SH_CONFIG, "Configuracion"))
    If wsConfig Is Nothing Then Set wsConfig = CrearConfiguracionBase()

    Set wsConcEnvase = ObtenerOCrearHoja(SH_CONC_ENVASE)
    Set wsConcCaja = ObtenerOCrearHoja(SH_CONC_CAJA)
    Set wsIncidencias = ObtenerOCrearHoja(SH_INCIDENCIAS)
    Set wsDashboard = ObtenerOCrearHoja(SH_DASHBOARD)

    Application.StatusBar = "Indexando MB51 y validando duplicados..."
    Set idxMB51 = IndexarMB51(wsMB51)
    Set conteoProveedor = ConstruirConteoProveedor(wsProvEnvase, wsProvCaja)
    referenciasDuplicadasProveedor = ContarReferenciasDuplicadas(conteoProveedor)
    Set catalogo = ConstruirCatalogo(wsConfig)
    Set duplicadosMB51Afectados = NuevoDiccionario()

    LeerTolerancias wsConfig, toleranciaCantidad, toleranciaFecha

    PrepararAuxiliaresMB51 wsMB51, idxMB51
    PrepararAuxiliaresProveedor wsProvEnvase, 6, 13, 14, conteoProveedor
    PrepararAuxiliaresProveedor wsProvCaja, 8, 17, 18, conteoProveedor

    Application.StatusBar = "Conciliando envases..."
    estEnvase = ConciliarEnvase(wsProvEnvase, wsConcEnvase, idxMB51, conteoProveedor, _
                                catalogo, toleranciaCantidad, toleranciaFecha, _
                                duplicadosMB51Afectados)

    Application.StatusBar = "Conciliando cajas..."
    estCaja = ConciliarCaja(wsProvCaja, wsConcCaja, idxMB51, conteoProveedor, _
                            catalogo, toleranciaCantidad, toleranciaFecha, _
                            duplicadosMB51Afectados)

    Application.StatusBar = "Generando incidencias y dashboard..."
    GenerarIncidencias wsConcEnvase, wsConcCaja, wsIncidencias
    ActualizarDashboard wsDashboard, estEnvase, estCaja, _
                        referenciasDuplicadasProveedor, duplicadosMB51Afectados.Count

    totalRegistros = CLng(estEnvase(0)) + CLng(estCaja(0))
    totalConciliados = CLng(estEnvase(1)) + CLng(estCaja(1))
    totalRevision = totalRegistros - totalConciliados
    filasDuplicadasProveedor = CLng(estEnvase(3)) + CLng(estCaja(3))

    segundos = Timer - t0
    If segundos < 0 Then segundos = segundos + 86400#

    RegistrarLog totalRegistros, totalConciliados, totalRevision, _
                 referenciasDuplicadasProveedor, duplicadosMB51Afectados.Count, segundos

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFull

    If referenciasDuplicadasProveedor > 0 Or duplicadosMB51Afectados.Count > 0 Then
        mensaje = "Conciliaci" & ChrW(243) & "n terminada con alertas." & vbCrLf & vbCrLf & _
                  "Registros procesados: " & Format$(totalRegistros, "#,##0") & vbCrLf & _
                  "Conciliados: " & Format$(totalConciliados, "#,##0") & vbCrLf & _
                  "Por revisar: " & Format$(totalRevision, "#,##0") & vbCrLf & _
                  "Referencias duplicadas en proveedor: " & _
                  Format$(referenciasDuplicadasProveedor, "#,##0") & vbCrLf & _
                  "Filas involucradas: " & Format$(filasDuplicadasProveedor, "#,##0") & vbCrLf & _
                  "Referencias duplicadas en MB51: " & Format$(duplicadosMB51Afectados.Count, "#,##0") & vbCrLf & vbCrLf & _
                  "Consulta la hoja Incidencias."
        If referenciasDuplicadasProveedor > 0 Then
            mensaje = mensaje & vbCrLf & vbCrLf & "Duplicados: " & _
                      ResumenReferenciasDuplicadas(conteoProveedor, 12)
        End If
        MsgBox mensaje, vbExclamation + vbOKOnly, "Conciliaci" & ChrW(243) & "n con alertas"
    Else
        mensaje = "Conciliaci" & ChrW(243) & "n terminada." & vbCrLf & vbCrLf & _
                  "Registros procesados: " & Format$(totalRegistros, "#,##0") & vbCrLf & _
                  "Conciliados: " & Format$(totalConciliados, "#,##0") & vbCrLf & _
                  "Por revisar: " & Format$(totalRevision, "#,##0")
        MsgBox mensaje, vbInformation + vbOKOnly, "Conciliaci" & ChrW(243) & "n terminada"
    End If

SalidaSegura:
    On Error Resume Next
    Application.StatusBar = False
    Application.Calculation = calcAnterior
    Application.EnableEvents = eventosAnteriores
    Application.ScreenUpdating = pantallaAnterior
    Application.DisplayAlerts = alertasAnteriores
    On Error GoTo 0
    Exit Sub

ControlError:
    mensaje = "No se pudo completar la conciliaci" & ChrW(243) & "n." & vbCrLf & vbCrLf & _
              "Error " & Err.Number & ": " & Err.Description
    MsgBox mensaje, vbCritical + vbOKOnly, "Error de conciliaci" & ChrW(243) & "n"
    Resume SalidaSegura
End Sub

Private Function IndexarMB51(ByVal ws As Worksheet) As Object
    Dim dic As Object
    Dim ultima As Long
    Dim datos As Variant
    Dim i As Long
    Dim clave As String
    Dim pedido As String
    Dim documento As String
    Dim material As String
    Dim textoMaterial As String
    Dim movimiento As String
    Dim cantidad As Double
    Dim importe As Double
    Dim fechaSAP As Double
    Dim okCantidad As Boolean
    Dim okImporte As Boolean
    Dim registro As Variant
    Dim filaDatos As Long

    Set dic = NuevoDiccionario()
    filaDatos = DetectarPrimeraFilaDatos(ws, 3, Array("REFERENCIA", "REMISION"))
    ultima = UltimaFila(ws, 3)
    If ultima < filaDatos Then
        Set IndexarMB51 = dic
        Exit Function
    End If

    datos = ws.Range(ws.Cells(filaDatos, 1), ws.Cells(ultima, 10)).Value2

    For i = 1 To UBound(datos, 1)
        clave = NormalizarClave(datos(i, 3))
        If Len(clave) > 0 Then
            pedido = TextoSeguro(datos(i, 1))
            documento = TextoSeguro(datos(i, 2))
            material = NormalizarClave(datos(i, 4))
            textoMaterial = TextoSeguro(datos(i, 5))
            movimiento = NormalizarClave(datos(i, 8))
            cantidad = NumeroSeguro(datos(i, 6), okCantidad)
            importe = NumeroSeguro(datos(i, 7), okImporte)
            fechaSAP = FechaSerial(datos(i, 9))

            If Not okCantidad Then cantidad = 0
            If Not okImporte Then importe = 0
            cantidad = AplicarSignoMovimiento(cantidad, movimiento)
            importe = AplicarSignoMovimiento(importe, movimiento)

            If Not dic.Exists(clave) Then
                registro = Array(1&, pedido, documento, material, textoMaterial, _
                                 cantidad, importe, fechaSAP, False)
                dic.Add clave, registro
            Else
                registro = dic(clave)
                registro(mbConteo) = CLng(registro(mbConteo)) + 1
                registro(mbPedidos) = AgregarUnico(CStr(registro(mbPedidos)), pedido)
                registro(mbDocumentos) = AgregarUnico(CStr(registro(mbDocumentos)), documento)

                If Len(material) > 0 Then
                    If Len(CStr(registro(mbMateriales))) > 0 Then
                        If InStr(1, " | " & CStr(registro(mbMateriales)) & " | ", _
                                 " | " & material & " | ", vbTextCompare) = 0 Then
                            registro(mbMaterialMultiple) = True
                        End If
                    End If
                    registro(mbMateriales) = AgregarUnico(CStr(registro(mbMateriales)), material)
                End If

                registro(mbTextos) = AgregarUnico(CStr(registro(mbTextos)), textoMaterial)
                registro(mbCantidad) = CDbl(registro(mbCantidad)) + cantidad
                registro(mbImporte) = CDbl(registro(mbImporte)) + importe

                If fechaSAP > 0 Then
                    If CDbl(registro(mbFecha)) <= 0 Or fechaSAP > CDbl(registro(mbFecha)) Then
                        registro(mbFecha) = fechaSAP
                    End If
                End If

                dic(clave) = registro
            End If
        End If
    Next i

    Set IndexarMB51 = dic
End Function

Private Function ConstruirConteoProveedor(ByVal wsEnvase As Worksheet, _
                                          ByVal wsCaja As Worksheet) As Object
    Dim dic As Object
    Set dic = NuevoDiccionario()

    AcumularClavesProveedor wsEnvase, 6, dic
    AcumularClavesProveedor wsCaja, 8, dic

    Set ConstruirConteoProveedor = dic
End Function

Private Sub AcumularClavesProveedor(ByVal ws As Worksheet, ByVal columnaClave As Long, _
                                    ByVal dic As Object)
    Dim ultima As Long
    Dim valores As Variant
    Dim i As Long
    Dim clave As String
    Dim filaDatos As Long

    filaDatos = DetectarPrimeraFilaDatos(ws, columnaClave, _
                                         Array("REMISION", "NO DE OE", "NO. DE O.E."))
    ultima = UltimaFila(ws, columnaClave)
    If ultima < filaDatos Then Exit Sub

    valores = ws.Range(ws.Cells(filaDatos, columnaClave), ws.Cells(ultima, columnaClave)).Value2
    For i = 1 To UBound(valores, 1)
        clave = NormalizarClave(valores(i, 1))
        If Len(clave) > 0 Then
            If dic.Exists(clave) Then
                dic(clave) = CLng(dic(clave)) + 1
            Else
                dic.Add clave, 1&
            End If
        End If
    Next i
End Sub

Private Function ContarReferenciasDuplicadas(ByVal conteoProveedor As Object) As Long
    Dim clave As Variant

    For Each clave In conteoProveedor.Keys
        If CLng(conteoProveedor(clave)) > 1 Then
            ContarReferenciasDuplicadas = ContarReferenciasDuplicadas + 1
        End If
    Next clave
End Function

Private Function ResumenReferenciasDuplicadas(ByVal conteoProveedor As Object, _
                                               ByVal maximoMostrar As Long) As String
    Dim clave As Variant
    Dim total As Long
    Dim mostradas As Long
    Dim resumen As String

    total = ContarReferenciasDuplicadas(conteoProveedor)
    If total = 0 Then Exit Function
    If maximoMostrar < 1 Then maximoMostrar = 1

    For Each clave In conteoProveedor.Keys
        If CLng(conteoProveedor(clave)) > 1 Then
            mostradas = mostradas + 1
            If Len(resumen) > 0 Then resumen = resumen & ", "
            resumen = resumen & CStr(clave) & " (" & _
                      CStr(conteoProveedor(clave)) & " veces)"
            If mostradas >= maximoMostrar Then Exit For
        End If
    Next clave

    If total > mostradas Then
        resumen = resumen & " y " & (total - mostradas) & " m" & ChrW(225) & "s"
    End If
    ResumenReferenciasDuplicadas = resumen
End Function

Private Function ConstruirCatalogo(ByVal ws As Worksheet) As Object
    Dim dic As Object
    Dim ultima As Long
    Dim datos As Variant
    Dim i As Long
    Dim tipo As String
    Dim producto As String
    Dim clave As String
    Dim factor As Double
    Dim okFactor As Boolean
    Dim material As String
    Dim baseCantidad As String
    Dim registro As Variant

    Set dic = NuevoDiccionario()
    ultima = UltimaFila(ws, 3)
    If ultima < 13 Then
        Set ConstruirCatalogo = dic
        Exit Function
    End If

    datos = ws.Range(ws.Cells(13, 1), ws.Cells(ultima, 6)).Value2
    For i = 1 To UBound(datos, 1)
        tipo = NormalizarTexto(datos(i, 2))
        producto = NormalizarTexto(datos(i, 3))
        If Len(tipo) > 0 And Len(producto) > 0 Then
            clave = tipo & "|" & producto
            factor = NumeroSeguro(datos(i, 4), okFactor)
            material = NormalizarClave(datos(i, 5))
            baseCantidad = NormalizarTexto(datos(i, 6))

            If Not dic.Exists(clave) Then
                registro = Array(factor, material, baseCantidad, False, okFactor)
                dic.Add clave, registro
            Else
                registro = dic(clave)
                registro(3) = True
                dic(clave) = registro
            End If
        End If
    Next i

    Set ConstruirCatalogo = dic
End Function

Private Function ConciliarEnvase(ByVal wsFuente As Worksheet, ByVal wsDestino As Worksheet, _
                                 ByVal idxMB51 As Object, ByVal conteoProveedor As Object, _
                                 ByVal catalogo As Object, ByVal toleranciaCantidad As Double, _
                                 ByVal toleranciaFecha As Long, _
                                 ByVal duplicadosMB51Afectados As Object) As Variant
    Dim ultima As Long
    Dim datos As Variant
    Dim salida() As Variant
    Dim filasSalida As Long
    Dim i As Long
    Dim j As Long
    Dim clave As String
    Dim producto As String
    Dim cfgClave As String
    Dim cfg As Variant
    Dim configExiste As Boolean
    Dim configDuplicada As Boolean
    Dim factorValido As Boolean
    Dim factor As Double
    Dim materialEsperado As String
    Dim baseCantidad As String
    Dim cajas As Double
    Dim tarimas As Double
    Dim okCajas As Boolean
    Dim okTarimas As Boolean
    Dim cantidadEsperada As Double
    Dim rec As Variant
    Dim movimientosMB51 As Long
    Dim cantidadSAP As Double
    Dim importeSAP As Double
    Dim fechaProveedor As Double
    Dim fechaSAP As Double
    Dim diferenciaCantidad As Variant
    Dim diferenciaDias As Variant
    Dim estado As String
    Dim riesgo As String
    Dim observacion As String
    Dim repeticionesProveedor As Long
    Dim conciliados As Long
    Dim duplicadosProveedor As Long
    Dim costoProveedor As Double
    Dim valorSAP As Double
    Dim okCosto As Boolean
    Dim costoFila As Double
    Dim filaDatos As Long

    PrepararHojaConciliacion wsDestino, True
    filaDatos = DetectarPrimeraFilaDatos(wsFuente, 6, Array("REMISION"))
    ultima = UltimaFila(wsFuente, 6)

    If ultima < filaDatos Then
        LimpiarResultado wsDestino
        ConciliarEnvase = Array(0&, 0&, 0&, 0&, 0#, 0#)
        Exit Function
    End If

    datos = wsFuente.Range(wsFuente.Cells(filaDatos, 1), wsFuente.Cells(ultima, 12)).Value2
    filasSalida = ContarClaves(datos, 6)
    If filasSalida = 0 Then
        LimpiarResultado wsDestino
        ConciliarEnvase = Array(0&, 0&, 0&, 0&, 0#, 0#)
        Exit Function
    End If

    ReDim salida(1 To filasSalida, 1 To 25)
    j = 0

    For i = 1 To UBound(datos, 1)
        clave = NormalizarClave(datos(i, 6))
        If Len(clave) > 0 Then
            j = j + 1
            producto = NormalizarTexto(datos(i, 9))
            cfgClave = "ENVASE|" & producto
            LeerConfiguracionProducto catalogo, cfgClave, configExiste, configDuplicada, _
                                       factorValido, factor, materialEsperado, baseCantidad

            cajas = NumeroSeguro(datos(i, 7), okCajas)
            tarimas = NumeroSeguro(datos(i, 8), okTarimas)
            If Not okCajas Then cajas = 0
            If Not okTarimas Then tarimas = 0
            cantidadEsperada = CantidadEsperadaProveedor("ENVASE", baseCantidad, factor, _
                                                          cajas, tarimas, factorValido)

            InicializarDatosMB51 movimientosMB51, cantidadSAP, importeSAP, fechaSAP, rec
            If idxMB51.Exists(clave) Then
                rec = idxMB51(clave)
                movimientosMB51 = CLng(rec(mbConteo))
                cantidadSAP = CDbl(rec(mbCantidad))
                importeSAP = CDbl(rec(mbImporte))
                fechaSAP = CDbl(rec(mbFecha))
                If movimientosMB51 > 1 Then AgregarClave duplicadosMB51Afectados, clave
            End If

            repeticionesProveedor = CLng(conteoProveedor(clave))
            fechaProveedor = FechaSerial(datos(i, 11))

            If movimientosMB51 > 0 And factorValido Then
                diferenciaCantidad = cantidadSAP - cantidadEsperada
            Else
                diferenciaCantidad = vbNullString
            End If

            If fechaProveedor > 0 And fechaSAP > 0 Then
                diferenciaDias = CLng(fechaSAP) - CLng(fechaProveedor)
            Else
                diferenciaDias = vbNullString
            End If

            EvaluarEstado repeticionesProveedor, movimientosMB51, configExiste, _
                          configDuplicada, factorValido, materialEsperado, _
                          MaterialRegistro(rec, movimientosMB51), cantidadEsperada, _
                          cantidadSAP, fechaProveedor, fechaSAP, toleranciaCantidad, _
                          toleranciaFecha, estado, riesgo, observacion

            salida(j, 1) = clave
            salida(j, 2) = datos(i, 2)
            salida(j, 3) = datos(i, 1)
            salida(j, 4) = datos(i, 5)
            salida(j, 5) = ValorFechaSalida(fechaProveedor)
            salida(j, 6) = datos(i, 9)
            salida(j, 7) = cajas
            salida(j, 8) = tarimas
            If factorValido Then salida(j, 9) = factor Else salida(j, 9) = vbNullString
            If factorValido Then salida(j, 10) = cantidadEsperada Else salida(j, 10) = vbNullString
            salida(j, 11) = movimientosMB51
            salida(j, 12) = CampoRegistro(rec, movimientosMB51, mbPedidos)
            salida(j, 13) = CampoRegistro(rec, movimientosMB51, mbDocumentos)
            salida(j, 14) = materialEsperado
            salida(j, 15) = CampoRegistro(rec, movimientosMB51, mbMateriales)
            salida(j, 16) = CampoRegistro(rec, movimientosMB51, mbTextos)
            If movimientosMB51 > 0 Then salida(j, 17) = cantidadSAP Else salida(j, 17) = vbNullString
            salida(j, 18) = diferenciaCantidad
            salida(j, 19) = ValorFechaSalida(fechaSAP)
            salida(j, 20) = diferenciaDias
            salida(j, 21) = datos(i, 10)
            If movimientosMB51 > 0 Then salida(j, 22) = importeSAP Else salida(j, 22) = vbNullString
            salida(j, 23) = estado
            salida(j, 24) = riesgo
            salida(j, 25) = observacion

            If estado = "CONCILIADO" Then conciliados = conciliados + 1
            If repeticionesProveedor > 1 Then duplicadosProveedor = duplicadosProveedor + 1

            costoFila = NumeroSeguro(datos(i, 10), okCosto)
            If okCosto Then costoProveedor = costoProveedor + costoFila
            If movimientosMB51 > 0 Then valorSAP = valorSAP + importeSAP
        End If
    Next i

    EscribirResultado wsDestino, salida, filasSalida, True
    ConciliarEnvase = Array(filasSalida, conciliados, filasSalida - conciliados, _
                            duplicadosProveedor, costoProveedor, valorSAP)
End Function

Private Function ConciliarCaja(ByVal wsFuente As Worksheet, ByVal wsDestino As Worksheet, _
                               ByVal idxMB51 As Object, ByVal conteoProveedor As Object, _
                               ByVal catalogo As Object, ByVal toleranciaCantidad As Double, _
                               ByVal toleranciaFecha As Long, _
                               ByVal duplicadosMB51Afectados As Object) As Variant
    Dim ultima As Long
    Dim datos As Variant
    Dim salida() As Variant
    Dim filasSalida As Long
    Dim i As Long
    Dim j As Long
    Dim clave As String
    Dim producto As String
    Dim cfgClave As String
    Dim cfg As Variant
    Dim configExiste As Boolean
    Dim configDuplicada As Boolean
    Dim factorValido As Boolean
    Dim factor As Double
    Dim materialEsperado As String
    Dim baseCantidad As String
    Dim cantidadProveedor As Double
    Dim okCantidadProveedor As Boolean
    Dim cantidadEsperada As Double
    Dim rec As Variant
    Dim movimientosMB51 As Long
    Dim cantidadSAP As Double
    Dim importeSAP As Double
    Dim fechaProveedor As Double
    Dim fechaSAP As Double
    Dim diferenciaCantidad As Variant
    Dim diferenciaDias As Variant
    Dim estado As String
    Dim riesgo As String
    Dim observacion As String
    Dim repeticionesProveedor As Long
    Dim conciliados As Long
    Dim duplicadosProveedor As Long
    Dim costoProveedor As Double
    Dim valorSAP As Double
    Dim okCosto As Boolean
    Dim costoFila As Double
    Dim filaDatos As Long

    PrepararHojaConciliacion wsDestino, False
    filaDatos = DetectarPrimeraFilaDatos(wsFuente, 8, Array("NO DE OE", "NO. DE O.E."))
    ultima = UltimaFila(wsFuente, 8)

    If ultima < filaDatos Then
        LimpiarResultado wsDestino
        ConciliarCaja = Array(0&, 0&, 0&, 0&, 0#, 0#)
        Exit Function
    End If

    datos = wsFuente.Range(wsFuente.Cells(filaDatos, 1), wsFuente.Cells(ultima, 16)).Value2
    filasSalida = ContarClaves(datos, 8)
    If filasSalida = 0 Then
        LimpiarResultado wsDestino
        ConciliarCaja = Array(0&, 0&, 0&, 0&, 0#, 0#)
        Exit Function
    End If

    ReDim salida(1 To filasSalida, 1 To 25)
    j = 0

    For i = 1 To UBound(datos, 1)
        clave = NormalizarClave(datos(i, 8))
        If Len(clave) > 0 Then
            j = j + 1
            producto = NormalizarTexto(datos(i, 9))
            cfgClave = "CAJA|" & producto
            LeerConfiguracionProducto catalogo, cfgClave, configExiste, configDuplicada, _
                                       factorValido, factor, materialEsperado, baseCantidad

            cantidadProveedor = NumeroSeguro(datos(i, 10), okCantidadProveedor)
            If Not okCantidadProveedor Then cantidadProveedor = 0
            cantidadEsperada = CantidadEsperadaProveedor("CAJA", baseCantidad, factor, _
                                                          cantidadProveedor, cantidadProveedor, _
                                                          factorValido)

            InicializarDatosMB51 movimientosMB51, cantidadSAP, importeSAP, fechaSAP, rec
            If idxMB51.Exists(clave) Then
                rec = idxMB51(clave)
                movimientosMB51 = CLng(rec(mbConteo))
                cantidadSAP = CDbl(rec(mbCantidad))
                importeSAP = CDbl(rec(mbImporte))
                fechaSAP = CDbl(rec(mbFecha))
                If movimientosMB51 > 1 Then AgregarClave duplicadosMB51Afectados, clave
            End If

            repeticionesProveedor = CLng(conteoProveedor(clave))
            fechaProveedor = FechaSerial(datos(i, 4))

            If movimientosMB51 > 0 And factorValido Then
                diferenciaCantidad = cantidadSAP - cantidadEsperada
            Else
                diferenciaCantidad = vbNullString
            End If

            If fechaProveedor > 0 And fechaSAP > 0 Then
                diferenciaDias = CLng(fechaSAP) - CLng(fechaProveedor)
            Else
                diferenciaDias = vbNullString
            End If

            EvaluarEstado repeticionesProveedor, movimientosMB51, configExiste, _
                          configDuplicada, factorValido, materialEsperado, _
                          MaterialRegistro(rec, movimientosMB51), cantidadEsperada, _
                          cantidadSAP, fechaProveedor, fechaSAP, toleranciaCantidad, _
                          toleranciaFecha, estado, riesgo, observacion

            salida(j, 1) = clave
            salida(j, 2) = datos(i, 2)
            salida(j, 3) = ValorFechaSalida(FechaSerial(datos(i, 3)))
            salida(j, 4) = ValorFechaSalida(fechaProveedor)
            salida(j, 5) = datos(i, 5)
            salida(j, 6) = datos(i, 6)
            salida(j, 7) = datos(i, 9)
            salida(j, 8) = cantidadProveedor
            If factorValido Then salida(j, 9) = factor Else salida(j, 9) = vbNullString
            If factorValido Then salida(j, 10) = cantidadEsperada Else salida(j, 10) = vbNullString
            salida(j, 11) = movimientosMB51
            salida(j, 12) = CampoRegistro(rec, movimientosMB51, mbPedidos)
            salida(j, 13) = CampoRegistro(rec, movimientosMB51, mbDocumentos)
            salida(j, 14) = materialEsperado
            salida(j, 15) = CampoRegistro(rec, movimientosMB51, mbMateriales)
            salida(j, 16) = CampoRegistro(rec, movimientosMB51, mbTextos)
            If movimientosMB51 > 0 Then salida(j, 17) = cantidadSAP Else salida(j, 17) = vbNullString
            salida(j, 18) = diferenciaCantidad
            salida(j, 19) = ValorFechaSalida(fechaSAP)
            salida(j, 20) = diferenciaDias
            salida(j, 21) = datos(i, 15)
            If movimientosMB51 > 0 Then salida(j, 22) = importeSAP Else salida(j, 22) = vbNullString
            salida(j, 23) = estado
            salida(j, 24) = riesgo
            salida(j, 25) = observacion

            If estado = "CONCILIADO" Then conciliados = conciliados + 1
            If repeticionesProveedor > 1 Then duplicadosProveedor = duplicadosProveedor + 1

            costoFila = NumeroSeguro(datos(i, 15), okCosto)
            If okCosto Then costoProveedor = costoProveedor + costoFila
            If movimientosMB51 > 0 Then valorSAP = valorSAP + importeSAP
        End If
    Next i

    EscribirResultado wsDestino, salida, filasSalida, False
    ConciliarCaja = Array(filasSalida, conciliados, filasSalida - conciliados, _
                          duplicadosProveedor, costoProveedor, valorSAP)
End Function

Private Sub EvaluarEstado(ByVal repeticionesProveedor As Long, ByVal movimientosMB51 As Long, _
                          ByVal configExiste As Boolean, ByVal configDuplicada As Boolean, _
                          ByVal factorValido As Boolean, ByVal materialEsperado As String, _
                          ByVal materialSAP As String, ByVal cantidadEsperada As Double, _
                          ByVal cantidadSAP As Double, ByVal fechaProveedor As Double, _
                          ByVal fechaSAP As Double, ByVal toleranciaCantidad As Double, _
                          ByVal toleranciaFecha As Long, ByRef estado As String, _
                          ByRef riesgo As String, ByRef observacion As String)

    If repeticionesProveedor > 1 And movimientosMB51 > 1 Then
        estado = "DUPLICADO PROVEEDOR Y MB51"
        riesgo = "ALTO"
        observacion = "Doble alerta: la referencia aparece " & repeticionesProveedor & _
                      " veces en proveedor y " & movimientosMB51 & " veces en MB51."
    ElseIf repeticionesProveedor > 1 Then
        estado = "DUPLICADO EN PROVEEDOR"
        riesgo = "ALTO"
        observacion = "La referencia aparece " & repeticionesProveedor & _
                      " veces en los reportes de proveedor."
    ElseIf configDuplicada Then
        estado = "CONFIGURACI" & ChrW(211) & "N DUPLICADA"
        riesgo = "ALTO"
        observacion = "El producto tiene m" & ChrW(225) & "s de una regla en Configuraci" & ChrW(243) & "n."
    ElseIf Not configExiste Or Not factorValido Or Len(materialEsperado) = 0 Then
        estado = "SIN FACTOR"
        riesgo = "ALTO"
        observacion = "Falta una regla " & ChrW(250) & "nica y completa de conversi" & ChrW(243) & "n/material esperado."
    ElseIf movimientosMB51 = 0 Then
        estado = "PENDIENTE SAP"
        riesgo = "ALTO"
        observacion = "La referencia no aparece en MB51."
    ElseIf movimientosMB51 > 1 Then
        estado = "DUPLICADO EN MB51"
        riesgo = "ALTO"
        observacion = "La referencia aparece " & movimientosMB51 & _
                      " veces en MB51; revisar documentos, reversas y cantidades agregadas."
    ElseIf NormalizarClave(materialEsperado) <> NormalizarClave(materialSAP) Then
        estado = "DIFERENCIA MATERIAL"
        riesgo = "MEDIO"
        observacion = "Material esperado " & materialEsperado & _
                      "; material SAP " & materialSAP & "."
    ElseIf Abs(cantidadSAP - cantidadEsperada) > toleranciaCantidad Then
        estado = "DIFERENCIA CANTIDAD"
        riesgo = "MEDIO"
        observacion = "La diferencia de cantidad excede la tolerancia de " & _
                      Format$(toleranciaCantidad, "#,##0.00") & "."
    ElseIf fechaProveedor <= 0 Or fechaSAP <= 0 Then
        estado = "FECHA INV" & ChrW(193) & "LIDA"
        riesgo = "ALTO"
        observacion = "Falta una fecha v" & ChrW(225) & "lida en proveedor o MB51."
    ElseIf Abs(CLng(fechaSAP) - CLng(fechaProveedor)) > toleranciaFecha Then
        estado = "DIFERENCIA FECHA"
        riesgo = "MEDIO"
        observacion = "La diferencia de fecha excede la tolerancia de " & _
                      toleranciaFecha & " d" & ChrW(237) & "a(s)."
    Else
        estado = "CONCILIADO"
        riesgo = "BAJO"
        observacion = "Referencia, material, cantidad y fecha conciliados."
    End If
End Sub

Private Sub GenerarIncidencias(ByVal wsEnvase As Worksheet, ByVal wsCaja As Worksheet, _
                               ByVal wsDestino As Worksheet)
    Dim total As Long
    Dim datos() As Variant
    Dim fila As Long
    Dim ultimaLimpiar As Long

    PrepararHojaIncidencias wsDestino
    total = ContarIncidencias(wsEnvase) + ContarIncidencias(wsCaja)

    ultimaLimpiar = MaxLong(UltimaFila(wsDestino, 1), 154)
    wsDestino.Range(wsDestino.Cells(FILA_DATOS_SALIDA, 1), wsDestino.Cells(ultimaLimpiar, 14)).ClearContents
    wsDestino.Range(wsDestino.Cells(FILA_DATOS_SALIDA, 12), wsDestino.Cells(ultimaLimpiar, 13)).Interior.Pattern = xlNone
    wsDestino.Range(wsDestino.Cells(FILA_DATOS_SALIDA, 12), wsDestino.Cells(ultimaLimpiar, 13)).Font.ColorIndex = xlAutomatic

    If total = 0 Then
        wsDestino.Cells(FILA_DATOS_SALIDA, 1).Value = "Sin incidencias"
        Exit Sub
    End If

    ReDim datos(1 To total, 1 To 14)
    fila = 0
    AgregarIncidenciasHoja wsEnvase, "Envase", datos, fila
    AgregarIncidenciasHoja wsCaja, "Caja", datos, fila

    wsDestino.Range("A5").Resize(total, 14).Value2 = datos
    wsDestino.Range("E5:G" & 4 + total).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    wsDestino.Range("H5:I" & 4 + total).NumberFormat = "dd-mmm-yyyy"
    wsDestino.Range("J5:K" & 4 + total).NumberFormat = "0;[Red]-0"
    wsDestino.Range("N5:N" & 4 + total).WrapText = True

    FormatearAlertas wsDestino, 5, 4 + total, 12, 13
End Sub

Private Function ContarIncidencias(ByVal ws As Worksheet) As Long
    Dim ultima As Long
    Dim r As Long
    Dim estado As String

    ultima = UltimaFila(ws, 1)
    For r = FILA_DATOS_SALIDA To ultima
        estado = UCase$(Trim$(CStr(ws.Cells(r, 23).Value2)))
        If Len(NormalizarClave(ws.Cells(r, 1).Value2)) > 0 And _
           Len(estado) > 0 And estado <> "CONCILIADO" Then
            ContarIncidencias = ContarIncidencias + 1
        End If
    Next r
End Function

Private Sub AgregarIncidenciasHoja(ByVal ws As Worksheet, ByVal fuente As String, _
                                   ByRef datos() As Variant, ByRef fila As Long)
    Dim ultima As Long
    Dim r As Long
    Dim estado As String
    Dim esEnvase As Boolean

    esEnvase = (UCase$(fuente) = "ENVASE")
    ultima = UltimaFila(ws, 1)

    For r = FILA_DATOS_SALIDA To ultima
        estado = UCase$(Trim$(CStr(ws.Cells(r, 23).Value2)))
        If Len(NormalizarClave(ws.Cells(r, 1).Value2)) > 0 And _
           Len(estado) > 0 And estado <> "CONCILIADO" Then
            fila = fila + 1
            datos(fila, 1) = fuente
            datos(fila, 2) = ws.Cells(r, 1).Value2
            datos(fila, 3) = ws.Cells(r, 2).Value2
            If esEnvase Then
                datos(fila, 4) = ws.Cells(r, 6).Value2
                datos(fila, 8) = ws.Cells(r, 5).Value2
            Else
                datos(fila, 4) = ws.Cells(r, 7).Value2
                datos(fila, 8) = ws.Cells(r, 4).Value2
            End If
            datos(fila, 5) = ws.Cells(r, 10).Value2
            datos(fila, 6) = ws.Cells(r, 17).Value2
            datos(fila, 7) = ws.Cells(r, 18).Value2
            datos(fila, 9) = ws.Cells(r, 19).Value2
            datos(fila, 10) = ws.Cells(r, 20).Value2
            datos(fila, 11) = ws.Cells(r, 11).Value2
            datos(fila, 12) = ws.Cells(r, 23).Value2
            datos(fila, 13) = ws.Cells(r, 24).Value2
            datos(fila, 14) = ws.Cells(r, 25).Value2
        End If
    Next r
End Sub

Private Sub ActualizarDashboard(ByVal ws As Worksheet, ByVal estEnvase As Variant, _
                                ByVal estCaja As Variant, _
                                ByVal referenciasDuplicadasProveedor As Long, _
                                ByVal duplicadosMB51 As Long)
    Dim totalRegistros As Long
    Dim totalConciliados As Long
    Dim totalRevision As Long
    Dim duplicadosProveedor As Long
    Dim costoTotal As Double
    Dim valorSAPTotal As Double
    Dim porcentaje As Double

    totalRegistros = CLng(estEnvase(0)) + CLng(estCaja(0))
    totalConciliados = CLng(estEnvase(1)) + CLng(estCaja(1))
    totalRevision = totalRegistros - totalConciliados
    duplicadosProveedor = CLng(estEnvase(3)) + CLng(estCaja(3))
    costoTotal = CDbl(estEnvase(4)) + CDbl(estCaja(4))
    valorSAPTotal = CDbl(estEnvase(5)) + CDbl(estCaja(5))
    If totalRegistros > 0 Then porcentaje = totalConciliados / totalRegistros

    ws.Range("A4:M23").UnMerge
    If Len(CStr(ws.Range("A1").Value2)) = 0 Then ws.Range("A1").Value = "Dashboard de conciliaci" & ChrW(243) & "n"

    ws.Range("A4").Value = "REGISTROS PROVEEDOR"
    ws.Range("D4").Value = "CONCILIADOS"
    ws.Range("G4").Value = "PENDIENTE / REVISAR"
    ws.Range("J4").Value = "% CONCILIACI" & ChrW(211) & "N"
    ws.Range("A5").Value = totalRegistros
    ws.Range("D5").Value = totalConciliados
    ws.Range("G5").Value = totalRevision
    ws.Range("J5").Value = porcentaje
    ws.Range("J5").NumberFormat = "0.0%"

    EscribirMatriz ws.Range("A10"), Array( _
        Array("Fuente", "Conciliados", "Por revisar"), _
        Array("Envase", estEnvase(1), estEnvase(2)), _
        Array("Caja", estCaja(1), estCaja(2)))

    EscribirMatriz ws.Range("A16"), Array( _
        Array("Fuente", "Registros", "% conciliado", "Costo log" & ChrW(237) & "stico proveedor", "Valoraci" & ChrW(243) & "n SAP", "Interpretaci" & ChrW(243) & "n"), _
        Array("Envase", estEnvase(0), PorcentajeSeguro(estEnvase(1), estEnvase(0)), estEnvase(4), estEnvase(5), "CXC vs valor material: no comparar"), _
        Array("Caja", estCaja(0), PorcentajeSeguro(estCaja(1), estCaja(0)), estCaja(4), estCaja(5), "Subtotal vs valor material: no comparar"), _
        Array("Total", totalRegistros, porcentaje, costoTotal, valorSAPTotal, "Totales informativos de distinta naturaleza"))
    ws.Range("C17:C19").NumberFormat = "0.0%"
    ws.Range("D17:E19").NumberFormat = "$#,##0.00;[Red]-$#,##0.00"

    ws.Range("A21").Value = "REFERENCIAS DUPLICADAS EN PROVEEDOR"
    ws.Range("D21").Value = referenciasDuplicadasProveedor
    ws.Range("G21").Value = "REFERENCIAS CON >1 MOVIMIENTO EN MB51"
    ws.Range("K21").Value = duplicadosMB51
    ws.Range("A22").Value = "Lectura del corte actual"

    If referenciasDuplicadasProveedor > 0 Or duplicadosMB51 > 0 Then
        ws.Range("A23").Value = "ALERTA: " & referenciasDuplicadasProveedor & _
                                " referencia(s) duplicada(s) en proveedor (" & _
                                duplicadosProveedor & " fila(s)) y " & duplicadosMB51 & _
                                " referencia(s) duplicada(s) en MB51. Consulta Incidencias."
    Else
        ws.Range("A23").Value = "Hay " & totalRevision & _
                                " registro(s) por revisar. Consulta Incidencias."
    End If

    ws.Range("D21").Interior.Color = RGB(253, 226, 226)
    ws.Range("K21").Interior.Color = RGB(253, 226, 226)
    ws.Range("D21").Font.Color = RGB(164, 22, 43)
    ws.Range("K21").Font.Color = RGB(164, 22, 43)
    ws.Range("D21").Font.Bold = True
    ws.Range("K21").Font.Bold = True
End Sub

Private Sub PrepararAuxiliaresMB51(ByVal ws As Worksheet, ByVal idxMB51 As Object)
    Dim ultima As Long
    Dim ultimaLimpiar As Long
    Dim r As Long
    Dim clave As String
    Dim registro As Variant
    Dim repeticiones As Long
    Dim filaDatos As Long
    Dim filaEncabezado As Long

    filaDatos = DetectarPrimeraFilaDatos(ws, 3, Array("REFERENCIA", "REMISION"))
    filaEncabezado = filaDatos - 1
    ultima = UltimaFila(ws, 3)
    ultimaLimpiar = MaxLong(ultima, MaxLong(UltimaFila(ws, 11), UltimaFila(ws, 12)))

    ws.Cells(filaEncabezado, 11).Value = "CLAVE NORMALIZADA"
    ws.Cells(filaEncabezado, 12).Value = "VECES EN MB51"
    EstiloEncabezadoAuxiliar ws.Range(ws.Cells(filaEncabezado, 11), ws.Cells(filaEncabezado, 12))

    If ultimaLimpiar >= filaDatos Then
        ws.Range(ws.Cells(filaDatos, 11), ws.Cells(ultimaLimpiar, 12)).ClearContents
        ws.Range(ws.Cells(filaDatos, 12), ws.Cells(ultimaLimpiar, 12)).Interior.Color = RGB(242, 244, 246)
        ws.Range(ws.Cells(filaDatos, 3), ws.Cells(ultimaLimpiar, 3)).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(filaDatos, 3), ws.Cells(ultimaLimpiar, 3)).Font.Bold = False
    End If

    For r = filaDatos To ultima
        clave = NormalizarClave(ws.Cells(r, 3).Value2)
        ws.Cells(r, 11).Value = clave
        If Len(clave) > 0 And idxMB51.Exists(clave) Then
            registro = idxMB51(clave)
            repeticiones = CLng(registro(mbConteo))
            ws.Cells(r, 12).Value = repeticiones
            If repeticiones > 1 Then
                MarcarDuplicado ws.Cells(r, 3), ws.Cells(r, 12)
            End If
        End If
    Next r
End Sub

Private Sub PrepararAuxiliaresProveedor(ByVal ws As Worksheet, ByVal columnaClave As Long, _
                                        ByVal columnaNormalizada As Long, ByVal columnaConteo As Long, _
                                        ByVal conteoProveedor As Object)
    Dim ultima As Long
    Dim ultimaLimpiar As Long
    Dim r As Long
    Dim clave As String
    Dim repeticiones As Long
    Dim filaDatos As Long
    Dim filaEncabezado As Long

    filaDatos = DetectarPrimeraFilaDatos(ws, columnaClave, _
                                         Array("REMISION", "NO DE OE", "NO. DE O.E."))
    filaEncabezado = filaDatos - 1
    ultima = UltimaFila(ws, columnaClave)
    ultimaLimpiar = MaxLong(ultima, MaxLong(UltimaFila(ws, columnaNormalizada), _
                                            UltimaFila(ws, columnaConteo)))

    ws.Cells(filaEncabezado, columnaNormalizada).Value = "CLAVE NORMALIZADA"
    ws.Cells(filaEncabezado, columnaConteo).Value = "VECES EN PROVEEDOR"
    EstiloEncabezadoAuxiliar ws.Range(ws.Cells(filaEncabezado, columnaNormalizada), _
                                      ws.Cells(filaEncabezado, columnaConteo))

    If ultimaLimpiar >= filaDatos Then
        ws.Range(ws.Cells(filaDatos, columnaNormalizada), _
                 ws.Cells(ultimaLimpiar, columnaConteo)).ClearContents
        ws.Range(ws.Cells(filaDatos, columnaConteo), _
                 ws.Cells(ultimaLimpiar, columnaConteo)).Interior.Color = RGB(242, 244, 246)
        ws.Range(ws.Cells(filaDatos, columnaClave), _
                 ws.Cells(ultimaLimpiar, columnaClave)).Font.ColorIndex = xlAutomatic
        ws.Range(ws.Cells(filaDatos, columnaClave), _
                 ws.Cells(ultimaLimpiar, columnaClave)).Font.Bold = False
    End If

    For r = filaDatos To ultima
        clave = NormalizarClave(ws.Cells(r, columnaClave).Value2)
        ws.Cells(r, columnaNormalizada).Value = clave
        If Len(clave) > 0 And conteoProveedor.Exists(clave) Then
            repeticiones = CLng(conteoProveedor(clave))
            ws.Cells(r, columnaConteo).Value = repeticiones
            If repeticiones > 1 Then
                MarcarDuplicado ws.Cells(r, columnaClave), ws.Cells(r, columnaConteo)
            End If
        End If
    Next r
End Sub

Private Sub MarcarDuplicado(ByVal celdaClave As Range, ByVal celdaConteo As Range)
    celdaClave.Font.Color = RGB(164, 22, 43)
    celdaClave.Font.Bold = True
    celdaConteo.Interior.Color = RGB(255, 199, 206)
    celdaConteo.Font.Color = RGB(156, 0, 6)
    celdaConteo.Font.Bold = True
End Sub

Private Sub EstiloEncabezadoAuxiliar(ByVal rango As Range)
    With rango
        .Interior.Color = RGB(38, 50, 56)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With
End Sub

Private Sub PrepararHojaConciliacion(ByVal ws As Worksheet, ByVal esEnvase As Boolean)
    Dim encabezados As Variant

    If esEnvase Then
        encabezados = Array("Remisi" & ChrW(243) & "n", "Pedido proveedor", "Planta", "Transportista", _
                            "Fecha proveedor", "Insumo", "Cajas", "Tarimas", "Factor", _
                            "Cant. esperada", "Coincidencias SAP", "Pedido SAP", _
                            "Doc material SAP", "Material esperado", "Material SAP", _
                            "Texto material SAP", "Cantidad SAP neta", "Dif. cantidad", _
                            "Fecha SAP", "Dif. d" & ChrW(237) & "as", "Total CXC proveedor", "Importe SAP", _
                            "Estatus", "Riesgo", "Observaci" & ChrW(243) & "n")
        If Len(CStr(ws.Range("A1").Value2)) = 0 Then ws.Range("A1").Value = "Conciliaci" & ChrW(243) & "n " & ChrW(183) & " Envase"
    Else
        encabezados = Array("No. O.E.", "Bit" & ChrW(225) & "cora", "Fecha carga", "Fecha entrega", _
                            "Origen", "Destino", "Producto", "Cant. proveedor", "Factor", _
                            "Cant. esperada", "Coincidencias SAP", "Pedido SAP", _
                            "Doc material SAP", "Material esperado", "Material SAP", _
                            "Texto material SAP", "Cantidad SAP neta", "Dif. cantidad", _
                            "Fecha SAP", "Dif. d" & ChrW(237) & "as", "Subtotal proveedor", "Importe SAP", _
                            "Estatus", "Riesgo", "Observaci" & ChrW(243) & "n")
        If Len(CStr(ws.Range("A1").Value2)) = 0 Then ws.Range("A1").Value = "Conciliaci" & ChrW(243) & "n " & ChrW(183) & " Caja"
    End If

    EscribirMatriz ws.Range("A4"), Array(encabezados)
    With ws.Range("A4:Y4")
        .Interior.Color = RGB(38, 50, 56)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Range("A1:Y2").Interior.Color = RGB(176, 24, 50)
    ws.Range("A1").Font.Color = vbWhite
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
End Sub

Private Sub PrepararHojaIncidencias(ByVal ws As Worksheet)
    Dim encabezados As Variant

    encabezados = Array("Fuente", "Folio", "Pedido / Bit" & ChrW(225) & "cora", "Producto / Insumo", _
                        "Cant. esperada", "Cant. SAP", "Dif. cantidad", "Fecha proveedor", _
                        "Fecha SAP", "Dif. d" & ChrW(237) & "as", "Coincidencias", "Estatus", "Riesgo", _
                        "Observaci" & ChrW(243) & "n")

    If Len(CStr(ws.Range("A1").Value2)) = 0 Then ws.Range("A1").Value = "Incidencias activas"
    ws.Range("A3").Value = "Incluye pendientes, diferencias y duplicados de proveedor o MB51."
    EscribirMatriz ws.Range("A4"), Array(encabezados)

    With ws.Range("A4:N4")
        .Interior.Color = RGB(38, 50, 56)
        .Font.Color = vbWhite
        .Font.Bold = True
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .WrapText = True
    End With

    ws.Range("A1:N2").Interior.Color = RGB(176, 24, 50)
    ws.Range("A1").Font.Color = vbWhite
    ws.Range("A1").Font.Bold = True
    ws.Range("A1").Font.Size = 16
End Sub

Private Sub EscribirResultado(ByVal ws As Worksheet, ByRef datos() As Variant, _
                              ByVal numeroFilas As Long, ByVal esEnvase As Boolean)
    Dim ultimaLimpiar As Long

    ultimaLimpiar = MaxLong(UltimaFila(ws, 1), 154)
    ws.Range(ws.Cells(FILA_DATOS_SALIDA, 1), ws.Cells(ultimaLimpiar, 25)).ClearContents
    ws.Range(ws.Cells(FILA_DATOS_SALIDA, 23), ws.Cells(ultimaLimpiar, 24)).Interior.Pattern = xlNone
    ws.Range(ws.Cells(FILA_DATOS_SALIDA, 23), ws.Cells(ultimaLimpiar, 24)).Font.ColorIndex = xlAutomatic
    ws.Range(ws.Cells(FILA_DATOS_SALIDA, 23), ws.Cells(ultimaLimpiar, 24)).Font.Bold = False

    If numeroFilas <= 0 Then Exit Sub

    ws.Range("A5").Resize(numeroFilas, 25).Value2 = datos
    If esEnvase Then
        ws.Range("E5:E" & 4 + numeroFilas).NumberFormat = "dd-mmm-yyyy"
    Else
        ws.Range("C5:D" & 4 + numeroFilas).NumberFormat = "dd-mmm-yyyy"
    End If
    ws.Range("S5:S" & 4 + numeroFilas).NumberFormat = "dd-mmm-yyyy"
    ws.Range("G5:J" & 4 + numeroFilas).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("Q5:R" & 4 + numeroFilas).NumberFormat = "#,##0.00;[Red]-#,##0.00"
    ws.Range("T5:T" & 4 + numeroFilas).NumberFormat = "0;[Red]-0"
    ws.Range("U5:V" & 4 + numeroFilas).NumberFormat = "$#,##0.00;[Red]-$#,##0.00"
    ws.Range("Y5:Y" & 4 + numeroFilas).WrapText = True

    FormatearAlertas ws, 5, 4 + numeroFilas, 23, 24

    If ws.ListObjects.Count = 0 Then
        If ws.AutoFilterMode Then ws.AutoFilterMode = False
        ws.Range("A4:Y" & 4 + numeroFilas).AutoFilter
    End If
End Sub

Private Sub LimpiarResultado(ByVal ws As Worksheet)
    Dim ultima As Long
    ultima = MaxLong(UltimaFila(ws, 1), 154)
    ws.Range(ws.Cells(FILA_DATOS_SALIDA, 1), ws.Cells(ultima, 25)).ClearContents
End Sub

Private Sub FormatearAlertas(ByVal ws As Worksheet, ByVal filaInicial As Long, _
                            ByVal filaFinal As Long, ByVal columnaEstado As Long, _
                            ByVal columnaRiesgo As Long)
    Dim r As Long
    Dim estado As String
    Dim riesgo As String

    If filaFinal < filaInicial Then Exit Sub

    For r = filaInicial To filaFinal
        estado = UCase$(Trim$(CStr(ws.Cells(r, columnaEstado).Value2)))
        riesgo = UCase$(Trim$(CStr(ws.Cells(r, columnaRiesgo).Value2)))

        With ws.Cells(r, columnaEstado)
            .Interior.Pattern = xlSolid
            .Font.Bold = True
            Select Case True
                Case estado = "CONCILIADO"
                    .Interior.Color = RGB(230, 244, 234)
                    .Font.Color = RGB(19, 122, 69)
                Case InStr(1, estado, "DUPLICADO", vbTextCompare) > 0 Or _
                     estado = "PENDIENTE SAP" Or estado = "SIN FACTOR" Or _
                     estado = "FECHA INV" & ChrW(193) & "LIDA" Or _
                     estado = "CONFIGURACI" & ChrW(211) & "N DUPLICADA"
                    .Interior.Color = RGB(253, 226, 226)
                    .Font.Color = RGB(164, 22, 43)
                Case Else
                    .Interior.Color = RGB(255, 243, 205)
                    .Font.Color = RGB(122, 83, 0)
            End Select
        End With

        With ws.Cells(r, columnaRiesgo)
            .Interior.Pattern = xlSolid
            .Font.Bold = True
            Select Case riesgo
                Case "ALTO"
                    .Interior.Color = RGB(253, 226, 226)
                    .Font.Color = RGB(164, 22, 43)
                Case "MEDIO"
                    .Interior.Color = RGB(255, 243, 205)
                    .Font.Color = RGB(122, 83, 0)
                Case Else
                    .Interior.Color = RGB(230, 244, 234)
                    .Font.Color = RGB(19, 122, 69)
            End Select
        End With
    Next r
End Sub

Private Sub LeerConfiguracionProducto(ByVal catalogo As Object, ByVal clave As String, _
                                      ByRef existe As Boolean, ByRef duplicada As Boolean, _
                                      ByRef factorValido As Boolean, ByRef factor As Double, _
                                      ByRef material As String, ByRef baseCantidad As String)
    Dim cfg As Variant

    existe = False
    duplicada = False
    factorValido = False
    factor = 0
    material = vbNullString
    baseCantidad = vbNullString

    If catalogo.Exists(clave) Then
        cfg = catalogo(clave)
        existe = True
        factor = CDbl(cfg(0))
        material = CStr(cfg(1))
        baseCantidad = CStr(cfg(2))
        duplicada = CBool(cfg(3))
        factorValido = CBool(cfg(4)) And factor > 0
    End If
End Sub

Private Function CantidadEsperadaProveedor(ByVal tipo As String, ByVal baseCantidad As String, _
                                           ByVal factor As Double, ByVal valor1 As Double, _
                                           ByVal valor2 As Double, ByVal factorValido As Boolean) As Double
    If Not factorValido Then Exit Function

    If InStr(1, baseCantidad, "TARIMA", vbTextCompare) > 0 Then
        CantidadEsperadaProveedor = valor2 * factor
    ElseIf InStr(1, baseCantidad, "CAJA", vbTextCompare) > 0 Then
        CantidadEsperadaProveedor = valor1 * factor
    ElseIf InStr(1, baseCantidad, "CANTIDAD", vbTextCompare) > 0 Then
        CantidadEsperadaProveedor = valor1 * factor
    ElseIf UCase$(tipo) = "ENVASE" Then
        CantidadEsperadaProveedor = valor2 * factor
    Else
        CantidadEsperadaProveedor = valor1 * factor
    End If
End Function

Private Sub InicializarDatosMB51(ByRef movimientos As Long, ByRef cantidad As Double, _
                                 ByRef importe As Double, ByRef fecha As Double, _
                                 ByRef registro As Variant)
    movimientos = 0
    cantidad = 0
    importe = 0
    fecha = 0
    registro = Empty
End Sub

Private Function CampoRegistro(ByVal registro As Variant, ByVal movimientos As Long, _
                               ByVal indice As eMB51) As Variant
    If movimientos > 0 And IsArray(registro) Then
        CampoRegistro = registro(indice)
    Else
        CampoRegistro = vbNullString
    End If
End Function

Private Function MaterialRegistro(ByVal registro As Variant, ByVal movimientos As Long) As String
    If movimientos > 0 And IsArray(registro) Then
        MaterialRegistro = CStr(registro(mbMateriales))
    Else
        MaterialRegistro = vbNullString
    End If
End Function

Private Function ContarClaves(ByVal datos As Variant, ByVal columna As Long) As Long
    Dim i As Long
    For i = 1 To UBound(datos, 1)
        If Len(NormalizarClave(datos(i, columna))) > 0 Then ContarClaves = ContarClaves + 1
    Next i
End Function

Private Sub LeerTolerancias(ByVal ws As Worksheet, ByRef toleranciaCantidad As Double, _
                            ByRef toleranciaFecha As Long)
    Dim ok As Boolean
    Dim valor As Double

    toleranciaCantidad = NumeroSeguro(ws.Range("B7").Value2, ok)
    If Not ok Or toleranciaCantidad < 0 Then toleranciaCantidad = 0

    valor = NumeroSeguro(ws.Range("B8").Value2, ok)
    If Not ok Or valor < 0 Then valor = 1
    toleranciaFecha = CLng(valor)
End Sub

Private Function CrearConfiguracionBase() As Worksheet
    Dim ws As Worksheet
    Set ws = ObtenerOCrearHoja(SH_CONFIG)

    ws.Range("A1").Value = "Configuraci" & ChrW(243) & "n y reglas"
    ws.Range("A5").Value = "Par" & ChrW(225) & "metros de tolerancia"
    EscribirMatriz ws.Range("A6"), Array( _
        Array("Par" & ChrW(225) & "metro", "Valor"), _
        Array("Tolerancia cantidad (unidades)", 0), _
        Array("Tolerancia fecha (d" & ChrW(237) & "as)", 1))
    ws.Range("A11").Value = "Cat" & ChrW(225) & "logo de conversi" & ChrW(243) & "n y material esperado"
    EscribirMatriz ws.Range("A12"), Array( _
        Array("Clave", "Tipo", "Producto / Insumo", "Factor piezas por unidad proveedor", "Material SAP esperado", "Base de cantidad"), _
        Array("Envase|GARRAFON", "Envase", "GARRAFON", 100, "<MATERIAL_ENVASE>", "Tarimas"), _
        Array("Caja|CAJA 3 LT", "Caja", "CAJA 3 LT", 1, "<MATERIAL_CAJA>", "Cantidad proveedor"))

    Set CrearConfiguracionBase = ws
End Function

Private Sub RegistrarLog(ByVal registros As Long, ByVal conciliados As Long, _
                         ByVal revision As Long, ByVal duplicadosProveedor As Long, _
                         ByVal duplicadosMB51 As Long, ByVal segundos As Double)
    Dim ws As Worksheet
    Dim fila As Long

    On Error GoTo Salir
    Set ws = ObtenerOCrearHoja(SH_LOG)

    If Len(CStr(ws.Range("A1").Value2)) = 0 Then
        EscribirMatriz ws.Range("A1"), Array(Array("Fecha y hora", "Usuario", "Registros", _
                                                   "Conciliados", "Por revisar", _
                                                   "Duplicados proveedor", "Duplicados MB51", _
                                                   "Segundos"))
        ws.Range("A1:H1").Font.Bold = True
    End If

    fila = UltimaFila(ws, 1) + 1
    If fila < 2 Then fila = 2
    ws.Cells(fila, 1).Value = Now
    ws.Cells(fila, 2).Value = Application.UserName
    ws.Cells(fila, 3).Value = registros
    ws.Cells(fila, 4).Value = conciliados
    ws.Cells(fila, 5).Value = revision
    ws.Cells(fila, 6).Value = duplicadosProveedor
    ws.Cells(fila, 7).Value = duplicadosMB51
    ws.Cells(fila, 8).Value = segundos
    ws.Columns(1).NumberFormat = "yyyy-mm-dd hh:mm:ss"
    If ThisWorkbook.ActiveSheet Is ws Then
        ThisWorkbook.Worksheets(SH_DASHBOARD).Activate
    End If
    ws.Visible = xlSheetVeryHidden
Salir:
End Sub

Private Function BuscarHoja(ByVal nombres As Variant) As Worksheet
    Dim nombre As Variant
    Dim ws As Worksheet
    Dim normalizado As String

    For Each nombre In nombres
        Set ws = Nothing
        On Error Resume Next
        Set ws = ThisWorkbook.Worksheets(CStr(nombre))
        On Error GoTo 0
        If Not ws Is Nothing Then
            Set BuscarHoja = ws
            Exit Function
        End If
    Next nombre

    For Each ws In ThisWorkbook.Worksheets
        normalizado = NormalizarNombreHoja(ws.Name)
        For Each nombre In nombres
            If normalizado = NormalizarNombreHoja(CStr(nombre)) Then
                Set BuscarHoja = ws
                Exit Function
            End If
        Next nombre
    Next ws
End Function

Private Function ObtenerOCrearHoja(ByVal nombre As String) As Worksheet
    Dim ws As Worksheet

    Set ws = BuscarHoja(Array(nombre))
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = nombre
    End If
    Set ObtenerOCrearHoja = ws
End Function

Private Function NuevoDiccionario() As Object
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")
    dic.CompareMode = vbTextCompare
    Set NuevoDiccionario = dic
End Function

Private Sub AgregarClave(ByVal dic As Object, ByVal clave As String)
    If Len(clave) = 0 Then Exit Sub
    If Not dic.Exists(clave) Then dic.Add clave, True
End Sub

Private Function UltimaFila(ByVal ws As Worksheet, ByVal columna As Long) As Long
    Dim fila As Long
    fila = ws.Cells(ws.Rows.Count, columna).End(xlUp).Row
    If fila < 1 Then fila = 1
    UltimaFila = fila
End Function

Private Function DetectarPrimeraFilaDatos(ByVal ws As Worksheet, _
                                          ByVal columnaClave As Long, _
                                          ByVal encabezadosEsperados As Variant) As Long
    Dim ultima As Long
    Dim limite As Long
    Dim r As Long
    Dim esperado As Variant
    Dim textoCelda As String
    Dim textoEsperado As String

    ultima = UltimaFila(ws, columnaClave)
    limite = MaxLong(ultima, 1)
    If limite > 30 Then limite = 30

    For r = 1 To limite
        textoCelda = NormalizarEncabezado(TextoSeguro(ws.Cells(r, columnaClave).Value2))
        If Len(textoCelda) > 0 Then
            For Each esperado In encabezadosEsperados
                textoEsperado = NormalizarEncabezado(CStr(esperado))
                If Len(textoEsperado) > 0 And _
                   InStr(1, textoCelda, textoEsperado, vbTextCompare) > 0 Then
                    DetectarPrimeraFilaDatos = r + 1
                    Exit Function
                End If
            Next esperado
        End If
    Next r

    Err.Raise vbObjectError + 1010, , _
              "No se encontr" & ChrW(243) & " el encabezado de la clave en la columna " & _
              columnaClave & " de la hoja '" & ws.Name & "'."
End Function

Private Function NormalizarEncabezado(ByVal texto As String) As String
    texto = UCase$(QuitarAcentos(Trim$(texto)))
    texto = Replace(texto, Chr$(160), "")
    texto = Replace(texto, " ", "")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ":", "")
    texto = Replace(texto, "(", "")
    texto = Replace(texto, ")", "")
    texto = Replace(texto, "-", "")
    texto = Replace(texto, "_", "")
    texto = Replace(texto, "/", "")
    NormalizarEncabezado = texto
End Function

Private Function NormalizarClave(ByVal valor As Variant) As String
    Dim texto As String
    Dim numero As Double

    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function

    If VarType(valor) = vbByte Or VarType(valor) = vbInteger Or _
       VarType(valor) = vbLong Or VarType(valor) = vbSingle Or _
       VarType(valor) = vbDouble Or VarType(valor) = vbCurrency Or _
       VarType(valor) = vbDecimal Then
        numero = CDbl(valor)
        If numero = Fix(numero) Then
            NormalizarClave = Format$(numero, "0")
        Else
            NormalizarClave = UCase$(Trim$(CStr(valor)))
        End If
        Exit Function
    End If

    texto = CStr(valor)
    texto = Replace(texto, Chr$(160), " ")
    texto = Replace(texto, vbTab, "")
    texto = Replace(texto, vbCr, "")
    texto = Replace(texto, vbLf, "")
    texto = Replace(texto, "'", "")
    texto = Replace(texto, " ", "")
    texto = Trim$(texto)

    If Len(texto) > 0 And IsNumeric(texto) Then
        On Error Resume Next
        numero = CDbl(texto)
        If Err.Number = 0 And numero = Fix(numero) Then texto = Format$(numero, "0")
        Err.Clear
        On Error GoTo 0
    End If

    NormalizarClave = UCase$(texto)
End Function

Private Function NormalizarTexto(ByVal valor As Variant) As String
    Dim texto As String

    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function
    texto = UCase$(Trim$(CStr(valor)))
    texto = Replace(texto, Chr$(160), " ")
    texto = Replace(texto, vbTab, " ")
    Do While InStr(1, texto, "  ", vbBinaryCompare) > 0
        texto = Replace(texto, "  ", " ")
    Loop
    NormalizarTexto = texto
End Function

Private Function NormalizarNombreHoja(ByVal nombre As String) As String
    Dim texto As String
    texto = UCase$(nombre)
    texto = QuitarAcentos(texto)
    texto = Replace(texto, " ", "")
    texto = Replace(texto, "_", "")
    texto = Replace(texto, "-", "")
    NormalizarNombreHoja = texto
End Function

Private Function QuitarAcentos(ByVal texto As String) As String
    texto = Replace(texto, "" & ChrW(193) & "", "A")
    texto = Replace(texto, "" & ChrW(201) & "", "E")
    texto = Replace(texto, "" & ChrW(205) & "", "I")
    texto = Replace(texto, "" & ChrW(211) & "", "O")
    texto = Replace(texto, "" & ChrW(218) & "", "U")
    texto = Replace(texto, "" & ChrW(220) & "", "U")
    texto = Replace(texto, "" & ChrW(209) & "", "N")
    QuitarAcentos = texto
End Function

Private Function TextoSeguro(ByVal valor As Variant) As String
    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function
    TextoSeguro = Trim$(CStr(valor))
End Function

Private Function NumeroSeguro(ByVal valor As Variant, ByRef esValido As Boolean) As Double
    Dim texto As String
    Dim posicionPunto As Long
    Dim posicionComa As Long
    Dim cantidadPuntos As Long
    Dim cantidadComas As Long
    Dim decimales As Long
    Dim negativo As Boolean
    Dim tipo As VbVarType

    esValido = False
    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function

    tipo = VarType(valor)
    If tipo = vbByte Or tipo = vbInteger Or tipo = vbLong Or _
       tipo = vbSingle Or tipo = vbDouble Or tipo = vbCurrency Or _
       tipo = vbDecimal Then
        NumeroSeguro = CDbl(valor)
        esValido = True
        Exit Function
    End If

    texto = Trim$(CStr(valor))
    If Len(texto) = 0 Then Exit Function
    negativo = (Left$(texto, 1) = "(" And Right$(texto, 1) = ")")
    If negativo Then texto = Mid$(texto, 2, Len(texto) - 2)

    texto = Replace(texto, Chr$(160), "")
    texto = Replace(texto, "$", "")
    texto = Replace(texto, " ", "")
    texto = Replace(texto, "MXN", "", 1, -1, vbTextCompare)

    posicionPunto = InStrRev(texto, ".")
    posicionComa = InStrRev(texto, ",")
    cantidadPuntos = Len(texto) - Len(Replace(texto, ".", ""))
    cantidadComas = Len(texto) - Len(Replace(texto, ",", ""))

    If posicionPunto > 0 And posicionComa > 0 Then
        If posicionPunto > posicionComa Then
            texto = Replace(texto, ",", "")
        Else
            texto = Replace(texto, ".", "")
            texto = Replace(texto, ",", ".")
        End If
    ElseIf posicionComa > 0 Then
        decimales = Len(texto) - posicionComa
        If cantidadComas = 1 And decimales > 0 And decimales <= 2 Then
            texto = Replace(texto, ",", ".")
        Else
            texto = Replace(texto, ",", "")
        End If
    ElseIf posicionPunto > 0 Then
        decimales = Len(texto) - posicionPunto
        If cantidadPuntos > 1 Or decimales = 0 Or decimales > 2 Then
            texto = Replace(texto, ".", "")
        End If
    End If

    If EsNumeroCanonico(texto) Then
        NumeroSeguro = Val(texto)
        If negativo Then NumeroSeguro = -Abs(NumeroSeguro)
        esValido = True
    End If
End Function

Private Function EsNumeroCanonico(ByVal texto As String) As Boolean
    Dim i As Long
    Dim caracter As String
    Dim puntos As Long
    Dim digitos As Long

    If Len(texto) = 0 Then Exit Function

    For i = 1 To Len(texto)
        caracter = Mid$(texto, i, 1)
        If caracter >= "0" And caracter <= "9" Then
            digitos = digitos + 1
        ElseIf caracter = "." Then
            puntos = puntos + 1
            If puntos > 1 Then Exit Function
        ElseIf (caracter = "-" Or caracter = "+") And i = 1 Then
            ' Signo permitido unicamente al inicio.
        Else
            Exit Function
        End If
    Next i

    EsNumeroCanonico = (digitos > 0)
End Function

Private Function FechaSerial(ByVal valor As Variant) As Double
    If IsError(valor) Or IsNull(valor) Or IsEmpty(valor) Then Exit Function

    If IsDate(valor) Then
        FechaSerial = CDbl(CDate(valor))
    ElseIf IsNumeric(valor) Then
        If CDbl(valor) > 0 Then FechaSerial = CDbl(valor)
    End If
End Function

Private Function ValorFechaSalida(ByVal fecha As Double) As Variant
    If fecha > 0 Then
        ValorFechaSalida = fecha
    Else
        ValorFechaSalida = vbNullString
    End If
End Function

Private Function AplicarSignoMovimiento(ByVal valor As Double, ByVal movimiento As String) As Double
    If valor < 0 Then
        AplicarSignoMovimiento = valor
    ElseIf EsMovimientoReversa(movimiento) Then
        AplicarSignoMovimiento = -Abs(valor)
    Else
        AplicarSignoMovimiento = valor
    End If
End Function

Private Function EsMovimientoReversa(ByVal movimiento As String) As Boolean
    Select Case CLng(Val(movimiento))
        Case 102, 122, 162, 202, 222, 262, 302, 312, 322, 342, _
             502, 512, 522, 532, 542, 552, 562, 602
            EsMovimientoReversa = True
    End Select
End Function

Private Function AgregarUnico(ByVal actual As String, ByVal nuevo As String) As String
    Dim partes As Variant
    Dim parte As Variant

    nuevo = Trim$(nuevo)
    If Len(nuevo) = 0 Then
        AgregarUnico = actual
        Exit Function
    End If

    If Len(actual) = 0 Then
        AgregarUnico = nuevo
        Exit Function
    End If

    partes = Split(actual, " | ")
    For Each parte In partes
        If StrComp(Trim$(CStr(parte)), nuevo, vbTextCompare) = 0 Then
            AgregarUnico = actual
            Exit Function
        End If
    Next parte

    AgregarUnico = actual & " | " & nuevo
End Function

Private Function PorcentajeSeguro(ByVal numerador As Variant, ByVal denominador As Variant) As Double
    If CDbl(denominador) <> 0 Then PorcentajeSeguro = CDbl(numerador) / CDbl(denominador)
End Function

Private Sub EscribirMatriz(ByVal celdaInicial As Range, ByVal filas As Variant)
    Dim matriz() As Variant
    Dim i As Long
    Dim j As Long
    Dim numeroFilas As Long
    Dim numeroColumnas As Long
    Dim filaDatos As Variant

    numeroFilas = UBound(filas) - LBound(filas) + 1
    filaDatos = filas(LBound(filas))
    numeroColumnas = UBound(filaDatos) - LBound(filaDatos) + 1
    ReDim matriz(1 To numeroFilas, 1 To numeroColumnas)

    For i = LBound(filas) To UBound(filas)
        filaDatos = filas(i)
        If UBound(filaDatos) - LBound(filaDatos) + 1 <> numeroColumnas Then
            Err.Raise vbObjectError + 1020, , "La matriz contiene filas con distinto n" & ChrW(250) & "mero de columnas."
        End If
        For j = LBound(filaDatos) To UBound(filaDatos)
            matriz(i - LBound(filas) + 1, j - LBound(filaDatos) + 1) = filaDatos(j)
        Next j
    Next i

    celdaInicial.Resize(numeroFilas, numeroColumnas).Value2 = matriz
End Sub

Private Function MaxLong(ByVal a As Long, ByVal b As Long) As Long
    If a > b Then MaxLong = a Else MaxLong = b
End Function
