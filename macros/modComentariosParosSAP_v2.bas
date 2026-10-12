Attribute VB_Name = "modComentariosParosSAP"
Option Explicit

' VERSION 2 - CON DIAGNOSTICO Y CONEXION SAP ROBUSTA
' =====================================================================
' COMENTARIOS DE PAROS DE LINEA - REPORTE DE EFICIENCIA <PLANTA>
'
' Objetivo:
'   - Lee la fecha de REPORTE EFICIENCIA!F3.
'   - Ejecuta ZREP_PAROS en SAP con variante /VAR_PAROS.
'   - Descarga SOLO el detalle del dia.
'   - Suma los minutos por causa de paro para cada linea.
'   - Toma las 3 causas con mas minutos.
'   - Escribe el resumen en REPORTE EFICIENCIA!W6:W9.
'
' Ejemplo:
'   111 min Llenadora 1 L2 -PLT Sist. sensor/fotocel. falla
'   55 min Llenadora 1 L2 -PLT Variador de frecuencia falla
'   38 min Llenadora 1 L2 -PLT PLC falla
'
' No modifica REP2, CTT, Dia, Ago 26 ni otras celdas del reporte.
' =====================================================================

Private Const CENTRO As String = "<CENTRO>"   ' centro (WERKS) real
Private Const TRANSACCION As String = "ZREP_PAROS"
Private Const VARIANTE As String = "/VAR_PAROS"
Private Const HOJA_REPORTE As String = "REPORTE EFICIENCIA"
Private Const CELDA_FECHA As String = "F3"
Private Const MAX_COMENTARIOS As Long = 3

Public Sub ActualizarComentariosParosSAP()
    Dim ses As Object
    Dim fechaReporte As Date
    Dim carpeta As String
    Dim archivo As String
    Dim wbSAP As Workbook
    Dim wsSAP As Worksheet
    Dim estadoPantalla As Boolean
    Dim estadoEventos As Boolean
    Dim estadoAlertas As Boolean

    On Error GoTo ERROR_GENERAL

    estadoPantalla = Application.ScreenUpdating
    estadoEventos = Application.EnableEvents
    estadoAlertas = Application.DisplayAlerts

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False

    fechaReporte = ObtenerFechaReporte()
    carpeta = ObtenerCarpetaTemporal()
    archivo = carpeta & "\PAROS_" & Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivo

    Set ses = ObtenerSesionSAP()
    If ses Is Nothing Then
        Err.Raise vbObjectError + 2000, , _
            "No se encontro una sesion activa de SAP GUI."
    End If

    ExportarDetalleParos ses, fechaReporte, archivo
    EsperarArchivo archivo, 30

    Set wbSAP = AbrirExportSAP(archivo)
    If wbSAP Is Nothing Then
        Err.Raise vbObjectError + 2001, , _
            "No fue posible abrir el archivo exportado por SAP."
    End If

    Set wsSAP = wbSAP.Worksheets(1)
    ProcesarComentarios wsSAP, fechaReporte

    wbSAP.Close SaveChanges:=False
    Set wbSAP = Nothing

    ThisWorkbook.Worksheets(HOJA_REPORTE).Activate
    ThisWorkbook.Worksheets(HOJA_REPORTE).Range("W6").Select

    MsgBox "Comentarios de paros actualizados." & vbCrLf & vbCrLf & _
           "Fecha: " & Format$(fechaReporte, "dd/mm/yyyy") & vbCrLf & _
           "Destino: " & HOJA_REPORTE & "!W6:W9", _
           vbInformation, "PAROS DE LINEA"

SALIDA:
    On Error Resume Next
    If Not wbSAP Is Nothing Then wbSAP.Close SaveChanges:=False
    BorrarSiExiste archivo
    Application.DisplayAlerts = estadoAlertas
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoPantalla
    Exit Sub

ERROR_GENERAL:
    MsgBox "No se pudieron actualizar los comentarios de paros." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "PAROS DE LINEA"
    Resume SALIDA
End Sub

' =====================================================================
' SAP
' =====================================================================
Private Function ObtenerSesionSAP() As Object
    Dim SapGuiAuto As Object
    Dim appSAP As Object
    Dim con As Object
    Dim ses As Object
    Dim rot As Object
    Dim i As Long
    Dim j As Long
    Dim errGetObject As Long
    Dim descGetObject As String
    Dim wnd0 As Object

    ' -------------------------------------------------------------
    ' Metodo 1 (oficial SAP): adjuntarse al objeto SAPGUI del ROT.
    ' -------------------------------------------------------------
    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    errGetObject = Err.Number
    descGetObject = Err.Description
    Err.Clear
    On Error GoTo 0

    ' -------------------------------------------------------------
    ' Metodo 2: SAP ROT Wrapper. Sirve como respaldo cuando Excel
    ' no puede obtener directamente la entrada SAPGUI del ROT.
    ' -------------------------------------------------------------
    If SapGuiAuto Is Nothing Then
        On Error Resume Next
        Set rot = CreateObject("SapROTWr.SapROTWrapper")
        If Not rot Is Nothing Then
            Set SapGuiAuto = rot.GetROTEntry("SAPGUI")
        End If
        On Error GoTo 0
    End If

    If SapGuiAuto Is Nothing Then
        Err.Raise vbObjectError + 2100, , _
            "Excel no puede adjuntarse al proceso SAP GUI." & vbCrLf & _
            "Abre SAP GUI for Windows, inicia sesion y deja una ventana SAP abierta." & vbCrLf & _
            "Si SAP y Excel se ejecutan con distinto nivel de privilegios (Administrador/normal), cierralos y abre ambos de la misma forma." & _
            IIf(errGetObject <> 0, vbCrLf & "Detalle GetObject: " & errGetObject & " - " & descGetObject, "")
    End If

    On Error GoTo ERROR_ENGINE
    Set appSAP = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0

    If appSAP Is Nothing Then
        Err.Raise vbObjectError + 2101, , _
            "SAP GUI existe, pero no fue posible obtener GetScriptingEngine. Revisa que SAP GUI Scripting este habilitado."
    End If

    If appSAP.Children.Count = 0 Then
        Err.Raise vbObjectError + 2102, , _
            "SAP GUI esta abierto, pero no hay conexiones activas. Inicia sesion en el sistema SAP antes de ejecutar la macro."
    End If

    ' Buscar una sesion real y utilizable, no asumir siempre Children(0).
    For i = 0 To appSAP.Children.Count - 1
        Set con = Nothing
        On Error Resume Next
        Set con = appSAP.Children(i)
        On Error GoTo 0

        If Not con Is Nothing Then
            For j = 0 To con.Children.Count - 1
                Set ses = Nothing
                On Error Resume Next
                Set ses = con.Children(j)
                On Error GoTo 0

                If Not ses Is Nothing Then
                    ' Verifica que la sesion exponga una ventana principal.
                    On Error Resume Next
                    Err.Clear
                    Set wnd0 = Nothing
                    Set wnd0 = ses.findById("wnd[0]")
                    If Err.Number = 0 And Not wnd0 Is Nothing Then
                        On Error GoTo 0
                        Set ObtenerSesionSAP = ses
                        Exit Function
                    End If
                    Err.Clear
                    On Error GoTo 0
                End If
            Next j
        End If
    Next i

    Err.Raise vbObjectError + 2103, , _
        "Se detecto SAP GUI, pero no se encontro una sesion utilizable. Cierra ventanas de login/popup, entra al menu SAP y vuelve a ejecutar."
    Exit Function

ERROR_ENGINE:
    Err.Raise vbObjectError + 2104, , _
        "No se pudo acceder al motor de SAP GUI Scripting." & vbCrLf & _
        "Error original " & Err.Number & ": " & Err.Description
End Function

Public Sub DiagnosticarSAPGUI()
    Dim SapGuiAuto As Object
    Dim appSAP As Object
    Dim con As Object
    Dim ses As Object
    Dim rot As Object
    Dim i As Long
    Dim j As Long
    Dim txt As String
    Dim metodo As String

    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Not SapGuiAuto Is Nothing Then metodo = "GetObject(SAPGUI)"
    On Error GoTo 0

    If SapGuiAuto Is Nothing Then
        On Error Resume Next
        Set rot = CreateObject("SapROTWr.SapROTWrapper")
        If Not rot Is Nothing Then Set SapGuiAuto = rot.GetROTEntry("SAPGUI")
        If Not SapGuiAuto Is Nothing Then metodo = "SapROTWrapper"
        On Error GoTo 0
    End If

    If SapGuiAuto Is Nothing Then
        MsgBox "DIAGNOSTICO SAP GUI" & vbCrLf & vbCrLf & _
               "NO se encontro el objeto SAPGUI en Windows." & vbCrLf & vbCrLf & _
               "1. Abre SAP GUI for Windows." & vbCrLf & _
               "2. Inicia sesion completamente." & vbCrLf & _
               "3. Deja abierta una pantalla SAP." & vbCrLf & _
               "4. Ejecuta SAP y Excel con el mismo nivel de privilegios.", _
               vbCritical, "DIAGNOSTICO SAP"
        Exit Sub
    End If

    On Error GoTo ERR_ENGINE
    Set appSAP = SapGuiAuto.GetScriptingEngine

    txt = "DIAGNOSTICO SAP GUI" & vbCrLf & vbCrLf & _
          "Conexion al ROT: OK" & vbCrLf & _
          "Metodo: " & metodo & vbCrLf & _
          "Conexiones detectadas: " & appSAP.Children.Count & vbCrLf

    For i = 0 To appSAP.Children.Count - 1
        Set con = appSAP.Children(i)
        txt = txt & vbCrLf & "Conexion " & (i + 1) & ": " & con.Children.Count & " sesion(es)"

        For j = 0 To con.Children.Count - 1
            Set ses = con.Children(j)
            txt = txt & vbCrLf & "  Sesion " & (j + 1)
            On Error Resume Next
            txt = txt & " | Sistema: " & ses.Info.SystemName
            txt = txt & " | Mandante: " & ses.Info.Client
            txt = txt & " | Usuario: " & ses.Info.User
            txt = txt & " | TX: " & ses.Info.Transaction
            On Error GoTo 0
        Next j
    Next i

    MsgBox txt, vbInformation, "DIAGNOSTICO SAP"
    Exit Sub

ERR_ENGINE:
    MsgBox "SAP GUI fue detectado, pero GetScriptingEngine fallo." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "Revisa que SAP GUI Scripting este habilitado en cliente y servidor.", _
           vbCritical, "DIAGNOSTICO SAP"
End Sub

Private Sub ExportarDetalleParos(ByVal ses As Object, _
                                 ByVal fechaReporte As Date, _
                                 ByVal rutaCompleta As String)
    Dim carpeta As String
    Dim nombre As String

    carpeta = CarpetaDeRuta(rutaCompleta)
    nombre = NombreDeRuta(rutaCompleta)

    ses.findById("wnd[0]").maximize
    IrATransaccion ses, TRANSACCION

    ses.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = CENTRO

    ' Para los comentarios solo necesitamos el dia mostrado en F3.
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-LOW").Text = _
        Format$(fechaReporte, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").Text = _
        Format$(fechaReporte, "dd.mm.yyyy")

    ses.findById("wnd[0]/usr/radP_DET").Select
    ses.findById("wnd[0]/usr/ctxtP_VARI").Text = VARIANTE

    ses.findById("wnd[0]").sendVKey 8
    EsperarSAP ses

    ' Archivo local. Corresponde al flujo grabado con SAP GUI Scripting.
    ses.findById("wnd[0]").sendVKey 45
    EsperarSAP ses

    ses.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/" & _
                 "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    ses.findById("wnd[1]/tbar[0]/btn[0]").press
    EsperarSAP ses

    PonerTextoSiExiste ses, "wnd[1]/usr/ctxtDY_PATH", carpeta
    ses.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre
    ses.findById("wnd[1]/tbar[0]/btn[11]").press
    EsperarSAP ses
End Sub

Private Sub IrATransaccion(ByVal ses As Object, ByVal tx As String)
    EsperarSAP ses
    ses.findById("wnd[0]/tbar[0]/okcd").Text = "/n" & tx
    ses.findById("wnd[0]").sendVKey 0
    EsperarSAP ses
End Sub

Private Sub EsperarSAP(ByVal ses As Object, Optional ByVal segundos As Long = 30)
    Dim limite As Date
    limite = DateAdd("s", segundos, Now)

    Do
        DoEvents
        If Not ses.Busy Then Exit Do
        If Now > limite Then
            Err.Raise vbObjectError + 2002, , _
                "SAP tardo demasiado en responder."
        End If
    Loop
End Sub

Private Sub PonerTextoSiExiste(ByVal ses As Object, _
                               ByVal id As String, _
                               ByVal valor As String)
    Dim obj As Object

    On Error Resume Next
    Set obj = ses.findById(id)
    If Not obj Is Nothing Then obj.Text = valor
    On Error GoTo 0
End Sub

' =====================================================================
' LECTURA DEL EXPORT DE SAP Y GENERACION DE COMENTARIOS
' =====================================================================
Private Sub ProcesarComentarios(ByVal ws As Worksheet, ByVal fechaReporte As Date)
    Dim colCentro As Long
    Dim colFecha As Long
    Dim colLinea As Long
    Dim colTipo As Long
    Dim colClave As Long
    Dim colSubclave As Long
    Dim colClave1 As Long
    Dim colObs As Long
    Dim colMin As Long
    Dim filaInicio As Long
    Dim ultimaFila As Long
    Dim fila As Long
    Dim idx As Long
    Dim minutos As Double
    Dim descripcion As String
    Dim claveDic As String
    Dim tipo As String
    Dim linea As String
    Dim registrosFecha As Long
    Dim registrosLinea(0 To 3) As Long
    Dim dicMin(0 To 3) As Object
    Dim dicTexto(0 To 3) As Object
    Dim i As Long
    Dim wsRep As Worksheet
    Dim filasDestino As Variant
    Dim lineasSAP As Variant

    lineasSAP = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")
    filasDestino = Array(6, 7, 8, 9)

    For i = 0 To 3
        Set dicMin(i) = CreateObject("Scripting.Dictionary")
        Set dicTexto(i) = CreateObject("Scripting.Dictionary")
        dicMin(i).CompareMode = vbTextCompare
        dicTexto(i).CompareMode = vbTextCompare
    Next i

    DetectarColumnas ws, filaInicio, colCentro, colFecha, colLinea, _
                     colTipo, colClave, colSubclave, colClave1, colObs, colMin

    ultimaFila = UltimaFilaUsada(ws)

    For fila = filaInicio To ultimaFila
        If MismaFecha(ws.Cells(fila, colFecha).Value, fechaReporte) Then
            registrosFecha = registrosFecha + 1

            linea = UCase$(Trim$(CStr(ws.Cells(fila, colLinea).Value)))
            idx = IndiceLinea(linea)

            If idx >= 0 Then
                registrosLinea(idx) = registrosLinea(idx) + 1
                tipo = UCase$(Trim$(CStr(ws.Cells(fila, colTipo).Value)))

                ' PRODUCCION no es paro y no debe aparecer en comentarios.
                If tipo <> vbNullString And InStr(1, tipo, "PRODUCCION", vbTextCompare) = 0 Then
                    minutos = NumeroSeguro(ws.Cells(fila, colMin).Value)

                    If minutos > 0 Then
                        descripcion = ConstruirDescripcion( _
                            CStr(ws.Cells(fila, colTipo).Value), _
                            CStr(ws.Cells(fila, colClave).Value), _
                            CStr(ws.Cells(fila, colSubclave).Value), _
                            CStr(ws.Cells(fila, colClave1).Value), _
                            CStr(ws.Cells(fila, colObs).Value))

                        descripcion = LimpiarTexto(descripcion)

                        If descripcion <> vbNullString Then
                            claveDic = UCase$(descripcion)

                            If dicMin(idx).Exists(claveDic) Then
                                dicMin(idx).Item(claveDic) = CDbl(dicMin(idx).Item(claveDic)) + minutos
                            Else
                                dicMin(idx).Add claveDic, minutos
                                dicTexto(idx).Add claveDic, descripcion
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next fila

    If registrosFecha = 0 Then
        Err.Raise vbObjectError + 2010, , _
            "El archivo exportado no contiene registros para " & _
            Format$(fechaReporte, "dd/mm/yyyy") & "."
    End If

    Set wsRep = ThisWorkbook.Worksheets(HOJA_REPORTE)

    ' Solo se tocan las cuatro celdas de comentarios.
    For i = 0 To 3
        If dicMin(i).Count > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = _
                CrearResumen(dicMin(i), dicTexto(i), MAX_COMENTARIOS)
        ElseIf InStr(1, UCase$(CStr(wsRep.Cells(filasDestino(i), 3).Value)), _
                           "PARO", vbTextCompare) > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = "PARO PROGRAMADO"
        ElseIf registrosLinea(i) > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = "SIN PAROS"
        Else
            wsRep.Range("W" & filasDestino(i)).Value = "SIN PRODUCCION / SIN REGISTROS"
        End If

        wsRep.Range("W" & filasDestino(i)).WrapText = True
    Next i
End Sub

Private Sub DetectarColumnas(ByVal ws As Worksheet, _
                             ByRef filaInicio As Long, _
                             ByRef colCentro As Long, _
                             ByRef colFecha As Long, _
                             ByRef colLinea As Long, _
                             ByRef colTipo As Long, _
                             ByRef colClave As Long, _
                             ByRef colSubclave As Long, _
                             ByRef colClave1 As Long, _
                             ByRef colObs As Long, _
                             ByRef colMin As Long)
    Dim filaEncabezado As Long

    colCentro = BuscarEncabezado(ws, "Centro", filaEncabezado)
    colFecha = BuscarEncabezado(ws, "Fecha Contable", filaEncabezado)
    colLinea = BuscarEncabezado(ws, "L" & ChrW(237) & "nea", filaEncabezado)
    If colLinea = 0 Then colLinea = BuscarEncabezado(ws, "Linea", filaEncabezado)

    colTipo = BuscarEncabezado(ws, "Tipo de Paro", filaEncabezado)
    colClave = BuscarEncabezado(ws, "Clave de Paro", filaEncabezado)
    colSubclave = BuscarEncabezado(ws, "Subclave de Paro", filaEncabezado)
    colClave1 = BuscarEncabezado(ws, "Clave 1 de Paro", filaEncabezado)
    colObs = BuscarEncabezado(ws, "Observaciones", filaEncabezado)
    colMin = BuscarEncabezado(ws, "Minutos de paro", filaEncabezado)

    If colFecha > 0 And colLinea > 0 And colTipo > 0 And colMin > 0 Then
        filaInicio = filaEncabezado + 1

        ' Columnas opcionales: si alguna no viene en el layout, se evita error.
        If colCentro = 0 Then colCentro = 1
        If colClave = 0 Then colClave = colTipo + 1
        If colSubclave = 0 Then colSubclave = colClave + 1
        If colClave1 = 0 Then colClave1 = colSubclave + 1
        If colObs = 0 Then colObs = colClave1 + 1
        Exit Sub
    End If

    ' Fallback basado en el Script4.xls de la macro grabada por SAP:
    ' los datos empiezan en B13:U...
    filaInicio = 13
    colCentro = 2       ' B
    colFecha = 3        ' C
    colLinea = 4        ' D
    colTipo = 11        ' K
    colClave = 12       ' L
    colSubclave = 13    ' M
    colClave1 = 14      ' N
    colObs = 15         ' O
    colMin = 20         ' T
End Sub

Private Function BuscarEncabezado(ByVal ws As Worksheet, _
                                  ByVal texto As String, _
                                  ByRef filaEncabezado As Long) As Long
    Dim rng As Range
    Dim c As Range

    Set rng = ws.Range("A1:AZ30")

    On Error Resume Next
    Set c = rng.Find(What:=texto, After:=rng.Cells(rng.Cells.Count), _
                     LookIn:=xlValues, LookAt:=xlWhole, _
                     SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                     MatchCase:=False)
    On Error GoTo 0

    If Not c Is Nothing Then
        BuscarEncabezado = c.Column
        filaEncabezado = c.Row
    End If
End Function

Private Function ConstruirDescripcion(ByVal tipo As String, _
                                      ByVal clave As String, _
                                      ByVal subclave As String, _
                                      ByVal clave1 As String, _
                                      ByVal observacion As String) As String
    tipo = LimpiarTexto(tipo)
    clave = LimpiarTexto(clave)
    subclave = LimpiarTexto(subclave)
    clave1 = LimpiarTexto(clave1)
    observacion = LimpiarTexto(observacion)

    ' En P.EQ.LINEA interesa mostrar equipo + falla.
    ' Ej.: Llenadora 1 L2 -PLT Sist. sensor/fotocel. falla
    If InStr(1, UCase$(tipo), "EQ.LINEA", vbTextCompare) > 0 Then
        If clave <> vbNullString And subclave <> vbNullString Then
            ConstruirDescripcion = clave & " " & subclave
        ElseIf clave <> vbNullString And clave1 <> vbNullString Then
            ConstruirDescripcion = clave & " " & clave1
        ElseIf subclave <> vbNullString Then
            ConstruirDescripcion = subclave
        ElseIf clave1 <> vbNullString Then
            ConstruirDescripcion = clave1
        ElseIf observacion <> vbNullString Then
            ConstruirDescripcion = observacion
        Else
            ConstruirDescripcion = clave
        End If
        Exit Function
    End If

    ' Para operativos, ajenos y servicios auxiliares la subclave suele ser
    ' el texto que se usa en la tabla: TORMENTA / INUNDACION,
    ' AJUSTE TRANSP. CAJAS, EXPLOSION DE ENVASES, etc.
    If subclave <> vbNullString Then
        ConstruirDescripcion = subclave
    ElseIf observacion <> vbNullString Then
        ConstruirDescripcion = observacion
    ElseIf clave1 <> vbNullString Then
        ConstruirDescripcion = clave1
    Else
        ConstruirDescripcion = clave
    End If
End Function

Private Function CrearResumen(ByVal dicMin As Object, _
                              ByVal dicTexto As Object, _
                              ByVal maxElementos As Long) As String
    Dim keys As Variant
    Dim usados() As Boolean
    Dim i As Long
    Dim j As Long
    Dim mejor As Long
    Dim mejorValor As Double
    Dim valor As Double
    Dim salida As String
    Dim cantidad As Long

    If dicMin.Count = 0 Then Exit Function

    keys = dicMin.Keys
    ReDim usados(LBound(keys) To UBound(keys))

    cantidad = dicMin.Count
    If cantidad > maxElementos Then cantidad = maxElementos

    For i = 1 To cantidad
        mejor = -1
        mejorValor = -1

        For j = LBound(keys) To UBound(keys)
            If Not usados(j) Then
                valor = CDbl(dicMin.Item(keys(j)))
                If valor > mejorValor Then
                    mejorValor = valor
                    mejor = j
                End If
            End If
        Next j

        If mejor >= 0 Then
            usados(mejor) = True

            If salida <> vbNullString Then salida = salida & vbLf
            salida = salida & FormatearMinutos(mejorValor) & " min " & _
                     CStr(dicTexto.Item(keys(mejor)))
        End If
    Next i

    CrearResumen = salida
End Function

Private Function FormatearMinutos(ByVal valor As Double) As String
    Dim redondeado As Double

    redondeado = Application.WorksheetFunction.Round(valor, 0)
    FormatearMinutos = Format$(redondeado, "0")
End Function

Private Function IndiceLinea(ByVal linea As String) As Long
    linea = UCase$(Replace(Replace(Trim$(linea), " ", vbNullString), "-", vbNullString))

    Select Case linea
        Case "LINEA001", "LINEA01", "LINEA1", "L1"
            IndiceLinea = 0
        Case "LINEA002", "LINEA02", "LINEA2", "L2"
            IndiceLinea = 1
        Case "LINEA003", "LINEA03", "LINEA3", "L3"
            IndiceLinea = 2
        Case "LINEA004", "LINEA04", "LINEA4", "L4"
            IndiceLinea = 3
        Case Else
            IndiceLinea = -1
    End Select
End Function

Private Function MismaFecha(ByVal valor As Variant, ByVal fecha As Date) As Boolean
    Dim s As String

    On Error Resume Next
    If IsDate(valor) Then
        MismaFecha = (DateValue(CDate(valor)) = DateValue(fecha))
        Exit Function
    End If
    On Error GoTo 0

    s = Trim$(CStr(valor))

    MismaFecha = (s = Format$(fecha, "dd.mm.yyyy") Or _
                  s = Format$(fecha, "dd/mm/yyyy") Or _
                  s = Format$(fecha, "d.m.yyyy") Or _
                  s = Format$(fecha, "d/m/yyyy"))
End Function

Private Function NumeroSeguro(ByVal valor As Variant) As Double
    Dim s As String

    If IsNumeric(valor) Then
        NumeroSeguro = CDbl(valor)
        Exit Function
    End If

    s = Trim$(CStr(valor))
    If s = vbNullString Then Exit Function

    ' SAP puede exportar decimal con punto o coma dependiendo del usuario.
    If Application.DecimalSeparator = "," Then
        s = Replace(s, ".", ",")
    Else
        s = Replace(s, ",", ".")
    End If

    If IsNumeric(s) Then NumeroSeguro = CDbl(s)
End Function

Private Function LimpiarTexto(ByVal texto As String) As String
    texto = Replace(texto, vbCr, " ")
    texto = Replace(texto, vbLf, " ")
    texto = Replace(texto, vbTab, " ")
    texto = Trim$(texto)

    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop

    LimpiarTexto = texto
End Function

' =====================================================================
' ARCHIVOS / UTILIDADES
' =====================================================================
Private Function ObtenerFechaReporte() As Date
    Dim v As Variant

    v = ThisWorkbook.Worksheets(HOJA_REPORTE).Range(CELDA_FECHA).Value

    If IsDate(v) Then
        ObtenerFechaReporte = CDate(v)
    Else
        Err.Raise vbObjectError + 2020, , _
            "La celda " & HOJA_REPORTE & "!" & CELDA_FECHA & _
            " no contiene una fecha valida."
    End If
End Function

Private Function ObtenerCarpetaTemporal() As String
    Dim base As String
    Dim ruta As String

    base = ThisWorkbook.Path
    If Len(base) = 0 Then base = Environ$("TEMP")

    ruta = base & "\SAP_TEMP"
    If Dir$(ruta, vbDirectory) = vbNullString Then MkDir ruta

    ObtenerCarpetaTemporal = ruta
End Function

Private Function AbrirExportSAP(ByVal ruta As String) As Workbook
    On Error Resume Next

    ' SAP suele crear .XLS que realmente es texto tabulado.
    Workbooks.OpenText Filename:=ruta, Origin:=xlWindows, StartRow:=1, _
        DataType:=xlDelimited, TextQualifier:=xlTextQualifierDoubleQuote, _
        ConsecutiveDelimiter:=False, Tab:=True, Semicolon:=False, _
        Comma:=False, Space:=False, Other:=False, TrailingMinusNumbers:=True

    If Err.Number = 0 Then
        Set AbrirExportSAP = ActiveWorkbook
        Exit Function
    End If

    Err.Clear
    Set AbrirExportSAP = Workbooks.Open(Filename:=ruta, ReadOnly:=True)
    On Error GoTo 0
End Function

Private Sub EsperarArchivo(ByVal ruta As String, Optional ByVal segundos As Long = 30)
    Dim limite As Date
    limite = DateAdd("s", segundos, Now)

    Do While Dir$(ruta) = vbNullString
        DoEvents
        If Now > limite Then
            Err.Raise vbObjectError + 2030, , _
                "SAP no genero el archivo esperado:" & vbCrLf & ruta
        End If
    Loop
End Sub

Private Sub BorrarSiExiste(ByVal ruta As String)
    On Error Resume Next
    If Len(ruta) > 0 Then
        If Dir$(ruta) <> vbNullString Then Kill ruta
    End If
    On Error GoTo 0
End Sub

Private Function CarpetaDeRuta(ByVal ruta As String) As String
    CarpetaDeRuta = Left$(ruta, InStrRev(ruta, "\") - 1)
End Function

Private Function NombreDeRuta(ByVal ruta As String) As String
    NombreDeRuta = Mid$(ruta, InStrRev(ruta, "\") + 1)
End Function

Private Function UltimaFilaUsada(ByVal ws As Worksheet) As Long
    Dim c As Range

    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByRows, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If c Is Nothing Then
        UltimaFilaUsada = 1
    Else
        UltimaFilaUsada = c.Row
    End If
End Function
