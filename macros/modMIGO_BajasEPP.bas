Attribute VB_Name = "modMIGO_BajasEPP"
Option Explicit

' ============================================================
' BAJAS DE EPP EN MIGO  (A07 Salida de mercancias / R10 Otros / mov. 201 a CeCo)
' Autor: Jose Luis Cruz Prieto
' Version 3.0
'
' Flujo:
'   1) Pegar el export de la app (con encabezados) en PEGAR_APP.
'   2) Preparar_Bajas_Desde_App   -> llena BAJAS_SAP como VALORES (sin formulas),
'                                    conserva el estado SAP de filas ya procesadas.
'   3) Validar_Bajas_EPP          -> marca LISTO / REVISAR por fila.
'   4) Procesar_Bajas_EPP_MIGO    -> MODO=SIMULACION: no se conecta a SAP.
'                                    MODO=VERIFICAR : captura en MIGO y presiona
'                                    Verificar. NUNCA contabiliza.
'   5) Contabilizar a mano en MIGO y registrar el documento con
'      Registrar_Contabilizacion_Manual.
'
' Utilerias: Probar_Conexion_SAP, Diagnosticar_Layout_MIGO.
'
' Toda la configuracion vive en la hoja CONFIG (columna A = parametro,
' columna B = valor). Ver docs/layout-bajas-epp.md.
' ============================================================

Private Const SHEET_APP As String = "PEGAR_APP"
Private Const SHEET_BAJAS As String = "BAJAS_SAP"
Private Const SHEET_CONFIG As String = "CONFIG"
Private Const SHEET_CATALOGOS As String = "CATALOGOS"
Private Const SHEET_LOG As String = "LOG_SAP"
Private Const SHEET_DIAG As String = "DIAG_MIGO"

Private Const ID_MAIN As String = "wnd[0]"
Private Const ID_OKCODE As String = "wnd[0]/tbar[0]/okcd"
Private Const ID_SBAR As String = "wnd[0]/sbar"
Private Const NOMBRE_TABLA_ITEMS As String = "SAPLMIGOTV_GOITEM"

' Estados de BAJAS_SAP
Private Const ST_PENDIENTE As String = "PENDIENTE"
Private Const ST_SIMULADO As String = "SIMULADO"
Private Const ST_CAPTURADO As String = "CAPTURADO"
Private Const ST_VERIFICADO As String = "VERIFICADO_SIN_CONTABILIZAR"
Private Const ST_CONTABILIZADO As String = "CONTABILIZADO"
Private Const ST_REVISAR As String = "REVISAR"
Private Const ST_ERROR As String = "ERROR"

Private Const ERR_CONFIG As Long = vbObjectError + 100
Private Const ERR_LAYOUT As Long = vbObjectError + 110
Private Const ERR_FATAL As Long = vbObjectError + 900

' Encabezados de BAJAS_SAP (en este orden; la macro los busca por nombre)
Private Const HEADERS_BAJAS As String = "FOLIO_APP|FECHA_DOC|CMV|MATERIAL|DESCRIPCION_MATERIAL|CANTIDAD|UME|CECO|DESTINATARIO|TEXTO|REFERENCIA|MOTIVO|NUMERO_EMPLEADO|NOMBRE|AREA_LINEA|STATUS_SAP|MENSAJE_SAP|FECHA_PROCESO|DOC_MATERIAL_SAP|OBSERVACIONES|VALIDACION|DETALLE_VALIDACION|LOTE"
Private Const HEADERS_LOG As String = "FECHA_HORA|LOTE|MODO|SISTEMA|MANDANTE|FILA|FOLIO_APP|MATERIAL|CANTIDAD|CECO|STATUS_SAP|MENSAJE_SAP"

' Longitudes maximas de los campos SAP
Private Const LEN_MATERIAL As Long = 40
Private Const LEN_CECO As Long = 10
Private Const LEN_DESTINATARIO As Long = 12   ' GOITEM-WEMPF
Private Const LEN_TEXTO As Long = 50          ' GOITEM-SGTXT

' Estado de Excel para restaurar
Private mPrevScreen As Boolean
Private mPrevCalc As Long
Private mPrevEvents As Boolean
Private mAppGuardada As Boolean

Private mTimeoutSeg As Double
Private mCols As Object          ' nombre tecnico de columna MIGO -> indice

' ============================================================
' API PUBLICA
' ============================================================

Public Sub Preparar_Bajas_Desde_App()
    Dim wsApp As Worksheet, wsB As Worksheet
    Dim src As Variant, outArr() As Variant
    Dim nApp As Long, i As Long, n As Long, c As Long
    Dim cFolio As Long, cFecha As Long, cNomina As Long, cNombre As Long, cArea As Long
    Dim cCeco As Long, cMat As Long, cSku As Long, cDesc As Long, cCat As Long
    Dim cTalla As Long, cCant As Long, cMotivo As Long, cOrigen As Long, cEstado As Long
    Dim previos As Object, vistos As Object, umes As Object
    Dim llave As String, prev As Variant, folio As String, material As String, motivo As String
    Dim nomina As String, nombre As String, perdidos As Long
    Dim hdr As Variant, k As Variant

    On Error GoTo ErrHandler
    GuardarApp

    Set wsApp = Hoja(SHEET_APP)
    Set wsB = Hoja(SHEET_BAJAS)

    cFolio = ColIdx(wsApp, "Folio")
    cFecha = ColIdx(wsApp, "Fecha")
    cNomina = ColIdx(wsApp, "Nomina", False)
    cNombre = ColIdx(wsApp, "Nombre", False)
    cArea = ColIdx(wsApp, "Area consumo", False)
    cCeco = ColIdx(wsApp, "Centro costo")
    cMat = ColIdx(wsApp, "Material")
    cSku = ColIdx(wsApp, "SKU", False)
    cDesc = ColIdx(wsApp, "Descripcion", False)
    cCat = ColIdx(wsApp, "Categoria", False)
    cTalla = ColIdx(wsApp, "Talla", False)
    cCant = ColIdx(wsApp, "Cantidad")
    cMotivo = ColIdx(wsApp, "Motivo")
    cOrigen = ColIdx(wsApp, "Origen", False)
    cEstado = ColIdx(wsApp, "Estado", False)

    src = LeerBloque(wsApp)
    nApp = ContarFilasApp(wsApp)

    ' Estado SAP previo por FOLIO|MATERIAL para no perder lo ya procesado
    Set previos = LeerEstadosPrevios(wsB)
    Set vistos = CreateObject("Scripting.Dictionary")
    Set umes = LeerUmes()

    hdr = Split(HEADERS_BAJAS, "|")
    If nApp > 0 Then ReDim outArr(1 To nApp, 1 To UBound(hdr) + 1)

    n = 0
    If IsArray(src) Then
        For i = 2 To UBound(src, 1)
            If FilaAppConDatos(src, i, cFolio, cMat, cCant) Then
                n = n + 1
                folio = Txt(src(i, cFolio))
                material = Txt(src(i, cMat))
                If material = "" And cSku > 0 Then material = Txt(src(i, cSku))
                motivo = Txt(src(i, cMotivo))
                nomina = TxtOpc(src, i, cNomina)
                nombre = TxtOpc(src, i, cNombre)

                outArr(n, B("FOLIO_APP")) = folio
                outArr(n, B("FECHA_DOC")) = FechaSap(src(i, cFecha))
                outArr(n, B("CMV")) = Cfg("CMV", "201")
                outArr(n, B("MATERIAL")) = material
                outArr(n, B("DESCRIPCION_MATERIAL")) = TxtOpc(src, i, cDesc)
                outArr(n, B("CANTIDAD")) = NumeroOTexto(src(i, cCant))
                outArr(n, B("UME")) = UmeDe(umes, TxtOpc(src, i, cCat))
                outArr(n, B("CECO")) = UCase$(Txt(src(i, cCeco)))
                outArr(n, B("DESTINATARIO")) = ValorSegun(Cfg("DESTINATARIO_DESDE", "MOTIVO"), motivo, nomina, nombre, folio)
                outArr(n, B("TEXTO")) = ValorSegun(Cfg("TEXTO_DESDE", "MOTIVO_FOLIO"), motivo, nomina, nombre, folio)
                outArr(n, B("REFERENCIA")) = ValorSegun(Cfg("REFERENCIA_DESDE", "FOLIO"), motivo, nomina, nombre, folio)
                outArr(n, B("MOTIVO")) = motivo
                outArr(n, B("NUMERO_EMPLEADO")) = nomina
                outArr(n, B("NOMBRE")) = nombre
                outArr(n, B("AREA_LINEA")) = TxtOpc(src, i, cArea)
                outArr(n, B("OBSERVACIONES")) = "Origen: " & TxtOpc(src, i, cOrigen) & " | Estado app: " & TxtOpc(src, i, cEstado) & " | Talla: " & TxtOpc(src, i, cTalla) & " | Fila app: " & i

                llave = UCase$(folio & "|" & material)
                If previos.Exists(llave) Then
                    prev = previos(llave)
                    For c = 0 To 4
                        outArr(n, B(Choose(c + 1, "STATUS_SAP", "MENSAJE_SAP", "FECHA_PROCESO", "DOC_MATERIAL_SAP", "LOTE"))) = prev(c)
                    Next c
                    previos.Remove llave
                Else
                    outArr(n, B("STATUS_SAP")) = ST_PENDIENTE
                End If

                If vistos.Exists(llave) Then
                    outArr(n, B("OBSERVACIONES")) = "DUPLICADO de la fila app " & vistos(llave) & " | " & outArr(n, B("OBSERVACIONES"))
                Else
                    vistos.Add llave, i
                End If
            End If
        Next i
    End If

    ' Filas ya contabilizadas que desaparecerian del pegado
    For Each k In previos.Keys
        prev = previos(k)
        If prev(0) = ST_CONTABILIZADO Or prev(0) = ST_VERIFICADO Then perdidos = perdidos + 1
    Next k
    If perdidos > 0 Then
        If MsgBox(perdidos & " fila(s) VERIFICADAS o CONTABILIZADAS de BAJAS_SAP ya no vienen en PEGAR_APP " & _
                  "y se perderia su estado." & vbCrLf & "Continuar de todos modos?", vbYesNo + vbExclamation) = vbNo Then GoTo Salir
    End If

    ' Escribir de un solo golpe (sin formulas)
    wsB.Cells.ClearContents
    wsB.Range(wsB.Cells(1, 1), wsB.Cells(1, UBound(hdr) + 1)).Value = FilaArreglo(hdr)
    If n > 0 Then wsB.Range(wsB.Cells(2, 1), wsB.Cells(n + 1, UBound(hdr) + 1)).Value = outArr

    ValidarTodo wsB, False

    If n <> nApp Then Err.Raise ERR_FATAL, , "Se esperaban " & nApp & " filas y se escribieron " & n & "."

    RestaurarApp
    MsgBox "BAJAS_SAP preparada: " & n & " fila(s) desde PEGAR_APP (" & nApp & " con datos).", vbInformation
    Exit Sub

Salir:
    RestaurarApp
    Exit Sub
ErrHandler:
    RestaurarApp
    MsgBox "No se pudo preparar BAJAS_SAP: " & Err.Description, vbCritical
End Sub

Public Sub Validar_Bajas_EPP()
    Dim ws As Worksheet, listos As Long, revisar As Long

    On Error GoTo ErrHandler
    GuardarApp
    Set ws = Hoja(SHEET_BAJAS)
    VerificarSincronia ws
    ValidarTodo ws, False, listos, revisar
    RestaurarApp
    MsgBox "Validacion terminada." & vbCrLf & "LISTO: " & listos & vbCrLf & "REVISAR: " & revisar, _
           IIf(revisar > 0, vbExclamation, vbInformation)
    Exit Sub
ErrHandler:
    RestaurarApp
    MsgBox "Validacion detenida: " & Err.Description, vbCritical
End Sub

Public Sub Procesar_Bajas_EPP_MIGO()
    Dim ws As Worksheet, session As Object
    Dim modo As String, lote As String, sid As String, mandt As String
    Dim centro As String, almacen As String
    Dim filas() As Long, nFilas As Long, previas As Long
    Dim capturadas() As Long, linea As Long
    Dim k As Long, r As Long, errores As Long, revisar As Long, verificadas As Long
    Dim tipo As String, msg As String, abortado As Boolean, motivoAborto As String
    Dim errNum As Long, errDesc As String

    On Error GoTo ErrGeneral

    Set ws = Hoja(SHEET_BAJAS)
    modo = UCase$(Cfg("MODO", "SIMULACION"))
    If modo <> "SIMULACION" And modo <> "VERIFICAR" Then
        Err.Raise ERR_CONFIG, , "CONFIG MODO debe ser SIMULACION o VERIFICAR (la macro nunca contabiliza)."
    End If
    mTimeoutSeg = Val(Cfg("SAP_TIMEOUT_SEG", "60"))
    If mTimeoutSeg <= 0 Then mTimeoutSeg = 60

    VerificarSincronia ws
    GuardarApp
    ValidarTodo ws, False
    Application.ScreenUpdating = True

    ' Candidatas: LISTO y no contabilizadas
    ReDim filas(1 To UltimaFila(ws))
    For r = 2 To UltimaFila(ws)
        If Celda(ws, r, "VALIDACION") = "LISTO" Then
            If Celda(ws, r, "STATUS_SAP") <> ST_CONTABILIZADO And Celda(ws, r, "DOC_MATERIAL_SAP") = "" Then
                nFilas = nFilas + 1
                filas(nFilas) = r
                If Celda(ws, r, "STATUS_SAP") = ST_VERIFICADO Then previas = previas + 1
            End If
        End If
    Next r

    If nFilas = 0 Then
        RestaurarApp
        MsgBox "No hay filas LISTO pendientes de SAP.", vbExclamation
        Exit Sub
    End If

    If previas > 0 Then
        If MsgBox(previas & " fila(s) ya se verificaron en una corrida anterior (" & ST_VERIFICADO & ")." & vbCrLf & _
                  "Si ya las contabilizaste a mano, registra el documento con Registrar_Contabilizacion_Manual " & _
                  "antes de seguir, o se capturarian de nuevo." & vbCrLf & vbCrLf & _
                  "Confirmas que NO estan contabilizadas?", vbYesNo + vbExclamation) = vbNo Then
            RestaurarApp
            Exit Sub
        End If
    End If

    lote = Format$(Now, "yyyymmdd-hhnnss")

    ' ---------------- SIMULACION (dry-run) ----------------
    If modo = "SIMULACION" Then
        For k = 1 To nFilas
            r = filas(k)
            SetEstado ws, r, ST_SIMULADO, "SIMULACION: se capturaria en MIGO linea " & k & _
                " (centro " & Cfg("CENTRO", "?") & ", almacen " & Cfg("ALMACEN", "?") & ", mov. " & Cfg("CMV", "201") & ").", lote
            RegistrarLog lote, modo, "-", "-", ws, r
        Next k
        RestaurarApp
        MsgBox "SIMULACION terminada. No se toco SAP." & vbCrLf & nFilas & " fila(s) se capturarian. Lote " & lote & ".", vbInformation
        Exit Sub
    End If

    ' ---------------- VERIFICAR ----------------
    sid = CfgRequerido("SAP_SISTEMA")
    mandt = CfgRequerido("SAP_MANDANTE")
    centro = CfgRequerido("CENTRO")
    almacen = CfgRequerido("ALMACEN")

    Set session = ObtenerSesionSAP(sid, mandt)

    If MsgBox("Se capturaran " & nFilas & " fila(s) en MIGO." & vbCrLf & _
              "Sistema: " & sid & "   Mandante: " & mandt & vbCrLf & _
              "Centro: " & centro & "   Almacen: " & almacen & "   Mov.: " & Cfg("CMV", "201") & vbCrLf & vbCrLf & _
              "Solo se VERIFICA; no se contabiliza. Continuar?", vbYesNo + vbQuestion, "Bajas EPP - MIGO") = vbNo Then
        RestaurarApp
        Exit Sub
    End If

    Application.ScreenUpdating = False
    PrepararMigo session

    ReDim capturadas(1 To nFilas)

    For k = 1 To nFilas
        r = filas(k)
        On Error GoTo ErrFila

        CapturarLinea session, ws, r, linea + 1, centro, almacen, tipo, msg

        Select Case tipo
            Case "E", "A", "X", "P"
                If Not LimpiarLinea(session, linea + 1) Then
                    Err.Raise ERR_FATAL, , "No pude limpiar la linea " & (linea + 1) & " despues del error: " & msg
                End If
                If tipo = "P" Then
                    SetEstado ws, r, ST_REVISAR, "SAP mostro una ventana al capturar: " & msg & ". Linea retirada del lote.", lote
                    revisar = revisar + 1
                Else
                    SetEstado ws, r, ST_ERROR, "SAP rechazo la linea: " & msg, lote
                    errores = errores + 1
                End If
            Case Else
                linea = linea + 1
                capturadas(linea) = r
                If msg = "" Then msg = "aceptada sin mensaje"
                SetEstado ws, r, ST_CAPTURADO, "Linea MIGO " & linea & ": " & msg, lote
        End Select
        RegistrarLog lote, modo, sid, mandt, ws, r

SiguienteFila:
        On Error GoTo ErrGeneral
        If abortado Then Exit For
        DoEvents
    Next k

    If abortado Then
        For k = 1 To linea
            SetEstado ws, capturadas(k), ST_REVISAR, "Lote abortado antes de verificar (" & motivoAborto & "). Sal de MIGO SIN grabar y vuelve a procesar.", lote
            RegistrarLog lote, modo, sid, mandt, ws, capturadas(k)
        Next k
        RestaurarApp
        MsgBox "Proceso detenido: " & motivoAborto & vbCrLf & "Sal de MIGO sin grabar. Nada se contabilizo.", vbCritical
        Exit Sub
    End If

    If linea = 0 Then
        RestaurarApp
        MsgBox "Ninguna fila quedo capturada en MIGO." & vbCrLf & "Errores: " & errores & "   Revisar: " & revisar, vbExclamation
        Exit Sub
    End If

    ' Verificar el documento completo y leer el resultado ANTES de cerrar ventanas
    PresionarVerificar session
    LeerMensajeSAP session, tipo, msg
    CerrarPopup session

    For k = 1 To linea
        r = capturadas(k)
        If tipo = "S" And Len(msg) > 0 Then
            SetEstado ws, r, ST_VERIFICADO, Celda(ws, r, "MENSAJE_SAP") & " | Verificacion: " & msg, lote
            verificadas = verificadas + 1
        ElseIf Len(msg) = 0 Then
            SetEstado ws, r, ST_REVISAR, Celda(ws, r, "MENSAJE_SAP") & " | Verificacion sin mensaje de SAP: revisar en pantalla.", lote
            revisar = revisar + 1
        Else
            SetEstado ws, r, ST_REVISAR, Celda(ws, r, "MENSAJE_SAP") & " | Verificacion (" & tipo & "): " & msg, lote
            revisar = revisar + 1
        End If
        RegistrarLog lote, modo, sid, mandt, ws, r
    Next k

    RestaurarApp
    MsgBox "Proceso terminado (lote " & lote & ")." & vbCrLf & _
           "Lineas en MIGO: " & linea & vbCrLf & _
           "Verificadas sin contabilizar: " & verificadas & vbCrLf & _
           "Revisar: " & revisar & vbCrLf & _
           "Errores: " & errores & vbCrLf & vbCrLf & _
           "Revisa MIGO y contabiliza a mano. Despues ejecuta Registrar_Contabilizacion_Manual con el lote " & lote & ".", _
           IIf(revisar + errores > 0, vbExclamation, vbInformation)
    Exit Sub

ErrFila:
    errNum = Err.Number
    errDesc = Err.Description
    If errNum = ERR_FATAL Then
        abortado = True
        motivoAborto = errDesc
        SetEstado ws, r, ST_ERROR, "Fatal: " & errDesc, lote
    Else
        SetEstado ws, r, ST_ERROR, errDesc & " | SAP: " & MensajeSeguro(session), lote
        errores = errores + 1
        If Not LimpiarLinea(session, linea + 1) Then
            abortado = True
            motivoAborto = "no pude limpiar la linea " & (linea + 1) & " tras un error"
        End If
    End If
    RegistrarLog lote, modo, sid, mandt, ws, r
    Resume SiguienteFila

ErrGeneral:
    errDesc = Err.Description
    RestaurarApp
    MsgBox "Error general: " & errDesc & vbCrLf & "Si MIGO quedo abierta, sal SIN grabar.", vbCritical
End Sub

Public Sub Registrar_Contabilizacion_Manual()
    Dim ws As Worksheet, lote As String, doc As String, r As Long, n As Long

    On Error GoTo ErrHandler
    Set ws = Hoja(SHEET_BAJAS)

    lote = Trim$(InputBox("Lote (columna LOTE) que contabilizaste a mano en MIGO:", "Registrar contabilizacion"))
    If lote = "" Then Exit Sub
    doc = Trim$(InputBox("Numero de documento de material que mostro SAP:", "Registrar contabilizacion"))
    If doc = "" Then Exit Sub
    If Not (doc Like String$(Len(doc), "#")) Or Len(doc) < 8 Then
        MsgBox "El documento debe ser numerico (8 a 10 digitos).", vbExclamation
        Exit Sub
    End If

    For r = 2 To UltimaFila(ws)
        If Celda(ws, r, "LOTE") = lote And Celda(ws, r, "STATUS_SAP") = ST_VERIFICADO Then
            ws.Cells(r, B("STATUS_SAP")).Value = ST_CONTABILIZADO
            ws.Cells(r, B("DOC_MATERIAL_SAP")).Value = "'" & doc
            ws.Cells(r, B("MENSAJE_SAP")).Value = "Contabilizado manualmente en MIGO. Doc " & doc
            ws.Cells(r, B("FECHA_PROCESO")).Value = Now
            RegistrarLog lote, "MANUAL", "-", "-", ws, r
            n = n + 1
        End If
    Next r

    MsgBox n & " fila(s) del lote " & lote & " marcadas como CONTABILIZADO.", IIf(n = 0, vbExclamation, vbInformation)
    Exit Sub
ErrHandler:
    MsgBox "Error: " & Err.Description, vbCritical
End Sub

Public Sub Probar_Conexion_SAP()
    Dim sapGui As Object, app As Object, conn As Object, ses As Object
    Dim txt As String, sid As String, mandt As String

    On Error GoTo ErrHandler
    sid = UCase$(Cfg("SAP_SISTEMA", ""))
    mandt = Cfg("SAP_MANDANTE", "")

    Set sapGui = GetObject("SAPGUI")
    Set app = sapGui.GetScriptingEngine
    For Each conn In app.Children
        For Each ses In conn.Children
            txt = txt & vbCrLf & ses.Info.SystemName & " / " & ses.Info.Client & " / " & ses.Info.Transaction & _
                  IIf(UCase$(ses.Info.SystemName) = sid And CStr(ses.Info.Client) = mandt, "   <- coincide con CONFIG", "")
        Next ses
    Next conn

    If txt = "" Then txt = vbCrLf & "(sin sesiones abiertas)"
    MsgBox "Sesiones SAP (sistema / mandante / transaccion):" & txt & vbCrLf & vbCrLf & _
           "CONFIG espera: " & sid & " / " & mandt, vbInformation
    Exit Sub
ErrHandler:
    MsgBox "No se pudo leer SAP GUI: " & Err.Description & vbCrLf & _
           "Verifica que SAP este abierto y que el scripting este habilitado.", vbCritical
End Sub

Public Sub Diagnosticar_Layout_MIGO()
    Dim session As Object, tbl As Object, ws As Worksheet, c As Long, nombre As String, titulo As String

    On Error GoTo ErrHandler
    mTimeoutSeg = Val(Cfg("SAP_TIMEOUT_SEG", "60"))
    If mTimeoutSeg <= 0 Then mTimeoutSeg = 60
    Set session = ObtenerSesionSAP(CfgRequerido("SAP_SISTEMA"), CfgRequerido("SAP_MANDANTE"))
    If session.Info.Transaction <> "MIGO" Then
        MsgBox "Abre MIGO (A07 / R10 / 201) en la sesion " & session.Info.SystemName & " y vuelve a ejecutar.", vbExclamation
        Exit Sub
    End If
    Set tbl = TablaItems(session)

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_DIAG)
    On Error GoTo ErrHandler
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = SHEET_DIAG
    End If
    ws.Cells.ClearContents
    ws.Range("A1:C1").Value = Array("INDICE", "NOMBRE_TECNICO", "TITULO")

    For c = 0 To tbl.Columns.Count - 1
        nombre = NombreColumna(tbl, c)
        titulo = ""
        On Error Resume Next
        titulo = tbl.Columns.ElementAt(c).Title
        On Error GoTo ErrHandler
        ws.Cells(c + 2, 1).Value = c
        ws.Cells(c + 2, 2).Value = nombre
        ws.Cells(c + 2, 3).Value = titulo
    Next c
    ws.Columns("A:C").AutoFit
    MsgBox "Columnas de la tabla de posiciones de MIGO en la hoja " & SHEET_DIAG & ".", vbInformation
    Exit Sub
ErrHandler:
    MsgBox "Diagnostico fallido: " & Err.Description, vbCritical
End Sub

' ============================================================
' SAP
' ============================================================

Private Function ObtenerSesionSAP(ByVal sid As String, ByVal mandt As String) As Object
    Dim sapGui As Object, app As Object, conn As Object, ses As Object, vistas As String

    On Error Resume Next
    Set sapGui = GetObject("SAPGUI")
    If Not sapGui Is Nothing Then Set app = sapGui.GetScriptingEngine
    On Error GoTo 0
    If app Is Nothing Then
        Err.Raise ERR_FATAL, , "No se pudo conectar a SAP GUI. Abre SAP, inicia sesion y habilita el scripting."
    End If

    For Each conn In app.Children
        For Each ses In conn.Children
            vistas = vistas & " " & ses.Info.SystemName & "/" & ses.Info.Client
            If UCase$(ses.Info.SystemName) = UCase$(sid) And CStr(ses.Info.Client) = CStr(mandt) Then
                If Not ses.Busy Then
                    Set ObtenerSesionSAP = ses
                    Exit Function
                End If
            End If
        Next ses
    Next conn

    Err.Raise ERR_FATAL, , "No hay una sesion libre de " & sid & " mandante " & mandt & _
        ". Sesiones abiertas:" & IIf(vistas = "", " ninguna", vistas)
End Function

Private Sub EsperarSAP(ByVal session As Object)
    Dim t0 As Double
    t0 = Timer
    Do While session.Busy
        DoEvents
        If SegundosDesde(t0) > mTimeoutSeg Then
            Err.Raise ERR_FATAL, , "SAP no respondio en " & mTimeoutSeg & " s."
        End If
    Loop
End Sub

Private Function SegundosDesde(ByVal t0 As Double) As Double
    SegundosDesde = Timer - t0
    If SegundosDesde < 0 Then SegundosDesde = SegundosDesde + 86400#
End Function

Private Sub PrepararMigo(ByVal session As Object)
    Dim tipo As String, msg As String

    session.findById(ID_MAIN).maximize
    session.findById(ID_OKCODE).Text = "/nMIGO"
    session.findById(ID_MAIN).sendVKey 0
    EsperarSAP session
    CerrarPopup session

    If session.Info.Transaction <> "MIGO" Then
        Err.Raise ERR_FATAL, , "No se pudo abrir MIGO (transaccion actual: " & session.Info.Transaction & ")."
    End If

    ' Cabecera desde CONFIG: operacion, referencia y clase de movimiento
    SetCampoCabecera session, "GODYNPRO-ACTION", "GuiComboBox", Cfg("MIGO_ACCION", "A07"), Cfg("MIGO_ID_ACCION", "")
    SetCampoCabecera session, "GODYNPRO-REFDOC", "GuiComboBox", Cfg("MIGO_REFERENCIA", "R10"), Cfg("MIGO_ID_REFERENCIA", "")
    SetCampoCabecera session, "GODEFAULT_TV-BWART", "GuiCTextField", Cfg("CMV", "201"), Cfg("MIGO_ID_CMV", "")
    session.findById(ID_MAIN).sendVKey 0
    EsperarSAP session

    LeerMensajeSAP session, tipo, msg
    If tipo = "E" Or tipo = "A" Or tipo = "X" Or tipo = "P" Then
        CerrarPopup session
        Err.Raise ERR_FATAL, , "MIGO rechazo la cabecera: " & msg
    End If

    CargarColumnas session
End Sub

Private Sub SetCampoCabecera(ByVal session As Object, ByVal nombre As String, ByVal tipoGui As String, _
                            ByVal valor As String, ByVal idAlterno As String)
    Dim ctl As Object

    On Error Resume Next
    Set ctl = session.findById("wnd[0]/usr").FindByName(nombre, tipoGui)
    If ctl Is Nothing And Len(idAlterno) > 0 Then Set ctl = session.findById(idAlterno)
    On Error GoTo 0

    If ctl Is Nothing Then
        Err.Raise ERR_FATAL, , "No encontre " & nombre & " en la cabecera de MIGO. Graba su ID y ponlo en CONFIG."
    End If

    If tipoGui = "GuiComboBox" Then
        If ctl.Key <> valor Then ctl.Key = valor
    Else
        ctl.Text = valor
    End If
    EsperarSAP session
End Sub

Private Function TablaItems(ByVal session As Object) As Object
    On Error Resume Next
    Set TablaItems = session.findById("wnd[0]/usr").FindByName(NOMBRE_TABLA_ITEMS, "GuiTableControl")
    On Error GoTo 0
    If TablaItems Is Nothing Then
        Err.Raise ERR_FATAL, , "No encontre la tabla de posiciones de MIGO (" & NOMBRE_TABLA_ITEMS & ")."
    End If
End Function

Private Function NombreColumna(ByVal tbl As Object, ByVal c As Long) As String
    On Error Resume Next
    NombreColumna = UCase$(tbl.GetCell(0, c).Name)
    On Error GoTo 0
End Function

Private Sub CargarColumnas(ByVal session As Object)
    Dim tbl As Object, c As Long, nombre As String

    Set tbl = TablaItems(session)
    Set mCols = CreateObject("Scripting.Dictionary")
    For c = 0 To tbl.Columns.Count - 1
        nombre = NombreColumna(tbl, c)
        If Len(nombre) > 0 Then
            If Not mCols.Exists(nombre) Then mCols.Add nombre, c
        End If
    Next c
End Sub

Private Function FilaVisible(ByVal session As Object, ByVal linea As Long) As Long
    Dim tbl As Object, visibles As Long, objetivo As Long

    Set tbl = TablaItems(session)
    visibles = tbl.VisibleRowCount
    If visibles <= 0 Then visibles = 1

    If linea <= visibles Then
        objetivo = 0
        FilaVisible = linea - 1
    Else
        objetivo = linea - visibles
        FilaVisible = visibles - 1
    End If

    If tbl.VerticalScrollbar.Position <> objetivo Then
        tbl.VerticalScrollbar.Position = objetivo
        EsperarSAP session
    End If
End Function

Private Sub SetCelda(ByVal session As Object, ByVal nombres As Variant, ByVal fila As Long, _
                     ByVal valor As String, ByVal etiqueta As String)
    Dim i As Long, tbl As Object

    For i = LBound(nombres) To UBound(nombres)
        If mCols.Exists(UCase$(CStr(nombres(i)))) Then
            Set tbl = TablaItems(session)
            tbl.GetCell(fila, mCols(UCase$(CStr(nombres(i))))).Text = valor
            Exit Sub
        End If
    Next i
    Err.Raise ERR_FATAL, , "La columna " & etiqueta & " no esta en la tabla de MIGO. Ejecuta Diagnosticar_Layout_MIGO."
End Sub

Private Sub CapturarLinea(ByVal session As Object, ByVal ws As Worksheet, ByVal r As Long, ByVal linea As Long, _
                          ByVal centro As String, ByVal almacen As String, ByRef tipo As String, ByRef msg As String)
    Dim fila As Long, tipo2 As String, msg2 As String

    fila = FilaVisible(session, linea)

    SetCelda session, Array("GOITEM-MATNR", "GOITEM-MAKTX"), fila, Celda(ws, r, "MATERIAL"), "Material"
    SetCelda session, Array("GOITEM-ERFMG"), fila, Replace(Celda(ws, r, "CANTIDAD"), ",", "."), "Cantidad"
    If centro <> "*" Then SetCelda session, Array("GOITEM-WERKS", "GOITEM-NAME1"), fila, centro, "Centro"
    SetCelda session, Array("GOITEM-LGORT", "GOITEM-LGOBE"), fila, almacen, "Almacen"
    SetCelda session, Array("GOITEM-KOSTL"), fila, Celda(ws, r, "CECO"), "Centro de costo"
    SetCelda session, Array("GOITEM-SGTXT"), fila, Celda(ws, r, "TEXTO"), "Texto"
    SetCelda session, Array("GOITEM-WEMPF"), fila, Celda(ws, r, "DESTINATARIO"), "Destinatario"

    session.findById(ID_MAIN).sendVKey 0
    EsperarSAP session
    LeerMensajeSAP session, tipo, msg      ' se lee ANTES de cerrar cualquier ventana

    If tipo = "W" Then                     ' advertencia: SAP pide confirmar con Enter
        session.findById(ID_MAIN).sendVKey 0
        EsperarSAP session
        LeerMensajeSAP session, tipo2, msg2
        msg = "Aviso: " & msg & IIf(Len(msg2) > 0, " | " & msg2, "")
        If Len(tipo2) > 0 And tipo2 <> "W" Then tipo = tipo2 Else tipo = "W"
    End If

    If tipo = "P" Then CerrarPopup session
End Sub

' Borra la linea capturada para que no entre al lote. Devuelve False si SAP sigue con error.
Private Function LimpiarLinea(ByVal session As Object, ByVal linea As Long) As Boolean
    Dim fila As Long, tbl As Object, k As Variant, tipo As String, msg As String

    On Error GoTo Fallo
    CerrarPopup session
    fila = FilaVisible(session, linea)
    Set tbl = TablaItems(session)
    For Each k In Array("GOITEM-WEMPF", "GOITEM-SGTXT", "GOITEM-KOSTL", "GOITEM-ERFMG", "GOITEM-MATNR", "GOITEM-MAKTX")
        If mCols.Exists(CStr(k)) Then
            On Error Resume Next
            tbl.GetCell(fila, mCols(CStr(k))).Text = ""
            On Error GoTo Fallo
        End If
    Next k
    session.findById(ID_MAIN).sendVKey 0
    EsperarSAP session
    LeerMensajeSAP session, tipo, msg
    CerrarPopup session
    LimpiarLinea = Not (tipo = "E" Or tipo = "A" Or tipo = "X")
    Exit Function
Fallo:
    LimpiarLinea = False
End Function

Private Sub PresionarVerificar(ByVal session As Object)
    Dim idBoton As String
    idBoton = Cfg("MIGO_BOTON_VERIFICAR", "wnd[0]/tbar[1]/btn[7]")
    ' Por seguridad solo se presiona el boton configurado de Verificar; nunca Contabilizar.
    session.findById(idBoton).press
    EsperarSAP session
End Sub

' Lee tipo y texto de la barra de estado y, si hay ventana emergente, su texto.
' Tipo "P" = hubo ventana emergente sin mensaje en la barra.
Private Sub LeerMensajeSAP(ByVal session As Object, ByRef tipo As String, ByRef msg As String)
    Dim popup As String

    tipo = ""
    msg = ""
    On Error Resume Next
    tipo = UCase$(Trim$(session.findById(ID_SBAR).MessageType))
    msg = Trim$(session.findById(ID_SBAR).Text)
    If session.Children.Count > 1 Then
        popup = TextoPopup(session)
        If Len(popup) > 0 Then msg = msg & IIf(Len(msg) > 0, " | ", "") & "Ventana: " & popup
        If tipo = "" Then tipo = "P"
    End If
    On Error GoTo 0
End Sub

Private Function TextoPopup(ByVal session As Object) As String
    Dim w As Object, s As String
    On Error Resume Next
    Set w = session.findById("wnd[1]")
    s = w.Text
    s = s & " " & TextosHijos(w.findById("usr"), 0)
    TextoPopup = Left$(Application.WorksheetFunction.Trim(s), 250)
    On Error GoTo 0
End Function

Private Function TextosHijos(ByVal cont As Object, ByVal nivel As Long) As String
    Dim ch As Object, s As String
    On Error Resume Next
    If nivel > 4 Then Exit Function
    For Each ch In cont.Children
        Select Case ch.Type
            Case "GuiLabel", "GuiTextField", "GuiTextedit"
                If Len(Trim$(ch.Text)) > 0 Then s = s & " " & Trim$(ch.Text)
            Case "GuiSimpleContainer", "GuiUserArea", "GuiScrollContainer", "GuiBox"
                s = s & TextosHijos(ch, nivel + 1)
        End Select
        If Len(s) > 400 Then Exit For
    Next ch
    TextosHijos = s
End Function

Private Sub CerrarPopup(ByVal session As Object)
    On Error Resume Next
    Do While session.Children.Count > 1
        session.findById("wnd[" & (session.Children.Count - 1) & "]").Close
        EsperarSAP session
        If Err.Number <> 0 Then Exit Do
    Loop
    On Error GoTo 0
End Sub

Private Function MensajeSeguro(ByVal session As Object) As String
    Dim t As String, m As String
    On Error Resume Next
    If Not session Is Nothing Then LeerMensajeSAP session, t, m
    If Len(m) = 0 Then m = "sin mensaje"
    MensajeSeguro = m
End Function

' ============================================================
' VALIDACION Y SINCRONIA
' ============================================================

Private Sub VerificarSincronia(ByVal ws As Worksheet)
    Dim nApp As Long, nB As Long
    nApp = ContarFilasApp(Hoja(SHEET_APP))
    nB = UltimaFila(ws) - 1
    If nB < 0 Then nB = 0
    If HayFormulas(ws) Then
        Err.Raise ERR_FATAL, , "BAJAS_SAP tiene formulas. Ejecuta Preparar_Bajas_Desde_App (escribe valores)."
    End If
    If nApp <> nB Then
        Err.Raise ERR_FATAL, , "PEGAR_APP tiene " & nApp & " fila(s) con datos y BAJAS_SAP tiene " & nB & _
            ". Ejecuta Preparar_Bajas_Desde_App antes de validar o procesar."
    End If
End Sub

Private Sub ValidarTodo(ByVal ws As Worksheet, ByVal mostrar As Boolean, _
                       Optional ByRef listos As Long, Optional ByRef revisar As Long)
    Dim data As Variant, r As Long, det As String, ult As Long, outV() As Variant
    Dim vMat As String, vCant As Variant, vCeco As String, vDest As String, vTxt As String, vFolio As String

    ult = UltimaFila(ws)
    listos = 0
    revisar = 0
    If ult < 2 Then Exit Sub

    data = ws.Range(ws.Cells(2, 1), ws.Cells(ult, B("LOTE"))).Value
    ReDim outV(1 To ult - 1, 1 To 2)

    For r = 1 To ult - 1
        det = ""
        vFolio = Txt(data(r, B("FOLIO_APP")))
        vMat = Txt(data(r, B("MATERIAL")))
        vCant = data(r, B("CANTIDAD"))
        vCeco = Txt(data(r, B("CECO")))
        vDest = Txt(data(r, B("DESTINATARIO")))
        vTxt = Txt(data(r, B("TEXTO")))

        If vFolio = "" Then det = det & "Falta FOLIO; "
        If vMat = "" Then det = det & "Falta MATERIAL; "
        If Len(vMat) > LEN_MATERIAL Then det = det & "MATERIAL excede " & LEN_MATERIAL & "; "
        If Not IsNumeric(vCant) Then
            det = det & "CANTIDAD invalida; "
        ElseIf CDbl(vCant) <= 0 Then
            det = det & "CANTIDAD debe ser mayor a 0; "
        End If
        If vCeco = "" Then det = det & "Falta CECO; "
        If Len(vCeco) > LEN_CECO Then det = det & "CECO excede " & LEN_CECO & "; "
        If vDest = "" Then det = det & "Falta DESTINATARIO; "
        If Len(vDest) > LEN_DESTINATARIO Then det = det & "DESTINATARIO excede " & LEN_DESTINATARIO & " (WEMPF); "
        If vTxt = "" Then det = det & "Falta TEXTO; "
        If Len(vTxt) > LEN_TEXTO Then det = det & "TEXTO excede " & LEN_TEXTO & "; "
        If InStr(1, Txt(data(r, B("OBSERVACIONES"))), "DUPLICADO", vbTextCompare) > 0 Then det = det & "Folio+material duplicado; "

        If det = "" Then
            outV(r, 1) = "LISTO"
            outV(r, 2) = ""
            listos = listos + 1
        Else
            outV(r, 1) = "REVISAR"
            outV(r, 2) = det
            revisar = revisar + 1
        End If
    Next r

    ws.Range(ws.Cells(2, B("VALIDACION")), ws.Cells(ult, B("DETALLE_VALIDACION"))).Value = outV
    If mostrar Then MsgBox "LISTO: " & listos & "   REVISAR: " & revisar, vbInformation
End Sub

Private Function HayFormulas(ByVal ws As Worksheet) As Boolean
    Dim rng As Range
    On Error Resume Next
    Set rng = ws.UsedRange.SpecialCells(xlCellTypeFormulas)
    On Error GoTo 0
    HayFormulas = Not rng Is Nothing
End Function

' ============================================================
' HOJAS, CONFIG Y LOG
' ============================================================

Private Function Hoja(ByVal nombre As String) As Worksheet
    On Error Resume Next
    Set Hoja = ThisWorkbook.Worksheets(nombre)
    On Error GoTo 0
    If Hoja Is Nothing Then Err.Raise ERR_LAYOUT, , "No existe la hoja " & nombre & "."
End Function

Private Function Cfg(ByVal clave As String, Optional ByVal defecto As String = "") As String
    Dim ws As Worksheet, r As Long, v As Variant

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CONFIG)
    On Error GoTo 0
    Cfg = defecto
    If ws Is Nothing Then Exit Function

    For r = 1 To ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
        If UCase$(Trim$(Txt(ws.Cells(r, 1).Value))) = UCase$(clave) Then
            v = ws.Cells(r, 2).Value
            If Len(Txt(v)) > 0 Then Cfg = Trim$(Txt(v))
            Exit Function
        End If
    Next r
End Function

Private Function CfgRequerido(ByVal clave As String) As String
    Dim v As String
    v = Cfg(clave, "")
    If Len(v) = 0 Or Left$(v, 1) = "<" Then
        Err.Raise ERR_CONFIG, , "CONFIG: captura un valor real para " & clave & " (hoy: '" & v & "')."
    End If
    CfgRequerido = v
End Function

Private Sub RegistrarLog(ByVal lote As String, ByVal modo As String, ByVal sid As String, ByVal mandt As String, _
                         ByVal ws As Worksheet, ByVal r As Long)
    Dim lg As Worksheet, nxt As Long, hdr As Variant

    On Error Resume Next
    Set lg = ThisWorkbook.Worksheets(SHEET_LOG)
    If lg Is Nothing Then
        Set lg = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        lg.Name = SHEET_LOG
    End If
    hdr = Split(HEADERS_LOG, "|")
    If Txt(lg.Cells(1, 1).Value) = "" Then lg.Range(lg.Cells(1, 1), lg.Cells(1, UBound(hdr) + 1)).Value = FilaArreglo(hdr)

    nxt = lg.Cells(lg.Rows.Count, 1).End(xlUp).Row + 1
    lg.Range(lg.Cells(nxt, 1), lg.Cells(nxt, UBound(hdr) + 1)).Value = Array( _
        Now, "'" & lote, modo, sid, "'" & mandt, r, _
        Celda(ws, r, "FOLIO_APP"), Celda(ws, r, "MATERIAL"), Celda(ws, r, "CANTIDAD"), Celda(ws, r, "CECO"), _
        Celda(ws, r, "STATUS_SAP"), Celda(ws, r, "MENSAJE_SAP"))
    On Error GoTo 0
End Sub

Private Sub SetEstado(ByVal ws As Worksheet, ByVal r As Long, ByVal estado As String, ByVal msg As String, ByVal lote As String)
    ws.Cells(r, B("STATUS_SAP")).Value = estado
    ws.Cells(r, B("MENSAJE_SAP")).Value = Left$(msg, 500)
    ws.Cells(r, B("FECHA_PROCESO")).Value = Now
    ws.Cells(r, B("LOTE")).Value = "'" & lote
End Sub

Private Sub GuardarApp()
    If mAppGuardada Then Exit Sub
    mPrevScreen = Application.ScreenUpdating
    mPrevCalc = Application.Calculation
    mPrevEvents = Application.EnableEvents
    mAppGuardada = True
    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.Calculation = xlCalculationManual
End Sub

Private Sub RestaurarApp()
    If Not mAppGuardada Then Exit Sub
    On Error Resume Next
    Application.Calculation = mPrevCalc
    Application.EnableEvents = mPrevEvents
    Application.ScreenUpdating = mPrevScreen
    Application.StatusBar = False
    mAppGuardada = False
    On Error GoTo 0
End Sub

' ============================================================
' UTILERIAS DE DATOS
' ============================================================

' Indice de columna (1-based) en BAJAS_SAP segun HEADERS_BAJAS
Private Function B(ByVal nombre As String) As Long
    Dim hdr As Variant, i As Long
    hdr = Split(HEADERS_BAJAS, "|")
    For i = 0 To UBound(hdr)
        If hdr(i) = nombre Then B = i + 1: Exit Function
    Next i
    Err.Raise ERR_LAYOUT, , "Columna interna desconocida: " & nombre
End Function

Private Function Celda(ByVal ws As Worksheet, ByVal r As Long, ByVal nombre As String) As String
    Celda = Trim$(Txt(ws.Cells(r, B(nombre)).Value))
End Function

Private Function UltimaFila(ByVal ws As Worksheet) As Long
    UltimaFila = ws.Cells(ws.Rows.Count, B("FOLIO_APP")).End(xlUp).Row
    Dim u2 As Long
    u2 = ws.Cells(ws.Rows.Count, B("MATERIAL")).End(xlUp).Row
    If u2 > UltimaFila Then UltimaFila = u2
End Function

Private Function ColIdx(ByVal ws As Worksheet, ByVal nombre As String, Optional ByVal requerido As Boolean = True) As Long
    Dim c As Long, ult As Long
    ult = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    For c = 1 To ult
        If NormEnc(ws.Cells(1, c).Value) = NormEnc(nombre) Then ColIdx = c: Exit Function
    Next c
    If requerido Then Err.Raise ERR_LAYOUT, , "No encontre la columna '" & nombre & "' en la fila 1 de " & ws.Name & "."
End Function

Private Function NormEnc(ByVal v As Variant) As String
    Dim s As String
    s = UCase$(Trim$(Txt(v)))
    s = Replace(s, ChrW(193), "A"): s = Replace(s, ChrW(201), "E"): s = Replace(s, ChrW(205), "I")
    s = Replace(s, ChrW(211), "O"): s = Replace(s, ChrW(218), "U"): s = Replace(s, ChrW(209), "N")
    s = Replace(s, ChrW(225), "A"): s = Replace(s, ChrW(233), "E"): s = Replace(s, ChrW(237), "I")
    s = Replace(s, ChrW(243), "O"): s = Replace(s, ChrW(250), "U"): s = Replace(s, ChrW(241), "N")
    s = Replace(s, " ", "_")
    NormEnc = s
End Function

Private Function LeerBloque(ByVal ws As Worksheet) As Variant
    Dim ult As Long, ultCol As Long
    ult = ws.Cells.Find(What:="*", SearchOrder:=xlByRows, SearchDirection:=xlPrevious).Row
    ultCol = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
    If ult < 2 Then
        LeerBloque = Empty
    Else
        LeerBloque = ws.Range(ws.Cells(1, 1), ws.Cells(ult, ultCol)).Value
    End If
End Function

Private Function ContarFilasApp(ByVal wsApp As Worksheet) As Long
    Dim src As Variant, i As Long, cF As Long, cM As Long, cC As Long
    If Application.WorksheetFunction.CountA(wsApp.Cells) <= 1 Then Exit Function
    cF = ColIdx(wsApp, "Folio")
    cM = ColIdx(wsApp, "Material")
    cC = ColIdx(wsApp, "Cantidad")
    src = LeerBloque(wsApp)
    If Not IsArray(src) Then Exit Function
    For i = 2 To UBound(src, 1)
        If FilaAppConDatos(src, i, cF, cM, cC) Then ContarFilasApp = ContarFilasApp + 1
    Next i
End Function

Private Function FilaAppConDatos(ByVal src As Variant, ByVal i As Long, ByVal cF As Long, ByVal cM As Long, ByVal cC As Long) As Boolean
    FilaAppConDatos = (Txt(src(i, cF)) <> "" Or Txt(src(i, cM)) <> "" Or Txt(src(i, cC)) <> "")
End Function

Private Function LeerEstadosPrevios(ByVal wsB As Worksheet) As Object
    Dim d As Object, r As Long, k As String, ult As Long, data As Variant
    Set d = CreateObject("Scripting.Dictionary")
    Set LeerEstadosPrevios = d
    If NormEnc(wsB.Cells(1, 1).Value) <> "FOLIO_APP" Then Exit Function
    If NormEnc(wsB.Cells(1, B("LOTE")).Value) <> "LOTE" Then Exit Function
    ult = UltimaFila(wsB)
    If ult < 2 Then Exit Function
    data = wsB.Range(wsB.Cells(2, 1), wsB.Cells(ult, B("LOTE"))).Value
    For r = 1 To ult - 1
        k = UCase$(Txt(data(r, B("FOLIO_APP"))) & "|" & Txt(data(r, B("MATERIAL"))))
        If Txt(data(r, B("STATUS_SAP"))) <> "" And Not d.Exists(k) Then
            d.Add k, Array(Txt(data(r, B("STATUS_SAP"))), Txt(data(r, B("MENSAJE_SAP"))), _
                           data(r, B("FECHA_PROCESO")), Txt(data(r, B("DOC_MATERIAL_SAP"))), Txt(data(r, B("LOTE"))))
        End If
    Next r
End Function

Private Function LeerUmes() As Object
    Dim d As Object, ws As Worksheet, cCat As Long, cUme As Long, r As Long
    Set d = CreateObject("Scripting.Dictionary")
    d.CompareMode = 1
    Set LeerUmes = d
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(SHEET_CATALOGOS)
    On Error GoTo 0
    If ws Is Nothing Then Exit Function
    cCat = ColIdx(ws, "Categoria", False)
    cUme = ColIdx(ws, "UME", False)
    If cCat = 0 Or cUme = 0 Then Exit Function
    For r = 2 To ws.Cells(ws.Rows.Count, cCat).End(xlUp).Row
        If Txt(ws.Cells(r, cCat).Value) <> "" And Not d.Exists(Txt(ws.Cells(r, cCat).Value)) Then
            d.Add Txt(ws.Cells(r, cCat).Value), Txt(ws.Cells(r, cUme).Value)
        End If
    Next r
End Function

Private Function UmeDe(ByVal umes As Object, ByVal categoria As String) As String
    If umes.Exists(categoria) Then UmeDe = umes(categoria) Else UmeDe = Cfg("UME_DEFAULT", "PZA")
End Function

Private Function ValorSegun(ByVal opcion As String, ByVal motivo As String, ByVal nomina As String, _
                           ByVal nombre As String, ByVal folio As String) As String
    Select Case UCase$(opcion)
        Case "MOTIVO": ValorSegun = LCase$(motivo)
        Case "NOMINA": ValorSegun = nomina
        Case "NOMINA_NOMBRE": ValorSegun = nomina & " - " & nombre
        Case "FOLIO": ValorSegun = folio
        Case "MOTIVO_FOLIO": ValorSegun = LCase$(motivo) & " F:" & folio
        Case Else: Err.Raise ERR_CONFIG, , "CONFIG: opcion no valida '" & opcion & "'. Usa MOTIVO, NOMINA, NOMINA_NOMBRE, FOLIO o MOTIVO_FOLIO."
    End Select
End Function

Private Function FechaSap(ByVal v As Variant) As String
    If IsDate(v) Then
        FechaSap = Format$(CDate(v), "dd.mm.yyyy")
    Else
        FechaSap = Txt(v)
    End If
End Function

Private Function NumeroOTexto(ByVal v As Variant) As Variant
    If IsNumeric(v) Then NumeroOTexto = CDbl(v) Else NumeroOTexto = Txt(v)
End Function

Private Function Txt(ByVal v As Variant) As String
    If IsError(v) Or IsNull(v) Or IsEmpty(v) Then
        Txt = ""
    Else
        Txt = Trim$(Replace(Replace(CStr(v), vbCr, " "), vbLf, " "))
    End If
End Function

Private Function TxtOpc(ByVal src As Variant, ByVal i As Long, ByVal c As Long) As String
    If c > 0 Then TxtOpc = Txt(src(i, c))
End Function

Private Function FilaArreglo(ByVal v As Variant) As Variant
    Dim a() As Variant, i As Long
    ReDim a(1 To 1, 1 To UBound(v) - LBound(v) + 1)
    For i = LBound(v) To UBound(v)
        a(1, i - LBound(v) + 1) = v(i)
    Next i
    FilaArreglo = a
End Function
