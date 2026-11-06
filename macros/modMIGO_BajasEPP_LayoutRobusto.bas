Attribute VB_Name = "modMIGO_BajasEPP_LayoutRobusto"
Option Explicit

' ============================================================
' DESARROLLADO POR JOSE LUIS CRUZ PRIETO:
' Libro_Maestro_Bajas_EPP_SAP_Ajustado_App
' VERSION ROBUSTA MULTIUSUARIO - MIGO
' Evita depender de posiciones fijas de columnas en el layout.
' Favor de no modificar el codigo
' Lee BAJAS_SAP, NO PEGAR_APP.
' Escribe resultados en:
'   P = STATUS_SAP
'   Q = MENSAJE_SAP
'   R = FECHA_PROCESO
'   S = DOC_MATERIAL_SAP
' ============================================================

Private Const SHEET_BAJAS As String = "BAJAS_SAP"
Private Const SHEET_LOG As String = "LOG_SAP"

Private Const FALLBACK_ALMACEN As String = "<ALMACEN>"

Private Const ID_OK_CODE As String = "wnd[0]/tbar[0]/okcd"
Private Const ID_MAIN_WINDOW As String = "wnd[0]"
Private Const ID_STATUS_BAR As String = "wnd[0]/sbar"

Private Const ID_ITEM_TABLE As String = _
    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM/"
Private Const ID_ITEM_TABLE_OBJ As String = _
    "wnd[0]/usr/ssubSUB_MAIN_CARRIER:SAPLMIGO:0002/subSUB_ITEMLIST:SAPLMIGO:0200/tblSAPLMIGOTV_GOITEM"

' Columnas reales de BAJAS_SAP
Private Const COL_FOLIO As String = "A"
Private Const COL_MATERIAL As String = "D"
Private Const COL_CANTIDAD As String = "F"
Private Const COL_CECO As String = "H"
Private Const COL_DESTINATARIO As String = "I"
Private Const COL_TEXTO As String = "J"
Private Const COL_REFERENCIA As String = "K"
Private Const COL_STATUS As String = "P"
Private Const COL_MENSAJE As String = "Q"
Private Const COL_FECHA_PROC As String = "R"
Private Const COL_DOC_MAT As String = "S"
Private Const COL_VALIDACION As String = "U"
Private Const COL_DETALLE_VALIDACION As String = "V"
Public Sub Validar_Bajas_EPP()
    Dim ws As Worksheet
    Dim lastRow As Long

    Set ws = ThisWorkbook.Worksheets(SHEET_BAJAS)
    lastRow = ws.Cells(ws.Rows.Count, COL_FOLIO).End(xlUp).Row

    If ValidarLayout(ws, lastRow, True) Then
        MsgBox "Validacion exitosa. Las filas LISTO estan preparadas para MIGO.", vbInformation
    End If
End Sub
Public Sub Procesar_Bajas_EPP_MIGO()
    Dim ws As Worksheet
    Dim session As Object
    Dim lastRow As Long
    Dim i As Long
    Dim n As Long
    Dim estado As String
    Dim msgSap As String
    Dim tipoSap As String
    Dim procesadas As Long
    Dim omitidas As Long
    Dim errores As Long
    Dim itemIndex As Long
    Dim lineaMigo As Long
    Dim filasProcesadas() As Long

    On Error GoTo ErrGeneral

    Set ws = ThisWorkbook.Worksheets(SHEET_BAJAS)
    lastRow = ws.Cells(ws.Rows.Count, COL_FOLIO).End(xlUp).Row

    If lastRow < 2 Then
        MsgBox "No hay registros para procesar en " & SHEET_BAJAS & ".", vbExclamation
        Exit Sub
    End If

    If Not ValidarLayout(ws, lastRow, True) Then Exit Sub

    ReDim filasProcesadas(1 To lastRow - 1)

    Set session = GetSapSession()

    ' Abrir MIGO una sola vez
    PrepararMigoA07 session

    For i = 2 To lastRow

        If Trim$(CStr(ws.Cells(i, COL_FOLIO).Value)) = "" Then
            omitidas = omitidas + 1
            GoTo SiguienteFila
        End If

        estado = UCase$(Trim$(CStr(ws.Cells(i, COL_STATUS).Value)))

        If estado = "VERIFICADO" Or estado = "OK" Or estado = "CONTABILIZADO" Then
            omitidas = omitidas + 1
            GoTo SiguienteFila
        End If

        If UCase$(Trim$(CStr(ws.Cells(i, COL_VALIDACION).Value))) <> "LISTO" Then
            ws.Cells(i, COL_STATUS).Value = "REVISAR"
            ws.Cells(i, COL_MENSAJE).Value = CStr(ws.Cells(i, COL_DETALLE_VALIDACION).Value)
            ws.Cells(i, COL_FECHA_PROC).Value = Now
            omitidas = omitidas + 1
            GoTo SiguienteFila
        End If

        On Error GoTo ErrFila

        lineaMigo = itemIndex + 1

        ' Captura en línea dinámica de MIGO
        CapturarFilaMigoA07_Lote session, ws, i, lineaMigo

        itemIndex = lineaMigo
        filasProcesadas(itemIndex) = i

        ws.Cells(i, COL_STATUS).Value = "CAPTURADO"
        ws.Cells(i, COL_MENSAJE).Value = "Capturado en MIGO línea " & itemIndex & ". Pendiente de verificar."
        ws.Cells(i, COL_FECHA_PROC).Value = Now

        On Error GoTo ErrGeneral
        GoTo SiguienteFila

ErrFila:
        ws.Cells(i, COL_STATUS).Value = "ERROR"
        ws.Cells(i, COL_MENSAJE).Value = Err.Description & " | SAP: " & SafeSapMessage(session)
        ws.Cells(i, COL_FECHA_PROC).Value = Now
        errores = errores + 1
        Err.Clear
        On Error GoTo ErrGeneral

SiguienteFila:
        DoEvents
    Next i

    If itemIndex = 0 Then
        MsgBox "No hubo filas LISTO para capturar en MIGO." & vbCrLf & _
               "Omitidas: " & omitidas & vbCrLf & _
               "Errores: " & errores, vbExclamation
        Exit Sub
    End If

    ' Verificar una sola vez, después de capturar todos los materiales
    PressVerificarMigo session
    EsperarSAP session
    CerrarPopupSiExiste session

    msgSap = GetSapMessage(session)
    tipoSap = GetSapMessageType(session)

    For n = 1 To itemIndex
        i = filasProcesadas(n)

        If tipoSap = "E" Or tipoSap = "A" Then
            ws.Cells(i, COL_STATUS).Value = "ERROR"
            errores = errores + 1
        Else
            ws.Cells(i, COL_STATUS).Value = "VERIFICADO"
            procesadas = procesadas + 1
        End If

        ws.Cells(i, COL_MENSAJE).Value = msgSap
        ws.Cells(i, COL_FECHA_PROC).Value = Now
        ws.Cells(i, COL_DOC_MAT).Value = ExtraerDocumentoMaterial(msgSap)

        RegistrarLog ws, i
    Next n

    MsgBox "Proceso terminado." & vbCrLf & _
           "Materiales capturados en MIGO: " & itemIndex & vbCrLf & _
           "Filas verificadas: " & procesadas & vbCrLf & _
           "Filas omitidas: " & omitidas & vbCrLf & _
           "Errores: " & errores, vbInformation
    Exit Sub

ErrGeneral:
    MsgBox "Error general: " & Err.Description, vbCritical
End Sub
Private Sub CapturarFilaMigoA07(ByVal session As Object, ByVal ws As Worksheet, ByVal rowIndex As Long)
    Dim material As String
    Dim cantidad As String
    Dim almacen As String
    Dim ceco As String
    Dim texto As String
    Dim destinatario As String
    Dim referencia As String

    material = CleanSapText(ws.Cells(rowIndex, COL_MATERIAL).Value, 40)
    cantidad = Replace(Trim$(CStr(ws.Cells(rowIndex, COL_CANTIDAD).Value)), ",", ".")
    almacen = FALLBACK_ALMACEN
    ceco = UCase$(CleanSapText(ws.Cells(rowIndex, COL_CECO).Value, 20))
    destinatario = CleanSapText(ws.Cells(rowIndex, COL_DESTINATARIO).Value, 50)
    texto = CleanSapText(ws.Cells(rowIndex, COL_TEXTO).Value, 50)
    referencia = CleanSapText(ws.Cells(rowIndex, COL_REFERENCIA).Value, 16)

    session.findById(ID_MAIN_WINDOW).maximize
    session.findById(ID_OK_CODE).Text = "/nMIGO"
    session.findById(ID_MAIN_WINDOW).sendVKey 0
    EsperarSAP session

    ' IMPORTANTE:
    ' La pantalla MIGO debe estar en A07 Salida de mercancias / R10 Otros / Movimiento 201.
    ' Si tu SAP no queda en esa pantalla al entrar a /nMIGO, graba esos combos con Script Recording.

    SetSapFieldByCandidates session, material, "Material", _
        ID_ITEM_TABLE & "ctxtGOITEM-MATNR[1,0]", _
        ID_ITEM_TABLE & "ctxtGOITEM-MAKTX[1,0]"

    SetSapFieldByCandidates session, cantidad, "Cantidad", _
        ID_ITEM_TABLE & "txtGOITEM-ERFMG[3,0]"

    SetSapFieldByCandidates session, almacen, "Almacen", _
        ID_ITEM_TABLE & "ctxtGOITEM-LGORT[5,0]", _
        ID_ITEM_TABLE & "ctxtGOITEM-LGOBE[5,0]"

    SetSapFieldByCandidates session, ceco, "Centro de costo", _
        ID_ITEM_TABLE & "ctxtGOITEM-KOSTL[6,0]"

    SetSapFieldByCandidates session, texto, "Texto", _
        ID_ITEM_TABLE & "txtGOITEM-SGTXT[32,0]"

    SetSapFieldByCandidates session, destinatario, "Destinatario", _
        ID_ITEM_TABLE & "txtGOITEM-WEMPF[34,0]"

    session.findById(ID_MAIN_WINDOW).sendVKey 0
    EsperarSAP session

    PressVerificarMigo session
    EsperarSAP session
    CerrarPopupSiExiste session
End Sub

Private Sub SetSapFieldByCandidates(ByVal session As Object, ByVal valor As String, ByVal nombreCampo As String, ParamArray ids() As Variant)
    Dim i As Long
    Dim id As String

    On Error Resume Next
    For i = LBound(ids) To UBound(ids)
        id = CStr(ids(i))
        Err.Clear
        session.findById(id).Text = valor
        If Err.Number = 0 Then
            session.findById(id).SetFocus
            On Error GoTo 0
            Exit Sub
        End If
    Next i
    On Error GoTo 0

    Err.Raise vbObjectError + 2100, , "No encontre el campo SAP para: " & nombreCampo & _
        ". Revisa el ID con SAP GUI Script Recording."
End Sub

Private Sub PressVerificarMigo(ByVal session As Object)
    Dim ids As Variant
    Dim i As Long
    Dim msgAntes As String

    ' Por seguridad NO se prueban botones alternos que pudieran contabilizar.
    ' Primero se ejecuta el comando global establecido por SAP luego el codigo de funcion CHECK.
    ids = Array("wnd[0]/tbar[1]/btn[7]")

    On Error Resume Next
    For i = LBound(ids) To UBound(ids)
        Err.Clear
        msgAntes = GetSapMessage(session)
        session.findById(CStr(ids(i))).Press
        If Err.Number = 0 Then
            EsperarSAP session
            On Error GoTo 0
            Exit Sub
        End If
    Next i

    Err.Clear
    session.findById(ID_OK_CODE).Text = "=CHECK"
    session.findById(ID_MAIN_WINDOW).sendVKey 0
    EsperarSAP session
    If Err.Number = 0 Then
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0

    Err.Raise vbObjectError + 2200, , "No pude presionar Verificar en MIGO. Graba el boton con SAP GUI Script Recording."
End Sub

Private Function ValidarLayout(ByVal ws As Worksheet, ByVal lastRow As Long, ByVal showMessage As Boolean) As Boolean
    Dim i As Long
    Dim errores As String
    Dim estado As String

    If lastRow < 2 Then
        If showMessage Then MsgBox "No hay datos en " & SHEET_BAJAS & ".", vbExclamation
        ValidarLayout = False
        Exit Function
    End If

    For i = 2 To lastRow
        If Trim$(CStr(ws.Cells(i, COL_FOLIO).Value)) <> "" Then
            estado = UCase$(Trim$(CStr(ws.Cells(i, COL_STATUS).Value)))

            If estado <> "VERIFICADO" And estado <> "OK" And estado <> "CONTABILIZADO" Then
                If Trim$(CStr(ws.Cells(i, COL_MATERIAL).Value)) = "" Then errores = errores & vbCrLf & "Fila " & i & ": falta MATERIAL."
                If Not IsNumeric(ws.Cells(i, COL_CANTIDAD).Value) Then
                    errores = errores & vbCrLf & "Fila " & i & ": CANTIDAD invalida."
                ElseIf CDbl(ws.Cells(i, COL_CANTIDAD).Value) <= 0 Then
                    errores = errores & vbCrLf & "Fila " & i & ": CANTIDAD invalida."
                End If
                If Trim$(CStr(ws.Cells(i, COL_CECO).Value)) = "" Then errores = errores & vbCrLf & "Fila " & i & ": falta CECO."
                If Trim$(CStr(ws.Cells(i, COL_DESTINATARIO).Value)) = "" Then errores = errores & vbCrLf & "Fila " & i & ": falta DESTINATARIO."
                If Trim$(CStr(ws.Cells(i, COL_TEXTO).Value)) = "" Then errores = errores & vbCrLf & "Fila " & i & ": falta TEXTO."
            End If
        End If
    Next i

    If Len(errores) > 0 Then
        If showMessage Then MsgBox "Corrige antes de procesar:" & vbCrLf & errores, vbCritical
        ValidarLayout = False
    Else
        ValidarLayout = True
    End If
End Function

Private Function GetSapSession() As Object
    Dim SapGuiAuto As Object
    Dim SapApp As Object
    Dim Connections As Object
    Dim Sessions As Object
    Dim Connection As Object
    Dim s As Object
    Dim candidato As Object
    Dim c As Long
    Dim n As Long
    Dim trx As String
    Dim activa As Boolean
    Dim detalle As String

    On Error GoTo ErrHandler

    Set SapGuiAuto = GetObject("SAPGUI")
    Set SapApp = SapGuiAuto.GetScriptingEngine
    Set Connections = SapApp.Connections

    If Connections.Count = 0 Then
        Err.Raise vbObjectError + 1001, , "SAP GUI esta abierto, pero no hay conexiones SAP iniciadas."
    End If

    ' IMPORTANTE: usar ElementAt(Long), no el indexador implicito Children(variable).
    ' Algunas versiones de SAP GUI devuelven Gui_Err_Bad_Index_Type (618)
    ' cuando el acceso a una coleccion se resuelve por Item(Variant).
    For c = 0 To Connections.Count - 1
        Set Connection = Connections.ElementAt(c)

        On Error Resume Next
        Set Sessions = Connection.Sessions
        If Err.Number <> 0 Then
            detalle = Err.Description
            Err.Clear
            Set Sessions = Nothing
        End If
        On Error GoTo ErrHandler

        If Not Sessions Is Nothing Then
            For n = 0 To Sessions.Count - 1
                Set s = Sessions.ElementAt(n)

                If Not s Is Nothing Then
                    If Not s.Busy Then
                        trx = ""
                        activa = False

                        On Error Resume Next
                        trx = UCase$(CStr(s.Info.Transaction))
                        activa = CBool(s.IsActive)
                        Err.Clear
                        On Error GoTo ErrHandler

                        ' Prioridad 1: MIGO activa.
                        If trx = "MIGO" And activa Then
                            Set GetSapSession = s
                            Exit Function
                        End If

                        ' Prioridad 2: cualquier sesion activa.
                        If activa Then
                            Set GetSapSession = s
                            Exit Function
                        End If

                        ' Respaldo: primera sesion disponible.
                        If candidato Is Nothing Then Set candidato = s
                    End If
                End If
            Next n
        End If

        Set Sessions = Nothing
    Next c

    If Not candidato Is Nothing Then
        Set GetSapSession = candidato
        Exit Function
    End If

    If Len(detalle) > 0 Then
        Err.Raise vbObjectError + 1002, , "No encontre una sesion SAP util. Detalle: " & detalle
    Else
        Err.Raise vbObjectError + 1002, , "No encontre una sesion SAP disponible."
    End If

ErrHandler:
    detalle = Err.Description
    Err.Raise vbObjectError + 1000, , _
        "No se pudo conectar a SAP GUI. Verifica SAP abierto, sesion iniciada y SAP GUI Scripting habilitado. Detalle: " & detalle
End Function

Private Sub EsperarSAP(ByVal session As Object)
    Do While session.Busy
        DoEvents
    Loop
End Sub

Private Function GetSapMessage(ByVal session As Object) As String
    On Error Resume Next
    GetSapMessage = session.findById(ID_STATUS_BAR).Text
    If Len(GetSapMessage) = 0 Then GetSapMessage = "Sin mensaje en barra SAP."
    On Error GoTo 0
End Function

Private Function SafeSapMessage(ByVal session As Object) As String
    On Error Resume Next
    SafeSapMessage = GetSapMessage(session)
    If Len(SafeSapMessage) = 0 Then SafeSapMessage = "Sin mensaje SAP."
    On Error GoTo 0
End Function

Private Function GetSapMessageType(ByVal session As Object) As String
    On Error Resume Next
    GetSapMessageType = UCase$(session.findById(ID_STATUS_BAR).MessageType)
    On Error GoTo 0
End Function
Private Function ObtenerFilaVisibleMigo(ByVal session As Object, ByVal itemIndex As Long) As Long
    Dim tbl As Object
    Dim visibleRows As Long
    Dim topRow As Long

    Set tbl = GetMigoItemTable(session)

    On Error Resume Next
    visibleRows = tbl.VisibleRowCount
    On Error GoTo 0

    If visibleRows <= 0 Then visibleRows = 10

    If itemIndex <= visibleRows Then
        On Error Resume Next
        tbl.VerticalScrollbar.Position = 0
        On Error GoTo 0

        ObtenerFilaVisibleMigo = itemIndex - 1
    Else
        topRow = itemIndex - visibleRows

        On Error Resume Next
        tbl.VerticalScrollbar.Position = topRow
        On Error GoTo 0

        EsperarSAP session

        ObtenerFilaVisibleMigo = visibleRows - 1
    End If
End Function
Private Sub PrepararMigoA07(ByVal session As Object)
    session.findById(ID_MAIN_WINDOW).maximize
    session.findById(ID_OK_CODE).Text = "/nMIGO"
    session.findById(ID_MAIN_WINDOW).sendVKey 0
    EsperarSAP session
    CerrarPopupSiExiste session

    ' IMPORTANTE:
    ' La pantalla debe quedar en A07 Salida de mercancías / R10 Otros / Movimiento 201.
    ' Si no queda así automáticamente, hay que grabar esos combos con SAP GUI Script Recording.
End Sub
Private Sub CapturarFilaMigoA07_Lote(ByVal session As Object, ByVal ws As Worksheet, ByVal rowIndex As Long, ByVal itemIndex As Long)
    Dim material As String
    Dim cantidad As String
    Dim almacen As String
    Dim ceco As String
    Dim texto As String
    Dim destinatario As String
    Dim sapRow As Long
    Dim tbl As Object

    material = CleanSapText(ws.Cells(rowIndex, COL_MATERIAL).Value, 40)
    cantidad = Replace(Trim$(CStr(ws.Cells(rowIndex, COL_CANTIDAD).Value)), ",", ".")
    almacen = FALLBACK_ALMACEN
    ceco = UCase$(CleanSapText(ws.Cells(rowIndex, COL_CECO).Value, 20))
    destinatario = CleanSapText(ws.Cells(rowIndex, COL_DESTINATARIO).Value, 50)
    texto = CleanSapText(ws.Cells(rowIndex, COL_TEXTO).Value, 50)

    sapRow = ObtenerFilaVisibleMigo(session, itemIndex)
    Set tbl = GetMigoItemTable(session)

    ' IMPORTANTE: ya no usamos posiciones fijas como [32,fila] o [34,fila].
    ' Se busca cada columna por su nombre tecnico SAP en el layout actual.
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-MATNR", material, "Material linea " & itemIndex
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-ERFMG", cantidad, "Cantidad linea " & itemIndex
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-LGORT", almacen, "Almacen linea " & itemIndex
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-KOSTL", ceco, "Centro de costo linea " & itemIndex
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-SGTXT", texto, "Texto linea " & itemIndex
    SetMigoFieldByTechnicalName tbl, sapRow, "GOITEM-WEMPF", destinatario, "Destinatario linea " & itemIndex

    ' Enter para que SAP acepte la linea y prepare la siguiente.
    session.findById(ID_MAIN_WINDOW).sendVKey 0
    EsperarSAP session
    CerrarPopupSiExiste session
End Sub

' ============================================================
' CAPA ROBUSTA MIGO - MULTIUSUARIO / MULTILAYOUT
' ============================================================

Private Function GetMigoItemTable(ByVal session As Object) As Object
    Dim tbl As Object
    Dim usr As Object

    ' Primero intentar el ID conocido de esta version de MIGO.
    On Error Resume Next
    Set tbl = session.findById(ID_ITEM_TABLE_OBJ)
    On Error GoTo 0

    If Not tbl Is Nothing Then
        Set GetMigoItemTable = tbl
        Exit Function
    End If

    ' Respaldo: buscar recursivamente el GuiTableControl de GOITEM.
    On Error Resume Next
    Set usr = session.findById("wnd[0]/usr")
    On Error GoTo 0

    If Not usr Is Nothing Then
        Set tbl = FindTableByToken(usr, "GOITEM")
    End If

    If tbl Is Nothing Then
        Err.Raise vbObjectError + 2300, , _
            "No encontre la tabla de posiciones de MIGO (GOITEM). " & _
            "Verifica que estes dentro de MIGO y que la lista de posiciones este visible."
    End If

    Set GetMigoItemTable = tbl
End Function

Private Function FindTableByToken(ByVal parentObj As Object, ByVal token As String) As Object
    Dim i As Long
    Dim child As Object
    Dim found As Object
    Dim tipo As String
    Dim nombre As String
    Dim tableField As String
    Dim esContainer As Boolean

    On Error GoTo Salir

    For i = 0 To parentObj.Children.Count - 1
        Set child = parentObj.Children.ElementAt(i)

        tipo = ""
        nombre = ""
        tableField = ""
        esContainer = False

        On Error Resume Next
        tipo = UCase$(CStr(child.Type))
        nombre = UCase$(CStr(child.Name))
        tableField = UCase$(CStr(child.TableFieldName))
        esContainer = CBool(child.ContainerType)
        On Error GoTo Salir

        If tipo = "GUITABLECONTROL" Then
            If InStr(1, nombre, UCase$(token), vbTextCompare) > 0 Or _
               InStr(1, tableField, UCase$(token), vbTextCompare) > 0 Then
                Set FindTableByToken = child
                Exit Function
            End If
        End If

        If esContainer Then
            Set found = FindTableByToken(child, token)
            If Not found Is Nothing Then
                Set FindTableByToken = found
                Exit Function
            End If
        End If
    Next i

Salir:
End Function

Private Function FindMigoColumn(ByVal tbl As Object, ByVal technicalName As String) As Long
    Dim c As Long
    Dim r As Long
    Dim maxRows As Long
    Dim cell As Object
    Dim nm As String
    Dim wanted As String

    FindMigoColumn = -1
    wanted = UCase$(technicalName)

    On Error Resume Next
    maxRows = tbl.VisibleRowCount
    On Error GoTo 0
    If maxRows <= 0 Then maxRows = 1
    If maxRows > 3 Then maxRows = 3

    For c = 0 To tbl.Columns.Count - 1
        For r = 0 To maxRows - 1
            Set cell = Nothing
            nm = ""

            On Error Resume Next
            Set cell = tbl.GetCell(r, c)
            If Not cell Is Nothing Then nm = UCase$(CStr(cell.Name))
            On Error GoTo 0

            If Len(nm) > 0 Then
                If nm = wanted Or Right$(nm, Len(wanted)) = wanted Then
                    FindMigoColumn = c
                    Exit Function
                End If
            End If
        Next r
    Next c
End Function

Private Sub SetMigoFieldByTechnicalName(ByVal tbl As Object, ByVal sapRow As Long, _
                                        ByVal technicalName As String, ByVal valor As String, _
                                        ByVal nombreCampo As String)
    Dim col As Long
    Dim cell As Object

    col = FindMigoColumn(tbl, technicalName)

    If col < 0 Then
        Err.Raise vbObjectError + 2310, , _
            "No encontre '" & nombreCampo & "' (" & technicalName & ") en el layout MIGO del usuario. " & _
            "Ejecuta Normalizar_Layout_MIGO o agrega esa columna desde Configurar layout."
    End If

    On Error GoTo ErrCampo
    Set cell = tbl.GetCell(sapRow, col)
    cell.Text = valor
    cell.SetFocus
    Exit Sub

ErrCampo:
    Err.Raise vbObjectError + 2311, , _
        "No pude escribir '" & nombreCampo & "' en MIGO. Columna tecnica: " & technicalName & _
        ". Detalle: " & Err.Description
End Sub

Private Function BuildMigoPermutation(ByVal tbl As Object, ByVal desiredFields As Variant, _
                                      ByRef missing As String) As String
    Dim used() As Boolean
    Dim i As Long
    Dim c As Long
    Dim col As Long
    Dim result As String

    ReDim used(0 To tbl.Columns.Count - 1)

    For i = LBound(desiredFields) To UBound(desiredFields)
        col = FindMigoColumn(tbl, CStr(desiredFields(i)))
        If col >= 0 Then
            If Not used(col) Then
                If Len(result) > 0 Then result = result & " "
                result = result & CStr(col)
                used(col) = True
            End If
        Else
            If Len(missing) > 0 Then missing = missing & vbCrLf
            missing = missing & "- " & CStr(desiredFields(i))
        End If
    Next i

    ' Mantener el resto de las columnas sin perderlas.
    For c = 0 To tbl.Columns.Count - 1
        If Not used(c) Then
            If Len(result) > 0 Then result = result & " "
            result = result & CStr(c)
        End If
    Next c

    BuildMigoPermutation = result
End Function

Public Sub Normalizar_Layout_MIGO()
    Dim session As Object
    Dim tbl As Object
    Dim desired As Variant
    Dim permutation As String
    Dim missing As String

    On Error GoTo ErrHandler

    Set session = GetSapSession()

    If UCase$(CStr(session.Info.Transaction)) <> "MIGO" Then
        PrepararMigoA07 session
    End If

    Set tbl = GetMigoItemTable(session)

    desired = Array( _
        "GOITEM-MATNR", _
        "GOITEM-ERFMG", _
        "GOITEM-LGORT", _
        "GOITEM-KOSTL", _
        "GOITEM-SGTXT", _
        "GOITEM-WEMPF")

    permutation = BuildMigoPermutation(tbl, desired, missing)

    If Len(missing) > 0 Then
        ' Si faltan columnas, SAP expone oficialmente ConfigureLayout para abrir el dialogo.
        On Error Resume Next
        tbl.ConfigureLayout
        On Error GoTo 0

        MsgBox "El layout no contiene todas las columnas requeridas:" & vbCrLf & vbCrLf & _
               missing & vbCrLf & vbCrLf & _
               "Se intento abrir Configurar layout. Agrega las columnas faltantes y vuelve a ejecutar esta macro.", _
               vbExclamation, "MIGO - Layout incompleto"
        Exit Sub
    End If

    tbl.ReorderTable permutation

    MsgBox "Layout MIGO normalizado para esta sesion." & vbCrLf & _
           "Orden principal: Material, Cantidad, Almacen, CECO, Texto y Destinatario.", _
           vbInformation, "MIGO - Layout OK"
    Exit Sub

ErrHandler:
    MsgBox "No se pudo normalizar MIGO: " & Err.Description, vbCritical
End Sub

Public Sub Probar_Conexion_SAP()
    Dim session As Object
    Dim msg As String

    On Error GoTo ErrHandler

    Set session = GetSapSession()

    msg = "Conexion SAP GUI OK." & vbCrLf & vbCrLf & _
          "Sistema: " & CStr(session.Info.SystemName) & vbCrLf & _
          "Usuario: " & CStr(session.Info.User) & vbCrLf & _
          "Transaccion: " & CStr(session.Info.Transaction) & vbCrLf & _
          "Sesion: " & CStr(session.Info.SessionNumber)

    MsgBox msg, vbInformation, "SAP GUI Scripting - Conexion OK"
    Exit Sub

ErrHandler:
    MsgBox "Prueba de conexion SAP fallo: " & Err.Description, vbCritical, "SAP GUI Scripting"
End Sub

Public Sub Diagnosticar_Layout_MIGO()
    Dim session As Object
    Dim tbl As Object
    Dim fields As Variant
    Dim labels As Variant
    Dim i As Long
    Dim col As Long
    Dim reporte As String
    Dim readOnlyMode As String

    On Error GoTo ErrHandler

    Set session = GetSapSession()

    reporte = "Sistema: " & CStr(session.Info.SystemName) & vbCrLf & _
              "Usuario SAP: " & CStr(session.Info.User) & vbCrLf & _
              "Transaccion: " & CStr(session.Info.Transaction) & vbCrLf

    On Error Resume Next
    readOnlyMode = CStr(session.Info.ScriptingModeReadOnly)
    On Error GoTo ErrHandler
    If Len(readOnlyMode) > 0 Then reporte = reporte & "Scripting solo lectura: " & readOnlyMode & vbCrLf

    If UCase$(CStr(session.Info.Transaction)) <> "MIGO" Then
        reporte = reporte & vbCrLf & "Abre MIGO y vuelve a ejecutar el diagnostico."
        MsgBox reporte, vbExclamation, "Diagnostico MIGO"
        Exit Sub
    End If

    Set tbl = GetMigoItemTable(session)

    fields = Array("GOITEM-MATNR", "GOITEM-ERFMG", "GOITEM-LGORT", "GOITEM-KOSTL", "GOITEM-SGTXT", "GOITEM-WEMPF")
    labels = Array("Material", "Cantidad", "Almacen", "CECO", "Texto", "Destinatario")

    reporte = reporte & "Columnas detectadas: " & tbl.Columns.Count & vbCrLf & vbCrLf

    For i = LBound(fields) To UBound(fields)
        col = FindMigoColumn(tbl, CStr(fields(i)))
        If col >= 0 Then
            reporte = reporte & "OK  " & labels(i) & " -> columna " & (col + 1) & " (" & fields(i) & ")" & vbCrLf
        Else
            reporte = reporte & "FALTA  " & labels(i) & " (" & fields(i) & ")" & vbCrLf
        End If
    Next i

    MsgBox reporte, vbInformation, "Diagnostico MIGO"
    Exit Sub

ErrHandler:
    MsgBox "Diagnostico MIGO fallo: " & Err.Description, vbCritical
End Sub

Private Sub CerrarPopupSiExiste(ByVal session As Object)
    On Error Resume Next
    If session.Children.Count > 1 Then
        session.findById("wnd[1]").Close
    End If
    On Error GoTo 0
End Sub

Private Sub RegistrarLog(ByVal ws As Worksheet, ByVal rowIndex As Long)
    Dim logWs As Worksheet
    Dim nextRow As Long

    On Error Resume Next
    Set logWs = ThisWorkbook.Worksheets(SHEET_LOG)
    On Error GoTo 0

    If logWs Is Nothing Then Exit Sub

    nextRow = logWs.Cells(logWs.Rows.Count, "A").End(xlUp).Row + 1

    logWs.Cells(nextRow, "A").Value = Now
    logWs.Cells(nextRow, "B").Value = rowIndex
    logWs.Cells(nextRow, "C").Value = ws.Cells(rowIndex, COL_FOLIO).Value
    logWs.Cells(nextRow, "D").Value = ws.Cells(rowIndex, COL_MATERIAL).Value
    logWs.Cells(nextRow, "E").Value = ws.Cells(rowIndex, COL_CANTIDAD).Value
    logWs.Cells(nextRow, "F").Value = ws.Cells(rowIndex, COL_CECO).Value
    logWs.Cells(nextRow, "G").Value = ws.Cells(rowIndex, COL_DESTINATARIO).Value
    logWs.Cells(nextRow, "H").Value = ws.Cells(rowIndex, COL_TEXTO).Value
    logWs.Cells(nextRow, "I").Value = ws.Cells(rowIndex, COL_STATUS).Value
    logWs.Cells(nextRow, "J").Value = ws.Cells(rowIndex, COL_MENSAJE).Value
    logWs.Cells(nextRow, "K").Value = ws.Cells(rowIndex, COL_DOC_MAT).Value
End Sub

Private Function CleanSapText(ByVal Value As Variant, ByVal maxLen As Long) As String
    Dim s As String

    s = Trim$(CStr(Value))
    s = Replace(s, vbCr, " ")
    s = Replace(s, vbLf, " ")

    Do While InStr(s, "  ") > 0
        s = Replace(s, "  ", " ")
    Loop

    If maxLen > 0 And Len(s) > maxLen Then s = Left$(s, maxLen)
    CleanSapText = s
End Function

Private Function ExtraerDocumentoMaterial(ByVal mensaje As String) As String
    Dim i As Long
    Dim ch As String
    Dim buffer As String
    Dim best As String

    For i = 1 To Len(mensaje)
        ch = Mid$(mensaje, i, 1)
        If ch Like "#" Then
            buffer = buffer & ch
        Else
            If Len(buffer) > Len(best) Then best = buffer
            buffer = ""
        End If
    Next i

    If Len(buffer) > Len(best) Then best = buffer

    If Len(best) >= 8 Then
        ExtraerDocumentoMaterial = best
    Else
        ExtraerDocumentoMaterial = ""
    End If
End Function
