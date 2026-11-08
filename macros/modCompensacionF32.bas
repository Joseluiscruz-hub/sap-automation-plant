Attribute VB_Name = "modCompensacionF32"
' Macro mejorada para compensación masiva F-32 (RH / RT / X1)
' Usada en el saneamiento FI post-S/4HANA · 410k documentos

Sub Compensarf32_VariosCte()
    Dim wsNotas As Worksheet
    Dim rngCtes As Range, rngDoctos As Range
    Dim totalCtes As Long, totalDoctos As Long
    Dim filasCtesF32 As Long, filasDoctosF32 As Long
    Dim clienteInicial As String
    Dim respuesta As VbMsgBoxResult
    Dim SapGuiAuto As Object
    Dim Appl As Object
    Dim Connection As Object
    Dim session As Object
    Dim cuadre As Double
    Dim i As Long, j As Long
    Dim regPendientes As Long
    
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set wsNotas = ThisWorkbook.Sheets("Notas")
    
    respuesta = MsgBox("¿Desea compensar documentos?", vbOKCancel + vbQuestion, "H.V.H")
    If respuesta <> vbOK Then GoTo CleanUp
    
    If wsNotas.Range("A13").Value = "" Then
        MsgBox "No hay clientes desde A13.", vbExclamation
        GoTo CleanUp
    End If
    
    Set rngCtes = wsNotas.Range("A13", wsNotas.Range("A13").End(xlDown))
    Set rngDoctos = wsNotas.Range("B13", wsNotas.Range("B13").End(xlDown))
    
    totalCtes = rngCtes.Rows.Count
    totalDoctos = rngDoctos.Rows.Count
    
    filasCtesF32 = Val(wsNotas.Range("A11").Value)
    filasDoctosF32 = Val(wsNotas.Range("B11").Value)
    If filasCtesF32 <= 0 Then filasCtesF32 = 8
    If filasDoctosF32 <= 0 Then filasDoctosF32 = 7
    
    clienteInicial = CStr(wsNotas.Range("A13").Value)
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set Appl = SapGuiAuto.GetScriptingEngine
    Set Connection = Appl.Children(0)
    Set session = Connection.Children(0)
    
    With session
        .findById("wnd[0]").Maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "f-32"
        .findById("wnd[0]").sendVKey 0
        
        .findById("wnd[0]/usr/ctxtRF05A-AGKON").Text = clienteInicial
        .findById("wnd[0]/usr/txtBKPF-MONAT").Text = Format(Date, "mm")
        .findById("wnd[0]/usr/ctxtBKPF-BUKRS").Text = "F142"
        .findById("wnd[0]/usr/ctxtBKPF-WAERS").Text = "MXN"
        .findById("wnd[0]").sendVKey 0
        
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z-").press
        .findById("wnd[0]/tbar[1]/btn[7]").press
        .findById("wnd[0]/tbar[1]/btn[6]").press
        .findById("wnd[0]/usr/chkRF05A-XMULK").Selected = True
        .findById("wnd[0]/usr/sub:SAPMF05A:0710/radRF05A-XPOS1[0,0]").SetFocus
        .findById("wnd[0]").sendVKey 0
        
        ' Clientes en lotes
        regPendientes = totalCtes
        i = 1
        Do While i <= totalCtes
            If regPendientes > filasCtesF32 Then
                For j = 0 To filasCtesF32 - 1
                    If i + j <= totalCtes Then
                        .findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[" & j & ",0]").Text = rngCtes.Cells(i + j, 1).Value
                        .findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[" & j & ",29]").Selected = True
                    End If
                Next j
                .findById("wnd[1]/tbar[0]/btn[0]").press
                regPendientes = regPendientes - filasCtesF32
                i = i + filasCtesF32
            Else
                For j = 0 To regPendientes - 1
                    .findById("wnd[1]/usr/sub:SAPMF05A:0609/ctxtRF05A-AGKON[" & j & ",0]").Text = rngCtes.Cells(i + j, 1).Value
                    .findById("wnd[1]/usr/sub:SAPMF05A:0609/chkRF05A-XNOPS[" & j & ",29]").Selected = True
                Next j
                .findById("wnd[1]/tbar[0]/btn[0]").press
                Exit Do
            End If
        Loop
        
        ' Documentos
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z-").press
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SEARCH").press
        .findById("wnd[1]/usr/sub:SAPDF05X:2000/radRF05A-XPOS1[3,0]").Select
        .findById("wnd[1]/tbar[0]/btn[0]").press
        
        regPendientes = totalDoctos
        i = 1
        Do While i <= totalDoctos
            If regPendientes > filasDoctosF32 Then
                For j = 0 To filasDoctosF32 - 1
                    If i + j <= totalDoctos Then
                        .findById("wnd[2]/usr/sub:SAPDF05X:0731/txtRF05A-SEL01[" & j & ",0]").Text = rngDoctos.Cells(i + j, 1).Value
                    End If
                Next j
                .findById("wnd[2]/tbar[0]/btn[13]").press
                regPendientes = regPendientes - filasDoctosF32
                i = i + filasDoctosF32
            Else
                For j = 0 To regPendientes - 1
                    .findById("wnd[2]/usr/sub:SAPDF05X:0731/txtRF05A-SEL01[" & j & ",0]").Text = rngDoctos.Cells(i + j, 1).Value
                Next j
                Exit Do
            End If
        Loop
        
        .findById("wnd[2]/tbar[0]/btn[0]").press
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnICON_SELECT_ALL").press
        .findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/btnIC_Z+").press
        
        On Error Resume Next
        cuadre = Val(.findById("wnd[0]/usr/tabsTS/tabpMAIN/ssubPAGE:SAPDF05X:6102/txtRF05A-DIFFB").Text)
        On Error GoTo ErrorHandler
        
        If cuadre = 0 Then
            .findById("wnd[0]/tbar[0]/btn[11]").press
            wsNotas.Range("P2").Value = .findById("wnd[0]/sbar").Text
            .findById("wnd[0]/tbar[0]/btn[15]").press
            MsgBox "Documentos compensados.", vbInformation
        Else
            .findById("wnd[0]/tbar[0]/btn[15]").press
            On Error Resume Next
            .findById("wnd[1]/usr/btnSPOP-OPTION1").press
            On Error GoTo ErrorHandler
            MsgBox "Existen diferencias. Favor de validar.", vbExclamation
        End If
    End With
    
    GoTo CleanUp

ErrorHandler:
    MsgBox "Error en compensación: " & Err.Description, vbCritical
CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
End Sub
