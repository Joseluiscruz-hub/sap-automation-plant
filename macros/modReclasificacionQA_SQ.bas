Attribute VB_Name = "modReclasificacionQA_SQ"
' Macro mejorada para reclasificación QA → SQ (cockpit por ruta)
' Basada en el caso de estudio FI post-S/4HANA · sociedad F142

Sub Reclasificar_NCA()
    Dim wsNotas As Worksheet
    Dim wsNombres As Worksheet
    Dim wsFBL5N As Worksheet
    Dim rutaReparto As String
    Dim respuesta As VbMsgBoxResult
    
    On Error GoTo ErrorHandler
    Application.DisplayAlerts = False
    Application.ScreenUpdating = False
    
    Set wsNotas = ThisWorkbook.Sheets("Notas")
    Set wsNombres = ThisWorkbook.Sheets("Nombre de los archivos")
    Set wsFBL5N = ThisWorkbook.Sheets("FBL5N")
    
    rutaReparto = Trim(CStr(wsNotas.Range("B4").Value))
    
    respuesta = MsgBox("¿Generar formatos de Reclasificación - QA " & rutaReparto & "?", _
                       vbOKCancel + vbQuestion, "GENERA FORMATOS")
    If respuesta <> vbOK Then GoTo CleanUp
    
    If Trim(CStr(wsNombres.Range("A3").Value)) <> "" Then
        MsgBox "Ya hay archivos generados. Limpia primero.", vbExclamation
        GoTo CleanUp
    End If
    
    If Val(wsFBL5N.Range("A2").Value) <= 1 Then
        MsgBox "No hay registros a generar.", vbInformation
        GoTo CleanUp
    End If
    
    Genera_CargoAbono
    Divide_Archivos
    
CleanUp:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    Exit Sub

ErrorHandler:
    MsgBox "Error en Reclasificar_NCA: " & Err.Description, vbCritical
    Resume CleanUp
End Sub

' Nota: Genera_CargoAbono y Divide_Archivos mantienen la lógica original de layout
' y división por lotes de 900. Se recomienda importar el módulo completo
' y ajustar nombres de tablas si es necesario.
