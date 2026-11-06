Attribute VB_Name = "modInstaladorConciliacion"
Option Explicit

Public Sub Instalar_Conciliacion_Limpia()
    On Error GoTo ErrorHandler

    Dim rutaBas As Variant
    Dim vbProj As Object
    Dim comp As Object
    Dim nombresEliminar As Variant
    Dim nombre As Variant
    Dim i As Long

    rutaBas = ThisWorkbook.Path & Application.PathSeparator & "Modulo6_Limpio_Integrado.bas"
    If Dir(CStr(rutaBas)) = vbNullString Then
        rutaBas = Application.GetOpenFilename("Modulos VBA (*.bas),*.bas", , "Selecciona Modulo6_Limpio_Integrado.bas")
        If VarType(rutaBas) = vbBoolean Then Exit Sub
    End If

    Set vbProj = ThisWorkbook.VBProject

    nombresEliminar = Array("Módulo1", "Módulo2", "Módulo3", "Módulo4", "Módulo5", "Módulo6", _
                            "Modulo1", "Modulo2", "Modulo3", "Modulo4", "Modulo5", "Modulo6")

    Application.ScreenUpdating = False
    Application.DisplayAlerts = False

    For i = LBound(nombresEliminar) To UBound(nombresEliminar)
        nombre = nombresEliminar(i)
        On Error Resume Next
        Set comp = vbProj.VBComponents(CStr(nombre))
        If Not comp Is Nothing Then vbProj.VBComponents.Remove comp
        Set comp = Nothing
        On Error GoTo ErrorHandler
    Next i

    vbProj.VBComponents.Import CStr(rutaBas)
    ReasignarBotonesConciliacion

    Application.DisplayAlerts = True
    Application.ScreenUpdating = True

    MsgBox "Instalacion completa." & vbCrLf & _
           "Se conservaron ThisWorkbook y las hojas; se reemplazaron los modulos viejos por Módulo6 limpio.", vbInformation
    Exit Sub

ErrorHandler:
    Application.DisplayAlerts = True
    Application.ScreenUpdating = True
    MsgBox "No se pudo instalar automaticamente." & vbCrLf & vbCrLf & _
           "Detalle: " & Err.Number & " - " & Err.Description & vbCrLf & vbCrLf & _
           "Activa: Archivo > Opciones > Centro de confianza > Configuracion de macros > " & _
           "Confiar en el acceso al modelo de objetos de proyectos de VBA.", vbCritical
End Sub

Private Sub ReasignarBotonesConciliacion()
    On Error Resume Next

    Dim ws As Worksheet
    Dim shp As Shape
    Dim accion As String

    For Each ws In ThisWorkbook.Worksheets
        For Each shp In ws.Shapes
            accion = shp.OnAction
            If Len(accion) > 0 Then
                If InStr(1, accion, "RestablecerDatos", vbTextCompare) > 0 Then
                    shp.OnAction = "'" & ThisWorkbook.Name & "'!Módulo6.RestablecerDatos"
                ElseIf InStr(1, accion, "ConciliarTransferenciasOptimizado", vbTextCompare) > 0 Then
                    shp.OnAction = "'" & ThisWorkbook.Name & "'!ConciliarTransferenciasOptimizado"
                End If
            End If
        Next shp
    Next ws
End Sub
