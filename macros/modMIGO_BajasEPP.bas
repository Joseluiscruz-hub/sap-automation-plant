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
Private Const FALLBACK_ALMACEN As String = "3100"

Public Sub Validar_Bajas_EPP()
    MsgBox "Modulo modMIGO_BajasEPP - ver codigo completo en el repositorio local."
End Sub

Public Sub Procesar_Bajas_EPP_MIGO()
    MsgBox "Modulo modMIGO_BajasEPP - ver codigo completo en el repositorio local."
End Sub

Public Sub Normalizar_Layout_MIGO()
    MsgBox "Modulo modMIGO_BajasEPP - ver codigo completo en el repositorio local."
End Sub
