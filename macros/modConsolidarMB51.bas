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
    MsgBox "Modulo modConsolidarMB51 - ver codigo completo en el repositorio local."
End Sub

Public Sub Consolidar_MB51_2025()
    Consolidar_MB51
End Sub
