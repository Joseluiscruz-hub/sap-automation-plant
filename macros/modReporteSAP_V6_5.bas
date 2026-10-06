Option Explicit

' VERSION 6.5 - REP2 + PRODUCCION + PDP PRODUCTION PLAN + PAROS / SAP GUI
' =====================================================================
' ACTUALIZACION INTEGRAL - REPORTE DE EFICIENCIA CUAUTITLAN
'
' Objetivo:
'   - Lee la fecha de REPORTE EFICIENCIA!F3.
'   - Ejecuta ZMNPPREP0002 P_ACU una vez SOLO para el dia de corte.
'   - Ejecuta ZMNPPREP0002 P_ACU otra vez desde el dia 1 hasta el corte para el MES.
'   - Ejecuta ZMNPPREP0002 P_DET para Fecha de inicio Real y % Eficiencia Acumulada.
'   - Actualiza REP2 sin perder sus formulas auxiliares y guarda el detalle en REP2_DETALLE.
'   - Ejecuta ZMNPPREP0004 con variante /SMOSKE_ALT solo para el dia de corte.
'   - Agrupa las causas exactas reportadas por SAP, toma TOP 3 y escribe W6:W9.
'   - Genera PAROS_AGRUPADOS incluyendo PRODUCCION para validar 1440 min por linea.
'   - Ejecuta ZMNPPREP0006 con variante /BRBR y actualiza la hoja Ago 26.
'   - Carga Cajas Unitarias por dia/linea en BE:BH, actualiza MTD y dia de corte.
'   - Busca el Production Plan valido mas reciente en PDP\ENTRADA.
'   - Lee Production Plan_1, convierte CF / Factor CF_CU y actualiza PLAN Ago 26 BJ:BN.
'   - Recalcula el libro para refrescar REPORTE EFICIENCIA y cumplimiento PDP.
'
' Modifica REP2, Ago 26 (REAL SAP + PLAN PDP), REP2_DETALLE, PAROS_AGRUPADOS y REPORTE EFICIENCIA!W6:W9.
' No modifica CTT ni Dia.
' =====================================================================

Private Const CENTRO As String = "MABC"
Private Const TRANSACCION As String = "ZMNPPREP0004"
Private Const VARIANTE As String = "/SMOSKE_ALT"
Private Const HOJA_REPORTE As String = "REPORTE EFICIENCIA"
Private Const CELDA_FECHA As String = "F3"
Private Const MAX_COMENTARIOS As Long = 3
Private Const TRANSACCION_REP2 As String = "ZMNPPREP0002"
Private Const VARIANTE_REP2 As String = "/DSB"
Private Const CENTRO_TRABAJO_REP2 As String = "Linea*"
Private Const HOJA_REP2 As String = "REP2"
Private Const HOJA_REP2_DETALLE As String = "REP2_DETALLE"
Private Const HOJA_PAROS_AGRUPADOS As String = "PAROS_AGRUPADOS"
Private Const TRANSACCION_PRODUCCION As String = "ZMNPPREP0006"
Private Const VARIANTE_PRODUCCION As String = "/BRBR"
Private Const HOJA_PDP As String = "Ago 26"
Private Const HOJA_PDP_FUENTE As String = "Production Plan_1"
Private Const HOJA_CATALOGO_PDP As String = "CATALOGO_PDP"
Private Const HOJA_PDP_CONVERSION As String = "PDP_CONVERSION"
Private Const CARPETA_PDP_RELATIVA As String = "PDP\ENTRADA"
Private Const CELDA_META_PDP_ARCHIVO As String = "Y37"
Private Const CELDA_META_PDP_FECHA As String = "Y38"
Private Const CELDA_META_PDP_PERIODO As String = "Y39"
Private Const CELDA_META_FECHA As String = "X36"
Private Const CELDA_META_EFI_ACUM As String = "Y36"

Public Sub V6_ActualizarReporteCompletoSAP()
    MsgBox "Modulo modReporteSAP_V6_5 - ver codigo completo en el repositorio local."
End Sub
