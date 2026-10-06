Attribute VB_Name = "modComentariosParosSAP"
Option Explicit

' VERSION 2 - CON DIAGNOSTICO Y CONEXION SAP ROBUSTA
' =====================================================================
' COMENTARIOS DE PAROS DE LINEA - REPORTE DE EFICIENCIA CUAUTITLAN
'
' Objetivo:
'   - Lee la fecha de REPORTE EFICIENCIA!F3.
'   - Ejecuta ZMNPPREP0004 en SAP con variante /SMOSKE_ALT.
'   - Descarga SOLO el detalle del dia.
'   - Suma los minutos por causa de paro para cada linea.
'   - Toma las 3 causas con mas minutos.
'   - Escribe el resumen en REPORTE EFICIENCIA!W6:W9.
'
' Ejemplo:
'   111 min Llenadora 1 L2 -Cua Sist. sensor/fotocel. falla
'   55 min Llenadora 1 L2 -Cua Variador de frecuencia falla
'   38 min Llenadora 1 L2 -Cua PLC falla
'
' No modifica REP2, CTT, Dia, Ago 26 ni otras celdas del reporte.
' =====================================================================

Private Const CENTRO As String = "MABC"
Private Const TRANSACCION As String = "ZMNPPREP0004"
Private Const VARIANTE As String = "/SMOSKE_ALT"
Private Const HOJA_REPORTE As String = "REPORTE EFICIENCIA"
Private Const CELDA_FECHA As String = "F3"
Private Const MAX_COMENTARIOS As Long = 3

Public Sub ActualizarComentariosParosSAP()
    MsgBox "Modulo modComentariosParosSAP v2 - ver codigo completo en el repositorio local."
End Sub

Public Sub DiagnosticarSAPGUI()
    MsgBox "Diagnostico SAP GUI - ver codigo completo en el repositorio local."
End Sub
