Attribute VB_Name = "modReporteSAP_V6_5"
Option Explicit

' VERSION 6.5 - REP2 + PRODUCCION + PDP PRODUCTION PLAN + PAROS / SAP GUI
' =====================================================================
' ACTUALIZACION INTEGRAL - REPORTE DE EFICIENCIA <PLANTA>
'
' Objetivo:
'   - Lee la fecha de REPORTE EFICIENCIA!F3.
'   - Ejecuta ZREP_EFICIENCIA P_ACU una vez SOLO para el dia de corte.
'   - Ejecuta ZREP_EFICIENCIA P_ACU otra vez desde el dia 1 hasta el corte para el MES.
'   - Ejecuta ZREP_EFICIENCIA P_DET para Fecha de inicio Real y % Eficiencia Acumulada.
'   - Actualiza REP2 sin perder sus formulas auxiliares y guarda el detalle en REP2_DETALLE.
'   - Ejecuta ZREP_PAROS con variante /VAR_PAROS solo para el dia de corte.
'   - Agrupa las causas exactas reportadas por SAP, toma TOP 3 y escribe W6:W9.
'   - Genera PAROS_AGRUPADOS incluyendo PRODUCCION para validar 1440 min por linea.
'   - Ejecuta ZREP_PRODUCCION con variante /VAR_PROD y actualiza la hoja Ago 26.
'   - Carga Cajas Unitarias por dia/linea en BE:BH, actualiza MTD y dia de corte.
'   - Busca el Production Plan valido mas reciente en PDP\ENTRADA.
'   - Lee Production Plan_1, convierte CF / Factor CF_CU y actualiza PLAN Ago 26 BJ:BN.
'   - Recalcula el libro para refrescar REPORTE EFICIENCIA y cumplimiento PDP.
'
' Ejemplo:
'   111 min Llenadora 1 L2 -PLT Sist. sensor/fotocel. falla
'   55 min Llenadora 1 L2 -PLT Variador de frecuencia falla
'   38 min Llenadora 1 L2 -PLT PLC falla
'
' Modifica REP2, Ago 26 (REAL SAP + PLAN PDP), REP2_DETALLE, PAROS_AGRUPADOS y REPORTE EFICIENCIA!W6:W9.
' No modifica CTT ni Dia.
' =====================================================================

Private Const CENTRO As String = "<CENTRO>"   ' centro (WERKS) real
Private Const TRANSACCION As String = "ZREP_PAROS"
Private Const VARIANTE As String = "/VAR_PAROS"
Private Const HOJA_REPORTE As String = "REPORTE EFICIENCIA"
Private Const CELDA_FECHA As String = "F3"
Private Const MAX_COMENTARIOS As Long = 3
Private Const TRANSACCION_REP2 As String = "ZREP_EFICIENCIA"
Private Const VARIANTE_REP2 As String = "/VAR_EFIC"
Private Const CENTRO_TRABAJO_REP2 As String = "Linea*"
Private Const HOJA_REP2 As String = "REP2"
Private Const HOJA_REP2_DETALLE As String = "REP2_DETALLE"
Private Const HOJA_PAROS_AGRUPADOS As String = "PAROS_AGRUPADOS"
Private Const TRANSACCION_PRODUCCION As String = "ZREP_PRODUCCION"
Private Const VARIANTE_PRODUCCION As String = "/VAR_PROD"
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

Public Sub V6_ProbarModoAcumuladoSAP()
    Dim ses As Object

    On Error GoTo EH

    Set ses = ObtenerSesionSAP()
    If ses Is Nothing Then
        MsgBox "No se encontro una sesion activa de SAP GUI.", _
               vbExclamation, "PRUEBA P_ACU"
        Exit Sub
    End If

    IrATransaccion ses, TRANSACCION_REP2
    EsperarSAP ses

    ses.findById("wnd[0]/usr/radP_ACU").SetFocus
    ses.findById("wnd[0]/usr/radP_ACU").Select

    MsgBox "MODO ACUMULADO SELECCIONADO CORRECTAMENTE." & vbCrLf & vbCrLf & _
           "ID: wnd[0]/usr/radP_ACU" & vbCrLf & _
           "Texto esperado: Acumulado", _
           vbInformation, "PRUEBA P_ACU"
    Exit Sub

EH:
    MsgBox "No se pudo seleccionar P_ACU." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "PRUEBA P_ACU"
End Sub


Public Sub V6_ProbarPACUDia()

    Dim ses As Object
    Dim fechaReporte As Date
    Dim carpeta As String
    Dim archivo As String
    Dim wb As Workbook

    On Error GoTo EH

    fechaReporte = ObtenerFechaReporte()
    carpeta = ObtenerCarpetaTemporal()

    archivo = carpeta & "\PRUEBA_PACU_DIA_" & _
              Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivo

    Set ses = ObtenerSesionSAP()

    ExportarREP2AcumuladoV6 ses, _
                            fechaReporte, _
                            fechaReporte, _
                            archivo

    EsperarArchivo archivo, 45

    Set wb = AbrirExportSAP(archivo)

    If wb Is Nothing Then
        Err.Raise vbObjectError + 2630, , _
            "No fue posible abrir la prueba P_ACU DIA."
    End If

    CargarREP2AcumuladoV6 wb.Worksheets(1), 3, 7, 8

    wb.Close SaveChanges:=False
    Set wb = Nothing

    Application.CalculateFullRebuild

    MsgBox "P_ACU DIA PROCESADO CORRECTAMENTE." & vbCrLf & vbCrLf & _
           "Periodo consultado:" & vbCrLf & _
           Format$(fechaReporte, "dd/mm/yyyy") & _
           " a " & Format$(fechaReporte, "dd/mm/yyyy"), _
           vbInformation, "PRUEBA P_ACU DIA V6"

    Exit Sub

EH:

    On Error Resume Next
    If Not wb Is Nothing Then wb.Close SaveChanges:=False

    MsgBox "Fallo P_ACU DIA." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "PRUEBA P_ACU DIA V6"

End Sub


Public Sub V6_ActualizarReporteCompletoSAP()

    Dim ses As Object
    Dim fechaReporte As Date
    Dim fechaInicio As Date
    Dim carpeta As String

    Dim archivoDia As String
    Dim archivoMes As String
    Dim archivoDetalle As String
    Dim archivoProduccion As String
    Dim archivoPDP As String
    Dim archivoParos As String

    Dim wbDia As Workbook
    Dim wbMes As Workbook
    Dim wbDetalle As Workbook
    Dim wbProduccion As Workbook
    Dim wbPDPFuente As Workbook
    Dim wbParos As Workbook

    Dim fechaInicioReal As Date
    Dim eficienciaAcumulada As Double

    Dim estadoPantalla As Boolean
    Dim estadoEventos As Boolean
    Dim estadoAlertas As Boolean
    Dim calcAnterior As XlCalculation
    Dim etapa As String

    On Error GoTo ERROR_GENERAL

    estadoPantalla = Application.ScreenUpdating
    estadoEventos = Application.EnableEvents
    estadoAlertas = Application.DisplayAlerts
    calcAnterior = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    etapa = "Validando libro"
    ValidarEstructuraV6

    fechaReporte = ObtenerFechaReporte()
    fechaInicio = DateSerial(Year(fechaReporte), Month(fechaReporte), 1)
    carpeta = ObtenerCarpetaTemporal()

    archivoDia = carpeta & "\REP002_DIA_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoMes = carpeta & "\REP002_MES_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoDetalle = carpeta & "\REP002_DET_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoProduccion = carpeta & "\PRODUCCION_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoParos = carpeta & "\PAROS_" & Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivoDia
    BorrarSiExiste archivoMes
    BorrarSiExiste archivoDetalle
    BorrarSiExiste archivoProduccion
    BorrarSiExiste archivoParos

    etapa = "Conectando a SAP GUI"
    Application.StatusBar = "SAP V6: conectando..."
    Set ses = ObtenerSesionSAP()

    If ses Is Nothing Then
        Err.Raise vbObjectError + 2600, , _
            "No se encontro una sesion activa de SAP GUI."
    End If


    ' ==========================================================
    ' PASO 1 - EFICIENCIA DEL DIA
    '
    ' IMPORTANTE:
    ' P_ACU solo devuelve UNA tabla para el periodo consultado.
    ' Para el DIA consultamos F3 -> F3.
    ' ==========================================================
    etapa = "Descargando eficiencia del DIA"
    Application.StatusBar = "SAP V6: descargando eficiencia del dia..."

    ExportarREP2AcumuladoV6 ses, fechaReporte, fechaReporte, archivoDia
    EsperarArchivo archivoDia, 45

    etapa = "Procesando eficiencia del DIA"
    Set wbDia = AbrirExportSAP(archivoDia)

    If wbDia Is Nothing Then
        Err.Raise vbObjectError + 2601, , _
            "No fue posible abrir el REP2 del DIA."
    End If

    CargarREP2AcumuladoV6 wbDia.Worksheets(1), 3, 7, 8
    wbDia.Close SaveChanges:=False
    Set wbDia = Nothing


    ' ==========================================================
    ' PASO 2 - EFICIENCIA ACUMULADA DEL MES
    '
    ' Segunda consulta P_ACU:
    ' 01 del mes -> F3.
    ' ==========================================================
    etapa = "Descargando eficiencia acumulada del MES"
    Application.StatusBar = "SAP V6: descargando acumulado del mes..."

    ExportarREP2AcumuladoV6 ses, fechaInicio, fechaReporte, archivoMes
    EsperarArchivo archivoMes, 45

    etapa = "Procesando eficiencia acumulada del MES"
    Set wbMes = AbrirExportSAP(archivoMes)

    If wbMes Is Nothing Then
        Err.Raise vbObjectError + 2602, , _
            "No fue posible abrir el REP2 del MES."
    End If

    CargarREP2AcumuladoV6 wbMes.Worksheets(1), 16, 20, 21
    wbMes.Close SaveChanges:=False
    Set wbMes = Nothing


    ' ==========================================================
    ' PASO 3 - DETALLE
    '
    ' Este es el flujo que comprobamos en video:
    ' P_DET -> Ejecutar -> btn[32] -> fila 0 -> fila 15
    ' -> Fecha inicio Real + % Eficiencia Acumulada.
    ' ==========================================================
    etapa = "Descargando REP2 DETALLE"
    Application.StatusBar = _
        "SAP V6: obteniendo Fecha de inicio Real y Eficiencia Acumulada..."

    ExportarREP2DetalleV6 ses, fechaInicio, fechaReporte, archivoDetalle
    EsperarArchivo archivoDetalle, 45

    etapa = "Procesando REP2 DETALLE"
    Set wbDetalle = AbrirExportSAP(archivoDetalle)

    If wbDetalle Is Nothing Then
        Err.Raise vbObjectError + 2603, , _
            "No fue posible abrir el REP2 DETALLE."
    End If

    ProcesarREP2DetalleV6 wbDetalle.Worksheets(1), _
                           fechaReporte, _
                           fechaInicioReal, _
                           eficienciaAcumulada

    wbDetalle.Close SaveChanges:=False
    Set wbDetalle = Nothing

    Application.CalculateFullRebuild


    ' ==========================================================
    ' PASO 4 - PRODUCCION REAL / HOJA AGO 26
    '
    ' ZREP_PRODUCCION /VAR_PROD
    ' SAP entrega Cajas Unitarias por fecha, linea, orden y material.
    ' Se agrupa por DIA + LINEA y se escribe en:
    '   BE = L1
    '   BF = L2
    '   BG = L3
    '   BH = L4
    ' Solo se consideran datos hasta la fecha de corte F3.
    ' ==========================================================
    etapa = "Descargando produccion para Ago 26"
    Application.StatusBar = "SAP V6.3: descargando produccion ZREP_PRODUCCION..."

    ExportarProduccionAgo26V6 ses, fechaInicio, fechaReporte, archivoProduccion
    EsperarArchivo archivoProduccion, 45

    etapa = "Procesando produccion para Ago 26"
    Set wbProduccion = AbrirExportSAP(archivoProduccion)

    If wbProduccion Is Nothing Then
        Err.Raise vbObjectError + 2640, , _
            "No fue posible abrir el archivo de produccion ZREP_PRODUCCION."
    End If

    ProcesarProduccionAgo26V6 wbProduccion.Worksheets(1), fechaReporte

    wbProduccion.Close SaveChanges:=False
    Set wbProduccion = Nothing

    Application.CalculateFullRebuild


    ' ==========================================================
    ' PASO 5 - PDP / PRODUCTION PLAN DESDE CARPETA CONTROLADA
    '
    ' Carpeta:
    '   [ruta del libro]\PDP\ENTRADA
    '
    ' Archivo esperado:
    '   Report of Printout Production Plan <PLANTA>, dd-mm-yy.xlsx
    '   (el nombre puede variar)
    '
    ' Hoja fuente:
    '   Production Plan_1
    '
    ' Logica:
    '   - Detecta cada bloque de dia por su columna Total.
    '   - Lee Linea + SKU + CF del Total diario.
    '   - Convierte CF / Factor CF_CU.
    '   - Suma CU por LINEA001..LINEA004.
    '   - Actualiza solamente los dias presentes en BJ:BM.
    '   - BN = Planta.
    '
    ' NO se toca la tabla REAL BE:BI alimentada desde SAP.
    ' ==========================================================
    etapa = "Actualizando PDP desde carpeta"
    Application.StatusBar = "V6.5: buscando Production Plan..."

    archivoPDP = ObtenerArchivoPDPRecienteV65(fechaReporte)

    If Len(archivoPDP) = 0 Then
        Err.Raise vbObjectError + 2680, , _
            "No se encontro un Production Plan valido en:" & vbCrLf & _
            ObtenerCarpetaPDPV65() & vbCrLf & vbCrLf & _
            "Coloca ahi el archivo Report of Printout Production Plan <PLANTA> mas reciente y vuelve a ejecutar."
    End If

    Set wbPDPFuente = AbrirLibroPDPV65(archivoPDP)

    If wbPDPFuente Is Nothing Then
        Err.Raise vbObjectError + 2681, , _
            "No fue posible abrir el archivo PDP:" & vbCrLf & archivoPDP
    End If

    ProcesarPDPDesdeLibroV65 wbPDPFuente, fechaReporte, archivoPDP

    wbPDPFuente.Close SaveChanges:=False
    Set wbPDPFuente = Nothing

    Application.CalculateFullRebuild


    ' ==========================================================
    ' PASO 6 - COMENTARIOS DE PARO
    ' ==========================================================
    etapa = "Descargando detalle de paros"
    Application.StatusBar = "SAP V6: descargando comentarios de paros..."

    ExportarDetalleParos ses, fechaReporte, archivoParos
    EsperarArchivo archivoParos, 45

    etapa = "Procesando comentarios de paros"
    Set wbParos = AbrirExportSAP(archivoParos)

    If wbParos Is Nothing Then
        Err.Raise vbObjectError + 2604, , _
            "No fue posible abrir el detalle de paros."
    End If

    ProcesarComentarios wbParos.Worksheets(1), fechaReporte

    wbParos.Close SaveChanges:=False
    Set wbParos = Nothing


    ' ==========================================================
    ' FINAL
    ' ==========================================================
    etapa = "Recalculando reporte"
    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    ThisWorkbook.Worksheets(HOJA_REPORTE).Activate

    MsgBox "REPORTE ACTUALIZADO CORRECTAMENTE" & vbCrLf & vbCrLf & _
           "Fecha de corte: " & Format$(fechaReporte, "dd/mm/yyyy") & vbCrLf & _
           "Inicio real SAP: " & Format$(fechaInicioReal, "dd/mm/yyyy") & vbCrLf & _
           "Eficiencia acumulada SAP: " & _
               Format$(eficienciaAcumulada, "0.00") & "%" & vbCrLf & vbCrLf & _
           "P_ACU DIA: OK" & vbCrLf & _
           "P_ACU MES: OK" & vbCrLf & _
           "P_DET: OK" & vbCrLf & _
           "Produccion / Ago 26: OK" & vbCrLf & _
           "PDP: " & NombreDeRuta(archivoPDP) & vbCrLf & _
           "Comentarios W6:W9: OK", _
           vbInformation, "REPORTE SAP V6"


SALIDA:
    On Error Resume Next

    If Not wbDia Is Nothing Then wbDia.Close SaveChanges:=False
    If Not wbMes Is Nothing Then wbMes.Close SaveChanges:=False
    If Not wbDetalle Is Nothing Then wbDetalle.Close SaveChanges:=False
    If Not wbProduccion Is Nothing Then wbProduccion.Close SaveChanges:=False
    If Not wbPDPFuente Is Nothing Then wbPDPFuente.Close SaveChanges:=False
    If Not wbParos Is Nothing Then wbParos.Close SaveChanges:=False

    BorrarSiExiste archivoDia
    BorrarSiExiste archivoMes
    BorrarSiExiste archivoDetalle
    BorrarSiExiste archivoProduccion
    BorrarSiExiste archivoParos

    Application.StatusBar = False
    Application.Calculation = calcAnterior
    Application.DisplayAlerts = estadoAlertas
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoPantalla

    Exit Sub


ERROR_GENERAL:

    MsgBox "No se pudo completar la actualizacion V6." & vbCrLf & vbCrLf & _
           "Etapa: " & etapa & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "REPORTE SAP V6"

    Resume SALIDA

End Sub


Public Sub V6_ProbarSoloDetalleConCamposSAP()
    Dim ses As Object
    Dim wsReporte As Worksheet
    Dim fechaReporte As Date
    Dim fechaInicio As Date
    Dim carpeta As String
    Dim archivo As String

    On Error GoTo EH

    Set wsReporte = ThisWorkbook.Worksheets(HOJA_REPORTE)
    fechaReporte = ObtenerFechaReporte()

    If fechaReporte <= 0 Then
        Err.Raise vbObjectError + 2501, , _
            "No se pudo determinar la fecha desde " & HOJA_REPORTE & "!" & CELDA_FECHA
    End If

    fechaInicio = DateSerial(Year(fechaReporte), Month(fechaReporte), 1)
    carpeta = ObtenerCarpetaTemporal()
    archivo = carpeta & "\PRUEBA_REP2_DET_" & Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivo

    Set ses = ObtenerSesionSAP()
    If ses Is Nothing Then
        Err.Raise vbObjectError + 2502, , "No se encontro una sesion activa de SAP."
    End If

    Application.StatusBar = _
        "SAP V6: Detalle -> seleccionando Fecha inicio Real y Eficiencia Acumulada..."

    ExportarREP2DetalleV6 ses, fechaInicio, fechaReporte, archivo
    EsperarArchivo archivo, 45

    Application.StatusBar = False

    MsgBox "PRUEBA COMPLETADA" & vbCrLf & vbCrLf & _
           "SAP debio ejecutar exactamente estos pasos:" & vbCrLf & _
           "1. Seleccionar DETALLE" & vbCrLf & _
           "2. Abrir layout (boton 32)" & vbCrLf & _
           "3. Agregar fila 0 = Fecha inicio Real" & vbCrLf & _
           "4. Agregar fila 15 = % Eficiencia Acumulada" & vbCrLf & _
           "5. Exportar" & vbCrLf & vbCrLf & _
           "Archivo: " & archivo, _
           vbInformation, "PRUEBA REP2 DETALLE V6"
    Exit Sub

EH:
    Application.StatusBar = False
    MsgBox "Fallo la prueba de REP2 DETALLE." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "PRUEBA REP2 DETALLE V6"
End Sub


Public Sub V6_ActualizarSoloREP2SAP()

    Dim ses As Object
    Dim fechaReporte As Date
    Dim fechaInicio As Date
    Dim carpeta As String

    Dim archivoDia As String
    Dim archivoMes As String
    Dim archivoDetalle As String

    Dim wbDia As Workbook
    Dim wbMes As Workbook
    Dim wbDetalle As Workbook

    Dim fechaInicioReal As Date
    Dim eficienciaAcumulada As Double
    Dim calcAnterior As XlCalculation
    Dim etapa As String

    On Error GoTo EH

    ValidarEstructuraV6

    calcAnterior = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    fechaReporte = ObtenerFechaReporte()
    fechaInicio = DateSerial(Year(fechaReporte), Month(fechaReporte), 1)
    carpeta = ObtenerCarpetaTemporal()

    archivoDia = carpeta & "\REP002_DIA_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoMes = carpeta & "\REP002_MES_" & Format$(fechaReporte, "yyyymmdd") & ".xls"
    archivoDetalle = carpeta & "\REP002_DET_" & Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivoDia
    BorrarSiExiste archivoMes
    BorrarSiExiste archivoDetalle

    Set ses = ObtenerSesionSAP()

    If ses Is Nothing Then
        Err.Raise vbObjectError + 2610, , _
            "No se encontro una sesion activa de SAP GUI."
    End If


    ' -------------------------------
    ' DIA = fecha corte -> fecha corte
    ' -------------------------------
    etapa = "P_ACU DIA"
    Application.StatusBar = "SAP V6: P_ACU del dia..."

    ExportarREP2AcumuladoV6 ses, fechaReporte, fechaReporte, archivoDia
    EsperarArchivo archivoDia, 45

    Set wbDia = AbrirExportSAP(archivoDia)

    If wbDia Is Nothing Then
        Err.Raise vbObjectError + 2611, , _
            "No fue posible abrir el archivo P_ACU DIA."
    End If

    CargarREP2AcumuladoV6 wbDia.Worksheets(1), 3, 7, 8

    wbDia.Close SaveChanges:=False
    Set wbDia = Nothing


    ' -------------------------------
    ' MES = dia 1 -> fecha corte
    ' -------------------------------
    etapa = "P_ACU MES"
    Application.StatusBar = "SAP V6: P_ACU acumulado mes..."

    ExportarREP2AcumuladoV6 ses, fechaInicio, fechaReporte, archivoMes
    EsperarArchivo archivoMes, 45

    Set wbMes = AbrirExportSAP(archivoMes)

    If wbMes Is Nothing Then
        Err.Raise vbObjectError + 2612, , _
            "No fue posible abrir el archivo P_ACU MES."
    End If

    CargarREP2AcumuladoV6 wbMes.Worksheets(1), 16, 20, 21

    wbMes.Close SaveChanges:=False
    Set wbMes = Nothing


    ' -------------------------------
    ' DETALLE = fecha real + efi acum.
    ' -------------------------------
    etapa = "P_DET"
    Application.StatusBar = "SAP V6: P_DET + campos adicionales..."

    ExportarREP2DetalleV6 ses, fechaInicio, fechaReporte, archivoDetalle
    EsperarArchivo archivoDetalle, 45

    Set wbDetalle = AbrirExportSAP(archivoDetalle)

    If wbDetalle Is Nothing Then
        Err.Raise vbObjectError + 2613, , _
            "No fue posible abrir REP2 DETALLE."
    End If

    ProcesarREP2DetalleV6 wbDetalle.Worksheets(1), _
                           fechaReporte, _
                           fechaInicioReal, _
                           eficienciaAcumulada

    wbDetalle.Close SaveChanges:=False
    Set wbDetalle = Nothing


    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    MsgBox "REP2 ACTUALIZADO" & vbCrLf & vbCrLf & _
           "DIA: " & Format$(fechaReporte, "dd/mm/yyyy") & vbCrLf & _
           "MES: " & Format$(fechaInicio, "dd/mm/yyyy") & _
               " al " & Format$(fechaReporte, "dd/mm/yyyy") & vbCrLf & _
           "Inicio real: " & Format$(fechaInicioReal, "dd/mm/yyyy") & vbCrLf & _
           "Eficiencia acumulada: " & _
               Format$(eficienciaAcumulada, "0.00") & "%", _
           vbInformation, "REP2 SAP V6"


SALIDA_REP2:
    On Error Resume Next

    If Not wbDia Is Nothing Then wbDia.Close SaveChanges:=False
    If Not wbMes Is Nothing Then wbMes.Close SaveChanges:=False
    If Not wbDetalle Is Nothing Then wbDetalle.Close SaveChanges:=False

    BorrarSiExiste archivoDia
    BorrarSiExiste archivoMes
    BorrarSiExiste archivoDetalle

    Application.StatusBar = False
    Application.Calculation = calcAnterior
    Application.DisplayAlerts = True
    Application.EnableEvents = True
    Application.ScreenUpdating = True

    Exit Sub


EH:

    MsgBox "No se pudo actualizar REP2 V6." & vbCrLf & vbCrLf & _
           "Etapa: " & etapa & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbExclamation, "REP2 SAP V6"

    Resume SALIDA_REP2

End Sub


Private Sub ValidarEstructuraV6()
    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_REPORTE)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2410, , _
            "No existe la hoja '" & HOJA_REPORTE & "'."
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_REP2)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2411, , _
            "No existe la hoja '" & HOJA_REP2 & "'. Usa el archivo operativo original que contiene REP2."
    End If

    Set ws = Nothing
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_PDP)
    On Error GoTo 0
    If ws Is Nothing Then
        Err.Raise vbObjectError + 2412, , _
            "No existe la hoja '" & HOJA_PDP & "'."
    End If
End Sub

Public Sub V6_ActualizarSoloAgo26SAP()

    Dim ses As Object
    Dim fechaReporte As Date
    Dim fechaInicio As Date
    Dim carpeta As String
    Dim archivo As String
    Dim wbProduccion As Workbook
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

    ValidarEstructuraV6

    fechaReporte = ObtenerFechaReporte()
    fechaInicio = DateSerial(Year(fechaReporte), Month(fechaReporte), 1)

    carpeta = ObtenerCarpetaTemporal()
    archivo = carpeta & "\PRODUCCION_" & Format$(fechaReporte, "yyyymmdd") & ".xls"

    BorrarSiExiste archivo

    Set ses = ObtenerSesionSAP()

    If ses Is Nothing Then
        Err.Raise vbObjectError + 2641, , _
            "No se encontro una sesion activa de SAP GUI."
    End If

    Application.StatusBar = _
        "SAP V6.3: actualizando Produccion Real de " & HOJA_PDP & "..."

    ExportarProduccionAgo26V6 ses, fechaInicio, fechaReporte, archivo
    EsperarArchivo archivo, 45

    Set wbProduccion = AbrirExportSAP(archivo)

    If wbProduccion Is Nothing Then
        Err.Raise vbObjectError + 2642, , _
            "No fue posible abrir el archivo PRODUCCION exportado por SAP."
    End If

    ProcesarProduccionAgo26V6 wbProduccion.Worksheets(1), fechaReporte

    wbProduccion.Close SaveChanges:=False
    Set wbProduccion = Nothing

    Application.CalculateFullRebuild

    ThisWorkbook.Worksheets(HOJA_PDP).Activate
    ThisWorkbook.Worksheets(HOJA_PDP).Range("BE3").Select

    MsgBox _
        "AGO 26 ACTUALIZADO CORRECTAMENTE" & vbCrLf & vbCrLf & _
        "Fuente SAP: " & TRANSACCION_PRODUCCION & vbCrLf & _
        "Variante: " & VARIANTE_PRODUCCION & vbCrLf & _
        "Corte: " & Format$(fechaReporte, "dd/mm/yyyy") & vbCrLf & _
        "Dato usado: Cajas Unitarias" & vbCrLf & vbCrLf & _
        "REAL diario BE:BH: OK" & vbCrLf & _
        "Produccion Real MTD I6:I10: OK" & vbCrLf & _
        "Produccion real dia R6:R10: OK", _
        vbInformation, "PRODUCCION SAP / AGO 26"

SALIDA_AGO26:
    On Error Resume Next
    If Not wbProduccion Is Nothing Then wbProduccion.Close SaveChanges:=False
    BorrarSiExiste archivo
    Application.StatusBar = False
    Application.DisplayAlerts = estadoAlertas
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoPantalla
    Exit Sub

ERROR_GENERAL:
    MsgBox _
        "No se pudo actualizar la hoja " & HOJA_PDP & "." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, "PRODUCCION SAP / AGO 26"
    Resume SALIDA_AGO26

End Sub



' =====================================================================
' PDP DESDE CARPETA CONTROLADA
' =====================================================================

Public Sub V6_AbrirCarpetaPDP()

    Dim carpeta As String

    On Error GoTo EH

    carpeta = ObtenerCarpetaPDPV65()

    Shell "explorer.exe """ & carpeta & """", vbNormalFocus

    Exit Sub

EH:
    MsgBox _
        "No se pudo abrir la carpeta PDP." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, "PDP"

End Sub


Public Sub V6_ActualizarSoloPDPDesdeCarpeta()

    Dim fechaReporte As Date
    Dim archivoPDP As String
    Dim wbFuente As Workbook

    Dim estadoPantalla As Boolean
    Dim estadoEventos As Boolean
    Dim estadoAlertas As Boolean
    Dim calcAnterior As XlCalculation

    On Error GoTo ERROR_GENERAL

    estadoPantalla = Application.ScreenUpdating
    estadoEventos = Application.EnableEvents
    estadoAlertas = Application.DisplayAlerts
    calcAnterior = Application.Calculation

    Application.ScreenUpdating = False
    Application.EnableEvents = False
    Application.DisplayAlerts = False
    Application.Calculation = xlCalculationManual

    ValidarEstructuraV6

    fechaReporte = ObtenerFechaReporte()

    Application.StatusBar = "V6.5: buscando Production Plan..."

    archivoPDP = ObtenerArchivoPDPRecienteV65(fechaReporte)

    If Len(archivoPDP) = 0 Then
        Err.Raise vbObjectError + 2682, , _
            "No se encontro un Production Plan valido." & vbCrLf & vbCrLf & _
            "Carpeta esperada:" & vbCrLf & _
            ObtenerCarpetaPDPV65() & vbCrLf & vbCrLf & _
            "El archivo debe contener la hoja '" & HOJA_PDP_FUENTE & "' " & _
            "con Linea, Producto, SKU, Tarimas y bloques T1/T2/T3/Total."
    End If

    Set wbFuente = AbrirLibroPDPV65(archivoPDP)

    If wbFuente Is Nothing Then
        Err.Raise vbObjectError + 2683, , _
            "No fue posible abrir:" & vbCrLf & archivoPDP
    End If

    ProcesarPDPDesdeLibroV65 wbFuente, fechaReporte, archivoPDP

    wbFuente.Close SaveChanges:=False
    Set wbFuente = Nothing

    Application.Calculation = xlCalculationAutomatic
    Application.CalculateFullRebuild

    ThisWorkbook.Worksheets(HOJA_PDP).Activate
    ThisWorkbook.Worksheets(HOJA_PDP).Range("BJ1").Select

    MsgBox _
        "PDP ACTUALIZADO CORRECTAMENTE" & vbCrLf & vbCrLf & _
        "Archivo utilizado:" & vbCrLf & _
        NombreDeRuta(archivoPDP) & vbCrLf & vbCrLf & _
        "Fecha del archivo: " & _
        Format$(FileDateTime(archivoPDP), "dd/mm/yyyy hh:mm") & vbCrLf & _
        "Periodo del reporte: " & _
        Format$(fechaReporte, "mmmm yyyy") & vbCrLf & vbCrLf & _
        "Production Plan_1: OK" & vbCrLf & _
        "Conversion CF -> CU: OK" & vbCrLf & _
        "PLAN BJ:BN: OK" & vbCrLf & _
        "PDP dia S6:S10: OK" & vbCrLf & _
        "Cumplimiento BO:BS: OK", _
        vbInformation, "PDP V6.5"


SALIDA_PDP:

    On Error Resume Next

    If Not wbFuente Is Nothing Then
        wbFuente.Close SaveChanges:=False
    End If

    Application.StatusBar = False
    Application.Calculation = calcAnterior
    Application.DisplayAlerts = estadoAlertas
    Application.EnableEvents = estadoEventos
    Application.ScreenUpdating = estadoPantalla

    Exit Sub


ERROR_GENERAL:

    MsgBox _
        "No se pudo actualizar el PDP." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, "PDP V6.5"

    Resume SALIDA_PDP

End Sub


Public Sub V6_ActualizarSoloComentariosSAP()
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
    Dim i As Integer
    Dim j As Integer
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

    ' Ruta rapida: replica exactamente el patron generado por el grabador SAP.
    ' Los literales 0 evitan el error 618 de tipo de indice en algunas versiones de SAP GUI.
    On Error Resume Next
    Set con = appSAP.Children(0)
    If Not con Is Nothing Then
        If con.Children.Count > 0 Then
            Set ses = con.Children(0)
            If Not ses Is Nothing Then
                Set wnd0 = Nothing
                Set wnd0 = ses.findById("wnd[0]")
                If Err.Number = 0 And Not wnd0 Is Nothing Then
                    On Error GoTo 0
                    Set ObtenerSesionSAP = ses
                    Exit Function
                End If
            End If
        End If
    End If
    Err.Clear
    On Error GoTo 0

    ' Si hubiera varias conexiones/sesiones, recorrerlas usando Integer/CInt.
    For i = 0 To appSAP.Children.Count - 1
        Set con = Nothing
        On Error Resume Next
        Set con = appSAP.Children(CInt(i))
        On Error GoTo 0

        If Not con Is Nothing Then
            For j = 0 To con.Children.Count - 1
                Set ses = Nothing
                On Error Resume Next
                Set ses = con.Children(CInt(j))
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

Public Sub V6_ProbarConexionSAPMinima()
    Dim SapGuiAuto As Object
    Dim applicationSAP As Object
    Dim connectionSAP As Object
    Dim sessionSAP As Object
    Dim mensaje As String

    On Error GoTo EH

    ' Mismo patron que genera el grabador oficial de SAP GUI Scripting.
    Set SapGuiAuto = GetObject("SAPGUI")
    Set applicationSAP = SapGuiAuto.GetScriptingEngine
    Set connectionSAP = applicationSAP.Children(0)
    Set sessionSAP = connectionSAP.Children(0)

    mensaje = "CONEXION SAP OK" & vbCrLf & vbCrLf
    On Error Resume Next
    mensaje = mensaje & "Sistema: " & sessionSAP.Info.SystemName & vbCrLf
    mensaje = mensaje & "Mandante: " & sessionSAP.Info.Client & vbCrLf
    mensaje = mensaje & "Usuario: " & sessionSAP.Info.User & vbCrLf
    mensaje = mensaje & "Transaccion: " & sessionSAP.Info.Transaction
    On Error GoTo 0

    MsgBox mensaje, vbInformation, "SAP GUI - PRUEBA MINIMA"
    Exit Sub
EH:
    MsgBox "Fallo la prueba minima." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "SAP GUI - PRUEBA MINIMA"
End Sub

Public Sub V6_DiagnosticarSAPGUI()
    Dim SapGuiAuto As Object
    Dim appSAP As Object
    Dim con As Object
    Dim ses As Object
    Dim rot As Object
    Dim i As Integer
    Dim j As Integer
    Dim txt As String
    Dim metodo As String
    Dim nCon As Long
    Dim nSes As Long

    On Error Resume Next
    Set SapGuiAuto = GetObject("SAPGUI")
    If Not SapGuiAuto Is Nothing Then metodo = "GetObject(SAPGUI)"
    Err.Clear
    On Error GoTo 0

    If SapGuiAuto Is Nothing Then
        On Error Resume Next
        Set rot = CreateObject("SapROTWr.SapROTWrapper")
        If Not rot Is Nothing Then Set SapGuiAuto = rot.GetROTEntry("SAPGUI")
        If Not SapGuiAuto Is Nothing Then metodo = "SapROTWrapper"
        Err.Clear
        On Error GoTo 0
    End If

    If SapGuiAuto Is Nothing Then
        MsgBox "DIAGNOSTICO SAP GUI" & vbCrLf & vbCrLf & _
               "NO se encontro el objeto SAPGUI en Windows." & vbCrLf & vbCrLf & _
               "Abre SAP GUI, inicia sesion y deja una pantalla SAP abierta.", _
               vbCritical, "DIAGNOSTICO SAP"
        Exit Sub
    End If

    ' PASO 1: motor de scripting
    On Error GoTo ERR_GET_ENGINE
    Set appSAP = SapGuiAuto.GetScriptingEngine
    On Error GoTo 0

    If appSAP Is Nothing Then
        MsgBox "SAPGUI fue detectado, pero GetScriptingEngine devolvio Nothing.", _
               vbCritical, "DIAGNOSTICO SAP"
        Exit Sub
    End If

    ' PASO 2: contar conexiones. Si falla aqui ya NO se reporta falsamente
    ' como error de GetScriptingEngine.
    On Error GoTo ERR_COLLECTION
    nCon = appSAP.Children.Count
    On Error GoTo 0

    txt = "DIAGNOSTICO SAP GUI - V6" & vbCrLf & vbCrLf & _
          "GetScriptingEngine: OK" & vbCrLf & _
          "Metodo: " & metodo & vbCrLf & _
          "Conexiones detectadas: " & nCon & vbCrLf

    If nCon = 0 Then
        txt = txt & vbCrLf & "No hay conexiones activas. Inicia sesion en SAP."
        MsgBox txt, vbExclamation, "DIAGNOSTICO SAP"
        Exit Sub
    End If

    ' Primero probamos exactamente como lo hace el grabador SAP: Children(0).
    On Error GoTo ERR_COLLECTION
    Set con = appSAP.Children(0)
    nSes = con.Children.Count
    txt = txt & vbCrLf & "Conexion 1: " & nSes & " sesion(es)"

    If nSes > 0 Then
        Set ses = con.Children(0)
        txt = txt & vbCrLf & "Sesion 1: OK"
        On Error Resume Next
        txt = txt & " | Sistema: " & ses.Info.SystemName
        txt = txt & " | Mandante: " & ses.Info.Client
        txt = txt & " | Usuario: " & ses.Info.User
        txt = txt & " | TX: " & ses.Info.Transaction
        Err.Clear
        On Error GoTo 0
    End If

    ' Si hay mas conexiones, recorrerlas con Integer/CInt para evitar Error 618.
    If nCon > 1 Then
        For i = 1 To CInt(nCon - 1)
            On Error GoTo ERR_COLLECTION
            Set con = appSAP.Children(CInt(i))
            nSes = con.Children.Count
            txt = txt & vbCrLf & "Conexion " & (i + 1) & ": " & nSes & " sesion(es)"

            If nSes > 0 Then
                For j = 0 To CInt(nSes - 1)
                    Set ses = con.Children(CInt(j))
                    txt = txt & vbCrLf & "  Sesion " & (j + 1) & ": OK"
                Next j
            End If
        Next i
    End If

    txt = txt & vbCrLf & vbCrLf & "RESULTADO: conexion SAP GUI Scripting correcta."
    MsgBox txt, vbInformation, "DIAGNOSTICO SAP"
    Exit Sub

ERR_GET_ENGINE:
    MsgBox "SAPGUI fue detectado, pero GetScriptingEngine SI fallo." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description, _
           vbCritical, "DIAGNOSTICO SAP"
    Exit Sub

ERR_COLLECTION:
    MsgBox "GetScriptingEngine funciono, pero fallo al acceder a Children/Sessions." & vbCrLf & vbCrLf & _
           "Error " & Err.Number & ": " & Err.Description & vbCrLf & vbCrLf & _
           "La V6.1 usa indices Integer/CInt para corregir el Error 618.", _
           vbCritical, "DIAGNOSTICO SAP"
End Sub


' =====================================================================

 ' =====================================================================
' PDP V6.5 - PRODUCTION PLAN -> CF / FACTOR CF_CU -> CU
' =====================================================================

Private Function ObtenerCarpetaPDPV65() As String

    Dim base As String
    Dim carpetaPadre As String
    Dim carpetaEntrada As String

    base = ThisWorkbook.Path

    If Len(base) = 0 Then
        base = Environ$("USERPROFILE") & "\Documents"
    End If

    carpetaPadre = base & "\PDP"
    carpetaEntrada = base & "\" & CARPETA_PDP_RELATIVA

    If Dir$(carpetaPadre, vbDirectory) = vbNullString Then
        MkDir carpetaPadre
    End If

    If Dir$(carpetaEntrada, vbDirectory) = vbNullString Then
        MkDir carpetaEntrada
    End If

    ObtenerCarpetaPDPV65 = carpetaEntrada

End Function


Private Function ObtenerArchivoPDPRecienteV65( _
                    ByVal fechaCorte As Date) As String

    Dim carpeta As String
    Dim nombre As String
    Dim ruta As String

    Dim fechaArchivo As Date
    Dim fechaMasReciente As Date

    carpeta = ObtenerCarpetaPDPV65()

    ' La carpeta es controlada, por lo que se revisan todos los Excel.
    ' Esto permite pequenas variaciones en el nombre del archivo.
    nombre = Dir$(carpeta & "\*.xls*")

    Do While Len(nombre) > 0

        If Left$(nombre, 2) <> "~$" Then

            ruta = carpeta & "\" & nombre

            If LCase$(ruta) <> LCase$(ThisWorkbook.FullName) Then

                If EsArchivoPDPValidoV65(ruta, fechaCorte) Then

                    On Error Resume Next
                    fechaArchivo = FileDateTime(ruta)

                    If Err.Number = 0 Then

                        If Len(ObtenerArchivoPDPRecienteV65) = 0 Or _
                           fechaArchivo > fechaMasReciente Then

                            fechaMasReciente = fechaArchivo
                            ObtenerArchivoPDPRecienteV65 = ruta

                        End If

                    End If

                    Err.Clear
                    On Error GoTo 0

                End If

            End If

        End If

        nombre = Dir$()

    Loop

End Function


Private Function EsArchivoPDPValidoV65( _
                    ByVal ruta As String, _
                    ByVal fechaCorte As Date) As Boolean

    Dim wb As Workbook
    Dim ws As Worksheet

    Dim filaCab As Long
    Dim colLinea As Long
    Dim colProducto As Long
    Dim colSKU As Long
    Dim colTarimas As Long

    Dim ultCol As Long
    Dim c As Long
    Dim f As Date
    Dim txt As String

    On Error GoTo NO_VALIDO

    Set wb = Workbooks.Open( _
                    Filename:=ruta, _
                    UpdateLinks:=False, _
                    ReadOnly:=True, _
                    AddToMru:=False)

    On Error Resume Next
    Set ws = wb.Worksheets(HOJA_PDP_FUENTE)
    On Error GoTo NO_VALIDO

    If ws Is Nothing Then GoTo NO_VALIDO

    filaCab = BuscarFilaCabeceraPlanV65( _
                    ws, _
                    colLinea, _
                    colProducto, _
                    colSKU, _
                    colTarimas)

    If filaCab = 0 Then GoTo NO_VALIDO

    ultCol = UltimaColumnaUsadaV6(ws)

    For c = colTarimas + 1 To ultCol

        txt = NormalizarSAPV6(CStr( _
                    ValorCeldaFusionadaPDPV65( _
                        ws.Cells(filaCab, c))))

        If txt = "TOTAL" Then

            f = FechaDesdeEncabezadoPDPV65( _
                    ValorCeldaFusionadaPDPV65( _
                        ws.Cells(filaCab - 1, c)))

            If f > 0 Then

                If Year(f) = Year(fechaCorte) And _
                   Month(f) = Month(fechaCorte) Then

                    EsArchivoPDPValidoV65 = True
                    GoTo SALIR

                End If

            End If

        End If

    Next c


NO_VALIDO:

    EsArchivoPDPValidoV65 = False


SALIR:

    On Error Resume Next

    If Not wb Is Nothing Then
        wb.Close SaveChanges:=False
    End If

End Function


Private Function AbrirLibroPDPV65(ByVal ruta As String) As Workbook

    On Error Resume Next

    Set AbrirLibroPDPV65 = Workbooks.Open( _
                                Filename:=ruta, _
                                UpdateLinks:=False, _
                                ReadOnly:=True, _
                                AddToMru:=False)

    On Error GoTo 0

End Function


Private Sub ProcesarPDPDesdeLibroV65( _
                    ByVal wbFuente As Workbook, _
                    ByVal fechaCorte As Date, _
                    ByVal rutaFuente As String)

    Dim wsSrc As Worksheet
    Dim wsDst As Worksheet
    Dim wsRep As Worksheet
    Dim wsAud As Worksheet

    Dim filaCab As Long
    Dim colLinea As Long
    Dim colProducto As Long
    Dim colSKU As Long
    Dim colTarimas As Long

    Dim ultFila As Long
    Dim ultCol As Long

    Dim dictFactor As Object
    Dim dictTotales As Object
    Dim dictDias As Object
    Dim dictFaltantes As Object

    Dim detalle As Collection

    Dim c As Long
    Dim r As Long
    Dim i As Long

    Dim fDia As Date
    Dim claveDia As String
    Dim claveTotal As String

    Dim linea As String
    Dim producto As String
    Dim sku As String

    Dim cf As Double
    Dim factor As Double
    Dim cu As Double

    Dim k As Variant
    Dim lineas As Variant
    Dim filaDestino As Long
    Dim valorLinea As Double

    Dim primeraFecha As Date
    Dim ultimaFecha As Date

    On Error GoTo EH

    Set wsSrc = wbFuente.Worksheets(HOJA_PDP_FUENTE)
    Set wsDst = ThisWorkbook.Worksheets(HOJA_PDP)
    Set wsRep = ThisWorkbook.Worksheets(HOJA_REPORTE)

    filaCab = BuscarFilaCabeceraPlanV65( _
                    wsSrc, _
                    colLinea, _
                    colProducto, _
                    colSKU, _
                    colTarimas)

    If filaCab = 0 Then

        Err.Raise vbObjectError + 2690, , _
            "No se encontro la cabecera del Production Plan. " & _
            "Se esperaban las columnas Linea, Producto, SKU y Tarimas."

    End If

    ultFila = UltimaFilaUsada(wsSrc)
    ultCol = UltimaColumnaUsadaV6(wsSrc)

    Set dictFactor = CrearCatalogoFactorCUV65()
    Set dictTotales = CreateObject("Scripting.Dictionary")
    Set dictDias = CreateObject("Scripting.Dictionary")
    Set dictFaltantes = CreateObject("Scripting.Dictionary")
    Set detalle = New Collection

    dictFactor.CompareMode = vbTextCompare
    dictTotales.CompareMode = vbTextCompare
    dictDias.CompareMode = vbTextCompare
    dictFaltantes.CompareMode = vbTextCompare

    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")


    ' =========================================================
    ' RECORRER TODOS LOS BLOQUES DE DIA
    '
    ' En la fila de cabecera existen grupos:
    '   T1 | T2 | T3 | Total
    '
    ' La fecha se obtiene de la celda fusionada inmediatamente
    ' superior al grupo.
    ' =========================================================
    For c = colTarimas + 1 To ultCol

        If NormalizarSAPV6(CStr( _
                ValorCeldaFusionadaPDPV65( _
                    wsSrc.Cells(filaCab, c)))) = "TOTAL" Then

            fDia = FechaDesdeEncabezadoPDPV65( _
                        ValorCeldaFusionadaPDPV65( _
                            wsSrc.Cells(filaCab - 1, c)))

            If fDia > 0 Then

                ' Solo el mes del reporte.
                If Year(fDia) = Year(fechaCorte) And _
                   Month(fDia) = Month(fechaCorte) Then

                    claveDia = CStr(CLng(fDia))

                    If Not dictDias.Exists(claveDia) Then
                        dictDias.Add claveDia, True
                    End If

                    If primeraFecha = 0 Or fDia < primeraFecha Then
                        primeraFecha = fDia
                    End If

                    If ultimaFecha = 0 Or fDia > ultimaFecha Then
                        ultimaFecha = fDia
                    End If


                    ' -----------------------------------------
                    ' LEER SKU POR SKU
                    ' -----------------------------------------
                    For r = filaCab + 1 To ultFila

                        linea = NormalizarLineaPDPV65( _
                                    CStr( _
                                        ValorCeldaFusionadaPDPV65( _
                                            wsSrc.Cells(r, colLinea))))

                        If Len(linea) > 0 Then

                            sku = NormalizarSKUPDPV65( _
                                        wsSrc.Cells(r, colSKU).Value)

                            cf = NumeroPDPV65( _
                                        wsSrc.Cells(r, c).Value)

                            If Len(sku) > 0 And Abs(cf) > 0.0000001 Then

                                producto = Trim$(CStr( _
                                                wsSrc.Cells(r, colProducto).Value))

                                If dictFactor.Exists(sku) Then

                                    factor = CDbl(dictFactor.Item(sku))

                                    If factor <= 0 Then

                                        Err.Raise vbObjectError + 2691, , _
                                            "El SKU " & sku & _
                                            " tiene Factor CF_CU igual a cero."

                                    End If

                                    ' REGLA SOLICITADA:
                                    ' CU = CF / Factor CF_CU
                                    cu = cf / factor

                                    claveTotal = claveDia & "|" & linea

                                    If dictTotales.Exists(claveTotal) Then
                                        dictTotales.Item(claveTotal) = _
                                            CDbl(dictTotales.Item(claveTotal)) + cu
                                    Else
                                        dictTotales.Add claveTotal, cu
                                    End If

                                    detalle.Add Array( _
                                        fDia, _
                                        linea, _
                                        producto, _
                                        sku, _
                                        cf, _
                                        factor, _
                                        cu, _
                                        "CONVERTIDO", _
                                        NombreDeRuta(rutaFuente))

                                ElseIf EsSKUNumericoPDPV65(sku) Then

                                    If Not dictFaltantes.Exists(sku) Then
                                        dictFaltantes.Add sku, producto
                                    End If

                                    detalle.Add Array( _
                                        fDia, _
                                        linea, _
                                        producto, _
                                        sku, _
                                        cf, _
                                        Empty, _
                                        Empty, _
                                        "SIN FACTOR CF_CU", _
                                        NombreDeRuta(rutaFuente))

                                Else

                                    ' Ejemplos vistos en el plan:
                                    ' ZMAN01 / ZSANCAM.
                                    ' No son SKU numericos del catalogo CF->CU.
                                    detalle.Add Array( _
                                        fDia, _
                                        linea, _
                                        producto, _
                                        sku, _
                                        cf, _
                                        Empty, _
                                        Empty, _
                                        "IGNORADO AUXILIAR", _
                                        NombreDeRuta(rutaFuente))

                                End If

                            End If

                        End If

                    Next r

                End If

            End If

        End If

    Next c


    ' =========================================================
    ' AUDITORIA DE CONVERSION
    ' Se genera incluso si falta un SKU para que sea facil
    ' identificar exactamente cual se debe agregar al catalogo.
    ' =========================================================
    Set wsAud = PrepararHojaPDPConversionV65()

    EscribirDetallePDPConversionV65 _
        wsAud, _
        detalle, _
        dictTotales, _
        dictDias, _
        rutaFuente


    If dictDias.Count = 0 Then

        Err.Raise vbObjectError + 2692, , _
            "No se encontraron bloques de fecha del mes " & _
            Format$(fechaCorte, "mmmm yyyy") & _
            " en la hoja " & HOJA_PDP_FUENTE & "."

    End If


    ' =========================================================
    ' NO ACTUALIZAR EL REPORTE SI FALTA UN SKU NUMERICO.
    '
    ' Esto evita subestimar el PDP por una conversion incompleta.
    ' =========================================================
    If dictFaltantes.Count > 0 Then

        Err.Raise vbObjectError + 2693, , _
            "Faltan factores CF_CU para estos SKU:" & vbCrLf & _
            Join(dictFaltantes.Keys, ", ") & vbCrLf & vbCrLf & _
            "Agregalos en la hoja " & HOJA_CATALOGO_PDP & _
            " y vuelve a ejecutar." & vbCrLf & vbCrLf & _
            "Se genero " & HOJA_PDP_CONVERSION & _
            " para revisar el detalle."

    End If


    ' =========================================================
    ' ACTUALIZAR SOLO LOS DIAS PRESENTES EN EL ARCHIVO
    '
    ' Ago 26:
    '   BJ = L1
    '   BK = L2
    '   BL = L3
    '   BM = L4
    '   BN = Planta
    '
    ' Dia 1  -> fila 3
    ' Dia 17 -> fila 19
    ' Dia 31 -> fila 33
    '
    ' Los dias que NO vienen en el archivo semanal se conservan.
    ' =========================================================
    For Each k In dictDias.Keys

        fDia = CDate(CLng(k))
        filaDestino = Day(fDia) + 2

        For i = 0 To 3

            claveTotal = CStr(k) & "|" & CStr(lineas(i))

            If dictTotales.Exists(claveTotal) Then
                valorLinea = CDbl(dictTotales.Item(claveTotal))
            Else
                valorLinea = 0
            End If

            wsDst.Cells(filaDestino, 62 + i).Value = valorLinea

        Next i

        wsDst.Cells(filaDestino, 66).Formula = _
            "=SUM(BJ" & filaDestino & ":BM" & filaDestino & ")"

    Next k


    ' =========================================================
    ' TOTALES MENSUALES PLAN
    ' =========================================================
    wsDst.Range("BJ35").Formula = "=SUM(BJ3:BJ34)"
    wsDst.Range("BK35").Formula = "=SUM(BK3:BK34)"
    wsDst.Range("BL35").Formula = "=SUM(BL3:BL34)"
    wsDst.Range("BM35").Formula = "=SUM(BM3:BM34)"
    wsDst.Range("BN35").Formula = "=SUM(BJ35:BM35)"


    ' =========================================================
    ' CUMPLIMIENTO REAL / PDP
    ' =========================================================
    For r = 3 To 34

        wsDst.Range("BO" & r).Formula = _
            "=IFERROR(BE" & r & "/BJ" & r & ",0)"

        wsDst.Range("BP" & r).Formula = _
            "=IFERROR(BF" & r & "/BK" & r & ",0)"

        wsDst.Range("BQ" & r).Formula = _
            "=IFERROR(BG" & r & "/BL" & r & ",0)"

        wsDst.Range("BR" & r).Formula = _
            "=IFERROR(BH" & r & "/BM" & r & ",0)"

        wsDst.Range("BS" & r).Formula = _
            "=IFERROR(BI" & r & "/BN" & r & ",0)"

    Next r

    wsDst.Range("BO35").Formula = "=IFERROR(BE35/BJ35,0)"
    wsDst.Range("BP35").Formula = "=IFERROR(BF35/BK35,0)"
    wsDst.Range("BQ35").Formula = "=IFERROR(BG35/BL35,0)"
    wsDst.Range("BR35").Formula = "=IFERROR(BH35/BM35,0)"
    wsDst.Range("BS35").Formula = "=IFERROR(BI35/BN35,0)"


    ' =========================================================
    ' PDP DEL DIA DE CORTE
    ' =========================================================
    wsDst.Range("S6").Formula = "=INDEX(BJ$3:BJ$34,$H$2)"
    wsDst.Range("S7").Formula = "=INDEX(BK$3:BK$34,$H$2)"
    wsDst.Range("S8").Formula = "=INDEX(BL$3:BL$34,$H$2)"
    wsDst.Range("S9").Formula = "=INDEX(BM$3:BM$34,$H$2)"
    wsDst.Range("S10").Formula = "=INDEX(BN$3:BN$34,$H$2)"


    ' =========================================================
    ' FORMATOS
    ' =========================================================
    wsDst.Range("BJ3:BN35").NumberFormat = "#,##0.00"
    wsDst.Range("BO3:BS35").NumberFormat = "0.00%"
    wsDst.Range("S6:S10").NumberFormat = "#,##0.00"


    ' =========================================================
    ' TRAZABILIDAD
    ' =========================================================
    wsRep.Range("X37").Value = "Archivo PDP"
    wsRep.Range(CELDA_META_PDP_ARCHIVO).Value = _
        NombreDeRuta(rutaFuente)

    wsRep.Range("X38").Value = "Fecha archivo PDP"
    wsRep.Range(CELDA_META_PDP_FECHA).Value = _
        FileDateTime(rutaFuente)

    wsRep.Range(CELDA_META_PDP_FECHA).NumberFormat = _
        "dd/mm/yyyy hh:mm"

    wsRep.Range("X39").Value = "Periodo PDP importado"

    If primeraFecha > 0 And ultimaFecha > 0 Then
        wsRep.Range(CELDA_META_PDP_PERIODO).Value = _
            Format$(primeraFecha, "dd/mm/yyyy") & _
            " - " & _
            Format$(ultimaFecha, "dd/mm/yyyy")
    End If

    Exit Sub


EH:

    Err.Raise Err.Number, , _
        "Error procesando PDP '" & _
        NombreDeRuta(rutaFuente) & "': " & _
        Err.Description

End Sub


' =====================================================================
' DETECTAR CABECERA DEL PRODUCTION PLAN
' =====================================================================
Private Function BuscarFilaCabeceraPlanV65( _
                    ByVal ws As Worksheet, _
                    ByRef colLinea As Long, _
                    ByRef colProducto As Long, _
                    ByRef colSKU As Long, _
                    ByRef colTarimas As Long) As Long

    Dim ultFila As Long
    Dim ultCol As Long

    Dim r As Long
    Dim c As Long
    Dim txt As String

    ultFila = UltimaFilaUsada(ws)
    ultCol = UltimaColumnaUsadaV6(ws)

    If ultFila > 25 Then ultFila = 25
    If ultCol > 30 Then ultCol = 30

    For r = 1 To ultFila

        colLinea = 0
        colProducto = 0
        colSKU = 0
        colTarimas = 0

        For c = 1 To ultCol

            txt = NormalizarSAPV6(CStr( _
                    ValorCeldaFusionadaPDPV65( _
                        ws.Cells(r, c))))

            Select Case txt

                Case "LINEA"
                    colLinea = c

                Case "PRODUCTO"
                    colProducto = c

                Case "SKU"
                    colSKU = c

                Case "TARIMAS"
                    colTarimas = c

            End Select

        Next c

        If colLinea > 0 And _
           colProducto > 0 And _
           colSKU > 0 And _
           colTarimas > 0 Then

            BuscarFilaCabeceraPlanV65 = r
            Exit Function

        End If

    Next r

End Function


' =====================================================================
' DEVUELVE EL VALOR VISIBLE DE UNA CELDA AUNQUE ESTE FUSIONADA
' =====================================================================
Private Function ValorCeldaFusionadaPDPV65( _
                    ByVal celda As Range) As Variant

    On Error GoTo NORMAL

    If celda.MergeCells Then
        ValorCeldaFusionadaPDPV65 = _
            celda.MergeArea.Cells(1, 1).Value
        Exit Function
    End If

NORMAL:

    ValorCeldaFusionadaPDPV65 = celda.Value

End Function


' =====================================================================
' CONVERTIR ENCABEZADOS COMO:
'   lunes 17-agosto-2026
'   martes 18-agosto-2026
'   17-08-2026
' =====================================================================
Private Function FechaDesdeEncabezadoPDPV65( _
                    ByVal valor As Variant) As Date

    Dim s As String
    Dim partes() As String
    Dim pos As Long

    Dim d As Long
    Dim m As Long
    Dim y As Long

    On Error GoTo SALIR

    If IsDate(valor) Then
        FechaDesdeEncabezadoPDPV65 = DateValue(CDate(valor))
        Exit Function
    End If

    s = LCase$(Trim$(CStr(valor)))

    If Len(s) = 0 Then Exit Function

    s = Replace(s, "" & ChrW(225) & "", "a")
    s = Replace(s, "" & ChrW(233) & "", "e")
    s = Replace(s, "" & ChrW(237) & "", "i")
    s = Replace(s, "" & ChrW(243) & "", "o")
    s = Replace(s, "" & ChrW(250) & "", "u")

    ' Quitar el nombre del dia de la semana.
    For pos = 1 To Len(s)

        If Mid$(s, pos, 1) Like "#" Then
            s = Mid$(s, pos)
            Exit For
        End If

    Next pos

    s = Replace(s, "/", "-")
    s = Replace(s, ".", "-")
    s = Replace(s, " ", "")

    partes = Split(s, "-")

    If UBound(partes) < 2 Then Exit Function

    d = Val(partes(0))
    m = MesPDPV65(partes(1))
    y = Val(partes(2))

    If y < 100 Then y = 2000 + y

    If d <= 0 Or m <= 0 Or y <= 0 Then Exit Function

    FechaDesdeEncabezadoPDPV65 = DateSerial(y, m, d)

SALIR:

End Function


Private Function MesPDPV65(ByVal texto As String) As Long

    texto = LCase$(Trim$(texto))

    If IsNumeric(texto) Then

        MesPDPV65 = CLng(texto)
        Exit Function

    End If

    Select Case texto

        Case "enero", "ene"
            MesPDPV65 = 1

        Case "febrero", "feb"
            MesPDPV65 = 2

        Case "marzo", "mar"
            MesPDPV65 = 3

        Case "abril", "abr"
            MesPDPV65 = 4

        Case "mayo", "may"
            MesPDPV65 = 5

        Case "junio", "jun"
            MesPDPV65 = 6

        Case "julio", "jul"
            MesPDPV65 = 7

        Case "agosto", "ago"
            MesPDPV65 = 8

        Case "septiembre", "setiembre", "sep", "set"
            MesPDPV65 = 9

        Case "octubre", "oct"
            MesPDPV65 = 10

        Case "noviembre", "nov"
            MesPDPV65 = 11

        Case "diciembre", "dic"
            MesPDPV65 = 12

    End Select

End Function


Private Function NormalizarLineaPDPV65( _
                    ByVal texto As String) As String

    texto = NormalizarSAPV6(texto)

    Select Case texto

        Case "LINEA001", "LINEA01", "LINEA1", "L1"
            NormalizarLineaPDPV65 = "LINEA001"

        Case "LINEA002", "LINEA02", "LINEA2", "L2"
            NormalizarLineaPDPV65 = "LINEA002"

        Case "LINEA003", "LINEA03", "LINEA3", "L3"
            NormalizarLineaPDPV65 = "LINEA003"

        Case "LINEA004", "LINEA04", "LINEA4", "L4"
            NormalizarLineaPDPV65 = "LINEA004"

    End Select

End Function


Private Function NormalizarSKUPDPV65( _
                    ByVal valor As Variant) As String

    Dim s As String

    On Error GoTo TEXTO

    If IsNumeric(valor) Then

        NormalizarSKUPDPV65 = _
            Format$(CDbl(valor), "0")

        Exit Function

    End If

TEXTO:

    s = UCase$(Trim$(CStr(valor)))
    s = Replace(s, " ", "")

    NormalizarSKUPDPV65 = s

End Function


Private Function EsSKUNumericoPDPV65( _
                    ByVal sku As String) As Boolean

    sku = Trim$(sku)

    If Len(sku) = 0 Then Exit Function

    EsSKUNumericoPDPV65 = IsNumeric(sku)

End Function


' =====================================================================
' CATALOGO CF -> CU
'
' Se crea automaticamente la hoja CATALOGO_PDP.
' Si aparece un nuevo SKU solo hay que agregar su Factor CF_CU.
' =====================================================================
Private Function CrearCatalogoFactorCUV65() As Object

    Dim ws As Worksheet
    Dim dic As Object

    Dim ultFila As Long
    Dim r As Long

    Dim sku As String
    Dim factor As Double

    AsegurarCatalogoPDPV65

    Set ws = ThisWorkbook.Worksheets(HOJA_CATALOGO_PDP)
    Set dic = CreateObject("Scripting.Dictionary")

    dic.CompareMode = vbTextCompare

    ultFila = UltimaFilaUsada(ws)

    For r = 2 To ultFila

        sku = NormalizarSKUPDPV65(ws.Cells(r, 1).Value)
        factor = NumeroPDPV65(ws.Cells(r, 5).Value)

        If Len(sku) > 0 And factor > 0 Then

            If dic.Exists(sku) Then
                dic.Item(sku) = factor
            Else
                dic.Add sku, factor
            End If

        End If

    Next r

    Set CrearCatalogoFactorCUV65 = dic

End Function


Private Sub AsegurarCatalogoPDPV65()

    Dim ws As Worksheet
    Dim datos As Variant
    Dim fila As Long
    Dim i As Long
    Dim sku As String
    Dim c As Range

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_CATALOGO_PDP)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets( _
                        ThisWorkbook.Worksheets.Count))

        ws.Name = HOJA_CATALOGO_PDP

    End If

    ws.Range("A1:F1").Value = Array( _
        "ECC SKU", _
        "Descripcion_Est", _
        "Cjas/tar", _
        "Factor CF a CFConv", _
        "Factor CF_CU", _
        "Tipo")

    datos = Array( _
        Array("9001", "PRODUCTO DEMO 01", 64, 0.5, 2.11338, "Refresco"), _
        Array("9002", "PRODUCTO DEMO 02", 36, 0.75, 3.17, "Refresco"), _
        Array("9003", "PRODUCTO DEMO 03", 64, 0.5, 2.11338, "Refresco"), _
        Array("9004", "PRODUCTO DEMO 04", 64, 0.5, 1.937262, "Refresco"), _
        Array("9005", "PRODUCTO DEMO 05", 32, 1, 3.874524, "Refresco"), _
        Array("9006", "PRODUCTO DEMO 06", 75, 1, 2.853058, "Refresco"), _
        Array("9007", "PRODUCTO DEMO 07", 40, 1, 3.170065, "Refresco"), _
        Array("9008", "PRODUCTO DEMO 08", 36, 0.75, 3.164556962, "Refresco"), _
        Array("9009", "PRODUCTO DEMO 09", 40, 0.666667, 2.465606, "Refresco"), _
        Array("9010", "PRODUCTO DEMO 10", 108, 0.5, 1.26806, "Refresco"), _
        Array("9011", "PRODUCTO DEMO 11", 64, 0.5, 2.113432, "Refresco"), _
        Array("9012", "PRODUCTO DEMO 12", 54, 1, 2.53612, "Refresco") _
    )

    For i = LBound(datos) To UBound(datos)

        sku = CStr(datos(i)(0))

        Set c = Nothing

        On Error Resume Next

        Set c = ws.Columns(1).Find( _
                    What:=sku, _
                    LookIn:=xlValues, _
                    LookAt:=xlWhole, _
                    MatchCase:=False)

        On Error GoTo 0

        If c Is Nothing Then

            fila = UltimaFilaUsada(ws) + 1

            ws.Cells(fila, 1).Resize(1, 6).Value = _
                Array( _
                    datos(i)(0), _
                    datos(i)(1), _
                    datos(i)(2), _
                    datos(i)(3), _
                    datos(i)(4), _
                    datos(i)(5))

        End If

    Next i

    ws.Range("A1:F1").Font.Bold = True
    ws.Columns("A:F").AutoFit
    ws.Columns("B").ColumnWidth = 42
    ws.Range("C2:E1000").NumberFormat = "0.000000000"

End Sub


' =====================================================================
' AUDITORIA PDP_CONVERSION
' =====================================================================
Private Function PrepararHojaPDPConversionV65() As Worksheet

    Dim ws As Worksheet

    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_PDP_CONVERSION)
    On Error GoTo 0

    If ws Is Nothing Then

        Set ws = ThisWorkbook.Worksheets.Add( _
                    After:=ThisWorkbook.Worksheets( _
                        ThisWorkbook.Worksheets.Count))

        ws.Name = HOJA_PDP_CONVERSION

    Else

        ws.Cells.Clear

    End If

    ws.Range("A1:I1").Value = Array( _
        "Fecha", _
        "Linea", _
        "Producto", _
        "SKU", _
        "CF Plan", _
        "Factor CF_CU", _
        "CU Convertidas", _
        "Estado", _
        "Archivo")

    ws.Range("A1:I1").Font.Bold = True

    Set PrepararHojaPDPConversionV65 = ws

End Function


Private Sub EscribirDetallePDPConversionV65( _
                    ByVal ws As Worksheet, _
                    ByVal detalle As Collection, _
                    ByVal dictTotales As Object, _
                    ByVal dictDias As Object, _
                    ByVal rutaFuente As String)

    Dim i As Long
    Dim r As Long
    Dim item As Variant

    Dim k As Variant
    Dim lineas As Variant
    Dim linea As Variant

    Dim clave As String
    Dim f As Date
    Dim totalCU As Double

    r = 2

    For i = 1 To detalle.Count

        item = detalle(i)

        ws.Cells(r, 1).Value = item(0)
        ws.Cells(r, 2).Value = item(1)
        ws.Cells(r, 3).Value = item(2)
        ws.Cells(r, 4).Value = item(3)
        ws.Cells(r, 5).Value = item(4)
        ws.Cells(r, 6).Value = item(5)
        ws.Cells(r, 7).Value = item(6)
        ws.Cells(r, 8).Value = item(7)
        ws.Cells(r, 9).Value = item(8)

        r = r + 1

    Next i

    r = r + 2

    ws.Cells(r, 1).Value = "RESUMEN CU POR DIA / LINEA"
    ws.Cells(r, 1).Font.Bold = True

    r = r + 1

    ws.Cells(r, 1).Resize(1, 4).Value = _
        Array("Fecha", "Linea", "CU Total", "Archivo")

    ws.Cells(r, 1).Resize(1, 4).Font.Bold = True

    r = r + 1

    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")

    For Each k In dictDias.Keys

        f = CDate(CLng(k))

        For Each linea In lineas

            clave = CStr(k) & "|" & CStr(linea)

            If dictTotales.Exists(clave) Then
                totalCU = CDbl(dictTotales.Item(clave))
            Else
                totalCU = 0
            End If

            ws.Cells(r, 1).Value = f
            ws.Cells(r, 2).Value = CStr(linea)
            ws.Cells(r, 3).Value = totalCU
            ws.Cells(r, 4).Value = NombreDeRuta(rutaFuente)

            r = r + 1

        Next linea

    Next k

    ws.Columns("A:I").AutoFit
    ws.Columns("C").ColumnWidth = 42
    ws.Columns("I").ColumnWidth = 45

    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("E:G").NumberFormat = "#,##0.000000"
    ws.Columns("C").WrapText = True

End Sub


Private Function NumeroPDPV65(ByVal valor As Variant) As Double

    Dim s As String

    On Error GoTo SALIR

    If IsNumeric(valor) Then

        NumeroPDPV65 = CDbl(valor)
        Exit Function

    End If

    s = Trim$(CStr(valor))

    If Len(s) = 0 Then Exit Function

    s = Replace(s, " ", "")
    s = Replace(s, Chr$(160), "")

    ' En los archivos de Excel normalmente el valor ya es numerico.
    ' Este bloque solo es respaldo si llega como texto.
    If IsNumeric(s) Then

        NumeroPDPV65 = CDbl(s)
        Exit Function

    End If

    If InStr(s, ",") > 0 And InStr(s, ".") = 0 Then

        ' Ejemplo 48,600 exportado como texto con separador de miles.
        If Len(s) - InStrRev(s, ",") = 3 Then
            s = Replace(s, ",", "")
        End If

    End If

    If IsNumeric(s) Then
        NumeroPDPV65 = CDbl(s)
    End If

SALIR:

End Function


' PRODUCCION SAP - ZREP_PRODUCCION /VAR_PROD
' =====================================================================
Private Sub ExportarProduccionAgo26V6(ByVal ses As Object, _
                                      ByVal fechaInicioMes As Date, _
                                      ByVal fechaCorte As Date, _
                                      ByVal rutaCompleta As String)

    Dim carpeta As String
    Dim nombre As String
    Dim fechaFinMes As Date

    carpeta = CarpetaDeRuta(rutaCompleta)
    nombre = NombreDeRuta(rutaCompleta)

    ' El grabador SAP usa del dia 1 al ultimo dia del mes.
    ' Al cargar Ago 26 se filtran los registros posteriores a F3,
    ' evitando incorporar produccion parcial del dia actual.
    fechaFinMes = DateSerial(Year(fechaCorte), Month(fechaCorte) + 1, 0)

    ses.findById("wnd[0]").maximize
    IrATransaccion ses, TRANSACCION_PRODUCCION

    ses.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = CENTRO
    ses.findById("wnd[0]/usr/ctxtS_ARBPL-LOW").Text = CENTRO_TRABAJO_REP2
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-LOW").Text = _
        Format$(fechaInicioMes, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").Text = _
        Format$(fechaFinMes, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtP_VARI").Text = VARIANTE_PRODUCCION

    ses.findById("wnd[0]/tbar[1]/btn[8]").press
    EsperarSAP ses

    ' Lista > Exportar > Archivo local
    ses.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
    EsperarSAP ses

    ses.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/" & _
                 "sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[1,0]").Select
    ses.findById("wnd[1]/tbar[0]/btn[0]").press
    EsperarSAP ses

    PonerTextoSiExiste ses, "wnd[1]/usr/ctxtDY_PATH", carpeta
    ses.findById("wnd[1]/usr/ctxtDY_FILENAME").Text = nombre

    ' En el recorder esta ventana se confirma con Enter.
    On Error Resume Next
    ses.findById("wnd[1]/tbar[0]/btn[11]").press
    If Err.Number <> 0 Then
        Err.Clear
        ses.findById("wnd[1]").sendVKey 0
    End If
    On Error GoTo 0

    EsperarSAP ses

End Sub


Private Sub ProcesarProduccionAgo26V6(ByVal wsRaw As Worksheet, _
                                      ByVal fechaCorte As Date)

    Dim wsPDP As Worksheet

    Dim filaEncabezado As Long
    Dim colLinea As Long
    Dim colFecha As Long
    Dim colCU As Long

    Dim ultimaFila As Long
    Dim r As Long
    Dim d As Long
    Dim i As Long
    Dim idx As Long
    Dim filaDestino As Long

    Dim linea As String
    Dim f As Date
    Dim cu As Double

    Dim realDia(1 To 31, 0 To 3) As Double
    Dim totalMes(0 To 3) As Double

    Dim diasMes As Long
    Dim diaCorte As Long

    Dim columnasReal As Variant
    Dim formulaCol As String

    filaEncabezado = BuscarFilaEncabezadoProduccionV6(wsRaw)

    If filaEncabezado = 0 Then
        Err.Raise vbObjectError + 2650, , _
            "No se encontro la fila de encabezados de PRODUCCION.XLS."
    End If

    colLinea = BuscarColumnaProduccionV6( _
                    wsRaw, filaEncabezado, _
                    Array("Linea", "L" & ChrW(237) & "nea"))

    colFecha = BuscarColumnaProduccionV6( _
                    wsRaw, filaEncabezado, _
                    Array("Fecha de Operacion", _
                          "Fecha de Operaci" & ChrW(243) & "n"))

    colCU = BuscarColumnaProduccionV6( _
                    wsRaw, filaEncabezado, _
                    Array("Cajas Unitarias", _
                          "Cajas Unitarias Producidas"))

    If colLinea = 0 Or colFecha = 0 Or colCU = 0 Then
        Err.Raise vbObjectError + 2651, , _
            "El layout de ZREP_PRODUCCION debe contener: " & _
            "Linea, Fecha de Operacion y Cajas Unitarias."
    End If

    ultimaFila = UltimaFilaUsada(wsRaw)

    For r = filaEncabezado + 1 To ultimaFila

        linea = Trim$(CStr(wsRaw.Cells(r, colLinea).Value))
        idx = IndiceLinea(linea)

        If idx >= 0 Then

            f = FechaDesdeValorV6(wsRaw.Cells(r, colFecha).Value)

            If f > 0 Then

                If Year(f) = Year(fechaCorte) And _
                   Month(f) = Month(fechaCorte) And _
                   CLng(f) <= CLng(fechaCorte) Then

                    If NumeroProduccionSeguroV6( _
                           wsRaw.Cells(r, colCU).Value, cu) Then

                        d = Day(f)

                        If d >= 1 And d <= 31 Then
                            realDia(d, idx) = realDia(d, idx) + cu
                            totalMes(idx) = totalMes(idx) + cu
                        End If

                    End If

                End If

            End If

        End If

    Next r


    Set wsPDP = ThisWorkbook.Worksheets(HOJA_PDP)

    diasMes = Day(DateSerial(Year(fechaCorte), Month(fechaCorte) + 1, 0))
    diaCorte = Day(fechaCorte)

    columnasReal = Array("BE", "BF", "BG", "BH")


    ' =========================================================
    ' METADATOS DEL CORTE
    ' =========================================================
    wsPDP.Range("H2").Value = diaCorte
    wsPDP.Range("M4").Value = diaCorte
    wsPDP.Range("M2").Value = diasMes
    wsPDP.Range("M3").Value = diasMes
    wsPDP.Range("P4").Value = _
        DateSerial(Year(fechaCorte), Month(fechaCorte) + 1, 0)
    wsPDP.Range("P4").NumberFormat = "dd-mmm"


    ' =========================================================
    ' REAL DIARIO
    '
    ' BD = dia
    ' BE = L1
    ' BF = L2
    ' BG = L3
    ' BH = L4
    ' BI = Planta
    ' =========================================================
    wsPDP.Range("BE3:BH34").ClearContents

    For d = 1 To 31

        filaDestino = d + 2
        wsPDP.Range("BD" & filaDestino).Value = d

        For i = 0 To 3

            If d <= diasMes And d <= diaCorte Then
                wsPDP.Range(columnasReal(i) & filaDestino).Value = _
                    realDia(d, i)
            Else
                wsPDP.Range(columnasReal(i) & filaDestino).Value = 0
            End If

        Next i

        wsPDP.Range("BI" & filaDestino).Formula = _
            "=SUM(BE" & filaDestino & ":BH" & filaDestino & ")"

    Next d


    ' =========================================================
    ' TOTALES MTD
    ' =========================================================
    wsPDP.Range("BE35").Formula = "=SUM(BE3:BE34)"
    wsPDP.Range("BF35").Formula = "=SUM(BF3:BF34)"
    wsPDP.Range("BG35").Formula = "=SUM(BG3:BG34)"
    wsPDP.Range("BH35").Formula = "=SUM(BH3:BH34)"
    wsPDP.Range("BI35").Formula = "=SUM(BE35:BH35)"


    ' =========================================================
    ' RESUMEN SUPERIOR
    '
    ' I = Produccion Real MTD
    ' R = Produccion real dia
    '
    ' Se dejan como formulas hacia la tabla REAL SAP para que
    ' el resumen y cumplimiento PDP queden ligados a una sola
    ' fuente de verdad.
    ' =========================================================
    wsPDP.Range("I6").Formula = "=BE35"
    wsPDP.Range("I7").Formula = "=BF35"
    wsPDP.Range("I8").Formula = "=BG35"
    wsPDP.Range("I9").Formula = "=BH35"
    wsPDP.Range("I10").Formula = "=BI35"

    wsPDP.Range("R6").Formula = _
        "=INDEX(BE$3:BE$34,$H$2)"
    wsPDP.Range("R7").Formula = _
        "=INDEX(BF$3:BF$34,$H$2)"
    wsPDP.Range("R8").Formula = _
        "=INDEX(BG$3:BG$34,$H$2)"
    wsPDP.Range("R9").Formula = _
        "=INDEX(BH$3:BH$34,$H$2)"
    wsPDP.Range("R10").Formula = _
        "=INDEX(BI$3:BI$34,$H$2)"


    ' =========================================================
    ' EFICIENCIAS YA TRAIDAS POR EL PASO REP2
    '
    ' J = Eficiencia MTD
    ' Q = Eficiencia dia
    '
    ' Esto integra Ago 26 con el mismo proceso SAP completo.
    ' =========================================================
    wsPDP.Range("J6").Formula = "='REP2'!G16"
    wsPDP.Range("J7").Formula = "='REP2'!G17"
    wsPDP.Range("J8").Formula = "='REP2'!G18"
    wsPDP.Range("J9").Formula = "='REP2'!G19"
    wsPDP.Range("J10").Formula = "='REP2'!G21"

    wsPDP.Range("Q6").Formula = "='REP2'!G3"
    wsPDP.Range("Q7").Formula = "='REP2'!G4"
    wsPDP.Range("Q8").Formula = "='REP2'!G5"
    wsPDP.Range("Q9").Formula = "='REP2'!G6"
    wsPDP.Range("Q10").Formula = "='REP2'!G8"


    ' =========================================================
    ' FORMATOS
    ' =========================================================
    wsPDP.Range("BE3:BI35").NumberFormat = "#,##0"
    wsPDP.Range("I6:I10").NumberFormat = "#,##0"
    wsPDP.Range("R6:R10").NumberFormat = "#,##0"
    wsPDP.Range("J6:J10").NumberFormat = "0.00"
    wsPDP.Range("Q6:Q10").NumberFormat = "0.00"

End Sub


Private Function BuscarFilaEncabezadoProduccionV6( _
                    ByVal ws As Worksheet) As Long

    Dim ultFila As Long
    Dim ultCol As Long
    Dim r As Long
    Dim c As Long
    Dim encontrados As Long
    Dim texto As String

    ultFila = UltimaFilaUsada(ws)
    ultCol = UltimaColumnaUsadaV6(ws)

    If ultFila > 30 Then ultFila = 30

    For r = 1 To ultFila

        encontrados = 0

        For c = 1 To ultCol

            texto = NormalizarEncabezadoProduccionV6( _
                        CStr(ws.Cells(r, c).Value))

            If texto = "LINEA" Then encontrados = encontrados + 1
            If texto = "FECHADEOPERACION" Then encontrados = encontrados + 1
            If texto = "CAJASUNITARIAS" Then encontrados = encontrados + 1

        Next c

        If encontrados >= 3 Then
            BuscarFilaEncabezadoProduccionV6 = r
            Exit Function
        End If

    Next r

End Function


Private Function BuscarColumnaProduccionV6( _
                    ByVal ws As Worksheet, _
                    ByVal fila As Long, _
                    ByVal nombres As Variant) As Long

    Dim ultCol As Long
    Dim c As Long
    Dim i As Long
    Dim actual As String
    Dim esperado As String

    ultCol = UltimaColumnaUsadaV6(ws)

    For c = 1 To ultCol

        actual = NormalizarEncabezadoProduccionV6( _
                    CStr(ws.Cells(fila, c).Value))

        For i = LBound(nombres) To UBound(nombres)

            esperado = NormalizarEncabezadoProduccionV6( _
                        CStr(nombres(i)))

            If actual = esperado Then
                BuscarColumnaProduccionV6 = c
                Exit Function
            End If

        Next i

    Next c

End Function


Private Function NormalizarEncabezadoProduccionV6( _
                    ByVal texto As String) As String

    texto = UCase$(Trim$(texto))

    texto = Replace(texto, "" & ChrW(193) & "", "A")
    texto = Replace(texto, "" & ChrW(201) & "", "E")
    texto = Replace(texto, "" & ChrW(205) & "", "I")
    texto = Replace(texto, "" & ChrW(211) & "", "O")
    texto = Replace(texto, "" & ChrW(218) & "", "U")
    texto = Replace(texto, "" & ChrW(220) & "", "U")

    texto = Replace(texto, " ", "")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, "_", "")
    texto = Replace(texto, "-", "")

    NormalizarEncabezadoProduccionV6 = texto

End Function


Private Function NumeroProduccionSeguroV6( _
                    ByVal valor As Variant, _
                    ByRef resultado As Double) As Boolean

    Dim s As String
    Dim posComa As Long
    Dim posPunto As Long
    Dim decimalSAP As String
    Dim milesSAP As String

    On Error GoTo NO_ES

    If IsNumeric(valor) Then
        resultado = CDbl(valor)
        NumeroProduccionSeguroV6 = True
        Exit Function
    End If

    s = Trim$(CStr(valor))

    If Len(s) = 0 Then Exit Function

    s = Replace(s, " ", "")

    posComa = InStrRev(s, ",")
    posPunto = InStrRev(s, ".")

    If posComa > 0 And posPunto > 0 Then

        ' El separador que aparece mas a la derecha es el decimal.
        If posPunto > posComa Then
            decimalSAP = "."
            milesSAP = ","
        Else
            decimalSAP = ","
            milesSAP = "."
        End If

        s = Replace(s, milesSAP, "")

        If Application.DecimalSeparator <> decimalSAP Then
            s = Replace(s, decimalSAP, Application.DecimalSeparator)
        End If

    ElseIf posPunto > 0 Then

        If Application.DecimalSeparator <> "." Then
            s = Replace(s, ".", Application.DecimalSeparator)
        End If

    ElseIf posComa > 0 Then

        If Application.DecimalSeparator <> "," Then
            s = Replace(s, ",", Application.DecimalSeparator)
        End If

    End If

    If IsNumeric(s) Then
        resultado = CDbl(s)
        NumeroProduccionSeguroV6 = True
    End If

    Exit Function

NO_ES:
    NumeroProduccionSeguroV6 = False

End Function


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
' CONTROL DEL MODO RESUMEN / DETALLE DE ZREP_EFICIENCIA
' =====================================================================
Private Sub ForzarModoResumenREP2V6(ByVal ses As Object)
    ' ==========================================================
    ' MODO RESUMEN / ACUMULADO
    ' Confirmado directamente en tu SAP el 11/08/2026:
    '
    '   ID:   wnd[0]/usr/radP_ACU
    '   Name: P_ACU
    '   Text: Acumulado
    '
    ' Ya no hacemos deteccion automatica.
    ' ==========================================================
    On Error GoTo EH

    ses.findById("wnd[0]/usr/radP_ACU").SetFocus
    ses.findById("wnd[0]/usr/radP_ACU").Select
    DoEvents

    ' Confirmacion adicional para evitar exportar DETALLE por error.
    If ses.findById("wnd[0]/usr/radP_DET").Selected Then
        Err.Raise vbObjectError + 2442, , _
            "No fue posible activar P_ACU. SAP mantiene P_DET seleccionado."
    End If

    Exit Sub

EH:
    Err.Raise vbObjectError + 2440, , _
        "No se pudo seleccionar el modo ACUMULADO (P_ACU) de ZREP_EFICIENCIA. " & _
        "Error SAP/VBA: " & Err.Description
End Sub

Private Function EsRadioSAPV6(ByVal obj As Object) As Boolean
    Dim t As String
    On Error Resume Next
    t = UCase$(CStr(obj.Type))
    If Len(t) = 0 Then t = UCase$(TypeName(obj))
    On Error GoTo 0

    EsRadioSAPV6 = (InStr(1, t, "RADIO", vbTextCompare) > 0)
End Function

Private Function PropiedadSAPV6(ByVal obj As Object, ByVal propiedad As String) As String
    On Error Resume Next
    Select Case UCase$(propiedad)
        Case "ID"
            PropiedadSAPV6 = CStr(obj.Id)
        Case "NAME"
            PropiedadSAPV6 = CStr(obj.Name)
        Case "TEXT"
            PropiedadSAPV6 = CStr(obj.Text)
        Case "TYPE"
            PropiedadSAPV6 = CStr(obj.Type)
    End Select
    On Error GoTo 0
End Function

Private Function BuscarTextoParcialV6(ByVal ws As Worksheet, ByVal texto As String) As Range
    Dim rng As Range

    Set rng = ws.UsedRange
    On Error Resume Next
    Set BuscarTextoParcialV6 = rng.Find(What:=texto, _
                                         After:=rng.Cells(rng.Cells.Count), _
                                         LookIn:=xlValues, _
                                         LookAt:=xlPart, _
                                         SearchOrder:=xlByRows, _
                                         SearchDirection:=xlNext, _
                                         MatchCase:=False)
    On Error GoTo 0
End Function

Public Sub V6_DiagnosticarModosREP2()
    Dim ses As Object
    Dim det As Object
    Dim cont As Object
    Dim obj As Object
    Dim i As Integer
    Dim salida As String
    Dim seleccionado As String

    On Error GoTo EH

    Set ses = ObtenerSesionSAP()
    If ses Is Nothing Then
        MsgBox "No se encontro una sesion activa de SAP GUI.", vbExclamation, "DIAGNOSTICO REP2"
        Exit Sub
    End If

    IrATransaccion ses, TRANSACCION_REP2
    EsperarSAP ses

    On Error Resume Next
    Set det = ses.findById("wnd[0]/usr/radP_DET")
    On Error GoTo EH

    If det Is Nothing Then
        MsgBox "No se encontro wnd[0]/usr/radP_DET en ZREP_EFICIENCIA.", _
               vbExclamation, "DIAGNOSTICO REP2"
        Exit Sub
    End If

    Set cont = det.Parent

    salida = "RADIO BUTTONS DETECTADOS EN ZREP_EFICIENCIA" & vbCrLf & _
             String$(55, "-") & vbCrLf

    For i = 0 To CInt(cont.Children.Count) - 1
        Set obj = Nothing
        On Error Resume Next
        Set obj = cont.Children(CInt(i))
        On Error GoTo EH

        If Not obj Is Nothing Then
            If EsRadioSAPV6(obj) Then
                seleccionado = ""
                On Error Resume Next
                If obj.Selected Then seleccionado = "  [SELECCIONADO]"
                On Error GoTo EH

                salida = salida & vbCrLf & _
                         "ID: " & PropiedadSAPV6(obj, "Id") & vbCrLf & _
                         "Name: " & PropiedadSAPV6(obj, "Name") & vbCrLf & _
                         "Text: " & PropiedadSAPV6(obj, "Text") & seleccionado & vbCrLf
            End If
        End If
    Next i

    MsgBox salida, vbInformation, "DIAGNOSTICO MODOS REP2"
    Exit Sub

EH:
    MsgBox "Error de diagnostico: " & Err.Number & vbCrLf & Err.Description, _
           vbExclamation, "DIAGNOSTICO REP2"
End Sub

Private Sub ExportarREP2AcumuladoV6(ByVal ses As Object, _
                                 ByVal fIni As Date, _
                                 ByVal fFin As Date, _
                                 ByVal rutaCompleta As String)
    Dim carpeta As String
    Dim nombre As String

    carpeta = CarpetaDeRuta(rutaCompleta)
    nombre = NombreDeRuta(rutaCompleta)

    ses.findById("wnd[0]").maximize
    IrATransaccion ses, TRANSACCION_REP2

    ses.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = CENTRO
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-LOW").Text = Format$(fIni, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").Text = Format$(fFin, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_ARBPL-LOW").Text = CENTRO_TRABAJO_REP2
    ses.findById("wnd[0]/usr/ctxtP_VARI").Text = VARIANTE_REP2

    ' IMPORTANTE:
    ' SAP puede recordar que P_DET quedo seleccionado de una ejecucion anterior.
    ' P_ACU devuelve una sola tabla para el periodo indicado.
    ForzarModoResumenREP2V6 ses

    ses.findById("wnd[0]/tbar[1]/btn[8]").press
    EsperarSAP ses

    ' P_ACU fue confirmado directamente en este sistema SAP.
    ' No hacemos detecciones adicionales de modo: el parser validara el archivo exportado.

    ' Lista > Exportar > Archivo local. Es el flujo de tu grabacion SAP.
    ses.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
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


Private Sub ExportarREP2DetalleV6(ByVal ses As Object, _
                                  ByVal fIni As Date, _
                                  ByVal fFin As Date, _
                                  ByVal rutaCompleta As String)
    Dim carpeta As String
    Dim nombre As String

    carpeta = CarpetaDeRuta(rutaCompleta)
    nombre = NombreDeRuta(rutaCompleta)

    ses.findById("wnd[0]").maximize
    IrATransaccion ses, TRANSACCION_REP2

    ses.findById("wnd[0]/usr/ctxtS_WERKS-LOW").Text = CENTRO
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-LOW").Text = Format$(fIni, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_GSTRI-HIGH").Text = Format$(fFin, "dd.mm.yyyy")
    ses.findById("wnd[0]/usr/ctxtS_ARBPL-LOW").Text = CENTRO_TRABAJO_REP2
    ses.findById("wnd[0]/usr/ctxtP_VARI").Text = VARIANTE_REP2

    ' El detalle es indispensable para disponer de GSTRI y POR_EACU.
' Aqui SI forzamos P_DET de forma explicita.
    ' MODO DETALLE confirmado en SAP: P_DET
    ses.findById("wnd[0]/usr/radP_DET").SetFocus
    ses.findById("wnd[0]/usr/radP_DET").Select
    ses.findById("wnd[0]/tbar[1]/btn[8]").press
    EsperarSAP ses

    AplicarLayoutREP2DetalleV6 ses

    ' Exportar archivo local.
    ses.findById("wnd[0]/mbar/menu[0]/menu[3]/menu[2]").Select
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

Private Sub AplicarLayoutREP2DetalleV6(ByVal ses As Object)
    Dim appSAP As Object
    Dim grid As Object
    Dim gridLayout As Object
    Dim coll1 As Object
    Dim filaFecha As Long
    Dim filaEficiencia As Long
    Dim idGridLayout As String

    idGridLayout = _
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/" & _
        "ssubSUB_DYN0510:SAPLSKBH:0620/" & _
        "cntlCONTAINER1_LAYO/shellcont/shell"

    Set appSAP = ses.Parent.Parent
    Set grid = ses.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")

    ' ==========================================================
    ' CORRECCION V6.1
    ' Ya NO se usan filas fijas 0 y 15.
    ' Al mover Fecha de inicio Real, SAP reordena la lista.
    ' Por eso cada campo se busca por nombre y el grid se vuelve
    ' a leer despues del primer movimiento.
    ' ==========================================================
    ses.findById("wnd[0]/tbar[1]/btn[32]").press
    EsperarSAP ses

    Set gridLayout = ses.findById(idGridLayout)

    ' ----------------------------------------------------------
    ' 1) FECHA DE INICIO REAL / GSTRI
    ' ----------------------------------------------------------
    filaFecha = BuscarFilaCampoLayoutSAPV6( _
                    gridLayout, _
                    "Fecha de inicio Real", _
                    "GSTRI")

    If filaFecha < 0 Then
        Err.Raise vbObjectError + 2701, , _
            "No se encontro el campo 'Fecha de inicio Real' (GSTRI) " & _
            "en la seleccion de columnas SAP."
    End If

    SeleccionarFilaGridSAPV6 gridLayout, filaFecha

    ses.findById( _
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/" & _
        "ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
    EsperarSAP ses

    ' SAP modifica la lista al mover GSTRI: volver a obtener el grid.
    Set gridLayout = Nothing
    Set gridLayout = ses.findById(idGridLayout)

    ' ----------------------------------------------------------
    ' 2) % EFICIENCIA ACUMULADA / POR_EACU
    ' ----------------------------------------------------------
    filaEficiencia = BuscarFilaCampoLayoutSAPV6( _
                        gridLayout, _
                        "Eficiencia Acumulada", _
                        "POR_EACU")

    If filaEficiencia < 0 Then
        Err.Raise vbObjectError + 2702, , _
            "No se encontro el campo '% Eficiencia Acumulada' (POR_EACU) " & _
            "en la seleccion de columnas SAP."
    End If

    SeleccionarFilaGridSAPV6 gridLayout, filaEficiencia

    ses.findById( _
        "wnd[1]/usr/tabsG_TS_ALV/tabpALV_M_R1/" & _
        "ssubSUB_DYN0510:SAPLSKBH:0620/btnAPP_WL_SING").press
    EsperarSAP ses

    ' Confirmar layout
    ses.findById("wnd[1]/tbar[0]/btn[0]").press
    EsperarSAP ses

    ' Volver a tomar el grid principal tras aplicar el layout.
    Set grid = Nothing
    Set grid = ses.findById("wnd[0]/usr/cntlGRID1/shellcont/shell/shellcont[1]/shell")

    ' Orden final requerido por el procesamiento de Excel.
    Set coll1 = appSAP.createGuiCollection
    CargarOrdenColumnasREP2V6 coll1
    grid.columnOrder = coll1
    Set coll1 = Nothing

    ' Dejar Fecha de inicio Real visible. No es necesario SelectColumn.
    grid.setCurrentCell -1, "GSTRI"
    grid.firstVisibleColumn = "GSTRI"
End Sub

Private Function BuscarFilaCampoLayoutSAPV6( _
                    ByVal gridLayout As Object, _
                    ByVal descripcion As String, _
                    ByVal nombreTecnico As String) As Long

    Dim columnas As Object
    Dim fila As Long
    Dim col As Long
    Dim idColumna As String
    Dim contenido As String
    Dim contenidoNormal As String
    Dim descripcionNormal As String
    Dim tecnicoNormal As String

    BuscarFilaCampoLayoutSAPV6 = -1

    descripcionNormal = NormalizarSAPV6(descripcion)
    tecnicoNormal = NormalizarSAPV6(nombreTecnico)

    On Error GoTo ERROR_BUSQUEDA
    Set columnas = gridLayout.ColumnOrder

    For fila = 0 To CLng(gridLayout.RowCount) - 1
        For col = 0 To CLng(columnas.Count) - 1

            idColumna = CStr(columnas(CInt(col)))
            contenido = vbNullString

            On Error Resume Next
            contenido = CStr(gridLayout.GetCellValue(CLng(fila), idColumna))
            On Error GoTo ERROR_BUSQUEDA

            If Len(Trim$(contenido)) > 0 Then
                contenidoNormal = NormalizarSAPV6(contenido)

                If Len(descripcionNormal) > 0 Then
                    If InStr(1, contenidoNormal, descripcionNormal, vbTextCompare) > 0 Then
                        BuscarFilaCampoLayoutSAPV6 = fila
                        Exit Function
                    End If
                End If

                If Len(tecnicoNormal) > 0 Then
                    If InStr(1, contenidoNormal, tecnicoNormal, vbTextCompare) > 0 Then
                        BuscarFilaCampoLayoutSAPV6 = fila
                        Exit Function
                    End If
                End If
            End If
        Next col
    Next fila

    Exit Function

ERROR_BUSQUEDA:
    BuscarFilaCampoLayoutSAPV6 = -1
End Function

Private Sub SeleccionarFilaGridSAPV6( _
                ByVal grid As Object, _
                ByVal fila As Long)

    Dim primeraVisible As Long

    If fila < 0 Then Exit Sub

    If fila > 4 Then
        primeraVisible = fila - 4
    Else
        primeraVisible = 0
    End If

    On Error GoTo ERROR_SELECCION

    grid.firstVisibleRow = primeraVisible
    grid.currentCellRow = fila
    grid.selectedRows = CStr(fila)

    Exit Sub

ERROR_SELECCION:
    Err.Raise vbObjectError + 2705, , _
        "No se pudo seleccionar la fila " & fila & _
        " en el selector de columnas SAP. Error: " & Err.Description
End Sub

Private Sub CargarOrdenColumnasREP2V6(ByVal coll1 As Object)
    coll1.Add "WERKS"
    coll1.Add "GSTRI"
    coll1.Add "ARBPL"
    coll1.Add "ZPRODT1"
    coll1.Add "ZPRODT2"
    coll1.Add "ZPRODTT"
    coll1.Add "ZPEFIT1"
    coll1.Add "ZPEFIT2"
    coll1.Add "ZPEFITT"
    coll1.Add "ZPETF02"
    coll1.Add "ZPETF05"
    coll1.Add "ZPETF06"
    coll1.Add "ZPETF07"
    coll1.Add "ZPETF08"
    coll1.Add "UTILIN"
    coll1.Add "POR_EACU"
End Sub

Private Sub ProcesarREP2DetalleV6(ByVal wsRaw As Worksheet, _
                                  ByVal fechaCorte As Date, _
                                  ByRef fechaInicioReal As Date, _
                                  ByRef eficienciaAcumulada As Double)
    Dim filaCab As Long
    Dim colLinea As Long
    Dim colFecha As Long
    Dim colAcum As Long
    Dim ultFila As Long
    Dim r As Long
    Dim txtLinea As String
    Dim f As Date
    Dim valorAcum As Double
    Dim i As Integer
    Dim lineas As Variant
    Dim primeraFecha(0 To 3) As Date
    Dim wsRep2 As Worksheet
    Dim wsReporte As Worksheet

    filaCab = BuscarFilaCabeceraV6(wsRaw, "Fecha de inicio Real")
    If filaCab = 0 Then
        Err.Raise vbObjectError + 2430, , _
            "El REP2 detalle no contiene la columna 'Fecha de inicio Real'."
    End If

    colLinea = BuscarColumnaCabeceraV6(wsRaw, filaCab, "L" & ChrW(237) & "nea")
    colFecha = BuscarColumnaCabeceraV6(wsRaw, filaCab, "Fecha de inicio Real")
    colAcum = BuscarColumnaCabeceraV6(wsRaw, filaCab, "% Eficiencia Acumulada")

    If colLinea = 0 Or colFecha = 0 Or colAcum = 0 Then
        Err.Raise vbObjectError + 2431, , _
            "Faltan columnas requeridas en REP2 detalle: L" & ChrW(237) & "nea, Fecha de inicio Real o % Eficiencia Acumulada."
    End If

    ultFila = UltimaFilaUsada(wsRaw)
    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")

    For r = filaCab + 1 To ultFila
        txtLinea = Trim$(CStr(wsRaw.Cells(r, colLinea).Value))
        f = FechaDesdeValorV6(wsRaw.Cells(r, colFecha).Value)

        If f > 0 Then
            If fechaInicioReal = 0 Or f < fechaInicioReal Then fechaInicioReal = f

            For i = 0 To 3
                If UCase$(txtLinea) = lineas(i) Then
                    If primeraFecha(i) = 0 Or f < primeraFecha(i) Then primeraFecha(i) = f
                    Exit For
                End If
            Next i
        End If

        If EsNumeroFlexibleV6(wsRaw.Cells(r, colAcum).Value, valorAcum) Then
            ' SAP imprime POR_EACU secuencialmente; el ultimo valor numerico
            ' es el acumulado oficial del periodo solicitado.
            eficienciaAcumulada = valorAcum
        End If
    Next r

    If fechaInicioReal = 0 Then
        Err.Raise vbObjectError + 2432, , "No se encontro una Fecha de inicio Real valida."
    End If

    Set wsRep2 = ThisWorkbook.Worksheets(HOJA_REP2)

    ' Ampliamos REP2 sin alterar A:N, que sigue alimentando el reporte original.
    wsRep2.Range("O2").Value = "Fecha inicio Real"
    wsRep2.Range("P2").Value = "% Eficiencia Acumulada"
    wsRep2.Range("O15").Value = "Fecha inicio Real"
    wsRep2.Range("P15").Value = "% Eficiencia Acumulada"

    ' DIA: la fecha real del corte para las lineas presentes.
    For i = 0 To 3
        If ExisteLineaEnFechaV6(wsRaw, filaCab + 1, ultFila, colLinea, colFecha, lineas(i), fechaCorte) Then
            wsRep2.Cells(3 + i, 15).Value = fechaCorte
        Else
            wsRep2.Cells(3 + i, 15).ClearContents
        End If
        wsRep2.Cells(3 + i, 16).ClearContents
    Next i

    wsRep2.Range("O8").Value = fechaCorte
    wsRep2.Range("P8").Value = eficienciaAcumulada

    ' MES: primera fecha real encontrada por linea.
    For i = 0 To 3
        If primeraFecha(i) > 0 Then
            wsRep2.Cells(16 + i, 15).Value = primeraFecha(i)
        Else
            wsRep2.Cells(16 + i, 15).ClearContents
        End If
        wsRep2.Cells(16 + i, 16).ClearContents
    Next i

    wsRep2.Range("O21").Value = fechaInicioReal
    wsRep2.Range("P21").Value = eficienciaAcumulada
    wsRep2.Range("O3:O21").NumberFormat = "dd/mm/yyyy"
    wsRep2.Range("P3:P21").NumberFormat = "0.00"

    ' Metadatos para el dashboard PRO. Se colocan fuera de la tabla visible.
    Set wsReporte = ThisWorkbook.Worksheets(HOJA_REPORTE)
    wsReporte.Range("X35").Value = "Fecha inicio SAP"
    wsReporte.Range("Y35").Value = "Eficiencia acumulada SAP"
    wsReporte.Range(CELDA_META_FECHA).Value = fechaInicioReal
    wsReporte.Range(CELDA_META_EFI_ACUM).Value = eficienciaAcumulada
    wsReporte.Range(CELDA_META_FECHA).NumberFormat = "dd/mm/yyyy"
    wsReporte.Range(CELDA_META_EFI_ACUM).NumberFormat = "0.00"

    GuardarREP2DetalleV6 wsRaw
End Sub

Private Sub GuardarREP2DetalleV6(ByVal wsRaw As Worksheet)
    Dim wsDst As Worksheet
    Dim rng As Range

    On Error Resume Next
    Set wsDst = ThisWorkbook.Worksheets(HOJA_REP2_DETALLE)
    On Error GoTo 0

    If wsDst Is Nothing Then
        Set wsDst = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        wsDst.Name = HOJA_REP2_DETALLE
    Else
        wsDst.Cells.Clear
    End If

    Set rng = wsRaw.UsedRange
    wsDst.Range("A1").Resize(rng.Rows.Count, rng.Columns.Count).Value = rng.Value
    wsDst.Columns.AutoFit
End Sub

Private Function BuscarFilaCabeceraV6(ByVal ws As Worksheet, ByVal texto As String) As Long
    Dim c As Range
    Set c = ws.UsedRange.Find(What:=texto, After:=ws.UsedRange.Cells(ws.UsedRange.Cells.Count), _
                              LookIn:=xlValues, LookAt:=xlWhole, _
                              SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                              MatchCase:=False)
    If Not c Is Nothing Then BuscarFilaCabeceraV6 = c.Row
End Function

Private Function BuscarColumnaCabeceraV6(ByVal ws As Worksheet, _
                                         ByVal fila As Long, _
                                         ByVal texto As String) As Long
    Dim c As Range
    Dim ultCol As Long

    ultCol = UltimaColumnaUsadaV6(ws)
    Set c = ws.Range(ws.Cells(fila, 1), ws.Cells(fila, ultCol)).Find( _
            What:=texto, After:=ws.Cells(fila, ultCol), LookIn:=xlValues, _
            LookAt:=xlWhole, SearchOrder:=xlByColumns, SearchDirection:=xlNext, _
            MatchCase:=False)
    If Not c Is Nothing Then BuscarColumnaCabeceraV6 = c.Column
End Function

Private Function FechaDesdeValorV6(ByVal v As Variant) As Date
    Dim s As String
    Dim p() As String

    On Error GoTo SALIR
    If IsDate(v) Then
        FechaDesdeValorV6 = CDate(v)
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    p = Split(s, ".")
    If UBound(p) = 2 Then
        FechaDesdeValorV6 = DateSerial(CInt(p(2)), CInt(p(1)), CInt(p(0)))
    End If
SALIR:
End Function

Private Function EsNumeroFlexibleV6(ByVal v As Variant, ByRef resultado As Double) As Boolean
    Dim s As String
    On Error GoTo NO_ES

    If IsNumeric(v) Then
        resultado = CDbl(v)
        EsNumeroFlexibleV6 = True
        Exit Function
    End If

    s = Trim$(CStr(v))
    If Len(s) = 0 Then Exit Function
    s = Replace(s, "%", "")
    s = Replace(s, " ", "")

    If Application.DecimalSeparator = "," Then
        s = Replace(s, ".", ",")
    Else
        s = Replace(s, ",", ".")
    End If

    If IsNumeric(s) Then
        resultado = CDbl(s)
        EsNumeroFlexibleV6 = True
    End If
    Exit Function
NO_ES:
    EsNumeroFlexibleV6 = False
End Function

Private Function ExisteLineaEnFechaV6(ByVal ws As Worksheet, _
                                      ByVal filaIni As Long, _
                                      ByVal filaFin As Long, _
                                      ByVal colLinea As Long, _
                                      ByVal colFecha As Long, _
                                      ByVal linea As String, _
                                      ByVal fechaBuscada As Date) As Boolean
    Dim r As Long
    Dim f As Date

    For r = filaIni To filaFin
        If UCase$(Trim$(CStr(ws.Cells(r, colLinea).Value))) = UCase$(linea) Then
            f = FechaDesdeValorV6(ws.Cells(r, colFecha).Value)
            If f > 0 Then
                If CLng(f) = CLng(fechaBuscada) Then
                    ExisteLineaEnFechaV6 = True
                    Exit Function
                End If
            End If
        End If
    Next r
End Function


' =====================================================================
' P_ACU - PARSER REAL OBSERVADO EN EL VIDEO
'
' P_ACU NO genera bloques DIA/MES.
' Genera UNA tabla para el periodo solicitado:
'
'   LINEA001
'   LINEA002
'   LINEA003
'   LINEA004
'   ...
'   Tot. | Acum.
'
' Por eso V6 ejecuta P_ACU dos veces:
'   1) fecha corte -> fecha corte = DIA
'   2) dia 1 -> fecha corte       = MES
' =====================================================================
Private Sub CargarREP2AcumuladoV6(ByVal wsRaw As Worksheet, _
                                  ByVal filaBase As Long, _
                                  ByVal filaTotal As Long, _
                                  ByVal filaAcum As Long)

    Dim wsDst As Worksheet
    Dim filaCabecera As Long
    Dim colLinea As Long
    Dim ultFila As Long
    Dim i As Integer
    Dim r As Long
    Dim linea As String
    Dim lineas As Variant
    Dim filaEncontrada As Long
    Dim filaAcumRaw As Long
    Dim j As Integer

    Set wsDst = ThisWorkbook.Worksheets(HOJA_REP2)

    filaCabecera = BuscarFilaCabeceraAcumuladoV6(wsRaw)

    If filaCabecera = 0 Then
        Err.Raise vbObjectError + 2620, , _
            "No se encontro la cabecera 'Linea' del archivo P_ACU."
    End If

    colLinea = BuscarColumnaLineaV6(wsRaw, filaCabecera)

    If colLinea = 0 Then
        Err.Raise vbObjectError + 2621, , _
            "No se identifico la columna Linea en P_ACU."
    End If

    ultFila = UltimaFilaUsada(wsRaw)

    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")

    ' Limpiar solo la seccion SAP de destino.
    wsDst.Range("A" & filaBase & ":N" & filaAcum).ClearContents


    ' ---------------------------------------------------------
    ' LINEAS 1 - 4
    ' ---------------------------------------------------------
    For i = 0 To 3

        linea = CStr(lineas(i))
        filaEncontrada = 0

        For r = filaCabecera + 1 To ultFila

            If UCase$(Trim$(CStr(wsRaw.Cells(r, colLinea).Value))) = linea Then
                filaEncontrada = r
                Exit For
            End If

        Next r

        ' Siempre conservar el identificador de linea.
        wsDst.Cells(filaBase + i, 1).Value = linea

        If filaEncontrada > 0 Then

            ' Desde LINEA existen exactamente 14 campos que alimentan A:N:
            ' Linea
            ' Prod T1
            ' Prod T2
            ' Total Prod
            ' Efi T1
            ' Efi T2
            ' Efi Total
            ' Cambio/Saneamiento
            ' P. Ajenos
            ' P. Operativos
            ' P. Eq. Linea
            ' P.Serv.Aux.
            ' Util Linea
            ' % Utilizacion
            For j = 0 To 13
                wsDst.Cells(filaBase + i, 1 + j).Value = _
                    wsRaw.Cells(filaEncontrada, colLinea + j).Value
            Next j

        End If

    Next i


    ' ---------------------------------------------------------
    ' TOTAL PRODUCCION B:D
    ' ---------------------------------------------------------
    wsDst.Cells(filaTotal, 2).Formula = _
        "=SUM(B" & filaBase & ":B" & (filaBase + 3) & ")"

    wsDst.Cells(filaTotal, 3).Formula = _
        "=SUM(C" & filaBase & ":C" & (filaBase + 3) & ")"

    wsDst.Cells(filaTotal, 4).Formula = _
        "=SUM(D" & filaBase & ":D" & (filaBase + 3) & ")"


    ' ---------------------------------------------------------
    ' FILA Tot. / Acum.
    '
    ' En el video la salida P_ACU muestra:
    '   Tot.   Acum.
    '
    ' "Acum." esta en la MISMA columna donde normalmente aparece
    ' LINEA001, por eso copiamos desde esa celda hasta N.
    ' ---------------------------------------------------------
    filaAcumRaw = 0

    For r = filaCabecera + 1 To ultFila

        If NormalizarSAPV6(CStr(wsRaw.Cells(r, colLinea).Value)) = "ACUM" Then
            filaAcumRaw = r
            Exit For
        End If

    Next r

    If filaAcumRaw = 0 Then
        Err.Raise vbObjectError + 2622, , _
            "No se encontro la fila Tot./Acum. en la salida P_ACU."
    End If

    For j = 0 To 13
        wsDst.Cells(filaAcum, 1 + j).Value = _
            wsRaw.Cells(filaAcumRaw, colLinea + j).Value
    Next j

End Sub


Private Function BuscarFilaCabeceraAcumuladoV6(ByVal ws As Worksheet) As Long

    Dim ultFila As Long
    Dim ultCol As Long
    Dim r As Long
    Dim c As Long
    Dim txt As String

    ultFila = UltimaFilaUsada(ws)
    ultCol = UltimaColumnaUsadaV6(ws)

    If ultFila > 60 Then ultFila = 60

    For r = 1 To ultFila

        For c = 1 To ultCol

            txt = NormalizarSAPV6(CStr(ws.Cells(r, c).Value))

            If txt = "LINEA" Then
                BuscarFilaCabeceraAcumuladoV6 = r
                Exit Function
            End If

        Next c

    Next r

End Function


Private Function BuscarColumnaLineaV6(ByVal ws As Worksheet, _
                                      ByVal filaCabecera As Long) As Long

    Dim ultCol As Long
    Dim c As Long

    ultCol = UltimaColumnaUsadaV6(ws)

    For c = 1 To ultCol

        If NormalizarSAPV6(CStr(ws.Cells(filaCabecera, c).Value)) = "LINEA" Then
            BuscarColumnaLineaV6 = c
            Exit Function
        End If

    Next c

End Function


Private Function NormalizarSAPV6(ByVal texto As String) As String

    texto = UCase$(Trim$(texto))

    texto = Replace(texto, "" & ChrW(193) & "", "A")
    texto = Replace(texto, "" & ChrW(201) & "", "E")
    texto = Replace(texto, "" & ChrW(205) & "", "I")
    texto = Replace(texto, "" & ChrW(211) & "", "O")
    texto = Replace(texto, "" & ChrW(218) & "", "U")
    texto = Replace(texto, "" & ChrW(220) & "", "U")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ":", "")
    texto = Replace(texto, " ", "")

    NormalizarSAPV6 = texto

End Function


Private Sub CargarREP2DesdeRawV6(ByVal wsRaw As Worksheet)
    Dim ws As Worksheet
    Dim dia As Range
    Dim mes As Range
    Dim acum1 As Range
    Dim acum2 As Range
    Dim rIniDia As Long
    Dim rFinDia As Long
    Dim rIniMes As Long
    Dim rFinMes As Long
    Dim ultimaFila As Long
    Dim cabDetalle As Range

    Set ws = ThisWorkbook.Worksheets(HOJA_REP2)
    ultimaFila = UltimaFilaUsada(wsRaw)

    Set dia = BuscarExactoV6(wsRaw, "DIA", 1, ultimaFila)
    Set mes = BuscarExactoV6(wsRaw, "MES", 1, ultimaFila)

    ' Si aparece esta cabecera, SAP devolvio el modo DETALLE durante la etapa RESUMEN.
    ' Es exactamente el formato del REP002_20260810.xls que compartiste.
    Set cabDetalle = BuscarTextoParcialV6(wsRaw, "Fecha de inicio Real")
    If Not cabDetalle Is Nothing Then
        Err.Raise vbObjectError + 2421, , _
            "SAP devolvio REP2 en modo DETALLE cuando se esperaba RESUMEN. " & _
            "La opcion P_DET quedo activa. Ejecuta V6_DiagnosticarModosREP2 si vuelve a ocurrir."
    End If

    If Not dia Is Nothing And Not mes Is Nothing Then
        rIniDia = dia.Row + 1
        rFinDia = mes.Row - 1
        rIniMes = mes.Row + 1
        rFinMes = ultimaFila
    Else
        Set acum1 = BuscarNOcurrenciaV6(wsRaw, "Acum.", 1)
        Set acum2 = BuscarNOcurrenciaV6(wsRaw, "Acum.", 2)

        If acum1 Is Nothing Or acum2 Is Nothing Then
            Err.Raise vbObjectError + 2420, , _
                "No se encontraron DIA/MES en el archivo de RESUMEN. " & _
                "La V6 ya valida el modo antes de exportar; ejecuta V6_DiagnosticarModosREP2."
        End If

        rIniDia = 1
        rFinDia = acum1.Row
        rIniMes = acum1.Row + 1
        rFinMes = acum2.Row
    End If

    ' Solo se reemplazan las zonas que contienen informacion proveniente de SAP.
    ' Las formulas auxiliares del REP2, por ejemplo filas 10:13, se conservan.
    ws.Range("A3:N8").ClearContents
    ws.Range("A16:N21").ClearContents

    CargarSeccionREP2V6 wsRaw, rIniDia, rFinDia, ws, 3, 7, 8
    CargarSeccionREP2V6 wsRaw, rIniMes, rFinMes, ws, 16, 20, 21
End Sub

Private Sub CargarSeccionREP2V6(ByVal wsRaw As Worksheet, _
                                ByVal filaIni As Long, _
                                ByVal filaFin As Long, _
                                ByVal wsDst As Worksheet, _
                                ByVal filaBase As Long, _
                                ByVal filaTotal As Long, _
                                ByVal filaAcum As Long)
    Dim lineas As Variant
    Dim i As Integer
    Dim c As Range
    Dim dstRow As Long
    Dim valores As Variant
    Dim totalRaw As Range

    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")

    For i = LBound(lineas) To UBound(lineas)
        dstRow = filaBase + i
        wsDst.Range("A" & dstRow & ":N" & dstRow).ClearContents
        wsDst.Cells(dstRow, 1).Value = lineas(i)

        Set c = BuscarExactoV6(wsRaw, CStr(lineas(i)), filaIni, filaFin)
        If Not c Is Nothing Then
            valores = wsRaw.Range(wsRaw.Cells(c.Row, c.Column), _
                                  wsRaw.Cells(c.Row, c.Column + 13)).Value2
            wsDst.Range("A" & dstRow & ":N" & dstRow).Value2 = valores
        End If
    Next i

    ' Total de produccion: B:D. Las demas columnas de esta fila se dejan limpias.
    wsDst.Range("A" & filaTotal & ":N" & filaTotal).ClearContents
    wsDst.Cells(filaTotal, 2).Formula = "=SUM(B" & filaBase & ":B" & (filaBase + 3) & ")"
    wsDst.Cells(filaTotal, 3).Formula = "=SUM(C" & filaBase & ":C" & (filaBase + 3) & ")"
    wsDst.Cells(filaTotal, 4).Formula = "=SUM(D" & filaBase & ":D" & (filaBase + 3) & ")"

    ' Fila Acum. del SAP: contiene la eficiencia y distribucion consolidada de planta.
    wsDst.Range("A" & filaAcum & ":N" & filaAcum).ClearContents
    Set c = BuscarExactoV6(wsRaw, "Acum.", filaIni, filaFin)
    If Not c Is Nothing Then
        valores = wsRaw.Range(wsRaw.Cells(c.Row, c.Column), _
                              wsRaw.Cells(c.Row, c.Column + 13)).Value2
        wsDst.Range("A" & filaAcum & ":N" & filaAcum).Value2 = valores
    Else
        wsDst.Cells(filaAcum, 1).Value = "Acum."
    End If
End Sub

Private Function BuscarExactoV6(ByVal ws As Worksheet, _
                                ByVal texto As String, _
                                ByVal filaIni As Long, _
                                ByVal filaFin As Long) As Range
    Dim ultimaCol As Long
    Dim rng As Range

    If filaIni < 1 Then filaIni = 1
    If filaFin < filaIni Then Exit Function

    ultimaCol = UltimaColumnaUsadaV6(ws)
    If ultimaCol < 1 Then ultimaCol = 1

    Set rng = ws.Range(ws.Cells(filaIni, 1), ws.Cells(filaFin, ultimaCol))

    On Error Resume Next
    Set BuscarExactoV6 = rng.Find(What:=texto, After:=rng.Cells(rng.Cells.Count), _
                                  LookIn:=xlValues, LookAt:=xlWhole, _
                                  SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                                  MatchCase:=False)
    On Error GoTo 0
End Function

Private Function BuscarNOcurrenciaV6(ByVal ws As Worksheet, _
                                     ByVal texto As String, _
                                     ByVal numero As Long) As Range
    Dim rng As Range
    Dim c As Range
    Dim primero As String
    Dim n As Long

    Set rng = ws.UsedRange
    Set c = rng.Find(What:=texto, After:=rng.Cells(rng.Cells.Count), _
                     LookIn:=xlValues, LookAt:=xlWhole, _
                     SearchOrder:=xlByRows, SearchDirection:=xlNext, _
                     MatchCase:=False)

    If c Is Nothing Then Exit Function

    primero = c.Address
    Do
        n = n + 1
        If n = numero Then
            Set BuscarNOcurrenciaV6 = c
            Exit Function
        End If
        Set c = rng.FindNext(c)
        If c Is Nothing Then Exit Do
    Loop While c.Address <> primero
End Function

Private Function UltimaColumnaUsadaV6(ByVal ws As Worksheet) As Long
    Dim c As Range

    On Error Resume Next
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), _
                          LookIn:=xlFormulas, LookAt:=xlPart, _
                          SearchOrder:=xlByColumns, SearchDirection:=xlPrevious)
    On Error GoTo 0

    If c Is Nothing Then
        UltimaColumnaUsadaV6 = 1
    Else
        UltimaColumnaUsadaV6 = c.Column
    End If
End Function

' =====================================================================
' LECTURA DEL EXPORT /VAR_PAROS Y GENERACION DE COMENTARIOS
'
' Layout esperado:
'   Linea | Fecha Contable | Tipo de Paro | Clave de Paro |
'   Subclave de Paro | Minutos de paro | Ptos. Efi. Perdid. |
'   Tripulacion | Intervalo | Observaciones
'
' Reglas:
'   - Detecta columnas por encabezado, nunca por posicion fija.
'   - Clave de Paro, Subclave y Observaciones son opcionales.
'   - Causa exacta: Clave > Subclave > Observaciones > Tipo.
'   - PRODUCCION se conserva en PAROS_AGRUPADOS para validar 1440 min,
'     pero NO participa en el TOP 3 de paros.
'   - Los minutos conservan hasta 2 decimales.
' =====================================================================
Private Sub ProcesarComentarios(ByVal ws As Worksheet, ByVal fechaReporte As Date)
    Dim filaInicio As Long, colFecha As Long, colLinea As Long
    Dim colTipo As Long, colClave As Long, colSubclave As Long
    Dim colObs As Long, colMin As Long
    Dim ultimaFila As Long, fila As Long, idx As Long, i As Long
    Dim minutos As Double, tipo As String, clave As String
    Dim subclave As String, observacion As String, causa As String
    Dim claveDic As String, linea As String
    Dim registrosFecha As Long, registrosLinea(0 To 3) As Long
    Dim dicParoMin(0 To 3) As Object, dicParoTexto(0 To 3) As Object
    Dim dicTodoMin(0 To 3) As Object, dicTodoTexto(0 To 3) As Object
    Dim dicTodoTipo(0 To 3) As Object
    Dim wsRep As Worksheet, filasDestino As Variant

    filasDestino = Array(6, 7, 8, 9)

    For i = 0 To 3
        Set dicParoMin(i) = CreateObject("Scripting.Dictionary")
        Set dicParoTexto(i) = CreateObject("Scripting.Dictionary")
        Set dicTodoMin(i) = CreateObject("Scripting.Dictionary")
        Set dicTodoTexto(i) = CreateObject("Scripting.Dictionary")
        Set dicTodoTipo(i) = CreateObject("Scripting.Dictionary")
        dicParoMin(i).CompareMode = vbTextCompare
        dicParoTexto(i).CompareMode = vbTextCompare
        dicTodoMin(i).CompareMode = vbTextCompare
        dicTodoTexto(i).CompareMode = vbTextCompare
        dicTodoTipo(i).CompareMode = vbTextCompare
    Next i

    DetectarColumnasPAROS ws, filaInicio, colFecha, colLinea, _
                           colTipo, colClave, colSubclave, colObs, colMin
    ultimaFila = UltimaFilaUsada(ws)

    For fila = filaInicio To ultimaFila
        If MismaFecha(ws.Cells(fila, colFecha).Value, fechaReporte) Then
            registrosFecha = registrosFecha + 1
            linea = UCase$(Trim$(CStr(ws.Cells(fila, colLinea).Value)))
            idx = IndiceLinea(linea)

            If idx >= 0 Then
                registrosLinea(idx) = registrosLinea(idx) + 1
                minutos = NumeroSeguro(ws.Cells(fila, colMin).Value)

                If minutos > 0 Then
                    tipo = LimpiarTexto(ValorCeldaOpcional(ws, fila, colTipo))
                    clave = LimpiarTexto(ValorCeldaOpcional(ws, fila, colClave))
                    subclave = LimpiarTexto(ValorCeldaOpcional(ws, fila, colSubclave))
                    observacion = LimpiarTexto(ValorCeldaOpcional(ws, fila, colObs))
                    causa = ObtenerCausaSAPExacta(tipo, clave, subclave, observacion)

                    If causa <> vbNullString Then
                        claveDic = UCase$(causa)

                        If dicTodoMin(idx).Exists(claveDic) Then
                            dicTodoMin(idx).Item(claveDic) = CDbl(dicTodoMin(idx).Item(claveDic)) + minutos
                        Else
                            dicTodoMin(idx).Add claveDic, minutos
                            dicTodoTexto(idx).Add claveDic, causa
                            dicTodoTipo(idx).Add claveDic, tipo
                        End If

                        If Not EsProduccionSAP(tipo, causa) Then
                            If dicParoMin(idx).Exists(claveDic) Then
                                dicParoMin(idx).Item(claveDic) = CDbl(dicParoMin(idx).Item(claveDic)) + minutos
                            Else
                                dicParoMin(idx).Add claveDic, minutos
                                dicParoTexto(idx).Add claveDic, causa
                            End If
                        End If
                    End If
                End If
            End If
        End If
    Next fila

    If registrosFecha = 0 Then
        Err.Raise vbObjectError + 2010, , _
            "El archivo /VAR_PAROS no contiene registros para " & Format$(fechaReporte, "dd/mm/yyyy") & "."
    End If

    Set wsRep = ThisWorkbook.Worksheets(HOJA_REPORTE)
    For i = 0 To 3
        If dicParoMin(i).Count > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = CrearResumenExacto(dicParoMin(i), dicParoTexto(i), MAX_COMENTARIOS)
        ElseIf InStr(1, UCase$(CStr(wsRep.Cells(filasDestino(i), 3).Value)), "PARO", vbTextCompare) > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = "PARO PROGRAMADO"
        ElseIf registrosLinea(i) > 0 Then
            wsRep.Range("W" & filasDestino(i)).Value = "SIN PAROS"
        Else
            wsRep.Range("W" & filasDestino(i)).Value = "SIN PRODUCCION / SIN REGISTROS"
        End If
        wsRep.Range("W" & filasDestino(i)).WrapText = True
    Next i

    EscribirParosAgrupados fechaReporte, dicTodoMin, dicTodoTexto, dicTodoTipo
End Sub

Private Sub DetectarColumnasPAROS(ByVal ws As Worksheet, _
                                   ByRef filaInicio As Long, _
                                   ByRef colFecha As Long, ByRef colLinea As Long, _
                                   ByRef colTipo As Long, ByRef colClave As Long, _
                                   ByRef colSubclave As Long, ByRef colObs As Long, _
                                   ByRef colMin As Long)
    Dim filaEncabezado As Long
    filaEncabezado = BuscarFilaEncabezadosPAROS(ws)
    If filaEncabezado = 0 Then Err.Raise vbObjectError + 2720, , "No se encontro la fila de encabezados del layout /VAR_PAROS."

    colLinea = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("L" & ChrW(237) & "nea", "Linea"))
    colFecha = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Fecha Contable"))
    colTipo = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Tipo de Paro"))
    colClave = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Clave de Paro"))
    colSubclave = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Subclave de Paro"))
    colMin = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Minutos de paro", "Minutos de Paro"))
    colObs = BuscarEncabezadoEnFila(ws, filaEncabezado, Array("Observaciones", "Observacion", "Observaci" & ChrW(243) & "n"))

    If colLinea = 0 Or colFecha = 0 Or colTipo = 0 Or colMin = 0 Then
        Err.Raise vbObjectError + 2721, , _
            "El layout /VAR_PAROS requiere Linea, Fecha Contable, Tipo de Paro y Minutos de paro."
    End If
    filaInicio = filaEncabezado + 1
End Sub

Private Function BuscarFilaEncabezadosPAROS(ByVal ws As Worksheet) As Long
    Dim r As Long, cLinea As Long, cFecha As Long, cTipo As Long, limite As Long
    limite = 50
    If UltimaFilaUsada(ws) < limite Then limite = UltimaFilaUsada(ws)
    For r = 1 To limite
        cLinea = BuscarEncabezadoEnFila(ws, r, Array("L" & ChrW(237) & "nea", "Linea"))
        cFecha = BuscarEncabezadoEnFila(ws, r, Array("Fecha Contable"))
        cTipo = BuscarEncabezadoEnFila(ws, r, Array("Tipo de Paro"))
        If cLinea > 0 And cFecha > 0 And cTipo > 0 Then
            BuscarFilaEncabezadosPAROS = r
            Exit Function
        End If
    Next r
End Function

Private Function BuscarEncabezadoEnFila(ByVal ws As Worksheet, ByVal fila As Long, ByVal nombres As Variant) As Long
    Dim ultimaCol As Long, c As Long, i As Long
    Dim actual As String, buscado As String
    ultimaCol = UltimaColumnaUsadaV6(ws)
    For c = 1 To ultimaCol
        actual = NormalizarEncabezadoParos(CStr(ws.Cells(fila, c).Value))
        If actual <> vbNullString Then
            For i = LBound(nombres) To UBound(nombres)
                buscado = NormalizarEncabezadoParos(CStr(nombres(i)))
                If actual = buscado Then
                    BuscarEncabezadoEnFila = c
                    Exit Function
                End If
            Next i
        End If
    Next c
End Function

Private Function NormalizarEncabezadoParos(ByVal texto As String) As String
    texto = UCase$(LimpiarTexto(texto))
    texto = Replace(texto, "" & ChrW(193) & "", "A")
    texto = Replace(texto, "" & ChrW(201) & "", "E")
    texto = Replace(texto, "" & ChrW(205) & "", "I")
    texto = Replace(texto, "" & ChrW(211) & "", "O")
    texto = Replace(texto, "" & ChrW(218) & "", "U")
    texto = Replace(texto, ".", "")
    texto = Replace(texto, ":", "")
    Do While InStr(texto, "  ") > 0
        texto = Replace(texto, "  ", " ")
    Loop
    NormalizarEncabezadoParos = Trim$(texto)
End Function

Private Function ValorCeldaOpcional(ByVal ws As Worksheet, ByVal fila As Long, ByVal columna As Long) As String
    If columna <= 0 Then Exit Function
    On Error Resume Next
    ValorCeldaOpcional = CStr(ws.Cells(fila, columna).Value)
    On Error GoTo 0
End Function

Private Function ObtenerCausaSAPExacta(ByVal tipo As String, ByVal clave As String, _
                                        ByVal subclave As String, ByVal observacion As String) As String
    tipo = LimpiarTexto(tipo): clave = LimpiarTexto(clave)
    subclave = LimpiarTexto(subclave): observacion = LimpiarTexto(observacion)

    If InStr(1, UCase$(tipo), "PRODUCCION", vbTextCompare) > 0 Then
        ObtenerCausaSAPExacta = "PRODUCCION"
    ElseIf clave <> vbNullString Then
        ObtenerCausaSAPExacta = clave
    ElseIf subclave <> vbNullString Then
        ObtenerCausaSAPExacta = subclave
    ElseIf observacion <> vbNullString Then
        ObtenerCausaSAPExacta = observacion
    Else
        ObtenerCausaSAPExacta = tipo
    End If
End Function

Private Function EsProduccionSAP(ByVal tipo As String, ByVal causa As String) As Boolean
    EsProduccionSAP = (InStr(1, UCase$(tipo), "PRODUCCION", vbTextCompare) > 0 Or UCase$(Trim$(causa)) = "PRODUCCION")
End Function

Private Function CrearResumenExacto(ByVal dicMin As Object, ByVal dicTexto As Object, ByVal maxElementos As Long) As String
    Dim keysOrdenadas As Variant, i As Long, limite As Long, salida As String, clave As String
    If dicMin.Count = 0 Then Exit Function
    keysOrdenadas = ClavesOrdenadasPorMinutos(dicMin)
    limite = UBound(keysOrdenadas) - LBound(keysOrdenadas) + 1
    If limite > maxElementos Then limite = maxElementos
    For i = LBound(keysOrdenadas) To LBound(keysOrdenadas) + limite - 1
        clave = CStr(keysOrdenadas(i))
        If salida <> vbNullString Then salida = salida & vbLf
        salida = salida & FormatearMinutosExactos(CDbl(dicMin.Item(clave))) & " min " & CStr(dicTexto.Item(clave))
    Next i
    CrearResumenExacto = salida
End Function

Private Function FormatearMinutosExactos(ByVal valor As Double) As String
    valor = Application.WorksheetFunction.Round(valor, 2)
    FormatearMinutosExactos = Format$(valor, "0.##")
End Function

Private Function ClavesOrdenadasPorMinutos(ByVal dicMin As Object) As Variant
    Dim keys As Variant, i As Long, j As Long, tmp As Variant
    keys = dicMin.Keys
    For i = LBound(keys) To UBound(keys) - 1
        For j = i + 1 To UBound(keys)
            If CDbl(dicMin.Item(keys(j))) > CDbl(dicMin.Item(keys(i))) Then
                tmp = keys(i): keys(i) = keys(j): keys(j) = tmp
            End If
        Next j
    Next i
    ClavesOrdenadasPorMinutos = keys
End Function

Private Sub EscribirParosAgrupados(ByVal fechaReporte As Date, _
                                    ByRef dicTodoMin() As Object, _
                                    ByRef dicTodoTexto() As Object, _
                                    ByRef dicTodoTipo() As Object)
    Dim ws As Worksheet, lineas As Variant, i As Long, r As Long, k As Long
    Dim rankParo As Long, totalLinea As Double, keys As Variant, clave As String
    Dim minutos As Double, tipo As String, causa As String, clasificacion As String

    lineas = Array("LINEA001", "LINEA002", "LINEA003", "LINEA004")
    On Error Resume Next
    Set ws = ThisWorkbook.Worksheets(HOJA_PAROS_AGRUPADOS)
    On Error GoTo 0
    If ws Is Nothing Then
        Set ws = ThisWorkbook.Worksheets.Add(After:=ThisWorkbook.Worksheets(ThisWorkbook.Worksheets.Count))
        ws.Name = HOJA_PAROS_AGRUPADOS
    Else
        ws.Cells.Clear
    End If

    ws.Cells(1, 1).Value = "Fecha"
    ws.Cells(1, 2).Value = "Linea"
    ws.Cells(1, 3).Value = "Tipo de Paro"
    ws.Cells(1, 4).Value = "Causa SAP"
    ws.Cells(1, 5).Value = "Minutos"
    ws.Cells(1, 6).Value = "% del dia"
    ws.Cells(1, 7).Value = "Ranking Paro"
    ws.Cells(1, 8).Value = "Clasificacion"
    With ws.Range("A1:H1")
        .Font.Bold = True: .Interior.Color = RGB(35, 35, 35): .Font.Color = RGB(255, 255, 255)
    End With

    r = 2
    For i = 0 To 3
        totalLinea = 0: rankParo = 0
        If dicTodoMin(i).Count > 0 Then
            keys = ClavesOrdenadasPorMinutos(dicTodoMin(i))
            For k = LBound(keys) To UBound(keys)
                clave = CStr(keys(k))
                minutos = CDbl(dicTodoMin(i).Item(clave))
                causa = CStr(dicTodoTexto(i).Item(clave))
                tipo = CStr(dicTodoTipo(i).Item(clave))
                totalLinea = totalLinea + minutos
                If EsProduccionSAP(tipo, causa) Then
                    clasificacion = "PRODUCCION"
                Else
                    clasificacion = "PARO": rankParo = rankParo + 1
                End If
                ws.Cells(r, 1).Value = fechaReporte
                ws.Cells(r, 2).Value = lineas(i)
                ws.Cells(r, 3).Value = tipo
                ws.Cells(r, 4).Value = causa
                ws.Cells(r, 5).Value = minutos
                ws.Cells(r, 6).Value = minutos / 1440#
                If clasificacion = "PARO" Then ws.Cells(r, 7).Value = rankParo
                ws.Cells(r, 8).Value = clasificacion
                r = r + 1
            Next k
        End If

        ws.Cells(r, 1).Value = fechaReporte
        ws.Cells(r, 2).Value = lineas(i)
        ws.Cells(r, 4).Value = "TOTAL LINEA"
        ws.Cells(r, 5).Value = totalLinea
        ws.Cells(r, 6).Value = totalLinea / 1440#
        ws.Cells(r, 8).Value = "VALIDACION"
        ws.Range("A" & r & ":H" & r).Font.Bold = True
        r = r + 2
    Next i

    ws.Columns("A:H").AutoFit
    ws.Columns("A").NumberFormat = "dd/mm/yyyy"
    ws.Columns("E").NumberFormat = "0.00"
    ws.Columns("F").NumberFormat = "0.00%"
End Sub

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


' ============================================================
' LIMPIAR CAMPOS AUTOMATICOS - REPORTE EFICIENCIA V6
'
' Borra unicamente informacion generada por SAP / macro:
'   - REP2 DIA
'   - REP2 MES
'   - Fecha inicio real / eficiencia acumulada
'   - Comentarios W6:W9
'   - Metadatos SAP X36:Y36
'   - Hoja REP2_DETALLE (si existe)
'
' NO BORRA:
'   - SKU / programa
'   - CTT
'   - Ago 26 / PDP
'   - Formulas de REPORTE EFICIENCIA
' ============================================================

Public Sub LimpiarCamposAutomaticos()

    Dim wsREP2 As Worksheet
    Dim wsReporte As Worksheet
    Dim wsDetalle As Worksheet
    Dim wsPDP As Worksheet
    Dim respuesta As VbMsgBoxResult

    On Error GoTo ERROR_GENERAL

    respuesta = MsgBox( _
        "Se borraran solamente los datos cargados automaticamente por SAP." & vbCrLf & vbCrLf & _
        "Se conservaran:" & vbCrLf & _
        "- SKU / Programa" & vbCrLf & _
        "- CTT" & vbCrLf & _
        "- PLAN PDP" & vbCrLf & _
        "- Formulas del reporte" & vbCrLf & vbCrLf & _
        "Deseas continuar?", _
        vbQuestion + vbYesNo, _
        "LIMPIAR DATOS AUTOMATICOS")

    If respuesta <> vbYes Then Exit Sub

    Application.ScreenUpdating = False
    Application.EnableEvents = False

    Set wsREP2 = ThisWorkbook.Worksheets("REP2")
    Set wsReporte = ThisWorkbook.Worksheets("REPORTE EFICIENCIA")
    Set wsPDP = ThisWorkbook.Worksheets(HOJA_PDP)

    ' ========================================================
    ' REP2 - INFORMACION DEL DIA
    ' ========================================================

    ' LINEA001 a LINEA004
    wsREP2.Range("B3:N6").ClearContents

    ' Total produccion del dia
    wsREP2.Range("B7:D7").ClearContents

    ' Acumulado del dia
    wsREP2.Range("B8:N8").ClearContents

    ' Fecha inicio real / eficiencia acumulada
    wsREP2.Range("O3:P8").ClearContents

    ' ========================================================
    ' REP2 - INFORMACION DEL MES
    ' ========================================================

    ' LINEA001 a LINEA004
    wsREP2.Range("B16:N19").ClearContents

    ' Total produccion mes
    wsREP2.Range("B20:D20").ClearContents

    ' Acumulado mensual
    wsREP2.Range("B21:N21").ClearContents

    ' Fecha inicio real / eficiencia acumulada
    wsREP2.Range("O16:P21").ClearContents

    ' ========================================================
    ' REPORTE EFICIENCIA
    ' ========================================================

    ' TOP 3 comentarios de paros L1-L4
    wsReporte.Range("W6:W9").ClearContents

    ' Metadatos tecnicos utilizados por el Dashboard PRO
    wsReporte.Range("X36:Y36").ClearContents

    ' ========================================================
    ' AGO 26 - PRODUCCION REAL AUTOMATICA
    '
    ' NO se toca BJ:BN (PLAN PDP).
    ' ========================================================
    wsPDP.Range("BE3:BH34").ClearContents
    wsPDP.Range("I6:I10").ClearContents
    wsPDP.Range("R6:R10").ClearContents


    ' ========================================================
    ' REP2_DETALLE
    ' ========================================================

    On Error Resume Next
    Set wsDetalle = ThisWorkbook.Worksheets("REP2_DETALLE")
    On Error GoTo ERROR_GENERAL

    If Not wsDetalle Is Nothing Then
        wsDetalle.Cells.ClearContents
    End If

    ' ========================================================
    ' RECALCULAR
    ' ========================================================

    Application.CalculateFullRebuild

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "DATOS AUTOMATICOS BORRADOS CORRECTAMENTE." & vbCrLf & vbCrLf & _
        "El reporte esta listo para ejecutar nuevamente:" & vbCrLf & _
        "V6_ActualizarReporteCompletoSAP", _
        vbInformation, _
        "REPORTE EFICIENCIA"

    Exit Sub

ERROR_GENERAL:

    Application.EnableEvents = True
    Application.ScreenUpdating = True

    MsgBox _
        "No se pudo completar la limpieza." & vbCrLf & vbCrLf & _
        "Error " & Err.Number & ": " & Err.Description, _
        vbExclamation, _
        "LIMPIAR DATOS AUTOMATICOS"

End Sub
