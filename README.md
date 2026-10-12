# AssetGuard — Automatización SAP para planta industrial

Sistema de automatización SAP GUI Scripting en VBA, desarrollado por **José Luis Cruz Prieto**.

Automatiza la operación diaria de una planta de producción: eficiencia, paros de línea, producción real, plan de producción (PDP) y bajas de EPP en MIGO. Todo corre desde Excel con una sola macro.

## Estructura

```
sap-automation-plant/
├── README.md
├── macros/
│   ├── modReporteSAP_V6_5.bas        ← reporte de eficiencia completo
│   ├── modMIGO_BajasEPP.bas          ← bajas de EPP en MIGO, v3 (~1,100 líneas)
│   ├── modCorreosNotasCredito.bas    ← correos masivos de notas de crédito, v3 (~1,180 líneas)
│   ├── modComentariosParosSAP_v2.bas ← comentarios TOP 3 de paros
│   └── modConsolidarMB51.bas         ← consolidador de exportaciones MB51
├── docs/
│   ├── layout-bajas-epp.md           ← hojas, columnas, CONFIG y estados
│   └── layout-correos-notas-credito.md
└── samples/                          ← CSV con datos ficticios para armar los libros
    ├── bajas-epp/
    └── correos-notas-credito/
```

## Módulos

### modReporteSAP_V6_5.bas — Reporte de eficiencia

Una sola macro (`V6_ActualizarReporteCompletoSAP`) ejecuta seis pasos:

1. **Eficiencia del día** — ZMNPPREP0002 variante /DSB, modo P_ACU, solo fecha de corte.
2. **Eficiencia acumulada del mes** — misma transacción, del día 1 al corte.
3. **Detalle REP2** — P_DET: fecha de inicio real y % eficiencia acumulada.
4. **Producción real** — ZMNPPREP0006 variante /BRBR: cajas unitarias por día y línea → hoja Ago 26.
5. **PDP / Production Plan** — lee el archivo más reciente de `PDP\ENTRADA`, convierte CF a unidades de consumo y actualiza el plan.
6. **Comentarios de paros** — ZMNPPREP0004 variante /SMOSKE_ALT: descarga el detalle del día, suma minutos por causa, toma el TOP 3 por línea y escribe REPORTE EFICIENCIA!W6:W9.

También genera la hoja PAROS_AGRUPADOS con validación de 1,440 minutos por línea (24 h), y `LimpiarCamposAutomaticos` borra solo lo que cargó SAP sin tocar SKU, CTT ni fórmulas.

**API pública:** V6_ActualizarReporteCompletoSAP, V6_ActualizarSoloREP2SAP, V6_ActualizarSoloAgo26SAP, V6_ActualizarSoloPDPDesdeCarpeta, V6_ActualizarSoloComentariosSAP, V6_ProbarConexionSAPMinima, V6_DiagnosticarSAPGUI, LimpiarCamposAutomaticos.

### modMIGO_BajasEPP.bas — Bajas de EPP en MIGO (v3)

Captura en lote las bajas de equipo de protección personal en MIGO (A07 / R10 / movimiento 201, salida a centro de costo) a partir del export de la app de EPP.

- **Sin desfase de filas:** `Preparar_Bajas_Desde_App` lee PEGAR_APP en arreglos y escribe BAJAS_SAP como valores, sin fórmulas. Si los conteos no cuadran, Validar y Procesar se detienen.
- **Cabecera desde CONFIG:** operación, referencia, movimiento 201, centro y almacén. No hay valores de planta en el código.
- **Sesión correcta:** valida el ID de sistema y el mandante contra CONFIG antes de tocar SAP. La espera de SAP tiene tiempo límite.
- **Mensajes por fila:** lee la barra de estado y las ventanas emergentes *antes* de cerrarlas. Una línea rechazada se retira de MIGO y queda en `ERROR` con su mensaje.
- **Estados honestos:** un mensaje vacío deja la fila en `REVISAR`. Si la verificación pasa, la fila queda en `VERIFICADO_SIN_CONTABILIZAR`, que no bloquea el reproceso. Solo `CONTABILIZADO`, con documento, se omite.
- **Seguridad:** `MODO = SIMULACION` (dry-run, no se conecta a SAP) o `VERIFICAR`. **Nunca contabiliza.** La contabilización se hace a mano y se registra con `Registrar_Contabilizacion_Manual`.
- Columnas de MIGO por nombre técnico (GOITEM-MATNR, GOITEM-KOSTL…). Manejo de errores por fila con `Resume`, bitácora en LOG_SAP, y ScreenUpdating/Calculation siempre restaurados.

**API pública:** Preparar_Bajas_Desde_App, Validar_Bajas_EPP, Procesar_Bajas_EPP_MIGO, Registrar_Contabilizacion_Manual, Probar_Conexion_SAP, Diagnosticar_Layout_MIGO.

**Setup:** ver [docs/layout-bajas-epp.md](docs/layout-bajas-epp.md) y [samples/bajas-epp](samples/bajas-epp).

### modCorreosNotasCredito.bas — Correos masivos de notas de crédito (v3)

Envío por Outlook de una nota de crédito manual por distribuidor: asunto automático, carátula desde Resumen, detalle de partidas filtrado por distribuidor, tres textos configurables, firma y adjuntos.

- **Distribuidor obligatorio** por fila (columna `Distribuidor`) y validado contra Resumen. Ya no se usa la posición del ID, que podía mandarle a un distribuidor los datos de otro.
- **No reenvía:** las filas con Estado `Enviado` se omiten.
- **Modos:** `REVISAR` (abre cada correo), `BORRADOR` (lo guarda en Borradores) y `ENVIAR` (pide confirmación). **Al terminar, el modo regresa a REVISAR.**
- Si falta un adjunto o algo falla, el correo se descarta (`Close olDiscard`) y no queda abierto.
- Valida el formato de los correos y, opcionalmente, una lista de dominios permitidos. La bitácora en LOG_CORREOS guarda solo conteos, nunca direcciones.

**API pública:** PrepararConfiguracionCorreos, ValidarFilasCorreos, GenerarOCrearCorreosMasivos.

**Setup:** ver [docs/layout-correos-notas-credito.md](docs/layout-correos-notas-credito.md) y [samples/correos-notas-credito](samples/correos-notas-credito).

### modComentariosParosSAP_v2.bas — Comentarios de paros de línea

- Lee la fecha de REPORTE EFICIENCIA!F3.
- Ejecuta ZMNPPREP0004 con variante /SMOSKE_ALT y descarga solo el detalle del día.
- Suma los minutos por causa de paro para cada línea (L1-L4).
- Toma las 3 causas con más minutos y escribe el resumen en W6:W9.
- Ejemplo: `111 min Llenadora 1 L2 -Cua Sist. sensor/fotocel. falla`.
- No modifica REP2, CTT, Dia ni Ago 26.

**API pública:** ActualizarComentariosParosSAP, DiagnosticarSAPGUI.

### modConsolidarMB51.bas — Consolidador MB51

Consolida y limpia exportaciones de la transacción MB51 (movimientos de material).

- Procesa .XLS, .XLSX y .XLSM, incluyendo archivos de texto UTF-16 que SAP genera.
- Detecta encabezados por contenido (Material + Cantidad), no por posición.
- Regla de negativos configurable: ELIMINAR (recomendado para anulaciones) o ABSOLUTO.
- Omite filas de título repetidas y pies de página de la salida dinámica SAP.
- Genera tabla con formato, encabezados únicos, freeze panes y hoja Resumen con estadísticas y log de errores por archivo.
- Tolera celdas con errores de Excel (#N/A, #VALUE!).

**API pública:** Consolidar_MB51.

## Caso de estudio

Piloto de dos meses en planta de producción con 543 usuarios registrados, operación estable sin incidentes críticos. La validación detectó usuarios sin centro de costo asignado por RH y los bloqueó antes de llegar a SAP — el sistema no solo automatiza: protege la integridad de los datos.

## Requisitos

- SAP GUI for Windows con Scripting habilitado (cliente y servidor).
- Excel 2016 o superior.
- Outlook de escritorio (solo para correos masivos).
- SAP y Excel con el mismo nivel de privilegios.

## Notas

- Los datos de planta (centro, materiales, nóminas) han sido anonimizados.
- La ruta de usuario local en modConsolidarMB51.bas fue reemplazada por una ruta genérica.
- Los módulos de bajas EPP y correos no traen valores de planta, sistemas, correos ni clientes reales: todo va en CONFIG/Configuracion. Los CSV de `samples/` son ficticios.
- Para importar los `.bas`, en el editor de VBA usa Archivo → Importar archivo. Se guardan con fin de línea CRLF.

Desarrollado por José Luis Cruz Prieto.
