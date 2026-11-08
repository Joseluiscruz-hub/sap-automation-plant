# AssetGuard — Automatización SAP para planta industrial

Sistema de automatización SAP GUI Scripting en VBA, desarrollado por **José Luis Cruz Prieto**.

Automatiza la operación diaria de una planta de producción: eficiencia, paros de línea, producción real, plan de producción (PDP) y bajas de EPP en MIGO. Todo corre desde Excel con una sola macro.

## Estructura

```
sap-automation-plant/
├── README.md
├── macros/
│   ├── modReporteSAP_V6_5.bas
│   ├── modComentariosParosSAP_v2.bas
│   ├── modConciliacionRobusta.bas
│   ├── modConsolidarMB51.bas
│   ├── modConciliarTransferencias.bas
│   ├── modInstaladorConciliacion.bas
│   ├── modMIGO_BajasEPP.bas
│   ├── modCorreosNotasCredito.bas
│   ├── modConvertirSAP_EPP.bas
│   ├── modSAP_a_TXT_Kiosko_EPP.bas
│   ├── modConvertirSAP_EPP_CargaLimpia.bas
│   ├── modReclasificacionQA_SQ.bas      ← reclasificación QA→SQ (FI post-S/4HANA)
│   └── modCompensacionF32.bas           ← compensación masiva F-32
├── docs/
│   ├── layout-bajas-epp.md
│   └── layout-correos-notas-credito.md
└── samples/
```

**Valores a configurar** en los módulos de eficiencia y paros (constantes al inicio del módulo): `CENTRO = "<CENTRO>"`, las transacciones Z del reporte (`ZREP_EFICIENCIA`, `ZREP_PAROS`, `ZREP_PRODUCCION`) y sus variantes (`/VAR_EFIC`, `/VAR_PAROS`, `/VAR_PROD`). En conciliación se configuran los materiales `<MATERIAL_ENVASE>` y `<MATERIAL_CAJA>` de la hoja Configuración. En la cadena de conversión SAP a EPP se configuran `<CENTRO>`, `<PLANTA>` y `<VARIANTE>` en `ConvertirPlanta`. Son marcadores públicos: sustitúyelos por los valores reales de tu sistema.

## Módulos

### modReclasificacionQA_SQ.bas y modCompensacionF32.bas — Saneamiento FI post-S/4HANA

Versiones mejoradas del cockpit usado para regularizar la cartera detallista de México (sociedad F142, modelo DA) después de S/4HANA.

- **modReclasificacionQA_SQ.bas**: punto de entrada `Reclasificar_NCA` para generar layouts QA→SQ por ruta (158k partidas, 77.5k documentos SQ).
- **modCompensacionF32.bas**: `Compensarf32_VariosCte` para compensación masiva en F-32 (410k documentos RH/RT/X1, neto ≈ 0).

Manejo de errores, validaciones previas y menos dependencia de code names de hoja. La lógica de lotes y secuencia SAP se conserva.

### modReporteSAP_V6_5.bas — Reporte de eficiencia

Una sola macro (`V6_ActualizarReporteCompletoSAP`) ejecuta seis pasos:

1. **Eficiencia del día** — transacción Z de eficiencia (`TRANSACCION_REP2`) con su variante, modo P_ACU, solo fecha de corte.
2. **Eficiencia acumulada del mes** — misma transacción, del día 1 al corte.
3. **Detalle REP2** — P_DET: fecha de inicio real y % eficiencia acumulada.
4. **Producción real** — transacción Z de producción (`TRANSACCION_PRODUCCION`): cajas unitarias por día y línea → hoja Ago 26.
5. **PDP / Production Plan** — lee el archivo más reciente de `PDP\ENTRADA`, convierte CF a unidades de consumo y actualiza el plan.
6. **Comentarios de paros** — transacción Z de paros (`TRANSACCION`) con su variante: descarga el detalle del día, suma minutos por causa, toma el TOP 3 por línea y escribe REPORTE EFICIENCIA!W6:W9.

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

**API pública:** Preparar_Bajas_Desde_App, Validar_Bajas_EPP, Procesar_Bajas_EPP_MIGO, Registrar_Contabilizacion_Manual, Probar_Conexion_SAP, Diagnosticar_Layout_MIGO, Normalizar_Layout_MIGO.

**v3.1 — layout robusto integrado:** la variante multiusuario de bajas EPP se fusionó en este módulo y ya no existe como archivo aparte. `Normalizar_Layout_MIGO` reordena la tabla de posiciones al orden canónico (Material, Cantidad, Almacén, CECO, Texto, Destinatario) y, si faltan columnas, abre Configurar layout para agregarlas. Además `TablaItems` ahora busca la tabla por nombre y, si el nombre cambia según la versión de MIGO, la busca recursivamente por GOITEM; y `NombreColumna` revisa las primeras 3 filas visibles en vez de solo la fila 0. Todo lo demás de v3.0 sigue igual: modos SIMULACION/VERIFICAR, nunca contabiliza solo.

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
- Ejecuta la transacción Z de paros (`ZREP_PAROS`, configurable) y descarga solo el detalle del día.
- Suma los minutos por causa de paro para cada línea (L1-L4).
- Toma las 3 causas con más minutos y escribe el resumen en W6:W9.
- Ejemplo: `111 min Llenadora 1 L2 Sist. sensor/fotocel. falla`.
- No modifica REP2, CTT, Dia ni Ago 26.

**API pública:** ActualizarComentariosParosSAP, DiagnosticarSAPGUI.

### modConciliacionRobusta.bas — Conciliación de envases y cajas contra MB51

Una macro (`ConciliarTodo`) cruza los reportes de proveedor (hojas Proveedor Envase y Proveedor Caja) contra la exportación MB51.

- Detecta automáticamente la fila de encabezados de cada fuente y normaliza la clave (remisión / O.E.).
- Convierte cantidades con la regla de Configuración (factor y material SAP esperado por producto).
- Considera movimientos de reversa y marca duplicados en proveedor, en MB51 o en ambos.
- Clasifica cada registro (conciliado, pendiente SAP, diferencia de cantidad, fecha o material, duplicado) con tolerancias configurables.
- Genera Conc Envase, Conc Caja, Incidencias y un Dashboard, y registra cada ejecución en una bitácora oculta.

**API pública:** ConciliarTodo.

### modConsolidarMB51.bas — Consolidador MB51 (v3)

Consolida y limpia exportaciones de la transacción MB51 (movimientos de material).

- Procesa .XLS, .XLSX y .XLSM, incluyendo archivos de texto UTF-16 que SAP genera.
- Detecta encabezados por contenido (Material + Cantidad), no por posición.
- Regla de negativos configurable: ELIMINAR (recomendado para anulaciones) o ABSOLUTO.
- Omite filas de título repetidas y pies de página de la salida dinámica SAP.
- Genera tabla con formato, encabezados únicos, freeze panes y hoja Resumen con estadísticas y log de errores por archivo.
- Tolera celdas con errores de Excel (#N/A, #VALUE!).

**API pública:** Consolidar_MB51.

### modConciliarTransferencias.bas — Conciliación de transferencias / rutas

Cruza Liquidación de Ruta contra Detalle Cliente DA - Vendedor y concilia por importe con búsqueda dinámica de subconjuntos (con límite de estados para no explotar en combinaciones).

- Lee las hojas una sola vez en arreglos y trabaja con Currency/centavos para no perder por decimales.
- `ConciliarConMismoSignoOptimizado` concilia movimientos del mismo signo; `RestablecerDatos` y `GenerarEstadisticas` dejan el libro listo para la siguiente corrida.
- `ValidarDatosAntesProcesar` revisa los datos antes de tocar nada.

**API pública:** ConciliarTransferenciasOptimizado, ConciliarConMismoSignoOptimizado, RestablecerDatos, ValidarDatosAntesProcesar, GenerarEstadisticas.

**Setup:** `modInstaladorConciliacion.bas` (`Instalar_Conciliacion_Limpia`) importa el módulo, elimina los módulos viejos y reasigna los botones a las macros nuevas.

### modConvertirSAP_EPP.bas, modSAP_a_TXT_Kiosko_EPP.bas y modConvertirSAP_EPP_CargaLimpia.bas — Export SAP a EPP

Cadena de tres módulos que convierte un export de SAP a los tres formatos que usa la operación de EPP:

1. **Template EPP** (`Convertir_LibroSAP_A_TemplateEPP`): pasa el export al template EPP SAP, con catálogo de tallas.
2. **TXT para kiosko** (`Crear_TXT_Kiosko_EPP_Desde_SAP`): genera el TXT que lee el kiosko de entrega.
3. **Carga limpia** (`Convertir_LibroSAP_A_CargaLimpia_EPP`): genera el archivo de carga sin basura para el sistema.

La planta se normaliza en `ConvertirPlanta` con los marcadores `<CENTRO>`, `<PLANTA>` y `<VARIANTE>`: configúralos con tus valores antes de usarla.

**Nota — variante por distribuidor fusionada:** la variante de correos con asunto y carátula por distribuidor ya no existe como archivo aparte: su comportamiento ya estaba en este v3 (distribuidor obligatorio por fila, asunto, carátula y detalle por distribuidor). Lo único que tenía la variante y que aquí NO se conservó, a propósito, fue usar el ID consecutivo como ordinal para tomar el distribuidor 1, 2, 3... de Resumen: si cambiaba el orden de Resumen, el correo de un distribuidor podía llevar los datos de otro. Por eso el distribuidor sigue siendo obligatorio por fila y validado contra Resumen.


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
