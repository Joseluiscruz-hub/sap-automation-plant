# AssetGuard — Automatización SAP para planta industrial

Sistema de automatización SAP GUI Scripting en VBA, desarrollado por **José Luis Cruz Prieto**.

Automatiza la operación diaria de una planta de producción: eficiencia, paros de línea, producción real, plan de producción (PDP) y bajas de EPP en MIGO. Todo corre desde Excel con una sola macro.

## Estructura

```
sap-automation-plant/
├── README.md
└── macros/
    ├── modReporteSAP_V6_5.bas        ← reporte de eficiencia completo (4,678 líneas)
    ├── modMIGO_BajasEPP.bas          ← bajas de EPP en MIGO (901 líneas)
    ├── modComentariosParosSAP_v2.bas ← comentarios TOP 3 de paros (791 líneas)
    └── modConsolidarMB51.bas         ← consolidador de exportaciones MB51 (717 líneas)
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

### modMIGO_BajasEPP.bas — Bajas de EPP en MIGO

Captura en lote las bajas de equipo de protección personal en MIGO (movimiento 201, salida de mercancías a centro de costo).

- Lee la hoja BAJAS_SAP: la plataforma exporta CSV, se pega en PEGAR_APP y las fórmulas llenan BAJAS_SAP.
- Valida material, cantidad, centro de costo, destinatario y texto antes de procesar.
- Abre MIGO **una sola vez**, captura todas las filas LISTO y verifica al final.
- Busca columnas por nombre técnico (GOITEM-MATNR, GOITEM-KOSTL...) en lugar de posiciones fijas.
- `Normalizar_Layout_MIGO` reordena el layout si faltan columnas.
- Escribe STATUS_SAP, MENSAJE_SAP, FECHA_PROCESO y DOC_MATERIAL_SAP por fila, y registra todo en LOG_SAP.
- Modo seguro: verifica, no contabiliza.

**API pública:** Validar_Bajas_EPP, Procesar_Bajas_EPP_MIGO, Normalizar_Layout_MIGO, Probar_Conexion_SAP, Diagnosticar_Layout_MIGO.

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
- SAP y Excel con el mismo nivel de privilegios.

## Notas

- Los datos de planta (centro, materiales, nóminas) han sido anonimizados.
- La ruta de usuario local en modConsolidarMB51.bas fue reemplazada por una ruta genérica.

Desarrollado por José Luis Cruz Prieto.
