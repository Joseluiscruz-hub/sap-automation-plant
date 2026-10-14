# Libro de Bajas EPP (MIGO 201): layout

Módulo: [`macros/modMIGO_BajasEPP.bas`](../macros/modMIGO_BajasEPP.bas). El repositorio no trae el libro `.xlsm`. Créalo con estas hojas, importa el módulo (VBE → Archivo → Importar archivo) y guarda como `.xlsm`.

Los archivos de [`samples/bajas-epp/`](../samples/bajas-epp/) tienen datos **ficticios** para cada hoja.

## Hojas

| Hoja | Quién la llena | Contenido |
|---|---|---|
| `PEGAR_APP` | Usuario | Export de la app de EPP pegado tal cual, con encabezados en la fila 1 |
| `BAJAS_SAP` | Macro (`Preparar_Bajas_Desde_App`) | Layout SAP en **valores** (sin fórmulas) + estado por fila |
| `CONFIG` | Usuario | Parámetros (columna A = parámetro, B = valor) |
| `CATALOGOS` | Usuario (opcional) | `Categoria` → `UME` |
| `LOG_SAP` | Macro | Bitácora de cada corrida (se crea sola) |
| `DIAG_MIGO` | Macro (`Diagnosticar_Layout_MIGO`) | Columnas técnicas de la tabla de posiciones de MIGO |

## PEGAR_APP (fila 1 = encabezados; se buscan por nombre, el orden no importa)

| Encabezado | Obligatorio | Uso |
|---|---|---|
| `Folio` | sí | Identificador de la entrega. Con `Material` forma la llave de la fila |
| `Fecha` | sí | → `FECHA_DOC` (dd.mm.aaaa) |
| `Centro costo` | sí | → `CECO` (mayúsculas) |
| `Material` | sí | → `MATERIAL` (si viene vacío se usa `SKU`) |
| `Cantidad` | sí | → `CANTIDAD` |
| `Motivo` | sí | → `MOTIVO` y, según CONFIG, `DESTINATARIO` / `TEXTO` |
| `Nomina`, `Nombre` | no | Trazabilidad interna; solo viajan a SAP si CONFIG lo pide |
| `Area consumo`, `SKU`, `Descripcion`, `Categoria`, `Talla`, `Origen`, `Estado`, `Periodo`, `Hora` | no | Informativos u `OBSERVACIONES` |

## BAJAS_SAP (la escribe la macro)

`FOLIO_APP | FECHA_DOC | CMV | MATERIAL | DESCRIPCION_MATERIAL | CANTIDAD | UME | CECO | DESTINATARIO | TEXTO | REFERENCIA | MOTIVO | NUMERO_EMPLEADO | NOMBRE | AREA_LINEA | STATUS_SAP | MENSAJE_SAP | FECHA_PROCESO | DOC_MATERIAL_SAP | OBSERVACIONES | VALIDACION | DETALLE_VALIDACION | LOTE`

- Cada fila con datos de `PEGAR_APP` genera exactamente una fila. Si los conteos no cuadran o hay fórmulas en la hoja, **Validar y Procesar se detienen**.
- Al volver a preparar, el estado SAP (`STATUS_SAP`, `MENSAJE_SAP`, `FECHA_PROCESO`, `DOC_MATERIAL_SAP`, `LOTE`) se conserva por `FOLIO|MATERIAL`.
- Validaciones: folio, material (≤ 40), cantidad > 0, CeCo (≤ 10), destinatario (≤ 12, campo WEMPF), texto (≤ 50) y folio+material duplicado.

### Estados (`STATUS_SAP`)

| Estado | Significado | ¿Se vuelve a procesar? |
|---|---|---|
| `PENDIENTE` | Nunca procesada | sí |
| `SIMULADO` | Pasó por modo SIMULACION | sí |
| `CAPTURADO` | Capturada en MIGO; aún no termina la verificación del lote | sí |
| `VERIFICADO_SIN_CONTABILIZAR` | SAP respondió OK a Verificar; **no está contabilizada** | sí (con aviso) |
| `REVISAR` | Sin mensaje de SAP, ventana emergente, aviso o error en la verificación del lote | sí |
| `ERROR` | SAP rechazó la línea; se retiró de MIGO | sí |
| `CONTABILIZADO` | Registrada con `Registrar_Contabilizacion_Manual` (tiene documento) | **no** |

## CONFIG

| Parámetro | Ejemplo | Descripción |
|---|---|---|
| `MODO` | `SIMULACION` | `SIMULACION` = no toca SAP (dry-run). `VERIFICAR` = captura y verifica. No existe modo contabilizar |
| `SAP_SISTEMA` | `<SID>` | ID de sistema (p. ej. QAS o PRD). Debe coincidir con la sesión abierta |
| `SAP_MANDANTE` | `<MANDANTE>` | Mandante esperado |
| `SAP_TIMEOUT_SEG` | `60` | Espera máxima por respuesta de SAP |
| `CENTRO` | `<CENTRO>` | Centro (WERKS). `*` = no capturar centro (usar el default de SAP) |
| `ALMACEN` | `<ALMACEN>` | Almacén (LGORT) |
| `CMV` | `201` | Clase de movimiento |
| `MIGO_ACCION` | `A07` | Operación de MIGO (salida de mercancías) |
| `MIGO_REFERENCIA` | `R10` | Documento de referencia (Otros) |
| `MIGO_BOTON_VERIFICAR` | `wnd[0]/tbar[1]/btn[7]` | ID del botón Verificar (grabarlo con Script Recording si difiere) |
| `MIGO_ID_ACCION`, `MIGO_ID_REFERENCIA`, `MIGO_ID_CMV` | *(vacío)* | IDs alternos de la cabecera si `FindByName` no los encuentra |
| `DESTINATARIO_DESDE` | `MOTIVO` | `MOTIVO`, `NOMINA`, `NOMINA_NOMBRE`, `FOLIO` o `MOTIVO_FOLIO` |
| `TEXTO_DESDE` | `MOTIVO_FOLIO` | Igual que arriba |
| `REFERENCIA_DESDE` | `FOLIO` | Igual que arriba |
| `UME_DEFAULT` | `PZA` | Si la categoría no está en CATALOGOS |

Los valores entre `<…>` son marcadores. Con un marcador, la macro se niega a correr en modo VERIFICAR.

## Flujo recomendado

1. Pegar el export en `PEGAR_APP` → `Preparar_Bajas_Desde_App`.
2. `Validar_Bajas_EPP` y corregir las filas `REVISAR`.
3. `MODO = SIMULACION` → `Procesar_Bajas_EPP_MIGO` (no se conecta a SAP).
4. `Probar_Conexion_SAP` para confirmar sistema y mandante.
5. `MODO = VERIFICAR` → `Procesar_Bajas_EPP_MIGO`. Captura, verifica y deja MIGO abierta.
6. Revisar en pantalla y **contabilizar a mano**.
7. `Registrar_Contabilizacion_Manual` con el `LOTE` y el número de documento.
