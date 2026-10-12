# Libro de Correos de notas de crédito: layout

Módulo: [`macros/modCorreosNotasCredito.bas`](../macros/modCorreosNotasCredito.bas). El repositorio no trae el libro `.xlsm`. Crea las hojas, importa el módulo y ejecuta `PrepararConfiguracionCorreos` una vez.

Los archivos de [`samples/correos-notas-credito/`](../samples/correos-notas-credito/) tienen datos **ficticios** (correos `@example.com`).

## Hojas

| Hoja | Contenido |
|---|---|
| `Resumen` | Una fila por distribuidor. El encabezado (en las primeras 50 filas) debe incluir `RUTA`, `DISTRIBUIDOR`, `NOMBRE DISTRIBUIDOR` y una columna de total que contenga `TOTAL` y `NOTAS` o `CREDITO` |
| `Detalle de partidas` | Partidas. El encabezado (en las primeras 30 filas) debe incluir `Cliente` y una columna con `Referencia`. Se envían las columnas desde `Cliente` hasta la última. Las filas se filtran por `Cliente` = distribuidor |
| `Correos` | Una fila = un correo (ver columnas) |
| `Configuracion` | Columna A = parámetro, B = valor (la crea `PrepararConfiguracionCorreos`) |
| `LOG_CORREOS` | Bitácora (se crea sola). Guarda conteos, nunca direcciones |

## Correos (fila 1 = encabezados)

| Col | Encabezado | Uso |
|---|---|---|
| A | `ID` | Consecutivo libre (ya **no** se usa para elegir distribuidor) |
| B | `Enviar?` | `SI` / `NO` |
| C | `Estado` | `Pendiente`, `Generado`, `Borrador`, `Enviado`, `Error` u `Omitido`. Las filas `Enviado` **no se reenvían**; para reenviar, borra el estado |
| D–F | `Para 1..3` | Destinatarios |
| G | `Para adicional` | Lista separada por `;` |
| H–Q | `CC 1..10` | Copias |
| R | `CC adicional` | Lista separada por `;` |
| S | `CCO` | Copia oculta |
| T | `Asunto` | Se llena solo si `Usar asunto automatico = SI` |
| U | `Mensaje` | Si trae texto, reemplaza al Texto 2 de Configuracion |
| V | `Adjuntos` | Rutas separadas por `\|` (relativas a `Ruta base de adjuntos`). Si falta uno, el correo se descarta |
| W, X | `Conteo Para`, `Conteo CC` | Los llena la validación |
| Y | `Validación` | Resultado de `ValidarFilasCorreos` |
| Z | `Fecha procesado` | |
| AA | `Log/Error` | |
| AB | `Distribuidor` | **Obligatorio.** Debe existir en `Resumen` |

## Configuracion

| Parámetro | Default | Nota |
|---|---|---|
| `Modo` | `REVISAR` | `REVISAR` (abre cada correo), `BORRADOR` (guarda en Borradores) o `ENVIAR`. **Al terminar regresa a REVISAR** |
| `Pausa entre envios (segundos)` | `2` | Solo en ENVIAR |
| `Usar firma de Outlook` | `SI` | |
| `Ruta base de adjuntos (opcional)` | *(vacío)* | |
| `Asunto base` | `Notas de crédito Manual` | Asunto = base + `; ` + nombre del distribuidor |
| `Usar asunto automatico` | `SI` | |
| `Entrega nota de credito` | `segunda` | Variable `{ENTREGA}` |
| `Texto 1 antes de caratula` / `Texto 2 entre caratula y detalle` / `Texto 3 despues de detalle` | | Variables: `{ENTREGA}`, `{CLIENTE}`, `{NOMBRE_CLIENTE}`, `{DISTRIBUIDOR}`, `{NOMBRE_DISTRIBUIDOR}`, `{REMITENTE}`. Salto de línea con Alt+Enter o `\n` |
| `Nombre remitente cierre` | *(vacío)* | `{REMITENTE}` |
| `Filas detalle correo (incluye encabezado)` | `0` | `0` = todas las partidas del distribuidor |
| `Solo distribuidores con total distinto de cero` | `SI` | Una fila con total 0 queda en Error y no se envía |
| `Dominios permitidos (opcional, separados por ;)` | *(vacío)* | Si se llena, rechaza correos de otros dominios |

## Flujo recomendado

1. `ValidarFilasCorreos`: revisa distribuidor, formato y dominios de correo, mínimos de Para/CC y total ≠ 0.
2. `GenerarOCrearCorreosMasivos` en `REVISAR` (o `BORRADOR`) y revisar algunos correos.
3. Cambiar a `ENVIAR` y volver a ejecutar. La macro pide confirmación, omite las filas `Enviado` y deja el modo en `REVISAR`.
