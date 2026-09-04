# Nova · Capa de datos

Apps Script que conecta las hojas de Google con las cinco pantallas de Nova.

## Los workbooks

| Workbook | ID | Qué guarda |
|---|---|---|
| `Nova_Empresarial_TEMPLATE` | `1MEzF8O2qDHuBTU2jsGER5Q9MZx8o5jexdSB0RL8-2iQ` | Plantilla que se copia por cada cliente nuevo |
| `Nova_Central` | `1IDfY-zoc5lyPJWgqLGWWk_1CvFeedD_mVuauTvQ6FWM` | Clientes, planes, solicitudes, candidatas |
| `Nova_Soul` | `1xY3v7Fv5KPsZfj-Y8Ud2jpo8LbJg3oQOuGl6yOgLf1M` | Datos personales, días, tránsitos, gastos |
| `Nova_Academy` | `1KgMBwFFdiPbE4M18tP91iFSgSgi4lTt7_L62_BOdCNo` | Estudiantes, progreso, ejercicios, certificados |

Carpeta: [Nova](https://drive.google.com/drive/folders/1lxpyEhj3dfdwgdpotL7e8G8gCd_EaB7o)

## Cómo correr el bootstrap

1. Entra a [script.google.com](https://script.google.com) → **Nuevo proyecto**
2. Nómbralo `Nova API`
3. Pega el contenido de `00-bootstrap.gs`
4. **Ejecutar → `bootstrapTodo()`**
5. Autoriza cuando lo pida (pide permiso sobre tus hojas, es normal)

Es idempotente: si lo corres dos veces no duplica nada.

Después, solo para el workbook que uses tú (no el template):

```
sembrarTiendasNutrea()
```

Eso deja Nutrea GT (GTQ, America/Guatemala) y Nutrea EC (USD, America/Guayaquil) listas.

## Reglas que no se rompen

- **Nova nunca lee `_Import_*`.** Esas pestañas guardan el export crudo de Meta/Dropi/Shopify. Solo los importadores las tocan. La app lee únicamente las pestañas normalizadas.
- **Nunca se sobrescribe una fila importada.** Los campos que vienen de una exportación (`valor`, `costo_envio`, `cpm`) son de solo lectura para la app. Lo que la app escribe vive en columnas propias.
- **Nunca se borran filas.** Un pedido cancelado se marca `estado = cancelado`. El histórico es lo que hace posibles las tendencias.
- **Fechas en ISO** (`AAAA-MM-DD`), **montos como número puro** — sin `$`, sin puntos de miles.
- **Toda escritura deja rastro** en `Movimientos` y actualiza `actualizado_en` / `actualizado_por`.

## Dos hojas que agregué al spec

`DATOS-Y-ALARMAS.md` usa `tienda` como columna en Pedidos, Pauta, Inventario, Equipo y
Parametros, pero nunca define dónde vive el catálogo. Y exige convertir moneda *"con la
tasa del día de la transacción, no la de hoy"* sin decir dónde se guardan las tasas.

- **`Tiendas`** — `id · nombre · marca · pais · sociedad · nit · moneda · zona_horaria · corte_despacho · estado`
  Resuelve el caso de dos tiendas de la misma marca en países distintos, y el de dos
  sociedades distintas en el mismo país. También es de donde salen las horas locales
  para evaluar las alarmas 1, 2 y 4.

- **`Tasas`** — `fecha · moneda_origen · moneda_destino · tasa`
  Sin esto los márgenes salen mal y las alarmas de dinero disparan en falso.

Y una columna: **`sheet_id` en `Clientes`**, que es lo que permite que Nova Central sepa
a qué workbook apuntar cuando abre un cliente.

## Modelo multi-tienda

Un workbook por cuenta Nova (por facturación). `tienda` discrimina las filas.

| Caso | Cómo se maneja |
|---|---|
| Misma marca, distintos países | Un workbook, dos filas en `Tiendas` con su moneda y zona horaria |
| Mismo país, distintas sociedades | Un workbook, dos filas con distinto `sociedad` y `nit` |
| Marcas distintas, misma dueña | Un workbook, columna `marca`. Se separan solo si la contabilidad legal exige que no se mezclen |

El equipo nunca abre el Drive — entra por Nova. Los permisos por rol y por tienda los
aplica la app, no Google Drive.

## Fuentes

Siete, en **tres** familias. Una tienda usa una sola plataforma de pedidos —
el soporte múltiple existe para que un cliente con Effi o Mastershop pueda
usar la app sin restricción, no para mezclar fuentes dentro de una tienda.

| Familia | Fuentes | Alimenta |
|---|---|---|
| Pedidos / fulfillment | Dropi · Effi · Mastershop · Shopify | `Pedidos`, `Novedades` |
| Pauta | Meta · TikTok | `Pauta` |
| Central telefónica | IRIS | `Llamadas` |

El mapeo de columnas es **declarativo**, en `FUENTES` dentro de `20-importadores.gs`.
Agregar una plataforma es agregar un bloque de alias, no escribir un importador nuevo.

### IRIS no es una plataforma de pedidos

Revisando `IRIS (1).csv` resultó ser el log de la central telefónica: 2.782 llamadas
salientes con `Sentido · Estado · Extensión · Número · Usuario · Grabación`. No trae
pedidos. Alimenta una entidad propia, **`Llamadas`**, y cruza con `Pedidos`
**por teléfono normalizado**, no por número de orden.

Eso convierte la normalización de teléfono en pieza central, no en detalle. Medido
en tus datos:

| Fuente | Formato real |
|---|---|
| Dropi EC | `984712695` — 9 dígitos, sin prefijo |
| Mastershop CO | `3138880827` y `+573202241205` — **mezclados en el mismo archivo** |
| IRIS | 10, 11, 12 y 13 dígitos según el país |

Por eso cada fila guarda el número tal como llegó (`telefono`) **y** su forma canónica
E.164 (`telefono_norm`), que es la única que se usa para cruzar. Está en `10-estados.gs`.

### Estado de verificación

| Fuente | Estado | Cotejado contra |
|---|---|---|
| Dropi | ✅ verificado | `Dropi-Pedidos-NutreaShop.xlsx` (Maestro EC, 569 pedidos) |
| Mastershop | ✅ verificado | `reporte-historial-de-pedidos-*.xlsx` (30 col) |
| Effi guías | ✅ verificado | `Reporte de Guías de transporte.xlsx` (57 col, 231 guías) |
| Effi novedades | ✅ verificado | `Reporte de Novedades de guías.xlsx` (41 col, 66 novedades) |
| Meta campañas | ✅ verificado | `CP-Nutrea-Ecuador-Campañas.csv` (35 col) |
| Meta facturación | ✅ verificado | `Resumen_Facturación.csv` |
| IRIS | ✅ verificado | `IRIS (1).csv` (2.782 llamadas) |
| Shopify | ⚠️ por verificar | falta cotejar `orders_export.csv` |
| TikTok | ⚠️ por verificar | falta un export real |

Para cerrar las que faltan: pega un export en su pestaña `_Import_*` y corre
`diagnosticar('shopify', 'ec')`. Devuelve qué columnas no encontró y cómo quedó la
primera fila normalizada.

### Effi no se parece a Mastershop

Había asumido que Effi usaba el formato de Mastershop. **No.** Effi es un sistema de
**guías de transporte**, no de pedidos: habla de `Remitente` y `Destinatario`, no de
`Cliente`; y emite **dos reportes separados** que se cruzan por número de guía.

| | Mastershop | Effi |
|---|---|---|
| Unidad | Pedido | Guía de transporte |
| Reportes | 1 (30 columnas) | 2 — guías (57) + novedades (41) |
| Cliente | `Cliente` | `Destinatario` |
| Documento | `Cédula` | `ID. destinatario` = `"CC: 3223665889"` |
| Novedad | flag `Presento Novedad` Si/No | reporte aparte con **código numérico** |

Effi trae dos ventajas que las otras no:

- **`Cód. Novedad`** — 701, 828, 699… El código es más estable que el texto: Effi puede
  reescribir la descripción, el número no cambia. Están mapeados los 13 códigos reales.
- **`Estado global guía inicial`** — 7 valores limpios, frente a los 12 sucios de
  `Estado guía inicial` que traen la ciudad pegada (`DEVUELTA DESDE TUNJA`,
  `ENTREGADA DIGITALIZADA EN MEDELLIN`). Se usa el global; el otro queda como detalle.
  Los sucios se resuelven por prefijo, porque la ciudad es dinámica y no se enumera.

### Meta emite dos reportes con formatos incompatibles

- **Campañas** — CSV normal, encabezado en la línea 1, fechas ISO. Trae **CPA y ROAS ya
  calculados** (`Coste por compra`, `ROAS de compras`), mejor usarlos que recalcularlos.
  Ojo: dice *"coste"*, no *"costo"* — el alias que tenía escrito no habría enganchado.
- **Facturación** — **10 líneas de metainformación antes del encabezado real**, montos
  con espacio de miles (`195 866`), fechas `D/M/AAAA`, y una fila de total al final.

Ese preámbulo rompía la detección de encabezado, que buscaba "la primera fila con 3+
celdas llenas" — la línea 2 (`Meta Platforms Ireland Limited, Merrion Road, Dublin 4…`)
tiene 6. Ahora una fuente puede declarar un **ancla de encabezado** y **filas a
descartar**, para que la fila de total no entre en las sumas.

## Estados canónicos

Cada plataforma nombra los estados distinto — Dropi dice `ENTREGADO`, Mastershop dice
`Entregada`, el courier dice `ENTREGADA DIGITALIZADA`. Sin una tabla canónica la alarma
de *"entrega bajo 65%"* cuenta mal según de qué fuente venga la fila.

`10-estados.gs` define diez estados y mapea los reales de cada plataforma:

```
pendiente · confirmado · en_bodega · en_transito · en_oficina
novedad · novedad_resuelta · entregado · devolucion · cancelado
```

Un estado que no cae en el mapa se guarda como `__sin_mapear__:<texto>` — visible en
la hoja, nunca descartado en silencio.

También agrupa los motivos de novedad reales del Maestro en seis causas
(`no_contacta`, `rechaza`, `direccion`, `dinero`, `ausente`, `logistica`). El
agrupamiento es lo que hace funcionar la alarma de patrón: tres novedades del mismo
grupo en la semana es un problema de proceso, no tres casos sueltos.

## Columnas que salieron de los datos reales

- **`solucion`** en `Novedades` — en tu Maestro ya existe como columna, con instrucciones
  al courier (*"dejar en oficina y llamar"*, *"volver a pasar"*). Es la salida de la
  novedad, no una dirección corregida. Reemplaza el `direccion_corregida` que había
  puesto antes; la dirección cargada del cliente no se toca.
- **`telefono_2` + `telefono_2_norm`** — el número alterno. El caso de *"escribió desde
  otro número"* o *"pide que lo llamen a este otro"*, que hoy se pierde dentro de la nota.
  `extraerTelefonos()` lo saca del texto libre.
- **`estado` + `estado_transportadora` + `estado_canonico`** — Mastershop trae dos estados
  distintos (el del negocio y el del courier) y no siempre coinciden. Se guardan los dos
  crudos más el canónico calculado.
- **`fuente` + `id_externo`** — para rastrear cualquier fila de vuelta a su plataforma.
- **`departamento`**, `cedula`, `correo`, `metodo_pago`, `bodega`, `razon_cancelacion` —
  vienen en el reporte de Mastershop y se estaban perdiendo.
- **`moneda_gasto` + `gasto_normalizado`** en `Pauta` — Meta factura en COP, TikTok en USD.

Y dos hojas: **`Fuentes`** (de donde sale el "3 fuentes" que muestra tu login) y
**`Llamadas`** (IRIS).

## Pendiente

- `10-api.gs` — endpoints `doGet`/`doPost` para leer y escribir desde las pantallas
- `30-alarmas.gs` — evaluador central + correo, con trigger de 15 minutos
- `40-provisionar.gs` — copia del template al crear un cliente desde Nova Central
- Verificar los alias de las 7 fuentes contra exports reales
