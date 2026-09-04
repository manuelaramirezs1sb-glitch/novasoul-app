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

## Pendiente

- `10-api.gs` — endpoints `doGet`/`doPost` para leer y escribir desde las pantallas
- `20-importadores.gs` — una función por fuente: Meta, Dropi, Shopify
- `30-alarmas.gs` — evaluador central + correo, con trigger de 15 minutos
- `40-provisionar.gs` — copia del template al crear un cliente desde Nova Central
