# DATOS Y ALARMAS — Esquema Normalizado Nova

> Este documento define el esquema canónico al que se convierten todas las
> exportaciones de plataforma antes de ser escritas en las hojas MAESTRO-*.
> La app **nunca** lee hojas crudas; solo lee pestañas DASHBOARD-* ya calculadas.

---

## 1. MAESTRO-PEDIDOS (26 columnas)

| # | Nombre de columna     | Tipo    | Fuente / Descripción                                      |
|---|----------------------|---------|-----------------------------------------------------------|
| A | ID_ORDEN             | Texto   | ID único de la plataforma origen                          |
| B | FECHA_ORDEN          | Fecha   | Fecha de creación del pedido (YYYY-MM-DD)                 |
| C | MES                  | Texto   | Mes de la orden (YYYY-MM)                                 |
| D | PLATAFORMA           | Texto   | dropi_ec · dropi_gt · effi · mastershop · shopify · iris  |
| E | NOMBRE_CLIENTE       | Texto   | Nombre completo normalizado (trim + title case)           |
| F | TELEFONO             | Texto   | 10 dígitos, sin prefijo de país                           |
| G | CIUDAD               | Texto   | Ciudad destino normalizada                                |
| H | ZONA                 | Texto   | Departamento / provincia                                  |
| I | TRANSPORTADORA       | Texto   | Nombre de la transportadora                               |
| J | GUIA                 | Texto   | Número de guía                                            |
| K | PRODUCTO             | Texto   | Nombre del producto principal                             |
| L | CANTIDAD             | Número  | Unidades del pedido                                       |
| M | VALOR_ORDEN          | Número  | Valor bruto de la orden (moneda original)                 |
| N | GANANCIA_NETA        | Número  | Ganancia reportada por plataforma (puede ser 0)           |
| O | COSTO_ENVIO          | Número  | Costo de envío declarado                                  |
| P | ESTATUS_PLATAFORMA   | Texto   | Estatus tal como viene de la plataforma                   |
| Q | ESTATUS_NORM         | Texto   | Estatus normalizado (ver tabla MAPEO-ESTATUS)             |
| R | ESTADO_ENTREGA       | Texto   | entregado · devuelto · en_camino · pendiente · cancelado  |
| S | NOVEDAD_TIPO         | Texto   | Tipo de novedad logística (ver NOVEDADES-LOGÍSTICAS)      |
| T | NOVEDAD_DETALLE      | Texto   | Descripción de la novedad                                 |
| U | SOLUCION             | Texto   | Acción tomada por gestora                                 |
| V | GESTORA_ASIGNADA     | Texto   | Nombre de la gestora responsable                          |
| W | INTENTOS_CONTACTO    | Número  | Número de intentos de contacto registrados                |
| X | FECHA_ENTREGA        | Fecha   | Fecha real de entrega (YYYY-MM-DD)                        |
| Y | ULTIMA_ACTUALIZACION | Fecha   | Timestamp de la última actualización (YYYY-MM-DD HH:MM)   |
| Z | NOTAS_INTERNAS       | Texto   | Campo libre para gestora (motor NUNCA lo sobreescribe)    |

---

## 2. MAESTRO-CARTERA (15 columnas)

| # | Nombre de columna   | Tipo    | Descripción                                          |
|---|---------------------|---------|------------------------------------------------------|
| A | TELEFONO            | Texto   | Llave primaria — 10 dígitos sin prefijo              |
| B | NOMBRE_CLIENTE      | Texto   | Último nombre conocido                               |
| C | CIUDAD              | Texto   | Ciudad del cliente                                   |
| D | ZONA                | Texto   | Departamento / provincia                             |
| E | ESTATUS_ACTUAL      | Texto   | activo · dudoso · moroso · bloqueado · ok             |
| F | ORDENES_ACTIVAS     | Número  | Pedidos en curso                                     |
| G | ORDENES_ENTREGADAS  | Número  | Pedidos entregados exitosamente                      |
| H | ORDENES_DEVUELTAS   | Número  | Pedidos devueltos                                    |
| I | TOTAL_FACTURADO     | Número  | Suma de VALOR_ORDEN de todos sus pedidos             |
| J | TOTAL_COBRADO       | Número  | Suma efectivamente cobrada                           |
| K | SALDO_PENDIENTE     | Número  | Diferencia facturado − cobrado                       |
| L | DIAS_SIN_CONTACTO   | Número  | Días desde ULTIMO_CONTACTO hasta hoy                 |
| M | ULTIMO_CONTACTO     | Fecha   | Fecha del último intento de contacto (YYYY-MM-DD)    |
| N | GESTORA_ASIGNADA    | Texto   | Gestora responsable                                  |
| O | NOTAS               | Texto   | Campo libre (motor NUNCA lo sobreescribe)            |

---

## 3. MAESTRO-PAUTA (14 columnas)

| # | Nombre de columna   | Tipo    | Descripción                                                |
|---|---------------------|---------|------------------------------------------------------------|
| A | FECHA_INICIO        | Fecha   | Inicio del período reportado (YYYY-MM-DD)                  |
| B | FECHA_FIN           | Fecha   | Fin del período reportado (YYYY-MM-DD)                     |
| C | MES                 | Texto   | Mes (YYYY-MM) derivado de FECHA_INICIO                     |
| D | PLATAFORMA          | Texto   | meta · tiktok                                              |
| E | CAMPAÑA             | Texto   | Nombre de la campaña                                       |
| F | CONJUNTO            | Texto   | Nombre del conjunto de anuncios / Ad Set                   |
| G | PRESUPUESTO         | Número  | Presupuesto del conjunto (moneda original)                 |
| H | GASTO               | Número  | Importe gastado en el período                              |
| I | IMPRESIONES         | Número  | Total de impresiones                                       |
| J | CLICS               | Número  | Clics en enlace                                            |
| K | RESULTADOS          | Número  | Conversiones / compras reportadas                          |
| L | CPR                 | Número  | Costo por resultado                                        |
| M | ROAS                | Número  | Retorno sobre gasto publicitario                           |
| N | VALOR_CONVERSION    | Número  | Valor total de conversiones                                |

---

## 4. GASTOS (8 columnas)

| # | Nombre de columna | Tipo    | Descripción                                         |
|---|------------------|---------|-----------------------------------------------------|
| A | FECHA            | Fecha   | Fecha del gasto (YYYY-MM-DD)                        |
| B | MES              | Texto   | Mes (YYYY-MM)                                       |
| C | CATEGORIA        | Texto   | pauta · operativo · nomina · envios · otros         |
| D | DESCRIPCION      | Texto   | Descripción libre del gasto                         |
| E | MONTO            | Número  | Valor del gasto                                     |
| F | MONEDA           | Texto   | COP · USD · GTQ                                     |
| G | RESPONSABLE      | Texto   | Quién registró el gasto                             |
| H | COMPROBANTE      | Texto   | Número o URL del comprobante                        |

---

## 5. DASHBOARD-DUENO (solo lectura para la app)

Esta pestaña es **la única que la app web lee**, a través del proxy Netlify.
Se recalcula cada vez que el motor corre `procesarTodo()`.

### Bloque HOY (fila 2–20)
| Celda | Clave       | Descripción                                      |
|-------|-------------|--------------------------------------------------|
| A2    | fecha_hoy   | Fecha actual YYYY-MM-DD                          |
| A3    | pedidos_hoy | Conteo de pedidos con FECHA_ORDEN = hoy          |
| A4    | ventas_hoy  | Suma de VALOR_ORDEN de pedidos de hoy            |
| A5    | ganancia_hoy| Suma de GANANCIA_NETA de pedidos de hoy          |
| A6    | novedades_hoy | Conteo de pedidos con NOVEDAD_TIPO ≠ ""        |
| A7    | tasa_entrega_mes | % entregados del mes actual                 |

### Bloque MES (fila 22–50)
| Celda | Clave           | Descripción                                  |
|-------|-----------------|----------------------------------------------|
| A22   | mes_actual      | YYYY-MM                                      |
| A23   | pedidos_mes     | Total pedidos del mes                        |
| A24   | ventas_mes      | Suma VALOR_ORDEN del mes                     |
| A25   | ganancia_mes    | Suma GANANCIA_NETA del mes                   |
| A26   | devueltos_mes   | Conteo ESTADO_ENTREGA = 'devuelto'           |
| A27   | cancelados_mes  | Conteo ESTADO_ENTREGA = 'cancelado'          |
| A28   | gasto_pauta_mes | Suma GASTO de MAESTRO-PAUTA del mes          |
| A29   | roas_mes        | VALOR_CONVERSION / GASTO del mes             |

### Bloque PEDIDOS ACTIVOS (fila 52+)
Tabla con columnas: ID_ORDEN, NOMBRE_CLIENTE, TELEFONO, CIUDAD, GUIA,
ESTATUS_NORM, NOVEDAD_TIPO, GESTORA_ASIGNADA — solo pedidos en estado activo.

---

## 6. ALARMAS Y UMBRALES

| Alarma                  | Condición                                   | Nivel    |
|-------------------------|---------------------------------------------|----------|
| ROAS bajo              | ROAS < 1.5                                  | 🔴 CRITICA |
| Gasto sin retorno       | GASTO > 0 y RESULTADOS = 0                  | 🔴 CRITICA |
| Tasa de entrega baja   | entregados/total_mes < 0.60                 | 🔴 CRITICA |
| Novedades sin resolver  | NOVEDAD_TIPO ≠ "" y SOLUCION = "" > 3 días | 🟡 ALERTA  |
| Cartera morosa creciendo| clientes morosos > 5% del total             | 🟡 ALERTA  |
| Sin ventas hoy          | pedidos_hoy = 0 a las 14:00                 | 🟡 ALERTA  |
| ROAS aceptable          | ROAS 1.5 – 2.5                              | 🔵 INFO    |
| ROAS óptimo             | ROAS > 2.5                                  | 🟢 OK      |

---

## 7. MAPEO-PLATAFORMA — Estructura de la pestaña

Cada columna de esta pestaña corresponde a una plataforma.
La fila 1 contiene el nombre de la plataforma (`dropi_ec`, `effi`, etc.).
Las filas 2–N listan los headers que esa plataforma usa en su export.
El motor detecta la plataforma del CSV/XLSX pegado en IMP-ORDENES
comparando sus headers reales contra cada columna — gana la que más matches tenga (mínimo 4).

### Plataformas soportadas

| Plataforma   | Headers clave de identificación                            |
|--------------|------------------------------------------------------------|
| dropi_ec     | ID, FECHA, NOMBRE CLIENTE, TELÉFONO, NÚMERO GUIA, ESTATUS |
| dropi_gt     | Igual a dropi_ec + GANANCIA distinta (o config manual)    |
| effi         | Reporte ID Pedido, N° de Guía, Fecha del Pedido, Cliente   |
| mastershop   | Número de Orden, Fecha de Orden, Nombre Completo           |
| shopify      | Name, Email, Financial Status, Fulfillment Status          |
| meta_camp    | Nombre de la campaña, Importe gastado (COP), ROAS          |
| meta_adset   | Nombre del conjunto de anuncios, Importe gastado (COP)     |
| tiktok       | Campaign Name, Ad Group Name, Cost, Impressions            |
| iris         | Llamada ID, Fecha Llamada, Agente, Duración, Resultado      |

---

## 8. MAPEO-ESTATUS — Normalización de estados

| Estatus plataforma (ejemplos)                        | ESTATUS_NORM  | ESTADO_ENTREGA |
|------------------------------------------------------|---------------|----------------|
| ENTREGADO, Entregado, Delivered                      | entregado     | entregado      |
| EN CAMINO, In Transit, En tránsito                   | en_camino     | en_camino      |
| DEVUELTO, Returned, Devolución en proceso            | devuelto      | devuelto       |
| CANCELADO, Cancelled, Anulado                        | cancelado     | cancelado      |
| PENDIENTE, Pending, En preparación                   | pendiente     | pendiente      |
| NOVEDAD, Con Novedad, Exception                      | novedad       | en_camino      |
| EN BODEGA, Retenido en bodega                        | en_bodega     | en_camino      |

---

## 9. NOVEDADES LOGÍSTICAS — Máquina de estados

```
LOGISTIC_EXCEPTION
  └─ NOVELTY_DETECTED
       ├─ NO_CONTACT          → RESOLUTION_IN_PROGRESS → NOVELTY_RESOLVED
       ├─ BAD_ADDRESS         → RESOLUTION_IN_PROGRESS → NOVELTY_RESOLVED
       ├─ SCHEDULE_MISMATCH   → RESOLUTION_IN_PROGRESS → NOVELTY_RESOLVED
       ├─ RURAL_ZONE          → RESOLUTION_IN_PROGRESS → NOVELTY_RESOLVED
       └─ CANCELLED_BY_CARRIER→ RESOLUTION_IN_PROGRESS → NOVELTY_RESOLVED
```

Los tipos de novedad se mapean así desde las plataformas:
- Dropi `NOVEDAD` column → NOVEDAD_TIPO
- Effi `Cód. Novedad` + `Novedad` → NOVEDAD_TIPO + NOVEDAD_DETALLE
- Dropi novedades sheet → merge por GUIA

---

*Generado automáticamente por el motor Nova. No editar manualmente.*
