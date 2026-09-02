# Datos y alarmas — Familia Nova

Este documento cubre **de dónde salen los datos** y **cuándo cada panel debe avisar algo**.

---

## Parte 1 · Hojas como base de datos provisional

### La regla cardinal

> **La aplicación nunca lee la hoja cruda. Lee una capa intermedia.**

Las exportaciones cambian de formato sin avisar. Con una capa intermedia, un cambio de formato rompe **una** función de importación, no seis pantallas.

### Arquitectura bidireccional

```
Archivos subidos a la plataforma (CSV / XLSX)
        ↓  Netlify function /import — una función por fuente
Hojas normalizadas (Pedidos, Novedades, Pauta, Inventario, Equipo…)
        ↓  Apps Script — recalcula DASHBOARD-* después de cada escritura
DASHBOARD-* (tabs de solo lectura para la app)
        ↓  Netlify function /read — proxy seguro
Paneles (dueño, admin, gestora, Nova Central, NovaSoul, novAcademy)
        ↓  escritura de vuelta desde los paneles
Hojas normalizadas
```

**NovaSoul y novAcademy no consumen exportaciones.** Toda su información la genera la propia plataforma: lo que la persona marca en su día, lo que el profesor da de alta, lo que el estudiante responde en la práctica. Para ellos el flujo es solo de escritura y lectura, sin importador. La única excepción es `Transitos`, que se precalcula a partir de los datos natales con una librería de efemérides — no se importa de un archivo, se genera.

### El flujo va en dos direcciones

```
Archivos subidos ──importador──→ ┌─────────────────┐
                                 │ Hojas           │ ──lectura──→ Paneles
Paneles ──────escritura────────→ │ normalizadas    │
                                 └─────────────────┘
                                         ↓
                                 Alarmas y resúmenes
```

**Regla:** un dato que la plataforma captura y no escribe en la hoja se pierde al recargar y las alarmas que dependen de él nunca disparan.

### Qué escribe cada panel

| Panel | Escribe en | Campos |
|---|---|---|
| **Gestora** | `Novedades` | `estado`, `nota`, `intentos`, `resuelta_en`, `gestora` |
| | `Pedidos` | `estado`, `guia`, `direccion` corregida, `fecha_promesa` |
| | `Equipo` | `casos_resueltos`, `ultima_conexion` |
| **Admin** | `Pedidos` | `gestora_asignada`, `estado` |
| | `Novedades` | autorización de devolución, `nota` |
| | `Equipo` | `casos_asignados`, cierre de turno |
| **Dueño** | `Inventario` | `stock`, `costo_unitario`, `precio` |
| | `Equipo` | `rol`, permisos, `nota_auditoria` |
| | `Parametros` | techo de CPA, costos fijos, umbrales |
| **Nova Central** | `Clientes` | `plan`, `tarifa`, `estado`, `ultimo_pago`, `fecha_corte` |
| | `Solicitudes` | aprobación o rechazo |
| | `Candidatas` | `estado` |
| **NovaSoul** | `Pendientes` | alta, marcado, fecha |
| | `Dias` | comidas, movimiento, cierre, puntos, racha |
| **novAcademy** | `Estudiantes` | alta, perfil, permisos, documentos |
| | `Progreso` | lección completada, nota, `ultima_conexion` |
| | `Ejercicios` | respuestas, tiempo, nota, envío a revisión |

### Reglas de escritura

- **Nunca sobrescribir columnas importadas** (`valor`, `costo_envio`, `cpm`). Son de solo lectura para la app.
- **Siempre escribir `actualizado_en` y `actualizado_por`.**
- **Registro en `Movimientos`**: `fecha · usuario · entidad · entidad_id · campo · valor_anterior · valor_nuevo`.
- **Antirrepetición**: tabla `Alertas_enviadas` con `alarma_id · entidad_id · enviado_en · resuelta_en`.
- **Escritura optimista** — la interfaz marca el cambio de inmediato y confirma cuando la hoja responde.
- **Cola local** si no hay conexión.

### Fuentes externas identificadas (solo Nova Empresarial y Nova Central)

| Fuente | Qué trae | Alimenta |
|---|---|---|
| Dropi export (CSV/XLSX) | Pedidos, estados, novedades | `Pedidos`, `Novedades` |
| Effi export (XLSX) | Pedidos, novedades logísticas | `Pedidos`, `Novedades` |
| Mastershop export | Pedidos | `Pedidos` |
| Shopify export | Pedidos (Nutrea EC) | `Pedidos` |
| Meta campaigns (CSV) | Gasto, resultados, ROAS por campaña | `Pauta` |
| Meta ad sets (CSV) | Gasto, resultados, ROAS por conjunto | `Pauta` |
| TikTok ads (CSV) | Gasto, impresiones, conversiones | `Pauta` |
| IRIS export | Llamadas de auditoría, agente, resultado | `Auditoría` |
| Cartera (XLSX) | Clientes con saldo pendiente | `Cartera` |
| Tarjeta (XLSX) | Transacciones de crédito | `Gastos` |

**Moneda:** Meta viene en COP; otras fuentes en USD. La importación normaliza a moneda base de la tienda con la tasa del día de la transacción.

---

## Parte 2 · Esquemas normalizados

Seis entidades vienen de exportaciones externas. Las de NovaSoul y novAcademy **no tienen fuente externa**: las llena por completo la propia plataforma.

### Nova Empresarial

**Pedidos**
`id · fecha · tienda · cliente · telefono · ciudad · direccion · producto · cantidad · valor · costo_producto · costo_envio · estado · transportadora · guia · intentos · gestora_asignada · fecha_promesa · fecha_entrega · actualizado_en · actualizado_por`

**Novedades**
`id · pedido_id · fecha · tipo · motivo · estado · gestora · nota · intentos · resuelta_en · actualizado_en`

**Pauta**
`fecha · tienda · plataforma · campaña · conjunto · gasto · impresiones · clics · resultados · cpm · cpa · roas`

**Inventario**
`sku · producto · tienda · stock · costo_unitario · precio · dias_cobertura · ultimo_conteo`

**Equipo**
`id · nombre · correo · rol · tienda · estado · casos_asignados · casos_resueltos · nota_auditoria · ultima_conexion`

**Cartera**
`telefono · nombre · ciudad · tienda · pedidos_activos · pedidos_entregados · pedidos_devueltos · total_facturado · total_cobrado · saldo_pendiente · dias_sin_contacto · ultimo_contacto · gestora_asignada · notas`

**Parametros** — los umbrales que define el dueño
`tienda · clave · valor · actualizado_en · actualizado_por`
Claves mínimas: `cpa_techo`, `costos_fijos_mes`, `entrega_minima_pct`, `corte_despacho_hora`, `dias_reposicion_proveedor`.

### Nova Central

**Clientes**
`id · empresa · pais · plan · tarifa · costo · estado · fecha_alta · fecha_corte · ultimo_pago · tickets_mes · usuarios · tiendas`

**Planes** — el armado de módulos y límites
`id · nombre · modulos · limite_usuarios · limite_tiendas · costo_calculado · precio_sugerido · tarifa_fijada`

**Solicitudes** — primer ingreso de cada persona, que Central aprueba
`id · cliente_id · nombre · correo · rol_pedido · estado · fecha_solicitud · resuelta_en · resuelta_por`

**Candidatas** — postulantes a gestora
`id · nombre · pais · experiencia · equipo_disponible · estado · fecha_postulacion · nota`

### NovaSoul

Todo lo escribe la persona; no hay importación.

**Usuarios**
`id · nombre · correo · fecha_nacimiento · hora_nacimiento · lugar_nacimiento · zona_horaria · acento · modo · idioma`

**Pendientes**
`id · usuario_id · texto · tipo · origen · fecha · hecho · hecho_en · plataforma_id`
`origen` distingue lo que puso la persona de lo que propuso la app. `plataforma_id` liga un pendiente a una de sus tiendas.

**Dias** — una fila por persona y por día
`usuario_id · fecha · comidas_marcadas · movimiento_hecho · puntos · cerrado · cerrado_en · perdonado`

**Recompensas**
`id · usuario_id · nombre · costo_puntos · canjeada · canjeada_en`

**Transitos** — el cielo de cada día, precalculado
`usuario_id · fecha · casa · tema · intensidad_pct · texto_transito · por_que · como_trabajarlo · el_otro_lado`
Los cuatro últimos son el contenido editorial de la tarjeta y su panel de detalle. Se generan a partir de los datos natales con una librería de efemérides; **no los escribe la persona**.

### novAcademy

**Estudiantes**
`id · nombre · correo · empresa · perfil · estado · fecha_alta · alta_por · acceso_vence · primer_ingreso`
`perfil` es uno de: `asistente`, `emprendedor`, `empresario`, `administradora`.

**Permisos** — lo que el profesor habilita a cada estudiante
`estudiante_id · ver_dinero · descargar_plantillas · ver_avance_equipo · emitir_certificado · entrar_practica`

**Documentos** — los archivos que se suben al dar de alta
`id · estudiante_id · nombre_archivo · tipo · peso · subido_en · subido_por`

**Progreso**
`estudiante_id · leccion_id · modulo · estado · completada_en · nota_quiz · intentos · ultima_conexion`

**Ejercicios** — cada intento en el ambiente de práctica
`id · estudiante_id · perfil · fecha · respuestas · aciertos · total · nota · segundos · enviado_revision · comentario_profesor · comentado_en`

**Certificados**
`id · estudiante_id · ruta · emitido_en · vence · url_pdf`

### Comunes (todos los productos)

**Movimientos** — registro de cambios, obligatorio
`fecha · usuario · entidad · entidad_id · campo · valor_anterior · valor_nuevo`

**Alertas_enviadas** — antirrepetición del correo
`alarma_id · entidad_id · enviado_en · resuelta_en`

### Reglas de higiene

- Una fila = un hecho. Sin celdas combinadas ni subtotales.
- Fechas en ISO (`YYYY-MM-DD`).
- Montos como número puro. Sin `$`, sin puntos de miles.
- Nunca borrar filas: un cancelado se marca `estado=cancelado`.
- Límite práctico: ~50 000 filas por hoja antes de migrar a BD real.

---

## Parte 3 · DASHBOARD-* — lo que la app lee

Cada pestaña `DASHBOARD-*` la calcula el Apps Script. La app solo lee estas.

### DASHBOARD-DUENO

| Bloque | Contenido |
|---|---|
| HOY | `pedidos_hoy`, `ventas_hoy`, `novedades_abiertas`, `tasa_entrega_7d` |
| MES | `pedidos_mes`, `ventas_mes`, `devueltos_mes`, `gasto_pauta_mes`, `roas_mes` |
| EQUILIBRIO | `punto_equilibrio_entregas`, `van_hoy`, `cubre_costos` |
| ACTIVOS | tabla de pedidos en_camino/pendiente/novedad |
| ALERTAS | lista de alarmas activas con nivel y acción |

### DASHBOARD-ADMIN

| Bloque | Contenido |
|---|---|
| HOY | `pedidos_urgentes`, `corte_despacho`, `novedades_sin_resolver` |
| EQUIPO | `gestoras_conectadas`, `carga_por_gestora` |
| COLA | pedidos por gestora con estado |

### DASHBOARD-GESTORA

| Bloque | Contenido |
|---|---|
| MI_TURNO | `mis_casos`, `casos_urgentes`, `turno_cierra_en` |
| MIS_CASOS | tabla filtrada por `gestora_asignada` |

---

## Parte 4 · Alarmas por panel

### Las seis alarmas obligatorias

| # | Se dispara cuando | Panel | Correo |
|---|---|---|---|
| 1 | Pedido > 3 días sin movimiento | Admin + gestora | Sí |
| 2 | Novedad > 24 h sin resolver | Admin + gestora | Sí |
| 3 | Cliente > 15 días sin pagar (Nova Central) | Nova Central | Sí |
| 4 | Gestora no entró en su turno | Admin | Sí |
| 5 | Tasa de entrega < 65% (7 días) | Dueño + admin | Sí |
| 6 | CPA supera el techo de Parametros | Dueño | Sí |

Todas son **nivel Crítica** (banner `#a8563a` sobre `#f6e3db`). Se evalúan cada 15 min y al abrir. El correo se manda **una sola vez por incidencia** (ver `Alertas_enviadas`).

### Niveles

| Nivel | Color texto / fondo | Comportamiento |
|---|---|---|
| **Crítica** | `#a8563a` / `#f6e3db` | Banner arriba, no descartable hasta resolver |
| **Atención** | `#8a6a13` / `#f6ecd0` | Tarjeta en resumen del día |
| **Informativa** | `#546b7a` / `#e4eaee` | Lista, sin interrumpir |

### Nova Central
Suscripción por vencer (10→5 días Atención, 3→1 Crítica) · Cliente vencido · Prueba por terminar · Cuenta que no arranca (>7 días sin conectar hoja) · Soporte que consume (3+ tickets/mes) · Cliente en riesgo · Margen en rojo · Ingreso concentrado (>25% un solo cliente) · Solicitud de acceso pendiente >24 h · Candidata sin revisar >5 días

### Nova Empresarial — dueño
Debajo del punto de equilibrio · Margen negativo en producto · CPA subió >40% vs. últimos 7 días · ROAS <2.0× tres días · Presupuesto >90% antes de las 18:00 · Quiebre de stock inminente · Inventario sin rotar (30 días, >20 und) · Conteo vencido >15 días · Devoluciones >10% · Gestora con nota <7.0 o >8 casos al cierre · Nómina >15% del mes anterior · Datos desactualizados >48 h

### Nova Empresarial — admin
Pedidos urgentes · Corte cerca (<1 h) · Novedad repetida (2+ novedades mismo pedido) · Patrón de novedades (3+ del mismo motivo/transportadora/ciudad) · Carga despareja (una gestora con el doble) · Gestora sin conectarse · Caso estancado >48 h · Devolución sin autorizar >24 h · Turno cerrado con pendientes

### Nova Empresarial — gestora
Casos con hora de corte · Cliente esperando >2 h · Cliente insatisfecho · Tercer intento de entrega · Gestión sin registrar · Turno por terminar (30 min) · Auditoría publicada

### NovaSoul — recordatorios (tono de acompañamiento, no de urgencia)
Lo primero del día · Pendiente con fecha hoy · Pendiente vencido · Comida sin marcar · Movimiento atrasado · Racha en riesgo · Cierre del día (21:00) · Tema de la semana en su pico

### novAcademy — profesor
Estudiante sin entrar 3 días · Estudiante estancado 7 días · Acceso por vencer <24 h · Quiz reprobado dos veces · Ejercicio enviado a revisión · Ruta terminada (certificado listo)

### novAcademy — estudiante
Lección pendiente >3 días · Módulo desbloqueado · Comentario de la profesora · Certificado disponible

---

## Parte 5 · Implementación de alarmas

```javascript
evaluarAlertas(datos, rol) → [
  { id, nivel, titulo, detalle, accion: { texto, destino }, entidad: { tipo, id } }
]
```

- **Filtrado por rol al final**, no al principio.
- **Umbrales configurables** en `Parametros`, no en código.
- **Sin ruido**: máx. 3 críticas visibles; el resto agrupa en "y N más". Una descartada no vuelve el mismo día (salvo Crítica). Si el dato está desactualizado, se muestra la alarma de datos desactualizados en lugar de las demás.
- **Frecuencia**: al abrir + cada 15 min. Corte de despacho: cada 5 min en la hora previa.
- Las 6 obligatorias necesitan evaluación del lado del servidor aunque nadie abra el panel.
- **Un correo por incidencia, no un digest** para las críticas. El resumen diario agrupa las de Atención.

---

*Versión 2 — incluye esquemas completos para los 4 productos + Nova Central.*
