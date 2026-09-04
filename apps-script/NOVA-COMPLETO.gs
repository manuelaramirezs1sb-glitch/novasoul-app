/**
 * ███ NOVA · CAPA DE DATOS ███
 * ─────────────────────────────────────────────────────────────
 * Archivo único. Pégalo completo en Código.gs y listo.
 *
 * ┌─ PARA INSTALAR EN UNA CUENTA NUEVA ────────────────────────┐
 * │                                                            │
 * │   1) Ejecutar → instalarNova()                             │
 * │        crea la carpeta, los 4 workbooks y todas las        │
 * │        pestañas. Guarda los IDs solo: no hay que copiar    │
 * │        ni pegar ningún ID a mano.                          │
 * │                                                            │
 * │   2) Ejecutar → crearNutrea()                              │
 * │        crea TU hoja de operación, copiando el template.    │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * IMPORTANTE: corre esto en la cuenta de Nova, no en una personal.
 * El Web App se ejecuta con los permisos de quien es dueño del script,
 * así que el script y las hojas tienen que vivir en la misma cuenta.
 *
 * Contiene, en este orden:
 *   1. Bootstrap ....... instalación y estructura de los 4 workbooks
 *   2. Estados ......... estados canónicos + teléfonos de 24 países
 *   3. Monedas ......... 35 monedas y conversión con la tasa del día
 *   4. Importadores .... convierte las pestañas _Import_* en filas limpias
 *   5. Auto-mapeo ...... detecta columnas de fuentes sin export de muestra
 *   6. Provisionar ..... crea la hoja de un cliente copiando el template
 *
 * ÚTILES:
 *   verInstalacion()          a qué hojas apunta el script en esta cuenta
 *   listarClientes()          qué clientes hay y a qué hoja apunta cada uno
 *   tasasFaltantes()          qué tasas faltan para poder convertir dinero
 *   proponerMapeo(fuente,td)  mapea una fuente nueva sin export de muestra
 *   diagnosticar(fuente,td)   revisa el mapeo de una fuente ya configurada
 *
 * Esquemas tomados de design_handoff_nova/DATOS-Y-ALARMAS.md y verificados
 * contra los exports reales de la carpeta QKF + NOVA.
 */



/* ═══════════════════════════════════════════════════════════════
   1 · INSTALACIÓN Y BOOTSTRAP
   ═══════════════════════════════════════════════════════════════ */


// ─── IDs de los workbooks ────────────────────────────────────
/**
 * Los IDs NO se escriben a mano. Se guardan en las propiedades del
 * script cuando corres instalarNova(), y de ahí los lee todo lo demás.
 *
 * Así el mismo código funciona en cualquier cuenta de Google sin editar
 * una sola línea: si mañana esto se mueve de un correo a otro, se corre
 * instalarNova() allá y listo.
 *
 * Los valores de abajo son solo el respaldo de la instalación original.
 */
const IDS_DEFAULT = {
  empresarial: '1MEzF8O2qDHuBTU2jsGER5Q9MZx8o5jexdSB0RL8-2iQ', // Nova_Empresarial_TEMPLATE
  central:     '1IDfY-zoc5lyPJWgqLGWWk_1CvFeedD_mVuauTvQ6FWM', // Nova_Central
  soul:        '1xY3v7Fv5KPsZfj-Y8Ud2jpo8LbJg3oQOuGl6yOgLf1M', // Nova_Soul
  academy:     '1KgMBwFFdiPbE4M18tP91iFSgSgi4lTt7_L62_BOdCNo', // Nova_Academy
};

function IDS_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  return {
    empresarial: p.ID_EMPRESARIAL || IDS_DEFAULT.empresarial,
    central:     p.ID_CENTRAL     || IDS_DEFAULT.central,
    soul:        p.ID_SOUL        || IDS_DEFAULT.soul,
    academy:     p.ID_ACADEMY     || IDS_DEFAULT.academy,
    carpeta:     p.ID_CARPETA     || '',
  };
}

/**
 * ★ INSTALACIÓN DESDE CERO ★
 *
 * Crea la carpeta Nova, los 4 workbooks, guarda sus IDs y construye
 * todas las pestañas. Es lo ÚNICO que hay que correr en una cuenta nueva.
 *
 * Corre esto en la cuenta de Nova — no en una personal. El Web App se
 * ejecuta con los permisos de quien es dueño del script, así que el
 * script y las hojas tienen que vivir en la misma cuenta.
 *
 * Es seguro correrlo dos veces: si ya hay IDs guardados, no crea nada
 * nuevo, solo completa las pestañas que falten.
 */
function instalarNova() {
  const props = PropertiesService.getScriptProperties();
  const ya = props.getProperties();
  const log = [];

  // 1. Carpeta
  let carpeta;
  if (ya.ID_CARPETA) {
    carpeta = DriveApp.getFolderById(ya.ID_CARPETA);
    log.push('Carpeta existente: ' + carpeta.getName());
  } else {
    carpeta = DriveApp.createFolder('Nova');
    props.setProperty('ID_CARPETA', carpeta.getId());
    log.push('Carpeta creada: Nova');
  }

  // 2. Los 4 workbooks
  const aCrear = [
    { clave: 'ID_EMPRESARIAL', nombre: 'Nova_Empresarial_TEMPLATE' },
    { clave: 'ID_CENTRAL',     nombre: 'Nova_Central' },
    { clave: 'ID_SOUL',        nombre: 'Nova_Soul' },
    { clave: 'ID_ACADEMY',     nombre: 'Nova_Academy' },
  ];

  aCrear.forEach(function (w) {
    if (ya[w.clave]) {
      log.push('Ya existía: ' + w.nombre);
      return;
    }
    const ss = SpreadsheetApp.create(w.nombre);
    // create() lo deja en la raíz del Drive; hay que moverlo a la carpeta
    DriveApp.getFileById(ss.getId()).moveTo(carpeta);
    props.setProperty(w.clave, ss.getId());
    log.push('Creado: ' + w.nombre);
  });

  // 3. Construir todas las pestañas
  log.push('');
  log.push(bootstrapTodo());
  log.push('');
  log.push('Carpeta: ' + carpeta.getUrl());
  log.push('');
  log.push('LISTO. Ahora corre crearNutrea() para crear tu operación.');

  const salida = log.join('\n');
  Logger.log(salida);
  return salida;
}

/** Muestra a qué hojas está apuntando el script en esta cuenta. */
function verInstalacion() {
  const i = IDS_();
  const msg = Object.keys(i).map(function (k) {
    return pad(k, 14) + (i[k] || '(sin configurar)');
  }).join('\n');
  Logger.log(msg);
  return msg;
}

// Las dos hojas obligatorias que van en TODOS los workbooks
const COMUNES = {
  Movimientos:      ['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo'],
  Alertas_enviadas: ['alarma_id','entidad_id','enviado_en','resuelta_en'],
};

// ─── ESQUEMAS ────────────────────────────────────────────────

const ESQUEMA_EMPRESARIAL = {
  // Configuración — Tiendas y Tasas no están en el spec original,
  // se agregan porque `tienda` se usa como columna en todas partes
  // y la conversión de moneda exige la tasa del día de la transacción.
  Tiendas: ['id','nombre','marca','pais','sociedad','nit','moneda',
            'zona_horaria','corte_despacho','estado'],
  Parametros: ['tienda','clave','valor','actualizado_en','actualizado_por'],
  Tasas: ['fecha','moneda_origen','moneda_destino','tasa'],

  // Catálogo de fuentes por tienda. La pantalla de login ya muestra
  // "Shopify + Dropi · USD · 3 fuentes" — esto es de donde sale ese conteo.
  Fuentes: ['tienda','fuente','tipo','cuenta','activa','ultima_importacion',
            'filas_ultima','notas'],

  // Operación
  // fuente + id_externo: sin esto no se puede deduplicar ni rastrear una fila
  //   de vuelta a la plataforma de donde salió (con 5 fuentes de pedidos es obligatorio).
  // estado_nova y direccion_corregida: columnas propias de la app. El spec exige
  //   "nunca sobrescribir una fila importada" — estado y direccion los manda la
  //   plataforma, lo que el equipo cambia vive aparte. La UI muestra el _nova si existe.
  // telefono_norm es la clave de cruce con las llamadas de IRIS: el mismo
  //   cliente aparece como 984712695 en Dropi y como 593984712695 en la central.
  // telefono_2 guarda el número alterno — el caso de "escribió desde otro
  //   número" o "pide que lo llamen a este otro", que hoy se pierde en la nota.
  // estado_transportadora existe porque Mastershop/Effi traen DOS estados:
  //   el del negocio y el del courier, y no siempre coinciden.
  Pedidos: ['id','fuente','id_externo','fecha','tienda','cliente','cedula','correo',
            'telefono','telefono_norm','telefono_2','telefono_2_norm',
            'ciudad','departamento','direccion','producto','sku','cantidad',
            'valor','costo_producto','costo_envio','metodo_pago','bodega',
            'estado','estado_transportadora','estado_canonico','transportadora','guia',
            'intentos','gestora_asignada','fecha_promesa','fecha_entrega',
            'razon_cancelacion','estado_nova','nota','ultimo_movimiento',
            'actualizado_en','actualizado_por'],

  // `solucion` es la instrucción que se le da al courier para resolver la
  // novedad ("dejar en oficina y llamar", "volver a pasar"). No cambia la
  // dirección cargada del cliente — es la salida de la novedad.
  // `grupo` agrupa el motivo (no_contacta / rechaza / direccion / dinero...)
  // y es lo que permite la alarma de patrón: 3 del mismo grupo en la semana
  // es un problema de proceso, no tres casos sueltos.
  Novedades: ['id','fuente','id_externo','pedido_id','fecha','tipo','motivo','grupo',
              'estado','gestora','solucion','nota','intentos','resuelta_en',
              'actualizado_en','actualizado_por'],

  // IRIS no es una plataforma de pedidos — es la central telefónica.
  // Cruza con Pedidos por telefono_norm, no por id de orden.
  Llamadas: ['id','fuente','id_externo','fecha_hora','tienda','sentido','estado',
             'extension','agente','telefono','telefono_norm','pedido_id',
             'seg_conversado','seg_espera','seg_total','campana','grabacion',
             'etiqueta','observacion'],
  Pauta: ['fecha','tienda','plataforma','cuenta','campana','conjunto','gasto',
          'moneda_gasto','gasto_normalizado','impresiones','clics','resultados','cpm','cpa'],
  Inventario: ['sku','producto','tienda','fuente','stock','costo_unitario','precio',
               'dias_cobertura','ultimo_conteo','actualizado_en','actualizado_por'],
  Equipo: ['id','nombre','correo','rol','tienda','estado','casos_asignados',
           'casos_resueltos','nota_auditoria','ultima_conexion'],
};

// Staging crudo. Nova NUNCA lee estas pestañas — solo los importadores.
const IMPORTS_EMPRESARIAL = [
  // pauta — Meta emite dos reportes con formatos distintos
  '_Import_Meta', '_Import_Meta_Facturacion', '_Import_TikTok',
  // pedidos / fulfillment — una tienda usa UNA sola de estas plataformas.
  // Effi emite dos reportes separados que se cruzan por número de guía.
  '_Import_Dropi', '_Import_Mastershop', '_Import_Shopify',
  '_Import_Effi_Guias', '_Import_Effi_Novedades',
  // central telefónica — alimenta Llamadas, no Pedidos
  '_Import_Iris',
];

const ESQUEMA_CENTRAL = {
  // sheet_id no está en el spec pero es indispensable: es lo que permite
  // que Central sepa a qué workbook apuntar cuando abre un cliente.
  Clientes: ['id','empresa','pais','plan','tarifa','costo','estado','fecha_alta',
             'fecha_corte','ultimo_pago','tickets_mes','usuarios','tiendas','sheet_id'],
  Planes: ['id','nombre','modulos','limite_usuarios','limite_tiendas',
           'costo_calculado','precio_sugerido','tarifa_fijada'],
  Solicitudes: ['id','cliente_id','nombre','correo','rol_pedido','estado',
                'fecha_solicitud','resuelta_en','resuelta_por'],
  Candidatas: ['id','nombre','pais','experiencia','equipo_disponible','estado',
               'fecha_postulacion','nota'],
};

const ESQUEMA_SOUL = {
  Usuarios: ['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
             'lugar_nacimiento','zona_horaria','acento','modo','idioma'],
  Pendientes: ['id','usuario_id','texto','tipo','origen','fecha','hecho',
               'hecho_en','plataforma_id'],
  Dias: ['usuario_id','fecha','comidas_marcadas','movimiento_hecho','puntos',
         'cerrado','cerrado_en','perdonado'],
  Recompensas: ['id','usuario_id','nombre','costo_puntos','canjeada','canjeada_en'],
  Transitos: ['usuario_id','fecha','casa','tema','intensidad_pct','texto_transito',
              'por_que','como_trabajarlo','el_otro_lado'],
  // Gastos personales — pendiente comprometido, NovaSoul no corre pauta
  Gastos: ['id','usuario_id','fecha','concepto','categoria','tipo','monto',
           'moneda','recurrente','nota'],
};

const ESQUEMA_ACADEMY = {
  Estudiantes: ['id','nombre','correo','empresa','perfil','estado','fecha_alta',
                'alta_por','acceso_vence','primer_ingreso'],
  Permisos: ['estudiante_id','ver_dinero','descargar_plantillas','ver_avance_equipo',
             'emitir_certificado','entrar_practica'],
  Documentos: ['id','estudiante_id','nombre_archivo','tipo','peso','subido_en','subido_por'],
  Progreso: ['estudiante_id','leccion_id','modulo','estado','completada_en',
             'nota_quiz','intentos','ultima_conexion'],
  Ejercicios: ['id','estudiante_id','perfil','fecha','respuestas','aciertos','total',
               'nota','segundos','enviado_revision','comentario_profesor','comentado_en'],
  Certificados: ['id','estudiante_id','ruta','emitido_en','vence','url_pdf'],
};

// ─── PARÁMETROS POR DEFECTO ──────────────────────────────────
// Los cinco umbrales que el spec fija por defecto. Solo cpa_techo
// lo define el dueño; los otros cuatro vienen con estos valores.
const PARAMETROS_DEFAULT = [
  ['*', 'cpa_techo',                 '', '', ''],
  ['*', 'costos_fijos_mes',          '', '', ''],
  ['*', 'entrega_minima_pct',        65, '', ''],
  ['*', 'corte_despacho_hora',  '16:00', '', ''],
  ['*', 'dias_reposicion_proveedor', 15, '', ''],
  ['*', 'pedido_sin_movimiento_dias', 3, '', ''],
  ['*', 'novedad_sin_resolver_horas',24, '', ''],
  ['*', 'cliente_sin_pagar_dias',    15, '', ''],
];

// ─── MOTOR ───────────────────────────────────────────────────

function bootstrapTodo() {
  const log = [];
  log.push(construir(IDS_().empresarial, 'Nova_Empresarial_TEMPLATE',
                     ESQUEMA_EMPRESARIAL, IMPORTS_EMPRESARIAL));
  log.push(construir(IDS_().central,  'Nova_Central', ESQUEMA_CENTRAL));
  log.push(construir(IDS_().soul,     'Nova_Soul',    ESQUEMA_SOUL));
  log.push(construir(IDS_().academy,  'Nova_Academy', ESQUEMA_ACADEMY));

  sembrarParametros();
  Logger.log(log.join('\n'));
  return log.join('\n');
}

function construir(fileId, nombre, esquema, importsCrudos) {
  const ss = SpreadsheetApp.openById(fileId);
  const creadas = [];

  // Entidades + las dos comunes obligatorias
  const todas = Object.assign({}, esquema, COMUNES);

  Object.keys(todas).forEach(function (tab) {
    if (crearTab(ss, tab, todas[tab])) creadas.push(tab);
  });

  // Staging crudo: sin encabezados fijos, el formato lo dicta la exportación
  (importsCrudos || []).forEach(function (tab) {
    if (!ss.getSheetByName(tab)) {
      const sh = ss.insertSheet(tab);
      sh.getRange(1, 1).setValue(
        '⚠ Pestaña de importación cruda. Nova no la lee. ' +
        'Pega aquí el export sin modificar; el importador la normaliza.'
      ).setFontColor('#a8563a').setFontWeight('bold');
      sh.setTabColor('#cccccc');
      creadas.push(tab);
    }
  });

  limpiarHojaPorDefecto(ss);
  return nombre + ': ' + (creadas.length ? creadas.join(', ') : 'sin cambios');
}

function crearTab(ss, nombre, encabezados) {
  if (ss.getSheetByName(nombre)) return false;

  const sh = ss.insertSheet(nombre);
  sh.getRange(1, 1, 1, encabezados.length)
    .setValues([encabezados])
    .setFontWeight('bold')
    .setBackground('#0b1824')
    .setFontColor('#c9a84c');
  sh.setFrozenRows(1);
  sh.autoResizeColumns(1, encabezados.length);

  // Sobra de columnas: se eliminan para que la hoja no arrastre peso muerto
  const sobran = sh.getMaxColumns() - encabezados.length;
  if (sobran > 0) sh.deleteColumns(encabezados.length + 1, sobran);

  return true;
}

/** Borra la "Hoja 1" / "Sheet1" vacía que Drive crea por defecto. */
function limpiarHojaPorDefecto(ss) {
  const basura = ['Sheet1', 'Hoja 1', 'Hoja1'];
  if (ss.getSheets().length <= 1) return;
  basura.forEach(function (n) {
    const sh = ss.getSheetByName(n);
    if (sh && sh.getLastRow() === 0) ss.deleteSheet(sh);
  });
}

function sembrarParametros() {
  const sh = SpreadsheetApp.openById(IDS_().empresarial).getSheetByName('Parametros');
  if (!sh || sh.getLastRow() > 1) return; // ya sembrado
  sh.getRange(2, 1, PARAMETROS_DEFAULT.length, 5).setValues(PARAMETROS_DEFAULT);
}

// La siembra de tiendas NO va aquí: vive en crearCliente(), porque la
// operación de Nova también es una copia del template. Si se sembrara
// aquí, cada cliente nuevo nacería con las tiendas de Nutrea adentro.


/* ═══════════════════════════════════════════════════════════════
   2 · ESTADOS CANÓNICOS Y TELÉFONOS
   ═══════════════════════════════════════════════════════════════ */


// ─── ESTADOS CANÓNICOS ───────────────────────────────────────
// Diez estados. Todo lo que llegue de cualquier plataforma cae en uno.

const ESTADOS = {
  PENDIENTE:        'pendiente',        // sin confirmar con el cliente
  CONFIRMADO:       'confirmado',       // guía generada, aún no despachado
  EN_BODEGA:        'en_bodega',        // en operación logística
  EN_TRANSITO:      'en_transito',      // en reparto
  EN_OFICINA:       'en_oficina',       // esperando retiro del cliente
  NOVEDAD:          'novedad',          // problema abierto
  NOVEDAD_RESUELTA: 'novedad_resuelta', // solución aplicada, sigue en ruta
  ENTREGADO:        'entregado',        // terminal, cuenta como venta
  DEVOLUCION:       'devolucion',       // terminal, no cuenta como venta
  CANCELADO:        'cancelado',        // terminal, no cuenta como venta
};

// Los tres estados que cierran el ciclo. Las tendencias y el % de
// entrega solo miran estos.
const ESTADOS_TERMINALES = [ESTADOS.ENTREGADO, ESTADOS.DEVOLUCION, ESTADOS.CANCELADO];

// ─── MAPEO POR PLATAFORMA ────────────────────────────────────
// Clave: el texto tal como llega (normalizado sin tildes ni may/min).
// Verificado contra los exports reales.

const MAPA_ESTADOS = {

  // Dropi / Ecuador — VERIFICADO (569 pedidos, 17 estados distintos)
  dropi: {
    'entregado':                          ESTADOS.ENTREGADO,
    'devolucion':                         ESTADOS.DEVOLUCION,
    'en proceso de devolucion':           ESTADOS.DEVOLUCION,
    'cancelado':                          ESTADOS.CANCELADO,
    'novedad':                            ESTADOS.NOVEDAD,
    'novedad solucionada':                ESTADOS.NOVEDAD_RESUELTA,
    'solucion aprobada':                  ESTADOS.NOVEDAD_RESUELTA,
    'pendiente':                          ESTADOS.PENDIENTE,
    'pendiente confirmacion':             ESTADOS.PENDIENTE,
    'guia_generada':                      ESTADOS.CONFIRMADO,
    'guia generada':                      ESTADOS.CONFIRMADO,
    'en bodega origen':                   ESTADOS.EN_BODEGA,
    'ingresando operativo a bodega':      ESTADOS.EN_BODEGA,
    'ingresando operativo a':             ESTADOS.EN_BODEGA,
    'en reparto':                         ESTADOS.EN_TRANSITO,
    'zona de entrega':                    ESTADOS.EN_TRANSITO,
    'en distribucion a cliente':          ESTADOS.EN_TRANSITO,
    'en ruta a concesion':                ESTADOS.EN_TRANSITO,
    'para retiro en agencia servientrega':ESTADOS.EN_OFICINA,
    'para retiro en agencia':             ESTADOS.EN_OFICINA,
  },

  // Mastershop / Effi — Colombia. VERIFICADO.
  // Ojo: traen DOS columnas de estado. `Estado del Pedido` es el del
  // negocio; `Estado en Transportadora` es el del courier. Se guardan
  // las dos y el canónico sale del primero.
  mastershop: {
    'entregada':                ESTADOS.ENTREGADO,
    'entregado':                ESTADOS.ENTREGADO,
    'entregada digitalizada':   ESTADOS.ENTREGADO,
    'devuelta':                 ESTADOS.DEVOLUCION,
    'devolucion':               ESTADOS.DEVOLUCION,
    'en devolucion':            ESTADOS.DEVOLUCION,
    'cancelada':                ESTADOS.CANCELADO,
    'cancelado':                ESTADOS.CANCELADO,
    'novedad':                  ESTADOS.NOVEDAD,
    'con novedad':              ESTADOS.NOVEDAD,
    'pendiente':                ESTADOS.PENDIENTE,
    'por confirmar':            ESTADOS.PENDIENTE,
    'generada':                 ESTADOS.CONFIRMADO,
    'guia generada':            ESTADOS.CONFIRMADO,
    'en bodega':                ESTADOS.EN_BODEGA,
    'centro acopio':            ESTADOS.EN_BODEGA,
    'en terminal origen':       ESTADOS.EN_BODEGA,
    'en transito':              ESTADOS.EN_TRANSITO,
    'en transporte':            ESTADOS.EN_TRANSITO,
    'en reparto':               ESTADOS.EN_TRANSITO,
    'reparto':                  ESTADOS.EN_TRANSITO,
    'en distribucion':          ESTADOS.EN_TRANSITO,
    'en terminal destino':      ESTADOS.EN_TRANSITO,
    'intento de entrega':       ESTADOS.NOVEDAD,
    'en oficina':               ESTADOS.EN_OFICINA,
    'reclame en oficina':       ESTADOS.EN_OFICINA,
  },

  // Effi — VERIFICADO. Es un sistema de GUÍAS, no de pedidos.
  // Trae dos columnas: `Estado global guía inicial` (limpio, 7 valores) y
  // `Estado guía inicial` (sucio, con ciudad pegada). Se usa el global.
  // Los estados con ciudad ("DEVUELTA DESDE TUNJA", "ENTREGADA DIGITALIZADA
  // EN MEDELLIN") se resuelven por prefijo — la ciudad es dinámica y no se
  // puede enumerar.
  effi: {
    // Estado global — el bueno
    'entregada a destino':                ESTADOS.ENTREGADO,
    'generada':                           ESTADOS.CONFIRMADO,
    'devolucion a origen':                ESTADOS.DEVOLUCION,
    'en transito':                        ESTADOS.EN_TRANSITO,
    'en reparto':                         ESTADOS.EN_TRANSITO,
    'novedad':                            ESTADOS.NOVEDAD,
    'disponible para retiro en oficina':  ESTADOS.EN_OFICINA,
    // Estado de guía — el sucio, por si el global viene vacío
    'entregada':                          ESTADOS.ENTREGADO,
    'entregada digitalizada':             ESTADOS.ENTREGADO, // + " EN <CIUDAD>"
    'devuelta desde':                     ESTADOS.DEVOLUCION, // + " <CIUDAD>"
    'devolucion ratificada':              ESTADOS.DEVOLUCION,
    'generada effi':                      ESTADOS.CONFIRMADO,
    'admitida':                           ESTADOS.EN_BODEGA,
    'a recibir por':                      ESTADOS.EN_BODEGA, // + " <TRANSPORTADORA>"
    'cerrado por incidencia':             ESTADOS.NOVEDAD,   // + ", VER CAUSA"
    'en terminal destino':                ESTADOS.EN_TRANSITO,
  },

  // Shopify — POR VERIFICAR contra un export real
  shopify: {
    'fulfilled':      ESTADOS.ENTREGADO,
    'unfulfilled':    ESTADOS.PENDIENTE,
    'partial':        ESTADOS.EN_TRANSITO,
    'restocked':      ESTADOS.DEVOLUCION,
    'paid':           ESTADOS.CONFIRMADO,
    'pending':        ESTADOS.PENDIENTE,
    'refunded':       ESTADOS.DEVOLUCION,
    'voided':         ESTADOS.CANCELADO,
    'cancelled':      ESTADOS.CANCELADO,
  },
};

/** Traduce el estado de una plataforma al canónico de Nova. */
function estadoCanonico(fuente, texto) {
  if (!texto) return '';
  const mapa = MAPA_ESTADOS[fuente] || MAPA_ESTADOS[fuente === 'effi' ? 'mastershop' : ''] || {};
  const k = norm(texto);
  if (mapa[k]) return mapa[k];

  // Coincidencia por prefijo: los couriers agregan sufijos
  // ("ENTREGADA DIGITALIZADA", "PARA RETIRO EN AGENCIA SERVIENTREGA")
  const claves = Object.keys(mapa);
  for (let i = 0; i < claves.length; i++) {
    if (k.indexOf(claves[i]) === 0) return mapa[claves[i]];
  }
  return '__sin_mapear__:' + k; // visible, no silencioso
}

// ─── CATÁLOGO DE NOVEDADES ───────────────────────────────────
// Los motivos reales del Maestro de Ecuador, agrupados por causa.
// El agrupamiento es lo que hace útil la alarma de "patrón de novedades":
// 3 del mismo grupo en la semana es un problema de proceso.

const GRUPOS_NOVEDAD = {
  no_contacta: [
    'destinatario no contesta llamadas ni whatsapp',
    'no contesta cliente',
    'destinatario no se encuentra en lugar de entrega',
    'el cliente no estaba en su domicilio',
    'no hay quien reciba',
  ],
  rechaza: [
    'destinatario indica que ya no desea el producto',
    'destinatario indica que no ha comprado ningun producto',
    'titular se nego a recibir',
    'destinatario indica que no es el producto que solicito',
    'destinatario indica que ya recibio el producto',
  ],
  direccion: [
    'faltan datos en direccion',
    'destinatario solicita cambio de direccion',
    'fuera de cobertura',
    'no coincide la ciudad real destino con la ciudad ingresada',
  ],
  dinero: [
    'destinatario indica que no tiene el dinero para realizar el pago del recaudo',
  ],
  ausente: [
    'destinatario indica que se encuentra fuera de la ciudad',
    'destinatario agenda',
  ],
  logistica: [
    'en bodega', 'generada', 'devuelto de', 'devolucion de distribucion',
    'cerrado', 'no reclamo en oficina',
  ],
};

/**
 * Códigos numéricos de novedad de Effi — VERIFICADO.
 * El código es más estable que el texto: Effi puede reescribir la
 * descripción, el número no cambia. Cuando venga código, manda el código.
 */
const CODIGOS_NOVEDAD_EFFI = {
  '701': { grupo: 'no_contacta', txt: 'Se visita, no se logra entrega' },
  '81':  { grupo: 'no_contacta', txt: 'Coordinar la entrega' },
  '828': { grupo: 'dinero',      txt: 'No cancela el valor a recaudar (RCE)' },
  '801': { grupo: 'rechaza',     txt: 'Pedido cancelado' },
  '699': { grupo: 'direccion',   txt: 'Dirección incompleta' },
  '702': { grupo: 'direccion',   txt: 'No se localiza dirección del destinatario' },
  '703': { grupo: 'direccion',   txt: 'En dirección de entrega no conocen destinatario' },
  '706': { grupo: 'direccion',   txt: 'Destinatario solicita otra dirección' },
  '713': { grupo: 'direccion',   txt: 'Dirección destinatario no existe' },
  '728': { grupo: 'dinero',      txt: 'Solicita entrega en fecha posterior para pagar' },
  '44':  { grupo: 'rechaza',     txt: 'Destinatario se rehúsa a recibir' },
  '34':  { grupo: 'rechaza',     txt: 'Solicita inventario, unidades selladas' },
  '31':  { grupo: 'direccion',   txt: 'No conocen destinatario en dirección destino' },
};

/**
 * Agrupa un motivo de novedad en una de seis causas.
 * Si viene código de Effi, el código manda; si no, se busca por texto.
 */
function grupoNovedad(motivo, codigo) {
  if (codigo && CODIGOS_NOVEDAD_EFFI[String(codigo).trim()]) {
    return CODIGOS_NOVEDAD_EFFI[String(codigo).trim()].grupo;
  }
  if (!motivo) return '';
  const k = norm(motivo);
  const grupos = Object.keys(GRUPOS_NOVEDAD);
  for (let i = 0; i < grupos.length; i++) {
    const lista = GRUPOS_NOVEDAD[grupos[i]];
    for (let j = 0; j < lista.length; j++) {
      if (k.indexOf(lista[j]) !== -1) return grupos[i];
    }
  }
  return 'otro';
}

/**
 * Effi guarda el documento como "CC: 3223665889" — tipo y número pegados.
 * Devuelve { tipo, numero }.
 */
function partirDocumento(raw) {
  if (!raw) return { tipo: '', numero: '' };
  const s = String(raw).trim();
  const m = s.match(/^([A-Za-zÁÉÍÓÚÑ.]+)\s*:\s*(.+)$/);
  if (m) return { tipo: m[1].replace(/\./g, '').toUpperCase(), numero: m[2].replace(/\D/g, '') };
  return { tipo: '', numero: s.replace(/\D/g, '') };
}

// ─── TELÉFONO: PREFIJOS Y NORMALIZACIÓN ──────────────────────
/**
 * El problema real, medido en los datos:
 *   · Dropi Ecuador  → 984712695      (9 dígitos, sin prefijo)
 *   · Mastershop CO  → 3138880827     (10 dígitos, sin prefijo)
 *                    → +573202241205  (con prefijo, en el mismo archivo)
 *   · IRIS           → 3142935085 (10) · 34600019313 (11)
 *                    → 526391380301 (12) · 528461145190 (13)
 *
 * Sin normalizar, una llamada de IRIS nunca cruza con su pedido.
 * La regla: guardar SIEMPRE el número tal como llegó (`telefono`)
 * y además su forma canónica E.164 sin '+' (`telefono_norm`),
 * que es la única que se usa para cruzar.
 */

const PREFIJOS = {
  // Sudamérica
  AR: { cod: '54',  largo: 10, movil: /^9?[1-9]/ },  // Argentina
  BO: { cod: '591', largo: 8,  movil: /^[67]/ },     // Bolivia
  BR: { cod: '55',  largo: 11, movil: /^[1-9]/ },    // Brasil
  CL: { cod: '56',  largo: 9,  movil: /^9/ },        // Chile
  CO: { cod: '57',  largo: 10, movil: /^3/ },        // Colombia
  EC: { cod: '593', largo: 9,  movil: /^9/ },        // Ecuador
  GY: { cod: '592', largo: 7,  movil: /^6/ },        // Guyana
  PE: { cod: '51',  largo: 9,  movil: /^9/ },        // Perú
  PY: { cod: '595', largo: 9,  movil: /^9/ },        // Paraguay
  SR: { cod: '597', largo: 7,  movil: /^[78]/ },     // Surinam
  UY: { cod: '598', largo: 8,  movil: /^9/ },        // Uruguay
  VE: { cod: '58',  largo: 10, movil: /^4/ },        // Venezuela
  // Centroamérica y México
  BZ: { cod: '501', largo: 7,  movil: /^6/ },        // Belice
  CR: { cod: '506', largo: 8,  movil: /^[678]/ },    // Costa Rica
  GT: { cod: '502', largo: 8,  movil: /^[3-5]/ },    // Guatemala
  HN: { cod: '504', largo: 8,  movil: /^[389]/ },    // Honduras
  MX: { cod: '52',  largo: 10, movil: /^[1-9]/ },    // México
  NI: { cod: '505', largo: 8,  movil: /^[578]/ },    // Nicaragua
  PA: { cod: '507', largo: 8,  movil: /^6/ },        // Panamá
  SV: { cod: '503', largo: 8,  movil: /^[67]/ },     // El Salvador
  // Caribe
  CU: { cod: '53',  largo: 8,  movil: /^5/ },        // Cuba
  HT: { cod: '509', largo: 8,  movil: /^[34]/ },     // Haití
  // Europa / Norteamérica
  ES: { cod: '34',  largo: 9,  movil: /^[67]/ },     // España
  US: { cod: '1',   largo: 10, movil: /^[2-9]/ },    // EE.UU. y Canadá
};

/**
 * Países del plan NANP (+1). Comparten código con EE.UU., así que un
 * número dominicano y uno estadounidense son indistinguibles por el
 * prefijo: se separan por el código de área. Se listan para que un
 * número de RD no se marque como raro.
 */
const AREAS_NANP = {
  DO: ['809', '829', '849'],   // República Dominicana
  PR: ['787', '939'],          // Puerto Rico
  JM: ['876', '658'],          // Jamaica
  TT: ['868'],                 // Trinidad y Tobago
  BB: ['246'],                 // Barbados
  BS: ['242'],                 // Bahamas
};

// Códigos ordenados de más largo a más corto: '593' tiene que
// probarse antes que '59' o '5', si no se recorta mal.
const CODIGOS = Object.keys(PREFIJOS)
  .map(function (p) { return { pais: p, cod: PREFIJOS[p].cod, largo: PREFIJOS[p].largo }; })
  .sort(function (a, b) { return b.cod.length - a.cod.length; });

/**
 * Normaliza a E.164 sin '+'. `paisDefault` es el ISO de la tienda,
 * y se usa cuando el número llega sin prefijo.
 */
function telefonoNorm(raw, paisDefault) {
  if (!raw) return '';
  let s = String(raw).replace(/[^\d]/g, '');
  if (!s) return '';

  s = s.replace(/^0+/, '');       // ceros de marcación nacional al inicio
  s = s.replace(/^00/, '');       // prefijo internacional 00

  const def = PREFIJOS[paisDefault];

  /**
   * Cero de troncal DESPUÉS del código de país.
   * Shopify Ecuador exporta "+5930984635105": 593 + 0 + 984635105.
   * Sin quitar ese cero quedan 13 dígitos, no cuadra con ningún patrón
   * y el cruce con las llamadas de IRIS se rompe sin avisar.
   */
  function sinTroncal(num, cod, largo) {
    if (num.indexOf(cod) !== 0) return num;
    const resto = num.slice(cod.length);
    if (resto.length === largo + 1 && resto.charAt(0) === '0') {
      return cod + resto.slice(1);
    }
    return num;
  }

  if (def) s = sinTroncal(s, def.cod, def.largo);

  // Ya viene con el prefijo del país de la tienda y el largo cuadra
  if (def && s.indexOf(def.cod) === 0 && s.length === def.cod.length + def.largo) {
    return s;
  }
  // Viene sin prefijo, con el largo nacional exacto
  if (def && s.length === def.largo) {
    return def.cod + s;
  }
  // Trae el prefijo de otro país
  for (let i = 0; i < CODIGOS.length; i++) {
    const c = CODIGOS[i];
    const limpio = sinTroncal(s, c.cod, c.largo);
    if (limpio.indexOf(c.cod) === 0 && limpio.length === c.cod.length + c.largo) return limpio;
  }
  // No cuadra con ningún patrón: se devuelve tal cual, en dígitos.
  // Queda visible como número raro en vez de romper el cruce en silencio.
  return s;
}

/**
 * Extrae los números de contacto de un texto libre. Resuelve el caso
 * que pediste: el cliente escribe desde otro número, o pide que lo
 * contacten a uno distinto, y ese dato hoy se pierde en la nota.
 * Devuelve { principal, secundario }.
 */
function extraerTelefonos(campoTelefono, textoLibre, paisDefault) {
  const principal = telefonoNorm(campoTelefono, paisDefault);
  let secundario = '';

  if (textoLibre) {
    const encontrados = String(textoLibre).match(/(?:\+?\d[\d\s\-().]{6,})/g) || [];
    for (let i = 0; i < encontrados.length; i++) {
      const n = telefonoNorm(encontrados[i], paisDefault);
      if (n && n !== principal && n.length >= 8) { secundario = n; break; }
    }
  }
  return { principal: principal, secundario: secundario };
}


/* ═══════════════════════════════════════════════════════════════
   3 · MONEDAS Y CONVERSIÓN
   ═══════════════════════════════════════════════════════════════ */


// ─── CATÁLOGO ────────────────────────────────────────────────
// dec: decimales que usa la moneda en la práctica.
// CLP y PYG son 0: tratarlas con centavos infla los montos por 100.

const MONEDAS = {
  // Latinoamérica
  ARS: { nom: 'Peso argentino',        sim: '$',    dec: 2, pais: 'Argentina' },
  BOB: { nom: 'Boliviano',             sim: 'Bs',   dec: 2, pais: 'Bolivia' },
  BRL: { nom: 'Real brasileño',        sim: 'R$',   dec: 2, pais: 'Brasil' },
  CLP: { nom: 'Peso chileno',          sim: '$',    dec: 0, pais: 'Chile' },
  COP: { nom: 'Peso colombiano',       sim: '$',    dec: 0, pais: 'Colombia' },
  CRC: { nom: 'Colón costarricense',   sim: '₡',    dec: 2, pais: 'Costa Rica' },
  CUP: { nom: 'Peso cubano',           sim: '$',    dec: 2, pais: 'Cuba' },
  DOP: { nom: 'Peso dominicano',       sim: 'RD$',  dec: 2, pais: 'República Dominicana' },
  GTQ: { nom: 'Quetzal',               sim: 'Q',    dec: 2, pais: 'Guatemala' },
  HNL: { nom: 'Lempira',               sim: 'L',    dec: 2, pais: 'Honduras' },
  MXN: { nom: 'Peso mexicano',         sim: '$',    dec: 2, pais: 'México' },
  NIO: { nom: 'Córdoba',               sim: 'C$',   dec: 2, pais: 'Nicaragua' },
  PAB: { nom: 'Balboa',                sim: 'B/.',  dec: 2, pais: 'Panamá' },
  PEN: { nom: 'Sol',                   sim: 'S/',   dec: 2, pais: 'Perú' },
  PYG: { nom: 'Guaraní',               sim: '₲',    dec: 0, pais: 'Paraguay' },
  UYU: { nom: 'Peso uruguayo',         sim: '$U',   dec: 2, pais: 'Uruguay' },
  VES: { nom: 'Bolívar',               sim: 'Bs.',  dec: 2, pais: 'Venezuela' },
  // Caribe y Centroamérica
  BZD: { nom: 'Dólar beliceño',        sim: 'BZ$',  dec: 2, pais: 'Belice' },
  GYD: { nom: 'Dólar guyanés',         sim: 'G$',   dec: 2, pais: 'Guyana' },
  SRD: { nom: 'Dólar surinamés',       sim: '$',    dec: 2, pais: 'Surinam' },
  TTD: { nom: 'Dólar trinitense',      sim: 'TT$',  dec: 2, pais: 'Trinidad y Tobago' },
  JMD: { nom: 'Dólar jamaiquino',      sim: 'J$',   dec: 2, pais: 'Jamaica' },
  HTG: { nom: 'Gourde',                sim: 'G',    dec: 2, pais: 'Haití' },
  BSD: { nom: 'Dólar bahameño',        sim: '$',    dec: 2, pais: 'Bahamas' },
  BBD: { nom: 'Dólar barbadense',      sim: '$',    dec: 2, pais: 'Barbados' },
  AWG: { nom: 'Florín arubeño',        sim: 'ƒ',    dec: 2, pais: 'Aruba' },
  XCD: { nom: 'Dólar del Caribe Or.',  sim: '$',    dec: 2, pais: 'Caribe Oriental' },
  // Usadas en la región
  USD: { nom: 'Dólar estadounidense',  sim: '$',    dec: 2, pais: 'Ecuador · Panamá · El Salvador' },
  // Resto del mundo
  EUR: { nom: 'Euro',                  sim: '€',    dec: 2, pais: 'Zona euro' },
  GBP: { nom: 'Libra esterlina',       sim: '£',    dec: 2, pais: 'Reino Unido' },
  CAD: { nom: 'Dólar canadiense',      sim: 'C$',   dec: 2, pais: 'Canadá' },
  CHF: { nom: 'Franco suizo',          sim: 'CHF',  dec: 2, pais: 'Suiza' },
  JPY: { nom: 'Yen',                   sim: '¥',    dec: 0, pais: 'Japón' },
  CNY: { nom: 'Yuan',                  sim: '¥',    dec: 2, pais: 'China' },
  AUD: { nom: 'Dólar australiano',     sim: 'A$',   dec: 2, pais: 'Australia' },
};

/**
 * Devuelve la ficha de una moneda. Una moneda desconocida no rompe nada:
 * cae en 2 decimales y usa su propio código como símbolo.
 */
function moneda(cod) {
  const c = String(cod || '').trim().toUpperCase();
  return MONEDAS[c] || { nom: c || 'desconocida', sim: c, dec: 2, pais: '', desconocida: true };
}

/**
 * Formatea un monto en su moneda, con convención latinoamericana:
 * punto para miles, coma para decimales. Ej: Q 1.250,50
 *
 * El separador de miles se aplica SOLO a la parte entera; aplicarlo a
 * la cadena completa produce "1.250.50", que ya no es un número.
 */
function fmtMoneda(monto, cod) {
  const m = moneda(cod);
  const n = aNumero(monto);
  if (n === '') return '';

  const signo = n < 0 ? '-' : '';
  const partes = Math.abs(n).toFixed(m.dec).split('.');
  const entera = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const dec = partes.length > 1 ? ',' + partes[1] : '';

  return signo + m.sim + ' ' + entera + dec;
}

// ─── TASAS ───────────────────────────────────────────────────

/**
 * Busca la tasa de un día concreto en la hoja `Tasas`.
 *
 * Si no hay tasa exacta para esa fecha, usa la más reciente ANTERIOR
 * (nunca una posterior: sería usar información del futuro). Si tampoco
 * hay ninguna anterior, devuelve null.
 *
 * @return {Object|null} { tasa, fecha_usada, exacta }
 */
function buscarTasa(ss, fecha, origen, destino) {
  origen = String(origen || '').toUpperCase();
  destino = String(destino || '').toUpperCase();
  if (!origen || !destino) return null;
  if (origen === destino) return { tasa: 1, fecha_usada: fecha, exacta: true };

  const sh = ss.getSheetByName('Tasas');
  if (!sh || sh.getLastRow() < 2) return null;

  const filas = sh.getDataRange().getValues().slice(1);
  let mejor = null;

  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];
    const fFecha = aISO(f[0], 'UTC');
    const fOri = String(f[1] || '').toUpperCase();
    const fDes = String(f[2] || '').toUpperCase();
    const fTasa = aNumero(f[3]);
    if (!fFecha || fTasa === '' || fTasa <= 0) continue;

    let tasa = null;
    if (fOri === origen && fDes === destino) tasa = fTasa;
    else if (fOri === destino && fDes === origen) tasa = 1 / fTasa; // el par inverso sirve igual
    if (tasa === null) continue;

    if (fFecha === fecha) return { tasa: tasa, fecha_usada: fFecha, exacta: true };
    if (fFecha < fecha && (!mejor || fFecha > mejor.fecha_usada)) {
      mejor = { tasa: tasa, fecha_usada: fFecha, exacta: false };
    }
  }
  return mejor;
}

/**
 * Convierte un monto usando la tasa del día de la transacción.
 *
 * NO inventa un número cuando falta la tasa: devuelve valor null y un
 * motivo. Quien llama decide qué hacer, pero nunca recibe una cifra
 * inventada que se vea como buena.
 *
 * @return {Object} { valor, tasa, fecha_usada, exacta, motivo }
 */
function convertir(ss, monto, fecha, origen, destino) {
  const n = aNumero(monto);
  if (n === '') return { valor: null, motivo: 'monto no numérico' };

  const o = String(origen || '').toUpperCase();
  const d = String(destino || '').toUpperCase();
  if (!o || !d) return { valor: null, motivo: 'falta moneda origen o destino' };
  if (o === d) return { valor: n, tasa: 1, fecha_usada: fecha, exacta: true };

  const t = buscarTasa(ss, fecha, o, d);
  if (!t) {
    return {
      valor: null,
      motivo: 'sin tasa ' + o + '->' + d + ' para ' + fecha +
              ' ni ninguna anterior. Agrega la fila en la hoja Tasas.',
    };
  }

  const dec = moneda(d).dec;
  return {
    valor: Number((n * t.tasa).toFixed(dec)),
    tasa: t.tasa,
    fecha_usada: t.fecha_usada,
    exacta: t.exacta,
    motivo: t.exacta ? '' : 'tasa del ' + t.fecha_usada + ', no había del ' + fecha,
  };
}

/**
 * Revisa qué tasas hacen falta para poder convertir todo lo que ya está
 * cargado. Córrela antes de confiar en cualquier número de dinero.
 */
function tasasFaltantes(cliente, monedaDestino) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const destino = String(monedaDestino || 'COP').toUpperCase();
  const faltan = {};

  ['Pedidos', 'Pauta'].forEach(function (tab) {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    const datos = sh.getDataRange().getValues();
    const enc = datos[0].map(norm);
    const cFecha = enc.indexOf('fecha');
    const cMon = enc.indexOf('moneda_gasto') !== -1
      ? enc.indexOf('moneda_gasto') : enc.indexOf('moneda');
    if (cFecha === -1 || cMon === -1) return;

    datos.slice(1).forEach(function (f) {
      const fecha = aISO(f[cFecha], 'UTC');
      const mon = String(f[cMon] || '').toUpperCase();
      if (!fecha || !mon || mon === destino) return;
      if (!buscarTasa(ss, fecha, mon, destino)) {
        faltan[mon + ' -> ' + destino + '  desde ' + fecha] = true;
      }
    });
  });

  const lista = Object.keys(faltan).sort();
  const msg = lista.length
    ? 'FALTAN TASAS (' + lista.length + '):\n  ' + lista.join('\n  ') +
      '\n\nAgrégalas en la hoja Tasas: fecha · moneda_origen · moneda_destino · tasa'
    : 'No falta ninguna tasa para convertir a ' + destino + '. ✓';
  Logger.log(msg);
  return msg;
}


/* ═══════════════════════════════════════════════════════════════
   4 · IMPORTADORES
   ═══════════════════════════════════════════════════════════════ */


// ─── CATÁLOGO DE FUENTES ─────────────────────────────────────
// `alias` mapea: columna_normalizada -> [posibles nombres en el export]
// El importador toma el primero que encuentre, sin distinguir may/min ni tildes.

const FUENTES = {

  // ── PEDIDOS / FULFILLMENT ──────────────────────────────────

  // VERIFICADO contra el Maestro de Nutrea EC (Dropi-Pedidos-NutreaShop.xlsx).
  // Ese archivo ya es la salida normalizada de tu motor anterior, así que
  // estos alias cubren tanto el export crudo de Dropi como tu Maestro.
  dropi: {
    tipo: 'pedidos',
    verificado: true,
    alias: {
      id_externo:       ['id', 'no. orden', 'numero orden', 'orden'],
      fecha:            ['fecha', 'fecha creacion', 'created at'],
      cliente:          ['nombre cliente', 'cliente', 'nombre'],
      telefono:         ['telefono', 'celular', 'phone'],
      guia:             ['numero guia', 'no guia', 'guia', 'tracking'],
      estado:           ['estatus', 'estado', 'status'],
      ciudad:           ['ciudad destino', 'ciudad', 'city'],
      transportadora:   ['transportadora', 'courier', 'operador'],
      valor:            ['total de la orden', 'valor', 'total', 'monto'],
      ganancia:         ['ganancia'],
      producto:         ['producto', 'productos', 'nombre producto'],
      cantidad:         ['cantidad', 'qty', 'unidades'],
      motivo_novedad:   ['novedad'],
      solucion:         ['solucion'],
      ultimo_movimiento:['ultimo movimiento'],
      fecha_ingreso:    ['fecha ingreso'],
      actualizado:      ['ultima actualizacion'],
      direccion:        ['direccion', 'address'],
      sku:              ['sku', 'codigo'],
      costo_envio:      ['flete', 'costo envio', 'envio'],
    },
  },

  // VERIFICADO contra reporte-historial-de-pedidos-*.xlsx
  mastershop: {
    tipo: 'pedidos',
    verificado: true,
    alias: {
      id_externo:       ['id pedido'],
      guia:             ['n° de guia', 'no de guia', 'n de guia'],
      fecha:            ['fecha del pedido'],
      fecha_entrega:    ['fecha de entrega'],
      cliente:          ['cliente'],
      cedula:           ['cedula'],
      correo:           ['correo electronico'],
      telefono:         ['telefono'],
      direccion:        ['direccion'],
      ciudad:           ['ciudad'],
      departamento:     ['departamento'],
      transportadora:   ['transportadora'],
      bodega:           ['nombre de la bodega'],
      costo_envio:      ['valor del envio'],
      tiene_novedad:    ['presento novedad'],
      estado:           ['estado del pedido'],
      estado_transportadora: ['estado en transportadora'],
      ultimo_movimiento:['ultima actualizacion'],
      metodo_pago:      ['metodo de pago'],
      producto:         ['producto'],
      valor:            ['valor del pedido'],
      cantidad:         ['cantidad'],
      costo_producto:   ['costo de productos'],
      nota:             ['notas'],
      razon_cancelacion:['razon de cancelacion'],
    },
  },

  // Effi — VERIFICADO. NO hereda de Mastershop: es un sistema de GUÍAS,
  // con vocabulario de transporte (Remitente/Destinatario) en vez de
  // pedidos (Cliente). Emite DOS reportes separados que hay que cruzar
  // por número de guía.
  effi_guias: {
    tipo: 'pedidos',
    verificado: true,
    tab: '_Import_Effi_Guias',
    alias: {
      prefijo_guia:     ['prefijo id guia'],
      guia:             ['guia transportadora'],
      guia_devolucion:  ['guia devolucion transportadora'],
      fecha:            ['fecha de envio'],
      fecha_promesa:    ['fecha de entrega esperada'],
      fecha_entrega:    ['fecha de estado final'],
      cliente:          ['destinatario'],
      documento:        ['id. destinatario'],   // viene como "CC: 3223665889"
      pais:             ['pais destinatario'],
      departamento:     ['departamento destinatario'],
      ciudad:           ['ciudad destinatario'],
      direccion:        ['direccion destinatario'],
      telefono:         ['telefonos destinatario'],
      transportadora:   ['nombre transportadora efficommerce'],
      transportadora_det:['transportadora'],
      costo_envio:      ['precio flete total a cliente'],
      flete_base:       ['precio flete a cliente'],
      seguro:           ['precio manejo (seguro) a cliente'],
      descuento:        ['% descuento'],
      valor:            ['valor recaudo', 'valor declarado'],
      valor_declarado:  ['valor declarado'],
      metodo_pago:      ['forma de pago'],
      producto:         ['contenido'],
      cantidad:         ['cantidad de paquetes'],
      peso:             ['peso (kg)'],
      // El global es el limpio (7 valores); el otro trae la ciudad pegada
      estado:           ['estado global guia inicial'],
      estado_transportadora: ['estado guia inicial'],
      estado_devolucion:['estado global guia devolucion'],
      nota:             ['nota guia'],
      observacion:      ['observacion'],
      sucursal:         ['sucursal'],
      centro_costos:    ['centro de costos'],
      creado_en:        ['fecha de creacion'],
      creado_por:       ['responsable de creacion'],
    },
  },

  effi_novedades: {
    tipo: 'novedades',
    verificado: true,
    tab: '_Import_Effi_Novedades',
    alias: {
      id_externo:       ['id novedad'],
      fecha:            ['fecha novedad'],
      guia:             ['guia novedad'],
      codigo:           ['cod. novedad'],       // 701, 828... más estable que el texto
      motivo:           ['novedad'],
      aclaracion:       ['aclaracion'],
      solucion:         ['historico de soluciones'],
      foto:             ['url foto'],
      creada_en:        ['fecha creacion effi'],
      creada_por:       ['responsable creacion'],
      cliente:          ['destinatario'],
      telefono:         ['telefono destinatario'],
      ciudad:           ['ciudad destinatario'],
      departamento:     ['departamento destinatario'],
      direccion:        ['direccion destinatario'],
      transportadora:   ['transportadora efficommerce'],
      valor:            ['valor recaudo', 'valor declarado'],
      costo_envio:      ['valor flete'],
      producto:         ['contenido'],
      estado:           ['estado'],
    },
  },

  // IRIS NO es una plataforma de pedidos: es la central telefónica.
  // VERIFICADO contra IRIS (1).csv — 2.782 llamadas salientes.
  // Se usa en TODAS las tiendas, no solo en una.
  //
  // Ojo: el export de IRIS no trae columna de tienda. La llamada se
  // asigna cruzando `telefono_norm` contra Pedidos: la tienda sale del
  // pedido que coincide. Una llamada sin pedido que la reciba queda con
  // tienda vacía y `pedido_id` vacío — visible, no descartada. Eso es
  // justamente lo que hay que revisar: son llamadas a números que no
  // están en ningún pedido.
  iris: {
    tipo: 'llamadas',
    verificado: true,
    multi_tienda: true,      // se reparte por cruce, no por configuración
    alias: {
      id_externo:     ['id'],
      uid:            ['uniqueid'],
      fecha_hora:     ['fecha hora'],
      seg_conversado: ['tiempo conversado'],
      seg_espera:     ['tiempo espera'],
      seg_total:      ['tiempo total'],
      sentido:        ['sentido'],
      estado:         ['estado'],
      extension:      ['extension'],
      telefono:       ['numero'],
      agente:         ['usuario'],
      campana:        ['campana'],
      grabacion:      ['grabacion'],
      etiqueta:       ['etiqueta'],
      observacion:    ['observacion'],
    },
  },

  // Shopify — VERIFICADO (orders_export_1.csv, 729 pedidos, 79 columnas).
  //
  // SECUNDARIA, no principal. La verdad operativa vive en Dropi/Effi/Mastershop:
  // Shopify solo sabe si despachó o no (`fulfilled`/`unfulfilled`), y su
  // `Financial Status` es "pending" en las 729 filas porque todo es contraentrega.
  // Contra los 17 estados de Dropi, no sirve para operar.
  //
  // Sirve para tres cosas que ninguna otra fuente tiene:
  //   · el correo del cliente
  //   · descuentos, impuestos y código de cupón
  //   · la conciliación de fuga: pedidos que entraron a la tienda y nunca
  //     llegaron a tener guía. Sin cruzar Shopify contra la plataforma de
  //     fulfillment, esos pedidos no aparecen en ningún lado.
  shopify: {
    tipo: 'pedidos_secundario',
    verificado: true,
    moneda_default: 'USD',
    alias: {
      id_externo:      ['name'],              // "#1729"
      id_interno:      ['id'],
      correo:          ['email'],
      fecha:           ['created at'],
      fecha_pago:      ['paid at'],
      fecha_despacho:  ['fulfilled at'],
      fecha_cancelado: ['cancelled at'],
      estado:          ['fulfillment status'],
      estado_pago:     ['financial status'],
      moneda_gasto:    ['currency'],
      subtotal:        ['subtotal'],
      costo_envio:     ['shipping'],
      impuestos:       ['taxes'],
      valor:           ['total'],
      cupon:           ['discount code'],
      descuento:       ['discount amount'],
      metodo_envio:    ['shipping method'],
      metodo_pago:     ['payment method'],
      cantidad:        ['lineitem quantity'],
      producto:        ['lineitem name'],
      precio_unitario: ['lineitem price'],
      sku:             ['lineitem sku'],
      cliente:         ['shipping name', 'billing name'],
      telefono:        ['shipping phone', 'billing phone'],
      direccion:       ['shipping address1', 'billing address1'],
      direccion_2:     ['shipping address2'],
      ciudad:          ['shipping city', 'billing city'],
      departamento:    ['shipping province', 'billing province'],
      pais:            ['shipping country', 'billing country'],
      nota:            ['notes'],
      etiquetas:       ['tags'],
      riesgo:          ['risk level'],
      origen:          ['source'],
    },
  },

  // ── PAUTA ──────────────────────────────────────────────────

  // Meta CAMPAÑAS — VERIFICADO. Es reporte a nivel campaña, no de conjunto.
  // Trae CPA y ROAS ya calculados: mejor usarlos que recalcularlos.
  // Ojo: dice "coste", no "costo".
  meta: {
    tipo: 'pauta',
    verificado: true,
    tab: '_Import_Meta',
    moneda_default: 'COP',
    alias: {
      fecha:        ['inicio del informe', 'dia', 'fecha', 'date'],
      fecha_fin:    ['fin del informe'],
      campana:      ['nombre de la campana'],
      entrega:      ['entrega de la campana'],
      conjunto:     ['nombre del conjunto de anuncios', 'conjunto'],
      presupuesto:  ['presupuesto del conjunto de anuncios'],
      gasto:        ['importe gastado (cop)', 'importe gastado', 'amount spent'],
      cpm:          ['cpm (coste por 1000 impresiones) (cop)', 'cpm'],
      resultados:   ['resultados'],
      compras:      ['compras'],
      cpa:          ['coste por compra (cop)'],
      roas:         ['roas (retorno del gasto publicitario) de compras'],
      aov:          ['aov'],
      valor_conv:   ['valor de conversion de compras'],
      clics:        ['clics en el enlace', 'clics', 'clicks'],
      ctr:          ['ctr (tasa de clics en el enlace)'],
      cpc:          ['cpc (coste por clic en el enlace) (cop)'],
      impresiones:  ['impresiones', 'impressions'],
      alcance:      ['alcance'],
      frecuencia:   ['frecuencia'],
      visitas_lp:   ['visitas a la pagina de destino'],
    },
  },

  // Meta FACTURACIÓN — VERIFICADO. Formato distinto al de campañas:
  //   · 10 líneas de metainformación ANTES del encabezado real
  //   · montos con espacio de miles ("195 866")
  //   · una fila de total al final que hay que descartar
  //   · fechas D/M/AAAA (las de campañas vienen en ISO)
  // Son los cargos a la tarjeta, no el gasto por campaña. Sirve para
  // cuadrar que lo facturado coincida con lo reportado.
  meta_facturacion: {
    tipo: 'facturacion',
    verificado: true,
    tab: '_Import_Meta_Facturacion',
    moneda_default: 'COP',
    encabezado_tras: 'fecha,identificador de la transaccion', // ancla del header
    descartar_filas: ['importe total facturado'],
    alias: {
      fecha:        ['fecha'],
      id_externo:   ['identificador de la transaccion'],
      gasto:        ['importe'],
      moneda_gasto: ['divisa'],
    },
  },

  // TikTok — SIN VERIFICAR. No hay export de muestra todavía.
  //
  // Los alias de abajo son candidatos en los dos idiomas en que TikTok Ads
  // Manager exporta, NO nombres confirmados. No dependas de ellos.
  //
  // El camino real es el auto-mapeo: pega el export en _Import_Tiktok y corre
  //   proponerMapeo('tiktok', 'gt')
  // Detecta las columnas por la FORMA de los datos, no por el nombre, escribe
  // sus propuestas en la hoja `Mapeos`, y lo que quede ahí manda sobre esto.
  // Así TikTok entra sin que nadie tenga que ver el archivo primero.
  tiktok: {
    tipo: 'pauta',
    verificado: false,
    automapeo: true, // usa el detector si `Mapeos` no tiene nada
    moneda_default: 'USD',
    alias: {
      fecha:        ['date', 'fecha', 'dia', 'time', 'día'],
      campana:      ['campaign name', 'nombre de la campaña', 'nombre de campaña', 'campaña'],
      conjunto:     ['ad group name', 'nombre del grupo de anuncios', 'grupo de anuncios'],
      anuncio:      ['ad name', 'nombre del anuncio'],
      gasto:        ['cost', 'spend', 'total cost', 'gasto', 'costo', 'costo total'],
      impresiones:  ['impressions', 'impresiones'],
      clics:        ['clicks', 'clics', 'clicks (destination)'],
      resultados:   ['conversions', 'conversiones', 'results', 'resultados'],
      cpm:          ['cpm', 'costo por mil'],
      cpc:          ['cpc', 'costo por clic'],
      cpa:          ['cost per conversion', 'costo por conversión', 'cpa'],
      ctr:          ['ctr', 'tasa de clics'],
      moneda_gasto: ['currency', 'moneda', 'divisa'],
    },
  },
};

// ─── NORMALIZACIÓN DE TEXTO ──────────────────────────────────

/** Quita tildes, baja a minúscula, colapsa espacios. */
function norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().trim().replace(/\s+/g, ' ');
}

/** Monto a número puro: sin $, sin puntos de miles, sin espacios. */
function aNumero(v) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return '';
  let s = String(v).replace(/[^\d,.\-]/g, '').trim();
  if (!s) return '';
  const ultimaComa = s.lastIndexOf(',');
  const ultimoPunto = s.lastIndexOf('.');
  // El separador decimal es el que aparece de último
  if (ultimaComa > ultimoPunto) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    s = s.replace(/,/g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? '' : n;
}

/** Fecha a ISO AAAA-MM-DD. Causa número uno de errores al leer hojas. */
function aISO(v, zonaHoraria) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, zonaHoraria || 'UTC', 'yyyy-MM-dd');
  }
  const s = String(v).trim();

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);           // ya ISO
  if (m) return m[1] + '-' + m[2] + '-' + m[3];

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); // DD/MM/AAAA
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, zonaHoraria || 'UTC', 'yyyy-MM-dd');
  }
  return '';
}

// ─── MOTOR DE IMPORTACIÓN ────────────────────────────────────

/**
 * Lee una pestaña _Import_* y devuelve filas normalizadas.
 * No escribe nada: separar la lectura de la escritura es lo que
 * permite probar el mapeo sin ensuciar las hojas de producción.
 */
function leerCrudo(ss, fuenteId, tienda) {
  const cfg = FUENTES[fuenteId];
  if (!cfg) throw new Error('Fuente desconocida: ' + fuenteId);

  const tab = cfg.tab ||
    ('_Import_' + fuenteId.charAt(0).toUpperCase() + fuenteId.slice(1));
  const sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('Falta la pestaña ' + tab);

  const datos = sh.getDataRange().getValues();
  if (datos.length < 2) return { filas: [], sinMapear: [], tab: tab };

  // ── Encontrar la fila de encabezados ──
  // El informe de facturación de Meta trae 10 líneas de metainformación
  // antes del header real, y varias tienen 3+ celdas llenas. Por eso, cuando
  // la fuente declara un ancla, se busca esa; el conteo de celdas no sirve.
  let filaEnc = 0;
  if (cfg.encabezado_tras) {
    const ancla = norm(cfg.encabezado_tras);
    for (let i = 0; i < Math.min(datos.length, 40); i++) {
      if (datos[i].map(norm).join(',').indexOf(ancla) === 0) { filaEnc = i; break; }
    }
  } else {
    for (let i = 0; i < Math.min(datos.length, 10); i++) {
      const llenas = datos[i].filter(function (c) { return String(c).trim() !== ''; });
      if (llenas.length >= 3) { filaEnc = i; break; }
    }
  }

  const enc = datos[filaEnc].map(norm);
  let cuerpo = datos.slice(filaEnc + 1);

  // Filas de total / subtotal que no son hechos y romperían las sumas
  if (cfg.descartar_filas && cfg.descartar_filas.length) {
    const patrones = cfg.descartar_filas.map(norm);
    cuerpo = cuerpo.filter(function (r) {
      const linea = r.map(norm).join(' ');
      return !patrones.some(function (p) { return linea.indexOf(p) !== -1; });
    });
  }

  // ── Resolver el mapeo de columnas ──
  // Prioridad: lo que confirmaste en la hoja `Mapeos` gana sobre los alias
  // del código. Así una corrección tuya sobrevive a cualquier cambio del .gs
  const alias = Object.assign({}, cfg.alias || {}, aliasDesdeMapeos(ss, fuenteId) || {});

  const idx = {};
  Object.keys(alias).forEach(function (campo) {
    const opciones = alias[campo].map(norm);
    for (let i = 0; i < enc.length; i++) {
      if (opciones.indexOf(enc[i]) !== -1) { idx[campo] = i; return; }
    }
  });

  const sinMapear = Object.keys(alias).filter(function (c) {
    return idx[c] === undefined;
  });

  const tz = zonaHorariaDe(ss, tienda);
  const filas = cuerpo
    .filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); })
    .map(function (r) {
      const o = { fuente: fuenteId, tienda: tienda };
      Object.keys(idx).forEach(function (campo) {
        let v = r[idx[campo]];
        if (campo === 'fecha') v = aISO(v, tz);
        else if (['valor','costo_envio','costo_producto','cantidad','gasto',
                  'impresiones','clics','resultados','cpm'].indexOf(campo) !== -1) {
          v = aNumero(v);
        } else v = String(v == null ? '' : v).trim();
        o[campo] = v;
      });
      return o;
    });

  return { filas: filas, sinMapear: sinMapear, tab: tab, tipo: cfg.tipo };
}

/** Zona horaria de la tienda, para que las alarmas 1/2/4 evalúen en hora local. */
function zonaHorariaDe(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return 'UTC';
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cTz = enc.indexOf('zona_horaria');
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][cId]).trim() === tienda) return filas[i][cTz] || 'UTC';
  }
  return 'UTC';
}

/**
 * Diagnóstico. Córrelo con un export pegado en la pestaña y te dice
 * exactamente qué columnas no encontró, para completar el mapeo.
 */
function diagnosticar(fuenteId, tienda, cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const r = leerCrudo(ss, fuenteId, tienda || 'gt');
  const msg = [
    'Fuente: ' + fuenteId + '  (' + r.tipo + ')',
    'Pestaña: ' + r.tab,
    'Filas leídas: ' + r.filas.length,
    'Columnas SIN mapear: ' + (r.sinMapear.length ? r.sinMapear.join(', ') : 'ninguna ✓'),
    '',
    'Primera fila normalizada:',
    JSON.stringify(r.filas[0] || {}, null, 2),
  ].join('\n');
  Logger.log(msg);
  return msg;
}


/* ═══════════════════════════════════════════════════════════════
   5 · AUTO-MAPEO
   ═══════════════════════════════════════════════════════════════ */


// ─── BANCO DE SINÓNIMOS ──────────────────────────────────────
// Tokens en español e inglés. No hace falta el nombre exacto:
// basta con que el encabezado CONTENGA alguno.

const SINONIMOS = {
  fecha:       ['fecha','dia','date','day','inicio','start','reporting','periodo','time'],
  campana:     ['campana','campaign','campaña'],
  conjunto:    ['conjunto','grupo','adgroup','ad group','adset','ad set','grupo de anuncios'],
  anuncio:     ['anuncio','ad name','creative','creativo'],
  gasto:       ['gasto','gastado','costo','coste','cost','spend','importe','amount','inversion'],
  impresiones: ['impresiones','impressions','impr','views','visualizaciones'],
  clics:       ['clics','clicks','clic','click','pulsaciones'],
  resultados:  ['resultados','results','conversiones','conversions','acciones'],
  compras:     ['compras','purchases','pedidos','orders','ventas'],
  cpm:         ['cpm','coste por 1000','costo por 1000','cost per 1000','cost per mille'],
  cpc:         ['cpc','coste por clic','costo por clic','cost per click'],
  cpa:         ['cpa','coste por compra','costo por compra','cost per conversion',
                'cost per purchase','coste por resultado','costo por resultado'],
  ctr:         ['ctr','tasa de clics','click through','click-through'],
  roas:        ['roas','retorno','return on ad'],
  alcance:     ['alcance','reach'],
  frecuencia:  ['frecuencia','frequency'],
  moneda_gasto:['moneda','divisa','currency'],
  entrega:     ['entrega','delivery','estado','status','state'],
  presupuesto: ['presupuesto','budget'],
  valor_conv:  ['valor de conversion','conversion value','revenue','ingresos'],
};

// Forma esperada de cada campo. Es lo que desempata cuando el nombre no ayuda.
const FORMAS = {
  fecha:       'fecha',
  campana:     'texto_repetido',
  conjunto:    'texto_repetido',
  anuncio:     'texto_repetido',
  gasto:       'numero',
  impresiones: 'entero_grande',
  clics:       'entero',
  resultados:  'entero',
  compras:     'entero',
  cpm:         'numero',
  cpc:         'numero',
  cpa:         'numero',
  ctr:         'porcentaje',
  roas:        'numero',
  alcance:     'entero_grande',
  frecuencia:  'decimal_bajo',
  moneda_gasto:'texto_constante',
  entrega:     'texto_repetido',
  presupuesto: 'numero',
  valor_conv:  'numero',
};

// ─── PERFIL DE UNA COLUMNA ───────────────────────────────────

/** Mide qué forma tiene una columna, a partir de hasta 60 valores. */
function perfilar(valores) {
  const v = valores
    .filter(function (x) { return x !== null && x !== '' && x !== undefined; })
    .slice(0, 60);
  if (!v.length) return { vacia: true };

  let fechas = 0, numeros = 0, enteros = 0, sumaAbs = 0, maxAbs = 0, entre0y1 = 0;
  const unicos = {};

  v.forEach(function (x) {
    unicos[String(x)] = 1;
    if (x instanceof Date || aISO(x, 'UTC')) fechas++;
    const n = aNumero(x);
    if (n !== '' && !isNaN(n)) {
      numeros++;
      if (Math.abs(n) === Math.round(Math.abs(n))) enteros++;
      const a = Math.abs(n);
      sumaAbs += a;
      if (a > maxAbs) maxAbs = a;
      if (a > 0 && a <= 1) entre0y1++;
    }
  });

  const n = v.length;
  return {
    vacia: false,
    n: n,
    pFecha:   fechas / n,
    pNumero:  numeros / n,
    pEntero:  numeros ? enteros / numeros : 0,
    pEntre01: numeros ? entre0y1 / numeros : 0,
    prom:     numeros ? sumaAbs / numeros : 0,
    max:      maxAbs,
    cardinalidad: Object.keys(unicos).length / n,
  };
}

/** Qué tanto encaja una columna con la forma esperada. 0 a 50. */
function puntajeForma(forma, p) {
  if (p.vacia) return 0;
  switch (forma) {
    case 'fecha':
      return p.pFecha > 0.8 ? 50 : (p.pFecha > 0.5 ? 25 : 0);
    case 'numero':
      return p.pNumero > 0.85 && p.pFecha < 0.5 ? 30 : 0;
    case 'entero':
      return p.pNumero > 0.85 && p.pEntero > 0.9 ? 35 : 0;
    case 'entero_grande':
      return p.pNumero > 0.85 && p.pEntero > 0.9 && p.prom > 100 ? 40 : 0;
    case 'porcentaje':
      return p.pNumero > 0.85 && p.pEntre01 > 0.7 ? 40 : 0;
    case 'decimal_bajo':
      return p.pNumero > 0.85 && p.prom > 0 && p.prom < 20 ? 30 : 0;
    case 'texto_repetido':
      return p.pNumero < 0.3 && p.cardinalidad < 0.9 ? 30 : 0;
    case 'texto_constante':
      return p.cardinalidad < 0.1 ? 40 : 0;
    default:
      return 0;
  }
}

/** Qué tanto se parece el encabezado al campo. 0 a 100. */
function puntajeNombre(campo, encabezado) {
  const h = norm(encabezado);
  if (!h) return 0;
  const syns = SINONIMOS[campo] || [];
  let mejor = 0;
  for (let i = 0; i < syns.length; i++) {
    const s = norm(syns[i]);
    if (h === s) { mejor = Math.max(mejor, 100); continue; }
    if (h.indexOf(s) !== -1) {
      // Un token corto dentro de un encabezado largo vale menos:
      // "cpc" dentro de "cpc (coste por clic)" es fuerte;
      // "cost" dentro de "cost per 1000 impressions" es débil.
      const prop = s.length / h.length;
      mejor = Math.max(mejor, 55 + Math.round(prop * 35));
    }
  }
  return mejor;
}

// ─── PROPUESTA DE MAPEO ──────────────────────────────────────

/**
 * Analiza una pestaña cruda y propone qué columna corresponde a cada campo.
 * No escribe nada en las entidades: solo propone.
 */
function analizarFuente(ss, fuenteId, campos) {
  const cfg = FUENTES[fuenteId] || {};
  const tab = cfg.tab ||
    ('_Import_' + fuenteId.charAt(0).toUpperCase() + fuenteId.slice(1));
  const sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('Falta la pestaña ' + tab);

  const datos = sh.getDataRange().getValues();
  if (datos.length < 2) throw new Error(tab + ' está vacía. Pega el export primero.');

  // Encabezado: la primera fila con 3+ celdas de texto no numérico
  let filaEnc = 0;
  for (let i = 0; i < Math.min(datos.length, 15); i++) {
    const textos = datos[i].filter(function (c) {
      return String(c).trim() !== '' && aNumero(c) === '';
    });
    if (textos.length >= 3) { filaEnc = i; break; }
  }

  const enc = datos[filaEnc];
  const cuerpo = datos.slice(filaEnc + 1);
  const perfiles = enc.map(function (_, j) {
    return perfilar(cuerpo.map(function (r) { return r[j]; }));
  });

  const objetivo = campos || Object.keys(SINONIMOS);
  const propuestas = [];
  const usadas = {};

  // Se resuelven primero los campos con más señal, para que no se
  // "roben" columnas los que son ambiguos.
  const ranking = [];
  objetivo.forEach(function (campo) {
    enc.forEach(function (h, j) {
      if (!String(h).trim()) return;
      const pn = puntajeNombre(campo, h);
      const pf = puntajeForma(FORMAS[campo] || 'texto', perfiles[j]);
      if (pn === 0 && pf === 0) return;
      ranking.push({ campo: campo, col: j, encabezado: String(h), punt: pn + pf, pn: pn, pf: pf });
    });
  });
  ranking.sort(function (a, b) { return b.punt - a.punt; });

  const resueltos = {};
  ranking.forEach(function (r) {
    if (resueltos[r.campo] || usadas[r.col]) return;
    // Sin señal de nombre, la forma sola no alcanza para decidir
    if (r.pn === 0 && r.pf < 40) return;
    resueltos[r.campo] = true;
    usadas[r.col] = true;
    propuestas.push({
      campo: r.campo,
      columna: r.encabezado,
      indice: r.col,
      puntaje: r.punt,
      confianza: r.punt >= 110 ? 'alta' : (r.punt >= 70 ? 'media' : 'baja'),
    });
  });

  // Regla de coherencia: impresiones > clics > resultados por magnitud.
  // Es lo que desempata tres columnas de enteros con nombres parecidos.
  const porCampo = {};
  propuestas.forEach(function (p) { porCampo[p.campo] = p; });
  const orden = ['impresiones', 'clics', 'resultados'];
  for (let i = 0; i < orden.length - 1; i++) {
    const a = porCampo[orden[i]], b = porCampo[orden[i + 1]];
    if (a && b && perfiles[a.indice].prom < perfiles[b.indice].prom) {
      a.confianza = 'baja';  b.confianza = 'baja';
      a.aviso = 'magnitud invertida contra ' + orden[i + 1];
      b.aviso = 'magnitud invertida contra ' + orden[i];
    }
  }

  const sinResolver = objetivo.filter(function (c) { return !resueltos[c]; });
  return {
    tab: tab, filaEncabezado: filaEnc + 1,
    columnas: enc.filter(function (h) { return String(h).trim(); }).length,
    filas: cuerpo.length,
    propuestas: propuestas, sinResolver: sinResolver,
  };
}

/**
 * Corre el análisis y escribe las propuestas en la hoja `Mapeos`
 * para que las revises. Lo que quede en `Mapeos` manda sobre el código.
 */
function proponerMapeo(fuenteId, tienda, cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const r = analizarFuente(ss, fuenteId);

  let sh = ss.getSheetByName('Mapeos');
  if (!sh) {
    sh = ss.insertSheet('Mapeos');
    sh.getRange(1, 1, 1, 7)
      .setValues([['fuente','campo_nova','columna_origen','confianza','aviso','definido_por','fecha']])
      .setFontWeight('bold').setBackground('#0b1824').setFontColor('#c9a84c');
    sh.setFrozenRows(1);
  }

  // Se reemplazan solo las filas de esta fuente; lo que ya confirmaste
  // a mano en otras fuentes no se toca.
  const todo = sh.getDataRange().getValues();
  for (let i = todo.length - 1; i >= 1; i--) {
    if (String(todo[i][0]) === fuenteId && String(todo[i][5]) !== 'humano') {
      sh.deleteRow(i + 1);
    }
  }

  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const filas = r.propuestas.map(function (p) {
    return [fuenteId, p.campo, p.columna, p.confianza, p.aviso || '', 'auto', hoy];
  });
  if (filas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  }

  const msg = [
    'Fuente: ' + fuenteId + '   ·   pestaña ' + r.tab,
    'Encabezado en la fila ' + r.filaEncabezado +
      '  ·  ' + r.columnas + ' columnas  ·  ' + r.filas + ' filas',
    '',
    'PROPUESTAS (revísalas en la hoja Mapeos):',
  ];
  r.propuestas.forEach(function (p) {
    const marca = p.confianza === 'alta' ? '  ' : (p.confianza === 'media' ? ' ~' : ' ?');
    msg.push(marca + ' ' + pad(p.campo, 14) + ' <- ' + p.columna +
             (p.aviso ? '   [' + p.aviso + ']' : ''));
  });
  if (r.sinResolver.length) {
    msg.push('');
    msg.push('SIN RESOLVER: ' + r.sinResolver.join(', '));
    msg.push('Si alguno existe en el archivo, escríbelo a mano en Mapeos');
    msg.push('y pon "humano" en definido_por para que no se sobrescriba.');
  }
  const salida = msg.join('\n');
  Logger.log(salida);
  return salida;
}

function pad(s, n) {
  s = String(s);
  while (s.length < n) s += ' ';
  return s;
}

/**
 * Lee la hoja `Mapeos` y devuelve los alias de una fuente.
 * El importador consulta esto ANTES que los alias del código, para que
 * una corrección tuya gane siempre sobre lo que esté escrito en el .gs
 */
function aliasDesdeMapeos(ss, fuenteId) {
  const sh = ss.getSheetByName('Mapeos');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues().slice(1);
  const alias = {};
  let hay = false;
  filas.forEach(function (f) {
    if (String(f[0]).trim() !== fuenteId) return;
    const campo = String(f[1]).trim();
    const col = String(f[2]).trim();
    if (!campo || !col) return;
    alias[campo] = [col];
    hay = true;
  });
  return hay ? alias : null;
}


/* ═══════════════════════════════════════════════════════════════
   6 · PROVISIONAR CLIENTES
   ═══════════════════════════════════════════════════════════════ */


// La carpeta sale de las propiedades del script, no de una constante:
// así el mismo código sirve en cualquier cuenta.
function carpetaNova() {
  const id = IDS_().carpeta;
  if (id) return DriveApp.getFolderById(id);
  const its = DriveApp.getFoldersByName('Nova');
  if (its.hasNext()) return its.next();
  throw new Error('No encuentro la carpeta Nova. Corre instalarNova() primero.');
}

/**
 * Crea la hoja de un cliente y lo registra en Nova_Central.
 *
 * @param {string} empresa   Nombre del negocio. Ej: 'Nutrea'
 * @param {string} pais      País de la cuenta (para facturación). Ej: 'Colombia'
 * @param {Array}  tiendas   Una fila por tienda:
 *        { id, nombre, marca, pais, moneda, zona, sociedad, nit, corte }
 * @param {Array}  fuentes   Opcional: [{ tienda, fuente, tipo }]
 * @param {string} plan      Opcional. Ej: 'Base'
 * @return {Object} { clienteId, sheetId, url }
 */
function crearCliente(empresa, pais, tiendas, fuentes, plan) {
  if (!empresa) throw new Error('Falta el nombre de la empresa.');
  if (!tiendas || !tiendas.length) throw new Error('Hay que declarar al menos una tienda.');

  // Nombres de tienda duplicados romperían el filtro por tienda en toda la app
  const vistos = {};
  tiendas.forEach(function (t) {
    if (!t.id) throw new Error('Cada tienda necesita un id (ej: "gt").');
    if (vistos[t.id]) throw new Error('Dos tiendas con el mismo id: ' + t.id);
    vistos[t.id] = 1;
  });

  const central = SpreadsheetApp.openById(IDS_().central);
  const shClientes = central.getSheetByName('Clientes');
  if (!shClientes) throw new Error('Corre bootstrapTodo() primero: falta la hoja Clientes.');

  // No duplicar una cuenta que ya existe
  const yaHay = shClientes.getDataRange().getValues().slice(1)
    .filter(function (f) { return norm(f[1]) === norm(empresa); });
  if (yaHay.length) {
    throw new Error('Ya existe un cliente llamado "' + empresa +
                    '" (hoja ' + yaHay[0][13] + '). Bórralo o usa otro nombre.');
  }

  // ── 1. Copiar el template ──
  const carpeta = carpetaNova();
  const copia = DriveApp.getFileById(IDS_().empresarial)
    .makeCopy('Nova_Empresarial_' + empresa, carpeta);
  const sheetId = copia.getId();
  const ss = SpreadsheetApp.openById(sheetId);

  // ── 2. Sembrar las tiendas ──
  const shT = ss.getSheetByName('Tiendas');
  const filasT = tiendas.map(function (t) {
    return [
      t.id, t.nombre || t.id, t.marca || empresa, t.pais || pais || '',
      t.sociedad || '', t.nit || '', t.moneda || '', t.zona || 'UTC',
      t.corte || '16:00', 'activa',
    ];
  });
  shT.getRange(2, 1, filasT.length, 10).setValues(filasT);

  // ── 3. Sembrar las fuentes declaradas ──
  if (fuentes && fuentes.length) {
    const shF = ss.getSheetByName('Fuentes');
    const filasF = fuentes.map(function (f) {
      return [f.tienda, f.fuente, f.tipo || 'pedidos', f.cuenta || '', 'si', '', '', ''];
    });
    shF.getRange(2, 1, filasF.length, 8).setValues(filasF);
  }

  // ── 4. Parámetros por defecto, uno por tienda ──
  // El template los trae con tienda '*'; se materializan por tienda para
  // que cada una pueda tener su propio techo de CPA y su hora de corte.
  const shP = ss.getSheetByName('Parametros');
  const base = shP.getDataRange().getValues().slice(1)
    .filter(function (f) { return String(f[0]).trim() === '*'; });
  if (base.length) {
    const hoy = ahoraISO();
    const expandidos = [];
    tiendas.forEach(function (t) {
      base.forEach(function (p) {
        const valor = (p[1] === 'corte_despacho_hora' && t.corte) ? t.corte : p[2];
        expandidos.push([t.id, p[1], valor, hoy, 'sistema']);
      });
    });
    shP.getRange(shP.getLastRow() + 1, 1, expandidos.length, 5).setValues(expandidos);
  }

  // ── 5. Registrar en Nova_Central ──
  const clienteId = 'C' + String(shClientes.getLastRow()).padStart(4, '0');
  const hoy = ahoraISO();
  shClientes.appendRow([
    clienteId, empresa, pais || '', plan || 'Base', '', '', 'activo',
    hoy, '', '', 0, 0, tiendas.length, sheetId,
  ]);

  const url = 'https://docs.google.com/spreadsheets/d/' + sheetId + '/edit';
  Logger.log([
    'Cliente creado: ' + empresa,
    '  id      : ' + clienteId,
    '  tiendas : ' + tiendas.map(function (t) { return t.nombre || t.id; }).join(', '),
    '  hoja    : ' + url,
  ].join('\n'));

  return { clienteId: clienteId, sheetId: sheetId, url: url };
}

/** Fecha-hora ISO en la zona del script. */
function ahoraISO() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd HH:mm:ss');
}

// ─── TU PROPIA OPERACIÓN ─────────────────────────────────────
/**
 * Crea la hoja de Nutrea. Es una copia del template, igual que la de
 * cualquier cliente — el template se queda limpio.
 *
 * Córrela UNA vez, después de bootstrapTodo().
 */
function crearNutrea() {
  return crearCliente(
    'Nutrea',
    'Colombia',
    [
      { id: 'gt', nombre: 'Nutrea GT', marca: 'Nutrea', pais: 'Guatemala',
        moneda: 'GTQ', zona: 'America/Guatemala', corte: '16:00' },
      { id: 'ec', nombre: 'Nutrea EC', marca: 'Nutrea', pais: 'Ecuador',
        moneda: 'USD', zona: 'America/Guayaquil', corte: '16:00' },
    ],
    [
      { tienda: 'gt', fuente: 'dropi',   tipo: 'pedidos'  },
      { tienda: 'gt', fuente: 'meta',    tipo: 'pauta'    },
      { tienda: 'gt', fuente: 'iris',    tipo: 'llamadas' },
      { tienda: 'ec', fuente: 'dropi',   tipo: 'pedidos'  },
      { tienda: 'ec', fuente: 'shopify', tipo: 'pedidos_secundario' },
      { tienda: 'ec', fuente: 'meta',    tipo: 'pauta'    },
      { tienda: 'ec', fuente: 'iris',    tipo: 'llamadas' },
    ],
    'Interno'
  );
}

/**
 * Resuelve a qué hoja de cliente hay que trabajar.
 *
 * Las funciones de diagnóstico apuntaban al TEMPLATE, que está vacío por
 * diseño — no sirve para revisar nada. Esto resuelve la hoja real:
 *
 *   · sin argumento y hay un solo cliente  -> ese
 *   · con el nombre del cliente            -> el suyo
 *   · con un ID de hoja                    -> ese
 *   · con 'template'                       -> el template, si de verdad lo quieres
 */
function hojaCliente(ref) {
  if (ref === 'template') return IDS_().empresarial;
  if (ref && String(ref).length > 30) return ref; // ya es un ID

  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  const filas = (sh && sh.getLastRow() > 1)
    ? sh.getDataRange().getValues().slice(1).filter(function (f) { return f[0]; })
    : [];

  if (!filas.length) {
    throw new Error('No hay clientes todavía. Corre crearNutrea() primero, ' +
                    'o pasa "template" si de verdad quieres el template vacío.');
  }
  if (ref) {
    const m = filas.filter(function (f) { return norm(f[1]) === norm(ref); });
    if (!m.length) {
      throw new Error('No existe el cliente "' + ref + '". Hay: ' +
                      filas.map(function (f) { return f[1]; }).join(', '));
    }
    return m[0][13];
  }
  if (filas.length === 1) return filas[0][13];

  throw new Error('Hay ' + filas.length + ' clientes, dime cuál: ' +
                  filas.map(function (f) { return '"' + f[1] + '"'; }).join(', '));
}

/** Lista los clientes registrados y a qué hoja apunta cada uno. */
function listarClientes() {
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!sh || sh.getLastRow() < 2) { Logger.log('Sin clientes todavía.'); return []; }
  const filas = sh.getDataRange().getValues().slice(1);
  const out = filas.map(function (f) {
    return { id: f[0], empresa: f[1], estado: f[6], tiendas: f[12], sheetId: f[13] };
  });
  Logger.log(out.map(function (c) {
    return pad(c.id, 7) + pad(c.empresa, 20) + pad(String(c.tiendas) + ' tiendas', 12) + c.sheetId;
  }).join('\n'));
  return out;
}
