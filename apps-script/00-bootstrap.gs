/**
 * Nova · Bootstrap de hojas
 * ─────────────────────────────────────────────────────────────
 * Construye todas las pestañas y encabezados de los 4 workbooks.
 * Se corre UNA VEZ. Es idempotente: si la pestaña ya existe, no la toca.
 *
 * Esquemas tomados de design_handoff_nova/DATOS-Y-ALARMAS.md
 *
 * Cómo usarlo:
 *   1. script.google.com → Nuevo proyecto → pega este archivo
 *   2. Ejecutar → bootstrapTodo()
 *   3. Autorizar cuando lo pida
 */

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
/**
 * NO hay IDs de respaldo a propósito.
 *
 * Antes había unos escritos aquí, y eso hizo que en una instalación el
 * script escribiera en las hojas de OTRA cuenta sin avisar: corrió bien,
 * dijo que todo estaba listo, y construyó las pestañas donde no era.
 *
 * Un respaldo que apunta a la cuenta equivocada es peor que no tener
 * respaldo. Si no hay IDs configurados, esto falla y dice qué hacer.
 */
function IDS_() {
  const p = PropertiesService.getScriptProperties().getProperties();
  if (!p.ID_EMPRESARIAL || !p.ID_CENTRAL || !p.ID_SOUL || !p.ID_ACADEMY) {
    throw new Error(
      'Este script todavía no está instalado en esta cuenta.\n\n' +
      'Corre  instalarNova()  primero: crea la carpeta, las 4 hojas y ' +
      'todas las pestañas en la cuenta en la que estás ahora.\n\n' +
      'No corras bootstrapTodo() directamente — esa función construye ' +
      'pestañas en hojas que ya existen, y sin instalación previa no ' +
      'sabe cuáles son.'
    );
  }
  return {
    empresarial: p.ID_EMPRESARIAL,
    central:     p.ID_CENTRAL,
    soul:        p.ID_SOUL,
    academy:     p.ID_ACADEMY,
    carpeta:     p.ID_CARPETA || '',
  };
}

/** En qué cuenta de Google está corriendo esto. */
function cuentaActual() {
  const email = Session.getEffectiveUser().getEmail();
  Logger.log('Este script corre como: ' + email);
  return email;
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

  // Lo primero que se dice es en qué cuenta se está instalando. Es el dato
  // que más caro sale equivocarse, y el más fácil de no notar.
  const cuenta = Session.getEffectiveUser().getEmail();
  log.push('Instalando en la cuenta: ' + cuenta);
  log.push('');

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
  // Moneda en la que el dueño ve su plata. Puede ser distinta a la de
  // todas sus tiendas: se opera en USD y GTQ, pero se vive en COP.
  ['*', 'moneda_reporte',         'COP', '', ''],
  // Movimiento de tasa a 30 días que dispara aviso (%)
  ['*', 'tasa_alerta_pct_30d',        5, '', ''],
  // Opcional: la tasa con la que montaste el negocio, para medir cuánto
  // se corrió desde entonces. Ej: tasa_referencia_USD = 4000
  ['*', 'tasa_referencia_USD',       '', '', ''],
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
