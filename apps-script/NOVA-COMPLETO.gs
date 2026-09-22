/* ═══════════════════════════════════════════════════════════════
   1 · INSTALACIÓN
   ═══════════════════════════════════════════════════════════════ */

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
  /**
   * `modalidad` dice de dónde sale el inventario, que cambia según cómo
   * trabaje la tienda:
   *
   *   catalogo_publico  el stock es del proveedor, no tuyo. Lo que se
   *                     registra es lo que tú confirmas que hay.
   *   catalogo_privado  el proveedor te pasa un archivo con existencias.
   *   marca_propia      el stock es tuyo y lo llevas tú.
   *
   * Sin esto habría que elegir una sola forma, y la mitad de los clientes
   * tendría una pantalla de inventario que no corresponde a su negocio.
   */
  Tiendas: ['id','nombre','marca','pais','sociedad','nit','moneda',
            'zona_horaria','corte_despacho','modalidad','estado'],
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
            // Lo de oficina: el estado lo dice la transportadora, pero el
            // acuerdo con la clienta y el adelanto los pone el equipo.
            'adelanto','acuerdo_oficina','confirmado_oficina',
            'actualizado_en','actualizado_por'],

  // `solucion` es la instrucción que se le da al courier para resolver la
  // novedad ("dejar en oficina y llamar", "volver a pasar"). No cambia la
  // dirección cargada del cliente — es la salida de la novedad.
  // `grupo` agrupa el motivo (no_contacta / rechaza / direccion / dinero...)
  // y es lo que permite la alarma de patrón: 3 del mismo grupo en la semana
  // es un problema de proceso, no tres casos sueltos.
  // `desenlace` es cómo terminó el pedido de esa novedad. Separado del
  // estado de la novedad a propósito: una novedad puede resolverse y el
  // pedido devolverse igual. Son dos hechos distintos, y confundirlos
  // esconde justo el caso que hay que mirar.
  Novedades: ['id','fuente','id_externo','pedido_id','fecha','tipo','motivo','grupo',
              'estado','solucionada','fecha_solucion','desenlace',
              'gestora','solucion','nota','intentos','resuelta_en',
              'actualizado_en','actualizado_por'],

  // IRIS no es una plataforma de pedidos — es la central telefónica.
  // Cruza con Pedidos por telefono_norm, no por id de orden.
  Llamadas: ['id','fuente','id_externo','fecha_hora','tienda','sentido','estado',
             'extension','agente','telefono','telefono_norm','pedido_id',
             'seg_conversado','seg_espera','seg_total','campana','grabacion',
             'etiqueta','observacion'],
  // `fecha_fin` existe porque Meta no siempre exporta por día: el informe
  // de conjuntos trae una sola fila por todo el periodo. Sin esa columna,
  // un reporte de agosto a septiembre se leería como si todo el gasto
  // hubiera ocurrido el 1 de agosto.
  Pauta: ['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
          'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado',
          'impresiones','alcance','frecuencia','clics','ctr','cpc','cpm',
          'resultados','compras','cpa','roas','valor_conv','visitas_lp'],

  /**
   * Los anuncios, uno por uno. Hoja APARTE de Pauta, y es lo importante.
   *
   * El mismo gasto existe en los dos niveles: un conjunto de 100 con tres
   * anuncios adentro son 100 en Pauta y 100 repartidos aquí. Si vivieran
   * en la misma hoja, cualquier suma contaría el gasto dos veces y el
   * margen saldría a la mitad sin que nada avisara.
   *
   * Separados, cada hoja responde una pregunta distinta y ninguna miente:
   * Pauta dice cuánto cuesta el conjunto —que es donde se mueve el
   * presupuesto— y Anuncios dice cuál creativo está tirando del carro.
   */
  Anuncios: ['id','fecha','tienda','plataforma','cuenta','campana','conjunto',
             'anuncio','anuncio_id','gasto','moneda_gasto',
             'impresiones','alcance','frecuencia','clics','ctr','cpc','cpm',
             'resultados','compras','cpa','valor_conv','visitas_lp'],
  /**
   * Lo que la operación cuesta aunque no se venda nada.
   *
   * Nómina, arriendo, la central telefónica, la suscripción. No está en
   * ningún archivo que exporte una plataforma: lo sabe la dueña y punto.
   * Antes se escribía en la pantalla y se perdía al recargar, así que no
   * se podía usar para nada serio — ni el punto de equilibrio ni la
   * utilidad del mes salían de verdad.
   *
   * `mes` vacío significa que se repite todos los meses. Con un mes
   * concreto, es un gasto de una sola vez: la caja de insumos de marzo no
   * tiene por qué seguir restando en abril.
   */
  Gastos: ['id','tienda','mes','tipo','nombre','valor','moneda','nota',
           'activo','actualizado_en','actualizado_por'],

  /**
   * La factura no es el reporte de campañas, y por eso va aparte.
   *
   * El reporte dice lo que la plataforma contabiliza como gasto. La
   * factura dice lo que te cobraron de verdad, cuándo y en qué moneda —
   * que para quien opera en un país y vive en otro es justo el número que
   * define la utilidad. Guardar solo uno de los dos deja sin respuesta la
   * pregunta de por qué no cuadran.
   */
  Facturacion: ['id','fuente','id_externo','fecha','tienda','plataforma','concepto',
                'gasto','moneda_gasto','gasto_normalizado','moneda_reporte'],

  /**
   * La cartera es el extracto, y el extracto manda.
   *
   * El export de órdenes dice lo que un pedido DEBERÍA costar: un flete
   * estimado, un costo de proveedor de lista. La cartera dice lo que la
   * plataforma de verdad te cobró y te abonó, orden por orden y con
   * fecha. Cuando los dos no coinciden, el que tiene razón es este.
   *
   * En agosto de Nutrea EC eso son 63 cobros de devolución por 315,55 —
   * un promedio de 5,01— contra los 3,50 que traía el export: 95 dólares
   * que ningún cierre estaba contando.
   *
   * `clase` es lo que el movimiento significa, sacado de su descripción:
   * ganancia, devolucion, flete, retiro o recarga. Un retiro NO es gasto
   * —es plata tuya saliendo de la billetera, muchas veces para pagar la
   * pauta— y meterlo como gasto hundiría la utilidad del mes.
   */
  Cartera: ['id','fuente','tienda','fecha','tipo','clase','monto','saldo_previo',
            'orden_id','guia','descripcion','cuenta','concepto_retiro',
            'importado_en'],

  /**
   * CAS: el ticket que se le radica a la transportadora.
   *
   * Cuando un pedido lleva días sin cambiar de estado, llamar al cliente
   * no sirve —el paquete no está con él—. Lo que mueve la aguja es
   * presión administrativa sobre la transportadora: un ticket oficial
   * pidiendo prioridad de despacho.
   *
   * Nova detecta los candidatos sola, leyendo los pedidos quietos. Lo que
   * no puede saber es si alguien radicó el ticket, con qué número, y qué
   * contestaron. Eso lo escribe el equipo, y es justo lo que hoy vive en
   * un Drive suelto donde nadie más lo ve.
   *
   * `estado`: abierto · respondido · resuelto · sin_respuesta
   */
  CAS: ['id','tienda','pedido_id','id_externo','guia','transportadora',
        'abierto_en','abierto_por','ticket','estado','dias_quieto',
        'ultima_gestion','respuesta','cerrado_en','nota'],

  /**
   * `id` para poder editar una fila desde la app, y `origen` para saber
   * si ese número lo contó una persona o lo trajo un archivo. Mezclarlos
   * sin distinguir hace imposible saber en cuál confiar.
   *
   * El precio NO es uno solo. Casi nadie vende una unidad suelta: hay
   * promoción de dos y de tres, y esa es la que de verdad se paga. Con un
   * solo precio, quien gestiona tiene que acordarse de memoria de cuánto
   * vale el combo, y el margen por unidad que calcula Nova se parece poco
   * a lo que entra por pedido.
   *
   * Por eso van tres columnas y no una lista: se leen de un vistazo en la
   * hoja y se editan sin sintaxis. Vacías quedan apagadas — una tienda
   * que solo vende unidades sueltas no tiene por qué llenarlas.
   *
   * `landing` es la página desde donde se vende. Quien confirma un pedido
   * la necesita abierta para responder precio y beneficios sin colgar, y
   * hoy vive en la cabeza de cada quien o en un chat viejo.
   *
   * `categoria` es lo que la dueña decide que es ese producto —estrella,
   * complemento, testeo, frenado— y no tiene nada que ver con `origen`.
   * Estuvieron en la misma columna un tiempo, y eso obligaba a elegir
   * entre saber de dónde salió el número o saber qué papel juega el
   * producto. Son dos preguntas distintas y ahora tienen dos columnas.
   */
  Inventario: ['id','sku','producto','tienda','fuente','origen','categoria',
               'proveedor','landing','stock','costo_unitario',
               'precio','precio_2','precio_3','minimo',
               // Las respuestas a las cuatro preguntas de siempre, cada
               // una en su columna. Qué pregunta es cada una lo decide la
               // tienda en Parametros → preguntas_producto.
               'resp_1','resp_2','resp_3','resp_4',
               'dias_cobertura','ultimo_conteo','nota','activo',
               'actualizado_en','actualizado_por'],
  // "permisos" es lo que la dueña decide que esta persona puede hacer,
  // separado por comas. Vacío = lo que el rol trae por defecto.
  // Ver PERMISOS_POR_ROL en 60-api.gs.
  Equipo: ['id','nombre','correo','rol','tienda','estado','casos_asignados',
           'casos_resueltos','nota_auditoria','ultima_conexion','permisos'],

  /**
   * Los estados que cada plataforma inventa, y qué significan aquí.
   *
   * Existe para que agregar un estado nuevo NO exija publicar una versión
   * del código. Las transportadoras cambian nombres sin avisar, y esperar
   * a un desarrollador para volver a cerrar un mes es un cuello de
   * botella que no tiene por qué existir.
   *
   * `origen` dice de dónde salió la traducción, y es lo que permite
   * auditarla:
   *   catalogo  · de las tablas que trae Nova
   *   deducido  · Nova lo dedujo del texto. Solo pasa con tránsito:
   *               nunca se deduce un entregado ni una devolución
   *   manual    · lo decidió la dueña. Manda sobre todo lo demás
   *   nuevo     · Nova no sabe qué es y está esperando respuesta
   *
   * `pedidos` es cuántos hay con ese estado. Sirve para priorizar: uno
   * suelto es ruido, doscientos es un cierre mal hecho esperando a pasar.
   */
  Estados: ['fuente','texto','estado_nova','origen','pedidos',
            'primera_vez','ultima_vez','decidido_por','nota'],

  // Un mes no cierra el día 31: cierra cuando los pedidos de ese mes ya
  // se resolvieron. Un pedido del 28 de agosto se entrega el 5 de
  // septiembre, y hasta que eso pase la tasa de entrega y el margen de
  // agosto son provisionales.
  //
  // Al cerrar se CONGELAN las cifras. Si se recalcularan siempre, el
  // agosto que reportaste en septiembre cambiaría en octubre cuando una
  // devolución vieja por fin se resuelva — y un número que cambia solo
  // no sirve para decidir ni para rendir cuentas.
  /**
   * `snapshot` guarda el cierre entero, no solo el resumen.
   *
   * Las columnas sueltas se quedaron cortas: no había dónde poner el costo
   * de mercancía, el flete, los gastos fijos ni el desglose por estado. Al
   * reabrir un mes cerrado, esos costos volvían como cero y la utilidad
   * salía enorme — el informe mostraba ventas menos pauta y nada más.
   *
   * Las columnas sueltas se mantienen porque son las que se leen de un
   * vistazo en la hoja. El snapshot es la verdad completa.
   */
  Cierres: ['tienda','mes','estado','cerrado_en','cerrado_por',
            'pendientes_al_cierre','pedidos','entregados','devueltos',
            'ventas','gasto','margen','efectividad','nota','snapshot'],
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

  /**
   * Quién puede entrar a Nova Central.
   *
   * Es una hoja aparte de Equipo a propósito. Equipo son las personas de
   * UN cliente: la dueña de Nutrea, sus admins, sus gestoras. Esto es
   * quien opera la plataforma — quien crea cuentas, ve la facturación de
   * todos y enciende el demo. Mezclarlas haría que darle permisos a una
   * admin de un cliente pudiera, por un descuido, abrirle la consola de
   * todos los clientes.
   *
   * `rol`: socia (todo) · operadora (crea y acompaña clientes, no toca
   * planes ni tarifas).
   */
  Plataforma: ['id','nombre','correo','rol','estado','ultima_conexion','nota'],
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
  log.push(actualizarClientes());
  Logger.log(log.join('\n'));
  return log.join('\n');
}

/**
 * Pone al día las hojas de los clientes que ya existen.
 *
 * El template es el molde, no la operación: las tiendas de verdad viven
 * en las copias. Cuando el esquema crece —una pestaña Cierres, una
 * columna permisos— el template se actualiza y las copias se quedan
 * atrás, así que la función nueva sirve en una instalación recién hecha
 * y falla en la que lleva meses trabajando. Que es al revés de lo que
 * uno quiere.
 *
 * Solo agrega pestañas y columnas que faltan. No borra ni reordena ni
 * toca una sola celda de datos.
 */
function actualizarClientes() {
  const central = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!central || central.getLastRow() < 2) return 'Clientes: ninguno registrado todavía';

  const filas = central.getDataRange().getValues();
  const enc = filas[0].map(function (h) { return String(h || '').trim().toLowerCase(); });
  const cId = enc.indexOf('sheet_id') !== -1 ? enc.indexOf('sheet_id') : 13;
  const cNom = enc.indexOf('empresa') !== -1 ? enc.indexOf('empresa') : 1;

  const out = [];
  for (let i = 1; i < filas.length; i++) {
    const id = String(filas[i][cId] || '').trim();
    if (!id) continue;
    const nombre = String(filas[i][cNom] || id).trim();
    try {
      out.push('  ' + construir(id, nombre, ESQUEMA_EMPRESARIAL, IMPORTS_EMPRESARIAL));
    } catch (e) {
      // Una hoja borrada o sin permiso no puede detener a las demás
      out.push('  ' + nombre + ': NO se pudo abrir (' + e.message + ')');
    }
  }
  return 'Clientes:\n' + (out.length ? out.join('\n') : '  ninguno con hoja');
}

function construir(fileId, nombre, esquema, importsCrudos) {
  const ss = SpreadsheetApp.openById(fileId);
  const creadas = [];

  // Entidades + las dos comunes obligatorias
  const todas = Object.assign({}, esquema, COMUNES);

  Object.keys(todas).forEach(function (tab) {
    if (crearTab(ss, tab, todas[tab])) creadas.push(tab);
    else {
      const nuevas = agregarColumnasFaltantes(ss, tab, todas[tab]);
      if (nuevas.length) creadas.push(tab + ' (+' + nuevas.join(', ') + ')');
    }
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

/**
 * Agrega al final las columnas que el esquema tiene y la hoja todavía no.
 *
 * El esquema crece: cuando se agregó "permisos" a Equipo, las hojas ya
 * creadas se quedaron sin esa columna y crearTab() no las tocaba porque
 * la pestaña ya existía. El resultado era una función nueva que no
 * funcionaba en las cuentas viejas y sí en las nuevas.
 *
 * Solo agrega. Nunca renombra ni reordena ni borra: si alguien movió una
 * columna de sitio o le puso otro nombre, esa decisión se respeta y la
 * columna que falta se añade al final.
 */
function agregarColumnasFaltantes(ss, nombre, encabezados) {
  const sh = ss.getSheetByName(nombre);
  if (!sh) return [];

  const ancho = sh.getLastColumn();
  const actuales = ancho
    ? sh.getRange(1, 1, 1, ancho).getValues()[0]
        .map(function (h) { return String(h || '').trim().toLowerCase(); })
    : [];

  const faltan = encabezados.filter(function (h) {
    return actuales.indexOf(String(h).trim().toLowerCase()) === -1;
  });
  if (!faltan.length) return [];

  if (sh.getMaxColumns() < ancho + faltan.length) {
    sh.insertColumnsAfter(Math.max(ancho, 1), faltan.length);
  }
  sh.getRange(1, ancho + 1, 1, faltan.length)
    .setValues([faltan])
    .setFontWeight('bold')
    .setBackground('#0b1824')
    .setFontColor('#c9a84c');
  return faltan;
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
   2 · ESTADOS Y TELÉFONOS
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Estados canónicos y normalización de teléfono
 * ─────────────────────────────────────────────────────────────
 * Construido a partir de los datos REALES de la carpeta QKF + NOVA:
 *   · Dropi-Pedidos-NutreaShop.xlsx  (Maestro Ecuador, 569 pedidos)
 *   · reporte-historial-de-pedidos-*.xlsx  (Mastershop/Effi Colombia)
 *   · IRIS (1).csv  (central telefónica, 2.782 llamadas)
 *
 * Cada plataforma nombra los estados distinto. Sin una tabla canónica,
 * la alarma de "entrega bajo 65%" cuenta mal según de qué fuente venga.
 */

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

  /**
   * No sabemos qué pasó con este pedido.
   *
   * Va aparte de `pendiente` a propósito, porque no son lo mismo:
   * pendiente es "todavía no se resolvió", una afirmación sobre el
   * pedido. Sin clasificar es "la transportadora usó una palabra que no
   * reconozco", una afirmación sobre NOVA.
   *
   * Confundirlos escondía el problema: los desconocidos se sumaban a los
   * pendientes y nadie volvía a mirarlos. Separados, se pueden contar,
   * mostrar y preguntar.
   */
  SIN_CLASIFICAR:   'sin_clasificar',
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

    /**
     * VERIFICADO contra un archivo real de Dropi Colombia (871 órdenes,
     * abril a septiembre de 2026, 18 estados distintos).
     *
     * Los once de abajo no estaban. Y un estado que no se reconoce no da
     * error: cae en "pendiente", que es el estado de lo que todavía no
     * tiene desenlace. El efecto es silencioso y caro — en ese archivo,
     * 13 pedidos RECHAZADO y 19 RECLAME EN OFICINA se habrían contado
     * como "en camino" para siempre, el mes nunca habría cerrado, y la
     * tasa de entrega habría salido más baja de lo que fue.
     */

    // El cliente lo rechazó en la puerta. Es una devolución: el paquete
    // se devuelve y el flete de retorno se paga igual. Contarlo como
    // cancelado escondería el costo, y como pendiente, el desenlace.
    'rechazado':                          ESTADOS.DEVOLUCION,
    'rechazada':                          ESTADOS.DEVOLUCION,

    // La guía se anuló antes de despachar: nadie movió nada, nadie cobró.
    'guia_anulada':                       ESTADOS.CANCELADO,
    'guia anulada':                       ESTADOS.CANCELADO,

    // Va a recogerlo el cliente a la oficina. No está perdido, pero
    // tampoco entregado: es el caso que Nova sigue en "Seguimiento a
    // oficina", donde se decide si se espera o se devuelve.
    'reclame en oficina':                 ESTADOS.EN_OFICINA,
    'reclamo en oficina':                 ESTADOS.EN_OFICINA,

    // Todavía en manos de la bodega
    'en procesamiento':                   ESTADOS.CONFIRMADO,
    'preparado para transportadora':      ESTADOS.EN_BODEGA,
    'en bodega transportadora':           ESTADOS.EN_BODEGA,

    // Ya salió: desde aquí el flete está causado
    'despachada':                         ESTADOS.EN_TRANSITO,
    'despachado':                         ESTADOS.EN_TRANSITO,
    'en bodega destino':                  ESTADOS.EN_TRANSITO,
    'en terminal destino':                ESTADOS.EN_TRANSITO,
    'en reexpedicion':                    ESTADOS.EN_TRANSITO,
    'en espera de ruta domestica':        ESTADOS.EN_TRANSITO,
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

/**
 * ═══════════════════════════════════════════════════════════
 *  ESTADOS QUE NOVA NO CONOCE
 * ═══════════════════════════════════════════════════════════
 *
 * Las tablas de arriba son el punto de partida, no la verdad completa.
 * Cada transportadora inventa nombres y los cambia sin avisar: en un solo
 * archivo real de Dropi Colombia aparecieron dieciocho estados distintos
 * y Nova reconocía siete.
 *
 * Antes, un estado desconocido acababa contado como "pendiente". No daba
 * error, y esa es la parte cara: trece pedidos RECHAZADO —que ya
 * volvieron y cuyo flete ya se pagó— quedaban como "en camino" para
 * siempre, el mes nunca cerraba, y la tasa de entrega salía más baja de
 * lo que fue. Un mes entero podía estar mal sin un solo síntoma.
 *
 * LA LÍNEA NO ES "IMPORTANTE O BOBO". ES SI EL ESTADO CIERRA EL PEDIDO.
 *
 * Los estados de tránsito —en bodega destino, en reexpedición, preparado
 * para transportadora— significan todos lo mismo para las cuentas: el
 * pedido sigue vivo. Equivocarse entre ellos no mueve un solo número, así
 * que Nova los deduce sola cuando el texto lo delata, y deja constancia.
 *
 * Los terminales —entregado, devolución, cancelado— mueven la plata en
 * direcciones opuestas. Adivinar "entregado" es inventar una venta y
 * contaminar el ingreso, el margen y el ROAS. Adivinar "devolución" es
 * cobrar un flete que nadie pagó. Así que NOVA NUNCA LOS ADIVINA: deja
 * el pedido en `sin_clasificar`, lo cuenta aparte, y pregunta.
 *
 * Y lo que la dueña clasifica se guarda en la hoja, no en el código: se
 * aplica al instante y sin que nadie publique una versión nueva.
 */

/** Palabras que solo aparecen en estados de tránsito. */
const PISTAS_TRANSITO = [
  ['bodega',        ESTADOS.EN_BODEGA],
  ['almacen',       ESTADOS.EN_BODEGA],
  ['preparad',      ESTADOS.EN_BODEGA],
  ['alistamiento',  ESTADOS.EN_BODEGA],
  ['procesamiento', ESTADOS.CONFIRMADO],
  ['procesando',    ESTADOS.CONFIRMADO],
  ['reparto',       ESTADOS.EN_TRANSITO],
  ['transito',      ESTADOS.EN_TRANSITO],
  ['transporte',    ESTADOS.EN_TRANSITO],
  ['despach',       ESTADOS.EN_TRANSITO],
  ['ruta',          ESTADOS.EN_TRANSITO],
  ['terminal',      ESTADOS.EN_TRANSITO],
  ['distribucion',  ESTADOS.EN_TRANSITO],
  ['reexpedicion',  ESTADOS.EN_TRANSITO],
  ['camino',        ESTADOS.EN_TRANSITO],
  ['oficina',       ESTADOS.EN_OFICINA],
  ['agencia',       ESTADOS.EN_OFICINA],
  ['sucursal',      ESTADOS.EN_OFICINA],
];

/**
 * Palabras que aparecen en estados TERMINALES y prohíben deducir nada.
 *
 * Van primero y mandan sobre las pistas de tránsito. "DEVUELTO DESDE
 * BODEGA DESTINO" tiene la palabra bodega, pero es una devolución: sin
 * este freno, Nova lo daría por vivo y nunca cobraría su flete de retorno.
 */
const PISTAS_TERMINALES = ['entrega', 'devol', 'devuel', 'cancel', 'anulad',
                           'rechaz', 'reembols', 'perdid', 'siniestr', 'indemniz'];

/**
 * Lo que Nova puede deducir sin riesgo, o cadena vacía.
 * Solo tránsito, nunca un desenlace.
 */
function deducirTransito(k) {
  for (let i = 0; i < PISTAS_TERMINALES.length; i++) {
    if (k.indexOf(PISTAS_TERMINALES[i]) !== -1) return '';
  }
  for (let j = 0; j < PISTAS_TRANSITO.length; j++) {
    if (k.indexOf(PISTAS_TRANSITO[j][0]) !== -1) return PISTAS_TRANSITO[j][1];
  }
  return '';
}

/**
 * Traduce el estado de una plataforma al canónico de Nova.
 *
 * `aprendidos` es lo que ya está clasificado en la hoja Estados, con la
 * forma { 'dropi|rechazado': 'devolucion' }. Va primero porque una
 * decisión de la dueña gana sobre cualquier tabla de aquí: si ella dice
 * que en SU operación ese estado significa otra cosa, tiene razón.
 */
function estadoCanonico(fuente, texto, aprendidos) {
  if (!texto) return '';
  const k = norm(texto);

  if (aprendidos && aprendidos[fuente + '|' + k]) return aprendidos[fuente + '|' + k];

  const mapa = MAPA_ESTADOS[fuente] || MAPA_ESTADOS[fuente === 'effi' ? 'mastershop' : ''] || {};
  if (mapa[k]) return mapa[k];

  // Coincidencia por prefijo: los couriers agregan sufijos
  // ("ENTREGADA DIGITALIZADA", "PARA RETIRO EN AGENCIA SERVIENTREGA")
  const claves = Object.keys(mapa);
  for (let i = 0; i < claves.length; i++) {
    if (k.indexOf(claves[i]) === 0) return mapa[claves[i]];
  }

  const deducido = deducirTransito(k);
  if (deducido) return deducido;

  return ESTADOS.SIN_CLASIFICAR;
}

/**
 * Igual que la anterior, pero además dice CÓMO lo resolvió.
 * La usa el importador para saber qué anotar en la hoja Estados.
 */
function estadoConOrigen(fuente, texto, aprendidos) {
  if (!texto) return { estado: '', origen: '' };
  const k = norm(texto);

  if (aprendidos && aprendidos[fuente + '|' + k]) {
    return { estado: aprendidos[fuente + '|' + k], origen: 'guardado', clave: k };
  }
  const mapa = MAPA_ESTADOS[fuente] || MAPA_ESTADOS[fuente === 'effi' ? 'mastershop' : ''] || {};
  if (mapa[k]) return { estado: mapa[k], origen: 'catalogo', clave: k };
  const claves = Object.keys(mapa);
  for (let i = 0; i < claves.length; i++) {
    if (k.indexOf(claves[i]) === 0) return { estado: mapa[claves[i]], origen: 'catalogo', clave: k };
  }
  const deducido = deducirTransito(k);
  if (deducido) return { estado: deducido, origen: 'deducido', clave: k };

  return { estado: ESTADOS.SIN_CLASIFICAR, origen: 'nuevo', clave: k };
}

// ─── LA HOJA DE ESTADOS ──────────────────────────────────────

/**
 * Lo que ya está clasificado, como { 'dropi|rechazado': 'devolucion' }.
 *
 * Solo cuentan las filas con `estado_nova` puesto: una fila con estado
 * vacío es justamente una pregunta sin responder, y devolverla como
 * traducción haría que un estado desconocido se resolviera en nada.
 */
function estadosAprendidos(ss) {
  const sh = ss.getSheetByName('Estados');
  const out = {};
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cF = e.indexOf('fuente'), cT = e.indexOf('texto'), cE = e.indexOf('estado_nova');
  if (cF === -1 || cT === -1 || cE === -1) return out;

  for (let i = 1; i < d.length; i++) {
    const est = norm(d[i][cE]);
    if (!est || est === ESTADOS.SIN_CLASIFICAR) continue;
    out[norm(d[i][cF]) + '|' + norm(d[i][cT])] = est;
  }
  return out;
}

/**
 * Anota los estados que se vieron en una importación.
 *
 * Se escribe SIEMPRE, no solo cuando Nova no entiende: saber que un
 * estado conocido dejó de aparecer también dice algo, y tener el listado
 * completo con sus cuentas es lo que permite decidir cuál mirar primero.
 *
 * Lo que la dueña ya decidió no se pisa nunca. Solo se le actualiza la
 * cuenta y la última fecha.
 */
function anotarEstados(ss, vistos) {
  const claves = Object.keys(vistos);
  if (!claves.length) return;

  let sh = ss.getSheetByName('Estados');
  if (!sh) {
    sh = ss.insertSheet('Estados');
    sh.getRange(1, 1, 1, HOJAS.Estados.length).setValues([HOJAS.Estados]);
  }

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const hoy = ahoraISO().slice(0, 10);

  const donde = {};
  for (let i = 1; i < d.length; i++) {
    donde[norm(d[i][c('fuente')]) + '|' + norm(d[i][c('texto')])] = i;
  }

  const nuevas = [];
  claves.forEach(function (k) {
    const v = vistos[k];
    const i = donde[k];

    if (i === undefined) {
      const campos = {
        fuente: v.fuente, texto: v.texto,
        // Un estado nuevo entra SIN traducción: la casilla vacía es la
        // pregunta. Escribirle algo sería responderla por la dueña.
        estado_nova: v.origen === 'nuevo' ? '' : v.estado,
        origen: v.origen, pedidos: v.n,
        primera_vez: hoy, ultima_vez: hoy, decidido_por: '', nota: '',
      };
      nuevas.push(e.map(function (col) {
        return campos[col] === undefined ? '' : campos[col];
      }));
      return;
    }

    d[i][c('pedidos')] = num(d[i][c('pedidos')]) + v.n;
    d[i][c('ultima_vez')] = hoy;
    // Una decisión manual no se toca jamás
    if (norm(d[i][c('origen')]) !== 'manual' && !String(d[i][c('estado_nova')]).trim()) {
      d[i][c('estado_nova')] = v.origen === 'nuevo' ? '' : v.estado;
      d[i][c('origen')] = v.origen;
    }
  });

  if (d.length > 1) sh.getRange(1, 1, d.length, d[0].length).setValues(d);
  if (nuevas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, nuevas[0].length).setValues(nuevas);
  }
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
   3 · MONEDAS
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Monedas y conversión
 * ─────────────────────────────────────────────────────────────
 * Dos cosas distintas que se confunden fácil:
 *
 *   1. El CATÁLOGO — cuántos decimales tiene cada moneda y cómo se muestra.
 *      Es una tabla estática. Barata. Está completa para Latinoamérica.
 *
 *   2. La TASA — cuánto vale esa moneda un día concreto. Vive en la hoja
 *      `Tasas` y hay que alimentarla. Sin tasa NO hay conversión posible,
 *      por más completo que esté el catálogo.
 *
 * El spec es explícito: se convierte con "la tasa del DÍA DE LA TRANSACCIÓN,
 * no la de hoy". Si no, los márgenes salen mal y todas las alarmas de dinero
 * disparan en falso.
 *
 * Por eso, cuando falta una tasa, la conversión NO inventa un número:
 * devuelve null y marca la fila. Un margen vacío se ve; un margen calculado
 * con la tasa equivocada no se ve y miente.
 */

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
   4 · FUENTES E IMPORTADORES
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Importadores
 * ─────────────────────────────────────────────────────────────
 * Convierte las pestañas _Import_* (crudas) en filas normalizadas.
 *
 * Regla del spec:
 *   "un cambio de formato rompe UNA función de importación, no seis pantallas"
 *
 * Por eso el mapeo es declarativo. Agregar una fuente nueva es agregar
 * un bloque a FUENTES — no escribir un importador nuevo.
 *
 * ⚠ ESTADO DE LOS MAPEOS
 *   Los alias de columna de abajo son la primera aproximación. Cada
 *   plataforma nombra distinto la misma cosa y cambia los nombres sin avisar.
 *   Marcados VERIFICADO los que ya se cotejaron contra un export real.
 *   Marcados POR VERIFICAR los que hay que confirmar con un CSV de muestra.
 */

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
      // Dropi ya dice si la novedad se resolvió y cuándo. Sin leer esto,
      // toda novedad de un pedido que después se entregó quedaba abierta
      // para siempre: 72 novedades y 72 "abiertas".
      solucionada:      ['fue solucionada la novedad'],
      fecha_solucion:   ['fecha de solucion'],
      ultimo_movimiento:['ultimo movimiento'],
      fecha_ingreso:    ['fecha ingreso'],
      actualizado:      ['ultima actualizacion'],
      direccion:        ['direccion', 'address'],
      sku:              ['sku', 'codigo'],
      // VERIFICADO contra ordenes_productos_20260911 (53 columnas, Nutrea
      // EC). Ese export nombra el flete "precio flete" y el costo del
      // producto "precio proveedor": sin estos dos, el margen por pedido
      // sale mal y no hay aviso de que falte nada.
      costo_envio:      ['flete', 'costo envio', 'envio', 'precio flete'],
      costo_producto:   ['precio proveedor', 'costo proveedor',
                         'precio proveedor x cantidad'],
      departamento:     ['departamento destino', 'departamento', 'provincia'],
      correo:           ['email', 'correo'],
      cedula:           ['nro de identificacion', 'numero de identificacion',
                         'identificacion', 'cedula', 'documento'],
      // Va a `observacion`, NO a `nota`: nota es de la gestora y el
      // importador tiene prohibido pisarla. Además de aquí se saca el
      // segundo número cuando la clienta pidió que la llamen a otro.
      observacion:      ['notas', 'nota', 'observacion'],
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

  // ── EL HISTÓRICO DEL PROPIO CLIENTE ────────────────────────
  /**
   * Casi nadie llega en cero. Llega con su control diario en Excel,
   * hecho a mano, con los encabezados que a esa persona le hicieron
   * sentido: "Tel", "A quién se le entregó", "Plata".
   *
   * Esta fuente no trae alias porque no hay alias que valgan: cada
   * archivo es distinto. El mapeo se propone leyendo el archivo y lo
   * confirma la persona antes de importar, y queda guardado en Mapeos
   * para las siguientes veces.
   */
  propio: {
    tipo: 'pedidos',
    propio: true,
    alias: {},
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
  /**
   * El historial de cartera de Dropi.
   *
   * VERIFICADO contra historial_de_cartera-12-09-2026 (Nutrea EC, 275
   * movimientos entre agosto y septiembre).
   *
   * Es el único archivo que dice lo que de verdad se movió: cuánto te
   * abonaron por cada orden entregada, cuánto te cobraron por cada
   * devolución y cada flete, y cuánto retiraste. El export de órdenes
   * trae estimados; esto es el extracto.
   */
  dropi_cartera: {
    tipo: 'cartera',
    verificado: true,
    tab: '_Import_Cartera',
    alias: {
      id_externo:      ['id'],
      fecha:           ['fecha'],
      tipo_movimiento: ['tipo'],
      monto:           ['monto'],
      saldo_previo:    ['monto previo'],
      orden_id:        ['orden id', 'id orden'],
      guia:            ['numero de guia', 'numero guia', 'guia'],
      descripcion:     ['descripcion'],
      cuenta:          ['cuenta'],
      concepto_retiro: ['concepto de retiro'],
    },
  },

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
/**
 * Cómo leer una fecha ambigua, mientras dura la importación.
 *
 * aISO se llama desde una docena de sitios y pasarle el ajuste por
 * parámetro obligaría a tocarlos todos —incluidos los que ni siquiera
 * tienen la hoja a mano—. Así que el valor se pone una vez, antes de
 * importar, y se quita al terminar.
 *
 * Solo cambia lo que de verdad es ambiguo: 25-12-2026 es 25 de diciembre
 * en cualquier convención, y AAAA-MM-DD no admite discusión.
 */
let FORMATO_FECHA = 'dia_primero';

function conFormatoFecha(ss, tienda, fn) {
  const previo = FORMATO_FECHA;
  try {
    const a = ajustes(ss, tienda);
    FORMATO_FECHA = a.formato_fecha === 'mes_primero' ? 'mes_primero' : 'dia_primero';
    return fn();
  } finally { FORMATO_FECHA = previo; }
}

function aISO(v, zonaHoraria) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, zonaHoraria || 'UTC', 'yyyy-MM-dd');
  }
  const s = String(v).trim();

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);           // ya ISO
  if (m) return m[1] + '-' + m[2] + '-' + m[3];

  /**
   * Día primero. En toda América Latina 03/09 es 3 de septiembre.
   *
   * Si el primer número pasa de 12 no hay duda; si no, manda la
   * convención de la región. Lo que NO se hace es dejárselo a new Date(),
   * que asume el formato de Estados Unidos y convierte medio mes de
   * pedidos en fechas de otro mes sin avisar.
   */
  m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{4})/);
  if (m) {
    const a = parseInt(m[1], 10), b = parseInt(m[2], 10);
    let dia, mesN;
    if (a > 12 && b <= 12)      { dia = a; mesN = b; }   // solo puede ser día-mes
    else if (b > 12 && a <= 12) { dia = b; mesN = a; }   // solo puede ser mes-día
    else if (FORMATO_FECHA === 'mes_primero') { mesN = a; dia = b; }
    else                        { dia = a; mesN = b; }   // la convención de la región
    return m[3] + '-' + ('0' + mesN).slice(-2) + '-' + ('0' + dia).slice(-2);
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

  // Una pestaña por tienda cuando hace falta: si la misma plataforma
  // sirve a dos tiendas, cada una necesita su propio export. Se busca
  // primero _Import_Dropi_EC y se cae a _Import_Dropi si no existe.
  const base = cfg.tab ||
    ('_Import_' + fuenteId.charAt(0).toUpperCase() + fuenteId.slice(1));
  const propia = base + '_' + String(tienda || '').toUpperCase();
  const sh = ss.getSheetByName(propia) || ss.getSheetByName(base);
  if (!sh) throw new Error('Falta la pestaña ' + base + ' (ni ' + propia + ').');
  const tab = sh.getName();

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
/** Las tiendas de un cliente, en el orden en que están en la hoja. */
function tiendasDeCliente(ss) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const c = d[0].map(norm).indexOf('id');
  return d.slice(1).map(function (f) { return String(f[c]).trim(); }).filter(String);
}

function diagnosticar(fuenteId, tienda, cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  // Sin tienda se toma la primera del cliente, no una fija
  const t = tienda || (tiendasDeCliente(ss)[0] || '');
  const r = leerCrudo(ss, fuenteId, t);
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

/**
 * Qué significa un movimiento de cartera.
 *
 * Dropi no trae una columna con la clase: la describe en texto libre, con
 * mayúsculas y el número de orden pegado al final. Las cinco frases de
 * abajo son las que aparecen en el historial de Nutrea EC, y cubren los
 * 275 movimientos del archivo de muestra sin dejar ninguno en "otro".
 *
 * Se clasifica por la frase, no por ENTRADA/SALIDA: una recarga y una
 * ganancia son las dos entradas, pero una es plata que metiste tú y la
 * otra es plata que ganaste. Sumarlas juntas diría que el mes vendió el
 * doble.
 */
function claseMovimiento(descripcion, conceptoRetiro) {
  const d = norm(descripcion || '');
  if (d.indexOf('ganancia en la orden') !== -1)   return 'ganancia';
  if (d.indexOf('cobro de devolucion') !== -1)    return 'devolucion';
  if (d.indexOf('flete inicial') !== -1)          return 'flete';
  if (d.indexOf('cobro de flete') !== -1)         return 'flete';
  if (d.indexOf('retiro') !== -1)                 return 'retiro';
  if (d.indexOf('recarga') !== -1)                return 'recarga';
  if (d.indexOf('reversion') !== -1)              return 'reversion';
  // Un retiro puede venir descrito de otra forma pero traer concepto
  if (String(conceptoRetiro || '').trim())        return 'retiro';
  return 'otro';
}


/* ═══════════════════════════════════════════════════════════════
   5 · MAPEO AUTOMÁTICO
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Auto-mapeo de fuentes desconocidas
 * ─────────────────────────────────────────────────────────────
 * El problema: no siempre hay un export de muestra. TikTok hoy, y mañana
 * cualquier plataforma nueva — o Meta renombrando una columna sin avisar.
 *
 * La solución no es adivinar nombres, es leer la FORMA de los datos.
 * Una columna de gasto se reconoce porque es numérica, positiva, con
 * decimales y magnitud media — no porque se llame "Importe gastado".
 *
 * Flujo:
 *   1. Pegas el export en su pestaña _Import_*
 *   2. Corres proponerMapeo('tiktok', 'gt')
 *   3. Escribe sus propuestas en la hoja `Mapeos` con un nivel de confianza
 *   4. Revisas y corriges lo que esté mal — sin tocar código
 *   5. El importador usa `Mapeos` por encima de los alias del código
 *
 * Esto hace que agregar una plataforma sea trabajo tuyo de dos minutos,
 * no un cambio de código que tienes que esperar.
 */

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
function puntajeNombre(campo, encabezado, dicc) {
  const h = norm(encabezado);
  if (!h) return 0;
  const syns = (dicc || SINONIMOS)[campo] || [];
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

// ─── ARCHIVOS PROPIOS DEL CLIENTE ────────────────────────────
/**
 * El histórico que ya tiene el cliente.
 *
 * Casi nadie empieza en cero: llega con su control diario en Excel,
 * hecho a mano, con los encabezados que a esa persona le hicieron
 * sentido. "Tel", "Cel de la clienta", "A quién se le entregó".
 * Ningún catálogo de alias va a cubrir eso.
 *
 * Por eso este diccionario no pretende acertar siempre: pretende
 * proponer, y que la persona confirme antes de que se escriba nada.
 * Un histórico mal leído es peor que no cargarlo.
 */
const SINONIMOS_PEDIDOS = {
  id_externo:   ['id','no','num','numero','orden','pedido','no orden','numero orden',
                 'numero de orden','guia','order','order id','referencia','codigo'],
  fecha:        ['fecha','dia','date','fecha pedido','fecha de pedido','fecha creacion',
                 'creado','registro','fecha venta'],
  cliente:      ['cliente','nombre','nombre cliente','nombre del cliente','clienta',
                 'destinatario','comprador','nombre completo','customer'],
  telefono:     ['telefono','tel','celular','cel','movil','whatsapp','contacto',
                 'numero','numero de contacto','phone'],
  telefono_2:   ['telefono 2','tel 2','celular 2','segundo numero','otro numero',
                 'numero alterno','telefono alterno','contacto 2','whatsapp 2'],
  cedula:       ['cedula','documento','dni','identificacion','ci','nit','rut'],
  correo:       ['correo','email','mail','e mail'],
  direccion:    ['direccion','dir','domicilio','address','direccion de entrega'],
  ciudad:       ['ciudad','city','municipio','localidad','destino','ciudad destino'],
  // 'estado' NO va aquí aunque en México signifique región: en una
  // tienda esa columna es el estado del pedido nueve de cada diez veces,
  // y si departamento se la queda, los pedidos entran sin estado.
  departamento: ['departamento','provincia','region','depto','dpto'],
  producto:     ['producto','productos','product','products','articulo','item',
                 'descripcion','description','sku','referencia producto',
                 'nombre producto','item name'],
  cantidad:     ['cantidad','cant','unidades','qty','uds','n unidades'],
  valor:        ['valor','total','monto','precio','importe','valor total',
                 'total orden','total de la orden','venta','plata','dinero',
                 'valor a cobrar','recaudo','cobrar','amount'],
  costo_envio:  ['flete','envio','costo envio','domicilio','shipping','valor envio'],
  costo_producto:['costo','costo producto','cost','costo unitario'],
  estado:       ['estado','status','estatus','situacion','estado pedido',
                 'estado del pedido','en que va'],
  guia:         ['guia','numero guia','no guia','tracking','rastreo','guia transportadora'],
  transportadora:['transportadora','courier','operador','empresa envio','mensajeria'],
  gestora_asignada:['gestora','asesora','vendedora','responsable','encargada','atendido por',
                 'asignado a','quien atiende'],
  nota:         ['nota','notas','observacion','observaciones','comentario','comentarios',
                 'detalle','anotacion'],
  metodo_pago:  ['pago','metodo pago','forma de pago','medio de pago','payment'],
};

const FORMAS_PEDIDOS = {
  id_externo:   'texto',
  fecha:        'fecha',
  cliente:      'texto',
  telefono:     'texto',
  telefono_2:   'texto',
  cedula:       'texto',
  correo:       'texto',
  direccion:    'texto',
  ciudad:       'texto_repetido',
  departamento: 'texto_repetido',
  producto:     'texto_repetido',
  cantidad:     'entero',
  valor:        'numero',
  costo_envio:  'numero',
  costo_producto:'numero',
  estado:       'texto_repetido',
  guia:         'texto',
  transportadora:'texto_repetido',
  gestora_asignada:'texto_repetido',
  nota:         'texto',
  metodo_pago:  'texto_repetido',
};

/**
 * El mismo análisis, pero sobre filas en memoria en vez de una pestaña.
 *
 * Existe para que la app pueda proponer el mapeo de un archivo recién
 * subido sin escribir una sola fila todavía: primero se muestra qué
 * entendió, la persona corrige, y solo entonces se importa.
 */
function analizarFilas(datos, dicc, formas, campos) {
  dicc = dicc || SINONIMOS;
  formas = formas || FORMAS;
  if (!datos || datos.length < 2) {
    throw new Error('El archivo no tiene filas de datos.');
  }

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

  const objetivo = campos || Object.keys(dicc);
  const ranking = [];
  objetivo.forEach(function (campo) {
    enc.forEach(function (h, j) {
      if (!String(h).trim()) return;
      const pn = puntajeNombre(campo, h, dicc);
      const pf = puntajeForma(formas[campo] || 'texto', perfiles[j]);
      if (pn === 0 && pf === 0) return;
      ranking.push({ campo: campo, col: j, encabezado: String(h),
                     punt: pn + pf, pn: pn, pf: pf });
    });
  });
  ranking.sort(function (a, b) { return b.punt - a.punt; });

  const resueltos = {}, usadas = {}, propuestas = [];
  ranking.forEach(function (r) {
    if (resueltos[r.campo] || usadas[r.col]) return;
    if (r.pn === 0 && r.pf < 40) return;
    resueltos[r.campo] = true;
    usadas[r.col] = true;
    propuestas.push({
      campo: r.campo, columna: r.encabezado, indice: r.col, puntaje: r.punt,
      confianza: r.punt >= 110 ? 'alta' : (r.punt >= 70 ? 'media' : 'baja'),
    });
  });

  return {
    filaEncabezado: filaEnc + 1,
    encabezados: enc.map(function (h) { return String(h || '').trim(); }),
    filas: cuerpo.length,
    propuestas: propuestas,
    sinResolver: objetivo.filter(function (c) { return !resueltos[c]; }),
    // Tres filas de ejemplo para que la persona vea lo que va a cargar
    muestra: cuerpo.slice(0, 3).map(function (r) {
      return r.map(function (c) { return String(c == null ? '' : c).slice(0, 40); });
    }),
  };
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
   6 · ALARMAS
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Alarmas
 * ─────────────────────────────────────────────────────────────
 * Las seis cosas que no pueden esperar a que alguien abra la app.
 *
 * ┌─ POR QUÉ LOS UMBRALES LOS PONE EL CLIENTE ─────────────────┐
 * │                                                            │
 * │ Una tienda que sabe que vive con 40% de devolución no      │
 * │ necesita que le griten todos los días por eso. Otra que    │
 * │ vende un producto de 200 dólares se hunde con un 12%.      │
 * │                                                            │
 * │ Un umbral inventado por mí produce una de dos cosas: una   │
 * │ alarma que suena siempre —y que por eso se ignora— o una   │
 * │ que nunca suena. Las dos son igual de inútiles.            │
 * │                                                            │
 * │ Así que cada umbral vive en la hoja Parametros, por        │
 * │ tienda, y el dueño lo cambia desde la app. Los valores de  │
 * │ abajo son solo el punto de partida.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * El CPA es la excepción, y a propósito: su techo NO se escribe, se
 * calcula. Es lo que deja cada entrega después del producto y el flete,
 * porque pagar más que eso por conseguir un pedido es perder plata en
 * cada venta. Ese número cambia solo cuando cambian los costos.
 */

const ALARMAS_DEFAULT = {
  dias_sin_mover:     3,     // pedido sin movimiento
  horas_novedad:      24,    // novedad sin gestionar
  efectividad_min:    65,    // % de entrega sobre lo resuelto
  devoluciones_max:   '',    // % — lo pone el cliente; vacío = apagada
  cpa_aviso_pct:      85,    // % del techo a partir del cual avisa
  cpa_subida_pct:     25,    // % de subida contra el mes pasado que avisa
  stock_dias_min:     '',    // días de cobertura — vacío = apagada
  alarmas_a:          '',    // correos extra, separados por coma
  alarmas_hora:       7,     // hora local de la revisión diaria
};

/** Qué es cada alarma, en palabras de quien la va a leer. */
const ALARMAS = [
  { id: 'sin_mover',    nombre: 'Pedidos detenidos',
    param: 'dias_sin_mover', unidad: 'días' },
  { id: 'novedad_vieja', nombre: 'Novedades sin gestionar',
    param: 'horas_novedad', unidad: 'horas' },
  { id: 'efectividad',  nombre: 'Efectividad baja',
    param: 'efectividad_min', unidad: '%' },
  { id: 'devoluciones', nombre: 'Devoluciones altas',
    param: 'devoluciones_max', unidad: '%', opcional: true },
  { id: 'cpa',          nombre: 'CPA cerca del techo',
    param: 'cpa_aviso_pct', unidad: '% del techo' },
  { id: 'cpa_sube',     nombre: 'CPA subiendo',
    param: 'cpa_subida_pct', unidad: '% vs. mes pasado', opcional: true },
  { id: 'stock',        nombre: 'Stock por agotarse',
    param: 'stock_dias_min', unidad: 'días de cobertura', opcional: true },
  // Sin parámetro: no es cuestión de tolerancia, es que faltan datos.
  { id: 'estados_nuevos', nombre: 'Estados sin clasificar', param: '', unidad: '' },
];

/** Los umbrales de una tienda: lo que diga Parametros, o el de fábrica. */
/**
 * Ajustes de la tienda que no son alarmas.
 *
 * Viven en la misma hoja Parametros porque son la misma clase de cosa: un
 * valor que la dueña elige y Nova respeta. Pero se guardan en un catálogo
 * aparte para que la pantalla de Alarmas no tenga que esconder lo que no
 * le corresponde, y para que cada uno valide lo suyo.
 *
 * El canal de comunicación es un enlace al grupo que el equipo YA usa —
 * WhatsApp, Telegram, lo que sea. Nova no lo reemplaza: quien gestiona
 * tiene el celular abierto todo el día y no va a mudarse a otro chat
 * porque una app se lo pida. Lo que hace Nova es tenerlo a un clic desde
 * la pantalla donde está trabajando.
 */
const AJUSTES_DEFAULT = {
  canal_nombre: '',   // cómo lo llama el equipo: "Grupo de pedidos"
  canal_url:    '',   // a dónde lleva
  /**
   * Cómo se leen las fechas de los archivos que subes.
   *
   * 05-07-2026 son dos fechas distintas: 5 de julio si el día va primero,
   * 7 de mayo si va primero el mes. Ningún archivo dice cuál es, y
   * adivinar mal no da error — mueve medio mes de pedidos a otro mes y
   * nadie se entera hasta que un cierre no cuadra.
   *
   * Por defecto el día primero, que es como escribe toda América Latina.
   * Quien exporte desde una cuenta configurada en inglés lo cambia aquí
   * una vez y deja de pelear con esto.
   *
   * Solo afecta a lo ambiguo: un 25-12-2026 no tiene vuelta de hoja, y
   * lo que ya viene en AAAA-MM-DD tampoco.
   */
  formato_fecha: 'dia_primero',   // 'dia_primero' | 'mes_primero'

  /**
   * Las preguntas que la clienta hace siempre.
   *
   * La ficha de un producto no es un párrafo: es la respuesta a cuatro o
   * cinco cosas que se preguntan en toda llamada. Guardarlas juntas en un
   * bloque de texto obliga a quien confirma a leerlo entero por teléfono
   * buscando la frase que necesita, y por eso termina improvisando.
   *
   * Cuáles son depende del catálogo —no es lo mismo vender una crema que
   * un electrodoméstico—, así que las escribe cada tienda. Estas son el
   * punto de partida, y se cambian en Configuración.
   *
   * Separadas por |. Hasta cuatro.
   */
  preguntas_producto: '¿Para qué sirve?|¿Cómo se usa?|' +
                      '¿En cuánto tiempo se ven resultados?|' +
                      '¿Tiene contraindicaciones?',

  /**
   * Lo que cuesta mover la plata, que nadie factura pero se cobra igual.
   *
   * Dropi descuenta un porcentaje de cada retiro de la billetera a la
   * cuenta del banco. No aparece como gasto en ningún reporte: sale
   * restado del monto que llega, así que es invisible para cualquier
   * cierre que mire solo ventas y pauta. En agosto de Nutrea son 1.100
   * dólares retirados — a 3%, treinta y tres dólares que no estaban en
   * ninguna cuenta.
   *
   * La comisión internacional es otra cosa y por eso va aparte: el banco
   * la cobra por pagar en moneda extranjera, y no a todos los
   * proveedores. A Meta sí; a Shopify y a Claude, no. Aplicarla a todo
   * inflaría el costo de los que no la pagan.
   *
   * Los dos en cero apagan el cálculo. Nadie tiene que aceptar unas
   * comisiones que no son las suyas.
   */
  retiro_pct: 3,            // % que la plataforma descuenta de cada retiro
  comision_intl_pct: 2,     // % del banco por pagar en moneda extranjera
  comision_intl_a: 'meta',  // a qué plataformas se les aplica, separadas por coma

  /**
   * La cuenta publicitaria de Meta de esta tienda.
   *
   * Solo el número, sin el "act_" de adelante. Va aquí y no junto a la
   * llave porque no es un secreto: es como el número de una cuenta
   * bancaria, sirve para nombrarla y no para entrar en ella. La llave sí
   * es secreta, y vive en las Propiedades del Script — donde no la ve
   * quien abra la hoja ni quien se la descargue.
   */
  meta_cuenta: '',

  /**
   * Si Nova también baja los anuncios uno por uno.
   *
   * Cuesta llamadas: un conjunto con cuatro creativos son cuatro filas
   * por día en vez de una, y el nivel de acceso de una app nueva permite
   * 300 llamadas por hora. Por eso arranca encendido pero se puede
   * apagar: hay cuentas donde saber cuál creativo tira no cambia ninguna
   * decisión, y pagar por ese dato sin usarlo no tiene sentido.
   */
  meta_anuncios: 'si',

  /**
   * Los umbrales del semáforo semanal.
   *
   * Viven aquí, junto al resto de ajustes de la tienda, y no escondidos
   * dentro de un `if` del código que los usa: son decisiones de negocio.
   * Un umbral que nadie puede ver es un umbral que nadie discute, y el
   * día que el semáforo se equivoque hará falta poder moverlo sin tocar
   * el código.
   *
   * `ticket_minimo` arranca en cero, que apaga esa luz. Es a propósito:
   * el mínimo sano depende del catálogo y del país, y poner uno por
   * defecto sería inventarle a cada cliente una meta que no es suya.
   */
  ticket_minimo: 0,       // por debajo, el ticket es rojo. Moneda de la tienda.
  entrega_minima: 72,     // % de entrega sobre resuelto que se considera sano
  muestra_minima: 10,     // resueltos que hacen falta para que una cifra signifique algo
  cpa_verde_pct: 70,      // qué parte del techo se puede gastar antes del amarillo
};

/** Esquemas que sí se pueden abrir desde un enlace de la app. */
const ESQUEMAS_CANAL = ['http:', 'https:', 'whatsapp:', 'tg:', 'slack:', 'msteams:'];

/**
 * Un enlace que se pueda pegar en un href sin abrir una puerta.
 *
 * `javascript:` en el canal correría código en la sesión de cualquiera
 * del equipo que le diera clic, con su token al lado. Por eso no basta
 * con que el texto parezca un enlace: tiene que ser de un esquema que
 * lleve a otra aplicación, nunca a este mismo documento.
 */
function validarCanal(url) {
  const u = String(url || '').trim();
  if (!u) return '';                       // vacío apaga el canal, es válido
  const m = u.match(/^([a-z][a-z0-9+.-]*):/i);
  /**
   * Un enlace sin esquema se completa, no se rechaza.
   *
   * "nutrea.co/tag-recede" es un enlace perfectamente claro y es como se
   * copia de media parte. Devolver un error por eso hacía fallar el
   * guardado ENTERO de la ficha —nombre, stock, precios, todo— por una
   * cosa que se arregla poniendo cuatro letras delante.
   *
   * Rechazar sigue siendo lo correcto para lo que de verdad es peligroso
   * o no lleva a ninguna parte; para lo que solo está incompleto, se
   * completa.
   */
  if (!m) {
    if (/^[a-z0-9][a-z0-9.-]*\.[a-z]{2,}(\/|$|\?)/i.test(u)) return '';
    return 'Eso no parece un enlace. Cópialo de la barra del navegador, ' +
           'completo — algo como https://tutienda.com/producto.';
  }
  if (ESQUEMAS_CANAL.indexOf(m[1].toLowerCase() + ':') === -1) {
    return 'Ese tipo de enlace no se puede abrir desde Nova. Sirven los de ' +
           'WhatsApp, Telegram, Slack, Teams o cualquier página https.';
  }
  return '';
}

/**
 * El enlace tal como se va a guardar.
 *
 * Si venía sin esquema y pasó la validación, es un dominio: se le pone
 * https:// delante para que el href de la pantalla funcione. Guardarlo a
 * medias dejaría un botón que no lleva a ninguna parte.
 */
function normalizarEnlace(url) {
  const u = String(url || '').trim();
  if (!u) return '';
  return /^[a-z][a-z0-9+.-]*:/i.test(u) ? u : 'https://' + u;
}

function ajustes(ss, tienda) {
  const out = {};
  Object.keys(AJUSTES_DEFAULT).forEach(function (k) { out[k] = AJUSTES_DEFAULT[k]; });

  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda'), cK = e.indexOf('clave'), cV = e.indexOf('valor');
  for (let i = 1; i < d.length; i++) {
    const t = String(d[i][cT] || '').trim();
    if (t && t !== tienda) continue;
    const k = norm(d[i][cK]);
    if (!(k in out)) continue;
    out[k] = String(d[i][cV] == null ? '' : d[i][cV]).trim();
  }
  return out;
}

function umbrales(ss, tienda) {
  const out = {};
  Object.keys(ALARMAS_DEFAULT).forEach(function (k) { out[k] = ALARMAS_DEFAULT[k]; });

  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda'), cK = e.indexOf('clave'), cV = e.indexOf('valor');
  for (let i = 1; i < d.length; i++) {
    const t = String(d[i][cT] || '').trim();
    // Un parámetro sin tienda vale para todas: sirve de valor general
    if (t && t !== tienda) continue;
    const k = norm(d[i][cK]);
    if (!(k in out)) continue;
    const v = d[i][cV];
    out[k] = (v === '' || v === null) ? '' : v;
  }
  return out;
}

/**
 * Evalúa las seis alarmas de una tienda.
 *
 * Devuelve una lista, no manda correos. Separarlo permite que la pantalla
 * las muestre en vivo y que el correo diario use exactamente lo mismo:
 * si fueran dos cálculos distintos, tarde o temprano dirían cosas
 * distintas y no habría forma de saber cuál creer.
 */
function evaluarAlarmas(ss, tienda) {
  const u = umbrales(ss, tienda);
  const tz = zonaHorariaDe(ss, tienda);
  const hoy = new Date();
  const mes = Utilities.formatDate(hoy, tz || 'UTC', 'yyyy-MM');
  const out = [];
  const moneda = monedaDeTienda(ss, tienda);

  // ── Datos del mes, una sola lectura ──
  const sesionFalsa = { rol: 'dueno' };
  const m = agregarMes(ss, tienda, mes, sesionFalsa);

  // ── 1. Pedidos detenidos ──
  const dias = Number(u.dias_sin_mover) || 0;
  if (dias > 0) {
    const detenidos = [];
    const shP = ss.getSheetByName('Pedidos');
    if (shP && shP.getLastRow() > 1) {
      const d = shP.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        const est = norm(f[c('estado_nova')] || f[c('estado_canonico')]);
        if (['entregado','devolucion','cancelado'].indexOf(est) !== -1) continue;
        const ult = aISO(f[c('ultimo_movimiento')] || f[c('actualizado_en')] ||
                         f[c('fecha')], 'UTC');
        if (!ult) continue;
        const d2 = (hoy - new Date(ult + 'T00:00:00Z')) / 86400000;
        if (d2 >= dias) {
          detenidos.push({ id: f[c('id_externo')] || f[c('id')],
                           cliente: f[c('cliente')], dias: Math.floor(d2),
                           gestora: f[c('gestora_asignada')] });
        }
      }
    }
    if (detenidos.length) {
      detenidos.sort(function (a, b) { return b.dias - a.dias; });
      out.push(alarma('sin_mover', 'mal',
        pl(detenidos.length, '1 pedido lleva', '% pedidos llevan') + ' ' +
          dias + ' días o más sin moverse',
        'El más viejo lleva ' + detenidos[0].dias + ' días. Un pedido detenido ' +
        'no avisa solo: o se gestiona, o se convierte en devolución.',
        detenidos.slice(0, 10), 'Pedidos'));
    }
  }

  // ── 2. Novedades sin gestionar ──
  const horas = Number(u.horas_novedad) || 0;
  if (horas > 0) {
    const viejas = [];
    const shN = ss.getSheetByName('Novedades');
    if (shN && shN.getLastRow() > 1) {
      const d = shN.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (norm(f[c('estado')]) !== 'abierta') continue;
        const fch = aISO(f[c('fecha')], 'UTC');
        if (!fch) continue;
        const h = (hoy - new Date(fch + 'T00:00:00Z')) / 3600000;
        if (h >= horas) {
          viejas.push({ id: f[c('pedido_id')], motivo: f[c('motivo')],
                        horas: Math.floor(h), gestora: f[c('gestora')] });
        }
      }
    }
    if (viejas.length) {
      viejas.sort(function (a, b) { return b.horas - a.horas; });
      out.push(alarma('novedad_vieja', 'mal',
        pl(viejas.length, '1 novedad lleva', '% novedades llevan') +
          ' más de ' + horas + ' horas abierta' + (viejas.length === 1 ? '' : 's'),
        'Una novedad sin contactar a la clienta en el primer día se vuelve ' +
        'devolución en la mayoría de los casos.',
        viejas.slice(0, 10), 'Novedades'));
    }
  }

  /**
   * ── Estados que Nova no entiende ──
   *
   * Va antes que la efectividad a propósito: mientras haya pedidos sin
   * clasificar, la efectividad que se muestre debajo está calculada sobre
   * menos pedidos de los que hay. Avisar del número malo antes de
   * explicar por qué es malo sería enseñar a desconfiar de todo.
   *
   * No tiene umbral configurable: no es una cuestión de tolerancia, es
   * que faltan datos. Un solo pedido sin clasificar ya es una pregunta
   * sin responder, aunque no mueva la aguja.
   */
  if (m.sinClasificar > 0) {
    const cuales = Object.keys(m.estadosDesconocidos || {})
      .sort(function (a, b) {
        return m.estadosDesconocidos[b] - m.estadosDesconocidos[a];
      });
    out.push(alarma('estados_nuevos', 'mal',
      pl(m.sinClasificar, '1 pedido está', '% pedidos están') + ' en un estado ' +
        'que no reconozco',
      'No entran en ninguna cuenta: ni entregados, ni devueltos, ni en el ' +
      'flete. Son ' + moneda + ' ' + Math.round(m.valorSinClasificar) +
      ' sin clasificar. Los estados son: ' + cuales.slice(0, 5).join(' · ') +
      '.\n\nDinos qué significan en Configuración → Estados y las cifras del ' +
      'mes se rehacen solas.',
      cuales.slice(0, 10).map(function (k) {
        return { estado: k, pedidos: m.estadosDesconocidos[k] };
      }), 'Configuración'));
  }

  // ── 3. Efectividad baja ──
  const efMin = Number(u.efectividad_min) || 0;
  if (efMin > 0 && m.resueltos >= 10 && m.efectividad < efMin) {
    out.push(alarma('efectividad', 'mal',
      'Efectividad en ' + m.efectividad.toFixed(1) + '%, bajo tu meta de ' + efMin + '%',
      'De ' + m.resueltos + ' pedidos ya resueltos este mes llegaron ' +
      m.entregados + '. Cada punto por debajo son ventas que ya pagaste en pauta ' +
      'y no entraron.', [], 'Pedidos'));
  }

  // ── 4. Devoluciones altas (opcional) ──
  const devMax = u.devoluciones_max === '' ? null : Number(u.devoluciones_max);
  if (devMax !== null && devMax > 0 && m.resueltos >= 10) {
    const tasa = m.resueltos ? m.devueltos / m.resueltos * 100 : 0;
    if (tasa > devMax) {
      out.push(alarma('devoluciones', 'mal',
        'Devoluciones en ' + tasa.toFixed(1) + '%, sobre tu límite de ' + devMax + '%',
        m.devueltos + ' de ' + m.resueltos + ' pedidos resueltos volvieron. ' +
        'Cada uno cuesta el flete de ida y el de vuelta.', [], 'Novedades'));
    }
  }

  /**
   * ── 5. CPA cerca del techo ──
   *
   * El techo no se escribe en ninguna parte: es lo que deja cada entrega
   * después del producto y el flete. Pagar más que eso por traer un
   * pedido es perder plata en cada venta, por bien que se vea el ROAS.
   */
  /**
   * Diez entregas como mínimo, y aquí está el porqué.
   *
   * A principios de mes la pauta ya se gastó y los pedidos todavía no
   * llegan: el CPA se calcula sobre dos o tres entregas y sale disparado.
   * El 12 de septiembre esta alarma decía "pagas USD 86,41 por pedido"
   * cuando el mes cerrado anterior iba en 17,93. No era un problema de
   * la operación, era un mes que apenas empezaba.
   *
   * Una alarma que grita cada primero de mes es una alarma que se ignora
   * el resto del mes.
   */
  if (m.entregados >= 10 && m.gasto > 0) {
    const techo = (m.ventas - m.costoProducto - m.costoEnvio) / m.entregados;
    const pagado = m.gasto / m.entregados;

    /**
     * El techo dice si estás perdiendo. Esta otra dice si estás
     * empeorando, que es la que avisa a tiempo: un CPA que sube 30% en un
     * mes todavía puede estar bajo el techo, y aun así ser la señal de
     * que la campaña se está agotando o la competencia subió la puja.
     *
     * Se compara con el mes pasado completo, no con el promedio: un
     * promedio de varios meses suaviza justo lo que hay que ver.
     */
    const sub = u.cpa_subida_pct === '' ? null : Number(u.cpa_subida_pct);
    if (sub !== null && sub > 0) {
      const ant = agregarMes(ss, tienda, mesAnterior(mes), sesionFalsa);
      const cpaAnt = ant.entregados ? ant.gasto / ant.entregados : 0;
      // También el mes pasado necesita volumen: comparar contra un mes
      // de tres entregas produce porcentajes enormes que no dicen nada.
      if (cpaAnt > 0 && ant.entregados >= 10) {
        const delta = (pagado - cpaAnt) / cpaAnt * 100;
        if (delta >= sub) {
          out.push(alarma('cpa_sube', 'ojo',
            'El CPA subió ' + delta.toFixed(0) + '% contra el mes pasado',
            'Pagabas ' + moneda + ' ' + cpaAnt.toFixed(2) + ' por pedido y ahora ' +
            'pagas ' + moneda + ' ' + pagado.toFixed(2) + '. Todavía ' +
            (pagado < techo ? 'estás bajo el techo, pero la tendencia se come el colchón.'
                            : 'y además ya pasaste el techo.'),
            [], 'Dinero'));
        }
      }
    }
    const pct = techo > 0 ? pagado / techo * 100 : 999;
    const aviso = Number(u.cpa_aviso_pct) || 85;
    if (techo > 0 && pct >= aviso) {
      const grave = pct >= 100;
      out.push(alarma('cpa', grave ? 'mal' : 'ojo',
        grave
          ? 'Estás pagando más por pedido de lo que deja cada entrega'
          : 'CPA al ' + pct.toFixed(0) + '% del techo',
        'Cada entrega deja ' + moneda + ' ' + techo.toFixed(2) + ' después del ' +
        'producto y el flete, y estás pagando ' + moneda + ' ' + pagado.toFixed(2) +
        ' de pauta por conseguirla. ' +
        (grave ? 'Así, vender más es perder más.'
               : 'Queda poco colchón: si la entrega cae unos puntos, el mes se pone rojo.'),
        [], 'Dinero'));
    }
  }

  /**
   * ── 6. Stock por agotarse ──
   *
   * Dos formas de quedarse corto, y las dos avisan en la misma alarma.
   *
   * La primera es el mínimo que la dueña le puso a ese producto en su
   * ficha: un número que ella eligió sabiendo cuánto tarda su proveedor.
   * Esa no depende de ningún umbral global — si se molestó en escribirlo,
   * es porque quiere que le avisen.
   *
   * La segunda es la cobertura en días, que sí es opcional y global: al
   * ritmo de venta de este mes, cuántos días aguanta. Sirve para el
   * producto al que nadie le puso mínimo.
   *
   * Van juntas en un solo aviso porque la pregunta es una sola: qué hay
   * que reponer. Dos correos por lo mismo se vuelven ruido, y el ruido
   * termina en la carpeta de no leídos.
   */
  const stockDias = u.stock_dias_min === '' ? null : Number(u.stock_dias_min);
  const bajos = [];
  const shI = ss.getSheetByName('Inventario');
  if (shI && shI.getLastRow() > 1) {
    const d = shI.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    // Ritmo de venta de cada producto en el mes, para estimar cobertura
    const ritmo = {};
    Object.keys(m.productos || {}).forEach(function (k) {
      const dm = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
      ritmo[norm(k)] = (m.productos[k].entregados || 0) / dm;
    });
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      if (norm(f[c('activo')]) === 'no') continue;
      const nom = String(f[c('producto')] || '').trim();
      if (!nom) continue;
      const stock = num(f[c('stock')]);
      const minimo = c('minimo') !== -1 ? num(f[c('minimo')]) : 0;
      const r = ritmo[norm(nom)] || 0;
      const cobertura = r ? Math.floor(stock / r) : null;

      let motivo = '';
      if (minimo > 0 && stock <= minimo) {
        motivo = stock <= 0 ? 'sin stock' : 'bajo su mínimo de ' + minimo;
      } else if (stockDias !== null && stockDias > 0 && cobertura !== null &&
                 cobertura <= stockDias) {
        motivo = 'alcanza ' + cobertura + ' días';
      }
      if (!motivo) continue;
      bajos.push({ producto: nom, stock: stock, dias: cobertura, motivo: motivo });
    }
  }
  if (bajos.length) {
    // Primero lo que menos aguanta; lo que no se puede estimar, al final
    bajos.sort(function (a, b) {
      const da = a.dias === null ? 9999 : a.dias, db = b.dias === null ? 9999 : b.dias;
      return da - db;
    });
    out.push(alarma('stock', 'ojo',
      pl(bajos.length, '1 producto se está acabando', '% productos se están acabando'),
      'Quedarse sin stock con la pauta prendida es pagar por pedidos que no ' +
      'puedes despachar.',
      bajos.slice(0, 10), 'Inventario'));
  }

  return { alarmas: out, umbrales: u, tienda: tienda, mes: mes };
}

/** Uno o varios. "1 novedades" delata que lo escribió una máquina. */
function pl(n, uno, varios) {
  return n === 1 ? uno.replace('%', n) : varios.replace('%', n);
}

function alarma(id, nivel, titulo, detalle, casos, ir) {
  const def = ALARMAS.filter(function (a) { return a.id === id; })[0] || {};
  return { id: id, nivel: nivel, nombre: def.nombre || id,
           titulo: titulo, detalle: detalle, casos: casos || [], ir: ir || '' };
}

/**
 * La revisión diaria que manda el correo.
 *
 * Manda UNA vez por alarma y por día. Sin eso, una efectividad baja que
 * dura toda la semana produce siete correos idénticos, y al tercero ya
 * nadie los abre — que es exactamente cuando deja de servir.
 */
function revisarAlarmas(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const tiendas = tiendasDeCliente(ss);
  const log = [];

  tiendas.forEach(function (tienda) {
    const r = evaluarAlarmas(ss, tienda);
    if (!r.alarmas.length) { log.push(tienda + ': sin alarmas'); return; }

    const nuevas = r.alarmas.filter(function (a) {
      return !yaAvisada(ss, a.id + '|' + tienda);
    });
    if (!nuevas.length) {
      log.push(tienda + ': ' + r.alarmas.length + ' alarmas, ya avisadas hoy');
      return;
    }

    const destinos = destinatarios(ss, r.umbrales);
    if (destinos.length) {
      MailApp.sendEmail({
        to: destinos.join(','),
        subject: 'Nova · ' + nuevas.length + ' cosas que mirar en ' + nombreTienda(ss, tienda),
        body: cuerpoCorreo(nuevas, tienda, nombreTienda(ss, tienda)),
      });
    }
    nuevas.forEach(function (a) { marcarAvisada(ss, a.id + '|' + tienda); });
    log.push(tienda + ': avisadas ' + nuevas.length + ' a ' + destinos.join(', '));
  });

  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

/** A quién le llega: las dueñas activas, más los correos que se agreguen. */
function destinatarios(ss, u) {
  const out = [];
  const sh = ss.getSheetByName('Equipo');
  if (sh && sh.getLastRow() > 1) {
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const cC = e.indexOf('correo'), cR = e.indexOf('rol'), cE = e.indexOf('estado');
    for (let i = 1; i < d.length; i++) {
      if (norm(d[i][cE]) === 'inactivo') continue;
      if (rolCanonico(d[i][cR]) !== 'dueno') continue;
      const c = String(d[i][cC] || '').trim();
      if (c) out.push(c);
    }
  }
  String(u.alarmas_a || '').split(/[,;]/).forEach(function (c) {
    const x = c.trim();
    if (x && out.indexOf(x) === -1) out.push(x);
  });
  return out;
}

function cuerpoCorreo(alarmas, tienda, nombre) {
  const lineas = ['Hola,', '',
    'Esto es lo que Nova encontró hoy en ' + nombre + ':', ''];
  alarmas.forEach(function (a, i) {
    lineas.push((i + 1) + '. ' + a.titulo);
    lineas.push('   ' + a.detalle);
    if (a.casos && a.casos.length) {
      a.casos.slice(0, 5).forEach(function (c) {
        lineas.push('   · ' + (c.cliente || c.producto || c.id || '') +
          (c.dias !== undefined ? ' — ' + c.dias + ' días' : '') +
          (c.horas !== undefined ? ' — ' + c.horas + ' horas' : ''));
      });
      if (a.casos.length > 5) lineas.push('   · y ' + (a.casos.length - 5) + ' más');
    }
    lineas.push('');
  });
  lineas.push('Los umbrales de estas alarmas los cambias tú en Nova, ' +
              'en Configuración.');
  lineas.push('');
  lineas.push('— Nova');
  return lineas.join('\n');
}

/** Una alarma avisada hoy no se vuelve a avisar hoy. */
function yaAvisada(ss, clave) {
  const sh = ss.getSheetByName('Alertas_enviadas');
  if (!sh || sh.getLastRow() < 2) return false;
  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() !== clave) continue;
    if (String(aISO(d[i][2], 'UTC')) === hoy) return true;
  }
  return false;
}

function marcarAvisada(ss, clave) {
  const sh = ss.getSheetByName('Alertas_enviadas');
  if (!sh) return;
  sh.appendRow([clave, '', ahoraISO(), '']);
}

function nombreTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return tienda;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cN = e.indexOf('nombre');
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][cId]).trim() === tienda) return String(d[i][cN] || tienda);
  }
  return tienda;
}

/**
 * Deja la revisión diaria corriendo sola.
 *
 * Es lo que hace que Nova avise sin que nadie abra nada. Sin esto, las
 * alarmas solo existen para quien ya está mirando la pantalla — que es
 * justo quien menos las necesita.
 */
function instalarTriggerAlarmas() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'revisarAlarmasTodos') {
      ScriptApp.deleteTrigger(t);
    }
  });
  const hora = Number(ALARMAS_DEFAULT.alarmas_hora) || 7;
  ScriptApp.newTrigger('revisarAlarmasTodos').timeBased().atHour(hora).everyDays(1).create();
  const msg = 'Revisión diaria de alarmas instalada para las ' + hora + ':00.';
  Logger.log(msg);
  return msg;
}

/** Recorre todos los clientes registrados. */
function revisarAlarmasTodos() {
  const central = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!central || central.getLastRow() < 2) return 'Sin clientes.';
  const filas = central.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('sheet_id') !== -1 ? enc.indexOf('sheet_id') : 13;
  const log = [];
  for (let i = 1; i < filas.length; i++) {
    const id = String(filas[i][cId] || '').trim();
    if (!id) continue;
    try { log.push(revisarAlarmas(id)); }
    catch (e) { log.push('Cliente ' + id + ': ' + e.message); }
  }
  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}


/* ═══════════════════════════════════════════════════════════════
   7 · CLIENTES
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Provisionar clientes
 * ─────────────────────────────────────────────────────────────
 * Crea la hoja de un cliente nuevo copiando el template.
 *
 * Es la función que Nova Central llama cuando aprietas "Crear cuenta".
 * El cliente NUNCA entra a Apps Script ni ve su hoja: la app se la
 * administra. Este script vive una sola vez, en la cuenta de Nova.
 *
 * El template se queda SIEMPRE vacío de filas. La operación de Nova
 * (Nutrea) también es una copia, no el template — si no, cada cliente
 * nuevo nacería con las tiendas de Nutrea adentro.
 */

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
function crearCliente(empresa, pais, tiendas, fuentes, plan, dueno) {
  if (!empresa) throw new Error('Falta el nombre de la empresa.');
  if (!tiendas || !tiendas.length) throw new Error('Hay que declarar al menos una tienda.');

  /**
   * Sin dueña no hay cliente.
   *
   * Una hoja recién creada sin nadie en Equipo es una cuenta a la que
   * NADIE puede entrar: la app decide los roles leyendo esa hoja, así
   * que sin una fila ahí no hay quien dé permisos ni quien se los dé a
   * sí mismo. El cliente pagaría y no podría abrir su propia tienda.
   */
  if (!dueno || !dueno.correo || String(dueno.correo).indexOf('@') === -1) {
    throw new Error(
      'Falta la dueña. Se crea así:\n\n' +
      '  crearCliente("Empresa", "EC", tiendas, fuentes, "Base",\n' +
      '               { nombre: "Nombre", correo: "correo@dominio.com" })\n\n' +
      'Sin una dueña en la hoja Equipo, nadie puede entrar a esa cuenta.');
  }

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

  /**
   * Las filas se arman leyendo los encabezados, no contando columnas.
   *
   * Cuando Tiendas ganó la columna `modalidad`, esta función seguía
   * escribiendo diez valores en diez posiciones: 'activa' cayó en
   * modalidad y `estado` quedó vacío. Cada cliente nuevo nacía con sus
   * tiendas apagadas. Por posición, cualquier columna nueva corre todo
   * lo que viene detrás; por nombre, no.
   */
  sembrar(ss, 'Tiendas', tiendas.map(function (t) {
    return {
      id: t.id, nombre: t.nombre || t.id, marca: t.marca || empresa,
      pais: t.pais || pais || '', sociedad: t.sociedad || '', nit: t.nit || '',
      moneda: t.moneda || '', zona_horaria: t.zona || 'UTC',
      corte_despacho: t.corte || '16:00',
      modalidad: t.modalidad || 'catalogo_publico',
      estado: 'activa',
    };
  }), 2);

  // ── 2b. La dueña, sin la cual nadie puede entrar ──
  sembrar(ss, 'Equipo', [{
    id: 'eq-' + Utilities.getUuid().slice(0, 8),
    nombre: dueno.nombre || 'Dueña',
    correo: String(dueno.correo).toLowerCase().trim(),
    rol: 'dueno', tienda: '*', estado: 'activo', permisos: '',
  }], 2);

  // ── 3. Sembrar las fuentes declaradas ──
  if (fuentes && fuentes.length) {
    sembrar(ss, 'Fuentes', fuentes.map(function (f) {
      return { tienda: f.tienda, fuente: f.fuente, tipo: f.tipo || 'pedidos',
               cuenta: f.cuenta || '', activa: 'si' };
    }), 2);
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

/**
 * Escribe filas en una hoja emparejando por nombre de columna.
 *
 * @param {number} desde  fila donde empieza a escribir; sin esto habría
 *                        que saber si la hoja ya tiene datos.
 */
function sembrar(ss, hoja, objetos, desde) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('Falta la hoja ' + hoja + '. Corre bootstrapTodo().');
  if (!objetos || !objetos.length) return;

  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0]
    .map(function (h) { return String(h || '').trim().toLowerCase(); });

  const filas = objetos.map(function (o) {
    return enc.map(function (col) {
      return o[col] !== undefined ? o[col] : '';
    });
  });
  const inicio = desde || (sh.getLastRow() + 1);
  sh.getRange(inicio, 1, filas.length, enc.length).setValues(filas);
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
    'Interno',
    { nombre: 'Manuela', correo: 'nutreashop@gmail.com' }
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


/* ═══════════════════════════════════════════════════════════════
   8 · TASAS DE CAMBIO
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Tasas automáticas y efecto cambiario
 * ─────────────────────────────────────────────────────────────
 * El problema que resuelve, en concreto:
 *
 *   Operas Nutrea EC en dólares, pero vives en Colombia y tu plata es
 *   en pesos. Si el dólar pasa de 4.000 a 3.700, tu tienda puede
 *   facturar exactamente lo mismo en USD y aun así tu utilidad en COP
 *   cae 7,5%. Nada pasó en la operación.
 *
 *   El riesgo no es perder esa plata: es no saber por qué la perdiste,
 *   ver caer el número y ponerte a optimizar pauta cuando el problema
 *   no está en la pauta.
 *
 * Por eso hay tres piezas:
 *   1. actualizarTasas()   trae la tasa sola, todos los días
 *   2. alarmaTasa()        avisa cuando el movimiento ya pesa
 *   3. efectoCambiario()   separa cuánto es operación y cuánto es cambio
 */

// Celda de trabajo para GOOGLEFINANCE. Se usa y se limpia.
const TAB_FX = '_fx_tmp';

// ─── QUÉ PARES HACEN FALTA ───────────────────────────────────

/**
 * Los pares de moneda que este cliente necesita: de la moneda de cada
 * tienda a la moneda en que reporta. Si operas GT en quetzales y EC en
 * dólares, y reportas en pesos, los pares son GTQ->COP y USD->COP.
 */
function paresEnUso(ss) {
  const destino = monedaReporte(ss);
  const shT = ss.getSheetByName('Tiendas');
  if (!shT || shT.getLastRow() < 2) return [];

  const filas = shT.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cMon = enc.indexOf('moneda');
  const cEstado = enc.indexOf('estado');

  const vistos = {};
  filas.slice(1).forEach(function (f) {
    if (cEstado !== -1 && norm(f[cEstado]) === 'inactiva') return;
    const m = String(f[cMon] || '').toUpperCase();
    if (m && m !== destino) vistos[m] = 1;
  });

  return Object.keys(vistos).map(function (o) {
    return { origen: o, destino: destino };
  });
}

/**
 * Los pares que hacen falta para convertir el GASTO DE PAUTA.
 *
 * Son otros pares, y por eso existe esta función aparte. Meta no cobra
 * en la moneda de la tienda: le cobra en la del medio de pago. Una
 * tienda de Guatemala puede facturar en quetzales y que Meta le cobre en
 * dólares, y ese gasto se convierte a la moneda de la TIENDA —no a la
 * del reporte— porque es contra las ventas de esa tienda que se compara.
 *
 * Mientras esto no existió, paresEnUso decidía sola qué tasas traer.
 * A Nutrea le funcionaba de rebote: reporta en pesos, Meta le cobra en
 * pesos, y el par USD->COP servía invertido. A una tienda que facture y
 * reporte en la misma moneda le habría devuelto «no hacen falta tasas»,
 * y el gasto se habría quedado sin convertir para siempre sin que nada
 * lo dijera.
 */
function paresDeGasto(ss) {
  const sh = ss.getSheetByName('Pauta');
  if (!sh || sh.getLastRow() < 2) return [];

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cMon = enc.indexOf('moneda_gasto');
  const cTienda = enc.indexOf('tienda');
  if (cMon === -1 || cTienda === -1) return [];

  const monedaDe = {};   // se pregunta una vez por tienda, no una por fila
  const vistos = {};

  datos.slice(1).forEach(function (f) {
    const origen = String(f[cMon] || '').toUpperCase();
    const t = String(f[cTienda] || '').trim();
    if (!origen || !t) return;
    if (!(t in monedaDe)) {
      try { monedaDe[t] = String(monedaDeTienda(ss, t) || '').toUpperCase(); }
      catch (e) { monedaDe[t] = ''; }
    }
    const destino = monedaDe[t];
    if (!destino || destino === origen) return;
    vistos[origen + '|' + destino] = 1;
  });

  return Object.keys(vistos).map(function (k) {
    const p = k.split('|');
    return { origen: p[0], destino: p[1] };
  });
}

/**
 * Todo lo que hay que traer: lo del reporte y lo de la pauta, sin repetir.
 *
 * Un par y su inverso son el mismo par para buscarTasa —sabe dividir—,
 * así que traer los dos sería gastar dos consultas para el mismo dato.
 */
function paresNecesarios(ss) {
  const out = [];
  const vistos = {};
  paresEnUso(ss).concat(paresDeGasto(ss)).forEach(function (p) {
    if (!p.origen || !p.destino || p.origen === p.destino) return;
    const directo = p.origen + '|' + p.destino;
    const inverso = p.destino + '|' + p.origen;
    if (vistos[directo] || vistos[inverso]) return;
    vistos[directo] = 1;
    out.push(p);
  });
  return out;
}

/** La moneda en que el dueño ve su plata. Sale de Parametros. */
function monedaReporte(ss) {
  const sh = ss.getSheetByName('Parametros');
  if (sh && sh.getLastRow() > 1) {
    const filas = sh.getDataRange().getValues().slice(1);
    for (let i = 0; i < filas.length; i++) {
      if (norm(filas[i][1]) === 'moneda_reporte' && filas[i][2]) {
        return String(filas[i][2]).toUpperCase();
      }
    }
  }
  return 'COP';
}

// ─── TRAER LAS TASAS ─────────────────────────────────────────

/**
 * Rellena los huecos de la hoja `Tasas` con GOOGLEFINANCE.
 *
 * Se usa GOOGLEFINANCE y no una API externa por tres razones: no pide
 * llave, no tiene límite de llamadas, y no se cae un domingo dejando la
 * plataforma sin tasas. La contra es que solo vive dentro de Sheets, así
 * que hay que escribir la fórmula en una celda y leer el resultado.
 *
 * @param {string} cliente  nombre del cliente, o vacío si solo hay uno
 * @param {number} dias     cuántos días hacia atrás revisar (por defecto 90)
 */
function actualizarTasas(cliente, dias) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const pares = paresNecesarios(ss);
  if (!pares.length) {
    Logger.log('No hay pares que actualizar: las tiendas reportan en su misma ' +
               'moneda y la pauta se cobra en esa misma moneda.');
    return 'Sin pares que actualizar.';
  }

  const shTasas = ss.getSheetByName('Tasas');
  const hoy = new Date();
  const desde = new Date(hoy.getTime() - (dias || 90) * 86400000);

  // Qué fechas ya tenemos, para no repetir
  const existentes = {};
  if (shTasas.getLastRow() > 1) {
    shTasas.getDataRange().getValues().slice(1).forEach(function (f) {
      const k = aISO(f[0], 'UTC') + '|' + String(f[1]).toUpperCase() +
                '|' + String(f[2]).toUpperCase();
      existentes[k] = 1;
    });
  }

  let sh = ss.getSheetByName(TAB_FX);
  if (!sh) { sh = ss.insertSheet(TAB_FX); sh.hideSheet(); }

  const nuevas = [];
  const errores = [];

  pares.forEach(function (p) {
    const simbolo = 'CURRENCY:' + p.origen + p.destino;
    const formula = '=GOOGLEFINANCE("' + simbolo + '","price",DATE(' +
      desde.getFullYear() + ',' + (desde.getMonth() + 1) + ',' + desde.getDate() + '),DATE(' +
      hoy.getFullYear() + ',' + (hoy.getMonth() + 1) + ',' + hoy.getDate() + '),"DAILY")';

    sh.clear();
    sh.getRange(1, 1).setFormula(formula);
    SpreadsheetApp.flush();
    Utilities.sleep(1200); // GOOGLEFINANCE tarda en resolver

    const datos = sh.getDataRange().getValues();
    // La primera fila son encabezados (Date / Close); si vino un error,
    // la celda trae el texto del error en vez de la tabla.
    if (datos.length < 2) {
      errores.push(p.origen + '->' + p.destino + ': ' + (datos[0] ? datos[0][0] : 'sin datos'));
      return;
    }

    datos.slice(1).forEach(function (r) {
      const fecha = aISO(r[0], 'UTC');
      const tasa = aNumero(r[1]);
      if (!fecha || tasa === '' || tasa <= 0) return;
      const k = fecha + '|' + p.origen + '|' + p.destino;
      if (existentes[k]) return;
      existentes[k] = 1;
      nuevas.push([fecha, p.origen, p.destino, tasa]);
    });
  });

  sh.clear();

  if (nuevas.length) {
    nuevas.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    shTasas.getRange(shTasas.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);
  }

  const msg = [
    'Tasas agregadas: ' + nuevas.length,
    'Pares: ' + pares.map(function (p) { return p.origen + '->' + p.destino; }).join(', '),
    errores.length ? 'ERRORES:\n  ' + errores.join('\n  ') : '',
  ].filter(String).join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Deja la actualización corriendo sola, todos los días a las 6 a.m.
 * Córrelo una vez por cliente.
 */
function instalarTriggerTasas() {
  const ya = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'actualizarTasasDiario';
  });
  if (ya.length) { Logger.log('El trigger ya existía.'); return 'Ya existía.'; }

  ScriptApp.newTrigger('actualizarTasasDiario')
    .timeBased().atHour(6).everyDays(1).create();
  Logger.log('Trigger diario instalado (6 a.m.).');
  return 'Trigger diario instalado.';
}

/** Lo que corre el trigger: actualiza las tasas de todos los clientes. */
function actualizarTasasDiario() {
  const clientes = listarClientes();
  clientes.forEach(function (c) {
    if (!c.sheetId) return;
    try {
      actualizarTasas(c.sheetId, 10); // ventana corta: solo tapar huecos
    } catch (e) {
      Logger.log('Fallo actualizando ' + c.empresa + ': ' + e.message);
    }
  });
}

// ─── ALARMA DE MOVIMIENTO ────────────────────────────────────

/**
 * Avisa cuando la tasa se movió lo suficiente para que importe.
 *
 * Dos umbrales, porque son dos preocupaciones distintas:
 *   · corto plazo — se movió fuerte esta semana, revisa precios
 *   · desde el inicio — cuánto se corrió respecto a cuando arrancaste,
 *     que es el caso que planteaste del dólar a 4.000
 *
 * Umbrales configurables en Parametros:
 *   tasa_alerta_pct_30d   (por defecto 5)
 *   tasa_referencia_USD   (opcional: la tasa con la que montaste el negocio)
 */
function alarmaTasa(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const pares = paresEnUso(ss);
  const destino = monedaReporte(ss);
  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const hace30 = Utilities.formatDate(new Date(Date.now() - 30 * 86400000), 'UTC', 'yyyy-MM-dd');

  const umbral = Number(parametro(ss, 'tasa_alerta_pct_30d') || 5);
  const avisos = [];

  pares.forEach(function (p) {
    const ahora = buscarTasa(ss, hoy, p.origen, p.destino);
    const antes = buscarTasa(ss, hace30, p.origen, p.destino);
    if (!ahora) {
      avisos.push({
        nivel: 'Atención',
        titulo: 'Sin tasa ' + p.origen + '→' + p.destino,
        detalle: 'No hay tasa reciente. Todos los números de dinero de esa ' +
                 'tienda están sin convertir. Corre actualizarTasas().',
      });
      return;
    }
    if (antes) {
      const varPct = ((ahora.tasa - antes.tasa) / antes.tasa) * 100;
      if (Math.abs(varPct) >= umbral) {
        const sube = varPct > 0;
        avisos.push({
          nivel: Math.abs(varPct) >= umbral * 2 ? 'Crítica' : 'Atención',
          titulo: p.origen + '→' + p.destino + ' se movió ' + varPct.toFixed(1) + '% en 30 días',
          detalle: 'Pasó de ' + antes.tasa.toFixed(2) + ' a ' + ahora.tasa.toFixed(2) + '. ' +
            (sube
              ? 'Lo que factura esa tienda vale más en ' + destino + ' que hace un mes.'
              : 'Lo que factura esa tienda vale menos en ' + destino + '. ' +
                'Tu utilidad cae aunque la operación esté igual.'),
          accion: { texto: 'Ver efecto cambiario', destino: 'dinero' },
        });
      }
    }

    // Contra la tasa con la que se montó el negocio
    const ref = Number(parametro(ss, 'tasa_referencia_' + p.origen) || 0);
    if (ref > 0) {
      const varRef = ((ahora.tasa - ref) / ref) * 100;
      if (Math.abs(varRef) >= 10) {
        avisos.push({
          nivel: 'Informativa',
          titulo: p.origen + '→' + p.destino + ': ' + varRef.toFixed(0) +
                  '% contra tu tasa de referencia',
          detalle: 'Montaste el negocio con ' + ref.toFixed(0) + ' y hoy está en ' +
                   ahora.tasa.toFixed(0) + '. Si los precios no se movieron desde ' +
                   'entonces, el margen real cambió ' + varRef.toFixed(0) + '%.',
        });
      }
    }
  });

  Logger.log(avisos.length
    ? avisos.map(function (a) { return '[' + a.nivel + '] ' + a.titulo + '\n  ' + a.detalle; }).join('\n\n')
    : 'Sin movimientos de tasa relevantes.');
  return avisos;
}

/** Lee un parámetro. Busca primero por tienda, si no el global '*'. */
function parametro(ss, clave, tienda) {
  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues().slice(1);
  let global = null;
  for (let i = 0; i < filas.length; i++) {
    if (norm(filas[i][1]) !== norm(clave)) continue;
    if (tienda && String(filas[i][0]).trim() === tienda) return filas[i][2];
    if (String(filas[i][0]).trim() === '*') global = filas[i][2];
  }
  return global;
}

// ─── EFECTO CAMBIARIO ────────────────────────────────────────

/**
 * Separa cuánto de tu variación es operación y cuánto es tipo de cambio.
 *
 * Es la pieza que evita el error caro: ver caer la utilidad en pesos y
 * salir a optimizar pauta, cuando lo que se movió fue el dólar.
 *
 * La descomposición es exacta:
 *   total      = R1·T1 − R0·T0
 *   operación  = (R1 − R0) · T0     ... cuánto habrías cambiado a tasa fija
 *   cambio     = R1 · (T1 − T0)     ... cuánto movió la tasa sola
 *   y operación + cambio = total, siempre.
 *
 * @param {string} cliente   nombre del cliente
 * @param {string} tienda    id de la tienda. Ej: 'ec'
 * @param {string} mesA      mes base 'AAAA-MM'
 * @param {string} mesB      mes a comparar 'AAAA-MM'
 */
function efectoCambiario(cliente, tienda, mesA, mesB) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const destino = monedaReporte(ss);
  const origen = monedaDeTienda(ss, tienda);
  if (!origen) throw new Error('No encuentro la tienda "' + tienda + '" en la hoja Tiendas.');

  if (origen === destino) {
    const msg = 'La tienda ' + tienda + ' opera en ' + origen +
                ', la misma moneda en que reportas. No hay efecto cambiario.';
    Logger.log(msg);
    return { aplica: false, mensaje: msg };
  }

  const R0 = ventasDelMes(ss, tienda, mesA);
  const R1 = ventasDelMes(ss, tienda, mesB);
  const T0 = tasaPromedioMes(ss, mesA, origen, destino);
  const T1 = tasaPromedioMes(ss, mesB, origen, destino);

  if (T0 === null || T1 === null) {
    const msg = 'Faltan tasas para ' + (T0 === null ? mesA : mesB) +
                '. Corre actualizarTasas() primero.';
    Logger.log(msg);
    return { aplica: false, mensaje: msg };
  }

  const enDestino0 = R0 * T0;
  const enDestino1 = R1 * T1;
  const total = enDestino1 - enDestino0;
  const porOperacion = (R1 - R0) * T0;
  const porCambio = R1 * (T1 - T0);

  const pct = function (x) {
    return enDestino0 ? (x / enDestino0 * 100).toFixed(1) + '%' : 'n/d';
  };

  const msg = [
    'EFECTO CAMBIARIO · tienda ' + tienda + ' · ' + mesA + ' vs ' + mesB,
    '',
    'Ventas en ' + origen + ':',
    '  ' + mesA + ': ' + fmtMoneda(R0, origen),
    '  ' + mesB + ': ' + fmtMoneda(R1, origen) +
      (R0 ? '   (' + ((R1 - R0) / R0 * 100).toFixed(1) + '%)' : ''),
    '',
    'Tasa ' + origen + '→' + destino + ':',
    '  ' + mesA + ': ' + T0.toFixed(2),
    '  ' + mesB + ': ' + T1.toFixed(2) +
      (T0 ? '   (' + ((T1 - T0) / T0 * 100).toFixed(1) + '%)' : ''),
    '',
    'En ' + destino + ':',
    '  ' + mesA + ': ' + fmtMoneda(enDestino0, destino),
    '  ' + mesB + ': ' + fmtMoneda(enDestino1, destino),
    '',
    '  variación total ....... ' + fmtMoneda(total, destino) + '   ' + pct(total),
    '  por OPERACIÓN ......... ' + fmtMoneda(porOperacion, destino) + '   ' + pct(porOperacion),
    '  por TIPO DE CAMBIO .... ' + fmtMoneda(porCambio, destino) + '   ' + pct(porCambio),
    '',
    lecturaEfecto(porOperacion, porCambio, destino),
  ].join('\n');

  Logger.log(msg);
  return {
    aplica: true, origen: origen, destino: destino,
    ventas: { a: R0, b: R1 }, tasas: { a: T0, b: T1 },
    total: total, porOperacion: porOperacion, porCambio: porCambio,
    mensaje: msg,
  };
}

/** Traduce los dos números a una frase que dice qué hacer. */
function lecturaEfecto(op, fx, destino) {
  const aOp = Math.abs(op), aFx = Math.abs(fx);
  if (aOp === 0 && aFx === 0) return 'Sin variación.';

  if (aFx > aOp * 2) {
    return op >= 0 && fx < 0
      ? '→ La operación MEJORÓ, pero la tasa se comió la mejora. No es un problema ' +
        'de ventas ni de pauta: es cambiario. Revisa precios, no campañas.'
      : '→ El movimiento es principalmente cambiario. La operación explica poco.';
  }
  if (aOp > aFx * 2) {
    return '→ El movimiento es de operación. La tasa casi no influyó: ' +
           'lo que cambió está en las ventas.';
  }
  return '→ Operación y tasa pesan parecido. Mira las dos antes de decidir.';
}

/** Moneda de una tienda, según la hoja Tiendas. */
function monedaDeTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cMon = enc.indexOf('moneda');
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][cId]).trim() === tienda) {
      return String(filas[i][cMon] || '').toUpperCase();
    }
  }
  return null;
}

/** Ventas entregadas de una tienda en un mes 'AAAA-MM'. */
function ventasDelMes(ss, tienda, mes) {
  const sh = ss.getSheetByName('Pedidos');
  if (!sh || sh.getLastRow() < 2) return 0;
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cF = enc.indexOf('fecha'), cT = enc.indexOf('tienda');
  const cV = enc.indexOf('valor'), cE = enc.indexOf('estado_canonico');

  let suma = 0;
  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    if (String(f[cT]).trim() !== tienda) continue;
    if (String(aISO(f[cF], 'UTC')).slice(0, 7) !== mes) continue;
    // Solo lo entregado cuenta como venta: lo devuelto y lo cancelado no.
    if (cE !== -1 && norm(f[cE]) !== 'entregado') continue;
    const v = aNumero(f[cV]);
    if (v !== '') suma += v;
  }
  return suma;
}

/** Tasa promedio de un mes. Promedio y no cierre: las ventas se reparten. */
function tasaPromedioMes(ss, mes, origen, destino) {
  const sh = ss.getSheetByName('Tasas');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues().slice(1);
  let suma = 0, n = 0;
  filas.forEach(function (f) {
    const fecha = aISO(f[0], 'UTC');
    if (!fecha || fecha.slice(0, 7) !== mes) return;
    const o = String(f[1]).toUpperCase(), d = String(f[2]).toUpperCase();
    const t = aNumero(f[3]);
    if (t === '' || t <= 0) return;
    if (o === origen && d === destino) { suma += t; n++; }
    else if (o === destino && d === origen) { suma += 1 / t; n++; }
  });
  return n ? suma / n : null;
}


/* ═══════════════════════════════════════════════════════════════
   9 · API WEB
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · API web
 * ─────────────────────────────────────────────────────────────
 * El puente entre las hojas y las cinco pantallas.
 *
 * ┌─ POR QUÉ LA AUTENTICACIÓN ES ASÍ ──────────────────────────┐
 * │                                                            │
 * │ Una página HTML estática NO puede guardar un secreto: todo │
 * │ lo que esté en el JavaScript lo ve cualquiera que abra la  │
 * │ página. Un token fijo escrito en el HTML no es seguridad,  │
 * │ es una llave pegada en la puerta.                          │
 * │                                                            │
 * │ Por eso:                                                   │
 * │   · el correo se valida contra la hoja Equipo              │
 * │   · el código de 6 dígitos llega por correo de verdad      │
 * │   · el token se emite al verificar y vence en 12 horas     │
 * │   · el ROL y las TIENDAS los decide el servidor            │
 * │                                                            │
 * │ Esto último es lo importante: si el rol lo decidiera la    │
 * │ pantalla, bastaría abrir la consola del navegador y        │
 * │ cambiarlo para ver el dinero. El servidor no le cree nada  │
 * │ al cliente.                                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * DESPLIEGUE:
 *   Implementar → Nueva implementación → Aplicación web
 *   Ejecutar como:    Yo
 *   Quién tiene acceso: Cualquier usuario
 *   (Va a "cualquiera" porque el token es el que controla el acceso,
 *    no la capa de Google. Sin token válido no devuelve ni una fila.)
 */

const TTL_CODIGO_M = 10;   // minutos que dura el código de 6 dígitos

/**
 * Cuánto dura una sesión, y por qué son dos límites y no uno.
 *
 * ┌────────────────────────────────────────────────────────────┐
 * │ TOPE: desde que entras, pase lo que pase. Una jornada de   │
 * │ gestora es un turno; el de la dueña, un día de trabajo.    │
 * │ Que una sesión no dure para siempre es lo que impide que   │
 * │ un token robado sirva un mes después.                      │
 * │                                                            │
 * │ INACTIVIDAD: desde la última vez que hiciste algo. Este    │
 * │ es el que de verdad protege, porque el riesgo en una       │
 * │ tienda no es el tiempo: es el computador compartido que    │
 * │ alguien dejó abierto y se fue.                             │
 * └────────────────────────────────────────────────────────────┘
 *
 * Recargar la página NO cuenta como volver a entrar: el token sigue
 * vivo y la sesión se reanuda sola. Pedir código en cada recarga
 * enseñaría a no recargar nunca, que es peor para todos.
 */
const TTL_POR_ROL = { dueno: 12, admin: 6, gestora: 4 };   // horas de tope
const INACTIVIDAD_MIN = 90;                                // minutos sin hacer nada

// ─── ENTRADA ─────────────────────────────────────────────────

function doGet(e)  { return manejar(e, 'GET'); }
function doPost(e) { return manejar(e, 'POST'); }

function manejar(e, metodo) {
  try {
    const p = leerParams(e, metodo);
    const accion = String(p.accion || '').trim();
    if (!accion) {
      /**
       * "Falta la acción" a secas no deja avanzar: dice que el cuerpo
       * llegó sin acción, pero no si llegó vacío, si llegó con otra cosa
       * o si ni siquiera era un POST. Sin eso hay que adivinar, y adivinar
       * sobre una petición que no se ve cuesta horas.
       */
      const crudo = (e && e.postData && e.postData.contents) || '';
      return json({ ok: false, error: 'Falta la acción.',
        diagnostico: {
          metodo: metodo,
          hubo_cuerpo: !!crudo,
          largo: crudo.length,
          empieza: String(crudo).slice(0, 120),
          tipo: (e && e.postData && e.postData.type) || '',
          parametros: Object.keys((e && e.parameter) || {}),
        } });
    }

    /**
     * La consola de la plataforma va por su lado.
     *
     * Antes de buscar una sesión de cliente: estas acciones no tienen
     * ninguna que buscar, y hacerlo primero le respondería "sesión
     * inválida" a quien está tratando de entrar a Nova Central.
     */
    if (accion.indexOf('nc_') === 0) return json(manejarCentral(accion, p));

    // Las únicas dos que no piden token
    if (accion === 'login')     return json(apiLogin(p));
    if (accion === 'verificar') return json(apiVerificar(p));

    const real = sesion(p.token);
    if (!real) return json({ ok: false, error: 'Sesión vencida o inválida.', reautenticar: true });

    /**
     * "Ver como" baja de nivel, nunca sube.
     *
     * La dueña ya puede verlo todo, así que ponerse la vista de gestora
     * no le da nada nuevo: le quita. Es un filtro sobre lo que ya tiene
     * derecho a ver, no una llave nueva. Por eso dueña → admin → gestora
     * se permite y al revés jamás.
     *
     * Lo que NO cambia es quién es: el correo y el nombre siguen siendo
     * los suyos, así que la auditoría dice "Manuela", no "Katherin". Si
     * cambiara, la bitácora dejaría de servir justo para lo que existe.
     */
    const s = vistaEfectiva(real, p.vista);

    switch (accion) {
      case 'yo':        return json({ ok: true, sesion: publico(s) });
      case 'resumen':   return json(apiResumen(s, p));
      case 'listar':    return json(apiListar(s, p));
      case 'escribir':  return json(apiEscribir(s, p));
      case 'cierre':    return json(apiCierre(s, p));
      case 'importar':  return json(apiImportarArchivo(s, p));
      case 'fuentes':   return json(apiFuentes(s, p));
      case 'trozo':     return json(apiTrozo(s, p));
      case 'crear':     return json(apiCrear(s, p));
      case 'equipo':    return json(apiEquipo(s, p));
      case 'productos': return json(apiProductos(s, p));
      case 'recuento':  return json(apiRecuento(s, p));
      case 'historial': return json(apiHistorial(s, p));
      case 'cas':       return json(apiCas(s, p));
      case 'cas_escribir': return json(apiCasEscribir(s, p));
      case 'alarmas':   return json(apiAlarmas(s, p));
      case 'parametros':return json(apiParametros(s, p));
      case 'auditoria': return json(apiAuditoria(s, p));
      case 'estados':   return json(apiEstados(s, p));
      case 'meta_estado':  return json(apiMetaEstado(s, p));
      case 'meta_guardar': return json(apiMetaGuardar(s, p));
      case 'meta_probar':  return json(apiMetaProbar(s, p));
      case 'meta_traer':   return json(apiMetaTraer(s, p));
      case 'semaforo':     return json(apiSemaforo(s, p));
      case 'estado_clasificar': return json(apiEstadoClasificar(s, p));
      case 'borrar':    return json(apiBorrar(s, p));
      case 'cerrarmes': return json(apiCerrarMes(s, p));
      case 'salir':     return json(apiSalir(p.token));
      default:          return json({ ok: false, error: 'Acción desconocida: ' + accion });
    }
  } catch (err) {
    return json({ ok: false, error: String(err && err.message || err) });
  }
}

/**
 * El cliente manda POST con Content-Type text/plain a propósito.
 * Con application/json el navegador dispara una petición preflight
 * OPTIONS, que Apps Script no sabe responder y la llamada falla.
 */
function leerParams(e, metodo) {
  if (metodo === 'POST' && e && e.postData && e.postData.contents) {
    try { return JSON.parse(e.postData.contents); } catch (x) { /* cae abajo */ }
  }
  return (e && e.parameter) || {};
}

function json(obj) {
  return ContentService.createTextOutput(JSON.stringify(obj))
    .setMimeType(ContentService.MimeType.JSON);
}

// ─── SESIONES ────────────────────────────────────────────────

/**
 * Paso 1: valida el correo contra Equipo y manda el código.
 *
 * A un correo que no está en Equipo se le responde exactamente igual
 * que a uno que sí: si la respuesta cambiara, cualquiera podría probar
 * correos hasta descubrir quién trabaja en la empresa.
 */
function apiLogin(p) {
  const email = String(p.email || '').toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, error: 'Correo inválido.' };
  }

  const persona = buscarPersona(email);
  if (persona) {
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put(
      'cod_' + email, codigo, TTL_CODIGO_M * 60
    );
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Tu código de Nova: ' + codigo,
        body: 'Hola ' + (persona.nombre || '') + ',\n\n' +
              'Tu código para entrar a Nova es: ' + codigo + '\n\n' +
              'Vence en ' + TTL_CODIGO_M + ' minutos.\n' +
              'Si no fuiste tú, ignora este correo.\n',
      });
    } catch (err) {
      // Sin cuota de correo el login quedaría muerto y sin explicación.
      return { ok: false, error: 'No se pudo enviar el código: ' + err.message };
    }
  }

  return { ok: true, enviado: true, vence_en_min: TTL_CODIGO_M };
}

/** Paso 2: cambia el código por un token. */
function apiVerificar(p) {
  const email = String(p.email || '').toLowerCase().trim();
  const codigo = String(p.codigo || '').trim();
  const cache = CacheService.getScriptCache();
  const esperado = cache.get('cod_' + email);

  if (!esperado || esperado !== codigo) {
    return { ok: false, error: 'Código incorrecto o vencido.' };
  }
  cache.remove('cod_' + email);

  const persona = buscarPersona(email);
  if (!persona) return { ok: false, error: 'Esta cuenta ya no tiene acceso.' };

  const horas = TTL_POR_ROL[persona.rol] || 4;
  const token = Utilities.getUuid();
  const s = {
    email: email,
    nombre: persona.nombre,
    rol: persona.rol,
    clienteId: persona.clienteId,
    sheetId: persona.sheetId,
    tiendas: persona.tiendas,
    modulos: persona.modulos,
    permisos: persona.permisos,
    vence: Date.now() + horas * 3600000,
    ultimo: Date.now(),
    horas: horas,
  };
  cache.put('ses_' + token, JSON.stringify(s), horas * 3600);

  registrarMovimiento(s, 'Equipo', persona.id, 'ultima_conexion', '', ahoraISO());
  return { ok: true, token: token, sesion: publico(s) };
}

function sesion(token) {
  if (!token) return null;
  const cache = CacheService.getScriptCache();
  const raw = cache.get('ses_' + String(token));
  if (!raw) return null;
  const s = JSON.parse(raw);
  const ahora = Date.now();

  // Tope: se cumplió el turno
  if (s.vence <= ahora) return null;

  // Inactividad: se fue y dejó la pantalla abierta
  const quieto = (ahora - (s.ultimo || s.vence)) / 60000;
  if (s.ultimo && quieto > INACTIVIDAD_MIN) {
    cache.remove('ses_' + String(token));
    return null;
  }

  /**
   * Cada acción corre el reloj de inactividad, pero NUNCA el del tope.
   * Si el tope se moviera también, una sesión activa sería eterna y la
   * regla de las horas no serviría de nada.
   */
  s.ultimo = ahora;
  const quedan = Math.ceil((s.vence - ahora) / 1000);
  if (quedan > 0) cache.put('ses_' + String(token), JSON.stringify(s), quedan);

  /**
   * Las sesiones abiertas sobreviven a un despliegue: viven en el caché,
   * no en el código. Una sesión creada antes de que existieran los
   * permisos no trae el campo, y sin este respaldo la dueña se quedaba
   * sin poder importar hasta volver a entrar — justo después de
   * actualizar, que es cuando uno va a probar.
   */
  if (!s.permisos) s.permisos = permisosDe(s.rol, '');
  return s;
}

/** Jerarquía de roles. Solo se puede mirar hacia abajo. */
const NIVEL_ROL = { gestora: 1, admin: 2, dueno: 3 };

function vistaEfectiva(s, vista) {
  const v = String(vista || '').trim();
  if (!v || v === s.rol) return s;
  if (!NIVEL_ROL[v]) return s;
  if (!puedeVerComo(s)) return s;
  // Hacia arriba, nunca
  if (NIVEL_ROL[v] >= NIVEL_ROL[s.rol]) return s;

  const copia = JSON.parse(JSON.stringify(s));
  copia.rol = v;
  copia.rolReal = s.rol;
  copia.vistaComo = v;
  // Los permisos de dinero se van con el rol: una vista de gestora que
  // siguiera viendo la pauta no simularía nada.
  if (v === 'gestora') copia.permisos = ['subir_pedidos', 'subir_novedades'];
  return copia;
}

/** Mirar con menos permisos es un permiso más, y se puede quitar. */
function puedeVerComo(s) {
  if (s.rol === 'gestora') return false;
  const p = s.permisos || [];
  if (p.indexOf('-ver_como') !== -1) return false;
  return true;
}

function apiSalir(token) {
  CacheService.getScriptCache().remove('ses_' + String(token));
  return { ok: true };
}

/** Lo que el cliente puede saber de su propia sesión. Sin sheetId. */
/**
 * Qué puede hacer cada rol si la dueña no dice otra cosa.
 *
 * La gestora sube pedidos y novedades porque es quien gestiona: bajar el
 * archivo de Dropi y subirlo es parte de su trabajo diario, no una tarea
 * administrativa. Lo que no trae por defecto es la pauta, porque ahí está
 * el gasto y el margen.
 *
 * La admin tampoco trae pauta de entrada. La dueña se la da escribiendo
 * "subir_pauta" en la columna permisos de esa persona, en la hoja Equipo.
 *
 * La dueña no aparece aquí: puede todo, siempre. Un dueño que se queda sin
 * permisos deja la cuenta huérfana y sin quién los devuelva.
 */
const PERMISOS_POR_ROL = {
  admin:   ['subir_pedidos', 'subir_novedades'],
  gestora: ['subir_pedidos', 'subir_novedades'],
};

/** Qué permiso exige cada plataforma. */
const PERMISO_DE_FUENTE = {
  dropi: 'subir_pedidos', mastershop: 'subir_pedidos',
  effi_guias: 'subir_pedidos', shopify: 'subir_pedidos',
  effi_novedades: 'subir_novedades', iris: 'subir_novedades',
  meta: 'subir_pauta', meta_facturacion: 'subir_pauta', tiktok: 'subir_pauta',
  // La cartera es el extracto de la billetera: dice cuánto hay, cuánto se
  // retiró y con qué concepto. Es dinero, así que va con su propio permiso
  // y de entrada solo lo tiene la dueña, igual que la pauta.
  dropi_cartera: 'subir_cartera',
};

/**
 * `ver_dinero` es el único que no habilita a subir nada: habilita a MIRAR.
 *
 * Hacía falta porque el semáforo semanal muestra utilidad, gasto y techo
 * de CPA, y hasta ahora la única forma de que una admin los viera era
 * darle permiso para SUBIR la pauta — que es otra cosa. Poder leer lo que
 * la empresa gana y poder escribir lo que gastó no son el mismo riesgo.
 */
const PERMISOS_CONOCIDOS = ['subir_pedidos', 'subir_novedades', 'subir_pauta',
                            'subir_cartera', 'ver_dinero'];

/**
 * Los permisos de una persona: los de su rol, más lo que la dueña le haya
 * escrito en la columna permisos.
 *
 * La celda suma, no reemplaza. Si reemplazara, escribir "subir_pauta" para
 * darle Meta a una admin le quitaría en silencio los pedidos que ya subía.
 * Para quitar algo se antepone un menos: "-subir_novedades".
 */
function permisosDe(rol, celda) {
  if (rol === 'dueno') return PERMISOS_CONOCIDOS.slice();

  const out = (PERMISOS_POR_ROL[rol] || []).slice();
  String(celda || '').split(/[,;]/).forEach(function (t) {
    const raw = norm(t).replace(/\s+/g, '_');
    if (!raw) return;
    const quita = raw.charAt(0) === '-';
    const p = quita ? raw.slice(1) : raw;
    if (PERMISOS_CONOCIDOS.indexOf(p) === -1) return; // texto suelto: se ignora
    const i = out.indexOf(p);
    if (quita) { if (i !== -1) out.splice(i, 1); }
    else if (i === -1) out.push(p);
  });
  return out;
}

function publico(s, ss) {
  const o = { email: s.email, nombre: s.nombre, rol: s.rol,
              tiendas: s.tiendas, modulos: s.modulos || ['empresarial'],
              permisos: s.permisos || [],
              // Para que la pantalla avise antes de cortar, en vez de
              // cortar en mitad de algo sin explicación
              vence: s.vence, horas: s.horas || TTL_POR_ROL[s.rol] || 4,
              inactividad_min: INACTIVIDAD_MIN,
              puede_ver_como: puedeVerComo(s),
              rol_real: s.rolReal || s.rol };
  // La ficha va con moneda y país para que la pantalla no tenga que adivinarlos
  try {
    o.fichas = fichasDe(ss || SpreadsheetApp.openById(s.sheetId), s.tiendas);
  } catch (e) { o.fichas = []; }
  return o;
}

/**
 * Busca a la persona en el Equipo de cada cliente registrado.
 * De aquí salen el rol y las tiendas — nunca de lo que mande la pantalla.
 */
function buscarPersona(email) {
  const clientes = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!clientes || clientes.getLastRow() < 2) return null;

  const filas = clientes.getDataRange().getValues().slice(1);
  for (let i = 0; i < filas.length; i++) {
    const sheetId = filas[i][13];
    if (!sheetId || norm(filas[i][6]) === 'suspendido') continue;

    let ss;
    try { ss = SpreadsheetApp.openById(sheetId); } catch (x) { continue; }
    const sh = ss.getSheetByName('Equipo');
    if (!sh || sh.getLastRow() < 2) continue;

    const datos = sh.getDataRange().getValues();
    const enc = datos[0].map(norm);
    const c = function (n) { return enc.indexOf(n); };

    for (let j = 1; j < datos.length; j++) {
      const f = datos[j];
      if (String(f[c('correo')] || '').toLowerCase().trim() !== email) continue;
      if (norm(f[c('estado')]) === 'inactivo') return null;

      const rol = rolCanonico(f[c('rol')]);
      if (!rol) {
        throw new Error(
          'El rol "' + f[c('rol')] + '" de ' + f[c('nombre')] + ' no se reconoce.\n\n' +
          'En la hoja Equipo, la columna rol tiene que decir una de estas: ' +
          'dueno · admin · gestora\n\n' +
          'Se aceptan variantes (dueña, administradora, asesora...), pero no ' +
          'cualquier texto: un rol que no se entiende dejaría a la persona con ' +
          'menos permisos de los que le tocan, sin avisar.');
      }
      const tiendaCol = String(f[c('tienda')] || '').trim();
      return {
        modulos: modulosDelPlan(filas[i][3]),
        id: f[c('id')],
        nombre: f[c('nombre')],
        rol: rol,
        // c('permisos') es -1 en hojas creadas antes de que la columna
        // existiera: sin celda, quedan los permisos del rol.
        permisos: permisosDe(rol, c('permisos') === -1 ? '' : f[c('permisos')]),
        clienteId: filas[i][0],
        sheetId: sheetId,
        // '*' o vacío significa todas las tiendas del cliente
        tiendas: (tiendaCol && tiendaCol !== '*')
          ? tiendaCol.split(/[,;]/).map(function (x) { return x.trim(); })
          : tiendasDe(ss),
      };
    }
  }
  return null;
}

function tiendasDe(ss) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return [];
  const datos = sh.getDataRange().getValues();
  const e = datos[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  return datos.slice(1)
    .filter(function (f) { return f[c('id')] && norm(f[c('estado')]) !== 'inactiva'; })
    .map(function (f) { return String(f[c('id')]).trim(); });
}

/**
 * La ficha completa de cada tienda, no solo el id.
 *
 * La pantalla necesita la moneda para el símbolo y el país para el
 * formato de número: Guatemala escribe 1,234.56 y Colombia 1.234,56.
 * Mandar solo el id obligaría al HTML a adivinarlo, que es justo lo que
 * ataba la pantalla a dos países.
 */
function fichasDe(ss, ids) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return [];
  const datos = sh.getDataRange().getValues();
  const e = datos[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  return datos.slice(1)
    .filter(function (f) {
      const id = String(f[c('id')]).trim();
      return id && ids.indexOf(id) !== -1;
    })
    .map(function (f) {
      return {
        id:     String(f[c('id')]).trim(),
        nombre: f[c('nombre')] || '',
        moneda: String(f[c('moneda')] || '').toUpperCase(),
        pais:   f[c('pais')] || '',
      };
    });
}


/**
 * Traduce lo que esté escrito en la columna `rol` al rol canónico.
 *
 * Sin esto, escribir "Administradora" en vez de "admin" hacía que la
 * persona cayera en gestora por defecto: veía menos de lo que le
 * corresponde y nadie se enteraba. Un permiso mal asignado en silencio
 * es peor que un error visible, así que lo que no se reconoce falla.
 */
const ROLES_VALIDOS = {
  dueno:   ['dueno','duena','owner','propietario','propietaria','ceo','jefa','jefe'],
  admin:   ['admin','administrador','administradora','administrator','gerente','supervisor','supervisora'],
  gestora: ['gestora','gestor','agente','asesor','asesora','vendedor','vendedora'],
};

function rolCanonico(raw) {
  const k = norm(raw);
  if (!k) return null;
  const roles = Object.keys(ROLES_VALIDOS);
  for (let i = 0; i < roles.length; i++) {
    if (ROLES_VALIDOS[roles[i]].indexOf(k) !== -1) return roles[i];
  }
  return null;
}


/**
 * Qué productos de Nova ve esta cuenta.
 *
 * Un cliente que compró Nova Empresarial NO debe ver NovaSoul ni
 * novAcademy en su pantalla de inicio: no los compró, y NovaSoul además
 * guarda datos personales de quien la usa. Mostrar el enlace, aunque no
 * pueda entrar, ya es filtrar de más.
 *
 * El plan vive en Nova_Central → Clientes, y su composición en Planes.
 * Sin plan reconocido se cae al mínimo: solo Empresarial.
 */
function modulosDelPlan(plan) {
  const p = norm(plan);
  if (!p) return ['empresarial'];
  if (p === 'interno') return ['empresarial', 'soul', 'academy', 'central'];

  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Planes');
    if (sh && sh.getLastRow() > 1) {
      const d = sh.getDataRange().getValues();
      const e = d[0].map(norm);
      const cN = e.indexOf('nombre'), cM = e.indexOf('modulos');
      for (let i = 1; i < d.length; i++) {
        if (norm(d[i][cN]) !== p) continue;
        const m = String(d[i][cM] || '').split(/[,;]/)
          .map(function (x) { return norm(x); }).filter(String);
        return m.length ? m : ['empresarial'];
      }
    }
  } catch (err) { /* sin hoja Planes se usa el mínimo */ }
  return ['empresarial'];
}

// ─── PERMISOS ────────────────────────────────────────────────

/**
 * Qué entidades puede ver cada rol.
 *
 * Se evalúa en el SERVIDOR. La pantalla también esconde lo que no
 * corresponde, pero eso es comodidad visual, no seguridad: esconder un
 * div no impide que alguien llame la API directamente.
 */
const PERMISOS = {
  dueno:   { leer: '*', escribir: ['Pedidos','Novedades','Inventario','Equipo',
                                   'Parametros','Tiendas','Fuentes','Gastos'] },
  admin:   { leer: ['Pedidos','Novedades','Llamadas','Inventario','Equipo','Tiendas','Fuentes'],
             escribir: ['Pedidos','Novedades','Equipo'] },
  gestora: { leer: ['Pedidos','Novedades','Llamadas'], escribir: ['Pedidos','Novedades'] },
};

// Entidades con plata adentro: nunca para gestora, y para admin solo
// las que no son de dinero.
const ENTIDADES_DINERO = ['Pauta','Tasas','Facturacion'];

function puede(s, accion, entidad) {
  const p = PERMISOS[s.rol] || PERMISOS.gestora;
  const lista = p[accion];
  if (!lista) return false;
  if (lista === '*') return true;
  return lista.indexOf(entidad) !== -1;
}

/**
 * Qué tan arriba va cada estado en una lista de trabajo.
 *
 * 0 a 2 es lo que necesita a alguien hoy. 3 es lo que Nova no reconoció:
 * va arriba a propósito, porque un estado desconocido es justo lo que hay
 * que mirar. De 4 en adelante está en curso, y de 7 ya terminó.
 */
const PRIORIDAD_ESTADO = {
  // Pedidos
  novedad: 0, pendiente: 1, en_oficina: 2,
  novedad_resuelta: 4, confirmado: 5, en_bodega: 5, en_transito: 6,
  entregado: 7, devolucion: 8, cancelado: 9,
  // Novedades: los mismos criterios, sobre su propia columna estado
  abierta: 0, resuelta: 7, cerrada: 8,
};

function prioridadEstado(v) {
  const e = String(v || '').trim();
  if (!e) return 3;
  const p = PRIORIDAD_ESTADO[e];
  return p === undefined ? 3 : p;
}

/** La gestora solo ve lo suyo. Se aplica al leer, no al pintar. */
function filtrarPorRol(s, entidad, filas, enc) {
  if (s.rol !== 'gestora') return filas;
  const cg = enc.indexOf('gestora_asignada') !== -1
    ? enc.indexOf('gestora_asignada') : enc.indexOf('gestora');
  if (cg === -1) return filas;
  const mio = norm(s.nombre);
  return filas.filter(function (f) { return norm(f[cg]) === mio; });
}

// ─── LECTURA ─────────────────────────────────────────────────

function apiListar(s, p) {
  const entidad = String(p.entidad || '').trim();
  if (!puede(s, 'leer', entidad)) {
    return { ok: false, error: 'Tu rol no tiene acceso a ' + entidad + '.' };
  }
  if (ENTIDADES_DINERO.indexOf(entidad) !== -1 && s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña ve ' + entidad + '.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName(entidad);
  if (!sh) return { ok: false, error: 'No existe la hoja ' + entidad + '.' };
  if (sh.getLastRow() < 2) return { ok: true, filas: [], total: 0 };

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  let filas = datos.slice(1).filter(function (f) {
    return f.some(function (c) { return String(c).trim() !== ''; });
  });

  // Tienda: siempre dentro de las que la persona tiene asignadas
  const cT = enc.indexOf('tienda');
  if (cT !== -1) {
    const pedida = String(p.tienda || '').trim();
    const permitidas = pedida
      ? (s.tiendas.indexOf(pedida) !== -1 ? [pedida] : [])
      : s.tiendas;
    filas = filas.filter(function (f) {
      return permitidas.indexOf(String(f[cT]).trim()) !== -1;
    });
  }

  // Rango de fechas
  const cF = enc.indexOf('fecha') !== -1 ? enc.indexOf('fecha') : enc.indexOf('fecha_hora');
  if (cF !== -1 && (p.desde || p.hasta)) {
    filas = filas.filter(function (f) {
      const d = aISO(f[cF], 'UTC');
      if (!d) return false;
      if (p.desde && d < p.desde) return false;
      if (p.hasta && d > p.hasta) return false;
      return true;
    });
  }

  filas = filtrarPorRol(s, entidad, filas, enc);

  /**
   * El orden manda, y no es la fecha.
   *
   * Antes salían en el orden de la hoja, y como el importador agrega al
   * final, los pedidos de septiembre quedaban debajo de los doscientos de
   * agosto: pidiendo las primeras trescientas filas, el mes en curso
   * podía no aparecer nunca.
   *
   * Pero ordenar por fecha tampoco sirve. Una lista de pedidos es una
   * lista de trabajo: lo primero tiene que ser lo que hay que resolver
   * hoy —una novedad abierta, un pedido sin confirmar, uno esperando en
   * oficina— y lo último, lo que ya terminó. Un entregado de esta mañana
   * no le gana a una novedad de hace tres días.
   *
   * Y dentro de lo que necesita acción, primero lo más viejo: ahí la
   * antigüedad es deuda, no historia.
   */
  // Las novedades no tienen estado_canonico: su urgencia vive en `estado`
  const cE = enc.indexOf('estado_canonico') !== -1
    ? enc.indexOf('estado_canonico') : enc.indexOf('estado');
  if (cE !== -1 || cF !== -1) {
    filas.sort(function (a, b) {
      const pa = cE === -1 ? 5 : prioridadEstado(a[cE]);
      const pb = cE === -1 ? 5 : prioridadEstado(b[cE]);
      if (pa !== pb) return pa - pb;
      if (cF === -1) return 0;
      const fa = aISO(a[cF], 'UTC') || '';
      const fb = aISO(b[cF], 'UTC') || '';
      if (fa === fb) return 0;
      // Lo pendiente: primero lo más viejo. Lo cerrado: lo más reciente.
      return pa <= 2 ? (fa < fb ? -1 : 1) : (fa < fb ? 1 : -1);
    });
  }

  const total = filas.length;
  const desde = Math.max(0, parseInt(p.offset || 0, 10));
  const cuantas = Math.min(500, Math.max(1, parseInt(p.limite || 200, 10)));
  const pagina = filas.slice(desde, desde + cuantas);

  return {
    ok: true, total: total, offset: desde,
    columnas: datos[0],
    filas: pagina.map(function (f) {
      const o = {};
      enc.forEach(function (k, i) { o[k] = valorLimpio(f[i]); });
      return o;
    }),
  };
}

/** Las fechas salen en ISO; lo demás tal cual. */
function valorLimpio(v) {
  if (v instanceof Date) return Utilities.formatDate(v, 'UTC', 'yyyy-MM-dd');
  return v;
}

// ─── ESCRITURA ───────────────────────────────────────────────

/**
 * Escribe CELDA por CELDA, no la fila completa.
 *
 * Sheets no tiene bloqueo de filas: si la admin reasigna un caso mientras
 * la gestora lo resuelve y ambas escriben la fila entera, uno de los dos
 * cambios se pierde en silencio. Escribiendo solo las celdas que
 * cambiaron, los dos sobreviven.
 *
 * Y nunca toca columnas importadas: esas son de solo lectura para la app.
 */
const COLUMNAS_IMPORTADAS = [
  'id','fuente','id_externo','fecha','valor','costo_producto','costo_envio',
  'cpm','cpa','gasto','impresiones','clics','resultados','tienda',
];

/**
 * Crea una fila nueva.
 *
 * Solo en las entidades que la app puede crear, y con los campos que
 * declara cada una. Una acción genérica de "inserta lo que te manden"
 * dejaría escribir en cualquier hoja cualquier cosa, incluidas las
 * columnas que el importador y el equipo se reparten.
 */
const CREABLES = {
  Gastos:     ['tienda', 'mes', 'tipo', 'nombre', 'valor', 'moneda', 'nota'],
  // `origen` no está: lo pone el servidor. Es el rastro de quién dio ese
  // número —una persona contando o un archivo importado— y dejar que lo
  // mande el cliente sería dejar que un conteo a mano se firme como si
  // hubiera venido de la plataforma.
  Inventario: ['tienda', 'sku', 'producto', 'stock', 'costo_unitario',
               'precio', 'precio_2', 'precio_3',
               'minimo', 'categoria', 'proveedor', 'landing',
               'resp_1', 'resp_2', 'resp_3', 'resp_4', 'nota'],
  Equipo:     ['nombre', 'correo', 'rol', 'tienda', 'estado', 'permisos'],
};

/** Quién puede crear o quitar en cada entidad. */
const SOLO_DUENO = ['Equipo', 'Gastos'];

/**
 * El recuento del periodo: cómo viene la tienda.
 *
 * Mira los últimos meses COMPLETOS, no el que está corriendo: comparar
 * un mes a medias contra uno entero siempre dice que vas peor, y no es
 * verdad, es que todavía no termina.
 *
 * Si hay tres o más, habla de trimestre. Si hay menos —una tienda que
 * apenas empieza— habla del último mes cerrado y lo dice, en vez de
 * inventar un trimestre con un mes adentro.
 *
 * Lo que avanzó no se escribe a mano: sale de comparar el primer mes
 * del periodo contra el último.
 */
function apiRecuento(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const tz = zonaHorariaDe(ss, tienda) || 'UTC';
  const mesActual = Utilities.formatDate(new Date(), tz, 'yyyy-MM');

  // Hasta seis meses atrás, quedándonos con los que tienen movimiento
  const conDatos = [];
  let m = mesAnterior(mesActual);
  for (let i = 0; i < 6 && conDatos.length < 3; i++) {
    const d = agregarMes(ss, tienda, m, s);
    if (d.pedidos > 0) conDatos.push({ mes: m, d: d });
    m = mesAnterior(m);
  }
  if (!conDatos.length) {
    return { ok: true, tienda: tienda, vacio: true, mesActual: mesActual };
  }

  conDatos.reverse();                      // del más viejo al más nuevo
  const periodo = conDatos.length >= 3 ? 'trimestre' : 'mes';
  const usados = periodo === 'trimestre' ? conDatos : [conDatos[conDatos.length - 1]];

  const t = { pedidos: 0, entregados: 0, devueltos: 0, resueltos: 0,
              ventas: 0, gasto: 0, fijos: 0, costoProducto: 0, costoEnvio: 0 };
  usados.forEach(function (x) {
    ['pedidos','entregados','devueltos','resueltos','ventas','gasto','fijos',
     'costoProducto','costoEnvio'].forEach(function (k) {
      t[k] += x.d[k] || 0;
    });
  });
  t.efectividad = t.resueltos ? t.entregados / t.resueltos * 100 : 0;
  t.roas = t.gasto ? t.ventas / t.gasto : 0;

  // Con qué comparamos: el periodo anterior de la misma longitud
  const previos = [];
  let pm = mesAnterior(usados[0].mes);
  for (let i = 0; i < usados.length; i++) { previos.push(pm); pm = mesAnterior(pm); }
  const ant = { ventas: 0, entregados: 0, resueltos: 0, gasto: 0, fijos: 0 };
  previos.forEach(function (mm) {
    const d = agregarMes(ss, tienda, mm, s);
    ant.ventas += d.ventas || 0; ant.entregados += d.entregados || 0;
    ant.resueltos += d.resueltos || 0; ant.gasto += d.gasto || 0;
    ant.fijos += d.fijos || 0;
  });
  ant.efectividad = ant.resueltos ? ant.entregados / ant.resueltos * 100 : 0;

  /**
   * Lo que avanzó: se comparan el primer y el último mes del periodo y se
   * queda lo que más se movió. Sin datos anteriores no se dice nada: un
   * "creció 100%" contra un mes que no existió es ruido.
   */
  const avances = [];
  const pri = usados[0].d, ult = usados[usados.length - 1].d;
  if (usados.length > 1) {
    const dEf = (ult.efectividad || 0) - (pri.efectividad || 0);
    if (Math.abs(dEf) >= 3) {
      avances.push({ signo: dEf > 0 ? '+' : '', valor: dEf.toFixed(0) + ' pts',
        titulo: 'La tasa de entrega ' + (dEf > 0 ? 'mejoró' : 'cayó') + ' de ' +
          (pri.efectividad||0).toFixed(0) + '% a ' + (ult.efectividad||0).toFixed(0) + '%',
        detalle: 'Entre ' + usados[0].mes + ' y ' + usados[usados.length-1].mes + '.',
        bueno: dEf > 0 });
    }
    const dV = (ult.ventas || 0) - (pri.ventas || 0);
    if (pri.ventas && Math.abs(dV / pri.ventas) >= 0.1) {
      avances.push({ signo: dV > 0 ? '+' : '', valor: Math.round(dV / pri.ventas * 100) + '%',
        titulo: 'Las ventas ' + (dV > 0 ? 'subieron' : 'bajaron') + ' de ' +
          Math.round(pri.ventas) + ' a ' + Math.round(ult.ventas),
        detalle: 'Mes a mes dentro del periodo.', bueno: dV > 0 });
    }
  }
  if (ant.ventas) {
    const dv = (t.ventas - ant.ventas) / ant.ventas * 100;
    if (Math.abs(dv) >= 5) {
      avances.push({ signo: dv > 0 ? '+' : '', valor: dv.toFixed(0) + '%',
        titulo: 'Ingresos ' + (dv > 0 ? 'por encima' : 'por debajo') + ' del ' +
          (periodo === 'mes' ? 'mes' : 'trimestre') + ' anterior',
        detalle: Math.round(t.ventas) + ' contra ' + Math.round(ant.ventas) + '.',
        bueno: dv > 0 });
    }
  }

  return { ok: true, tienda: tienda, periodo: periodo,
           meses: usados.map(function (x) { return x.mes; }),
           total: t, anterior: ant, avances: avances.slice(0, 3),
           moneda: monedaDeTienda(ss, tienda), mesActual: mesActual,
           mesesConDatos: conDatos.length };
}

/**
 * Las alarmas de la tienda, calculadas al momento.
 *
 * La pantalla y el correo diario usan exactamente esta función. Si
 * fueran dos cálculos distintos, tarde o temprano dirían cosas distintas
 * y no habría forma de saber a cuál creerle.
 */
function apiAlarmas(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const r = evaluarAlarmas(ss, tienda);

  // La de dinero es solo de la dueña
  const alarmas = s.rol === 'dueno'
    ? r.alarmas
    : r.alarmas.filter(function (a) { return a.id !== 'cpa'; });

  // Los umbrales son de la dueña; el canal del equipo es de todo el
  // equipo. Va aquí y no en su propia llamada porque la pantalla ya hace
  // esta, y una llamada más por un enlace sería una llamada de más.
  return { ok: true, tienda: tienda, alarmas: alarmas,
           umbrales: s.rol === 'dueno' ? r.umbrales : null,
           ajustes: ajustes(ss, tienda),
           catalogo: ALARMAS };
}

/**
 * Leer y cambiar los umbrales.
 *
 * Solo la dueña: son las reglas con las que su propia operación se
 * juzga, y cualquiera que pudiera moverlas podría apagar la alarma que
 * lo señala.
 */
function apiParametros(s, p) {
  const ss = SpreadsheetApp.openById(s.sheetId);
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }

  if (!p.cambios) {
    return { ok: true, tienda: tienda, umbrales: umbrales(ss, tienda),
             ajustes: ajustes(ss, tienda),
             catalogo: ALARMAS, defaults: ALARMAS_DEFAULT };
  }

  if (s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña cambia la configuración de la tienda.' };
  }

  /**
   * El canal se valida antes de tocar la hoja.
   *
   * Si una de las dos claves está mal, no se escribe ninguna: guardar el
   * nombre del canal y rechazar su enlace dejaría un botón con etiqueta
   * y sin destino, que es peor que no tener botón.
   */
  if (p.cambios.canal_url !== undefined) {
    const err = validarCanal(p.cambios.canal_url);
    if (err) return { ok: false, error: err };
  }

  let sh = ss.getSheetByName('Parametros');
  if (!sh) return { ok: false, error: 'Falta la hoja Parametros. Corre bootstrapTodo().' };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda'), cK = e.indexOf('clave'), cV = e.indexOf('valor');
  const cA = e.indexOf('actualizado_en'), cP = e.indexOf('actualizado_por');

  Object.keys(p.cambios).forEach(function (clave) {
    // Nada fuera de los dos catálogos: umbrales de alarma y ajustes de
    // la tienda. Aceptar una clave cualquiera convertiría esta acción en
    // "escribe lo que quieras en Parametros".
    if (!(clave in ALARMAS_DEFAULT) && !(clave in AJUSTES_DEFAULT)) return;
    const valor = p.cambios[clave];
    let fila = -1;
    for (let i = 1; i < d.length; i++) {
      if (norm(d[i][cK]) === norm(clave) &&
          String(d[i][cT] || '').trim() === tienda) { fila = i; break; }
    }
    if (fila === -1) {
      const nueva = new Array(e.length).fill('');
      nueva[cT] = tienda; nueva[cK] = clave; nueva[cV] = valor;
      if (cA !== -1) nueva[cA] = ahoraISO();
      if (cP !== -1) nueva[cP] = s.email;
      sh.appendRow(nueva);
      registrarMovimiento(s, 'Parametros', tienda + '/' + clave, 'valor', '', valor);
    } else {
      const antes = d[fila][cV];
      if (String(antes) === String(valor)) return;
      sh.getRange(fila + 1, cV + 1).setValue(valor);
      if (cA !== -1) sh.getRange(fila + 1, cA + 1).setValue(ahoraISO());
      if (cP !== -1) sh.getRange(fila + 1, cP + 1).setValue(s.email);
      registrarMovimiento(s, 'Parametros', tienda + '/' + clave, 'valor', antes, valor);
    }
  });
  SpreadsheetApp.flush();
  return { ok: true, umbrales: umbrales(ss, tienda), ajustes: ajustes(ss, tienda) };
}

/** Las categorías que la dueña puede ponerle a un producto. */
const CATEGORIAS_PRODUCTO = ['estrella', 'complemento', 'testeo', 'frenado'];

/**
 * Lo que una ficha de inventario tiene que cumplir para entrar.
 *
 * El nombre es lo único imprescindible: es la llave con la que la ficha se
 * encuentra con los pedidos. Sin él la ficha existe pero no se junta con
 * nada, y el producto sigue apareciendo como si no tuviera ficha.
 */
function validarFicha(d) {
  if (d.producto !== undefined && !String(d.producto || '').trim()) {
    return 'La ficha necesita el nombre del producto: es con lo que se ' +
           'encuentra con los pedidos.';
  }
  const cat = norm(d.categoria || '');
  if (cat && CATEGORIAS_PRODUCTO.indexOf(cat) === -1) {
    return 'La categoría debe ser una de: ' + CATEGORIAS_PRODUCTO.join(', ') + '.';
  }
  const negativo = ['stock', 'minimo', 'costo_unitario', 'precio',
                    'precio_2', 'precio_3'].filter(function (k) {
    return d[k] !== undefined && d[k] !== '' && num(d[k]) < 0;
  });
  if (negativo.length) return 'No puede haber números negativos en ' + negativo.join(', ') + '.';

  /**
   * Un combo tiene que costar más que una unidad suelta, o el precio está
   * mal escrito. Se avisa en vez de aceptarlo: un 2x más barato que un 1x
   * no es una promoción, es un error de tecleo que después aparece como
   * un margen raro sin que nadie sepa de dónde salió.
   */
  const p1 = num(d.precio), p2 = num(d.precio_2), p3 = num(d.precio_3);
  if (p1 && p2 && p2 < p1) {
    return 'El precio de 2 unidades (' + p2 + ') es menor que el de 1 (' + p1 +
           '). Si es a propósito, dilo en la nota; si no, revísalo.';
  }
  if (p2 && p3 && p3 < p2) {
    return 'El precio de 3 unidades (' + p3 + ') es menor que el de 2 (' + p2 + ').';
  }

  // La landing se abre desde la app: mismo filtro que el canal del equipo
  if (d.landing !== undefined && String(d.landing).trim()) {
    const err = validarCanal(d.landing);
    if (err) return 'La landing: ' + err;
  }
  return '';
}

/** Lo que puede decir `origen`: de dónde salió el número, no qué es. */
const ORIGENES_INVENTARIO = ['manual', 'importado', 'archivo', 'plataforma'];

/**
 * La categoría de una ficha, tolerando las hojas viejas.
 *
 * Antes `origen` hacía los dos trabajos. Si esa columna todavía guarda una
 * categoría se respeta; si guarda una procedencia, no se confunde con una.
 */
function categoriaDeFicha(categoria, origen) {
  const cat = String(categoria || '').trim();
  if (cat) return cat;
  const org = String(origen || '').trim();
  return ORIGENES_INVENTARIO.indexOf(norm(org)) === -1 ? org : '';
}

/**
 * El catálogo, sacado de los pedidos.
 *
 * Nadie tiene que escribir una lista de productos: ya están todos en los
 * pedidos, repetidos. Aquí se agrupan y se mide cómo le va a cada uno.
 *
 * Lo que la ficha agrega es lo que ningún archivo trae: el costo real,
 * el precio de lista, la categoría, si es un testeo. Por eso la ficha y
 * el conteo viven aparte: el conteo se recalcula solo, la ficha es tuya.
 *
 * Un producto que vende y no tiene ficha no se esconde: se muestra con el
 * aviso de qué le falta. Sin ficha no se puede saber si deja plata, y un
 * producto que no sabes si deja plata es justo el que hay que mirar.
 */
function apiProductos(s, p) {
  /**
   * El catálogo se arma leyendo TODOS los pedidos de la tienda, no solo
   * los de quien pregunta. Para una gestora eso sería ver el negocio
   * entero por una puerta lateral, así que esta puerta no es suya.
   */
  if (!puede(s, 'leer', 'Inventario')) {
    return { ok: false, error: 'Tu rol no ve el catálogo de productos.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const desde = String(p.desde || '').trim();   // 'AAAA-MM', opcional

  const cat = {};
  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (desde && (!fecha || fecha.slice(0, 7) < desde)) continue;

      const nombre = String(f[c('producto')] || '').trim() || '(sin nombre)';
      const clave = norm(nombre);
      if (!cat[clave]) {
        cat[clave] = { nombre: nombre, sku: String(f[c('sku')] || '').trim(),
                       pedidos: 0, entregados: 0, devueltos: 0, cancelados: 0,
                       ventas: 0, unidades: 0, unidadesEntregadas: 0,
                       costoProducto: 0,
                       primera: fecha || '', ultima: fecha || '' };
      }
      const x = cat[clave];
      const unids = num(f[c('cantidad')]) || 1;
      x.pedidos++;
      x.unidades += unids;
      if (!x.sku) x.sku = String(f[c('sku')] || '').trim();
      if (fecha) {
        if (!x.primera || fecha < x.primera) x.primera = fecha;
        if (!x.ultima  || fecha > x.ultima)  x.ultima  = fecha;
      }
      const est = norm(f[c('estado_nova')] || f[c('estado_canonico')]);
      // Del inventario solo salen las unidades que se entregaron. Una
      // devolución vuelve a la bodega y un cancelado nunca salió: contarlos
      // como consumo haría creer que el stock se agota más rápido de lo que
      // se agota, y mandaría a reponer de más.
      if (est === 'entregado') {
        x.entregados++;
        x.ventas += num(f[c('valor')]);
        x.unidadesEntregadas += unids;
      }
      if (est === 'devolucion') x.devueltos++;
      if (est === 'cancelado')  x.cancelados++;
      x.costoProducto += num(f[c('costo_producto')]);
    }
  }

  // La ficha que completa la dueña, si existe
  const fichas = {};
  const shI = ss.getSheetByName('Inventario');
  if (shI && shI.getLastRow() > 1) {
    const d = shI.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      if (norm(f[c('activo')]) === 'no') continue;
      const nombreFicha = String(f[c('producto')] || '').trim();
      fichas[norm(nombreFicha)] = {
        id: f[c('id')], sku: String(f[c('sku')] || ''),
        // El nombre tal como lo escribió la dueña. Sin esto, un producto
        // con ficha pero sin ventas salía en pantalla con su clave
        // normalizada —minúsculas y sin tildes— como si fuera su nombre.
        nombre: nombreFicha,
        stock: num(f[c('stock')]), minimo: num(f[c('minimo')]),
        costo: num(f[c('costo_unitario')]), precio: num(f[c('precio')]),
        // La escalera de precios: lo que de verdad se cobra por combo
        precio2: c('precio_2') !== -1 ? num(f[c('precio_2')]) : 0,
        precio3: c('precio_3') !== -1 ? num(f[c('precio_3')]) : 0,
        landing: String((c('landing') !== -1 ? f[c('landing')] : '') || '').trim(),
        // Las respuestas a las cuatro preguntas de siempre
        respuestas: [1, 2, 3, 4].map(function (n) {
          return String((c('resp_' + n) !== -1 ? f[c('resp_' + n)] : '') || '').trim();
        }),
        // `categoria` es nueva. Las hojas escritas antes guardaban esto en
        // `origen`, así que se lee de ahí mientras nadie la haya llenado —
        // pero solo si lo que dice no es una palabra de procedencia, que
        // es para lo que `origen` sirve de ahora en adelante.
        categoria: categoriaDeFicha(c('categoria') !== -1 ? f[c('categoria')] : '',
                                    f[c('origen')]),
        proveedor: String((c('proveedor') !== -1 ? f[c('proveedor')] : '') || ''),
        nota: String(f[c('nota')] || ''),
        ultimoConteo: aISO(f[c('ultimo_conteo')], 'UTC') || '',
      };
    }
  }

  /**
   * A cuántos días de stock estás.
   *
   * El ritmo sale de las unidades que de verdad se entregaron, repartidas
   * entre los días que van desde el primer pedido hasta hoy —no hasta el
   * último pedido: si hace dos semanas que no vende, esas dos semanas son
   * parte del ritmo, y esconderlas diría que rota más rápido de lo que rota.
   *
   * Con menos de 14 días de historia no se calcula. Tres entregas en dos
   * días darían un ritmo de 1,5 al día y una cobertura que se derrumba
   * sola; es mejor decir que todavía no se sabe.
   */
  const hoyISO = new Date().toISOString().slice(0, 10);
  const MIN_DIAS_RITMO = 14;
  const cobertura = function (x, ficha) {
    const base = { ritmo: null, dias: null, ventana: 0 };
    if (!x || !x.primera) return base;
    const ini = new Date(x.primera + 'T00:00:00Z').getTime();
    const fin = new Date(hoyISO + 'T00:00:00Z').getTime();
    const ventana = Math.floor((fin - ini) / 86400000) + 1;
    base.ventana = ventana;
    if (ventana < MIN_DIAS_RITMO || !x.unidadesEntregadas) return base;
    base.ritmo = x.unidadesEntregadas / ventana;
    if (ficha && ficha.stock > 0) base.dias = Math.floor(ficha.stock / base.ritmo);
    return base;
  };

  const salida = Object.keys(cat).map(function (k) {
    const x = cat[k];
    const ficha = fichas[k] || null;
    const resueltos = x.entregados + x.devueltos;
    const cob = cobertura(x, ficha);

    // Qué le falta a este producto para poder decidir sobre él
    const falta = [];
    if (!ficha) falta.push('ficha');
    else {
      if (!ficha.costo)  falta.push('costo');
      if (!ficha.precio) falta.push('precio');
      if (!ficha.categoria) falta.push('categoría');
    }
    if (!x.sku && (!ficha || !ficha.sku)) falta.push('sku');

    return {
      clave: k, nombre: x.nombre, sku: x.sku || (ficha ? ficha.sku : ''),
      pedidos: x.pedidos, entregados: x.entregados, devueltos: x.devueltos,
      cancelados: x.cancelados, unidades: x.unidades, ventas: x.ventas,
      unidadesEntregadas: x.unidadesEntregadas,
      ticket: x.entregados ? x.ventas / x.entregados : 0,
      entrega: resueltos ? x.entregados / resueltos * 100 : 0,
      primera: x.primera, ultima: x.ultima,
      ficha: ficha, falta: falta,
      ritmo: cob.ritmo, coberturaDias: cob.dias, ventanaDias: cob.ventana,
      // Margen unitario solo si hay con qué calcularlo. Si no, null: un
      // margen estimado sobre un costo inventado es peor que no tenerlo.
      margen: (ficha && ficha.costo && x.entregados)
        ? (x.ventas / x.entregados) - ficha.costo : null,
    };
  }).sort(function (a, b) { return b.pedidos - a.pedidos; });

  // Fichas de productos que todavía no han vendido nada
  Object.keys(fichas).forEach(function (k) {
    if (cat[k]) return;
    const fi = fichas[k];
    salida.push({ clave: k, nombre: fi.nombre || k, sku: fi.sku, pedidos: 0,
                  entregados: 0, devueltos: 0, cancelados: 0, unidades: 0,
                  unidadesEntregadas: 0,
                  ventas: 0, ticket: 0, entrega: 0, primera: '', ultima: '',
                  ficha: fi, falta: [], margen: null,
                  ritmo: null, coberturaDias: null, ventanaDias: 0,
                  sinVentas: true });
  });

  // Las preguntas que esta tienda decidió que son las suyas
  const preg = String(ajustes(ss, tienda).preguntas_producto || '')
    .split('|').map(function (x) { return x.trim(); }).filter(String).slice(0, 4);

  /**
   * Qué columnas tiene DE VERDAD la hoja Inventario.
   *
   * Cuando Nova gana un campo nuevo, la hoja del cliente no lo tiene
   * hasta que alguien corre bootstrapTodo(). Mientras tanto, guardar una
   * ficha mandaba campos que no existían y el servidor los rechazaba —
   * sin que quedara claro que el problema era una migración pendiente y
   * no lo que se había escrito.
   *
   * Diciéndolo, la pantalla puede mandar solo lo que cabe y avisar de lo
   * que falta. Guardar a medias con explicación es mucho mejor que no
   * guardar sin ella.
   */
  let columnas = [];
  if (shI && shI.getLastColumn() > 0) {
    columnas = shI.getRange(1, 1, 1, shI.getLastColumn()).getValues()[0]
      .map(norm).filter(String);
  }

  return { ok: true, tienda: tienda, productos: salida, preguntas: preg,
           columnas: columnas,
           modalidad: modalidadDeTienda(ss, tienda),
           moneda: monedaDeTienda(ss, tienda) };
}

/** Cómo consigue el stock esta tienda: catálogo público, privado o marca propia. */
function modalidadDeTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return 'catalogo_publico';
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cM = e.indexOf('modalidad');
  if (cM === -1) return 'catalogo_publico';
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][cId]).trim() === tienda) {
      return String(d[i][cM] || '').trim() || 'catalogo_publico';
    }
  }
  return 'catalogo_publico';
}

/**
 * El equipo, con lo que cada persona hizo en el mes.
 *
 * El rendimiento no se escribe en ninguna parte: se cuenta sobre los
 * pedidos y las novedades que tiene asignados. Así no hay una cifra que
 * mantener al día a mano, y nadie puede maquillarla.
 */
function apiEquipo(s, p) {
  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Equipo');
  /**
   * `puedeEditar` también va aquí, y no solo en la salida de abajo.
   *
   * Cuando faltaba, una cuenta con el Equipo vacío le respondía a su
   * propia dueña que no podía editar: la pantalla escondía el formulario
   * y los botones, y quedaba sin forma de agregar a la primera persona.
   * Era el peor momento posible para negar el permiso — justo el día uno.
   */
  if (!sh || sh.getLastRow() < 2) {
    return { ok: true, personas: [], puedeEditar: s.rol === 'dueno' };
  }

  const mes = String(p.mes || Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM'));
  const tienda = String(p.tienda || '').trim();

  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const c = function (n) { return enc.indexOf(n); };

  const personas = [];
  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    if (!String(f[c('correo')] || '').trim()) continue;
    const rol = rolCanonico(f[c('rol')]);
    personas.push({
      id: f[c('id')], nombre: f[c('nombre')], correo: f[c('correo')],
      rol: rol || String(f[c('rol')] || ''),
      rolLegible: rol ? '' : 'no reconocido',
      tienda: String(f[c('tienda')] || '*'),
      estado: norm(f[c('estado')]) === 'inactivo' ? 'inactivo' : 'activo',
      permisos: c('permisos') === -1 ? '' : String(f[c('permisos')] || ''),
      ultima_conexion: c('ultima_conexion') === -1 ? '' : f[c('ultima_conexion')],
      pedidos: 0, entregados: 0, novedades: 0, resueltas: 0, sinMover: 0,
    });
  }

  const porNombre = {};
  personas.forEach(function (x) { porNombre[norm(x.nombre)] = x; });

  // Pedidos del mes, por gestora asignada
  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const cc = function (n) { return e.indexOf(n); };
    const hoy = new Date();
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (tienda && String(f[cc('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[cc('fecha')], 'UTC');
      if (!fecha || fecha.slice(0, 7) !== mes) continue;
      const g = porNombre[norm(f[cc('gestora_asignada')])];
      if (!g) continue;
      g.pedidos++;
      const est = norm(f[cc('estado_nova')] || f[cc('estado_canonico')]);
      if (est === 'entregado') g.entregados++;
      if (['entregado','devolucion','cancelado'].indexOf(est) === -1) {
        const ult = aISO(f[cc('ultimo_movimiento')] || f[cc('actualizado_en')], 'UTC') || fecha;
        if ((hoy - new Date(ult + 'T00:00:00Z')) / 86400000 > 3) g.sinMover++;
      }
    }
  }

  // Novedades del mes, por gestora
  const shN = ss.getSheetByName('Novedades');
  if (shN && shN.getLastRow() > 1) {
    const d = shN.getDataRange().getValues();
    const e = d[0].map(norm);
    const cF = e.indexOf('fecha'), cG = e.indexOf('gestora'), cE = e.indexOf('estado');
    for (let i = 1; i < d.length; i++) {
      const fecha = aISO(d[i][cF], 'UTC');
      if (!fecha || fecha.slice(0, 7) !== mes) continue;
      const g = porNombre[norm(d[i][cG])];
      if (!g) continue;
      g.novedades++;
      if (norm(d[i][cE]) === 'resuelta') g.resueltas++;
    }
  }

  // La jornada de hoy sale de la bitácora, y se cruza por correo: el
  // nombre se puede escribir de tres formas distintas en tres hojas, el
  // correo es el mismo con el que entró.
  const jornada = jornadaDeHoy(ss, tienda);
  personas.forEach(function (x) {
    x.efectividad = x.pedidos ? x.entregados / x.pedidos * 100 : 0;
    x.hoy = jornada[String(x.correo || '').trim().toLowerCase()] || null;
  });

  // La gestora solo se ve a sí misma
  const salida = s.rol === 'gestora'
    ? personas.filter(function (x) { return norm(x.nombre) === norm(s.nombre); })
    : personas;

  return { ok: true, mes: mes, personas: salida, puedeEditar: s.rol === 'dueno' };
}

/**
 * El rastro de quién cambió qué.
 *
 * Sale de Movimientos, que se escribe solo en cada edición. La gestora no
 * lo ve: es información sobre el equipo, no para el equipo.
 */
function apiAuditoria(s, p) {
  if (s.rol === 'gestora') return { ok: false, error: 'No tienes acceso a la auditoría.' };

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Movimientos');
  if (!sh || sh.getLastRow() < 2) return { ok: true, movimientos: [] };

  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const c = function (n) { return enc.indexOf(n); };
  const cuantas = Math.min(300, Math.max(1, parseInt(p.limite || 150, 10)));

  const out = [];
  for (let i = filas.length - 1; i >= 1 && out.length < cuantas; i--) {
    const f = filas[i];
    if (!String(f[c('fecha')] || '').trim()) continue;
    out.push({
      fecha: aISO(f[c('fecha')], 'UTC') || String(f[c('fecha')]),
      hora: String(f[c('fecha')]).slice(11, 16),
      usuario: f[c('usuario')], entidad: f[c('entidad')],
      entidad_id: f[c('entidad_id')], campo: f[c('campo')],
      antes: f[c('valor_anterior')], despues: f[c('valor_nuevo')],
    });
  }
  return { ok: true, movimientos: out };
}

/**
 * Los estados que cada plataforma usa, y qué sabe Nova de cada uno.
 *
 * Primero los que no entiende, ordenados por cuántos pedidos afectan: uno
 * suelto es ruido, doscientos es un cierre mal hecho esperando a pasar.
 */
function apiEstados(s, p) {
  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Estados');
  const out = { ok: true, sinClasificar: [], conocidos: [],
                opciones: OPCIONES_ESTADO, puedeEditar: s.rol === 'dueno' };
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  for (let i = 1; i < d.length; i++) {
    const texto = String(d[i][c('texto')] || '').trim();
    if (!texto) continue;
    const fila = {
      fuente: String(d[i][c('fuente')] || ''), texto: texto,
      estado: norm(d[i][c('estado_nova')]), origen: norm(d[i][c('origen')]),
      pedidos: num(d[i][c('pedidos')]),
      primera: d[i][c('primera_vez')], ultima: d[i][c('ultima_vez')],
      nota: String(d[i][c('nota')] || ''),
    };
    if (!fila.estado) out.sinClasificar.push(fila);
    else out.conocidos.push(fila);
  }

  out.sinClasificar.sort(function (a, b) { return b.pedidos - a.pedidos; });
  out.conocidos.sort(function (a, b) { return b.pedidos - a.pedidos; });
  return out;
}

/** A qué puede equivaler un estado, en palabras de quien va a elegir. */
const OPCIONES_ESTADO = [
  { id: 'entregado',  nombre: 'Entregado',
    ayuda: 'Llegó y se cobró. Cuenta como venta.', terminal: true },
  { id: 'devolucion', nombre: 'Devuelto',
    ayuda: 'Volvió. No es venta, y su flete se paga igual.', terminal: true },
  { id: 'cancelado',  nombre: 'Cancelado',
    ayuda: 'Nunca salió. No cuesta nada.', terminal: true },
  { id: 'en_transito', nombre: 'En camino',
    ayuda: 'Salió de bodega y va para allá.' },
  { id: 'en_bodega',  nombre: 'En bodega',
    ayuda: 'Todavía no sale.' },
  { id: 'en_oficina', nombre: 'En oficina',
    ayuda: 'Esperando que el cliente lo recoja.' },
  { id: 'novedad',    nombre: 'Con novedad',
    ayuda: 'Hubo un problema y hay que gestionarlo.' },
  { id: 'confirmado', nombre: 'Confirmado',
    ayuda: 'Confirmado con el cliente, sin despachar.' },
  { id: 'pendiente',  nombre: 'Pendiente',
    ayuda: 'Sin confirmar todavía.' },
];

/**
 * La dueña dice qué significa un estado.
 *
 * Queda con origen "manual" y desde ese momento manda sobre las tablas
 * del código: si ella dice que en SU operación ese estado significa otra
 * cosa, tiene razón — quien conoce su operación es ella.
 *
 * No se reescriben los pedidos ya guardados. No hace falta: cada pantalla
 * recalcula desde Pedidos, y el estado se vuelve a traducir al leer. Lo
 * que sí queda avisado son los meses ya cerrados, porque esos están
 * congelados a propósito y nadie debería cambiarlos a espaldas de quien
 * los reportó.
 */
function apiEstadoClasificar(s, p) {
  if (s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña decide qué significa un estado.' };
  }
  const fuente = norm(p.fuente), texto = norm(p.texto), estado = norm(p.estado);
  if (!fuente || !texto) return { ok: false, error: 'Falta el estado a clasificar.' };

  const validos = OPCIONES_ESTADO.map(function (o) { return o.id; });
  if (estado && validos.indexOf(estado) === -1) {
    return { ok: false, error: 'No conozco el estado "' + estado + '".' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Estados');
  if (!sh) return { ok: false, error: 'Falta la hoja Estados. Corre bootstrapTodo().' };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][c('fuente')]) !== fuente || norm(d[i][c('texto')]) !== texto) continue;

    const antes = norm(d[i][c('estado_nova')]);
    sh.getRange(i + 1, c('estado_nova') + 1).setValue(estado);
    sh.getRange(i + 1, c('origen') + 1).setValue(estado ? 'manual' : 'nuevo');
    sh.getRange(i + 1, c('decidido_por') + 1).setValue(s.email);
    if (p.nota !== undefined) sh.getRange(i + 1, c('nota') + 1).setValue(String(p.nota));

    registrarMovimiento(s, 'Estados', fuente + ' · ' + texto,
                        'estado_nova', antes, estado);

    return { ok: true, fuente: fuente, texto: texto, estado: estado,
             cerradosAfectados: cierresConEseEstado(ss, texto) };
  }
  return { ok: false, error: 'No encuentro ese estado en la hoja.' };
}

/**
 * Qué meses ya cerrados contienen pedidos con ese estado.
 *
 * Un cierre congelado no se rehace solo: las cifras que ya se reportaron
 * no pueden cambiar a espaldas de nadie. Pero sí hay que decir cuáles
 * quedaron hechos antes de saber esto, para que la dueña decida.
 */
function cierresConEseEstado(ss, texto) {
  const shC = ss.getSheetByName('Cierres');
  const shP = ss.getSheetByName('Pedidos');
  if (!shC || shC.getLastRow() < 2 || !shP || shP.getLastRow() < 2) return [];

  const cerrados = {};
  const dc = shC.getDataRange().getValues();
  const ec = dc[0].map(norm);
  for (let i = 1; i < dc.length; i++) {
    cerrados[String(dc[i][ec.indexOf('tienda')]).trim() + '|' +
             String(dc[i][ec.indexOf('mes')]).trim()] = true;
  }
  if (!Object.keys(cerrados).length) return [];

  const dp = shP.getDataRange().getValues();
  const ep = dp[0].map(norm);
  const cT = ep.indexOf('tienda'), cF = ep.indexOf('fecha'), cE = ep.indexOf('estado');
  const tocados = {};
  for (let i = 1; i < dp.length; i++) {
    if (norm(dp[i][cE]) !== texto) continue;
    const f = aISO(dp[i][cF], 'UTC');
    if (!f) continue;
    const k = String(dp[i][cT]).trim() + '|' + f.slice(0, 7);
    if (cerrados[k]) tocados[k] = (tocados[k] || 0) + 1;
  }
  return Object.keys(tocados).map(function (k) {
    return { tienda: k.split('|')[0], mes: k.split('|')[1], pedidos: tocados[k] };
  });
}

function apiCrear(s, p) {
  const entidad = String(p.entidad || '').trim();
  const campos = CREABLES[entidad];
  if (!campos) return { ok: false, error: 'No se pueden crear filas en ' + entidad + '.' };
  if (SOLO_DUENO.indexOf(entidad) !== -1 && s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña puede tocar ' + entidad + '.' };
  }
  if (!puede(s, 'escribir', entidad) && s.rol !== 'dueno') {
    return { ok: false, error: 'Tu rol no puede crear en ' + entidad + '.' };
  }
  if (ENTIDADES_DINERO.indexOf(entidad) !== -1 && s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña.' };
  }

  /**
   * Nadie entra al Equipo con un rol que el sistema no entiende, ni con
   * un correo que ya está adentro.
   *
   * Un rol mal escrito dejaría a esa persona sin acceso y sin explicación
   * el día que intente entrar; un correo repetido haría que dos filas
   * distintas peleen por la misma sesión.
   */
  if (entidad === 'Equipo') {
    const err = validarPersona(s, p.datos || {}, '');
    if (err) return { ok: false, error: err };
    // "asesora" y "propietaria" entran igual, pero en la hoja queda una
    // sola palabra por rol: así se puede filtrar y contar sin sorpresas.
    p.datos.rol = rolCanonico(p.datos.rol);
    if (!p.datos.estado) p.datos.estado = 'activo';
  }

  // Una categoría mal escrita no rompe nada hoy, pero mañana ese producto
  // no aparece en ningún filtro y nadie entiende por qué.
  if (entidad === 'Inventario') {
    if (!String((p.datos || {}).producto || '').trim()) {
      return { ok: false, error: 'La ficha necesita el nombre del producto: es ' +
               'con lo que se encuentra con los pedidos.' };
    }
    const err = validarFicha(p.datos || {});
    if (err) return { ok: false, error: err };
    if (p.datos.landing !== undefined) p.datos.landing = normalizarEnlace(p.datos.landing);
  }

  const datos = p.datos || {};
  const tienda = String(datos.tienda || p.tienda || '').trim();
  if (tienda && s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'Esa tienda no es tuya.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName(entidad);
  if (!sh) return { ok: false, error: 'Falta la hoja ' + entidad + '. Corre bootstrapTodo().' };

  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  const id = entidad.toLowerCase() + '-' + Utilities.getUuid().slice(0, 8);
  const hoy = new Date().toISOString().slice(0, 10);
  const fila = enc.map(function (col) {
    if (col === 'id') return id;
    if (col === 'activo') return 'si';
    if (col === 'actualizado_en') return ahoraISO();
    if (col === 'actualizado_por') return s.email;
    // Una ficha creada desde la app la escribió una persona, y el stock
    // con el que nace es el conteo del día.
    if (entidad === 'Inventario' && col === 'origen') return 'manual';
    if (entidad === 'Inventario' && col === 'ultimo_conteo') return hoy;
    if (campos.indexOf(col) !== -1) return datos[col] !== undefined ? datos[col] : '';
    return '';
  });
  sh.appendRow(fila);
  registrarMovimiento(s, entidad, id, 'creado', '', JSON.stringify(datos).slice(0, 200));
  return { ok: true, id: id };
}

/**
 * Marca una fila como inactiva. No la borra.
 *
 * Un gasto que se elimina de verdad se lleva consigo la explicación de
 * por qué el margen de marzo era ese. Desactivarlo lo saca de los
 * cálculos de aquí en adelante y deja el rastro.
 */
/**
 * Comprueba una persona antes de guardarla.
 *
 * @param {string} idActual  vacío al crear; el id de la fila al editar,
 *                           para no chocar consigo misma.
 */
function validarPersona(s, datos, idActual) {
  const correo = String(datos.correo || '').toLowerCase().trim();
  if (!correo || correo.indexOf('@') === -1) return 'Falta un correo válido.';
  if (!String(datos.nombre || '').trim()) return 'Falta el nombre.';

  const rol = rolCanonico(datos.rol);
  if (!rol) {
    return 'El rol "' + datos.rol + '" no se reconoce. Tiene que ser ' +
           'dueño, admin o gestora — se aceptan variantes como dueña, ' +
           'administradora o asesora.';
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return '';

  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cC = enc.indexOf('correo'), cId = enc.indexOf('id');
  for (let i = 1; i < filas.length; i++) {
    if (idActual && String(filas[i][cId]).trim() === idActual) continue;
    if (String(filas[i][cC] || '').toLowerCase().trim() === correo) {
      return 'Ya hay alguien en el equipo con el correo ' + correo + '.';
    }
  }
  return '';
}

/**
 * Cuántas dueñas activas quedarían si esta fila cambiara.
 *
 * Una cuenta sin dueña activa es una cuenta sin quién dé permisos: nadie
 * puede volver a entrar a arreglarlo, ni siquiera desde la hoja, porque
 * la app decide los roles leyendo justamente esa hoja.
 */
function duenosActivosSin(ss, idExcluido, rolNuevo) {
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return 0;
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cRol = enc.indexOf('rol'), cEst = enc.indexOf('estado');
  let n = 0;
  for (let i = 1; i < filas.length; i++) {
    const esta = String(filas[i][cId]).trim() === idExcluido;
    const rol = esta ? rolNuevo : rolCanonico(filas[i][cRol]);
    const estado = esta ? (rolNuevo ? 'activo' : 'inactivo') : norm(filas[i][cEst]);
    if (rol === 'dueno' && estado !== 'inactivo') n++;
  }
  return n;
}

function apiBorrar(s, p) {
  const entidad = String(p.entidad || '').trim();
  if (!CREABLES[entidad]) return { ok: false, error: 'No se puede borrar en ' + entidad + '.' };
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña.' };

  // Quitar a la última dueña deja la cuenta sin quién dé permisos
  if (entidad === 'Equipo') {
    const ss0 = SpreadsheetApp.openById(s.sheetId);
    if (duenosActivosSin(ss0, String(p.id || '').trim(), '') === 0) {
      return { ok: false, error: 'No puedes quitar a la única dueña: la cuenta ' +
               'se quedaría sin quién dé permisos. Nombra otra dueña primero.' };
    }
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName(entidad);
  if (!sh) return { ok: false, error: 'Falta la hoja ' + entidad + '.' };

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cId = enc.indexOf('id'), cAct = enc.indexOf('activo');
  const id = String(p.id || '').trim();
  const cEstado = enc.indexOf('estado');
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cId]).trim() !== id) continue;
    // Equipo marca 'estado: inactivo'; el resto, 'activo: no'
    if (entidad === 'Equipo' && cEstado !== -1) {
      sh.getRange(i + 1, cEstado + 1).setValue('inactivo');
      registrarMovimiento(s, entidad, id, 'estado', 'activo', 'inactivo');
    } else if (cAct !== -1) {
      sh.getRange(i + 1, cAct + 1).setValue('no');
      registrarMovimiento(s, entidad, id, 'activo', 'si', 'no');
    }
    return { ok: true };
  }
  return { ok: false, error: 'No encuentro esa fila.' };
}

function apiEscribir(s, p) {
  const entidad = String(p.entidad || '').trim();
  if (!puede(s, 'escribir', entidad)) {
    return { ok: false, error: 'Tu rol no puede escribir en ' + entidad + '.' };
  }

  const id = String(p.id || '').trim();
  const campos = p.campos || {};
  if (!id) return { ok: false, error: 'Falta el id de la fila.' };
  if (!Object.keys(campos).length) return { ok: false, error: 'No hay campos que escribir.' };

  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName(entidad);
  if (!sh) return { ok: false, error: 'No existe la hoja ' + entidad + '.' };

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cId = enc.indexOf('id');
  if (cId === -1) return { ok: false, error: entidad + ' no tiene columna id.' };

  let fila = -1;
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cId]).trim() === id) { fila = i; break; }
  }
  if (fila === -1) return { ok: false, error: 'No encuentro ' + entidad + ' con id ' + id + '.' };

  // La gestora solo escribe sobre lo suyo
  if (s.rol === 'gestora') {
    const cg = enc.indexOf('gestora_asignada') !== -1
      ? enc.indexOf('gestora_asignada') : enc.indexOf('gestora');
    if (cg !== -1 && norm(datos[fila][cg]) !== norm(s.nombre)) {
      return { ok: false, error: 'Ese caso no está asignado a ti.' };
    }
  }
  // Y solo dentro de sus tiendas
  const cT = enc.indexOf('tienda');
  if (cT !== -1 && s.tiendas.indexOf(String(datos[fila][cT]).trim()) === -1) {
    return { ok: false, error: 'Ese registro es de otra tienda.' };
  }

  /**
   * Editar a alguien del Equipo pasa por las mismas reglas que crearlo.
   *
   * Sin esto, bastaba con editar la fila propia y ponerse "admin" para
   * dejar la cuenta sin dueña, o escribir un rol inventado que dejaría a
   * esa persona fuera sin decir por qué.
   */
  if (entidad === 'Equipo') {
    if (s.rol !== 'dueno') {
      return { ok: false, error: 'Solo la dueña cambia el equipo.' };
    }
    const prop = {};
    ['nombre','correo','rol','tienda','estado','permisos'].forEach(function (k) {
      prop[k] = campos[k] !== undefined ? campos[k] : datos[fila][enc.indexOf(k)];
    });
    const err = validarPersona(s, prop, id);
    if (err) return { ok: false, error: err };

    const rolNuevo = rolCanonico(prop.rol);
    const quedaActivo = norm(prop.estado) !== 'inactivo';
    if (duenosActivosSin(ss, id, quedaActivo ? rolNuevo : '') === 0) {
      return { ok: false, error: 'Ese cambio dejaría la cuenta sin ninguna ' +
               'dueña activa, y nadie podría volver a dar permisos. ' +
               'Nombra otra dueña primero.' };
    }
    // El rol se guarda en su forma canónica: "dueña" y "propietaria"
    // entran igual, pero en la hoja queda una sola palabra.
    if (campos.rol !== undefined) campos.rol = rolNuevo;
  }

  if (entidad === 'Inventario') {
    const err = validarFicha(campos);
    if (err) return { ok: false, error: err };
    // Un dominio suelto se guarda completo, para que el enlace funcione
    if (campos.landing !== undefined) campos.landing = normalizarEnlace(campos.landing);
  }

  const escritos = [], rechazados = [];
  Object.keys(campos).forEach(function (k) {
    const col = enc.indexOf(norm(k));
    if (col === -1) {
      // Una columna que falta casi siempre es una migración pendiente, no
      // un campo inventado. Decirlo ahorra media hora de buscar dónde está
      // el error.
      rechazados.push(k + ' (esa columna todavía no existe en la hoja ' +
                      entidad + ' — corre bootstrapTodo())');
      return;
    }
    if (COLUMNAS_IMPORTADAS.indexOf(norm(k)) !== -1) {
      rechazados.push(k + ' (viene de la plataforma, es de solo lectura)');
      return;
    }
    const antes = datos[fila][col];
    const ahora = campos[k];
    if (String(antes) === String(ahora)) return;

    sh.getRange(fila + 1, col + 1).setValue(ahora);
    registrarMovimiento(s, entidad, id, norm(k), antes, ahora);
    escritos.push(k);
  });

  /**
   * Tocar el stock a mano es hacer un conteo, y queda fechado.
   *
   * Sin esto, una ficha que alguien corrigió hace tres meses y otra que
   * se contó esta mañana se ven igual, y la alarma de stock no sabe a
   * cuál de las dos creerle.
   */
  if (entidad === 'Inventario' && escritos.indexOf('stock') !== -1) {
    const hoy = new Date().toISOString().slice(0, 10);
    [['ultimo_conteo', hoy], ['origen', 'manual']].forEach(function (par) {
      const col = enc.indexOf(par[0]);
      if (col !== -1) sh.getRange(fila + 1, col + 1).setValue(par[1]);
    });
  }

  // Rastro de frescura: varias alarmas dependen de esto
  ['actualizado_en', 'actualizado_por'].forEach(function (k, i) {
    const col = enc.indexOf(k);
    if (col !== -1) sh.getRange(fila + 1, col + 1).setValue(i ? s.email : ahoraISO());
  });

  return { ok: true, escritos: escritos, rechazados: rechazados };
}

/** Sin este registro no hay auditoría ni vista sombra. */
function registrarMovimiento(s, entidad, entidadId, campo, antes, ahora) {
  try {
    const sh = SpreadsheetApp.openById(s.sheetId).getSheetByName('Movimientos');
    // Si estaba mirando con otra vista, queda dicho: sigue siendo ella,
    // pero conviene saber desde dónde lo hizo.
    const quien = s.vistaComo
      ? s.email + ' (viendo como ' + s.vistaComo + ')'
      : s.email;
    if (sh) sh.appendRow([ahoraISO(), quien, entidad, entidadId, campo, antes, ahora]);
  } catch (err) {
    // Que falle la bitácora no puede tumbar la operación
    Logger.log('No se pudo registrar el movimiento: ' + err.message);
  }
}

// ─── RESUMEN Y CIERRE ────────────────────────────────────────

/** Los KPIs de la pantalla Hoy, calculados en el servidor. */
function apiResumen(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const mes = String(p.mes || Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM'));
  const d = agregarMes(ss, tienda, mes, s);
  d.recaudo7 = recaudoUltimosDias(ss, tienda, 7);
  return { ok: true, tienda: tienda, mes: mes, datos: d };
}

/**
 * Lo que entró cada uno de los últimos días.
 *
 * Va aparte de agregarMes porque una semana no cabe dentro de un mes: el
 * 2 de septiembre, cinco de los siete días son de agosto. Calcularlo con
 * el agregado del mes habría dibujado media semana en cero.
 *
 * Recaudo es lo entregado, y se cuenta el día que se entregó, no el día
 * que se pidió. Un pedido del lunes que llega el jueves es plata del
 * jueves: ponerla en el lunes haría que el mejor día del gráfico fuera
 * siempre el día en que más se vendió, no aquel en que más entró.
 *
 * Cuando la transportadora no da fecha de entrega, ese pedido no se
 * reparte por ningún lado: se cuenta aparte y la pantalla lo dice. Una
 * plata que no se sabe qué día entró no puede inventarse un día.
 */
function recaudoUltimosDias(ss, tienda, dias) {
  const tz = zonaHorariaDe(ss, tienda) || 'UTC';
  const hoy = new Date();
  const serie = [];
  const indice = {};
  for (let i = dias - 1; i >= 0; i--) {
    const d = new Date(hoy.getTime() - i * 86400000);
    const iso = Utilities.formatDate(d, tz, 'yyyy-MM-dd');
    indice[iso] = serie.length;
    serie.push({ fecha: iso, total: 0, entregas: 0 });
  }

  const out = { serie: serie, sinFecha: 0, pedidosSinFecha: 0, moneda: monedaDeTienda(ss, tienda) };
  const sh = ss.getSheetByName('Pedidos');
  if (!sh || sh.getLastRow() < 2) return out;

  const datos = sh.getDataRange().getValues();
  const e = datos[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const desde = serie[0].fecha;

  for (let i = 1; i < datos.length; i++) {
    const f = datos[i];
    if (String(f[c('tienda')]).trim() !== tienda) continue;
    if (norm(f[c('estado_nova')] || f[c('estado_canonico')]) !== 'entregado') continue;

    const valor = num(f[c('valor')]);
    const entrega = c('fecha_entrega') !== -1 ? aISO(f[c('fecha_entrega')], tz) : '';
    if (!entrega) {
      // Solo cuenta como "sin fecha" si el pedido es reciente; uno de
      // marzo sin fecha de entrega no es un hueco de esta semana.
      const pedido = aISO(f[c('fecha')], tz);
      if (pedido && pedido >= desde) { out.sinFecha += valor; out.pedidosSinFecha++; }
      continue;
    }
    if (indice[entrega] === undefined) continue;
    const dia = serie[indice[entrega]];
    dia.total += valor;
    dia.entregas++;
  }
  return out;
}

/**
 * Quién estuvo trabajando hoy y desde cuándo.
 *
 * No hay reloj de entrada: lo que hay es la bitácora. Cada vez que
 * alguien entra, escribe una nota, cambia un estado o sube un archivo,
 * queda una fila en Movimientos con su correo y la hora. La jornada de
 * una persona es su primer movimiento del día, el último, y cuántos hizo.
 *
 * Eso no es lo mismo que horas trabajadas y la pantalla no lo llama así.
 * Alguien puede estar llamando dos horas sin tocar Nova; lo que se sabe
 * es cuándo tocó Nova, y eso es lo que se dice.
 */
function jornadaDeHoy(ss, tienda) {
  const tz = zonaHorariaDe(ss, tienda) || 'UTC';
  const hoy = Utilities.formatDate(new Date(), tz, 'yyyy-MM-dd');
  const porCorreo = {};

  const sh = ss.getSheetByName('Movimientos');
  if (sh && sh.getLastRow() > 1) {
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const cF = e.indexOf('fecha'), cU = e.indexOf('usuario');
    // De atrás hacia adelante: lo de hoy está al final, y así no se
    // recorren doce mil filas de meses pasados para nada.
    for (let i = d.length - 1; i > 0; i--) {
      const sello = String(d[i][cF] || '');
      const dia = sello.slice(0, 10);
      if (dia < hoy) break;
      if (dia !== hoy) continue;
      // "correo (viendo como gestora)" sigue siendo la misma persona
      const quien = String(d[i][cU] || '').split(' (')[0].trim().toLowerCase();
      if (!quien) continue;
      const hora = sello.slice(11, 16);
      if (!porCorreo[quien]) {
        porCorreo[quien] = { correo: quien, primero: hora, ultimo: hora, acciones: 0 };
      }
      const x = porCorreo[quien];
      x.acciones++;
      if (hora && hora < x.primero) x.primero = hora;
      if (hora && hora > x.ultimo)  x.ultimo = hora;
    }
  }

  /**
   * Cuánto hace de eso se calcula aquí, no en el navegador.
   *
   * La hora que se guarda es la de la tienda —Guatemala, Ecuador— y quien
   * mira puede estar en Colombia. Restar una contra el reloj del navegador
   * daba una hora de diferencia en Nutrea GT: Katherin salía "activa hace
   * un momento" cuando llevaba dos horas sin tocar nada, o al revés.
   */
  const ahoraMin = Number(Utilities.formatDate(new Date(), tz, 'HH')) * 60 +
                   Number(Utilities.formatDate(new Date(), tz, 'mm'));
  Object.keys(porCorreo).forEach(function (k) {
    const x = porCorreo[k];
    const p = String(x.ultimo || '').split(':');
    const ult = Number(p[0]) * 60 + Number(p[1] || 0);
    x.haceMin = isNaN(ult) ? null : Math.max(0, ahoraMin - ult);
  });
  return porCorreo;
}

/** El cierre de mes: este mes contra el anterior. */
function apiCierre(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const mes = String(p.mes || Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM'));
  const prev = mesAnterior(mes);

  // Un mes cerrado devuelve lo que se congeló, no un recálculo: si se
  // recalculara, las cifras que ya reportaste cambiarían solas.
  const congelado = cierreGuardado(ss, tienda, mes);
  const r = {
    ok: true, tienda: tienda, mes: mes,
    moneda: monedaDeTienda(ss, tienda),
    actual: congelado ? congelado.datos : agregarMes(ss, tienda, mes, s),
    anterior: agregarMes(ss, tienda, prev, s),
    cerrado: !!congelado,
    cerrado_en: congelado ? congelado.cerrado_en : '',
  };
  if (!congelado) {
    r.provisional = r.actual.pendientes > 0;
    r.pendientes = r.actual.pendientes;
  }

  // La moneda en la que la dueña piensa, y a cuánto estaba el último día
  // del mes. Sin esto el informe solo habla en la moneda de la tienda.
  r.monedaReporte = monedaReporte(ss);
  if (r.monedaReporte && r.monedaReporte !== r.moneda) {
    const t = buscarTasa(ss, ultimoDiaDelMes(mes), r.moneda, r.monedaReporte);
    r.tasa = t && t.tasa ? t.tasa : null;
    r.tasaFecha = t ? t.fecha_usada : '';
    r.tasaExacta = t ? !!t.exacta : false;
  }

  // El efecto cambiario es solo de la dueña: es información de dinero
  if (s.rol === 'dueno' && monedaDeTienda(ss, tienda) !== monedaReporte(ss)) {
    try {
      r.cambiario = efectoCambiarioSS(ss, tienda, prev, mes);
    } catch (err) {
      r.cambiario = { aplica: false, mensaje: err.message };
    }
  }
  return r;
}

/**
 * Agrega un mes de una tienda. Una sola pasada por Pedidos y otra por
 * Novedades: con miles de filas, recorrerlas por cada KPI es lo que hace
 * que la pantalla tarde.
 */
function agregarMes(ss, tienda, mes, s) {
  const out = {
    pedidos: 0, despachados: 0, entregados: 0, devueltos: 0, cancelados: 0,
    pendientes: 0,
    ventas: 0, costoProducto: 0, costoEnvio: 0, costoDevolucion: 0,
    valorAbierto: 0,
    // Pedidos cuyo estado Nova no reconoce. No entran en ninguna otra
    // cuenta: ni entregados, ni devueltos, ni despachados, ni costos.
    sinClasificar: 0, valorSinClasificar: 0, estadosDesconocidos: {},
    // Qué costo trae cada estado, contado o no. Es lo que permite
    // reconciliar el cierre de Nova contra el que el cliente hizo aparte.
    costosPorEstado: {
      producto: { entregado: 0, devolucion: 0, cancelado: 0, pendiente: 0 },
      envio:    { entregado: 0, devolucion: 0, cancelado: 0, pendiente: 0 },
    },
    novedades: 0, sinMover: 0,
    grupos: {}, transportadoras: {}, productos: {},
  };

  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const datos = shP.getDataRange().getValues();
    const e = datos[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    const hoy = new Date();

    for (let i = 1; i < datos.length; i++) {
      const f = datos[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha || fecha.slice(0, 7) !== mes) continue;

      out.pedidos++;
      // estado_nova gana sobre el importado: es lo que el equipo corrigió
      const est = norm(f[c('estado_nova')] || f[c('estado_canonico')] || f[c('estado')]);

      /**
       * Un estado que Nova no entiende no entra en ninguna cuenta.
       *
       * La tentación era tratarlo como pendiente, pero eso es afirmar que
       * el pedido sigue vivo —y bien puede estar entregado desde hace un
       * mes—. Tampoco se le cobra flete: decir que salió de bodega es
       * otra afirmación que nadie puede sostener.
       *
       * Se cuenta aparte, con su plata, y la pantalla lo pregunta. Un
       * hueco que se ve vale mucho más que un número que parece completo.
       */
      if (est === ESTADOS.SIN_CLASIFICAR) {
        out.sinClasificar++;
        out.valorSinClasificar += num(f[c('valor')]);
        const crudo = String(f[c('estado')] || '').trim() || '(vacío)';
        out.estadosDesconocidos[crudo] = (out.estadosDesconocidos[crudo] || 0) + 1;
        continue;
      }

      if (est === 'entregado')   { out.entregados++; out.ventas += num(f[c('valor')]); }
      if (est === 'devolucion')  out.devueltos++;
      if (est === 'cancelado')   out.cancelados++;
      if (['cancelado','pendiente'].indexOf(est) === -1) out.despachados++;

      /**
       * Un costo se cuenta cuando se incurrió, no cuando aparece en la fila.
       *
       * Dropi trae costo_producto y costo_envio en TODAS las filas, también
       * en las canceladas y en las que aún no salen de bodega. Sumarlas
       * todas era cargarle al mes el costo de mercancía que nunca se
       * despachó y fletes que nadie pagó. La pantalla del cierre ya decía
       * "cancelados: sin flete" mientras la suma sí se los cobraba.
       *
       *   producto → solo lo entregado: es la mercancía que se fue y se pagó
       *   flete    → lo despachado: salió del almacén, la transportadora cobra
       *   cancelado y pendiente → nada: no se despachó
       *
       * Se guarda además el desglose por estado, para poder mostrar en el
       * cierre qué se contó y qué se dejó fuera. Una cifra que no se puede
       * reconciliar contra la hoja propia del cliente no sirve de nada.
       */
      const cProd = num(f[c('costo_producto')]);
      const cEnv  = num(f[c('costo_envio')]);
      const grupoCosto = ['entregado', 'devolucion', 'cancelado'].indexOf(est) !== -1
        ? est : 'pendiente';
      out.costosPorEstado.producto[grupoCosto] += cProd;
      out.costosPorEstado.envio[grupoCosto]    += cEnv;

      if (est === 'entregado') out.costoProducto += cProd;
      if (['cancelado', 'pendiente'].indexOf(est) === -1) out.costoEnvio += cEnv;

      /**
       * El flete de una devolución se paga igual, y a veces doble. Va
       * aparte del flete de las entregas porque son dos cosas distintas:
       * uno es costo de vender, el otro es costo de no haber vendido.
       */
      if (est === 'devolucion') out.costoDevolucion += cEnv;

      // Por producto, para ver cuál se sostiene y cuál no
      const prod = String(f[c('producto')] || 'Sin producto').trim();
      if (!out.productos[prod]) {
        out.productos[prod] = { pedidos: 0, entregados: 0, ventas: 0 };
      }
      out.productos[prod].pedidos++;
      if (est === 'entregado') {
        out.productos[prod].entregados++;
        out.productos[prod].ventas += num(f[c('valor')]);
      }

      const t = String(f[c('transportadora')] || '').trim();
      if (t) {
        if (!out.transportadoras[t]) out.transportadoras[t] = { n: 0, entregados: 0 };
        out.transportadoras[t].n++;
        if (est === 'entregado') out.transportadoras[t].entregados++;
      }

      // Un pedido sin estado terminal es un desenlace que todavía no se
      // conoce: mientras haya alguno, el mes es provisional.
      if (['entregado','devolucion','cancelado'].indexOf(est) === -1) {
        out.pendientes++;
        out.valorAbierto += num(f[c('valor')]);
        const ult = aISO(f[c('ultimo_movimiento')] || f[c('actualizado_en')], 'UTC') || fecha;
        const dias = (hoy - new Date(ult + 'T00:00:00Z')) / 86400000;
        if (dias > 3) out.sinMover++;
      }
    }
  }

  const shN = ss.getSheetByName('Novedades');
  if (shN && shN.getLastRow() > 1) {
    const datos = shN.getDataRange().getValues();
    const e = datos[0].map(norm);
    const cF = e.indexOf('fecha'), cG = e.indexOf('grupo'), cM = e.indexOf('motivo');
    for (let i = 1; i < datos.length; i++) {
      const fecha = aISO(datos[i][cF], 'UTC');
      if (!fecha || fecha.slice(0, 7) !== mes) continue;
      out.novedades++;
      const g = String(datos[i][cG] || '').trim() ||
                grupoNovedad(datos[i][cM]) || 'otro';
      out.grupos[g] = (out.grupos[g] || 0) + 1;
    }
  }

  /**
   * La tasa de entrega se mide sobre lo RESUELTO, no sobre lo despachado.
   *
   * Un pedido que todavía está en ruta no ha fallado ni ha acertado: no
   * sabe. Meterlo en el denominador hunde la tasa del mes en curso y la
   * hace ver peor de lo que es, justo cuando más se mira.
   */
  out.resueltos = out.entregados + out.devueltos;
  out.efectividad = out.resueltos ? out.entregados / out.resueltos * 100 : 0;
  out.efectividadDespacho = out.despachados ? out.entregados / out.despachados * 100 : 0;
  out.tasaDevolucion = out.despachados ? out.devueltos / out.despachados * 100 : 0;
  out.ticket = out.entregados ? out.ventas / out.entregados : 0;

  // La pauta solo se agrega para la dueña
  if (s.rol === 'dueno') {
    const shPa = ss.getSheetByName('Pauta');
    out.gasto = 0; out.campanas = {}; out.gastoPorDia = {};
    out.gastoSinConvertir = 0; out.monedasSinTasa = {};
    const monTienda = monedaDeTienda(ss, tienda);
    if (shPa && shPa.getLastRow() > 1) {
      const datos = shPa.getDataRange().getValues();
      const e = datos[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < datos.length; i++) {
        const f = datos[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        const fecha = aISO(f[c('fecha')], 'UTC');
        if (!fecha || fecha.slice(0, 7) !== mes) continue;
        /**
         * El gasto SOLO cuenta si está en la moneda de la tienda.
         *
         * Meta le cobra a Nutrea en pesos y la tienda factura en dólares.
         * Cuando no había tasa de cambio, gasto_normalizado quedaba vacío
         * y esto caía al gasto crudo: 1.386.315 pesos se sumaban como si
         * fueran dólares contra ventas de 4.000. De ahí salían un CPA de
         * 277.263 y un margen de −181.817%.
         *
         * Restar pesos a dólares no es un error de redondeo: es una cifra
         * inventada con pinta de cierta. Así que si no hay conversión, no
         * se suma — se cuenta aparte y la pantalla lo dice.
         */
        const mon = String(f[c('moneda_gasto')] || '').toUpperCase();
        const crudo = num(f[c('gasto')]);
        let g = 0;
        if (!mon || mon === monTienda) {
          g = crudo;                       // ya viene en la moneda de la tienda
        } else {
          // Con la tasa del día del gasto, no la de hoy: una campaña de
          // agosto se pagó al dólar de agosto.
          const t = buscarTasa(ss, fecha, mon, monTienda);
          if (t && t.tasa) {
            g = crudo * t.tasa;
          } else {
            out.gastoSinConvertir += crudo;
            out.monedasSinTasa[mon] = (out.monedasSinTasa[mon] || 0) + crudo;
          }
        }
        out.gasto += g;
        // Conjunto si lo hay: es donde de verdad se decide el presupuesto
        const nom = String(f[c('conjunto')] || f[c('campana')] || 'Sin nombre').trim();
        if (!out.campanas[nom]) out.campanas[nom] = { gasto: 0, resultados: 0, sinTasa: 0 };
        out.campanas[nom].gasto += g;
        if (!g && crudo) out.campanas[nom].sinTasa += crudo;
        out.campanas[nom].resultados += num(f[c('resultados')]);

        /**
         * Gasto por día, solo si el reporte viene por día.
         *
         * El informe de conjuntos de Meta trae una fila por todo el
         * periodo. Repartirla entre los días dibujaría una curva que no
         * existe, y una curva inventada se lee como si fuera real.
         */
        const fin = aISO(f[c('fecha_fin')], 'UTC');
        if (!fin || fin === fecha) {
          out.gastoPorDia[fecha] = (out.gastoPorDia[fecha] || 0) + g;
        }
      }
    }
    /**
     * Los gastos fijos del mes.
     *
     * `mes` vacío es un gasto que se repite todos los meses; con un mes
     * concreto, es de una sola vez. Un gasto desactivado deja de contar
     * de aquí en adelante, pero su fila se queda: es lo que explica por
     * qué el margen de marzo era el que era.
     */
    out.fijos = 0; out.detalleFijos = [];
    const shG = ss.getSheetByName('Gastos');
    if (shG && shG.getLastRow() > 1) {
      const datos = shG.getDataRange().getValues();
      const e = datos[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < datos.length; i++) {
        const f = datos[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        if (norm(f[c('activo')]) === 'no') continue;
        const m = String(f[c('mes')] || '').trim();
        if (m && m !== mes) continue;   // gasto de otro mes
        const v = num(f[c('valor')]);
        out.fijos += v;
        out.detalleFijos.push({ nombre: String(f[c('nombre')] || ''), valor: v,
                                tipo: String(f[c('tipo')] || 'fijo'),
                                recurrente: !m, id: String(f[c('id')] || '') });
      }
    }

    /**
     * La cartera corrige el costo de devolución.
     *
     * El export de órdenes trae un flete de lista; la cartera trae lo que
     * la plataforma cobró de verdad. En agosto de Nutrea EC eso son 63
     * cobros por 315,55 —promedio 5,01— contra los 3,50 del export: 95
     * dólares que ningún cierre estaba contando.
     *
     * Se usa el promedio real por devolución sobre las devoluciones de la
     * cohorte, no el total del mes de cartera: los cobros de agosto
     * incluyen devoluciones de pedidos de julio. El promedio sí es
     * representativo; el total sería de otro conjunto de pedidos.
     */
    out.cartera = carteraDelMes(ss, tienda, mes);
    out.costoDevolucionEstimado = out.costoDevolucion;
    if (out.cartera.hay && out.cartera.devolucionPromedio && out.devueltos) {
      out.costoDevolucion = out.cartera.devolucionPromedio * out.devueltos;
      out.costoDevolucionFuente = 'cartera';
    } else {
      out.costoDevolucionFuente = 'export';
    }

    out.cpa  = out.entregados ? out.gasto / out.entregados : 0;
    out.roas = out.gasto ? out.ventas / out.gasto : 0;
    /**
     * El cobro de devolución entra en el margen SOLO si viene de la cartera.
     *
     * Una devolución cuesta dos veces: el flete de ida, que ya está en
     * costoEnvio porque el pedido sí salió de bodega, y el cobro de
     * retorno que la plataforma pasa aparte. En el historial de Nutrea son
     * dos líneas distintas: "SALIDA POR COBRO DE FLETE INICIAL" (~7,10) y
     * "SALIDA DE COBRO DE DEVOLUCIÓN" (~5,01).
     *
     * Sin cartera, lo único que Nova tiene es el flete del export, y
     * costoDevolucion es una copia de ese mismo número: restarlo sería
     * cobrar el flete de ida dos veces. Así que sin cartera se muestra
     * aparte y no se resta, y la pantalla dice que falta ese costo.
     *
     * Margen: antes de los gastos fijos. Utilidad: lo que queda de verdad.
     */
    const cobroRetorno = out.costoDevolucionFuente === 'cartera'
      ? out.costoDevolucion : 0;

    /**
     * Las dos comisiones que nadie factura.
     *
     * La de retiro sale de la cartera: es un porcentaje de lo que sacaste
     * de la billetera ese mes. La internacional la cobra el banco por
     * pagar en moneda extranjera, y solo a algunos proveedores — a Meta
     * sí, a Shopify y Claude no—, así que se aplica al gasto de las
     * plataformas que la dueña haya listado, no a todo.
     *
     * Las dos van con los gastos fijos y no dentro del margen: el margen
     * responde si el producto deja plata, y estas comisiones no dependen
     * del producto sino de cómo se mueve el dinero.
     */
    const ajM = ajustes(ss, tienda);
    out.comisionRetiro = (out.cartera && out.cartera.costoRetiros) || 0;
    const pctIntl = Number(ajM.comision_intl_pct) || 0;
    const plataformas = String(ajM.comision_intl_a || '').split(/[,;]/)
      .map(function (x) { return norm(x); }).filter(String);
    let baseIntl = 0;
    // El gasto ya convertido de las plataformas a las que sí se las cobran
    if (pctIntl && plataformas.indexOf('meta') !== -1) baseIntl = out.gasto;
    out.comisionIntl = baseIntl * (pctIntl / 100);
    out.comisiones = out.comisionRetiro + out.comisionIntl;
    out.margen = out.ventas - out.gasto - out.costoProducto - out.costoEnvio
                 - cobroRetorno;
    out.utilidad = out.margen - out.fijos - out.comisiones;
  }
  return out;
}

function num(v) { const n = aNumero(v); return n === '' ? 0 : n; }

/** efectoCambiario() pero recibiendo el spreadsheet ya abierto. */
function efectoCambiarioSS(ss, tienda, mesA, mesB) {
  const destino = monedaReporte(ss);
  const origen = monedaDeTienda(ss, tienda);
  if (origen === destino) return { aplica: false, mensaje: 'Misma moneda.' };

  const R0 = ventasDelMes(ss, tienda, mesA), R1 = ventasDelMes(ss, tienda, mesB);
  const T0 = tasaPromedioMes(ss, mesA, origen, destino);
  const T1 = tasaPromedioMes(ss, mesB, origen, destino);
  if (T0 === null || T1 === null) {
    return { aplica: false, mensaje: 'Faltan tasas. Corre actualizarTasas().' };
  }
  return {
    aplica: true, origen: origen, destino: destino,
    ventas: { a: R0, b: R1 }, tasas: { a: T0, b: T1 },
    total: R1 * T1 - R0 * T0,
    porOperacion: (R1 - R0) * T0,
    porCambio: R1 * (T1 - T0),
  };
}



// ─── CIERRE DE MES ───────────────────────────────────────────

/** Devuelve las cifras congeladas de un mes ya cerrado, o null. */
function cierreGuardado(ss, tienda, mes) {
  const sh = ss.getSheetByName('Cierres');
  if (!sh || sh.getLastRow() < 2) return null;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')]).trim() !== tienda) continue;
    if (String(d[i][c('mes')]).trim() !== mes) continue;
    if (norm(d[i][c('estado')]) !== 'cerrado') continue;

    /**
     * El cierre completo si se guardó; si no, lo que quepa en las columnas.
     *
     * Los meses cerrados antes de que existiera `snapshot` no tienen dónde
     * guardar el costo de mercancía ni el flete, y salían como cero: el
     * informe mostraba ventas menos pauta y una utilidad que nunca fue.
     * Ahora esos meses se marcan como incompletos y el informe lo dice, en
     * vez de enseñar un número de más.
     */
    const crudo = c('snapshot') === -1 ? '' : String(d[i][c('snapshot')] || '');
    if (crudo) {
      try {
        const snap = JSON.parse(crudo);
        snap.congelado = true;
        return { cerrado_en: d[i][c('cerrado_en')], datos: snap };
      } catch (err) { /* snapshot ilegible: se cae al resumen de columnas */ }
    }
    return {
      cerrado_en: d[i][c('cerrado_en')],
      datos: {
        pedidos: num(d[i][c('pedidos')]), entregados: num(d[i][c('entregados')]),
        devueltos: num(d[i][c('devueltos')]), ventas: num(d[i][c('ventas')]),
        gasto: num(d[i][c('gasto')]), margen: num(d[i][c('margen')]),
        efectividad: num(d[i][c('efectividad')]),
        pendientes: num(d[i][c('pendientes_al_cierre')]),
        despachados: 0, cancelados: 0, novedades: 0, sinMover: 0,
        grupos: {}, transportadoras: {},
        congelado: true, incompleto: true,
      },
    };
  }
  return null;
}

/**
 * Cierra un mes: congela sus cifras.
 *
 * Solo la dueña. Y avisa si quedan pedidos sin resolver, porque cerrar
 * con pendientes deja fuera ventas que todavía pueden entrar — pero no
 * lo prohíbe: a veces hay que cerrar contra una fecha aunque falten dos
 * guías perdidas.
 */
/**
 * El mes anterior a uno dado, en formato AAAA-MM.
 *
 * Existía solo en la pantalla. El servidor la llamaba desde apiCierre sin
 * tenerla, así que TODA petición del cierre fallaba con "mesAnterior is
 * not defined" — la pantalla lo interpretaba como "no hay datos" y caía
 * al ejemplo. Un error que se disfrazaba de falta de información.
 */
function mesAnterior(mes) {
  const a = parseInt(String(mes).slice(0, 4), 10);
  const m = parseInt(String(mes).slice(5, 7), 10);
  if (!a || !m) return String(mes);
  const d = new Date(Date.UTC(a, m - 2, 1));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM');
}

/** El último día real del mes: 28, 30 o 31 según toque. */
function ultimoDiaDelMes(mes) {
  const a = parseInt(mes.slice(0, 4), 10), m = parseInt(mes.slice(5, 7), 10);
  const d = new Date(Date.UTC(a, m, 0));
  return Utilities.formatDate(d, 'UTC', 'yyyy-MM-dd');
}

function apiCerrarMes(s, p) {
  if (s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña cierra un mes.' };
  }
  const tienda = String(p.tienda || '').trim();
  const mes = String(p.mes || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  if (!/^\d{4}-\d{2}$/.test(mes)) {
    return { ok: false, error: 'El mes va como AAAA-MM.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  if (cierreGuardado(ss, tienda, mes)) {
    return { ok: false, error: 'Ese mes ya está cerrado. Para rehacerlo, ' +
             'cambia su estado a "abierto" en la hoja Cierres.' };
  }

  const d = agregarMes(ss, tienda, mes, s);
  if (d.pendientes > 0 && !p.forzar) {
    return {
      ok: false, requiere_confirmacion: true, pendientes: d.pendientes,
      error: 'Quedan ' + d.pendientes + ' pedidos de ' + mes + ' sin resolver. ' +
        'Si cierras ahora, sus ventas no entran en este mes y las cifras ' +
        'quedan congeladas así. Normalmente se cierra cuando ya no hay ' +
        'nada pendiente por entregar.',
    };
  }

  const sh = ss.getSheetByName('Cierres');
  if (!sh) return { ok: false, error: 'Falta la hoja Cierres. Corre bootstrapTodo().' };

  /**
   * Se guarda el cierre entero, por nombre de columna.
   *
   * Antes se escribía con appendRow y una lista posicional: si alguien
   * agregaba una columna a Cierres, todo se corría un puesto. Y el
   * snapshot es lo que permite que un mes cerrado vuelva con sus costos
   * —mercancía, flete, fijos, el desglose por estado— en vez de ceros.
   */
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  const valores = {
    tienda: tienda, mes: mes, estado: 'cerrado', cerrado_en: ahoraISO(),
    cerrado_por: s.email, pendientes_al_cierre: d.pendientes,
    pedidos: d.pedidos, entregados: d.entregados, devueltos: d.devueltos,
    ventas: d.ventas, gasto: d.gasto || 0, margen: d.margen || 0,
    efectividad: d.efectividad, nota: p.nota || '',
    snapshot: JSON.stringify(d),
  };
  sh.appendRow(enc.map(function (col) {
    return valores[col] !== undefined ? valores[col] : '';
  }));
  registrarMovimiento(s, 'Cierres', tienda + '/' + mes, 'estado', 'abierto', 'cerrado');
  return { ok: true, mes: mes, tienda: tienda, datos: d, cerrado_en: ahoraISO() };
}

// ─── IMPORTAR DESDE LA APP ───────────────────────────────────
/**
 * Recibe un archivo subido desde la pantalla y lo importa.
 *
 * Esta es la vía real del producto. Pegar a mano en las pestañas
 * _Import_* era andamiaje de pruebas, y traía un problema que aquí
 * no existe: un archivo pegado no dice de qué tienda es, así que había
 * que declararlo aparte y podía quedar mal. Cuando la persona sube el
 * archivo desde la app, la tienda es la que tiene abierta. No hay nada
 * que adivinar.
 *
 * El crudo igual queda archivado en su pestaña _Import_*, porque es lo
 * que permite rehacer una importación cuando un mapeo se corrige.
 */
/**
 * Recibe un pedazo de archivo y lo guarda en el caché.
 *
 * Un archivo entero en una sola petición no llega: Google rechaza la
 * petición antes de que el script se entere —no aparece ni en el
 * registro de ejecuciones— y el navegador solo ve un 404 que no explica
 * nada. Por eso el archivo viaja partido, y se arma aquí.
 *
 * Los pedazos son de 30 KB. No porque el caché no admita más —admite
 * 100 KB— sino porque no sabemos dónde corta Google exactamente: 72 KB
 * en una sola petición no pasa, 80 bytes sí. 30 KB queda lejos de la duda.
 *
 * Viven 10 minutos: lo suficiente para terminar de subir, y no tanto
 * como para que un archivo abandonado se quede ocupando espacio.
 */
function apiTrozo(s, p) {
  const clave = claveSubida(s, p.clave);
  const i = parseInt(p.indice, 10);
  if (!clave || isNaN(i)) return { ok: false, error: 'Trozo mal identificado.' };
  const t = String(p.trozo || '');
  if (t.length > 100000) return { ok: false, error: 'Trozo demasiado grande.' };

  CacheService.getScriptCache().put(clave + '_' + i, t, 600);
  return { ok: true, indice: i, bytes: t.length };
}

/** La clave lleva el correo: nadie puede armar la subida de otra persona. */
function claveSubida(s, clave) {
  const c = String(clave || '').replace(/[^A-Za-z0-9]/g, '').slice(0, 40);
  if (!c) return '';
  return 'sub_' + String(s.email).replace(/[^a-z0-9]/g, '') + '_' + c;
}

/** Junta los pedazos. Si falta alguno, se dice cuál: no se importa a medias. */
function armarSubida(s, clave, trozos) {
  const base = claveSubida(s, clave);
  const cache = CacheService.getScriptCache();
  const partes = [];
  for (let i = 0; i < trozos; i++) {
    const t = cache.get(base + '_' + i);
    if (t === null) {
      throw new Error('Se perdió la parte ' + (i + 1) + ' de ' + trozos +
        ' mientras subía. Vuelve a intentarlo.');
    }
    partes.push(t);
  }
  for (let i = 0; i < trozos; i++) cache.remove(base + '_' + i);
  return partes.join('');
}

function apiImportarArchivo(s, p) {
  const fuente = String(p.fuente || '').trim();
  const tienda = String(p.tienda || '').trim();
  const nombre = String(p.nombre || 'archivo').trim();

  let b64;
  if (p.clave && p.trozos) {
    try { b64 = armarSubida(s, p.clave, parseInt(p.trozos, 10)); }
    catch (err) { return { ok: false, error: err.message }; }
  } else {
    b64 = String(p.contenido || '');
  }

  if (!FUENTES[fuente]) return { ok: false, error: 'Fuente desconocida: ' + fuente };
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a la tienda ' + tienda + '.' };
  }
  // Quién puede subir qué lo decide la dueña, no el rol a secas ni la
  // pantalla: esconder un botón no impide llamar a la API directamente.
  const necesita = PERMISO_DE_FUENTE[fuente];
  const tiene = s.permisos || [];
  if (necesita && tiene.indexOf(necesita) === -1) {
    return { ok: false, error:
      (necesita === 'subir_pauta' || necesita === 'subir_cartera')
        ? 'No tienes permiso para subir ' +
          (necesita === 'subir_pauta' ? 'pauta' : 'la cartera') +
          '. La dueña lo activa en Permisos.'
        : 'No tienes permiso para subir este tipo de archivo.' };
  }
  if (!b64) return { ok: false, error: 'El archivo llegó vacío.' };
  // ~8 MB en base64. Por encima, Apps Script se queda sin tiempo.
  if (b64.length > 11000000) {
    return { ok: false, error: 'El archivo es muy grande (más de 8 MB). ' +
             'Expórtalo por rangos de fecha más cortos.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  let filas;
  try {
    filas = leerArchivo(b64, nombre);
  } catch (err) {
    return { ok: false, error: 'No pude leer el archivo: ' + err.message };
  }
  if (!filas || filas.length < 2) {
    return { ok: false, error: 'El archivo no tiene filas de datos.' };
  }

  // El crudo se archiva en su pestaña, con el nombre por tienda para que
  // dos tiendas con la misma plataforma no se pisen.
  const cfg = FUENTES[fuente];
  const base = cfg.tab ||
    ('_Import_' + fuente.charAt(0).toUpperCase() + fuente.slice(1));
  const nomTab = base + '_' + tienda.toUpperCase();
  let sh = ss.getSheetByName(nomTab);
  if (!sh) { sh = ss.insertSheet(nomTab); sh.setTabColor('#cccccc'); }
  sh.clear();

  const ancho = Math.max.apply(null, filas.map(function (f) { return f.length; }));
  const rect = filas.map(function (f) {
    const r = f.slice();
    while (r.length < ancho) r.push('');
    return r;
  });
  /**
   * La pestaña cruda se escribe como TEXTO, a la fuerza.
   *
   * Sin esto, Google Sheets reinterpreta lo que escribimos según el
   * idioma de la hoja: "03-09-2026" se vuelve 9 de marzo en vez de 3 de
   * septiembre, y el daño ocurre al guardar, no al leer. Un pedido de
   * esta semana aparecía con 187 días de espera.
   *
   * Como texto, lo que entra es exactamente lo que traía el archivo, y
   * quien decide cómo se lee es el importador —que sabe que en América
   * Latina el día va primero— y no el idioma de la hoja.
   */
  sh.getRange(1, 1, rect.length, ancho).setNumberFormat('@');
  sh.getRange(1, 1, rect.length, ancho).setValues(rect);
  SpreadsheetApp.flush();

  /**
   * Un archivo propio no se importa a ciegas.
   *
   * El export de Dropi siempre trae las mismas columnas; el control
   * diario que alguien lleva a mano, no. Ahí "Tel" puede ser el
   * teléfono de la clienta o el de la transportadora, y nadie más que
   * quien lo escribió lo sabe. Así que se lee, se propone, y se espera
   * confirmación antes de escribir una sola fila en Pedidos.
   */
  if (cfg.propio) {
    if (p.mapeo && Object.keys(p.mapeo).length) {
      guardarMapeo(ss, fuente, p.mapeo, s.nombre);
    } else {
      let an;
      try {
        an = analizarFilas(filas, SINONIMOS_PEDIDOS, FORMAS_PEDIDOS);
      } catch (err) {
        return { ok: false, error: err.message, archivado: nomTab };
      }
      return { ok: true, revisar: true, archivo: nombre, tab: nomTab,
               filas: an.filas, encabezados: an.encabezados,
               propuestas: an.propuestas, sinResolver: an.sinResolver,
               muestra: an.muestra,
               campos: Object.keys(SINONIMOS_PEDIDOS) };
    }
  }

  let res;
  try {
    res = importar(fuente, tienda, s.sheetId);
  } catch (err) {
    return { ok: false, error: err.message, archivado: nomTab };
  }

  registrarMovimiento(s, 'Fuentes', fuente + '/' + tienda, 'importacion',
                      '', nombre + ' · ' + (filas.length - 1) + ' filas');
  return { ok: true, resumen: res, archivo: nombre,
           filas: filas.length - 1, tab: nomTab };
}

/**
 * Convierte el archivo subido en una matriz de filas.
 *
 * CSV se parsea directo. XLSX no: es un zip binario, así que se sube a
 * Drive pidiendo conversión a hoja de cálculo, se lee, y se borra.
 * El archivo temporal se elimina siempre, incluso si la lectura falla.
 */
function leerArchivo(b64, nombre) {
  const bytes = Utilities.base64Decode(b64);
  const ext = String(nombre).toLowerCase().split('.').pop();

  if (ext === 'csv' || ext === 'txt') {
    let texto = Utilities.newBlob(bytes).getDataAsString('UTF-8');
    // Un CSV de Excel suele venir en latin1: si aparece el carácter de
    // reemplazo, se reintenta con esa codificación.
    if (texto.indexOf('�') !== -1) {
      texto = Utilities.newBlob(bytes).getDataAsString('ISO-8859-1');
    }
    const sep = detectarSeparador(texto);
    return Utilities.parseCsv(texto, sep);
  }

  const blob = Utilities.newBlob(bytes,
    'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet', nombre);
  let id = null;
  try {
    id = convertirAHoja(blob);
    return abrirConvertida(id, nombre).getSheets()[0].getDataRange().getValues();
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
}

/**
 * Abrir el Excel recién convertido, esperando a que exista de verdad.
 *
 * Drive contesta con el id en cuanto acepta la subida, pero la hoja
 * todavía no está lista para Sheets: hay unos segundos en que ese id
 * existe y `openById` responde "No se puede abrir el archivo en estos
 * momentos" con un 404 de Drive. Cuanto más grande el archivo, más dura
 * esa ventana — un export de Meta de cuatro meses la abre de sobra.
 *
 * No es un error que se arregle reintentando la subida entera: es
 * esperar. Seis intentos con pausas crecientes suman unos veinte
 * segundos, que es mucho menos de lo que tarda la conversión misma.
 */
function abrirConvertida(id, nombre) {
  const esperas = [500, 1000, 2000, 4000, 6000, 6000];
  let ultimo = null;
  for (let i = 0; i < esperas.length; i++) {
    try { return SpreadsheetApp.openById(id); }
    catch (err) {
      ultimo = err;
      Utilities.sleep(esperas[i]);
    }
  }
  throw new Error(
    'Google convirtió "' + nombre + '" pero todavía no lo deja abrir. ' +
    'Suele pasar con archivos grandes: vuelve a intentarlo en un minuto, ' +
    'o exporta el reporte en CSV, que no necesita conversión y entra directo. ' +
    '(' + String(ultimo && ultimo.message || ultimo) + ')');
}

/**
 * Un .xlsx se lee convirtiéndolo a hoja de cálculo. Hay dos maneras y se
 * intentan las dos a propósito.
 *
 * La primera es el servicio avanzado de Drive, que es la vía documentada
 * pero hay que activarla a mano en cada proyecto. Cuando no está, el
 * error que salía era "Drive is not defined": cierto, inútil, y le caía
 * a quien subía su primer archivo.
 *
 * La segunda es la misma API por HTTP, con el token del propio script.
 * No hay que activar nada. Es la que hace que esto funcione recién
 * instalado en la cuenta de un cliente, que es donde tiene que funcionar.
 */
function convertirAHoja(blob) {
  if (typeof Drive !== 'undefined' && Drive.Files && Drive.Files.insert) {
    return Drive.Files.insert({ title: 'nova_tmp_' + Date.now(),
      mimeType: MimeType.GOOGLE_SHEETS }, blob).id;
  }

  const lim = '-nova-' + Utilities.getUuid();
  const meta = { name: 'nova_tmp_' + Date.now(), mimeType: MimeType.GOOGLE_SHEETS };
  const cabeza = Utilities.newBlob(
    '--' + lim + '\r\nContent-Type: application/json; charset=UTF-8\r\n\r\n' +
    JSON.stringify(meta) + '\r\n' +
    '--' + lim + '\r\nContent-Type: ' + blob.getContentType() + '\r\n\r\n').getBytes();
  const cola = Utilities.newBlob('\r\n--' + lim + '--').getBytes();

  const r = UrlFetchApp.fetch(
    'https://www.googleapis.com/upload/drive/v3/files?uploadType=multipart&fields=id', {
      method: 'post',
      contentType: 'multipart/related; boundary=' + lim,
      payload: Utilities.newBlob(cabeza.concat(blob.getBytes(), cola)).getBytes(),
      headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
      muteHttpExceptions: true,
    });

  if (r.getResponseCode() >= 300) {
    throw new Error(
      'Google no pudo convertir el Excel (' + r.getResponseCode() + '). ' +
      'Exporta el archivo en CSV y súbelo así, que no necesita conversión.');
  }
  const id = JSON.parse(r.getContentText()).id;
  if (!id) throw new Error('La conversión no devolvió un archivo.');
  return id;
}

/**
 * Guarda en `Mapeos` lo que la persona confirmó.
 *
 * Se marca 'humano' a propósito: proponerMapeo() borra sus propias
 * propuestas al recalcular, pero respeta las humanas. Una corrección
 * hecha a mano no debe perderse porque el análisis se volvió a correr.
 *
 * Y queda guardado para la próxima vez: quien sube su histórico mes a
 * mes con el mismo formato no vuelve a confirmar nada.
 */
function guardarMapeo(ss, fuenteId, mapeo, quien) {
  let sh = ss.getSheetByName('Mapeos');
  if (!sh) {
    sh = ss.insertSheet('Mapeos');
    sh.getRange(1, 1, 1, 7)
      .setValues([['fuente','campo_nova','columna_origen','confianza','aviso','definido_por','fecha']])
      .setFontWeight('bold').setBackground('#0b1824').setFontColor('#c9a84c');
    sh.setFrozenRows(1);
  }

  // Fuera las filas anteriores de esta fuente: el mapeo nuevo manda
  const todo = sh.getDataRange().getValues();
  for (let i = todo.length - 1; i >= 1; i--) {
    if (String(todo[i][0]) === fuenteId) sh.deleteRow(i + 1);
  }

  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const filas = Object.keys(mapeo)
    .filter(function (campo) { return String(mapeo[campo] || '').trim(); })
    .map(function (campo) {
      return [fuenteId, campo, String(mapeo[campo]).trim(), 'confirmada', '',
              'humano', hoy];
    });
  if (filas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, 7).setValues(filas);
  }
  SpreadsheetApp.flush();
  return filas.length;
}

/** Coma o punto y coma: Excel en español exporta con punto y coma. */
function detectarSeparador(texto) {
  const linea = texto.split(/\r?\n/)[0] || '';
  return (linea.split(';').length > linea.split(',').length) ? ';' : ',';
}

/** Las fuentes configuradas para una tienda, para poblar el selector. */
function apiFuentes(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('Fuentes');
  const out = [];
  if (sh && sh.getLastRow() > 1) {
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    d.slice(1).forEach(function (f) {
      if (String(f[c('tienda')]).trim() !== tienda) return;
      out.push({
        fuente: String(f[c('fuente')]).trim(),
        tipo: String(f[c('tipo')] || '').trim(),
        ultima: f[c('ultima_importacion')] || '',
        filas: f[c('filas_ultima')] || '',
      });
    });
  }
  // Todo lo que Nova sabe leer, por si la tienda aún no lo tiene declarado
  const catalogo = Object.keys(FUENTES).map(function (k) {
    return { fuente: k, tipo: FUENTES[k].tipo, verificado: !!FUENTES[k].verificado };
  });
  return { ok: true, tienda: tienda, configuradas: out, catalogo: catalogo };
}

// ─── PRUEBA ──────────────────────────────────────────────────

/**
 * Simula un login completo sin salir del editor. Úsalo antes de
 * publicar para confirmar que Equipo está bien cargado.
 */
function probarApi(email) {
  // Sin argumento usa la cuenta que corre el script, pero eso obliga a que
  // ese correo esté en Equipo. Pasando uno se prueba cualquier persona:
  //   probarApi('gestora@tutienda.com')
  email = String(email || Session.getEffectiveUser().getEmail()).toLowerCase().trim();
  const p = buscarPersona(email);
  if (!p) {
    const msg = 'El correo ' + email + ' no está en la hoja Equipo de ningún cliente.\n\n' +
      'Agrégalo en Nova_Empresarial_Nutrea → Equipo:\n' +
      '  id · nombre · correo · rol (dueno/admin/gestora) · tienda (* = todas) · estado (activo)\n\n' +
      'O prueba con un correo que ya esté en la lista:\n' +
      '  probarApi(\'otro@correo.com\')';
    Logger.log(msg);
    return msg;
  }
  const s = { email: email, nombre: p.nombre, rol: p.rol, clienteId: p.clienteId,
              sheetId: p.sheetId, tiendas: p.tiendas, vence: Date.now() + 3600000 };
  const msg = [
    'Persona encontrada:',
    '  nombre  : ' + p.nombre,
    '  rol     : ' + p.rol,
    '  cliente : ' + p.clienteId,
    '  tiendas : ' + p.tiendas.join(', '),
    '',
    'Resumen de ' + p.tiendas[0] + ':',
    JSON.stringify(apiResumen(s, { tienda: p.tiendas[0] }).datos, null, 2),
  ].join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Existe solo para pedirle a Google el permiso de salir a internet.
 *
 * La conversión de Excel usa la API de Drive por HTTP, y ese permiso no
 * se concede solo: Apps Script lo pide la primera vez que se ejecuta
 * algo que lo use, desde el editor. Si nunca se corre desde ahí, la
 * aplicación web falla con "No tienes permiso para llamar a
 * UrlFetchApp.fetch" — que es cierto, y no dice qué hacer.
 *
 * Correr esto una vez y aceptar resuelve eso para siempre.
 */
function autorizar() {
  UrlFetchApp.fetch('https://www.googleapis.com/drive/v3/about?fields=user', {
    headers: { Authorization: 'Bearer ' + ScriptApp.getOAuthToken() },
    muteHttpExceptions: true,
  });
  return 'Permiso concedido. Ya puedes subir archivos de Excel.';
}

/**
 * Lo que de verdad se movió en la billetera este mes.
 *
 * Nova calcula la utilidad desde el export de órdenes, que son estimados:
 * un flete de lista, un costo de proveedor de catálogo. La cartera es el
 * extracto — lo que la plataforma cobró y abonó de verdad. Cuando los dos
 * no coinciden, el que tiene razón es el extracto.
 *
 * Ojo con la fecha: la cartera se mueve el día que la plata cambia de
 * manos, no el día que se creó el pedido. Los movimientos de agosto
 * incluyen pedidos de julio que se entregaron en agosto, y los pedidos de
 * agosto entregados en septiembre están en el mes siguiente. Por eso esto
 * NO reemplaza el cierre por cohorte: lo acompaña, y sirve para cuadrar.
 *
 * Los retiros van aparte de todo lo demás. No son gasto: son plata tuya
 * saliendo de la billetera —casi siempre para pagar la pauta, que es lo
 * que dicen los conceptos del historial de Nutrea— y meterlos como gasto
 * hundiría la utilidad de un mes que estuvo bien.
 */
function carteraDelMes(ss, tienda, mes) {
  const out = {
    hay: false, mes: mes,
    ganancia: 0, devoluciones: 0, fletes: 0, otros: 0,
    nGanancia: 0, nDevoluciones: 0, nFletes: 0,
    retiros: 0, nRetiros: 0, conceptosRetiro: {},
    recargas: 0,
    netoOperativo: 0, saldo: null, ultimaFecha: '',
    devolucionPromedio: 0,
  };
  const sh = ss.getSheetByName('Cartera');
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  if (c('fecha') === -1) return out;

  let ultimo = null;
  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    if (c('tienda') !== -1 && String(f[c('tienda')]).trim() !== tienda) continue;
    const fecha = aISO(f[c('fecha')], 'UTC');
    if (!fecha) continue;

    // El saldo más reciente es de la tienda entera, no del mes: sirve para
    // saber cuánta plata hay ahora, que es una pregunta sin mes.
    if (!ultimo || fecha > ultimo.fecha) {
      ultimo = { fecha: fecha, saldo: num(f[c('saldo_previo')]) + num(f[c('monto')]) };
    }
    if (fecha.slice(0, 7) !== mes) continue;

    out.hay = true;
    const monto = num(f[c('monto')]);
    const abs = Math.abs(monto);
    switch (String(f[c('clase')] || '').trim()) {
      case 'ganancia':   out.ganancia += abs;     out.nGanancia++;     break;
      case 'devolucion': out.devoluciones += abs; out.nDevoluciones++; break;
      case 'flete':      out.fletes += abs;       out.nFletes++;       break;
      case 'recarga':    out.recargas += abs;                          break;
      case 'retiro': {
        out.retiros += abs; out.nRetiros++;
        const cp = String(c('concepto_retiro') === -1 ? '' : f[c('concepto_retiro')] || '').trim();
        const k = cp || 'sin concepto';
        out.conceptosRetiro[k] = (out.conceptosRetiro[k] || 0) + abs;
        break;
      }
      default: out.otros += monto;
    }
  }

  out.netoOperativo = out.ganancia - out.devoluciones - out.fletes + out.otros;

  /**
   * Lo que cuesta sacar la plata.
   *
   * La plataforma descuenta un porcentaje de cada retiro. No lo factura
   * ni aparece en ningún reporte: sale restado del monto que llega al
   * banco, así que un cierre que mire solo ventas y pauta nunca lo ve.
   * Es pequeño por retiro y nada pequeño al final del mes.
   */
  const aj = ajustes(ss, tienda);
  const pct = Number(aj.retiro_pct);
  out.retiroPct = isNaN(pct) ? 0 : pct;
  out.costoRetiros = out.retiros * (out.retiroPct / 100);
  out.devolucionPromedio = out.nDevoluciones ? out.devoluciones / out.nDevoluciones : 0;
  if (ultimo) { out.saldo = ultimo.saldo; out.ultimaFecha = ultimo.fecha; }
  return out;
}

/**
 * Qué tiene cada mes y qué le falta para poder cerrarse.
 *
 * Cargar el histórico de un cliente no es subir un archivo: son cuatro
 * cosas por mes, y si falta una el cierre sale mal sin decir por qué. La
 * peor es la tasa de cambio — sin ella la pauta en otra moneda no se
 * suma, y el mes aparece con una utilidad estupenda que nunca existió.
 *
 * Por eso esto no calcula el cierre: lo que hace es decir, mes a mes, si
 * hay pedidos, si hay pauta, si esa pauta se puede convertir, si hay
 * gastos fijos y si ya está cerrado. Convierte "¿por qué no me sale el
 * trimestre?" en una lista de lo que falta.
 *
 * Cada hoja se lee UNA vez y se reparte por mes. Llamar a agregarMes
 * ocho veces leería Pedidos entero ocho veces.
 */
function apiHistorial(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  if (s.rol !== 'dueno') {
    return { ok: false, error: 'El estado del histórico lo ve la dueña.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const tz = zonaHorariaDe(ss, tienda) || 'UTC';
  const monTienda = monedaDeTienda(ss, tienda);
  const cuantos = Math.min(Math.max(parseInt(p.meses, 10) || 8, 1), 18);

  // Los meses que se van a mirar: el actual y los anteriores
  const meses = [];
  let m = Utilities.formatDate(new Date(), tz, 'yyyy-MM');
  for (let i = 0; i < cuantos; i++) { meses.push(m); m = mesAnterior(m); }
  const idx = {};
  const out = meses.map(function (mm, i) {
    idx[mm] = i;
    return { mes: mm, pedidos: 0, entregados: 0, ventas: 0,
             pautaFilas: 0, pautaSinTasa: 0, monedasSinTasa: {},
             fijos: 0, nFijos: 0, cartera: 0, cerrado: false, falta: [] };
  });

  const bucket = function (fecha) {
    if (!fecha) return null;
    const k = fecha.slice(0, 7);
    return idx[k] === undefined ? null : out[idx[k]];
  };

  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const b = bucket(aISO(f[c('fecha')], tz));
      if (!b) continue;
      b.pedidos++;
      if (norm(f[c('estado_nova')] || f[c('estado_canonico')]) === 'entregado') {
        b.entregados++;
        b.ventas += num(f[c('valor')]);
      }
    }
  }

  // Las tasas se cargan una vez: preguntar por cada fila de pauta
  // volvería a leer la hoja entera cada vez.
  const tasas = {};
  const shT = ss.getSheetByName('Tasas');
  if (shT && shT.getLastRow() > 1) {
    shT.getDataRange().getValues().slice(1).forEach(function (f) {
      const fe = aISO(f[0], 'UTC');
      if (!fe) return;
      tasas[fe + '|' + String(f[1]).toUpperCase() + '|' + String(f[2]).toUpperCase()] = 1;
      tasas[fe + '|' + String(f[2]).toUpperCase() + '|' + String(f[1]).toUpperCase()] = 1;
    });
  }

  const shPa = ss.getSheetByName('Pauta');
  if (shPa && shPa.getLastRow() > 1) {
    const d = shPa.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], tz);
      const b = bucket(fecha);
      if (!b) continue;
      b.pautaFilas++;
      const mon = String(f[c('moneda_gasto')] || '').toUpperCase();
      if (!mon || mon === monTienda) continue;
      if (!tasas[fecha + '|' + mon + '|' + monTienda]) {
        b.pautaSinTasa++;
        b.monedasSinTasa[mon] = (b.monedasSinTasa[mon] || 0) + 1;
      }
    }
  }

  const shG = ss.getSheetByName('Gastos');
  if (shG && shG.getLastRow() > 1) {
    const d = shG.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      if (norm(f[c('activo')]) === 'no') continue;
      const v = num(f[c('valor')]);
      const mm = String(f[c('mes')] || '').trim();
      // Sin mes es recurrente: cuenta en todos
      if (!mm) { out.forEach(function (b) { b.fijos += v; b.nFijos++; }); continue; }
      const b = bucket(mm + '-01');
      if (b) { b.fijos += v; b.nFijos++; }
    }
  }

  const shC = ss.getSheetByName('Cartera');
  if (shC && shC.getLastRow() > 1) {
    const d = shC.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    if (c('fecha') !== -1) {
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (c('tienda') !== -1 && String(f[c('tienda')]).trim() !== tienda) continue;
        const b = bucket(aISO(f[c('fecha')], tz));
        if (b) b.cartera++;
      }
    }
  }

  const shCi = ss.getSheetByName('Cierres');
  if (shCi && shCi.getLastRow() > 1) {
    const d = shCi.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][c('tienda')]).trim() !== tienda) continue;
      if (norm(d[i][c('estado')]) !== 'cerrado') continue;
      const b = out[idx[String(d[i][c('mes')]).trim()]];
      if (b) b.cerrado = true;
    }
  }

  /**
   * Qué le falta a cada mes, en el orden en que importa.
   *
   * Sin pedidos no hay mes: lo demás da igual. La tasa va antes que la
   * pauta misma porque una pauta que no se puede convertir es peor que no
   * tener pauta — no se suma, y nadie lo nota.
   */
  const actual = meses[0];
  out.forEach(function (b) {
    if (!b.pedidos) { b.falta.push('pedidos'); return; }
    if (b.pautaSinTasa) b.falta.push('tasas');
    else if (!b.pautaFilas) b.falta.push('pauta');
    if (!b.nFijos) b.falta.push('gastos fijos');
    b.listo = !b.falta.length;
    b.enCurso = b.mes === actual;
  });

  return { ok: true, tienda: tienda, moneda: monTienda, meses: out };
}

/**
 * CAS: los pedidos estancados y los tickets que se les radicaron.
 *
 * Del SOP de Nutrea: cuando un pedido lleva días sin cambiar de estado,
 * llamar al cliente no sirve — el paquete no está con él. Lo que mueve
 * la aguja es presión administrativa sobre la transportadora: un ticket
 * oficial pidiendo prioridad de despacho, tres veces por semana, y
 * seguir insistiendo mientras no contesten.
 *
 * El criterio del SOP es "órdenes con días sin cambio de estado, de los
 * últimos 10 días". Los dos números son ajustables porque una tienda con
 * transportadora lenta no puede usar el mismo umbral que una rápida: se
 * toma `dias_sin_mover` de las alarmas, que ya es el umbral que la dueña
 * eligió para "esto lleva demasiado quieto".
 *
 * Nova detecta los candidatos sola. Lo que no puede saber —si alguien
 * radicó el ticket, con qué número, y qué contestaron— lo escribe el
 * equipo. Eso hoy vive en un Drive suelto que nadie más ve.
 */
const VENTANA_CAS_DIAS = 10;

function apiCas(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const tz = zonaHorariaDe(ss, tienda) || 'UTC';
  const u = umbrales(ss, tienda);
  const minDias = Number(u.dias_sin_mover) || 3;
  const hoy = new Date();
  /**
   * La ventana del SOP son 10 días, porque el CAS sirve para atajar a
   * tiempo. Pero quien nunca lo ha hecho tiene pedidos estancados de
   * antes, y esconderlos porque "ya no aplica el protocolo" sería
   * esconder plata parada. Se cuentan aparte y se pueden pedir.
   */
  const ventana = Math.max(1, parseInt(p.ventana, 10) || VENTANA_CAS_DIAS);

  // Los CAS ya radicados, por pedido
  const abiertos = {};
  const lista = [];
  const shC = ss.getSheetByName('CAS');
  if (shC && shC.getLastRow() > 1) {
    const d = shC.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const est = norm(f[c('estado')]) || 'abierto';
      const reg = {
        id: f[c('id')], pedidoId: String(f[c('pedido_id')] || ''),
        idExterno: String(f[c('id_externo')] || ''),
        guia: String(f[c('guia')] || ''),
        transportadora: String(f[c('transportadora')] || ''),
        ticket: String(f[c('ticket')] || ''),
        estado: est,
        abiertoEn: aISO(f[c('abierto_en')], tz) || '',
        abiertoPor: String(f[c('abierto_por')] || ''),
        ultimaGestion: aISO(f[c('ultima_gestion')], tz) || '',
        respuesta: String(f[c('respuesta')] || ''),
        nota: String(f[c('nota')] || ''),
        diasQuieto: num(f[c('dias_quieto')]),
      };
      // Cuántos días lleva el ticket sin que nadie lo toque
      const ref = reg.ultimaGestion || reg.abiertoEn;
      reg.diasSinTocar = ref
        ? Math.floor((hoy - new Date(ref + 'T00:00:00Z')) / 86400000) : null;
      lista.push(reg);
      if (est !== 'resuelto') abiertos[reg.pedidoId] = reg;
    }
  }

  // Candidatos: pedidos vivos, quietos, y de la ventana reciente
  const candidatos = [];
  let viejos = 0, valorViejos = 0;
  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const est = norm(f[c('estado_nova')] || f[c('estado_canonico')]);
      if (['entregado', 'devolucion', 'cancelado'].indexOf(est) !== -1) continue;

      const fecha = aISO(f[c('fecha')], tz);
      if (!fecha) continue;
      const edad = Math.floor((hoy - new Date(fecha + 'T00:00:00Z')) / 86400000);
      const ult = aISO(f[c('ultimo_movimiento')] || f[c('actualizado_en')], tz) || fecha;
      const quieto = Math.floor((hoy - new Date(ult + 'T00:00:00Z')) / 86400000);
      if (quieto < minDias) continue;

      // Estancado pero más viejo que la ventana: existe, y se dice
      if (edad > ventana) { viejos++; valorViejos += num(f[c('valor')]); continue; }

      const id = String(f[c('id')] || '');
      candidatos.push({
        pedidoId: id, idExterno: String(f[c('id_externo')] || ''),
        cliente: String(f[c('cliente')] || ''),
        ciudad: String(f[c('ciudad')] || ''),
        guia: String(f[c('guia')] || ''),
        transportadora: String(f[c('transportadora')] || ''),
        estado: est, valor: num(f[c('valor')]),
        fecha: fecha, diasQuieto: quieto, edad: edad,
        yaTiene: !!abiertos[id],
        cas: abiertos[id] || null,
      });
    }
  }

  // Lo más quieto primero: ahí la antigüedad es deuda
  candidatos.sort(function (a, b) { return b.diasQuieto - a.diasQuieto; });
  lista.sort(function (a, b) {
    const pa = a.estado === 'resuelto' ? 1 : 0, pb = b.estado === 'resuelto' ? 1 : 0;
    if (pa !== pb) return pa - pb;
    return (b.diasSinTocar || 0) - (a.diasSinTocar || 0);
  });

  return { ok: true, tienda: tienda, minDias: minDias, ventana: ventana,
           candidatos: candidatos, cas: lista,
           viejos: viejos, valorViejos: valorViejos,
           sinRadicar: candidatos.filter(function (x) { return !x.yaTiene; }).length };
}

/**
 * Radicar un CAS, o anotar qué pasó con uno ya radicado.
 *
 * El número de ticket no es obligatorio al abrirlo: en la práctica
 * primero se decide radicarlo y el número llega después. Obligarlo
 * empujaría a inventarse uno con tal de poder guardar.
 */
function apiCasEscribir(s, p) {
  const tienda = String(p.tienda || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = SpreadsheetApp.openById(s.sheetId);
  const sh = ss.getSheetByName('CAS');
  if (!sh) return { ok: false, error: 'Falta la hoja CAS. Corre bootstrapTodo().' };

  const ESTADOS = ['abierto', 'respondido', 'sin_respuesta', 'resuelto'];
  const estado = norm(p.estado || 'abierto');
  if (ESTADOS.indexOf(estado) === -1) {
    return { ok: false, error: 'Estado no válido. Sirven: ' + ESTADOS.join(', ') + '.' };
  }

  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  const hoy = new Date().toISOString().slice(0, 10);
  const id = String(p.id || '').trim();

  if (id) {
    const d = sh.getDataRange().getValues();
    const cId = enc.indexOf('id');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cId]).trim() !== id) continue;
      const cambios = {
        estado: estado, ultima_gestion: hoy,
        ticket: p.ticket !== undefined ? String(p.ticket).trim() : undefined,
        respuesta: p.respuesta !== undefined ? String(p.respuesta).trim() : undefined,
        nota: p.nota !== undefined ? String(p.nota).trim() : undefined,
        cerrado_en: estado === 'resuelto' ? hoy : undefined,
      };
      Object.keys(cambios).forEach(function (k) {
        if (cambios[k] === undefined) return;
        const col = enc.indexOf(k);
        if (col === -1) return;
        const antes = d[i][col];
        if (String(antes) === String(cambios[k])) return;
        sh.getRange(i + 1, col + 1).setValue(cambios[k]);
        registrarMovimiento(s, 'CAS', id, k, antes, cambios[k]);
      });
      return { ok: true, id: id };
    }
    return { ok: false, error: 'No encuentro ese CAS.' };
  }

  const pedidoId = String(p.pedido_id || '').trim();
  if (!pedidoId) return { ok: false, error: 'Falta decir de qué pedido es el CAS.' };

  const nuevo = 'cas-' + Utilities.getUuid().slice(0, 8);
  const valores = {
    id: nuevo, tienda: tienda, pedido_id: pedidoId,
    id_externo: String(p.id_externo || ''), guia: String(p.guia || ''),
    transportadora: String(p.transportadora || ''),
    abierto_en: hoy, abierto_por: s.email,
    ticket: String(p.ticket || '').trim(), estado: estado,
    dias_quieto: num(p.dias_quieto), ultima_gestion: hoy,
    respuesta: '', cerrado_en: '', nota: String(p.nota || '').trim(),
  };
  sh.appendRow(enc.map(function (col) {
    return valores[col] !== undefined ? valores[col] : '';
  }));
  registrarMovimiento(s, 'CAS', nuevo, 'radicado', '', pedidoId);
  return { ok: true, id: nuevo };
}


/* ═══════════════════════════════════════════════════════════════
   10 · NOVA CENTRAL
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Central (la consola de la plataforma)
 * ─────────────────────────────────────────────────────────────
 * Lo que Nova Central hace y Nova Empresarial no: crear clientes, ver
 * cuántos hay y en qué plan, y controlar el demo.
 *
 * ┌─ POR QUÉ UNA SESIÓN APARTE ────────────────────────────────┐
 * │                                                            │
 * │ La sesión de Empresarial se resuelve contra la hoja Equipo │
 * │ de UN cliente: dice de qué tiendas es esa persona y qué    │
 * │ permisos le dio su dueña. Aquí la pregunta es otra: quién  │
 * │ opera la plataforma. Es quien crea cuentas, ve la          │
 * │ facturación de todos los clientes y enciende el demo.      │
 * │                                                            │
 * │ Si fueran la misma, darle un permiso de más a la admin de  │
 * │ un cliente podría, por un descuido, abrirle la consola de  │
 * │ todos los demás. Son dos preguntas distintas y tienen dos  │
 * │ hojas, dos tokens y dos caducidades.                       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Todas las acciones de aquí empiezan por `nc_`, y el enrutador las
 * manda a este archivo ANTES de buscar una sesión de cliente.
 */

/** Cuánto dura una sesión de la consola. Más corta que la de una dueña. */
const TTL_CENTRAL_H = 6;

/** Qué puede hacer cada quien en la consola. */
const ROLES_CENTRAL = {
  socia:      ['ver', 'crear_cliente', 'planes', 'demo', 'facturacion'],
  operadora:  ['ver', 'crear_cliente', 'demo'],
};

function puedeCentral(s, permiso) {
  return (ROLES_CENTRAL[s.rol] || []).indexOf(permiso) !== -1;
}

/**
 * Quién es esta persona en la plataforma.
 *
 * Solo la hoja Plataforma de Nova_Central. No se cae hacia Equipo: si
 * alguien no está aquí, no entra, aunque sea dueña de tres tiendas.
 */
function buscarOperadora(email) {
  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName('Plataforma');
  if (!sh || sh.getLastRow() < 2) return null;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const buscado = String(email || '').toLowerCase().trim();

  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    if (String(f[c('correo')] || '').toLowerCase().trim() !== buscado) continue;
    if (norm(f[c('estado')]) === 'inactivo') return null;
    const rol = norm(f[c('rol')]) === 'operadora' ? 'operadora' : 'socia';
    return { id: f[c('id')], nombre: f[c('nombre')], correo: buscado,
             rol: rol, fila: i + 1 };
  }
  return null;
}

function sesionCentral(token) {
  const crudo = CacheService.getScriptCache().get('nc_' + String(token || ''));
  if (!crudo) return null;
  const s = JSON.parse(crudo);
  if (Date.now() > s.vence) return null;
  return s;
}

/**
 * El enrutador de la consola.
 *
 * Va antes que el de clientes porque estas acciones no tienen sesión de
 * cliente que buscar: buscarla primero devolvería "sesión inválida" a
 * quien está intentando entrar a la consola, que es confuso.
 */
function manejarCentral(accion, p) {
  if (accion === 'nc_login')     return centralLogin(p);
  if (accion === 'nc_verificar') return centralVerificar(p);

  const s = sesionCentral(p.token);
  if (!s) return { ok: false, error: 'Sesión vencida o inválida.', reautenticar: true };

  switch (accion) {
    case 'nc_yo':       return { ok: true, sesion: publicoCentral(s) };
    case 'nc_clientes': return centralClientes(s, p);
    case 'nc_planes':   return centralPlanes(s, p);
    case 'nc_crear':    return centralCrearCliente(s, p);
    case 'nc_automatico':        return centralAutomatico(s, p);
    case 'nc_automatico_prender':return centralPrenderAutomatico(s, p);
    case 'nc_salir':
      CacheService.getScriptCache().remove('nc_' + p.token);
      return { ok: true };
    default: return { ok: false, error: 'Acción desconocida: ' + accion };
  }
}

function publicoCentral(s) {
  return { correo: s.correo, nombre: s.nombre, rol: s.rol,
           vence: s.vence, horas: TTL_CENTRAL_H,
           permisos: ROLES_CENTRAL[s.rol] || [] };
}

/**
 * El código de acceso.
 *
 * Responde lo mismo exista la persona o no. Decir "ese correo no está en
 * la plataforma" convierte esta pantalla en una forma de averiguar quién
 * trabaja aquí, y esta consola crea cuentas de clientes.
 */
function centralLogin(p) {
  const email = String(p.email || '').toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, error: 'Correo inválido.' };
  }
  const op = buscarOperadora(email);
  if (op) {
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put('nccod_' + email, codigo, TTL_CODIGO_M * 60);
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Tu código de Nova Central: ' + codigo,
        body: 'Hola ' + (op.nombre || '') + ',\n\n' +
              'Tu código para entrar a Nova Central es: ' + codigo + '\n\n' +
              'Vence en ' + TTL_CODIGO_M + ' minutos.\n' +
              'Desde ahí se crean cuentas de clientes: si no fuiste tú, ' +
              'avísale a la otra socia.\n',
      });
    } catch (err) {
      return { ok: false, error: 'No se pudo enviar el código: ' + err.message };
    }
  }
  return { ok: true, enviado: true, vence_en_min: TTL_CODIGO_M };
}

function centralVerificar(p) {
  const email = String(p.email || '').toLowerCase().trim();
  const codigo = String(p.codigo || '').trim();
  const cache = CacheService.getScriptCache();
  const esperado = cache.get('nccod_' + email);

  if (!esperado || esperado !== codigo) {
    return { ok: false, error: 'Código incorrecto o vencido.' };
  }
  cache.remove('nccod_' + email);

  const op = buscarOperadora(email);
  if (!op) return { ok: false, error: 'Esta cuenta ya no tiene acceso.' };

  const token = Utilities.getUuid();
  const s = { correo: email, nombre: op.nombre, rol: op.rol, id: op.id,
              vence: Date.now() + TTL_CENTRAL_H * 3600000 };
  cache.put('nc_' + token, JSON.stringify(s), TTL_CENTRAL_H * 3600);

  // Rastro de quién entra a la consola, en la propia hoja
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Plataforma');
    const e = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const col = e.indexOf('ultima_conexion');
    if (col !== -1) sh.getRange(op.fila, col + 1).setValue(ahoraISO());
  } catch (err) { /* el rastro no puede impedir entrar */ }

  return { ok: true, token: token, sesion: publicoCentral(s) };
}

/**
 * Los clientes de la plataforma.
 *
 * Cada fila trae además lo que no está en Clientes y sí importa para
 * saber si esa cuenta está viva: cuántas personas tiene en su Equipo y
 * cuándo fue la última importación. Un cliente que pagó y no ha subido
 * un archivo en tres semanas es el que hay que llamar hoy.
 *
 * Abrir la hoja de cada cliente cuesta, así que se hace solo si lo
 * piden: la lista sola se dibuja al instante.
 */
function centralClientes(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve los clientes.' };

  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName('Clientes');
  if (!sh || sh.getLastRow() < 2) return { ok: true, clientes: [] };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const conDetalle = !!p.detalle;
  const verDinero = puedeCentral(s, 'facturacion');

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    const empresa = String(f[c('empresa')] || '').trim();
    if (!empresa) continue;

    const cl = {
      id: f[c('id')], empresa: empresa,
      pais: String(f[c('pais')] || ''),
      plan: String(f[c('plan')] || ''),
      estado: norm(f[c('estado')]) || 'activo',
      alta: aISO(f[c('fecha_alta')], 'UTC') || '',
      corte: aISO(f[c('fecha_corte')], 'UTC') || '',
      ultimoPago: aISO(f[c('ultimo_pago')], 'UTC') || '',
      tiendas: num(f[c('tiendas')]),
      usuarios: num(f[c('usuarios')]),
      sheetId: String(f[c('sheet_id')] || ''),
    };
    // La tarifa es dinero: solo la socia
    if (verDinero) {
      cl.tarifa = num(f[c('tarifa')]);
      cl.costo = num(f[c('costo')]);
    }

    if (conDetalle && cl.sheetId) {
      try {
        const cs = SpreadsheetApp.openById(cl.sheetId);
        const shE = cs.getSheetByName('Equipo');
        cl.personas = shE && shE.getLastRow() > 1 ? shE.getLastRow() - 1 : 0;
        const shF = cs.getSheetByName('Fuentes');
        if (shF && shF.getLastRow() > 1) {
          const df = shF.getDataRange().getValues();
          const ef = df[0].map(norm);
          const ci = ef.indexOf('ultima_importacion');
          let ult = '';
          for (let j = 1; j < df.length; j++) {
            const u = aISO(df[j][ci], 'UTC');
            if (u && u > ult) ult = u;
          }
          cl.ultimaImportacion = ult;
        }

        /**
         * Estados que ese cliente tiene sin clasificar.
         *
         * Aparece aquí porque el cliente no siempre se da cuenta: Nova se
         * lo dice en su pantalla, pero quien vende Nova necesita saberlo
         * antes de que le llamen diciendo que las cifras no cuadran. Un
         * estado nuevo de una transportadora suele afectar a varios
         * clientes del mismo país a la vez.
         */
        const shEst = cs.getSheetByName('Estados');
        if (shEst && shEst.getLastRow() > 1) {
          const de = shEst.getDataRange().getValues();
          const ee = de[0].map(norm);
          const cEst = ee.indexOf('estado_nova'), cTx = ee.indexOf('texto'),
                cPed = ee.indexOf('pedidos');
          const pend = [];
          for (let j = 1; j < de.length; j++) {
            if (String(de[j][cTx] || '').trim() && !norm(de[j][cEst])) {
              pend.push({ texto: String(de[j][cTx]), pedidos: num(de[j][cPed]) });
            }
          }
          pend.sort(function (a, b) { return b.pedidos - a.pedidos; });
          cl.estadosSinClasificar = pend.length;
          cl.pedidosSinClasificar = pend.reduce(function (t, x) { return t + x.pedidos; }, 0);
          cl.estadosNuevos = pend.slice(0, 5);
        }
      } catch (err) {
        // Una hoja borrada o sin permiso no puede tumbar la lista entera
        cl.problema = 'No se pudo abrir su hoja: ' + err.message;
      }
    }
    out.push(cl);
  }

  return { ok: true, clientes: out, verDinero: verDinero };
}

function centralPlanes(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve los planes.' };
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Planes');
  if (!sh || sh.getLastRow() < 2) return { ok: true, planes: [] };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const verDinero = puedeCentral(s, 'facturacion');

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const nombre = String(d[i][c('nombre')] || '').trim();
    if (!nombre) continue;
    const pl = { id: d[i][c('id')], nombre: nombre,
      modulos: String(d[i][c('modulos')] || '').split(/[,;]/)
        .map(function (x) { return x.trim(); }).filter(String),
      limiteUsuarios: num(d[i][c('limite_usuarios')]),
      limiteTiendas: num(d[i][c('limite_tiendas')]) };
    if (verDinero) {
      pl.precio = num(d[i][c('precio_sugerido')]);
      pl.tarifa = num(d[i][c('tarifa_fijada')]);
    }
    out.push(pl);
  }
  return { ok: true, planes: out, verDinero: verDinero };
}

/**
 * Crear un cliente desde la pantalla.
 *
 * Hasta ahora crearCliente() solo se podía llamar desde el editor de
 * Apps Script. Eso significa que vender una cuenta requería abrir el
 * código, y que nadie más que quien lo escribió podía hacerlo.
 *
 * Lo que se valida aquí y no dentro de crearCliente: que quien pide
 * tenga permiso, y que lo que llega de la pantalla tenga la forma que
 * esa función espera. Las reglas del negocio —que no falte la dueña, que
 * no haya dos tiendas con el mismo id, que la empresa no exista ya—
 * siguen viviendo allá, porque también valen cuando se llama a mano.
 */
function centralCrearCliente(s, p) {
  if (!puedeCentral(s, 'crear_cliente')) {
    return { ok: false, error: 'Tu rol no puede crear clientes.' };
  }

  const empresa = String(p.empresa || '').trim();
  const pais = String(p.pais || '').trim().toUpperCase();
  const plan = String(p.plan || 'Base').trim();
  const dueno = p.dueno || {};

  if (!empresa) return { ok: false, error: 'Falta el nombre de la empresa.' };
  if (!String(dueno.correo || '').trim() || String(dueno.correo).indexOf('@') === -1) {
    return { ok: false, error: 'Falta el correo de la dueña. Sin una dueña en su ' +
             'hoja Equipo, nadie podría entrar a esa cuenta — ni ella.' };
  }

  const tiendas = (p.tiendas || []).filter(function (t) {
    return t && String(t.id || '').trim();
  }).map(function (t) {
    return {
      id: String(t.id).trim().toLowerCase(),
      nombre: String(t.nombre || t.id).trim(),
      pais: String(t.pais || pais || '').toUpperCase(),
      moneda: String(t.moneda || '').toUpperCase(),
      zona: String(t.zona || '').trim(),
      modalidad: String(t.modalidad || 'catalogo_publico').trim(),
    };
  });
  if (!tiendas.length) {
    return { ok: false, error: 'Hay que declarar al menos una tienda.' };
  }
  const sinMoneda = tiendas.filter(function (t) { return !t.moneda; });
  if (sinMoneda.length) {
    return { ok: false, error: 'Falta la moneda de: ' +
             sinMoneda.map(function (t) { return t.nombre; }).join(', ') +
             '. Sin moneda, Nova no puede sumar ni convertir nada de esa tienda.' };
  }

  const fuentes = (p.fuentes || []).filter(function (f) {
    return f && f.tienda && f.fuente;
  });

  let r;
  try {
    r = crearCliente(empresa, pais, tiendas, fuentes, plan, {
      nombre: String(dueno.nombre || '').trim() || 'Dueña',
      correo: String(dueno.correo).toLowerCase().trim(),
    });
  } catch (err) {
    return { ok: false, error: err.message };
  }

  registrarCentral(s, 'Clientes', r.clienteId, 'creado', '',
                   empresa + ' · ' + tiendas.length + ' tienda(s) · plan ' + plan);
  return { ok: true, clienteId: r.clienteId, url: r.url, empresa: empresa,
           dueno: String(dueno.correo).toLowerCase().trim() };
}

/** La bitácora de la consola vive en Nova_Central, no en la del cliente. */
function registrarCentral(s, entidad, id, campo, antes, ahora) {
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Movimientos');
    if (sh) sh.appendRow([ahoraISO(), s.correo, entidad, id, campo, antes, ahora]);
  } catch (err) {
    Logger.log('No se pudo registrar en Central: ' + err.message);
  }
}

/**
 * ★ Darse de alta en la consola, la primera vez ★
 *
 * Se entra a Nova Central con un correo que esté en la hoja Plataforma,
 * y esa hoja nace vacía: sin esto nadie podría entrar nunca.
 *
 * SE CORRE SIN ESCRIBIR NADA. Eliges primeraSocia en el editor, le das a
 * Ejecutar, y se da de alta la cuenta desde la que estás corriendo el
 * script — que es la dueña de las hojas, así que ya es tuya.
 *
 * El botón Ejecutar del editor no sabe pasar argumentos, y pedirle a
 * alguien que edite el código para darse de alta es pedirle que toque lo
 * único que no debería tener que tocar. Si quieres otro correo, se puede
 * pasar: primeraSocia("Nombre", "otro@correo.com")
 *
 * Solo funciona mientras la hoja esté vacía. Después, quien agrega gente
 * es quien ya está dentro.
 */
function primeraSocia(nombre, correo) {
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Plataforma');
  if (!sh) throw new Error('Falta la hoja Plataforma. Corre bootstrapTodo().');
  if (sh.getLastRow() > 1) {
    const ya = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
      .map(function (f) { return f[2]; }).filter(String).join(', ');
    throw new Error(
      'La hoja Plataforma ya tiene gente: ' + ya + '\n\n' +
      'Para agregar a alguien más, escribe su fila directamente ahí: ' +
      'nombre, correo, rol (socia u operadora) y estado "activo".');
  }

  // Sin correo, el de la cuenta que está corriendo esto
  const mail = String(correo || Session.getEffectiveUser().getEmail() || '')
    .toLowerCase().trim();
  if (!mail || mail.indexOf('@') === -1) {
    throw new Error(
      'No pude averiguar tu correo solo. Córrela así:\n\n' +
      '  primeraSocia("Manuela", "tucorreo@dominio.com")');
  }

  sembrar(sh.getParent(), 'Plataforma', [{
    id: 'pl-' + Utilities.getUuid().slice(0, 8),
    nombre: nombre || mail.split('@')[0],
    correo: mail,
    rol: 'socia', estado: 'activo', ultima_conexion: '', nota: 'primera socia',
  }], 2);

  const msg = [
    'Listo.',
    '',
    'Entra a Nova Central con este correo:  ' + mail,
    'Ahí te llega el código de 6 dígitos.',
    '',
    'Si prefieres otro correo, bórralo de la hoja Plataforma y vuelve a',
    'correr esto pasándole el que quieras.',
  ].join('\n');
  Logger.log(msg);
  return msg;
}


// ─── LO AUTOMÁTICO, DESDE LA CONSOLA ─────────────────────────

/**
 * ═══════════════════════════════════════════════════════════
 *  EL INTERRUPTOR DE LO AUTOMÁTICO
 * ═══════════════════════════════════════════════════════════
 *
 * Los disparadores que hacen correr a Nova sola —las tasas de cambio y
 * la revisión de alarmas— vivían únicamente en el Apps Script. Estaban
 * escritos desde el principio y nunca se prendieron, porque prenderlos
 * exigía abrir el editor, encontrar la función y ejecutarla a mano.
 *
 * Eso no se le puede pedir a nadie, y menos a un cliente. Pero tampoco
 * va en la pantalla de la dueña: los disparadores son del proyecto
 * entero, no de cada cuenta, y un botón ahí le daría a cualquier cliente
 * un interruptor que afecta a todos los demás.
 *
 * Así que va aquí, en la consola, que es de Manuela.
 */
function centralAutomatico(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve esto.' };
  return { ok: true, automatico: estadoAutomatico() };
}

/**
 * Prende los dos trabajos y programa la carga de tasas para dentro de un
 * minuto.
 *
 * La carga NO se hace aquí. Bajar noventa días de tasas de GOOGLEFINANCE
 * para varios clientes tarda más de lo que un navegador espera, y el
 * botón se quedaría colgado sin que nadie supiera si funcionó. En vez de
 * eso se arma un disparador de un solo uso: la respuesta vuelve al
 * instante y el trabajo pesado corre solo, por detrás.
 */
function centralPrenderAutomatico(s, p) {
  if (!puedeCentral(s, 'crear_cliente')) {
    return { ok: false, error: 'Tu rol no puede prender lo automático.' };
  }

  let hechos;
  try {
    hechos = prenderTrabajos_();
  } catch (e) {
    /**
     * Crear disparadores pide un permiso que la autorización vieja del
     * script puede no tener. No se disfraza de otro error: se dice qué
     * pasó y cuál es la salida, que es abrir el editor una vez.
     */
    return { ok: false, error:
      'No pude crear los disparadores: ' + e.message + '\n\n' +
      'Suele ser que la autorización del script es anterior a esta ' +
      'función. Abre el Apps Script, corre prenderAutomatico() una vez ' +
      'y acepta los permisos; después este botón ya funciona.' };
  }

  // Y la carga inicial, por detrás.
  let cargando = false;
  try {
    const yaHay = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'cargaInicialTasas';
    });
    if (!yaHay) {
      ScriptApp.newTrigger('cargaInicialTasas').timeBased().after(60 * 1000).create();
    }
    cargando = true;
  } catch (e) {
    cargando = false;
  }

  return {
    ok: true,
    prendidos: hechos,
    cargando: cargando,
    automatico: estadoAutomatico(),
    mensaje: cargando
      ? 'Listo. Las tasas de los últimos 90 días empiezan a bajar en un minuto; ' +
        'según cuántas cuentas haya puede tardar varios. Vuelve a mirar en un rato.'
      : 'Los trabajos quedaron prendidos, pero no pude programar la carga de ' +
        'las tasas viejas. Córrela a mano: cargaInicialTasas() en el Apps Script.',
  };
}


/* ═══════════════════════════════════════════════════════════════
   11 · ESCRITURA
   ═══════════════════════════════════════════════════════════════ */

/**
 * Nova · Importar (el lado que escribe)
 * ─────────────────────────────────────────────────────────────
 * `leerCrudo()` normaliza pero no guarda nada. Esto es lo que convierte
 * un export pegado en filas que la app puede leer.
 *
 * Tres reglas que no se rompen:
 *
 *   1. NUNCA se sobrescribe lo que escribió el equipo.
 *      Si una gestora puso una nota o cambió el estado, la siguiente
 *      importación no lo borra. Se actualizan solo las columnas que
 *      vienen de la plataforma.
 *
 *   2. NUNCA se borran filas.
 *      Un pedido que desaparece del export no se elimina: el histórico
 *      es lo que hace posibles las tendencias.
 *
 *   3. Deduplicación por fuente + id_externo.
 *      Reimportar el mismo archivo dos veces no duplica nada, actualiza.
 *
 * USO:
 *   1. Pega el export en su pestaña _Import_*
 *   2. importar('dropi', 'ec')
 */

// Columnas que escribe la app y el importador jamás toca.
const COLUMNAS_DEL_EQUIPO = [
  'estado_nova', 'nota', 'solucion', 'gestora_asignada', 'fecha_promesa',
  'intentos', 'resuelta_en', 'telefono_2', 'telefono_2_norm',
  'actualizado_en', 'actualizado_por',
];

/**
 * Importa una fuente a su entidad.
 *
 * @param {string} fuenteId  'dropi' | 'mastershop' | 'effi_guias' | ...
 * @param {string} tienda    id de la tienda: 'gt' | 'ec'
 * @param {string} cliente   nombre del cliente (o vacío si solo hay uno)
 */
function importar(fuenteId, tienda, cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  if (!tienda) throw new Error('Falta la tienda. Ej: importar("dropi","ec")');
  if (!monedaDeTienda(ss, tienda)) {
    throw new Error('La tienda "' + tienda + '" no está en la hoja Tiendas.');
  }

  // Todo el import corre con el formato de fecha que eligió la dueña
  return conFormatoFecha(ss, tienda, function () {
    return importarConFormato(ss, fuenteId, tienda);
  });
}

function importarConFormato(ss, fuenteId, tienda) {
  const r = leerCrudo(ss, fuenteId, tienda);
  if (!r.filas.length) {
    const msg = r.tab + ' está vacía. Pega el export ahí primero.';
    Logger.log(msg);
    return msg;
  }

  const destino = { pedidos: 'Pedidos', novedades: 'Novedades',
                    llamadas: 'Llamadas', pauta: 'Pauta',
                    facturacion: 'Facturacion', cartera: 'Cartera',
                    pedidos_secundario: 'Pedidos' }[r.tipo];
  if (!destino) throw new Error('No sé dónde guardar una fuente de tipo ' + r.tipo);

  let extra = '';
  const pais = paisDeTienda(ss, tienda);

  /**
   * Lo que la dueña ya clasificó, leído una sola vez.
   *
   * Si cada fila abriera la hoja Estados, un archivo de cuarenta mil
   * pedidos la abriría cuarenta mil veces y el script se quedaría sin sus
   * seis minutos antes de escribir nada.
   */
  const aprendidos = estadosAprendidos(ss);
  const vistos = {};

  const preparadas = r.filas.map(function (f) {
    return prepararFila(f, r.tipo, fuenteId, tienda, pais, ss, aprendidos, vistos);
  });

  // Lo que se vio queda anotado, se entendiera o no
  anotarEstados(ss, vistos);

  /**
   * Qué cambió respecto a la carga anterior.
   *
   * "206 actualizadas" no dice nada. Lo que alguien necesita saber al
   * subir el archivo del día es qué pasó desde ayer: cuántos llegaron,
   * cuántos se entregaron, cuántos se devolvieron y qué queda por
   * resolver. Eso se calcula comparando el estado que había en la hoja
   * contra el que trae el archivo, ANTES de escribir encima.
   */
  const antes = (r.tipo === 'pedidos') ? estadosActuales(ss, destino, tienda) : null;

  const res = escribirFilas(ss, destino, preparadas, fuenteId);
  registrarImportacion(ss, tienda, fuenteId, res.nuevas + res.actualizadas);

  if (antes) extra = novedadesDeLaCarga(antes, preparadas) + extra;

  // Las novedades que vienen dentro del export de pedidos se derivan aparte
  if (r.tipo === 'pedidos' && fuenteId === 'dropi') {
    const nov = derivarNovedades(preparadas, fuenteId, tienda);
    if (nov.length) {
      const rn = escribirFilas(ss, 'Novedades', nov, fuenteId);
      extra = '\nNovedades derivadas: ' + rn.nuevas + ' nuevas, ' +
              rn.actualizadas + ' actualizadas, ' + rn.iguales + ' sin cambios';
    }
  }

  /**
   * Una fila de pauta que cubre dos meses no se puede repartir sin
   * inventar: el gasto de una campaña no se distribuye parejo por día.
   *
   * Así que no se reparte — se avisa. Si no, todo el gasto de agosto a
   * septiembre se contaría en agosto y septiembre saldría en cero, que
   * es peor que un número que falta: es un número que miente.
   */
  if (r.tipo === 'pauta') {
    const cruzan = preparadas.filter(function (p) {
      return p.fecha_fin && String(p.fecha).slice(0, 7) !== String(p.fecha_fin).slice(0, 7);
    });
    if (cruzan.length) {
      extra += '\n⚠ ' + cruzan.length + ' fila(s) cubren más de un mes (' +
        preparadas[0].fecha + ' a ' + preparadas[0].fecha_fin + ').\n' +
        '   Todo ese gasto se contará en el primer mes. Para que cada mes ' +
        'reciba lo suyo,\n   vuelve a exportar en Meta con Desglose → Por día.';
    }
  }

  const msg = [
    'Importado: ' + fuenteId + ' → ' + destino + ' (tienda ' + tienda + ')',
    '  filas leídas   : ' + r.filas.length,
    '  nuevas         : ' + res.nuevas,
    '  actualizadas   : ' + res.actualizadas,
    '  sin cambios    : ' + res.iguales,
    r.sinMapear.length ? '  columnas sin mapear: ' + r.sinMapear.join(', ') : '',
    res.sinEstado.length
      ? '  ⚠ ' + res.sinEstado.length + ' estado(s) que no reconozco: ' +
        res.sinEstado.slice(0, 8).join(' · ') +
        '\n    Esos pedidos quedan SIN CLASIFICAR: no cuentan como entregados ' +
        'ni como devueltos.\n    Dinos qué significan en Configuración → Estados ' +
        'y las cifras se rehacen solas.'
      : '',
    extra,
  ].filter(String).join('\n');
  Logger.log(msg);
  return msg;
}

/** El estado que tiene hoy cada pedido en la hoja, por id. */
function estadosActuales(ss, hoja, tienda) {
  const sh = ss.getSheetByName(hoja);
  if (!sh || sh.getLastRow() < 2) return {};
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cT = e.indexOf('tienda');
  const cE = e.indexOf('estado_canonico'), cN = e.indexOf('estado_nova');
  const out = {};
  for (let i = 1; i < d.length; i++) {
    if (cT !== -1 && String(d[i][cT]).trim() !== tienda) continue;
    const id = String(d[i][cId]).trim();
    if (id) out[id] = norm(d[i][cN] || d[i][cE]);
  }
  return out;
}

/**
 * El resumen de lo que cambió, en la lengua de quien opera.
 *
 * No dice "38 filas actualizadas": dice qué pasó con los pedidos.
 */
function novedadesDeLaCarga(antes, ahora) {
  let nuevos = 0, entregados = 0, devueltos = 0, cancelados = 0,
      aNovedad = 0, pendientes = 0, sinCambio = 0;

  ahora.forEach(function (p) {
    const previo = antes[p.id];
    const est = norm(p.estado_canonico);
    if (previo === undefined) {
      nuevos++;
    } else if (previo !== est) {
      if (est === 'entregado')  entregados++;
      else if (est === 'devolucion') devueltos++;
      else if (est === 'cancelado')  cancelados++;
      else if (est === 'novedad')    aNovedad++;
    } else {
      sinCambio++;
    }
    if (['entregado','devolucion','cancelado'].indexOf(est) === -1) pendientes++;
  });

  // Uno o varios: "se devolvieron 1" delata que lo escribió una máquina
  const pl = function (n, uno, varios) { return n === 1 ? uno : varios.replace('%', n); };

  const partes = [];
  if (nuevos)     partes.push(pl(nuevos, 'entró 1 pedido nuevo', 'entraron % pedidos nuevos'));
  if (entregados) partes.push(pl(entregados, 'se entregó 1', 'se entregaron %'));
  if (devueltos)  partes.push(pl(devueltos, 'se devolvió 1', 'se devolvieron %'));
  if (cancelados) partes.push(pl(cancelados, 'se canceló 1', 'se cancelaron %'));
  if (aNovedad)   partes.push(pl(aNovedad, '1 entró en novedad', '% entraron en novedad'));

  if (!partes.length && !pendientes) return '';

  let msg = '\n\nDesde la carga anterior: ';
  msg += partes.length ? partes.join(', ') + '.' : 'ningún pedido cambió de estado.';
  if (pendientes) {
    msg += '\n' + pl(pendientes, 'Queda 1 pedido sin desenlace.',
                     'Quedan % pedidos sin desenlace.');
  }
  return msg;
}

/** Convierte una fila normalizada en una fila lista para la entidad. */
/**
 * Las columnas que son fechas, en cualquier hoja.
 *
 * Todas se guardan en AAAA-MM-DD. Lo que llega del archivo puede venir
 * como 12-07-2026 o 07/12/2026 según la plataforma y el país, y esas dos
 * cosas se ven iguales: una es 12 de julio y la otra 7 de diciembre.
 * Mientras se guarde así, cada pantalla tiene que volver a adivinar, y
 * el navegador adivina a la americana —mes primero— sin avisar.
 */
const CAMPOS_FECHA = ['fecha', 'fecha_entrega', 'fecha_promesa', 'fecha_ingreso',
                      'fecha_solucion', 'fecha_fin', 'ultimo_movimiento',
                      'actualizado', 'creado_en', 'ultimo_conteo'];

function prepararFila(f, tipo, fuenteId, tienda, pais, ss, aprendidos, vistos) {
  const o = Object.assign({}, f);
  o.fuente = fuenteId;
  o.tienda = tienda;

  /**
   * La fecha se normaliza aquí, una vez, y ya nadie más tiene que dudar.
   *
   * Antes se guardaba tal como venía y cada lector la interpretaba por su
   * cuenta: el servidor con aISO —que sabe que en la región el día va
   * primero— y la pantalla con new Date(), que asume el formato de
   * Estados Unidos. Por eso un pedido del 12 de julio salía en pantalla
   * como 7 de diciembre.
   *
   * La hora se conserva cuando viene: el gráfico de la jornada y el
   * seguimiento de novedades la necesitan.
   */
  CAMPOS_FECHA.forEach(function (k) {
    if (o[k] === undefined || o[k] === '' || o[k] === null) return;
    const iso = aISO(o[k], 'UTC');
    if (!iso) return;
    const hora = String(o[k]).match(/\b(\d{1,2}:\d{2}(?::\d{2})?)\b/);
    o[k] = hora ? iso + ' ' + hora[1] : iso;
  });

  // El id es fuente + id externo: estable entre importaciones, y deja
  // ver de dónde salió cada fila sin abrir el export.
  const ext = String(o.id_externo || o.guia || '').trim();
  o.id_externo = ext;
  o.id = fuenteId + '-' + (ext || Utilities.getUuid().slice(0, 8));

  if (o.telefono) {
    o.telefono_norm = telefonoNorm(o.telefono, pais);
    // El número alterno suele estar escondido en la nota o la observación
    const libre = [o.nota, o.observacion, o.aclaracion].filter(String).join(' ');
    if (libre) {
      const t = extraerTelefonos(o.telefono, libre, pais);
      if (t.secundario) { o.telefono_2 = t.secundario; o.telefono_2_norm = t.secundario; }
    }
  }

  if (o.documento) {
    const d = partirDocumento(o.documento);
    o.cedula = d.numero;
    delete o.documento;
  }

  if (tipo === 'pedidos' || tipo === 'pedidos_secundario') {
    const r = estadoConOrigen(fuenteId, o.estado, aprendidos);
    o.estado_canonico = r.estado;
    if (vistos && r.clave) {
      const k = fuenteId + '|' + r.clave;
      if (!vistos[k]) {
        vistos[k] = { fuente: fuenteId, texto: r.clave,
                      estado: r.estado, origen: r.origen, n: 0 };
      }
      vistos[k].n++;
    }
  }
  if (tipo === 'novedades') {
    o.grupo = grupoNovedad(o.motivo, o.codigo);
    if (o.aclaracion && !o.nota) { o.nota = o.aclaracion; }
  }
  if (tipo === 'llamadas') {
    o.seg_conversado = aSegundos(o.seg_conversado);
    o.seg_espera     = aSegundos(o.seg_espera);
    o.seg_total      = aSegundos(o.seg_total);
    // IRIS no dice de qué tienda es la llamada: se cruza por teléfono
    o.tienda = '';
  }
  if (tipo === 'facturacion') {
    o.plataforma = fuenteId.replace('_facturacion', '');
    o.id = fuenteId + '-' + (o.id_externo || Utilities.getUuid().slice(0, 8));
    const mon = String(o.moneda_gasto || (FUENTES[fuenteId] || {}).moneda_default || '').toUpperCase();
    o.moneda_gasto = mon;
    const destino = monedaReporte(ss);
    o.moneda_reporte = destino;
    if (mon && o.fecha) {
      const c = convertir(ss, o.gasto, o.fecha, mon, destino);
      // Sin tasa no se inventa un número: queda vacío y visible
      o.gasto_normalizado = c.valor === null ? '' : c.valor;
    }
  }

  if (tipo === 'cartera') {
    o.id = fuenteId + '-' + (o.id_externo || Utilities.getUuid().slice(0, 8));
    o.tipo = String(o.tipo_movimiento || '').trim().toUpperCase();
    delete o.tipo_movimiento;
    o.clase = claseMovimiento(o.descripcion, o.concepto_retiro);
    o.importado_en = ahoraISO();
    /**
     * El signo se guarda en el monto, no en la cabeza de quien lee.
     *
     * Dropi manda todo positivo y dice aparte si fue ENTRADA o SALIDA.
     * Guardarlo así obliga a acordarse del signo cada vez que se suma, y
     * tarde o temprano alguien suma una devolución como ingreso.
     */
    if (o.tipo === 'SALIDA') o.monto = -Math.abs(num(o.monto));
    else o.monto = Math.abs(num(o.monto));
  }

  if (tipo === 'pauta') {
    o.plataforma = fuenteId;
    /**
     * Una fila de pauta no trae identificador propio, así que se arma uno
     * con lo que la hace única: plataforma, tienda, periodo y conjunto.
     *
     * El periodo entra entero —inicio y fin— porque Meta exporta el mismo
     * conjunto para rangos distintos. Sin el fin, volver a exportar con
     * otro rango pisaría la fila anterior y el gasto del mes cambiaría
     * solo, sin que nadie hubiera tocado nada.
     */
    o.id = [fuenteId, tienda, o.fecha || '', o.fecha_fin || '',
            norm(o.conjunto || o.campana || '')].join('-')
           .replace(/\s+/g, '_').slice(0, 180);
    const mon = String(o.moneda_gasto || (FUENTES[fuenteId] || {}).moneda_default || '').toUpperCase();
    o.moneda_gasto = mon;
    const destino = monedaReporte(ss);
    if (mon && o.fecha) {
      const c = convertir(ss, o.gasto, o.fecha, mon, destino);
      // Sin tasa no se inventa un número: queda vacío y visible
      o.gasto_normalizado = c.valor === null ? '' : c.valor;
    }
  }
  return o;
}

/** "00:01:23" → 83. Las llamadas vienen en hh:mm:ss. */
function aSegundos(v) {
  if (v === '' || v == null) return '';
  const s = String(v).trim();
  const m = s.match(/^(\d+):(\d{2}):(\d{2})$/);
  if (m) return (+m[1]) * 3600 + (+m[2]) * 60 + (+m[3]);
  const n = aNumero(s);
  return n === '' ? '' : n;
}

/**
 * Escribe las filas en su hoja: actualiza las que ya existen, agrega las
 * nuevas. Nunca borra, y nunca pisa las columnas del equipo.
 */
/**
 * Escribe las filas en su entidad, por bloques.
 *
 * Antes escribía celda por celda. Era lo más seguro sobre el papel —una
 * hoja no tiene bloqueo de fila, así que dos escrituras simultáneas se
 * pisan— pero con 206 pedidos y ocho columnas eran mil seiscientas
 * llamadas, una por una, y la importación tardaba más de un minuto.
 * Google corta la respuesta antes y el navegador solo ve un 404: el
 * archivo entraba a medias y parecía que había fallado la subida.
 *
 * Ahora se arma todo en memoria y se escribe por tramos de columnas
 * seguidas. Las columnas del equipo se quedan fuera del tramo, así que
 * siguen intocables; las de la plataforma solo las escribe esto, de modo
 * que devolver su valor actual a una fila que no cambió no pisa nada.
 * Y un candado impide que dos importaciones corran encima.
 */
function escribirFilas(ss, hoja, filas, fuenteId) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('No existe la hoja ' + hoja);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Hay otra importación corriendo. Espera a que termine.');
  }

  try {
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const cId = enc.indexOf('id');
    if (cId === -1) throw new Error(hoja + ' no tiene columna id.');

    // Lo que ya está, indexado, para no recorrer la hoja por cada fila
    let datos = [];
    const existentes = {};
    if (sh.getLastRow() > 1) {
      datos = sh.getRange(2, 1, sh.getLastRow() - 1, enc.length).getValues();
      datos.forEach(function (f, i) {
        const k = String(f[cId]).trim();
        if (k) existentes[k] = i;
      });
    }

    const nuevas = [], sinEstado = {}, tocadas = {};
    let actualizadas = 0, iguales = 0;

    filas.forEach(function (o) {
      if (o.estado_canonico === ESTADOS.SIN_CLASIFICAR) {
        sinEstado[norm(o.estado) || '(vacio)'] = 1;
      }

      const i = existentes[o.id];
      if (i === undefined) {
        nuevas.push(enc.map(function (col) {
          return o[col] !== undefined ? o[col] : '';
        }));
        return;
      }

      let cambio = false;
      enc.forEach(function (col, c) {
        if (COLUMNAS_DEL_EQUIPO.indexOf(col) !== -1) return;
        if (o[col] === undefined || o[col] === '') return;
        if (String(datos[i][c]) === String(o[col])) return;
        datos[i][c] = o[col];
        tocadas[c] = true;
        cambio = true;
      });
      if (cambio) actualizadas++; else iguales++;
    });

    // Tramos de columnas seguidas que hay que reescribir
    if (datos.length && Object.keys(tocadas).length) {
      const cols = Object.keys(tocadas).map(Number).sort(function (a, b) { return a - b; });
      let ini = cols[0], fin = cols[0];
      const escribirTramo = function (a, b) {
        const ancho = b - a + 1;
        const sub = datos.map(function (f) { return f.slice(a, b + 1); });
        sh.getRange(2, a + 1, datos.length, ancho).setValues(sub);
      };
      for (let k = 1; k < cols.length; k++) {
        if (cols[k] === fin + 1) { fin = cols[k]; continue; }
        escribirTramo(ini, fin);
        ini = fin = cols[k];
      }
      escribirTramo(ini, fin);
    }

    if (nuevas.length) {
      sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, enc.length).setValues(nuevas);
    }

    return {
      nuevas: nuevas.length, actualizadas: actualizadas, iguales: iguales,
      sinEstado: Object.keys(sinEstado),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dropi trae la novedad dentro de la misma fila del pedido, no en un
 * reporte aparte como Effi. Se extraen para que la alarma de patrón y
 * el cierre de mes puedan contarlas.
 */
/**
 * Saca las novedades que vienen dentro del export de pedidos.
 *
 * El estado de la novedad no es el estado del pedido, y confundirlos fue
 * un error caro: antes una novedad solo contaba como resuelta si el
 * pedido estaba EN ESE MOMENTO en "novedad solucionada". Pero un pedido
 * que tuvo novedad y después se entregó ya no está en ese estado, está
 * entregado — así que su novedad quedaba abierta para siempre. Setenta y
 * dos novedades y setenta y dos "abiertas", cuando de verdad quedaban dos.
 *
 * Dropi ya trae la respuesta en sus propias columnas, y ahora se leen.
 * Son tres hechos distintos y cada uno tiene su columna:
 *
 *   solucionada  lo que dice la plataforma: SI o NO
 *   desenlace    cómo terminó el pedido: entregado, devuelto, cancelado
 *   estado       qué hay que hacer hoy con ella
 *
 * Separarlos deja ver el caso que importa: la novedad que el equipo SÍ
 * resolvió y el pedido se devolvió igual. Eso es trabajo que no se
 * convirtió en venta, y mezclado con lo demás no se ve.
 */
function derivarNovedades(pedidos, fuenteId, tienda) {
  const TERMINALES = ['entregado', 'devolucion', 'cancelado'];

  return pedidos
    .filter(function (p) {
      const m = String(p.motivo_novedad || '').trim();
      return m && m !== '.' && m !== '-';
    })
    .map(function (p) {
      const sol = norm(p.solucionada || '');
      const solucionada = sol === 'si' || sol === 'sí' || sol === 'true' || sol === '1';
      const cerrado = TERMINALES.indexOf(p.estado_canonico) !== -1;

      let estado;
      if (solucionada || p.estado_canonico === 'novedad_resuelta') estado = 'resuelta';
      else if (cerrado) estado = 'cerrada';   // terminó sin resolverse
      else estado = 'abierta';                // sigue esperando a alguien

      return {
        id: fuenteId + '-nov-' + p.id_externo,
        fuente: fuenteId,
        id_externo: p.id_externo,
        pedido_id: p.id,
        fecha: p.ultimo_movimiento || p.fecha,
        tipo: 'novedad',
        motivo: p.motivo_novedad,
        grupo: grupoNovedad(p.motivo_novedad),
        estado: estado,
        solucionada: solucionada ? 'si' : 'no',
        fecha_solucion: p.fecha_solucion || '',
        desenlace: cerrado ? p.estado_canonico : '',
      };
    });
}

/** Deja constancia de cuándo se importó y cuántas filas entraron. */
function registrarImportacion(ss, tienda, fuenteId, filas) {
  const sh = ss.getSheetByName('Fuentes');
  if (!sh || sh.getLastRow() < 2) return;
  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cT = enc.indexOf('tienda'), cF = enc.indexOf('fuente');
  const cU = enc.indexOf('ultima_importacion'), cN = enc.indexOf('filas_ultima');

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cT]).trim() !== tienda) continue;
    if (norm(datos[i][cF]) !== norm(fuenteId.split('_')[0])) continue;
    if (cU !== -1) sh.getRange(i + 1, cU + 1).setValue(ahoraISO());
    if (cN !== -1) sh.getRange(i + 1, cN + 1).setValue(filas);
    return;
  }
}

function paisDeTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return '';
  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cId = enc.indexOf('id'), cP = enc.indexOf('pais');
  const ISO = { colombia:'CO', ecuador:'EC', guatemala:'GT', mexico:'MX',
                peru:'PE', chile:'CL', argentina:'AR', bolivia:'BO',
                paraguay:'PY', uruguay:'UY', venezuela:'VE', panama:'PA',
                'costa rica':'CR', honduras:'HN', nicaragua:'NI',
                'el salvador':'SV', brasil:'BR', espana:'ES' };
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cId]).trim() === tienda) {
      return ISO[norm(datos[i][cP])] || '';
    }
  }
  return '';
}

/**
 * Cruza las llamadas de IRIS con los pedidos, por teléfono normalizado.
 * Córrelo después de importar IRIS y los pedidos.
 *
 * Una llamada que no cruza con ningún pedido queda con tienda y pedido_id
 * vacíos — visible, no descartada. Son llamadas a números que no están en
 * ningún pedido, y vale la pena mirarlas.
 */
function cruzarLlamadas(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const shL = ss.getSheetByName('Llamadas');
  const shP = ss.getSheetByName('Pedidos');
  if (!shL || shL.getLastRow() < 2) { Logger.log('No hay llamadas.'); return 'Sin llamadas.'; }
  if (!shP || shP.getLastRow() < 2) { Logger.log('No hay pedidos.'); return 'Sin pedidos.'; }

  const dP = shP.getDataRange().getValues();
  const eP = dP[0].map(norm);
  const cTel = eP.indexOf('telefono_norm'), cPid = eP.indexOf('id'), cPt = eP.indexOf('tienda');

  // Un teléfono puede tener varios pedidos: se queda el más reciente
  const porTel = {};
  const cFe = eP.indexOf('fecha');
  for (let i = 1; i < dP.length; i++) {
    const t = String(dP[i][cTel] || '').trim();
    if (!t) continue;
    const fe = aISO(dP[i][cFe], 'UTC') || '';
    if (!porTel[t] || fe > porTel[t].fecha) {
      porTel[t] = { id: dP[i][cPid], tienda: dP[i][cPt], fecha: fe };
    }
  }

  const dL = shL.getDataRange().getValues();
  const eL = dL[0].map(norm);
  const cLtel = eL.indexOf('telefono_norm'), cLpid = eL.indexOf('pedido_id'),
        cLt = eL.indexOf('tienda');

  let cruzadas = 0, huerfanas = 0;
  for (let i = 1; i < dL.length; i++) {
    const t = String(dL[i][cLtel] || '').trim();
    const m = t && porTel[t];
    if (m) {
      if (String(dL[i][cLpid]) !== String(m.id)) {
        shL.getRange(i + 1, cLpid + 1).setValue(m.id);
        shL.getRange(i + 1, cLt + 1).setValue(m.tienda);
      }
      cruzadas++;
    } else { huerfanas++; }
  }

  const msg = 'Llamadas cruzadas: ' + cruzadas + '\n' +
    'Sin pedido que las reciba: ' + huerfanas +
    (huerfanas ? '\n  (números que no están en ningún pedido — vale la pena revisarlos)' : '');
  Logger.log(msg);
  return msg;
}

/**
 * Importa todo lo que esté pegado y tenga fuente activa. Es lo que
 * conviene correr después de pegar los exports del día.
 */
function importarTodo(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const sh = ss.getSheetByName('Fuentes');
  if (!sh || sh.getLastRow() < 2) return 'No hay fuentes configuradas.';

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cT = enc.indexOf('tienda'), cF = enc.indexOf('fuente'), cA = enc.indexOf('activa');

  // Cuántas tiendas activas usa cada plataforma. Si son dos y comparten
  // una sola pestaña, importar las dos etiquetaría las mismas filas
  // primero con una tienda y luego con la otra: los datos quedarían
  // asignados a la que corrió de último, sin que nadie lo note.
  const porFuente = {};
  datos.slice(1).forEach(function (f) {
    if (norm(f[cA]) !== 'si') return;
    const fu = String(f[cF]).trim();
    (porFuente[fu] = porFuente[fu] || []).push(String(f[cT]).trim());
  });

  const log = [];
  datos.slice(1).forEach(function (f) {
    if (norm(f[cA]) !== 'si') return;
    const fuente = String(f[cF]).trim(), tienda = String(f[cT]).trim();

    if (porFuente[fuente].length > 1) {
      const base = '_Import_' + fuente.charAt(0).toUpperCase() + fuente.slice(1);
      if (!ss.getSheetByName(base + '_' + tienda.toUpperCase())) {
        log.push('SALTADA: ' + fuente + ' / ' + tienda + '\n' +
          '  La usan ' + porFuente[fuente].length + ' tiendas (' +
          porFuente[fuente].join(', ') + ') y solo hay una pestaña ' + base + '.\n' +
          '  Importarlas juntas etiquetaría las mismas filas con la tienda\n' +
          '  equivocada. Elige una salida:\n' +
          '    a) Crea una pestaña por tienda: ' + base + '_GT y ' + base + '_EC\n' +
          '    b) O pon "no" en la columna activa de la tienda que no la usa');
        return;
      }
    }

    // Effi son dos reportes con nombres propios
    const ids = fuente === 'effi' ? ['effi_guias', 'effi_novedades'] : [fuente];
    ids.forEach(function (id) {
      try { log.push(importar(id, tienda, cliente)); }
      catch (err) { log.push('· ' + id + '/' + tienda + ': ' + err.message); }
    });
  });

  try { log.push(cruzarLlamadas(cliente)); } catch (err) { /* sin llamadas */ }

  const msg = log.join('\n\n');
  Logger.log(msg);
  return msg;
}


// ─── ATAJOS ──────────────────────────────────────────────────
/**
 * El botón Ejecutar de Apps Script no permite pasar argumentos, así que
 * para importar una sola fuente hace falta una función sin parámetros.
 *
 * No se listan por tienda a propósito: hacerlo ataría el código a las
 * tiendas de una cuenta, y Nova se vende a clientes de toda la región.
 * importarTodo() recorre la hoja Fuentes, que es donde vive esa lista.
 *
 * Si necesitas importar una fuente suelta, escribe la llamada en la
 * consola del editor:  importar('dropi', 'lima')
 */
function importarTodoAhora() { return importarTodo(); }


/* ═══════════════════════════════════════════════════════════════
   12 · META
   ═══════════════════════════════════════════════════════════════ */

/**
 * ═══════════════════════════════════════════════════════════
 *  META · LA PAUTA ENTRA SOLA
 * ═══════════════════════════════════════════════════════════
 *
 * Hasta hoy, para que Nova supiera cuánto se gastó en anuncios alguien
 * tenía que entrar a Meta, exportar un Excel y subirlo. Todos los días.
 * Y ese archivo trae fechas ambiguas, monedas mezcladas y una fila por
 * periodo en vez de por día — tres formas de equivocarse que no existen
 * si los datos se piden directamente.
 *
 * DÓNDE VIVE LA LLAVE
 *
 * En las Propiedades del Script, nunca en una hoja. Una llave en una
 * celda la ve cualquiera a quien le compartan el archivo, y la de Meta
 * no caduca: quien la copie la tiene para siempre. En las Propiedades no
 * la ve ni quien abra la hoja ni quien la descargue.
 *
 * Un mismo script atiende a todos los clientes, así que la clave lleva
 * el identificador de la hoja de cada uno: la llave de una tienda no
 * puede quedar al alcance de otra.
 *
 * LO QUE NOVA PUEDE HACER CON ELLA
 *
 * Leer. Nada más. El permiso que se pide es `ads_read` y el activo se
 * asigna como "ver rendimiento", así que ni esta función ni ninguna otra
 * puede crear un anuncio, pausarlo ni mover un presupuesto. Si esa llave
 * se filtrara, lo peor que alguien podría hacer es enterarse del gasto.
 */

const META_API = 'https://graph.facebook.com/v21.0/';

/** La llave de ESTE cliente. Nunca una global. */
function metaClave(s) { return 'META_TOKEN_' + String(s.sheetId || '').slice(0, 44); }

function metaToken(s) {
  return PropertiesService.getScriptProperties().getProperty(metaClave(s)) || '';
}

/**
 * Qué cuenta publicitaria le toca a cada tienda.
 * Vive en Parametros, que es donde ya viven los ajustes de la tienda.
 */
function metaCuenta(ss, tienda) {
  return String(ajustes(ss, tienda).meta_cuenta || '').replace(/^act_/, '').trim();
}

/**
 * El estado de la conexión, SIN devolver la llave.
 *
 * Nunca se manda de vuelta a la pantalla, ni recortada. Una llave que
 * viaja al navegador queda en la memoria del navegador, y de ahí a una
 * captura de pantalla hay un paso. Se dice si existe y cuándo se guardó;
 * para cambiarla se escribe una nueva.
 */
function apiMetaEstado(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña ve la conexión con Meta.' };
  const ss = SpreadsheetApp.openById(s.sheetId);
  const props = PropertiesService.getScriptProperties();
  /**
   * La última prueba que salió bien.
   *
   * Sin esto, lo único que decía si la conexión servía era el mensaje
   * que aparecía justo al probarla — y se perdía al recargar. Quien
   * volvía al día siguiente veía "hay una llave guardada" y no tenía
   * forma de saber si esa llave funcionaba.
   */
  let prueba = null;
  try {
    const crudo = props.getProperty(metaClave(s) + '_PRUEBA');
    if (crudo) prueba = JSON.parse(crudo);
  } catch (e) { prueba = null; }

  return {
    ok: true,
    hayLlave: !!metaToken(s),
    ultimaPrueba: prueba,
    guardadaEn: props.getProperty(metaClave(s) + '_FECHA') || '',
    cuentas: s.tiendas.map(function (t) {
      return { tienda: t, cuenta: metaCuenta(ss, t), moneda: monedaDeTienda(ss, t) };
    }),
  };
}

/** Guardar la llave y/o el número de cuenta de una tienda. */
function apiMetaGuardar(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña conecta Meta.' };
  const props = PropertiesService.getScriptProperties();

  /**
   * Se llama `llave` y no `token` a propósito.
   *
   * `token` es el de la sesión de Nova y viaja en cada petición. Cuando
   * este campo se llamaba igual, la llave de Meta ocupaba su lugar al
   * enviar: el servidor no reconocía la sesión y respondía cerrándola.
   * Guardar la llave expulsaba a quien la estaba guardando.
   */
  if (p.llave !== undefined) {
    const t = String(p.llave).trim();
    if (t && t.length < 50) {
      return { ok: false, error: 'Esa llave se ve incompleta. Las de Meta pasan de ' +
               'doscientos caracteres — puede que se haya cortado al copiarla.' };
    }
    if (t) {
      props.setProperty(metaClave(s), t);
      props.setProperty(metaClave(s) + '_FECHA', ahoraISO());
      // La prueba anterior era de la llave anterior. Dejarla puesta diría
      // "conectada" sobre una llave que nadie ha comprobado todavía.
      props.deleteProperty(metaClave(s) + '_PRUEBA');
    } else {
      props.deleteProperty(metaClave(s));
      props.deleteProperty(metaClave(s) + '_FECHA');
      props.deleteProperty(metaClave(s) + '_PRUEBA');
    }
    registrarMovimiento(s, 'Meta', 'llave', 'token', '(oculto)', t ? '(guardada)' : '(borrada)');
  }

  if (p.cuenta !== undefined && p.tienda) {
    if (s.tiendas.indexOf(p.tienda) === -1) {
      return { ok: false, error: 'No tienes acceso a esa tienda.' };
    }
    // Se escribe por la misma puerta que el resto de ajustes: valida,
    // registra el cambio y no acepta claves fuera del catálogo.
    const cta = String(p.cuenta).replace(/^act_/, '').replace(/\D/g, '');
    const rp = apiParametros(s, { tienda: p.tienda, cambios: { meta_cuenta: cta } });
    if (!rp.ok) return rp;
  }

  return apiMetaEstado(s, p);
}

/**
 * Probar la conexión y decir QUÉ falla.
 *
 * "No se pudo conectar" no sirve: quien lo lee no sabe si volver a Meta,
 * revisar el número de cuenta o esperar. Meta devuelve códigos distintos
 * para la llave vencida, el permiso que falta y la cuenta que no existe,
 * y cada uno tiene una respuesta distinta.
 */
function apiMetaProbar(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña prueba la conexión.' };

  const token = metaToken(s);
  if (!token) return { ok: false, error: 'Todavía no has guardado la llave.' };

  const ss = SpreadsheetApp.openById(s.sheetId);
  const tienda = String(p.tienda || s.tiendas[0]);
  const cuenta = metaCuenta(ss, tienda);
  if (!cuenta) {
    return { ok: false, error: 'Falta el número de cuenta publicitaria de esta tienda. ' +
             'Está en business.facebook.com, junto al nombre de la cuenta.' };
  }

  const url = META_API + 'act_' + cuenta +
    '?fields=name,currency,account_status,amount_spent' +
    '&access_token=' + encodeURIComponent(token);

  let r;
  try {
    r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (err) {
    return { ok: false, error: 'No se pudo hablar con Meta: ' + err.message };
  }

  let j;
  try { j = JSON.parse(r.getContentText()); }
  catch (err) { return { ok: false, error: 'Meta contestó algo que no entiendo.' }; }

  if (j.error) return Object.assign({ ok: false }, explicarErrorMeta(j.error));

  const monedaTienda = monedaDeTienda(ss, tienda);

  PropertiesService.getScriptProperties().setProperty(
    metaClave(s) + '_PRUEBA',
    JSON.stringify({ cuando: ahoraISO(), cuenta: cuenta,
                     nombre: j.name || '', moneda: j.currency || '', tienda: tienda }));

  return {
    ok: true, tienda: tienda, cuenta: cuenta,
    nombre: j.name || '', moneda: j.currency || '',
    // Si Meta cobra en una moneda y la tienda factura en otra, el gasto
    // no se puede sumar a las ventas sin convertirlo. Mejor decirlo aquí
    // que dejar que aparezca un margen absurdo dentro de tres semanas.
    avisoMoneda: (j.currency && monedaTienda && j.currency !== monedaTienda)
      ? 'Meta te cobra en ' + j.currency + ' y esta tienda factura en ' +
        monedaTienda + '. Nova va a convertir con la tasa del día de cada gasto.'
      : '',
  };
}

/** Traducir el error de Meta a algo que diga qué hacer. */
function explicarErrorMeta(e) {
  const cod = e.code, sub = e.error_subcode;

  if (cod === 190) {
    return { error: 'La llave ya no sirve.\n\n' +
      (sub === 463 ? 'Venció. ' : 'Puede que la hayan revocado, o que lleve 180 días sin usarse. ') +
      'Genera otra en Meta: Usuarios del sistema → Nova → Generar identificador.',
      accion: 'llave' };
  }
  if (cod === 200 || cod === 10) {
    return { error: 'La llave funciona, pero no tiene permiso para leer esta cuenta.\n\n' +
      'Revisa dos cosas en Meta: que la cuenta publicitaria esté asignada al usuario ' +
      'de sistema Nova con "ver rendimiento", y que la llave se haya generado con ' +
      'el permiso ads_read.', accion: 'permiso' };
  }
  if (cod === 100) {
    return { error: 'Meta no encuentra esa cuenta publicitaria.\n\n' +
      'Revisa el número: son solo dígitos, sin el "act_" adelante.', accion: 'cuenta' };
  }
  if (cod === 17 || cod === 4 || cod === 613) {
    return { error: 'Meta está limitando las consultas ahora mismo. ' +
      'Espera unos minutos y vuelve a intentar.', accion: 'esperar' };
  }
  return { error: 'Meta dijo: ' + (e.message || 'error ' + cod), accion: '' };
}


// ─── LO QUE CORRE SOLO ───────────────────────────────────────

/**
 * ═══════════════════════════════════════════════════════════
 *  UN SOLO INTERRUPTOR PARA TODO LO AUTOMÁTICO
 * ═══════════════════════════════════════════════════════════
 *
 * Nova tiene trabajos que deberían correr sin que nadie los pida: bajar
 * las tasas de cambio del día y revisar las alarmas. El código estaba,
 * pero cada uno se prendía con su propia función y había que acordarse de
 * las dos. Nadie se acuerda de dos cosas que se hacen una sola vez.
 *
 * El resultado era el peor posible: parecía automático y no lo era.
 * Alguien conectaba Meta, veía la pauta llegar en pesos contra una tienda
 * en dólares, y se quedaba esperando una conversión que nunca iba a pasar
 * porque la hoja Tasas estaba vacía — sin un error, sin una pista.
 *
 * Ahora es un interruptor. Se puede accionar desde el Apps Script
 * (`prenderAutomatico`) o desde la consola, con un botón — y se puede
 * comprobar por los dos lados, porque un automatismo que no se puede
 * comprobar es un automatismo en el que no se puede confiar.
 *
 * Una advertencia que vale por toda esta sección: los disparadores son
 * del PROYECTO, no de cada cliente. Un solo Apps Script atiende a todos,
 * así que se prenden una vez y sirven para todos. Por eso el botón vive
 * en Nova Central y no en la pantalla de la dueña: no es una decisión de
 * cada cliente, y ponerlo ahí le daría a cualquiera un interruptor que
 * afecta a los demás.
 * ═══════════════════════════════════════════════════════════
 */

/**
 * Los dos trabajos que Nova tiene que correr sola, y sus horas.
 *
 * Están en un solo sitio para que prender, revisar y mostrar hablen de
 * lo mismo. Cuando cada función tenía su propia lista, prender uno y
 * revisar otro era cuestión de tiempo.
 */
const TRABAJOS = [
  {
    fn: 'actualizarTasasDiario',
    nombre: 'Tasas de cambio',
    hora: 5,
    // Va primero a propósito: sin tasas el gasto de pauta en otra moneda
    // no se suma, así que las alarmas de CPA y margen estarían juzgando
    // una operación a la que le falta el gasto.
    porque: 'Sin esto, la pauta que Meta cobra en otra moneda no se puede sumar.',
  },
  {
    fn: 'leerMetaDiario',
    nombre: 'Lectura de Meta',
    hora: 6,
    porque: 'Sin esto, el gasto de anuncios solo entra si alguien sube el Excel.',
  },
  {
    fn: 'revisarAlarmasTodos',
    nombre: 'Revisión de alarmas',
    hora: 7,
    // De último: juzga el día con el gasto ya adentro y ya convertido.
    porque: 'Sin esto, las alarmas solo se calculan cuando alguien abre Nova.',
  },
  {
    fn: 'semaforoLunes',
    nombre: 'Semáforo semanal',
    hora: 8,
    // El único semanal. Va el lunes y a las 8, después de que las tasas,
    // Meta y las alarmas de esa misma mañana ya corrieron: si saliera
    // antes, juzgaría la semana con la pauta del viernes.
    dia: 'MONDAY',
    porque: 'Sin esto, el cierre de la semana solo existe si alguien lo pide.',
  },
];

/** Cómo se dice la frecuencia de un trabajo, en castellano. */
function cuandoCorre_(t) {
  return t.dia === 'MONDAY' ? 'los lunes a las ' + t.hora + ':00'
                            : 'todos los días a las ' + t.hora + ':00';
}

/** Qué disparadores hay puestos ahora mismo, por función. */
function trabajosPuestos_() {
  const puestos = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    puestos[t.getHandlerFunction()] = true;
  });
  return puestos;
}

/**
 * El orden de las horas importa, y por eso un disparador viejo se rehace.
 *
 * Las tasas tienen que estar antes de que entre el gasto de Meta, y las
 * alarmas después, para que juzguen el día con el gasto ya adentro. Si
 * una versión anterior dejó un disparador a otra hora, mantenerlo sería
 * dejar el orden al azar — así que se borra y se vuelve a crear.
 *
 * Google no dice a qué hora quedó un disparador (`atHour` define una
 * franja de una hora, no un minuto exacto), así que la hora se guarda
 * aparte, en las Propiedades. Sin eso no habría forma de saber si el que
 * está puesto es el de ahora o el de antes.
 */
const PROP_HORAS = 'NOVA_HORAS_TRABAJOS';

function horasGuardadas_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties()
      .getProperty(PROP_HORAS) || '{}');
  } catch (e) { return {}; }
}

/**
 * Prende los dos trabajos. No toca las tasas: eso se hace aparte.
 *
 * Separado a propósito. Instalar un disparador es instantáneo; bajar
 * noventa días de tasas de GOOGLEFINANCE tarda. Si fueran la misma
 * llamada, el botón de la consola se quedaría colgado hasta que
 * terminara lo lento, y quien lo aprieta no sabría si funcionó.
 */
function prenderTrabajos_() {
  const props = PropertiesService.getScriptProperties();
  const horas = horasGuardadas_();
  const disparadores = ScriptApp.getProjectTriggers();

  const hechos = TRABAJOS.map(function (t) {
    const mios = disparadores.filter(function (d) {
      return d.getHandlerFunction() === t.fn;
    });
    const correcta = horas[t.fn] === t.hora;

    if (mios.length === 1 && correcta) {
      return { fn: t.fn, nombre: t.nombre, ya: true, hora: t.hora,
               cuando: cuandoCorre_(t) };
    }
    // Sobrantes o a la hora equivocada: se rehace. Dos disparadores de la
    // misma función leerían Meta dos veces la misma mañana.
    mios.forEach(function (d) { try { ScriptApp.deleteTrigger(d); } catch (e) {} });
    const b = ScriptApp.newTrigger(t.fn).timeBased().atHour(t.hora);
    if (t.dia) b.onWeekDay(ScriptApp.WeekDay[t.dia]).create();
    else b.everyDays(1).create();
    horas[t.fn] = t.hora;
    return { fn: t.fn, nombre: t.nombre, ya: false, hora: t.hora,
             cuando: cuandoCorre_(t), rehecho: mios.length > 0 };
  });

  props.setProperty(PROP_HORAS, JSON.stringify(horas));
  return hechos;
}

/**
 * Las tasas de los últimos noventa días de cada cliente, ahora.
 *
 * El disparador diario solo tapa los huecos de aquí en adelante. Sin
 * esta primera carga, el gasto de pauta de los meses pasados se quedaría
 * sin convertir para siempre — y los cierres que ya se hicieron no
 * tendrían con qué cuadrar.
 *
 * Esta función es también la que corre el disparador de un solo uso que
 * arma la consola, así que no puede recibir parámetros ni depender de
 * una sesión.
 */
function cargaInicialTasas() {
  const hechos = [];
  let cargadas = 0;
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try { actualizarTasas(c.sheetId, 90); cargadas++; }
    catch (e) { hechos.push('OJO · no pude cargar las tasas de ' + c.empresa + ': ' + e.message); }
  });
  hechos.unshift('Tasas de los últimos 90 días cargadas en ' + cargadas + ' cuenta(s).');

  // Y se borra el disparador de un solo uso que la trajo hasta aquí, si
  // lo hubo. Un disparador "after" que nadie limpia se queda ocupando
  // una de las veinte ranuras que da Google, para siempre.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cargaInicialTasas' &&
        t.getEventType() === ScriptApp.EventType.CLOCK) {
      try { ScriptApp.deleteTrigger(t); } catch (e) {}
    }
  });

  const msg = hechos.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Prender lo automático, desde el Apps Script.
 *
 * La consola hace lo mismo con un botón (`centralPrenderAutomatico`).
 * Esta versión existe para cuando la consola todavía no está montada, o
 * cuando hay que arreglar algo sin depender de que el sitio cargue.
 */
function prenderAutomatico() {
  const hechos = prenderTrabajos_().map(function (r) {
    if (r.ya) return '· ' + r.nombre + ': ya estaba corriendo.';
    return '· ' + r.nombre + ': ' + (r.rehecho ? 'reprogramado' : 'prendido') +
           ', ' + r.cuando + '.';
  });
  hechos.push('· ' + cargaInicialTasas());

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + hechos.join('\n') +
    '\n\nPara comprobarlo cuando quieras, corre verAutomatico().';
  Logger.log(msg);
  return msg;
}

// ─── COMPROBAR ───────────────────────────────────────────────

/**
 * El estado de lo automático, en datos.
 *
 * Devuelve objetos y no texto porque lo consume la consola. `verAutomatico()`
 * es esto mismo escrito para leer en el Apps Script — una sola fuente,
 * dos formas de mirarla, para que no puedan contradecirse.
 */
function estadoAutomatico() {
  const puestos = trabajosPuestos_();
  const horas = horasGuardadas_();
  const trabajos = TRABAJOS.map(function (t) {
    // "Puesto a otra hora" cuenta como apagado: el orden entre los tres
    // es lo que hace que el gasto entre convertido y las alarmas lo vean.
    return { nombre: t.nombre, hora: t.hora, porque: t.porque,
             cuando: cuandoCorre_(t),
             prendido: !!puestos[t.fn] && horas[t.fn] === t.hora };
  });

  const clientes = [];
  let error = '';
  try {
    listarClientes().forEach(function (c) {
      if (!c.sheetId) return;
      clientes.push(estadoTasasDe_(c.sheetId, c.empresa));
    });
  } catch (e) {
    error = e.message;
  }

  return {
    trabajos: trabajos,
    todoPrendido: trabajos.every(function (t) { return t.prendido; }),
    clientes: clientes,
    // Un cliente "en falta" es uno que necesita tasas y no las tiene al
    // día. Es el número que decide si la consola avisa o se calla.
    enFalta: clientes.filter(function (c) { return c.necesita && !c.alDia; }).length,
    error: error,
  };
}

/**
 * El estado de las tasas de una cuenta.
 *
 * Distingue tres cosas que se confunden: que la hoja esté vacía, que
 * esté vieja, y que no haga falta. Una tienda que factura en la misma
 * moneda en que le cobran no necesita ninguna tasa, y decirle que le
 * "faltan" sería mandarla a arreglar algo que no está roto.
 */
function estadoTasasDe_(sheetId, nombre) {
  const out = { empresa: nombre, necesita: false, alDia: true, pares: [], error: '' };
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const pares = paresNecesarios(ss);
    if (!pares.length) return out;          // no necesita: alDia se queda en true
    out.necesita = true;

    /**
     * Par por par, no la hoja entera.
     *
     * Mirar solo la fecha más reciente de toda la hoja diría "al día"
     * mientras a un par le falta hasta la primera fila: basta con que
     * otro par se esté actualizando bien para tapar el hueco.
     */
    const ultimaDe = {};
    const sh = ss.getSheetByName('Tasas');
    if (sh && sh.getLastRow() > 1) {
      sh.getDataRange().getValues().slice(1).forEach(function (f) {
        const fecha = aISO(f[0], 'UTC');
        if (!fecha) return;
        const o = String(f[1] || '').toUpperCase();
        const d = String(f[2] || '').toUpperCase();
        [o + '|' + d, d + '|' + o].forEach(function (k) {   // el inverso sirve igual
          if (!ultimaDe[k] || fecha > ultimaDe[k]) ultimaDe[k] = fecha;
        });
      });
    }

    const hoy = ahoraISO().slice(0, 10);
    out.pares = pares.map(function (p) {
      const ultima = ultimaDe[p.origen + '|' + p.destino] || '';
      const dias = ultima
        ? Math.floor((new Date(hoy + 'T00:00:00Z') - new Date(ultima + 'T00:00:00Z')) / 86400000)
        : -1;
      const bien = dias >= 0 && dias <= 3;
      if (!bien) out.alDia = false;
      return {
        par: p.origen + '→' + p.destino,
        ultima: ultima, dias: dias, bien: bien,
      };
    });
  } catch (e) {
    out.error = e.message;
    out.alDia = false;
  }
  return out;
}

/**
 * Lo mismo, escrito para leer en el Apps Script.
 *
 * Existe porque "instalado" y "funcionando" no son lo mismo. Un
 * disparador puede estar puesto y fallar todos los días en silencio —
 * Google lo reintenta, no avisa, y la hoja se queda vieja sin que nadie
 * lo note.
 */
function verAutomatico() {
  const e = estadoAutomatico();

  const lineas = e.trabajos.map(function (t) {
    return (t.prendido ? '✓ ' : '✗ ') + t.nombre +
           (t.prendido ? ' (' + t.cuando + ')' : '  ← APAGADO. ' + t.porque);
  });

  lineas.push('');
  if (e.error) {
    lineas.push('No pude revisar las tasas: ' + e.error);
  } else if (!e.clientes.length) {
    lineas.push('No hay clientes registrados todavía, así que no hay tasas que revisar.');
  }

  e.clientes.forEach(function (c) {
    if (c.error) { lineas.push('· ' + c.empresa + ': no pude revisar — ' + c.error); return; }
    if (!c.necesita) {
      lineas.push('· ' + c.empresa + ': no necesita tasas — factura y le cobran en la misma moneda.');
      return;
    }
    const detalle = c.pares.map(function (p) {
      if (p.dias < 0) return p.par + ' SIN NINGUNA TASA';
      return p.par + (p.bien ? ' ✓ al ' + p.ultima
                             : ' ' + p.dias + ' días atrasado (última: ' + p.ultima + ')');
    });
    lineas.push('· ' + c.empresa + ': ' + detalle.join(' · ') +
                (c.alDia ? '' : '\n    ← corre prenderAutomatico()'));
  });

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + lineas.join('\n');
  Logger.log(msg);
  return msg;
}


/* ═══════════════════════════════════════════════════════════════
   13 · META · LECTURA DIARIA
   ═══════════════════════════════════════════════════════════════ */

/**
 * ═══════════════════════════════════════════════════════════
 *  META · LA LECTURA DIARIA
 * ═══════════════════════════════════════════════════════════
 *
 * Aquí es donde empieza a pasar el agua. 75-meta.gs deja la llave puesta
 * y comprobada; este archivo la usa para pedirle a Meta el gasto de cada
 * día y escribirlo en Pauta, sin que nadie exporte un Excel.
 *
 * TRES COSAS QUE SE ARREGLAN SOLAS AL PEDIRLO ASÍ
 *
 * 1. `time_increment=1` obliga a Meta a devolver UNA FILA POR DÍA. El
 *    export manual trae una sola fila por todo el periodo, y repartirla
 *    entre los días dibuja una curva que nunca existió.
 *
 * 2. La moneda viene declarada por Meta (`account_currency`), no
 *    adivinada por el nombre de una columna. El export dice "Importe
 *    gastado (COP)" solo a veces, y cuando no lo dice, el importador
 *    tiene que suponer.
 *
 * 3. Las fechas vienen en ISO. El Excel las trae en el formato de quien
 *    lo exportó, y 03/04 es marzo o abril según el país.
 *
 * LO QUE NO SE INVENTA
 *
 * Meta no entrega presupuesto ni estado de entrega en este informe. Esas
 * columnas quedan VACÍAS, no en cero. Un cero dice "no había
 * presupuesto"; vacío dice "no lo sé", que es la verdad.
 */

/**
 * Qué se le pide a Meta.
 *
 * A nivel de conjunto (`adset`) y no de campaña porque el presupuesto se
 * decide por conjunto: es la unidad que alguien puede subir, bajar o
 * apagar mañana. Un consejo sobre una campaña que adentro tiene tres
 * conjuntos con rendimientos distintos no se puede ejecutar.
 */
const META_NIVEL = 'adset';

/**
 * Y los anuncios, uno por uno, en su propia hoja.
 *
 * Es otra pregunta: el conjunto dice dónde va el presupuesto, el anuncio
 * dice cuál creativo tira del carro. Se piden aparte y se guardan aparte
 * porque el gasto está en los dos —el conjunto de 100 son los mismos 100
 * repartidos entre sus anuncios— y sumarlos juntos contaría todo dos
 * veces.
 *
 * Cuesta llamadas: un conjunto con cuatro creativos son cuatro filas por
 * día en vez de una. Por eso se puede apagar por tienda, y por eso la
 * lectura diaria trae menos días de anuncios que de conjuntos.
 */
const META_NIVEL_ANUNCIO = 'ad';

const META_CAMPOS_ANUNCIO = [
  'date_start', 'account_currency', 'campaign_name', 'adset_name',
  'ad_name', 'ad_id', 'spend', 'impressions', 'reach', 'frequency',
  'clicks', 'ctr', 'cpc', 'cpm', 'actions', 'action_values',
];

/** Cuántos días de anuncios se repiden a diario. Menos que los conjuntos. */
const META_DIAS_ANUNCIOS = 3;

const META_CAMPOS = [
  'date_start', 'date_stop', 'account_currency',
  'campaign_name', 'adset_name', 'adset_id',
  'spend', 'impressions', 'reach', 'frequency',
  'clicks', 'ctr', 'cpc', 'cpm',
  'actions', 'action_values',
];

/**
 * Cuántos días se vuelven a pedir en cada lectura diaria.
 *
 * No basta con pedir "ayer". Meta sigue atribuyendo conversiones hasta
 * días después de que ocurrió el clic, así que las cifras de ayer
 * cambian pasado mañana. Se vuelven a pedir siete días y se reescriben
 * encima: el número siempre es el último que Meta conoce, no el primero
 * que dijo.
 */
const META_DIAS_DIARIO = 7;

/**
 * Cómo se reconoce una compra, y en qué orden.
 *
 * Meta devuelve una lista de "acciones" con nombres técnicos y varias
 * pueden ser la misma venta contada de dos formas. Se toma la primera de
 * esta lista que exista, y se DICE cuál se tomó, porque de esto depende
 * el CPA entero y quien lo mire tiene derecho a saber de dónde salió.
 */
const META_ACCIONES_COMPRA = [
  'offsite_conversion.fb_pixel_purchase',
  'onsite_web_purchase',
  'omni_purchase',
  'purchase',
];

/**
 * Y cómo se reconoce un registro.
 *
 * En contraentrega mucha gente no optimiza a compra sino a formulario:
 * el pedido se confirma después, por teléfono. Si no hay compras, el
 * resultado del conjunto son esos registros — pero nunca se mezclan los
 * dos en la misma columna.
 */
const META_ACCIONES_LEAD = [
  'offsite_conversion.fb_pixel_lead',
  'onsite_conversion.lead_grouped',
  'lead',
];

const META_ACCION_LP = ['landing_page_view'];

/** El primer tipo de la lista que Meta reportó, con su valor. */
function metaAccion_(acciones, tipos) {
  if (!acciones || !acciones.length) return null;
  const porTipo = {};
  acciones.forEach(function (a) { porTipo[a.action_type] = a.value; });
  for (let i = 0; i < tipos.length; i++) {
    if (porTipo[tipos[i]] !== undefined) {
      return { tipo: tipos[i], valor: Number(porTipo[tipos[i]]) || 0 };
    }
  }
  return null;
}

/**
 * Una llamada a Meta, con sus páginas.
 *
 * Meta parte las respuestas largas y deja un enlace al resto. Sin seguir
 * ese enlace, una cuenta con muchos conjuntos devolvería los primeros y
 * el resto simplemente no existiría — un gasto que falta sin que nada
 * avise es peor que un error.
 */
function metaPedir_(url, token) {
  const filas = [];
  let siguiente = url;
  let vueltas = 0;

  while (siguiente && vueltas < 40) {   // tope: una cuenta rota no cuelga el script
    vueltas++;
    let r;
    try {
      r = UrlFetchApp.fetch(siguiente, { muteHttpExceptions: true });
    } catch (err) {
      return { ok: false, error: 'No se pudo hablar con Meta: ' + err.message };
    }

    let j;
    try { j = JSON.parse(r.getContentText()); }
    catch (err) { return { ok: false, error: 'Meta contestó algo que no entiendo.' }; }

    if (j.error) return Object.assign({ ok: false }, explicarErrorMeta(j.error));

    (j.data || []).forEach(function (f) { filas.push(f); });

    // El enlace siguiente ya trae la llave adentro; el primero no.
    siguiente = (j.paging && j.paging.next) ? j.paging.next : '';
    if (siguiente && siguiente.indexOf('access_token') === -1) {
      siguiente += '&access_token=' + encodeURIComponent(token);
    }
  }

  return { ok: true, filas: filas, paginas: vueltas };
}

/** yyyy-MM-dd de hace N días, en la zona del script. */
function metaFecha_(diasAtras) {
  const d = new Date(Date.now() - diasAtras * 86400000);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Trae el gasto de una tienda y lo escribe en Pauta.
 *
 * Devuelve un informe de lo que hizo, no un "listo". Quien paga por esto
 * necesita poder auditarlo: cuántas filas entraron, cuántas cambiaron,
 * qué acción se contó como compra y si hay filas viejas que se solapan.
 *
 * @param {string} sheetId  la hoja del cliente
 * @param {string} tienda   el id de la tienda
 * @param {string} token    la llave de Meta de ESE cliente
 * @param {number} dias     cuántos días hacia atrás pedir
 */
function metaLeerTienda_(sheetId, tienda, token, dias) {
  const ss = SpreadsheetApp.openById(sheetId);
  const cuenta = metaCuenta(ss, tienda);
  const informe = {
    tienda: tienda, cuenta: cuenta, ok: false, error: '',
    filas: 0, nuevas: 0, actualizadas: 0, iguales: 0,
    desde: '', hasta: '', moneda: '', monedaTienda: monedaDeTienda(ss, tienda) || '',
    gasto: 0, accionCompra: '', accionResultado: '', solapadas: 0, avisos: [],
  };

  if (!cuenta) {
    informe.error = 'Esta tienda no tiene número de cuenta publicitaria. ' +
                    'Se pone en Configuración → Conexión con Meta.';
    return informe;
  }

  const hasta = metaFecha_(0);
  const desde = metaFecha_(Math.max(1, dias || META_DIAS_DIARIO));
  informe.desde = desde; informe.hasta = hasta;

  const url = META_API + 'act_' + cuenta + '/insights' +
    '?level=' + META_NIVEL +
    '&time_increment=1' +
    '&time_range=' + encodeURIComponent(JSON.stringify({ since: desde, until: hasta })) +
    '&fields=' + META_CAMPOS.join(',') +
    '&limit=200' +
    '&access_token=' + encodeURIComponent(token);

  const r = metaPedir_(url, token);
  if (!r.ok) { informe.error = r.error; return informe; }

  informe.filas = r.filas.length;
  if (!r.filas.length) {
    informe.ok = true;
    informe.avisos.push('Meta no reportó gasto en esos días.');
    return informe;
  }

  const tiposCompra = {}, tiposResultado = {};
  const filas = r.filas.map(function (f) {
    const fecha = String(f.date_start || '').slice(0, 10);
    const moneda = String(f.account_currency || '').toUpperCase();
    if (moneda) informe.moneda = moneda;

    const compra = metaAccion_(f.actions, META_ACCIONES_COMPRA);
    const lead   = metaAccion_(f.actions, META_ACCIONES_LEAD);
    const lp     = metaAccion_(f.actions, META_ACCION_LP);
    const valor  = metaAccion_(f.action_values, META_ACCIONES_COMPRA);

    if (compra) tiposCompra[compra.tipo] = 1;
    const resultado = compra || lead;
    if (resultado) tiposResultado[resultado.tipo] = 1;

    const gasto = Number(f.spend) || 0;
    informe.gasto += gasto;

    /**
     * El identificador es el MISMO que arma el importador del Excel.
     *
     * A propósito: si alguien sube el export del mismo día y el mismo
     * conjunto, las dos filas son la misma fila y la segunda pisa a la
     * primera en vez de sumarse. Dos formas de traer el dato, un solo
     * número.
     */
    const conjunto = String(f.adset_name || f.campaign_name || 'Sin nombre').trim();
    const id = ['meta', tienda, fecha, fecha, norm(conjunto)].join('-')
      .replace(/\s+/g, '_').slice(0, 180);

    return {
      id: id,
      fecha: fecha,
      fecha_fin: fecha,          // una fila por día: el periodo es el día
      tienda: tienda,
      plataforma: 'meta',
      cuenta: cuenta,
      campana: String(f.campaign_name || '').trim(),
      conjunto: conjunto,
      // `entrega` y `presupuesto` no vienen en este informe. Vacías, no
      // en cero: un cero diría que no había presupuesto.
      gasto: gasto,
      moneda_gasto: moneda,
      impresiones: Number(f.impressions) || 0,
      alcance: Number(f.reach) || 0,
      frecuencia: Number(f.frequency) || 0,
      clics: Number(f.clicks) || 0,
      ctr: Number(f.ctr) || 0,
      cpc: Number(f.cpc) || 0,
      cpm: Number(f.cpm) || 0,
      resultados: resultado ? resultado.valor : 0,
      compras: compra ? compra.valor : 0,
      // CPA y ROAS se calculan aquí porque los dos números que los forman
      // están aquí y son del mismo día. Recalcularlos después obliga a
      // adivinar contra qué gasto se dividían.
      cpa: (resultado && resultado.valor) ? gasto / resultado.valor : '',
      roas: (valor && gasto) ? valor.valor / gasto : '',
      valor_conv: valor ? valor.valor : '',
      visitas_lp: lp ? lp.valor : '',
    };
  });

  informe.accionCompra = Object.keys(tiposCompra).join(', ');
  informe.accionResultado = Object.keys(tiposResultado).join(', ');

  // Si Meta no optimiza a compra, el CPA que sale no es por venta
  if (!informe.accionCompra && informe.accionResultado) {
    informe.avisos.push('Meta no reportó compras en estos conjuntos: el resultado ' +
      'que se contó es «' + informe.accionResultado + '». El CPA de esta pantalla ' +
      'es por ese resultado, no por venta entregada.');
  }
  if (informe.moneda && informe.monedaTienda && informe.moneda !== informe.monedaTienda) {
    informe.avisos.push('Meta cobra en ' + informe.moneda + ' y la tienda factura en ' +
      informe.monedaTienda + '. Nova convierte con la tasa del día de cada gasto; ' +
      'si faltan tasas, ese gasto se cuenta aparte y la pantalla lo dice.');
  }

  let esc;
  try {
    esc = escribirFilas(ss, 'Pauta', filas, 'meta');
  } catch (err) {
    informe.error = err.message;
    return informe;
  }
  informe.nuevas = esc.nuevas;
  informe.actualizadas = esc.actualizadas;
  informe.iguales = esc.iguales;

  informe.solapadas = metaFilasSolapadas_(ss, tienda, desde, hasta);
  if (informe.solapadas) {
    informe.avisos.push(informe.solapadas + ' fila(s) de un Excel subido a mano cubren ' +
      'un rango de varios días dentro de este periodo. Esas SÍ se suman por separado ' +
      'y el gasto quedaría contado dos veces. Bórralas en la hoja Pauta.');
  }

  informe.ok = true;
  return informe;
}

/**
 * Filas de Pauta que cubren un rango de varios días dentro del periodo.
 *
 * Son las únicas que se pueden contar dos veces. Las de un solo día
 * comparten identificador con las que trae la API, así que se pisan
 * entre ellas y no hay riesgo. Estas no, y no se borran solas: borrar
 * datos que alguien subió es una decisión suya, no de Nova.
 */
function metaFilasSolapadas_(ss, tienda, desde, hasta) {
  const sh = ss.getSheetByName('Pauta');
  if (!sh || sh.getLastRow() < 2) return 0;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')]).trim() !== tienda) continue;
    if (norm(d[i][c('plataforma')]) !== 'meta') continue;
    const f1 = aISO(d[i][c('fecha')], 'UTC');
    const f2 = aISO(d[i][c('fecha_fin')], 'UTC');
    if (!f1 || !f2 || f1 === f2) continue;      // de un solo día: no hay riesgo
    if (f2 < desde || f1 > hasta) continue;     // fuera del periodo leído
    n++;
  }
  return n;
}

/**
 * Los anuncios de una tienda, a su propia hoja.
 *
 * No toca Pauta. Nada de lo que escribe aquí entra en ninguna suma de
 * gasto: esta hoja existe para comparar creativos entre sí, no para
 * cuadrar plata. El total de la cuenta sigue saliendo de Pauta.
 */
function metaLeerAnuncios_(sheetId, tienda, token, dias) {
  const ss = SpreadsheetApp.openById(sheetId);
  const cuenta = metaCuenta(ss, tienda);
  const informe = { tienda: tienda, ok: false, error: '', filas: 0,
                    nuevas: 0, actualizadas: 0, iguales: 0, desde: '', hasta: '' };

  if (!cuenta) { informe.error = 'Esta tienda no tiene cuenta publicitaria.'; return informe; }
  if (!ss.getSheetByName('Anuncios')) {
    informe.error = 'Falta la hoja Anuncios. Corre bootstrapTodo() una vez.';
    return informe;
  }

  const hasta = metaFecha_(0);
  const desde = metaFecha_(Math.max(1, dias || META_DIAS_ANUNCIOS));
  informe.desde = desde; informe.hasta = hasta;

  const url = META_API + 'act_' + cuenta + '/insights' +
    '?level=' + META_NIVEL_ANUNCIO +
    '&time_increment=1' +
    '&time_range=' + encodeURIComponent(JSON.stringify({ since: desde, until: hasta })) +
    '&fields=' + META_CAMPOS_ANUNCIO.join(',') +
    '&limit=200' +
    '&access_token=' + encodeURIComponent(token);

  const r = metaPedir_(url, token);
  if (!r.ok) { informe.error = r.error; return informe; }

  informe.filas = r.filas.length;
  if (!r.filas.length) { informe.ok = true; return informe; }

  const filas = r.filas.map(function (f) {
    const fecha = String(f.date_start || '').slice(0, 10);
    const compra = metaAccion_(f.actions, META_ACCIONES_COMPRA);
    const lead   = metaAccion_(f.actions, META_ACCIONES_LEAD);
    const lp     = metaAccion_(f.actions, META_ACCION_LP);
    const valor  = metaAccion_(f.action_values, META_ACCIONES_COMPRA);
    const resultado = compra || lead;
    const gasto = Number(f.spend) || 0;

    // El id lleva el ad_id de Meta, que sí es único y estable. En Pauta
    // no se puede: el informe de conjuntos no siempre trae adset_id.
    return {
      id: ['meta', tienda, fecha, String(f.ad_id || norm(f.ad_name || ''))]
            .join('-').replace(/\s+/g, '_').slice(0, 180),
      fecha: fecha, tienda: tienda, plataforma: 'meta', cuenta: cuenta,
      campana: String(f.campaign_name || '').trim(),
      conjunto: String(f.adset_name || '').trim(),
      anuncio: String(f.ad_name || '').trim(),
      anuncio_id: String(f.ad_id || ''),
      gasto: gasto,
      moneda_gasto: String(f.account_currency || '').toUpperCase(),
      impresiones: Number(f.impressions) || 0,
      alcance: Number(f.reach) || 0,
      frecuencia: Number(f.frequency) || 0,
      clics: Number(f.clicks) || 0,
      ctr: Number(f.ctr) || 0,
      cpc: Number(f.cpc) || 0,
      cpm: Number(f.cpm) || 0,
      resultados: resultado ? resultado.valor : 0,
      compras: compra ? compra.valor : 0,
      cpa: (resultado && resultado.valor) ? gasto / resultado.valor : '',
      valor_conv: valor ? valor.valor : '',
      visitas_lp: lp ? lp.valor : '',
    };
  });

  try {
    const esc = escribirFilas(ss, 'Anuncios', filas, 'meta');
    informe.nuevas = esc.nuevas;
    informe.actualizadas = esc.actualizadas;
    informe.iguales = esc.iguales;
    informe.ok = true;
  } catch (err) {
    informe.error = err.message;
  }
  return informe;
}

/** Todas las tiendas de un cliente. */
function metaLeerCliente_(sheetId, dias) {
  const token = PropertiesService.getScriptProperties()
    .getProperty('META_TOKEN_' + String(sheetId).slice(0, 44)) || '';
  if (!token) return { sheetId: sheetId, sinLlave: true, tiendas: [] };

  const ss = SpreadsheetApp.openById(sheetId);
  const shT = ss.getSheetByName('Tiendas');
  if (!shT || shT.getLastRow() < 2) return { sheetId: sheetId, tiendas: [] };

  const filas = shT.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cEstado = enc.indexOf('estado');

  const out = [];
  filas.slice(1).forEach(function (f) {
    const id = String(f[cId] || '').trim();
    if (!id) return;
    if (cEstado !== -1 && norm(f[cEstado]) === 'inactiva') return;
    if (!metaCuenta(ss, id)) return;            // esa tienda no tiene cuenta conectada
    const r = metaLeerTienda_(sheetId, id, token, dias);
    // Los anuncios solo si esa tienda los quiere. Y si fallan, no
    // arrastran a la lectura de conjuntos: el gasto ya quedó escrito.
    if (norm(ajustes(ss, id).meta_anuncios) === 'si') {
      try {
        r.anuncios = metaLeerAnuncios_(sheetId, id, token, META_DIAS_ANUNCIOS);
        if (!r.anuncios.ok && r.anuncios.error) {
          r.avisos.push('Los conjuntos entraron bien, pero los anuncios no: ' +
                        r.anuncios.error);
        }
      } catch (e) {
        r.avisos.push('Los conjuntos entraron bien, pero los anuncios no: ' + e.message);
      }
    }
    out.push(r);
  });
  return { sheetId: sheetId, tiendas: out };
}

/**
 * Lo que corre el disparador: todos los clientes, todas las mañanas.
 *
 * Un cliente que falla no puede impedir que se lea el siguiente. Se anota
 * el fallo y se sigue.
 */
function leerMetaDiario() {
  const lineas = [];
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try {
      const r = metaLeerCliente_(c.sheetId, META_DIAS_DIARIO);
      if (r.sinLlave) return;                   // no está conectado: no es un fallo
      r.tiendas.forEach(function (t) {
        lineas.push(c.empresa + ' · ' + t.tienda + ': ' +
          (t.ok ? t.filas + ' filas (' + t.nuevas + ' nuevas, ' +
                  t.actualizadas + ' corregidas)' +
                  (t.avisos.length ? '  OJO · ' + t.avisos.join(' ') : '')
                : 'FALLÓ · ' + t.error));
      });
    } catch (e) {
      lineas.push(c.empresa + ': FALLÓ · ' + e.message);
    }
  });

  const msg = lineas.length ? lineas.join('\n') : 'Ninguna cuenta tiene Meta conectado.';
  Logger.log(msg);
  return msg;
}

/**
 * El botón «Traer ahora» de la pantalla.
 *
 * Existe por dos razones distintas. Una: al conectar, nadie quiere
 * esperar hasta mañana para ver si sirve. Dos: para traer el historial
 * de una cuenta vieja, que es lo que hace falta el primer día.
 */
function apiMetaTraer(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña trae la pauta.' };

  const token = metaToken(s);
  if (!token) return { ok: false, error: 'Todavía no has guardado la llave.' };

  /**
   * El tope de días no es un capricho.
   *
   * Meta solo guarda insights de los últimos 37 meses, y pedir de a un
   * día multiplica las llamadas: 400 días son 400 filas por conjunto y
   * varias páginas. El nivel de acceso de una app nueva permite 300
   * llamadas por hora, así que un historial largo se trae por tramos.
   */
  const dias = Math.min(Math.max(Number(p.dias) || META_DIAS_DIARIO, 1), 400);

  const tienda = String(p.tienda || s.tiendas[0]);
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }

  const informe = metaLeerTienda_(s.sheetId, tienda, token, dias);
  if (!informe.ok) return { ok: false, error: informe.error };

  /**
   * Los anuncios se traen con menos días que los conjuntos, aunque se
   * pidan muchos. Un año a nivel de anuncio son miles de filas y varias
   * páginas por cuenta: se llenaría la cuota de la hora y quedaría a
   * medias, que es peor que no traerlo.
   */
  const ss2 = SpreadsheetApp.openById(s.sheetId);
  if (norm(ajustes(ss2, tienda).meta_anuncios) === 'si') {
    try {
      informe.anuncios = metaLeerAnuncios_(s.sheetId, tienda, token, Math.min(dias, 90));
      if (!informe.anuncios.ok && informe.anuncios.error) {
        informe.avisos.push('Los anuncios no se pudieron traer: ' + informe.anuncios.error);
      }
    } catch (e) {
      informe.avisos.push('Los anuncios no se pudieron traer: ' + e.message);
    }
  }

  registrarMovimiento(s, 'Pauta', 'meta', 'traer', tienda,
    informe.nuevas + ' nuevas / ' + informe.actualizadas + ' corregidas');

  return { ok: true, informe: informe };
}


/* ═══════════════════════════════════════════════════════════════
   14 · SEMÁFORO SEMANAL
   ═══════════════════════════════════════════════════════════════ */

/**
 * ═══════════════════════════════════════════════════════════
 *  EL SEMÁFORO SEMANAL
 * ═══════════════════════════════════════════════════════════
 *
 * Cada lunes en la mañana: qué pasó la semana que cerró, cruzando la
 * pauta de Meta con lo que de verdad se entregó.
 *
 * POR QUÉ CRUZARLOS ES EL PUNTO
 *
 * Meta cuenta una compra cuando alguien confirma en el checkout. En
 * contraentrega, una parte de esas compras nunca se entrega: el ROAS que
 * muestra Meta está inflado y con él se deciden presupuestos. Nova sabe
 * cuáles llegaron, porque tiene los dos lados.
 *
 * LAS REGLAS DE CÁLCULO, QUE NO SON NEGOCIABLES
 *
 * Son las que ya venían usándose a mano, escritas aquí para que no
 * cambien sin que alguien lo decida:
 *
 *   · Solo un pedido ENTREGADO genera ganancia.
 *   · Cada DEVOLUCIÓN cuesta el flete: el producto vuelve, el envío no.
 *   · Un CANCELADO cuesta cero — nunca salió.
 *   · La tasa de entrega es entregados ÷ (entregados + devoluciones).
 *     Lo que sigue en tránsito no entra: todavía no falló.
 *   · El techo de CPA es la utilidad antes de pauta ÷ pedidos creados.
 *     Es cuánto se puede pagar por pedido sin perder plata.
 *
 * Y LA REGLA QUE PROTEGE A TODAS LAS DEMÁS
 *
 * Donde no hay dato, dice «sin dato». Nada se extrapola, nada se rellena
 * con un promedio. Una semana sin pedidos en la hoja no significa que no
 * se vendió: significa que la hoja no se actualizó, y son dos cosas
 * distintas que llevan a decisiones opuestas.
 */

/**
 * Los umbrales de esta tienda, como números.
 *
 * Viven en AJUSTES_DEFAULT con el resto de ajustes, y cada tienda los
 * cambia en Parametros. Aquí solo se convierten: lo que vuelve de una
 * hoja es texto, y "72" no es 72 cuando se compara con un porcentaje.
 *
 * No hay una segunda lista de valores por defecto a propósito. Dos
 * listas del mismo umbral acaban diciendo cosas distintas, y entonces el
 * semáforo se contradice con la pantalla que lo configura.
 */
const SEMAFORO_CLAVES = ['ticket_minimo', 'entrega_minima',
                         'muestra_minima', 'cpa_verde_pct'];

function ajustesSemaforo(ss, tienda) {
  const a = ajustes(ss, tienda) || {};
  const out = {};
  SEMAFORO_CLAVES.forEach(function (k) {
    const v = Number(a[k]);
    out[k] = (a[k] === '' || a[k] === undefined || isNaN(v))
      ? Number(AJUSTES_DEFAULT[k]) || 0 : v;
  });
  return out;
}

// ─── SEMANAS ─────────────────────────────────────────────────

/** El lunes de la semana de una fecha ISO. */
function lunesDe_(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  const dow = d.getUTCDay();                 // 0 domingo … 6 sábado
  const atras = (dow === 0) ? 6 : dow - 1;   // el domingo pertenece a la semana que cierra
  d.setUTCDate(d.getUTCDate() - atras);
  return d.toISOString().slice(0, 10);
}

function masDias_(iso, n) {
  const d = new Date(iso + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

/** La semana cerrada más reciente: de lunes a domingo, ya terminada. */
function semanaCerrada_(hoyISO) {
  return masDias_(lunesDe_(hoyISO), -7);
}

// ─── LOS NÚMEROS DE UNA SEMANA ───────────────────────────────

/**
 * Todo lo que se puede contar de una semana, sin juzgar nada todavía.
 *
 * Separado del semáforo a propósito: los números son hechos y los
 * colores son opiniones. Poder mirar los hechos sin la opinión encima es
 * lo que permite discutir la opinión.
 */
function numerosSemana_(pedidos, pauta, lunes) {
  const domingo = masDias_(lunes, 6);
  const o = {
    lunes: lunes, domingo: domingo,
    pedidos: 0, despachados: 0, entregados: 0, devoluciones: 0, cancelados: 0,
    enTransito: 0, sinClasificar: 0,
    ventas: 0, ganancia: 0, costoDevoluciones: 0,
    gasto: 0, gastoSinConvertir: 0,
    diasConPedidos: {}, productos: {},
    sinCosto: 0,                 // entregados sin costo de producto
  };

  pedidos.forEach(function (p) {
    if (p.fecha < lunes || p.fecha > domingo) return;
    o.pedidos++;
    o.diasConPedidos[p.fecha] = (o.diasConPedidos[p.fecha] || 0) + 1;

    const est = p.estado;
    if (est !== ESTADOS.CANCELADO && est !== ESTADOS.PENDIENTE &&
        est !== ESTADOS.SIN_CLASIFICAR) o.despachados++;

    const prod = p.producto || '(sin producto)';
    if (!o.productos[prod]) {
      o.productos[prod] = { pedidos: 0, entregados: 0, devoluciones: 0, ganancia: 0 };
    }
    const pr = o.productos[prod];
    pr.pedidos++;

    if (est === ESTADOS.ENTREGADO) {
      o.entregados++; pr.entregados++;
      o.ventas += p.valor;
      /**
       * La ganancia solo se calcula si hay costo de producto.
       *
       * Sin costo, valor − flete daría una ganancia enorme y falsa. Se
       * cuenta aparte y se dice cuántos quedaron fuera: una utilidad a la
       * que le faltan pedidos es mejor que una utilidad inventada, pero
       * solo si se sabe que le faltan.
       */
      if (p.costoProducto > 0) {
        const g = p.valor - p.costoProducto - p.costoEnvio;
        o.ganancia += g;
        pr.ganancia += g;
      } else {
        o.sinCosto++;
      }
    } else if (est === ESTADOS.DEVOLUCION) {
      o.devoluciones++; pr.devoluciones++;
      // El producto vuelve; el flete ya se pagó y no vuelve.
      o.costoDevoluciones += p.costoEnvio;
    } else if (est === ESTADOS.CANCELADO) {
      o.cancelados++;
    } else if (est === ESTADOS.SIN_CLASIFICAR) {
      o.sinClasificar++;
    } else {
      o.enTransito++;
    }
  });

  pauta.forEach(function (g) {
    if (g.fecha < lunes || g.fecha > domingo) return;
    if (g.convertido === null) { o.gastoSinConvertir += g.gasto; return; }
    o.gasto += g.convertido;
  });

  const resueltos = o.entregados + o.devoluciones;
  o.tasaEntrega = resueltos ? o.entregados / resueltos * 100 : null;
  o.ticket = o.entregados ? o.ventas / o.entregados : null;
  o.utilidadAntesPauta = o.ganancia - o.costoDevoluciones;
  o.hayPauta = o.gasto > 0 || o.gastoSinConvertir > 0;
  o.utilidadReal = o.hayPauta ? o.utilidadAntesPauta - o.gasto : null;
  o.cpaPorPedido = (o.hayPauta && o.pedidos) ? o.gasto / o.pedidos : null;
  o.cpaPorEntrega = (o.hayPauta && o.entregados) ? o.gasto / o.entregados : null;
  // El techo: cuánto se puede pagar por pedido sin perder plata.
  o.techo = o.pedidos ? o.utilidadAntesPauta / o.pedidos : null;
  o.diasConDatos = Object.keys(o.diasConPedidos).length;
  return o;
}

// ─── LEER LAS HOJAS UNA SOLA VEZ ─────────────────────────────

/**
 * Pedidos y pauta de las últimas N semanas, ya normalizados.
 *
 * Se leen de una vez y se reparten por semana después. Abrir la hoja una
 * vez por semana analizada multiplicaba por ocho el trabajo para leer
 * exactamente los mismos datos.
 */
function datosParaSemaforo_(ss, tienda, desde, hasta) {
  const monTienda = monedaDeTienda(ss, tienda) || '';
  const pedidos = [];
  const pauta = [];

  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha || fecha < desde || fecha > hasta) continue;
      pedidos.push({
        fecha: fecha,
        estado: norm(f[c('estado_nova')] || f[c('estado_canonico')]) || ESTADOS.SIN_CLASIFICAR,
        valor: num(f[c('valor')]),
        costoProducto: num(f[c('costo_producto')]),
        costoEnvio: num(f[c('costo_envio')]),
        producto: String(f[c('producto')] || '').trim(),
      });
    }
  }

  const shA = ss.getSheetByName('Pauta');
  if (shA && shA.getLastRow() > 1) {
    const d = shA.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha || fecha < desde || fecha > hasta) continue;
      const gasto = num(f[c('gasto')]);
      const mon = String(f[c('moneda_gasto')] || '').toUpperCase();
      let convertido = gasto;
      if (mon && monTienda && mon !== monTienda) {
        const t = buscarTasa(ss, fecha, mon, monTienda);
        convertido = (t && t.tasa) ? gasto * t.tasa : null;
      }
      pauta.push({ fecha: fecha, gasto: gasto, convertido: convertido,
                   conjunto: String(f[c('conjunto')] || f[c('campana')] || '').trim() });
    }
  }

  return { pedidos: pedidos, pauta: pauta, moneda: monTienda };
}

// ─── EL SEMÁFORO ─────────────────────────────────────────────

/**
 * Un semáforo: qué mide, cuánto dio, de qué color y por qué.
 *
 * El «por qué» es obligatorio y va junto al color a propósito. Un color
 * sin su razón es una orden; con ella es un argumento, y un argumento se
 * puede discutir cuando el umbral está mal puesto.
 */
function luz_(estado, valor, etiqueta, porque, delta) {
  return { estado: estado, valor: valor, etiqueta: etiqueta, porque: porque,
           delta: (delta === undefined ? null : delta) };
}

function semaforosDe_(hoy, antes, u, moneda) {
  const luces = [];

  // 1 · CPA por entrega contra el techo
  if (hoy.cpaPorEntrega === null || hoy.techo === null) {
    luces.push(luz_('sin_medir', null, 'CPA por entrega',
      !hoy.hayPauta ? 'No hay gasto de pauta cargado en esta semana.'
                    : 'No hubo entregas, así que no hay contra qué dividir.'));
  } else {
    const pct = hoy.techo > 0 ? hoy.cpaPorEntrega / hoy.techo * 100 : 999;
    const e = pct <= u.cpa_verde_pct ? 'verde' : (pct <= 100 ? 'amarillo' : 'rojo');
    luces.push(luz_(e, hoy.cpaPorEntrega, 'CPA por entrega',
      'El techo es ' + redondear_(hoy.techo) + ' ' + moneda + '. ' +
      (e === 'rojo' ? 'Está por encima: cada entrega pierde plata.'
        : 'Queda ' + redondear_(hoy.techo - hoy.cpaPorEntrega) + ' de margen.')));
  }

  // 2 · Ticket promedio
  if (hoy.ticket === null) {
    luces.push(luz_('sin_medir', null, 'Ticket promedio', 'No hubo entregas esta semana.'));
  } else if (!u.ticket_minimo) {
    luces.push(luz_('sin_umbral', hoy.ticket, 'Ticket promedio',
      'No hay un mínimo definido. Se pone en Configuración para que esto se pueda juzgar.',
      antes && antes.ticket !== null ? hoy.ticket - antes.ticket : null));
  } else {
    const e = hoy.ticket >= u.ticket_minimo ? 'verde'
            : hoy.ticket >= u.ticket_minimo * 0.9 ? 'amarillo' : 'rojo';
    luces.push(luz_(e, hoy.ticket, 'Ticket promedio',
      'El mínimo que definiste es ' + redondear_(u.ticket_minimo) + ' ' + moneda + '.',
      antes && antes.ticket !== null ? hoy.ticket - antes.ticket : null));
  }

  // 3 · Tasa de entrega
  if (hoy.tasaEntrega === null) {
    luces.push(luz_('sin_medir', null, 'Entrega sobre resuelto',
      'Ningún pedido de la semana llegó todavía a entregado ni a devolución.'));
  } else {
    const resueltos = hoy.entregados + hoy.devoluciones;
    const chica = resueltos < u.muestra_minima;
    const e = chica ? 'muestra_chica'
      : hoy.tasaEntrega >= u.entrega_minima ? 'verde'
      : hoy.tasaEntrega >= u.entrega_minima - 8 ? 'amarillo' : 'rojo';
    luces.push(luz_(e, hoy.tasaEntrega, 'Entrega sobre resuelto',
      hoy.entregados + ' de ' + resueltos + ' resueltos' +
      (hoy.enTransito ? ' · ' + hoy.enTransito +
        (hoy.enTransito === 1 ? ' sigue' : ' siguen') + ' en tránsito, fuera de esta cuenta' : '') +
      (chica ? '. Con menos de ' + u.muestra_minima + ' resueltos esta cifra no es una señal.' : ''),
      antes && antes.tasaEntrega !== null ? hoy.tasaEntrega - antes.tasaEntrega : null));
  }

  // 4 · Utilidad
  if (!hoy.hayPauta) {
    luces.push(luz_('incompleto', hoy.utilidadAntesPauta, 'Utilidad de la semana',
      'Es ANTES de pauta. Sin el gasto de anuncios no se puede cerrar.',
      antes ? hoy.utilidadAntesPauta - antes.utilidadAntesPauta : null));
  } else if (hoy.gastoSinConvertir > 0) {
    luces.push(luz_('incompleto', hoy.utilidadReal, 'Utilidad de la semana',
      redondear_(hoy.gastoSinConvertir) + ' de pauta quedaron sin convertir por falta ' +
      'de tasa de cambio, así que no están restados aquí.',
      antes && antes.utilidadReal !== null ? hoy.utilidadReal - antes.utilidadReal : null));
  } else {
    const e = hoy.utilidadReal > 0
      ? (antes && antes.utilidadReal !== null && hoy.utilidadReal < antes.utilidadReal
          ? 'amarillo' : 'verde')
      : 'rojo';
    luces.push(luz_(e, hoy.utilidadReal, 'Utilidad de la semana',
      e === 'rojo' ? 'La semana cerró en pérdida.'
        : 'Después de producto, flete, devoluciones y pauta.',
      antes && antes.utilidadReal !== null ? hoy.utilidadReal - antes.utilidadReal : null));
  }

  return luces;
}

/**
 * Un número para leer, no para operar.
 *
 * Con separador de miles: «1.386.315» se lee de un vistazo y «1386315»
 * hay que contarlo con el dedo. Esto es lo que va en el correo del lunes,
 * que alguien lee en el teléfono a las siete de la mañana.
 */
function redondear_(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  const x = Math.round(Number(n) * 100) / 100;
  try { return x.toLocaleString('es-CO', { maximumFractionDigits: 2 }); }
  catch (e) { return String(x); }
}

// ─── LAS ALERTAS ─────────────────────────────────────────────

/**
 * Lo que hay que mirar, con qué hacer al respecto.
 *
 * Cada alerta trae una acción concreta. Una alerta sin acción solo
 * informa de que algo va mal, y quien la lee queda igual de atascado que
 * antes pero además preocupado.
 */
function alertasDe_(hoy, antes, u, moneda, cobertura) {
  const a = [];

  // Lo primero: ¿los datos están?
  if (cobertura.semanasVacias.length) {
    a.push({ nivel: 'rojo',
      titulo: 'Hay ' + cobertura.semanasVacias.length + ' semana(s) sin un solo pedido en la hoja',
      accion: 'Eso no significa que no vendiste: significa que la hoja no se ' +
        'actualizó. Baja de Dropi el reporte con rango completo y súbelo en ' +
        'Importar. Sin eso el semáforo no mide nada. Semanas vacías: ' +
        cobertura.semanasVacias.join(', ') + '.' });
  }
  if (hoy.diasConDatos > 0 && hoy.diasConDatos < 7) {
    a.push({ nivel: 'amarillo',
      titulo: 'La semana tiene ' + hoy.diasConDatos + ' de 7 días con registros',
      accion: 'Todo lo de abajo está calculado sobre esos días. No se extrapoló ' +
        'nada, así que las cifras son parciales, no bajas.' });
  }
  if (hoy.sinClasificar) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.sinClasificar + ' pedido(s) con un estado que Nova no reconoce',
      accion: 'No entran en ninguna cuenta: ni entregados, ni devueltos, ni ' +
        'despachados. Se clasifican en Configuración → Estados, una sola vez ' +
        'por estado nuevo.' });
  }
  if (hoy.sinCosto) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.sinCosto + ' pedido(s) entregados no tienen costo de producto',
      accion: 'Su ganancia quedó fuera de la utilidad. Se llena en Inventario, ' +
        'y desde ahí cuenta para todos los cierres.' });
  }
  if (hoy.gastoSinConvertir > 0) {
    a.push({ nivel: 'rojo',
      titulo: redondear_(hoy.gastoSinConvertir) + ' de pauta sin convertir',
      accion: 'Meta cobra en otra moneda y faltan las tasas de esos días. Ese ' +
        'gasto NO está restado de la utilidad. Se arregla en Nova Central, ' +
        'con «Prender lo automático».' });
  }

  // Y después: ¿cómo va el negocio?
  if (hoy.tasaEntrega !== null &&
      (hoy.entregados + hoy.devoluciones) >= u.muestra_minima &&
      hoy.tasaEntrega < u.entrega_minima) {
    a.push({ nivel: 'rojo',
      titulo: 'Entrega en ' + redondear_(hoy.tasaEntrega) + '%, bajo tu mínimo de ' +
        u.entrega_minima + '%',
      accion: 'Cada punto de entrega que se cae baja el techo de CPA. Mira de qué ' +
        'ciudades y de qué anuncios vienen las devoluciones antes de tocar el ' +
        'presupuesto.' });
  }
  if (hoy.cpaPorEntrega !== null && hoy.techo !== null && hoy.cpaPorEntrega > hoy.techo) {
    a.push({ nivel: 'rojo',
      titulo: 'El CPA por entrega (' + redondear_(hoy.cpaPorEntrega) + ') está sobre el techo (' +
        redondear_(hoy.techo) + ')',
      accion: 'Cada entrega de esta semana costó más de lo que deja. O sube el ' +
        'ticket, o baja el costo por entrega, o se apaga lo que menos entrega.' });
  }
  if (antes && hoy.pedidos && antes.pedidos && hoy.pedidos < antes.pedidos * 0.6) {
    a.push({ nivel: 'amarillo',
      titulo: 'Los pedidos cayeron de ' + antes.pedidos + ' a ' + hoy.pedidos,
      accion: 'Confirma si bajaste el presupuesto o si se apagó algo solo. Una ' +
        'caída de este tamaño sin decisión detrás suele ser un anuncio rechazado.' });
  }
  if (hoy.pedidos && hoy.cancelados / hoy.pedidos > 0.2) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.cancelados + ' de ' + hoy.pedidos + ' pedidos cancelados (' +
        Math.round(hoy.cancelados / hoy.pedidos * 100) + '%)',
      accion: 'No cuestan flete, pero es 1 de cada ' +
        Math.round(hoy.pedidos / Math.max(hoy.cancelados, 1)) +
        ' leads pagados que nunca llega a despacho.' });
  }

  if (!a.length) {
    a.push({ nivel: 'ninguno', titulo: 'Nada que reportar esta semana',
      accion: 'Los números están dentro de lo que definiste. Es una respuesta ' +
        'válida y conviene que exista: si el semáforo siempre encontrara algo, ' +
        'dejaría de significar algo cuando lo encuentre.' });
  }
  return a;
}

// ─── COBERTURA ───────────────────────────────────────────────

/**
 * Qué días tiene la hoja, semana por semana.
 *
 * Va antes que cualquier número porque decide si los números valen. Una
 * semana entera ausente no es una semana mala: es una semana que no se
 * importó, y confundirlas lleva a apagar una campaña que estaba
 * funcionando.
 */
function coberturaDe_(pedidos, lunes, semanas) {
  const porDia = {};
  pedidos.forEach(function (p) { porDia[p.fecha] = (porDia[p.fecha] || 0) + 1; });

  const filas = [];
  const vacias = [];
  for (let s = semanas - 1; s >= 0; s--) {
    const ini = masDias_(lunes, -7 * s);
    const dias = [];
    let total = 0;
    for (let d = 0; d < 7; d++) {
      const f = masDias_(ini, d);
      const n = porDia[f] || 0;
      total += n;
      dias.push({ fecha: f, pedidos: n });
    }
    filas.push({ lunes: ini, dias: dias, total: total });
    if (!total) vacias.push(ini);
  }
  return { filas: filas, semanasVacias: vacias };
}

// ─── LO QUE SE ARMA Y SE ENTREGA ─────────────────────────────

/**
 * El semáforo completo de una tienda.
 *
 * @param {string} sheetId  hoja del cliente
 * @param {string} tienda   id de la tienda
 * @param {string} lunes    lunes de la semana a analizar; vacío = la última cerrada
 */
function semaforoSemanal(sheetId, tienda, lunes) {
  const ss = SpreadsheetApp.openById(sheetId);
  const hoyISO = ahoraISO().slice(0, 10);
  const L = lunes || semanaCerrada_(hoyISO);
  const SEMANAS = 8;

  const desde = masDias_(L, -7 * (SEMANAS - 1));
  const hasta = masDias_(L, 6);
  const u = ajustesSemaforo(ss, tienda);
  const d = datosParaSemaforo_(ss, tienda, desde, hasta);

  const hoy = numerosSemana_(d.pedidos, d.pauta, L);
  const antes = numerosSemana_(d.pedidos, d.pauta, masDias_(L, -7));

  // El promedio de las 4 semanas previas: por semana de CALENDARIO, no
  // por semana con datos. Dividir entre las que tienen datos esconde
  // justamente las que faltan.
  const previas = [];
  for (let k = 2; k <= 5; k++) previas.push(numerosSemana_(d.pedidos, d.pauta, masDias_(L, -7 * k)));
  const prom = {};
  ['pedidos', 'entregados', 'devoluciones', 'cancelados', 'ventas',
   'utilidadAntesPauta', 'gasto'].forEach(function (k) {
    prom[k] = previas.reduce(function (a, x) { return a + (x[k] || 0); }, 0) / previas.length;
  });

  const cobertura = coberturaDe_(d.pedidos, L, SEMANAS);

  return {
    tienda: tienda, moneda: d.moneda,
    semana: { lunes: L, domingo: hasta },
    generado: ahoraISO(),
    umbrales: u,
    hoy: hoy, anterior: antes, promedio4: prom,
    luces: semaforosDe_(hoy, antes, u, d.moneda),
    alertas: alertasDe_(hoy, antes, u, d.moneda, cobertura),
    cobertura: cobertura,
    productos: productosOrdenados_(hoy, u),
    hayDatos: hoy.pedidos > 0,
  };
}

/** Los productos de la semana, del que más pedidos trajo al que menos. */
function productosOrdenados_(hoy, u) {
  return Object.keys(hoy.productos).map(function (nombre) {
    const p = hoy.productos[nombre];
    const resueltos = p.entregados + p.devoluciones;
    return {
      producto: nombre, pedidos: p.pedidos,
      entregados: p.entregados, devoluciones: p.devoluciones,
      tasaEntrega: resueltos ? p.entregados / resueltos * 100 : null,
      gananciaPorPedido: p.pedidos ? p.ganancia / p.pedidos : null,
      // Por debajo de la muestra mínima no se opina: se dice cuánto falta.
      señal: resueltos >= u.muestra_minima ? 'medible'
             : 'faltan ' + (u.muestra_minima - resueltos) + ' resueltos para que signifique algo',
    };
  }).sort(function (a, b) { return b.pedidos - a.pedidos; });
}

// ─── LA API Y EL CORREO ──────────────────────────────────────

/**
 * Quién puede ver esto.
 *
 * Hacen falta las dos cosas: que el PLAN del cliente incluya pauta y
 * dinero —eso lo decide Nova Central, o sea Manuela— y que la PERSONA
 * tenga el permiso dentro de su equipo —eso lo decide su dueña.
 *
 * Son preguntas distintas y por eso se responden por separado. Mezclarlas
 * llevaría a que darle un permiso a una gestora le abriera un módulo que
 * su empresa no paga, o a que pagar un módulo se lo mostrara a todo el
 * equipo.
 */
function puedeVerSemaforo(s) {
  const modulos = s.modulos || [];
  const permisos = s.permisos || [];
  if (modulos.indexOf('pauta') === -1 || modulos.indexOf('dinero') === -1) {
    return { puede: false, porque: 'plan',
      error: 'El semáforo cruza la pauta con la plata, y tu plan no incluye ' +
             'los dos módulos. Se activa desde Nova Central.' };
  }
  if (s.rol !== 'dueno' && permisos.indexOf('ver_dinero') === -1) {
    return { puede: false, porque: 'permiso',
      error: 'El semáforo muestra utilidad y gasto. Tu dueña decide quién los ve.' };
  }
  return { puede: true };
}

function apiSemaforo(s, p) {
  const v = puedeVerSemaforo(s);
  if (!v.puede) return { ok: false, error: v.error, porque: v.porque };

  const tienda = String(p.tienda || s.tiendas[0]);
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  return { ok: true, semaforo: semaforoSemanal(s.sheetId, tienda, String(p.lunes || '')) };
}

/**
 * El correo del lunes, en texto.
 *
 * Texto plano y no HTML porque esto se lee en el teléfono a las siete de
 * la mañana, y lo que importa es que las cuatro luces se vean en la
 * primera pantalla sin cargar nada.
 */
function semaforoTexto(sem) {
  const m = sem.moneda ? ' ' + sem.moneda : '';
  const L = [];
  L.push('SEMÁFORO · ' + sem.tienda);
  L.push(sem.semana.lunes + ' a ' + sem.semana.domingo);
  L.push('');

  if (!sem.hayDatos) {
    L.push('No hay ni un pedido de esa semana en la hoja.');
    L.push('');
    L.push('Eso no significa que no se vendió:');
    L.push('significa que la hoja no se actualizó.');
    L.push('');
    L.push('Sube el reporte de Dropi con el rango completo y vuelve a mirarlo.');
    L.push('Mientras no esté, este semáforo no mide nada — y prefiero decirlo');
    L.push('a mandarte unos ceros que parecen una mala semana.');
    return L.join('\n');
  }

  const icono = { verde: '🟢', amarillo: '🟡', rojo: '🔴',
                  sin_medir: '⚪', incompleto: '🟡', muestra_chica: '⚪',
                  sin_umbral: '⚪' };
  sem.luces.forEach(function (l) {
    L.push((icono[l.estado] || '⚪') + '  ' + l.etiqueta + ': ' +
           (l.valor === null ? 'sin dato' : redondear_(l.valor)));
    if (l.delta !== null && l.delta !== undefined) {
      L.push('    ' + (l.delta >= 0 ? '▲ +' : '▼ ') + redondear_(l.delta) +
             ' vs. la semana anterior');
    }
    if (l.porque) L.push('    ' + l.porque);
    L.push('');
  });

  L.push('');
  L.push('LA SEMANA');
  L.push('  Pedidos creados      ' + sem.hoy.pedidos);
  L.push('  Entregados           ' + sem.hoy.entregados);
  L.push('  Devoluciones         ' + sem.hoy.devoluciones);
  L.push('  Cancelados           ' + sem.hoy.cancelados);
  L.push('  En tránsito          ' + sem.hoy.enTransito);
  L.push('  Utilidad antes pauta ' + redondear_(sem.hoy.utilidadAntesPauta) + m);
  L.push('  Pauta                ' + (sem.hoy.hayPauta ? redondear_(sem.hoy.gasto) + m : 'sin dato'));
  L.push('  Utilidad real        ' + (sem.hoy.utilidadReal === null ? 'sin dato'
                                      : redondear_(sem.hoy.utilidadReal) + m));
  L.push('  Techo de CPA         ' + (sem.hoy.techo === null ? 'sin dato'
                                      : redondear_(sem.hoy.techo) + m));

  L.push('');
  L.push('QUÉ MIRAR');
  sem.alertas.forEach(function (a, i) {
    L.push('  ' + (i + 1) + '. ' + a.titulo);
    L.push('     ' + a.accion);
  });

  L.push('');
  L.push('Cómo se calculó: solo los entregados generan ganancia; cada devolución');
  L.push('cuesta el flete; los cancelados cuestan cero. La tasa de entrega es');
  L.push('entregados ÷ (entregados + devoluciones). El techo de CPA es la');
  L.push('utilidad antes de pauta ÷ pedidos creados. Donde no hay dato, dice');
  L.push('sin dato: nada está extrapolado.');
  return L.join('\n');
}

/**
 * Lo que corre el lunes: un semáforo por tienda, a quien pueda verlo.
 *
 * Se le manda solo a la dueña. El semáforo lleva utilidad y gasto, y
 * quién más los ve es una decisión de ella, no de Nova.
 */
function semaforoLunes() {
  const log = [];
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try {
      const ss = SpreadsheetApp.openById(c.sheetId);
      const modulos = modulosDeCliente_(c.sheetId);
      if (modulos.indexOf('pauta') === -1 || modulos.indexOf('dinero') === -1) {
        log.push(c.empresa + ': el plan no incluye pauta y dinero. Saltado.');
        return;
      }
      duenasDe_(ss).forEach(function (due) {
        (due.tiendas === '*' ? tiendasActivas_(ss) : [due.tiendas]).forEach(function (t) {
          const sem = semaforoSemanal(c.sheetId, t, '');
          try {
            MailApp.sendEmail({
              to: due.correo,
              subject: 'Semáforo de la semana · ' + t + ' · ' + sem.semana.lunes,
              body: semaforoTexto(sem),
            });
            log.push(c.empresa + ' · ' + t + ' → ' + due.correo);
          } catch (e) {
            log.push(c.empresa + ' · ' + t + ': no se pudo enviar — ' + e.message);
          }
        });
      });
    } catch (e) {
      log.push(c.empresa + ': FALLÓ — ' + e.message);
    }
  });
  const msg = log.length ? log.join('\n') : 'Ningún cliente con plan de pauta y dinero.';
  Logger.log(msg);
  return msg;
}

/** Las dueñas activas de una cuenta, con sus tiendas. */
function duenasDe_(ss) {
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (rolCanonico(d[i][c('rol')]) !== 'dueno') continue;
    if (norm(d[i][c('estado')]) === 'inactivo') continue;
    const correo = String(d[i][c('correo')] || '').trim();
    if (!correo) continue;
    out.push({ correo: correo, tiendas: String(d[i][c('tienda')] || '*').trim() || '*' });
  }
  return out;
}

function tiendasActivas_(ss) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cEst = e.indexOf('estado');
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][cId] || '').trim();
    if (!id) continue;
    if (cEst !== -1 && norm(d[i][cEst]) === 'inactiva') continue;
    out.push(id);
  }
  return out;
}

/**
 * Qué módulos incluye el plan de un cliente, por su hoja.
 *
 * Se apoya en `modulosDelPlan`, que es la misma función que usa el login:
 * si fueran dos, un cliente podría ver el semáforo por correo y no en su
 * pantalla, o al revés, y nadie sabría cuál de las dos tiene razón.
 */
function modulosDeCliente_(sheetId) {
  try {
    const shC = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
    if (!shC || shC.getLastRow() < 2) return [];
    const d = shC.getDataRange().getValues();
    const e = d[0].map(norm);
    const cSheet = e.indexOf('sheet_id'), cPlan = e.indexOf('plan');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cSheet] || '').trim() === sheetId) {
        return modulosDelPlan(d[i][cPlan]);
      }
    }
    return [];
  } catch (e) {
    return [];
  }
}
