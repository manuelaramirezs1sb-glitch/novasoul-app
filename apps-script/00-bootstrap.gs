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

/**
 * ── ABRIR UN LIBRO UNA SOLA VEZ POR EJECUCIÓN ──
 *
 * `libro_()` parece gratis y no lo es: cada llamada es
 * una ida y vuelta a los servidores de Google. Se midió al responder
 * «¿por qué NovaSoul está tan lento?» y el número fue feo: pintar la
 * pantalla de entrada abría los libros SESENTA veces, casi todas el
 * mismo libro, porque cada función que necesitaba una pestaña lo abría
 * por su cuenta.
 *
 * Esto guarda el manejador mientras dura la ejecución. No es un caché
 * de DATOS —eso sí sería peligroso—: es el mismo objeto vivo de Google,
 * así que lo que una función escriba, la siguiente lo lee. Lo único que
 * se ahorra es volver a pedir la llave de una puerta que ya está
 * abierta.
 *
 * El objeto muere cuando muere la ejecución, que en Apps Script son
 * segundos. No hay nada que invalidar.
 */
var LIBROS_ABIERTOS_ = {};

/** Soltar los manejadores. Se llama al empezar cada petición. */
function libroOlvidar_() { LIBROS_ABIERTOS_ = {}; }

function libro_(id) {
  const k = String(id || '');
  if (!k) throw new Error('Me pidieron abrir un libro sin decirme cuál.');
  if (!LIBROS_ABIERTOS_[k]) LIBROS_ABIERTOS_[k] = SpreadsheetApp.openById(k);
  return LIBROS_ABIERTOS_[k];
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
            'razon_cancelacion','estado_nova','nota',
            /**
             * `nota` es del equipo y el importador no la toca.
             * `observacion` es lo que traía el archivo: Dropi la llama
             * «notas», y ahí es donde las gestoras escribían antes de
             * Nova. Sin esta columna el importador la leía —solo para
             * pescar un segundo teléfono— y la botaba.
             */
            'observacion',
            /**
             * ── QUIÉN LO HIZO ≠ QUIÉN LO VE ──
             *
             * `gestora_asignada` hace DOS trabajos que no son el mismo:
             * dice quién atendió el pedido, y decide quién puede verlo.
             * Mientras el equipo esté todo en Nova eso funciona.
             *
             * Deja de funcionar en cuanto llega un histórico. En el
             * archivo de una tienda real la columna GESTIONA trae JAIME
             * (399 pedidos), APOYO (68), ZULAY (48)… gente que no tiene
             * cuenta, y «APOYO» que ni siquiera es una persona. Si eso
             * cae en `gestora_asignada`, esos pedidos quedan asignados a
             * nadie que pueda entrar: NO LOS VE NINGUNA PERSONA.
             *
             * Así que se parten en dos. `gestionado_por` es una
             * etiqueta: texto libre, sin cuenta, para el registro y las
             * estadísticas por agente. `gestora_asignada` sigue siendo
             * control de acceso y solo acepta a alguien del Equipo.
             *
             * Se puede importar la historia entera sin crear una sola
             * cuenta, y el día que se creen, enlazar una etiqueta con
             * una persona no rompe nada de lo ya importado.
             */
            'gestionado_por',
            'ultimo_movimiento',
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
  /**
   * ── DOS COLUMNAS PARA LO MISMO, A PROPÓSITO ──
   *
   * `solucion` es lo que el equipo escribe DENTRO de Nova, y el
   * importador tiene prohibido pisarla.
   *
   * `solucion_plataforma` es lo que venía escrito en el archivo —lo que
   * la gestora escribió en Dropi o en Effi antes de que existiera Nova—.
   * Esa SÍ la actualiza cada importación, porque es un espejo de la
   * plataforma, no un dato de aquí.
   *
   * Antes solo existía la primera, y el resultado era que el histórico
   * de soluciones del archivo se leía y se tiraba: se importaban
   * setenta novedades y las setenta llegaban sin una palabra de lo que
   * ya se había hecho. Empezar de cero encima de trabajo que ya estaba
   * hecho es la peor forma de estrenar una herramienta.
   */
  Novedades: ['id','fuente','id_externo','pedido_id','fecha','tipo','motivo','grupo',
              'estado','solucionada','fecha_solucion','desenlace',
              'gestora','solucion','nota','intentos','resuelta_en',
              'solucion_plataforma','aclaracion','gestionado_por',
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
  /**
   * Las cinco últimas columnas se agregaron al leer el control real de
   * una tienda: su hoja de CAS trae CLIENTE, TELEFONO, la FECHA DE
   * ENVIO del pedido, lo que se hizo (GESTION) y quién lo hizo.
   *
   * `cliente` y `telefono` parecen redundantes —están en el pedido—
   * pero no lo son: un CAS puede llegar de un pedido que nunca se
   * importó, y sin el nombre ese reclamo no se puede ni buscar.
   */
  CAS: ['id','tienda','pedido_id','id_externo','guia','transportadora',
        'abierto_en','abierto_por','ticket','estado','dias_quieto',
        'ultima_gestion','respuesta','cerrado_en','nota',
        'cliente','telefono','fecha_envio','gestion','gestionado_por'],

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
   * Los veredictos de auditoría. Uno por novedad revisada.
   *
   * Existe porque la pantalla de Auditoría tenía un botón de «Guardar
   * veredicto» sin nada detrás: se marcaba el hallazgo, se escribía la
   * nota, se guardaba, y no había ningún sitio donde eso cayera.
   *
   * `senales` guarda lo que Nova había levantado sola en el momento de
   * revisar. Sirve para lo único que importa un mes después: saber si
   * el veredicto se puso mirando los mismos datos que hay hoy.
   */
  Auditorias: ['id','tienda','novedad_id','pedido_id','gestora','veredicto',
               'nota','senales','auditor','creada_en'],

  /**
   * El chat del equipo, dentro de Nova.
   *
   * Una sola hoja para los dos tipos de conversación, y la diferencia la
   * marca `para`:
   *   vacío  · es del grupo de `tienda` — lo lee quien tenga esa tienda
   *   correo · es privado entre `de` y `para`, y no lo lee nadie más
   *
   * `de_nombre` se guarda junto al correo a propósito, aunque esté en
   * Equipo. Un mensaje tiene que seguir diciendo quién lo escribió cuando
   * esa persona ya no esté en el equipo; si el nombre se buscara al
   * pintar, el historial se quedaría sin autores al primer cambio de
   * personal — y un chat de trabajo sirve justamente para eso.
   *
   * `leido_por` son correos separados por comas. `borrado` marca el
   * mensaje sin quitar la fila: queda la constancia de que hubo algo ahí.
   */
  Mensajes: ['id','tienda','de','de_nombre','para','texto','creado_en',
             'leido_por','borrado'],

  /**
   * Las dos puertas que una gestora abre cada mañana. Una fila por tienda.
   *
   *   url/usuario/clave  la PLATAFORMA donde se gestiona (Dropi, Shopify)
   *   correo_codigo      a dónde llega el código de verificación
   *   canal_*            el canal donde la tienda habla con SUS COMPRADORES
   *                      (Chat Center, WhatsApp) — no el chat del equipo,
   *                      que vive dentro de Nova en la hoja Mensajes
   *
   * `clave` guarda la contraseña, y eso es deliberado. Ella lo explicó:
   * es una subcuenta creada dentro de la plataforma que solo confirma y
   * gestiona —no puede sacar dinero— y además pide un código que llega a
   * `correo_codigo`, así que sin ese buzón no entra nadie.
   *
   * Hoja aparte y NO un Parámetro más, por dos razones que no dependen de
   * que nadie se acuerde: `apiParametros` registra cada cambio en
   * Movimientos con el valor viejo y el nuevo (la contraseña quedaría en
   * texto plano para siempre), y `ajustes` viaja a la pantalla en cada
   * carga del día. Ver 97-accesos.gs.
   */
  Accesos: ['tienda','plataforma','url','usuario','clave','correo_codigo',
            'canal_nombre','canal_url',
            /**
             * CÓMO LE PAGA LA CLIENTA.
             *
             * «no veo (…) la información para hacer los abonos o añadir
             *  un link de pago por PayPal o si tiene otra pasarela».
             *
             * `pago_url` es el enlace que la gestora le manda a la
             * compradora —PayPal, Wompi, un link de Nequi— y `pago_datos`
             * es lo que se dicta por teléfono cuando no hay enlace: el
             * banco, el número de cuenta y a nombre de quién.
             *
             * Son dos campos y no uno porque se usan en momentos
             * distintos: el enlace se pega en el chat, los datos se leen
             * en voz alta. Juntos en un solo campo, quien confirma tiene
             * que leer todo el bloque buscando la línea que necesita.
             */
            'pago_nombre','pago_url','pago_datos',
            'nota','actualizado_en','actualizado_por'],

  /**
   * La bitácora de intentos: una fila por nota, nunca una por caso.
   *
   * Existe porque había UNA columna `nota` por pedido y una caja de
   * texto encima: quien anotaba el intento del jueves abría la caja con
   * lo del miércoles dentro y escribía encima. El intento del miércoles
   * desaparecía sin error y sin aviso.
   *
   * Ella: «deben dejar la nota por cada intento de contacto, deben
   * aparecer todas, con fecha y nombre de quien puso la nota».
   *
   * `entidad` es Pedidos, Novedades o CAS. `autor_nombre` se guarda
   * junto al correo por la misma razón que en Mensajes: la nota tiene
   * que seguir diciendo quién la escribió cuando esa persona ya no esté
   * en el equipo.
   */
  Notas: ['id','tienda','entidad','entidad_id','texto','autor','autor_nombre','creado_en'],

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

  /**
   * ═══════════════════════════════════════════════════════════
   *  EL LADO DE ELLA: SUS TRABAJOS Y SU PLATA
   * ═══════════════════════════════════════════════════════════
   *
   * Nova Central no es solo la consola del producto: es el centro de
   * trabajo de Manuela, que además de Nova tiene clientes propios, la
   * universidad y un empleo. Todo eso tiene fechas y tiene plata, y
   * hasta hoy vivía en la cabeza.
   *
   * Vive AQUÍ y no en NovaSoul a propósito. Central se abre delante de
   * otra gente —una socia, un contador— y estas son cifras de negocio.
   * NovaSoul guarda lo que no se le muestra a nadie.
   */

  /**
   * Cada compromiso con nombre propio: Nutrea, PHH, Son de Sky, la U.
   *
   * `tipo` separa lo que se cobra de lo que no. La universidad no paga,
   * pero ocupa las mismas horas que un cliente que sí — y una lista que
   * solo mira lo que factura deja fuera justo lo que no se puede
   * incumplir.
   */
  /**
   * Un proyecto, visto por el lado del negocio.
   *
   * `mi_rol` no es lo mismo que `tipo`. El tipo dice qué es —cliente,
   * empleo, propio, universidad—; el rol dice qué es ELLA adentro: si
   * es socia se lleva un porcentaje, si es trabajadora le pagan, y si
   * es propio no hay a quién cobrarle. Con una sola columna había que
   * elegir, y las dos preguntas se responden distinto.
   *
   * `modalidad` dice CÓMO entra la plata: un precio fijo, un porcentaje
   * de la utilidad, por hora, o nada. Cuando es porcentaje, `tienda_id`
   * apunta a la tienda de Nova Empresarial de la que sale ese número, y
   * `base_porcentaje` dice sobre qué se calcula. Sin esas tres, un 50%
   * es un número suelto que no se puede verificar contra nada.
   *
   * `confidencial` decide si el TEXTO de las tareas sube a esta
   * pantalla o se queda en NovaSoul. Nace encendido para los empleos,
   * porque lo de PHH es confidencial y ella lo dejó dicho: Central se
   * abre delante de una socia o un contador, y un olvido no puede ser
   * lo único que proteja eso.
   */
  /**
   * `padre_id` deja que un trabajo contenga proyectos.
   *
   * Existe por un caso concreto: Upwork es UN trabajo —una forma de
   * cobro, una moneda, una quincena— con muchos encargos dentro, cada
   * uno con sus tareas y su fecha. Sin esto, cada encargo sería un
   * trabajo suelto y no habría forma de ver cuánto lleva ese cliente en
   * total sin sumarlo a mano.
   *
   * Está vacío en casi todos, y así debe ser: la jerarquía existe donde
   * hace falta, no en todas partes. PHH y Nutrea son uno solo.
   */
  Trabajos: ['id','nombre','contraparte','tipo','estado','moneda',
             'valor_acordado','forma_cobro','fecha_inicio','fecha_entrega',
             'horas_semana','especificacion','documento','nota',
             'mi_rol','modalidad','porcentaje','base_porcentaje',
             'cliente_id','tienda_id','confidencial','padre_id',
             // Cada cuánto paga, y a los cuántos días del trabajo hecho
             // entra la plata. Upwork: quincenal, 7 días.
             'periodicidad','dias_pago'],

  /**
   * De dónde sale la información profunda de cada proyecto.
   *
   * Solo ENLACES y notas: el Excel, el PDF, el Word, la presentación o
   * el artefacto de Claude se quedan donde están. Es la misma regla que
   * con las materias y con PHH — Nova sabe dónde están las cosas, no
   * guarda una copia de lo que dicen.
   */
  Fuentes: ['id','trabajo_id','nombre','tipo','enlace','nota','agregado_en'],

  /**
   * Lo que debe entrar, con su fecha ESPERADA aparte de la real.
   *
   * Sin la esperada no se puede saber que algo está atrasado: solo se
   * sabe que no ha llegado, que es otra cosa. Esa diferencia es la que
   * permite llamar a tiempo en vez de darse cuenta a fin de mes.
   */
  Cobros: ['id','trabajo_id','concepto','monto','moneda',
           'fecha_esperada','fecha_cobrada','estado','nota'],

  /**
   * Su plata personal, movimiento por movimiento.
   *
   * `cuenta` existe porque no es lo mismo tener el mes cubierto que
   * tenerlo cubierto EN la cuenta de donde se paga. Un ingreso en
   * dólares que todavía está en PayPal no paga un arriendo mañana.
   */
  Finanzas: ['id','fecha','flujo','categoria','concepto','monto','moneda',
             'cuenta','recurrente','trabajo_id','nota'],

  /**
   * Deudas y ahorros van juntos porque son la misma cosa con el signo
   * cambiado: un saldo que se mueve hacia una meta y una fecha.
   */
  Metas: ['id','tipo','nombre','con_quien','monto_meta','saldo','moneda',
          'cuota','dia_del_mes','fecha_meta','estado','nota'],
};

const ESQUEMA_SOUL = {
  Usuarios: ['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
             'lugar_nacimiento','zona_horaria','acento','modo','idioma'],
  /**
   * Un pendiente de NovaSoul.
   *
   * `trabajo_id` apunta a la hoja Trabajos de Nova_Central: es el mismo
   * proyecto visto por el otro lado. Central sabe cuánto vale y cuándo
   * se cobra; Soul sabe qué hay que entregar y cuánto cuesta en horas.
   *
   * `horas_estimadas` las pone ella. `horas_reales` se llenan al
   * cerrarlo, y son las que un día van a corregir sus estimaciones —
   * por eso son dos columnas y no una.
   *
   * `riesgo` no es la prioridad: es qué pasa si NO se entrega. Un
   * parcial es inamovible; un ajuste de carta se corre una semana. Sin
   * esa diferencia, elegir qué se cae es adivinar.
   */
  Pendientes: ['id','usuario_id','texto','tipo','origen','fecha','hecho',
               'hecho_en','plataforma_id','trabajo_id','estado','prioridad',
               'horas_estimadas','horas_reales','riesgo','nota','materia_id'],

  /**
   * Las materias del semestre.
   *
   * Una por materia y no una por semestre, porque cada una tiene su
   * profesor, su carpeta y sus fechas. Todas suben a Central como UN
   * solo proyecto —«Universidad»—, que es como pesa en su semana.
   *
   * `carpeta` es un ENLACE a Drive. Los archivos no se copian aquí:
   * misma regla que con PHH. Nova guarda dónde están, no qué dicen.
   */
  Materias: ['id','usuario_id','nombre','codigo','profesor','carpeta',
             'semestre','trabajo_id','estado','nota'],

  /**
   * Lo que se repite todas las semanas: turnos y clases.
   *
   * Se escribe UNA vez y se repite solo. Una fila por día de la semana
   * —el turno del viernes y el del sábado son dos filas— porque cada
   * uno tiene su horario y su paga.
   *
   * `desde` y `hasta` acotan el semestre o el contrato. Sin ellos, una
   * clase de este semestre seguiría apareciendo en marzo del año
   * entrante y la semana diría que no cabe por una materia que ya pasó.
   *
   * `paga_fija` es lo que vale el turno. Las propinas NO van aquí: son
   * distintas cada día y se escriben al día siguiente, en Turnos.
   */
  Rutina: ['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin',
           'lugar','trabajo_id','materia_id','paga_fija','moneda','desde','hasta',
           'activo','nota'],

  /**
   * Cada turno que de verdad se trabajó, con sus propinas.
   *
   * La rutina dice «los viernes»; esto dice «el viernes 25 fueron
   * 80.000 de base y 34.000 de propinas». Sin esta hoja, el mes
   * cuadraría con la paga fija y nunca con lo que entró.
   *
   * `finanza_id` es el enganche con el movimiento que se escribió en
   * Nova_Central: existe para no escribirlo dos veces si ella corrige
   * las propinas después.
   */
  Turnos: ['id','usuario_id','rutina_id','fecha','paga','propinas','moneda',
           'estado','finanza_id','nota'],

  /**
   * Cuántas horas libres tiene cada día de la semana.
   *
   * Es el techo contra el que se compara todo lo demás, y no se puede
   * adivinar: descontados turnos de Salsabor y clases, lo que queda
   * solo lo sabe ella. Mientras esta hoja esté vacía, NovaSoul dice que
   * falta el dato en vez de inventar una semana de 24 horas.
   */
  Horas: ['usuario_id','dia_semana','horas_libres','nota'],

  /**
   * Mindlab: doce semanas, una tarea cada una.
   *
   * Vive en su propia hoja y no suelta en Pendientes porque el plan es
   * una cosa y la tarea de esta semana es otra. De aquí baja a
   * Pendientes de a una, cuando toca.
   */
  Mindlab: ['id','usuario_id','semana','mes','tema','tarea','horas_estimadas',
            'desde','hasta','estado','nota'],

  /**
   * Los gastos fijos del mes: el PLAN, no lo que pasó.
   *
   * Arriendo, mercado, servicios, internet, crédito, deudas, móvil,
   * varios y ahorro. Lo que de verdad salió vive en Finanzas, en
   * Nova_Central, y esta pantalla compara una cosa contra la otra.
   *
   * Están separados a propósito: si el plan y el movimiento fueran la
   * misma fila, no habría forma de ver que el mercado se pasó — el
   * presupuesto se habría reescrito solo para darse la razón.
   */
  Fijos: ['id','usuario_id','categoria','concepto','monto','moneda',
          'dia_del_mes','activo','nota',
          /**
           * ── LO QUE ELLA RECLAMÓ DE ESTA PANTALLA ──
           *
           * «debería haber una distinción si es único pago, pago
           *  mensual, pago por cuotas; las cuotas y deudas deben estar
           *  en otra parte, separada; los ingresos y los gastos fijos
           *  también deben estar separados».
           *
           * Cinco columnas nuevas, y cada una responde una de esas:
           *
           *   flujo        · entra o sale. Sin esto, un sueldo fijo y
           *                  el arriendo caían en la misma lista.
           *   tipo_pago    · mensual, cuotas o único.
           *   cuotas_total · cuántas son en total.
           *   cuotas_pagadas· cuántas van.
           *   cuota_desde  · el mes de la primera (AAAA-MM), que es lo
           *                  único que permite decir CUÁNDO TERMINA.
           *   acreedor     · a quién se le debe. Ella tiene dos deudas
           *                  con nombre propio y la pantalla las
           *                  mostraba como una sola línea «Deudas».
           */
          'flujo','tipo_pago','cuotas_total','cuotas_pagadas','cuota_desde','acreedor',
          /**
           * `deuda_total` es lo que se debe EN TOTAL, no la cuota.
           *
           * Ella lo pidió así: «debes pedir la deuda total, el tiempo
           * de meses, y dar la oportunidad de cambiar el aporte cada
           * que vaya a subir un pago si pagué más o menos».
           *
           * Sin este número, lo único que se podía hacer era multiplicar
           * la cuota por las cuotas — que solo es verdad si paga
           * exactamente lo mismo todos los meses. Y ella dice que no.
           */
          'deuda_total'],
  /**
   * Cada abono a una deuda, con LO QUE DE VERDAD PAGÓ ese mes.
   *
   * Va en su propia hoja y no en una columna de Fijos porque son
   * muchos por deuda y cambian de monto. Con esto, «cuánto llevas
   * pagado» deja de ser una multiplicación optimista y pasa a ser una
   * suma de hechos.
   */
  Pagos: ['id','usuario_id','fijo_id','fecha','monto','moneda','nota'],
  Dias: ['usuario_id','fecha','comidas_marcadas','movimiento_hecho','puntos',
         'cerrado','cerrado_en','perdonado'],
  Recompensas: ['id','usuario_id','nombre','costo_puntos','canjeada','canjeada_en'],
  /**
   * La carta natal. Una fila por cuerpo.
   *
   * No se calcula: se carga. Ella la saca de Horus, que ya le da todo
   * bien, y Nova la guarda para poder cruzarla con su semana. Calcular
   * efemérides aquí sería rehacer mal algo que ya está bien hecho.
   */
  Carta: ['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota'],

  /**
   * Los tránsitos, CON SU DURACIÓN.
   *
   * `desde` y `hasta` no son un adorno: la Luna dura dos días y medio y
   * Saturno meses, y un consejo que no distingue eso es ruido. Un
   * tránsito sin fechas no se puede cruzar con una semana.
   */
  /**
   * Los tránsitos, con SUS DOS CASAS.
   *
   * `casa` es de casas enteras —un signo, una casa— que es como se
   * hacen las profecciones y lo que el resto del código usa.
   * `casa_placidus` es el sistema que usa Horus.
   *
   * Van las dos porque NO dan lo mismo: contra las diez casas de su
   * carta de Horus, Placidus acierta 10 de 10 y casas enteras 6 de 10.
   * Y en cinco años de sus tránsitos difieren en 65 de 135. Elegir uno
   * en silencio la habría dejado viendo un número que no cuadra con su
   * app, sin saber por qué. Ella dijo: «muéstrame las dos y yo decido».
   */
  Transitos: ['usuario_id','fecha','casa','casa_placidus','tema','intensidad_pct',
              'texto_transito','por_que','como_trabajarlo','el_otro_lado',
              'cuerpo','aspecto','a_natal','desde','hasta','fuente'],

  /**
   * El pensum kármico: qué pide cada temporada.
   *
   * Es suyo y lo carga ella. Nova no dice qué significa un tránsito —
   * dice qué ventana está abierta hoy, según lo que ella escribió, y la
   * cruza con lo que tiene que entregar.
   */
  /**
   * `casa_alterna` guarda la casa del OTRO sistema cuando los dos no
   * coinciden. No es un adorno: la casa decide el momento —angular
   * cambiar, sucedente descansar, cadente aprender—, así que con dos
   * casas hay dos consejos distintos para el mismo tránsito. Guardar
   * solo uno sería decidir por ella y que no se entere.
   */
  Pensum: ['id','usuario_id','desde','hasta','titulo','cuerpo','casa','casa_alterna',
           'momento','que_pide','que_evitar','nota'],

  /**
   * La revolución solar: su año, de cumpleaños a cumpleaños.
   *
   * La ventana la calcula Nova —es aritmética de calendario—. La carta
   * de ese año la trae ella de Horus.
   */
  Revolucion: ['usuario_id','anio','desde','hasta','ascendente','casa_sol',
               'tema','texto','nota'],
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
  const central = libro_(IDS_().central).getSheetByName('Clientes');
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
  const ss = libro_(fileId);
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
  const sh = libro_(IDS_().empresarial).getSheetByName('Parametros');
  if (!sh || sh.getLastRow() > 1) return; // ya sembrado
  sh.getRange(2, 1, PARAMETROS_DEFAULT.length, 5).setValues(PARAMETROS_DEFAULT);
}

// La siembra de tiendas NO va aquí: vive en crearCliente(), porque la
// operación de Nova también es una copia del template. Si se sembrara
// aquí, cada cliente nuevo nacería con las tiendas de Nutrea adentro.
