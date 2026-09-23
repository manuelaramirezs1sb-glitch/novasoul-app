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
 * ── POR QUÉ AQUÍ NO HAY UN `estadoCanonico()` ──
 *
 * Había uno, y no lo llamaba nadie. Existía para respaldar una frase
 * que era falsa: que las pantallas «vuelven a traducir el estado al
 * leer». No lo hacían. La traducción se escribe UNA vez, al importar
 * (`estadoConOrigen`, abajo), y queda en la columna `estado_canonico`
 * del pedido.
 *
 * Esa función muerta fue la coartada de un fallo real: la dueña
 * clasificaba un estado, la pantalla decía «las cifras se están
 * rehaciendo», y los pedidos ya importados no se movían. Quien leyera
 * el código encontraba una función que parecía hacerlo.
 *
 * Ahora, cuando ella clasifica, `reaplicarEstado_()` reescribe los
 * pedidos que ya estaban y devuelve cuántos movió. Se borró la función
 * muerta para que nadie vuelva a creerle.
 */

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
