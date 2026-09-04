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
    'en transito':              ESTADOS.EN_TRANSITO,
    'en reparto':               ESTADOS.EN_TRANSITO,
    'en distribucion':          ESTADOS.EN_TRANSITO,
    'en oficina':               ESTADOS.EN_OFICINA,
  },

  // Effi usa el mismo formato de reporte que Mastershop.
  // Si aparece una diferencia, se separa en su propio bloque.
  effi: null, // se resuelve con el mapa de mastershop

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

/** Agrupa un motivo de novedad. Devuelve 'otro' si no cae en ninguno. */
function grupoNovedad(motivo) {
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
  CO: { cod: '57',  largo: 10, movil: /^3/ },      // Colombia
  EC: { cod: '593', largo: 9,  movil: /^9/ },      // Ecuador
  GT: { cod: '502', largo: 8,  movil: /^[3-5]/ },  // Guatemala
  MX: { cod: '52',  largo: 10, movil: /^[1-9]/ },  // México
  PE: { cod: '51',  largo: 9,  movil: /^9/ },      // Perú
  CL: { cod: '56',  largo: 9,  movil: /^9/ },      // Chile
  ES: { cod: '34',  largo: 9,  movil: /^[67]/ },   // España
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

  s = s.replace(/^0+/, ''); // ceros de marcación nacional

  const def = PREFIJOS[paisDefault];

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
    if (s.indexOf(c.cod) === 0 && s.length === c.cod.length + c.largo) return s;
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
