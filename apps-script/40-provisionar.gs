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

const CARPETA_NOVA = '1lxpyEhj3dfdwgdpotL7e8G8gCd_EaB7o';

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

  const central = SpreadsheetApp.openById(IDS.central);
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
  const carpeta = DriveApp.getFolderById(CARPETA_NOVA);
  const copia = DriveApp.getFileById(IDS.empresarial)
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

/** Lista los clientes registrados y a qué hoja apunta cada uno. */
function listarClientes() {
  const sh = SpreadsheetApp.openById(IDS.central).getSheetByName('Clientes');
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
