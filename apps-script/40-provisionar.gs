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
