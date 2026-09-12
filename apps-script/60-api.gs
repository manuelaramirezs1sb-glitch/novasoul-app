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

const TTL_SESION_H = 12;   // horas que dura un token
const TTL_CODIGO_M = 10;   // minutos que dura el código de 6 dígitos

// ─── ENTRADA ─────────────────────────────────────────────────

function doGet(e)  { return manejar(e, 'GET'); }
function doPost(e) { return manejar(e, 'POST'); }

function manejar(e, metodo) {
  try {
    const p = leerParams(e, metodo);
    const accion = String(p.accion || '').trim();
    if (!accion) return json({ ok: false, error: 'Falta la acción.' });

    // Las únicas dos que no piden token
    if (accion === 'login')     return json(apiLogin(p));
    if (accion === 'verificar') return json(apiVerificar(p));

    const s = sesion(p.token);
    if (!s) return json({ ok: false, error: 'Sesión vencida o inválida.', reautenticar: true });

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
      case 'auditoria': return json(apiAuditoria(s, p));
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
    vence: Date.now() + TTL_SESION_H * 3600000,
  };
  cache.put('ses_' + token, JSON.stringify(s), TTL_SESION_H * 3600);

  registrarMovimiento(s, 'Equipo', persona.id, 'ultima_conexion', '', ahoraISO());
  return { ok: true, token: token, sesion: publico(s) };
}

function sesion(token) {
  if (!token) return null;
  const raw = CacheService.getScriptCache().get('ses_' + String(token));
  if (!raw) return null;
  const s = JSON.parse(raw);
  if (s.vence <= Date.now()) return null;

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
};

const PERMISOS_CONOCIDOS = ['subir_pedidos', 'subir_novedades', 'subir_pauta'];

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
              permisos: s.permisos || [] };
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
  Inventario: ['tienda', 'sku', 'producto', 'stock', 'costo_unitario', 'precio',
               'minimo', 'origen', 'nota'],
  Equipo:     ['nombre', 'correo', 'rol', 'tienda', 'estado', 'permisos'],
};

/** Quién puede crear o quitar en cada entidad. */
const SOLO_DUENO = ['Equipo', 'Gastos'];

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
  if (!sh || sh.getLastRow() < 2) return { ok: true, personas: [] };

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

  personas.forEach(function (x) {
    x.efectividad = x.pedidos ? x.entregados / x.pedidos * 100 : 0;
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
  const fila = enc.map(function (col) {
    if (col === 'id') return id;
    if (col === 'activo') return 'si';
    if (col === 'actualizado_en') return ahoraISO();
    if (col === 'actualizado_por') return s.email;
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

  const escritos = [], rechazados = [];
  Object.keys(campos).forEach(function (k) {
    const col = enc.indexOf(norm(k));
    if (col === -1) { rechazados.push(k + ' (no existe)'); return; }
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
    if (sh) sh.appendRow([ahoraISO(), s.email, entidad, entidadId, campo, antes, ahora]);
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
  return { ok: true, tienda: tienda, mes: mes, datos: d };
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
      if (est === 'entregado')   { out.entregados++; out.ventas += num(f[c('valor')]); }
      if (est === 'devolucion')  out.devueltos++;
      if (est === 'cancelado')   out.cancelados++;
      if (['cancelado','pendiente'].indexOf(est) === -1) out.despachados++;

      out.costoProducto += num(f[c('costo_producto')]);
      out.costoEnvio    += num(f[c('costo_envio')]);

      /**
       * El flete de una devolución se paga igual, y a veces doble. Va
       * aparte del flete de las entregas porque son dos cosas distintas:
       * uno es costo de vender, el otro es costo de no haber vendido.
       */
      if (est === 'devolucion') out.costoDevolucion += num(f[c('costo_envio')]);

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
    out.gasto = 0; out.campanas = {};
    if (shPa && shPa.getLastRow() > 1) {
      const datos = shPa.getDataRange().getValues();
      const e = datos[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < datos.length; i++) {
        const f = datos[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        const fecha = aISO(f[c('fecha')], 'UTC');
        if (!fecha || fecha.slice(0, 7) !== mes) continue;
        const g = num(f[c('gasto_normalizado')]) || num(f[c('gasto')]);
        out.gasto += g;
        const nom = String(f[c('campana')] || 'Sin nombre').trim();
        if (!out.campanas[nom]) out.campanas[nom] = { gasto: 0, resultados: 0 };
        out.campanas[nom].gasto += g;
        out.campanas[nom].resultados += num(f[c('resultados')]);
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

    out.cpa  = out.entregados ? out.gasto / out.entregados : 0;
    out.roas = out.gasto ? out.ventas / out.gasto : 0;
    // Margen: antes de los gastos fijos. Utilidad: lo que queda de verdad.
    out.margen = out.ventas - out.gasto - out.costoProducto - out.costoEnvio;
    out.utilidad = out.margen - out.fijos;
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
    return {
      cerrado_en: d[i][c('cerrado_en')],
      datos: {
        pedidos: num(d[i][c('pedidos')]), entregados: num(d[i][c('entregados')]),
        devueltos: num(d[i][c('devueltos')]), ventas: num(d[i][c('ventas')]),
        gasto: num(d[i][c('gasto')]), margen: num(d[i][c('margen')]),
        efectividad: num(d[i][c('efectividad')]),
        pendientes: num(d[i][c('pendientes_al_cierre')]),
        despachados: 0, cancelados: 0, novedades: 0, sinMover: 0,
        grupos: {}, transportadoras: {}, congelado: true,
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
  sh.appendRow([tienda, mes, 'cerrado', ahoraISO(), s.email, d.pendientes,
                d.pedidos, d.entregados, d.devueltos, d.ventas,
                d.gasto || 0, d.margen || 0, d.efectividad,
                p.nota || '']);
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
      necesita === 'subir_pauta'
        ? 'No tienes permiso para subir pauta. La dueña lo activa en Permisos.'
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
    return SpreadsheetApp.openById(id).getSheets()[0].getDataRange().getValues();
  } finally {
    if (id) { try { DriveApp.getFileById(id).setTrashed(true); } catch (e) {} }
  }
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
