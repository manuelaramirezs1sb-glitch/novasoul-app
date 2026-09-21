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

const PERMISOS_CONOCIDOS = ['subir_pedidos', 'subir_novedades', 'subir_pauta',
                            'subir_cartera'];

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
