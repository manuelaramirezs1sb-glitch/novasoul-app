/**
 * ═══════════════════════════════════════════════════════════════
 *  EL CHAT DEL EQUIPO · dentro de Nova
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ QUÉ PIDIÓ ───────────────────────────────────────────────┐
 * │                                                            │
 * │ «el dueño, la admin y las gestoras deben tener un mismo    │
 * │  chat en Nova (…) que pueda hablar con su equipo completo  │
 * │  o solo dejarle un mensaje al admin o gestor en privado,   │
 * │  dentro de Nova, no por fuera».                            │
 * │                                                            │
 * │ Hasta hoy la pestaña «Nova Chat» era un BOTÓN que abría    │
 * │ un grupo de WhatsApp. Eso es lo contrario de lo que pidió: │
 * │ mandaba la conversación afuera, donde Nova no sabe de qué  │
 * │ se habló ni quién quedó de hacer qué.                      │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ EL GRUPO ES POR TIENDA ───────────────────────────┐
 * │                                                            │
 * │ Porque el acceso ya es por tienda. Una gestora de Colombia │
 * │ no ve los pedidos de Guatemala; tampoco tiene por qué leer │
 * │ lo que se habla de esa tienda.                             │
 * │                                                            │
 * │ Una dueña con tres tiendas tiene tres grupos, y eso no es  │
 * │ un inconveniente: es que son tres equipos distintos que    │
 * │ hablan de cosas distintas.                                 │
 * │                                                            │
 * │ El privado es aparte: va entre dos personas del mismo      │
 * │ cliente, sin importar la tienda. Un mensaje a alguien es   │
 * │ para esa persona, no para el sitio donde trabaja.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE NOVA NO HACE AQUÍ ─────────────────────────────────┐
 * │                                                            │
 * │ No edita ni borra mensajes de nadie, ni los de uno mismo.  │
 * │ Un chat de trabajo donde se puede reescribir lo dicho no   │
 * │ sirve para lo único que sirve un chat de trabajo: saber    │
 * │ qué se acordó.                                             │
 * │                                                            │
 * │ Se pueden ocultar los propios —queda la marca de que       │
 * │ hubo un mensaje y quién lo borró—, y eso es todo.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Cuánto texto cabe en un mensaje. Un pegote de mil líneas no es un chat. */
const CHAT_LARGO = 2000;

/** Cuántos mensajes trae una conversación de una vez. */
const CHAT_PAGINA = 80;

/**
 * Con quién se puede hablar: el equipo del mismo cliente.
 *
 * Se incluye a las personas de OTRAS tiendas a propósito. La admin
 * coordina varias, y poder escribirle a la gestora de Guatemala sin
 * tener que estar parada en Guatemala es justamente para lo que sirve
 * un privado.
 */
function chatGente_(ss, yo) {
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const correo = String(d[i][c('correo')] || '').toLowerCase().trim();
    const nombre = String(d[i][c('nombre')] || '').trim();
    if (!correo || !nombre) continue;
    if (norm(d[i][c('estado')]) === 'inactivo') continue;
    if (correo === norm(yo)) continue;          // no se habla solo
    out.push({
      correo: correo, nombre: nombre,
      rol: rolCanonico(d[i][c('rol')]) || String(d[i][c('rol')] || ''),
      tiendas: String(d[i][c('tienda')] || '*').trim(),
    });
  }
  return out;
}

/**
 * Si esta persona puede leer esta conversación.
 *
 * Se comprueba al LEER y al ESCRIBIR, no solo al pintar la lista. Una
 * pantalla que no ofrece una conversación no impide pedirla: basta
 * cambiar un valor en la consola del navegador.
 */
function chatPuedeVer_(s, con) {
  const quien = String(con || '').trim();
  if (!quien) return false;

  // Grupo de una tienda: «g:ec»
  if (quien.indexOf('g:') === 0) {
    const t = quien.slice(2);
    return !!t && (s.tiendas || []).indexOf(t) !== -1;
  }
  // Privado: el correo de la otra persona.
  return quien.indexOf('@') !== -1;
}

/** El id de conversación en el que cae un mensaje, visto desde `yo`. */
function chatHiloDe_(m, yo) {
  if (!String(m.para || '').trim()) return 'g:' + String(m.tienda || '');
  const de = norm(m.de), para = norm(m.para);
  return de === norm(yo) ? para : de;
}

/**
 * Las conversaciones de esta persona, con lo último de cada una.
 *
 * Una sola lectura de la hoja para todo: la lista, los sin leer y el
 * último mensaje de cada hilo. Tres recorridos de la misma hoja para
 * tres números que salen del mismo sitio es lo que vuelve lento algo
 * que se abre veinte veces al día.
 */
function apiChat(s, p) {
  const yo = String(s.email || '').toLowerCase().trim();
  const ss = libro_(s.sheetId);
  const gente = chatGente_(ss, yo);

  const out = {
    ok: true, yo: yo, yoNombre: s.nombre || yo,
    gente: gente,
    grupos: (s.tiendas || []).map(function (t) {
      return { id: 'g:' + t, tienda: t, nombre: nombreTienda(ss, t) || t };
    }),
    hilos: {}, mensajes: [], con: '', sinLeer: 0,
  };

  const sh = ss.getSheetByName('Mensajes');
  if (!sh || sh.getLastRow() < 2) {
    out.con = String(p.con || out.grupos[0] && out.grupos[0].id || '');
    return out;
  }

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  const mios = [];
  for (let i = 1; i < d.length; i++) {
    const m = {
      id: String(d[i][c('id')] || ''),
      tienda: String(d[i][c('tienda')] || '').trim(),
      de: String(d[i][c('de')] || '').toLowerCase().trim(),
      deNombre: String(d[i][c('de_nombre')] || ''),
      para: String(d[i][c('para')] || '').toLowerCase().trim(),
      texto: String(d[i][c('texto')] || ''),
      cuando: String(d[i][c('creado_en')] || ''),
      leidoPor: String(d[i][c('leido_por')] || ''),
      borrado: norm(d[i][c('borrado')]) === 'si',
    };
    if (!m.id) continue;

    /**
     * Qué le toca a esta persona:
     *  · un privado en el que es una de las dos puntas;
     *  · o un mensaje de grupo de una tienda suya.
     * Cualquier otra cosa no se le manda — ni para contarla.
     */
    const esPrivado = !!m.para;
    if (esPrivado) {
      if (m.de !== yo && m.para !== yo) continue;
    } else {
      if ((s.tiendas || []).indexOf(m.tienda) === -1) continue;
    }
    mios.push(m);
  }

  mios.sort(function (a, b) { return a.cuando < b.cuando ? -1 : 1; });

  const leido = function (m) {
    return m.de === yo ||
           (',' + m.leidoPor.toLowerCase() + ',').indexOf(',' + yo + ',') !== -1;
  };

  mios.forEach(function (m) {
    const hilo = chatHiloDe_(m, yo);
    if (!out.hilos[hilo]) out.hilos[hilo] = { id: hilo, sinLeer: 0, ultimo: null };
    out.hilos[hilo].ultimo = {
      de: m.deNombre || m.de, texto: m.borrado ? '(mensaje borrado)' : m.texto,
      cuando: m.cuando, mio: m.de === yo,
    };
    if (!leido(m)) { out.hilos[hilo].sinLeer++; out.sinLeer++; }
  });

  /** La conversación abierta. Por defecto, el grupo de la tienda activa. */
  const con = String(p.con || '').trim() ||
              (out.grupos[0] ? out.grupos[0].id : '');
  if (con && chatPuedeVer_(s, con)) {
    out.con = con;
    out.mensajes = mios.filter(function (m) { return chatHiloDe_(m, yo) === con; })
      .slice(-CHAT_PAGINA)
      .map(function (m) {
        return { id: m.id, de: m.de, deNombre: m.deNombre || m.de,
                 texto: m.borrado ? '' : m.texto, borrado: m.borrado,
                 cuando: m.cuando, mio: m.de === yo };
      });
  }
  return out;
}

/**
 * Escribir. Al grupo de una tienda, o a una persona.
 *
 * `para` vacío significa grupo, y entonces la tienda es obligatoria: un
 * mensaje de grupo sin tienda no lo podría leer nadie, porque el filtro
 * de lectura pregunta justamente por eso.
 */
function apiChatEnviar(s, p) {
  const yo = String(s.email || '').toLowerCase().trim();
  const texto = String(p.texto || '').trim();
  if (!texto) return { ok: false, error: 'El mensaje está vacío.' };
  if (texto.length > CHAT_LARGO) {
    return { ok: false, error: 'El mensaje es muy largo. Máximo ' + CHAT_LARGO +
                               ' caracteres — para lo demás, mejor una nota en el caso.' };
  }

  const con = String(p.con || '').trim();
  if (!chatPuedeVer_(s, con)) {
    return { ok: false, error: 'No puedes escribir en esa conversación.' };
  }

  const ss = libro_(s.sheetId);
  let tienda = '', para = '';

  if (con.indexOf('g:') === 0) {
    tienda = con.slice(2);
  } else {
    para = con.toLowerCase();
    // Solo a alguien que existe y está activo en este cliente.
    const gente = chatGente_(ss, yo);
    if (!gente.filter(function (g) { return g.correo === para; }).length) {
      return { ok: false, error: 'Esa persona no está en tu equipo.' };
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro mensaje guardándose.' };
  try {
    const sh = ss.getSheetByName('Mensajes');
    if (!sh) {
      return { ok: false, error: 'Falta la hoja Mensajes. Corre bootstrapTodo() una vez.' };
    }
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const fila = {
      id: 'm' + Utilities.getUuid().slice(0, 10),
      tienda: tienda, de: yo, de_nombre: s.nombre || yo, para: para,
      texto: texto, creado_en: ahoraISO(),
      // Quien escribe ya lo leyó. Sin esto, todo mensaje propio contaría
      // como pendiente para uno mismo.
      leido_por: yo, borrado: '',
    };
    sh.appendRow(enc.map(function (c) { return fila[c] !== undefined ? fila[c] : ''; }));
    return { ok: true, id: fila.id, cuando: fila.creado_en };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

/**
 * Marcar como leída una conversación.
 *
 * Se escribe la columna entera de una vez. Marcar mensaje por mensaje
 * serían ochenta escrituras cada vez que alguien abre el chat, y el
 * chat se abre veinte veces al día.
 */
function apiChatVisto(s, p) {
  const yo = String(s.email || '').toLowerCase().trim();
  const con = String(p.con || '').trim();
  if (!chatPuedeVer_(s, con)) return { ok: false, error: 'Esa conversación no es tuya.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: true, marcados: 0 };
  try {
    const ss = libro_(s.sheetId);
    const sh = ss.getSheetByName('Mensajes');
    if (!sh || sh.getLastRow() < 2) return { ok: true, marcados: 0 };

    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const cL = e.indexOf('leido_por');
    if (cL === -1) return { ok: true, marcados: 0 };
    const cDe = e.indexOf('de'), cPara = e.indexOf('para'), cT = e.indexOf('tienda');

    let n = 0;
    for (let i = 1; i < d.length; i++) {
      const m = { de: String(d[i][cDe] || '').toLowerCase().trim(),
                  para: String(d[i][cPara] || '').toLowerCase().trim(),
                  tienda: String(d[i][cT] || '').trim() };
      if (chatHiloDe_(m, yo) !== con) continue;
      if (m.de === yo) continue;
      const leidos = String(d[i][cL] || '').toLowerCase();
      if ((',' + leidos + ',').indexOf(',' + yo + ',') !== -1) continue;
      d[i][cL] = leidos ? leidos + ',' + yo : yo;
      n++;
    }
    if (!n) return { ok: true, marcados: 0 };

    const columna = d.slice(1).map(function (f) { return [f[cL]]; });
    sh.getRange(2, cL + 1, columna.length, 1).setValues(columna);
    return { ok: true, marcados: n };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

/**
 * Borrar un mensaje PROPIO.
 *
 * No se borra la fila: se marca. Queda constancia de que hubo un
 * mensaje ahí y de quién lo quitó. En un chat de trabajo, poder hacer
 * desaparecer lo dicho sin rastro es exactamente lo que impide usarlo
 * para saber qué se acordó.
 */
function apiChatBorrar(s, p) {
  const yo = String(s.email || '').toLowerCase().trim();
  const id = String(p.id || '').trim();
  if (!id) return { ok: false, error: 'No sé qué mensaje.' };

  try {
    const ss = libro_(s.sheetId);
    const sh = ss.getSheetByName('Mensajes');
    if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'No hay mensajes.' };
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const cId = e.indexOf('id'), cDe = e.indexOf('de'), cB = e.indexOf('borrado');
    if (cB === -1) return { ok: false, error: 'Falta la columna borrado. Corre bootstrapTodo().' };

    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cId] || '').trim() !== id) continue;
      if (String(d[i][cDe] || '').toLowerCase().trim() !== yo) {
        return { ok: false, error: 'Solo puedes borrar tus propios mensajes.' };
      }
      sh.getRange(i + 1, cB + 1).setValue('si');
      return { ok: true };
    }
    return { ok: false, error: 'No encuentro ese mensaje.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
