/**
 * EL CHAT DEL EQUIPO · que sea un chat, y que respete el acceso.
 *
 * ┌─ LO QUE ESTA PRUEBA EXISTE PARA QUE NO VUELVA A PASAR ─────┐
 * │                                                            │
 * │ «el dueño, la admin y las gestoras deben tener un mismo    │
 * │  chat en Nova (…) dentro de Nova, no por fuera».           │
 * │                                                            │
 * │ La pestaña «Nova Chat» era un BOTÓN que abría un grupo de   │
 * │ WhatsApp. No fallaba: simplemente no era un chat.           │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Lo que se mide aquí, en este orden:
 *
 *  1. Que un mensaje escrito se pueda leer. Lo primero, porque es lo
 *     único que la pantalla anterior no hacía.
 *  2. Que el grupo de una tienda NO lo lea quien no tiene esa tienda.
 *     Y no solo que la lista no lo ofrezca: que pedirlo a mano tampoco
 *     lo entregue, porque cambiar un valor en la consola del navegador
 *     es exactamente lo que hace cualquiera que quiera mirar.
 *  3. Que un privado sea privado, incluso para la dueña.
 *  4. Que lo borrado quede marcado y no desaparecido.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
/** Cuántas escrituras de rango hizo el código. Marcar 80 mensajes uno por
 *  uno funciona igual de bien y tarda veinte veces más. */
let ESCRITURAS = 0;
function libro(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (nombre) => {
    const m = hojas[nombre];
    if (!m) return null;
    return {
      getName: () => nombre,
      getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (n) => m.splice(n - 1, 1),
      getRange: (f, c, nf, nc) => ({
        getValues: () => {
          const out = [];
          for (let i = 0; i < (nf || 1); i++) {
            const fila = m[f - 1 + i] || [];
            out.push(fila.slice(c - 1, c - 1 + (nc || fila.length)));
          }
          return out;
        },
        setValues: (v) => { ESCRITURAS++; v.forEach((fila, i) => {
          const d = m[f - 1 + i] || (m[f - 1 + i] = []);
          fila.forEach((val, j) => { d[c - 1 + j] = val; }); }); },
        setValue: (val) => { ESCRITURAS++;
          const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = val; },
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: (id) => libro(id), flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{}' }) };
let UUID = 0;
global.Utilities = { sleep: () => {}, getUuid: () => 'u' + (++UUID) + '0000000',
  formatDate: (d, tz, pat) => {
    const iso = (d && d.getTime && d.getTime() === new Date(HOY).getTime())
      ? HOY : new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(pat)) return iso.slice(0, 19).replace('T', ' ');
    return iso;
  } };
const RealDate = Date;
global.Date = class extends RealDate {
  constructor(...a) { return a.length ? new RealDate(...a) : new RealDate(HOY); }
  static now() { return new RealDate(HOY).getTime(); }
  static UTC(...a) { return RealDate.UTC(...a); }
};

(0, eval)(src + '\n;globalThis.__F = { apiChat, apiChatEnviar, apiChatVisto, apiChatBorrar,' +
  ' chatPuedeVer_, chatHiloDe_, chatGente_, libroOlvidar_, soulOlvidar_,' +
  ' ESQUEMA_EMPRESARIAL, CHAT_LARGO };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_MSG = F.ESQUEMA_EMPRESARIAL.Mensajes;
const C_EQU = F.ESQUEMA_EMPRESARIAL.Equipo;
const C_TIE = F.ESQUEMA_EMPRESARIAL.Tiendas;

function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

/**
 * El equipo de Sara: ella, una admin que cubre las dos tiendas, y dos
 * gestoras con UNA tienda cada una. Esa es la forma que importa: si las
 * dos gestoras tuvieran las dos tiendas, la prueba del acceso no probaría
 * nada.
 */
function sembrar() {
  UUID = 0; ESCRITURAS = 0;
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Tiendas: [C_TIE, f(C_TIE, { id: 'ec', nombre: 'Nutrea Ecuador' }),
                     f(C_TIE, { id: 'gt', nombre: 'Nutrea Guatemala' })],
    Equipo: [C_EQU,
      f(C_EQU, { id: 'e1', nombre: 'Sara',    correo: 'sara@x.com',  rol: 'dueno',   tienda: '*',  estado: 'activo' }),
      f(C_EQU, { id: 'e2', nombre: 'Katherin',correo: 'kat@x.com',   rol: 'admin',   tienda: 'ec,gt', estado: 'activo' }),
      f(C_EQU, { id: 'e3', nombre: 'Zulay',   correo: 'zulay@x.com', rol: 'gestora', tienda: 'ec', estado: 'activo' }),
      f(C_EQU, { id: 'e4', nombre: 'Jaime',   correo: 'jaime@x.com', rol: 'gestora', tienda: 'gt', estado: 'activo' }),
      f(C_EQU, { id: 'e5', nombre: 'Ana',     correo: 'ana@x.com',   rol: 'gestora', tienda: 'ec', estado: 'inactivo' }),
    ],
    Mensajes: [C_MSG],
  };
}

const S = {
  duena:   { email: 'sara@x.com',  nombre: 'Sara',     rol: 'dueno',   sheetId: 'emp', tiendas: ['ec', 'gt'] },
  admin:   { email: 'kat@x.com',   nombre: 'Katherin', rol: 'admin',   sheetId: 'emp', tiendas: ['ec', 'gt'] },
  zulay:   { email: 'zulay@x.com', nombre: 'Zulay',    rol: 'gestora', sheetId: 'emp', tiendas: ['ec'] },
  jaime:   { email: 'jaime@x.com', nombre: 'Jaime',    rol: 'gestora', sheetId: 'emp', tiendas: ['gt'] },
};

console.log('\n── 1 · escribir y leer, que era lo que no pasaba ──');
sembrar();

let r = F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'Buenos días equipo, hoy priorizamos Ecuador.' });
ok('la dueña escribe al grupo de su tienda', r.ok, JSON.stringify(r));
ok('el mensaje quedó en la hoja', LIBROS.emp.Mensajes.length === 2,
   'filas: ' + LIBROS.emp.Mensajes.length);

let v = F.apiChat(S.zulay, {});
igual('la gestora de ec lo lee', 1, v.mensajes.length);
igual('y lee el texto completo', 'Buenos días equipo, hoy priorizamos Ecuador.',
      v.mensajes[0] && v.mensajes[0].texto);
igual('con el nombre de quien lo escribió', 'Sara', v.mensajes[0] && v.mensajes[0].deNombre);
ok('no marcado como propio', v.mensajes[0] && v.mensajes[0].mio === false);
igual('le cuenta 1 sin leer', 1, v.sinLeer);

// Quien escribe no se cuenta a sí misma: sin esto, cada mensaje propio
// dejaría un punto rojo que no se puede quitar.
v = F.apiChat(S.duena, {});
igual('a la dueña su propio mensaje no le cuenta sin leer', 0, v.sinLeer);
ok('y lo ve como propio', v.mensajes[0] && v.mensajes[0].mio === true);

console.log('\n── 2 · el grupo de una tienda no lo lee quien no la tiene ──');
v = F.apiChat(S.jaime, {});
igual('la gestora de gt no ve nada de ec', 0, v.mensajes.length);
igual('ni le cuenta sin leer', 0, v.sinLeer);
igual('su conversación por defecto es su propia tienda', 'g:gt', v.con);

/**
 * La lista no ofrece esa conversación. Eso no impide pedirla: basta
 * escribir `con: 'g:ec'` en la consola del navegador. Por eso el filtro
 * está en el servidor y no en la pantalla.
 */
v = F.apiChat(S.jaime, { con: 'g:ec' });
igual('pedir a mano el grupo ajeno no devuelve mensajes', 0, v.mensajes.length);
igual('y no queda abierto', '', v.con);

r = F.apiChatEnviar(S.jaime, { con: 'g:ec', texto: 'hola' });
ok('tampoco puede escribir en el grupo ajeno', r.ok === false, JSON.stringify(r));
igual('y no se guardó nada', 2, LIBROS.emp.Mensajes.length);

igual('la dueña sí tiene los dos grupos', ['g:ec', 'g:gt'],
      F.apiChat(S.duena, {}).grupos.map(g => g.id));
igual('con el nombre de la tienda, no el código', 'Nutrea Ecuador',
      F.apiChat(S.duena, {}).grupos[0].nombre);
igual('la gestora de una sola tienda tiene un grupo', 1,
      F.apiChat(S.zulay, {}).grupos.length);

console.log('\n── 3 · el privado es privado, también para la dueña ──');
r = F.apiChatEnviar(S.admin, { con: 'zulay@x.com', texto: 'Zulay, te faltó confirmar tres de ayer.' });
ok('la admin le escribe en privado a la gestora', r.ok, JSON.stringify(r));

v = F.apiChat(S.zulay, { con: 'kat@x.com' });
igual('la gestora lo recibe', 'Zulay, te faltó confirmar tres de ayer.',
      v.mensajes[0] && v.mensajes[0].texto);
igual('el hilo se llama por la otra persona', 'kat@x.com', v.con);

v = F.apiChat(S.duena, { con: 'kat@x.com' });
igual('la dueña no lee el privado entre otras dos', 0, v.mensajes.length);
v = F.apiChat(S.duena, {});
ok('ni le aparece en la lista de hilos', !v.hilos['kat@x.com'] && !v.hilos['zulay@x.com'],
   JSON.stringify(Object.keys(v.hilos)));
igual('ni le cuenta sin leer', 0, v.sinLeer);

// La admin cubre las dos tiendas; escribirle a la gestora de Guatemala
// no exige estar parada en Guatemala. El privado va entre personas.
r = F.apiChatEnviar(S.admin, { con: 'jaime@x.com', texto: 'Jaime, mira las novedades de gt.' });
ok('el privado cruza tiendas', r.ok, JSON.stringify(r));
igual('y llega', 1, F.apiChat(S.jaime, { con: 'kat@x.com' }).mensajes.length);

r = F.apiChatEnviar(S.admin, { con: 'nadie@fuera.com', texto: 'hola' });
ok('no se le escribe a alguien de fuera del equipo', r.ok === false, JSON.stringify(r));
r = F.apiChatEnviar(S.admin, { con: 'ana@x.com', texto: 'hola' });
ok('ni a quien está inactiva', r.ok === false, JSON.stringify(r));

console.log('\n── 4 · sin leer, y marcar de una sola escritura ──');
sembrar();
F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'uno' });
F.apiChatEnviar(S.admin, { con: 'g:ec', texto: 'dos' });
F.apiChatEnviar(S.admin, { con: 'zulay@x.com', texto: 'privado' });
F.apiChatEnviar(S.duena, { con: 'g:gt', texto: 'otra tienda' });

v = F.apiChat(S.zulay, {});
igual('la gestora tiene 3 sin leer (2 del grupo + 1 privado)', 3, v.sinLeer);
igual('dos en el grupo de su tienda', 2, v.hilos['g:ec'].sinLeer);
igual('uno en el privado con la admin', 1, v.hilos['kat@x.com'].sinLeer);
igual('lo último del grupo es el último mensaje', 'dos', v.hilos['g:ec'].ultimo.texto);
ok('el hilo de la otra tienda no existe para ella', !v.hilos['g:gt'],
   JSON.stringify(Object.keys(v.hilos)));

ESCRITURAS = 0;
r = F.apiChatVisto(S.zulay, { con: 'g:ec' });
igual('marca los dos del grupo', 2, r.marcados);
igual('en UNA sola escritura', 1, ESCRITURAS);

v = F.apiChat(S.zulay, {});
igual('ya no tiene sin leer en el grupo', 0, v.hilos['g:ec'].sinLeer);
igual('el privado sigue sin leer', 1, v.sinLeer);

// Volver a marcar lo ya leído no debe escribir nada: el chat se abre
// veinte veces al día y cada apertura marcaría la hoja entera.
ESCRITURAS = 0;
r = F.apiChatVisto(S.zulay, { con: 'g:ec' });
igual('volver a marcar no marca nada', 0, r.marcados);
igual('y no escribe', 0, ESCRITURAS);

// La otra gestora no se ve afectada por lo que leyó Zulay.
igual('lo que leyó una no lo marca para la otra', 1, F.apiChat(S.jaime, {}).sinLeer);

r = F.apiChatVisto(S.jaime, { con: 'g:ec' });
ok('no se marca como leída una conversación ajena', r.ok === false, JSON.stringify(r));

console.log('\n── 4b · abrir y marcar, en un solo viaje ──');
/**
 * «cuenta todos los llamados que necesita Nova Empresarial para
 *  funcionar y no ser lenta».
 *
 * Abrir una conversación eran DOS peticiones —traer y marcar— para un
 * solo gesto, y la segunda releía la hoja entera que la primera acababa
 * de leer. Con `marcar` es una.
 */
sembrar();
F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'uno' });
F.apiChatEnviar(S.admin, { con: 'g:ec', texto: 'dos' });
v = F.apiChat(S.zulay, { con: 'g:ec', marcar: true });
igual('trae los mensajes', 2, v.mensajes.length);
igual('y de paso los marca', 2, v.marcados);
igual('y ya lo dice en el mismo vuelo', 0, v.hilos['g:ec'].sinLeer);
igual('y en el total', 0, v.sinLeer);
igual('al volver a preguntar sigue en cero', 0, F.apiChat(S.zulay, {}).sinLeer);

// Sin `marcar` no se marca nada: leer la lista no es abrir el hilo.
sembrar();
F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'uno' });
v = F.apiChat(S.zulay, { con: 'g:ec' });
igual('sin pedirlo, no marca', undefined, v.marcados);
igual('y sigue contando sin leer', 1, F.apiChat(S.zulay, {}).sinLeer);

// Y no se puede marcar una conversación ajena pidiéndola a mano.
r = F.apiChat(S.jaime, { con: 'g:ec', marcar: true });
igual('pedir a mano el grupo ajeno no lo abre', '', r.con);
igual('ni lo marca', 1, F.apiChat(S.zulay, {}).sinLeer);

console.log('\n── 5 · borrar deja marca, no hueco ──');
sembrar();
const m1 = F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'me equivoqué de tienda' });
F.apiChatEnviar(S.admin, { con: 'g:ec', texto: 'tranquila' });

r = F.apiChatBorrar(S.admin, { id: m1.id });
ok('nadie borra el mensaje de otra persona', r.ok === false, JSON.stringify(r));

r = F.apiChatBorrar(S.duena, { id: m1.id });
ok('la dueña borra el suyo', r.ok, JSON.stringify(r));
igual('la fila NO se borró', 3, LIBROS.emp.Mensajes.length);

v = F.apiChat(S.zulay, {});
igual('siguen siendo dos mensajes', 2, v.mensajes.length);
ok('el borrado va marcado', v.mensajes[0].borrado === true, JSON.stringify(v.mensajes[0]));
igual('y sin texto', '', v.mensajes[0].texto);
ok('el otro sigue entero', v.mensajes[1].texto === 'tranquila');

r = F.apiChatBorrar(S.duena, { id: 'noexiste' });
ok('borrar algo que no existe avisa', r.ok === false, JSON.stringify(r));

console.log('\n── 6 · los bordes ──');
sembrar();
r = F.apiChatEnviar(S.duena, { con: 'g:ec', texto: '   ' });
ok('un mensaje en blanco no se guarda', r.ok === false, JSON.stringify(r));
r = F.apiChatEnviar(S.duena, { con: 'g:ec', texto: 'x'.repeat(F.CHAT_LARGO + 1) });
ok('un pegote de mil líneas tampoco', r.ok === false, JSON.stringify(r));
r = F.apiChatEnviar(S.duena, { con: '', texto: 'hola' });
ok('sin conversación no se escribe', r.ok === false, JSON.stringify(r));
r = F.apiChatEnviar(S.duena, { con: 'g:', texto: 'hola' });
ok('un grupo sin tienda no existe', r.ok === false, JSON.stringify(r));
igual('nada de eso se guardó', 1, LIBROS.emp.Mensajes.length);

// Con la hoja vacía la pantalla tiene que abrir igual, no explotar.
v = F.apiChat(S.zulay, {});
ok('con la hoja vacía la pantalla abre', v.ok === true, JSON.stringify(v));
igual('y ofrece su grupo', 'g:ec', v.con);
igual('con el equipo para los privados', ['sara@x.com', 'kat@x.com', 'jaime@x.com'],
      v.gente.map(g => g.correo));
ok('sin ofrecerle hablar consigo misma',
   v.gente.every(g => g.correo !== 'zulay@x.com'));
ok('ni con quien está inactiva', v.gente.every(g => g.correo !== 'ana@x.com'));

// Y sin la hoja siquiera: es lo que pasa hasta que corra bootstrapTodo().
delete LIBROS.emp.Mensajes;
F.libroOlvidar_();
v = F.apiChat(S.zulay, {});
ok('sin la hoja Mensajes la pantalla abre igual', v.ok === true, JSON.stringify(v));
r = F.apiChatEnviar(S.zulay, { con: 'g:ec', texto: 'hola' });
ok('y escribir avisa qué falta', r.ok === false && /bootstrapTodo/.test(r.error || ''),
   JSON.stringify(r));

console.log('\n── 7 · la hoja está declarada ──');
ok('Mensajes está en el esquema', !!C_MSG, JSON.stringify(C_MSG));
['id', 'tienda', 'de', 'de_nombre', 'para', 'texto', 'creado_en', 'leido_por', 'borrado']
  .forEach(c => ok('columna ' + c, (C_MSG || []).indexOf(c) !== -1));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
