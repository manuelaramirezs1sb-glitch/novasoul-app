/**
 * REPARTIR EL TRABAJO · lo que le faltaba a Nova para tener equipo.
 *
 * ┌─ EL AGUJERO ───────────────────────────────────────────────┐
 * │                                                            │
 * │ Nova sabía FILTRAR —una gestora ve solo sus pedidos— y eso  │
 * │ estaba probado. Lo que no existía era la otra mitad: no    │
 * │ había ninguna manera de poner ese nombre en la columna.    │
 * │                                                            │
 * │ Mientras la dueña trabajó sola no se notó: ella lo ve      │
 * │ todo. Se habría notado el primer día que entrara su        │
 * │ hermana, viera CERO pedidos y CERO novedades, y concluyera │
 * │ que la herramienta no sirve.                               │
 * │                                                            │
 * │ Por eso la primera prueba de este archivo no es sobre      │
 * │ asignar: es sobre lo que ve una gestora ANTES y DESPUÉS.   │
 * │ Esa es la que importa.                                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libroStub(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (n) => {
    const m = hojas[n];
    if (!m) return null;
    return {
      getName: () => n, getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (k) => m.splice(k - 1, 1),
      getRange: (f, c, nf, nc) => ({
        getValues: () => {
          const o = [];
          for (let i = 0; i < (nf || 1); i++) {
            const r = m[f - 1 + i] || [];
            o.push(r.slice(c - 1, c - 1 + (nc || r.length)));
          }
          return o;
        },
        /** Escribe DENTRO del rango, nunca filas enteras. */
        setValues: (v) => { v.forEach((r, i) => {
          const d = m[f - 1 + i] || (m[f - 1 + i] = []);
          r.forEach((x, j) => { d[c - 1 + j] = x; }); }); },
        setValue: (x) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = x; },
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-25T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: libroStub, flush: () => {} };
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

(0, eval)(src + '\n;globalThis.__F = { apiReparto, asignablesDe_, apiListar,' +
  ' filtrarPorRol, libroOlvidar_, soulOlvidar_, ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;
const E = F.ESQUEMA_EMPRESARIAL;
const fila = (cols, o) => cols.map(c => (o[c] !== undefined ? o[c] : ''));

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

function sembrar() {
  UUID = 0; F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Pedidos: [E.Pedidos.slice()], Novedades: [E.Novedades.slice()],
    Equipo: [E.Equipo.slice()],
    Tiendas: [['id','nombre','moneda','estado','modalidad'],
              ['ec','Nutrea EC','USD','activa','marca_propia'],
              ['gt','Nutrea GT','GTQ','activa','marca_propia']],
    Movimientos: [['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo']],
    Estados: [['fuente','texto','estado_nova','origen','pedidos','primera_vez',
               'ultima_vez','decidido_por','nota']],
    Parametros: [['tienda','clave','valor','nota','tipo']],
  };
  // Su equipo: ella, su hermana en EC, y alguien solo de GT.
  LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e1', nombre:'Manuela Ramírez',
    correo:'m@nova.com', rol:'dueno', tienda:'*', estado:'activo' }));
  LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e2', nombre:'Andrea Ramírez',
    correo:'andrea@nutrea.com', rol:'gestora', tienda:'ec', estado:'activo' }));
  LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e3', nombre:'Solo de GT',
    correo:'gt@nutrea.com', rol:'gestora', tienda:'gt', estado:'activo' }));
  LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e4', nombre:'Ya no está',
    correo:'x@nutrea.com', rol:'gestora', tienda:'ec', estado:'inactivo' }));

  // Diez pedidos de EC, todos SIN asignar — como los importa Dropi.
  for (let i = 0; i < 10; i++) {
    LIBROS.emp.Pedidos.push(fila(E.Pedidos, { id:'d'+i, tienda:'ec',
      fecha:'2026-09-2'+(i%5), cliente:'Clienta '+i, valor:50,
      estado_canonico: i < 3 ? 'novedad' : (i < 7 ? 'en_transito' : 'entregado'),
      telefono_norm:'59399900'+i }));
  }
  // Y tres novedades colgando de los tres primeros.
  for (let i = 0; i < 3; i++) {
    LIBROS.emp.Novedades.push(fila(E.Novedades, { id:'n'+i, pedido_id:'d'+i,
      fecha:'2026-09-24', motivo:'No contesta', estado:'abierta' }));
  }
  // Un pedido de GT, para comprobar que no se cruza.
  LIBROS.emp.Pedidos.push(fila(E.Pedidos, { id:'gt1', tienda:'gt',
    fecha:'2026-09-24', cliente:'De Guate', valor:80, estado_canonico:'en_transito' }));
}

const DUENA = { sheetId:'emp', rol:'dueno', email:'m@nova.com', nombre:'Manuela Ramírez',
                tiendas:['ec','gt'], permisos:[], modulos:['empresarial'] };
const HERMANA = { sheetId:'emp', rol:'gestora', email:'andrea@nutrea.com',
                  nombre:'Andrea Ramírez', tiendas:['ec'], permisos:[],
                  modulos:['empresarial'] };
const ADMIN = { sheetId:'emp', rol:'admin', email:'a@nova.com', nombre:'Una Admin',
                tiendas:['ec'], permisos:[], modulos:['empresarial'] };

const cuantosVe = (s, entidad) =>
  (F.apiListar(s, { entidad: entidad, tienda: 'ec', limite: 300 }).filas || []).length;

// ══════════════════════════════════════════════════════════════
console.log('\n── SE ASIGNAN TIENDAS, NO PEDIDOS ──');
/**
 * ┌─ POR QUÉ CAMBIÓ TODO ESTE ARCHIVO ─────────────────────────┐
 * │                                                            │
 * │ Empezó probando un reparto pedido por pedido. Funcionaba.  │
 * │ Ella lo corrigió de raíz:                                  │
 * │                                                            │
 * │   «de nada sirve asignar pedidos si ya asignaste la        │
 * │    tienda. La gestora de la tienda gestiona la tienda que  │
 * │    tiene asignada: todo lo que son pedidos, novedades y    │
 * │    CAS. Porque si la tienda tiene pedidos pendientes de    │
 * │    confirmar hace más de 3 días, alguien debe recibir esa  │
 * │    información — y no solo la admin o la dueña».           │
 * │                                                            │
 * │ El filtro viejo producía lo contrario de lo que pretendía: │
 * │ un pedido sin repartir no lo veía NADIE salvo la dueña, y  │
 * │ los que llevan más tiempo quietos son justo los que nadie  │
 * │ tomó. El cuidado terminaba escondiendo el trabajo.         │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
sembrar();
igual('la dueña ve los 10 pedidos', 10, cuantosVe(DUENA, 'Pedidos'));
/** Sin que nadie le asigne un solo pedido, ve TODA su tienda. */
igual('la gestora también ve los 10, sin que nadie le asigne nada',
      10, cuantosVe(HERMANA, 'Pedidos'));
igual('y las novedades de su tienda', 3, cuantosVe(HERMANA, 'Novedades'));

console.log('\n── Pero solo de SU tienda ──');
/** El recorte que de verdad protege: la tienda, no el pedido. */
igual('no ve la tienda de Guatemala', 0,
      (F.apiListar(HERMANA, { entidad: 'Pedidos', tienda: 'gt', limite: 300 }).filas || []).length);
const SOLO_GT = { sheetId:'emp', rol:'gestora', email:'gt@nutrea.com', nombre:'Solo de GT',
                  tiendas:['gt'], permisos:[], modulos:['empresarial'] };
igual('quien solo tiene GT ve 1', 1,
      (F.apiListar(SOLO_GT, { entidad: 'Pedidos', limite: 300 }).filas || []).length);
igual('y no ve ni uno de Ecuador', 0,
      (F.apiListar(SOLO_GT, { entidad: 'Pedidos', tienda: 'ec', limite: 300 }).filas || []).length);

console.log('\n── Dos gestoras pueden compartir tienda ──');
sembrar();
LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e5', nombre:'Otra Gestora',
  correo:'otra@nutrea.com', rol:'gestora', tienda:'ec', estado:'activo' }));
const OTRA = { sheetId:'emp', rol:'gestora', email:'otra@nutrea.com', nombre:'Otra Gestora',
               tiendas:['ec'], permisos:[], modulos:['empresarial'] };
igual('las dos ven lo mismo',
      [cuantosVe(HERMANA, 'Pedidos'), cuantosVe(OTRA, 'Pedidos')], [10, 10]);

console.log('\n── Y una persona puede tener varias tiendas ──');
const DOS = { sheetId:'emp', rol:'gestora', email:'andrea@nutrea.com',
              nombre:'Andrea Ramírez', tiendas:['ec','gt'], permisos:[],
              modulos:['empresarial'] };
igual('ve las dos juntas cuando no pide una', 11,
      (F.apiListar(DOS, { entidad: 'Pedidos', limite: 300 }).filas || []).length);
igual('y una sola cuando la pide', 1,
      (F.apiListar(DOS, { entidad: 'Pedidos', tienda: 'gt', limite: 300 }).filas || []).length);

console.log('\n── Puede trabajar cualquier caso de su tienda ──');
sembrar();
igual('escribe sobre uno que nadie le asignó', true,
      globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'd3',
                                        campos: { nota: 'llamé yo' } }).ok);
/** Pero nunca fuera de su tienda. */
igual('y nunca sobre uno de otra tienda', false,
      globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'gt1',
                                        campos: { nota: 'no es mía' } }).ok);

console.log('\n── Quién lo tocó queda anotado solo ──');
sembrar();
const cHizo = E.Pedidos.indexOf('gestionado_por');
igual('antes, nadie', '', LIBROS.emp.Pedidos.filter(f => f[0] === 'd0')[0][cHizo]);
globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'd0',
                                  campos: { estado_nova: 'confirmado' } });
igual('después, quien lo movió', 'Andrea Ramírez',
      LIBROS.emp.Pedidos.filter(f => f[0] === 'd0')[0][cHizo]);
/**
 * Y la siguiente persona NO lo pisa: el crédito del trabajo es de quien
 * lo resolvió, no de quien lo miró de último.
 */
globalThis.apiEscribir(DUENA, { entidad: 'Pedidos', id: 'd0',
                                campos: { nota: 'revisado' } });
igual('y quien pase después no lo pisa', 'Andrea Ramírez',
      LIBROS.emp.Pedidos.filter(f => f[0] === 'd0')[0][cHizo]);

/** Lo que trajo el archivo histórico también se respeta. */
sembrar();
LIBROS.emp.Pedidos.filter(f => f[0] === 'd1')[0][cHizo] = 'JAIME';
globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'd1', campos: { nota: 'x' } });
igual('ni al que venía del archivo', 'JAIME',
      LIBROS.emp.Pedidos.filter(f => f[0] === 'd1')[0][cHizo]);

console.log('\n── Quién trabaja la tienda ──');
sembrar();
LIBROS.emp.Pedidos.filter(f => f[0] === 'd7')[0][cHizo] = 'JAIME';
LIBROS.emp.Pedidos.filter(f => f[0] === 'd8')[0][cHizo] = 'JAIME';
LIBROS.emp.Pedidos.filter(f => f[0] === 'd9')[0][cHizo] = 'Andrea Ramírez';
let rep = F.apiReparto(DUENA, { tienda: 'ec' });
igual('10 pedidos', 10, rep.total);
igual('7 sin tocar', 7, rep.sinTocar);
/** JAIME no tiene cuenta y cuenta igual: es quien hizo el trabajo. */
igual('JAIME lleva 2', 2, rep.etiquetas.filter(e => e.nombre === 'JAIME')[0].total);
igual('y se dice que no tiene cuenta', false,
      rep.etiquetas.filter(e => e.nombre === 'JAIME')[0].esUsuario);
igual('Andrea sí la tiene', true,
      rep.etiquetas.filter(e => e.nombre === 'Andrea Ramírez')[0].esUsuario);
igual('una gestora no ve el reparto', false, F.apiReparto(HERMANA, { tienda: 'ec' }).ok);

console.log('\n── A quién se le puede asignar una tienda ──');
sembrar();
const gente = F.asignablesDe_(libroStub('emp'), 'ec');
igual('solo activos y de esa tienda', ['Manuela Ramírez', 'Andrea Ramírez'],
      gente.map(g => g.nombre));
ok('la inactiva no está', gente.filter(g => g.nombre === 'Ya no está').length === 0);
ok('la de GT tampoco', gente.filter(g => g.nombre === 'Solo de GT').length === 0);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
