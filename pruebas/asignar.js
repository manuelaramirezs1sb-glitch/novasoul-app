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

(0, eval)(src + '\n;globalThis.__F = { apiAsignar, apiReparto, asignablesDe_, apiListar,' +
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
console.log('\n── EL PRIMER DÍA DE SU HERMANA, ANTES DE ESTO ──');
sembrar();
igual('la dueña ve los 10 pedidos', 10, cuantosVe(DUENA, 'Pedidos'));
/**
 * Esto es lo que habría pasado: entra, y no hay nada. Ni un pedido, ni
 * una novedad. No porque falle algo, sino porque nadie podía asignarle.
 */
igual('la hermana ve CERO pedidos', 0, cuantosVe(HERMANA, 'Pedidos'));
igual('y CERO novedades', 0, cuantosVe(HERMANA, 'Novedades'));

console.log('\n── Y AHORA, DESPUÉS DE ASIGNAR ──');
let r = F.apiAsignar(DUENA, { ids: ['d0','d1','d2','d3'], a: 'e2', tienda: 'ec' });
igual('se asignan los cuatro', 4, r.asignados);
igual('con el nombre de la hoja Equipo', 'Andrea Ramírez', r.a);
igual('la hermana ya ve sus cuatro', 4, cuantosVe(HERMANA, 'Pedidos'));
/** Y —lo que se olvida siempre— las novedades se van con el pedido. */
igual('las novedades se fueron con ellos', 3, r.novedades);
igual('y las ve', 3, cuantosVe(HERMANA, 'Novedades'));
igual('la dueña sigue viéndolo todo', 10, cuantosVe(DUENA, 'Pedidos'));

console.log('\n── El nombre se canoniza contra la hoja, siempre ──');
sembrar();
/**
 * Se asigna escribiendo «andrea ramirez» en minúsculas y sin tilde. Lo
 * que se guarda tiene que ser «Andrea Ramírez», el de la hoja: si se
 * guardara lo escrito, un día entraría «Andrea» y otro «Andrea R.» y el
 * filtro dejaría de encontrar la mitad de sus pedidos.
 */
F.apiAsignar(DUENA, { ids: ['d0'], a: 'andrea ramirez', tienda: 'ec' });
const cG = E.Pedidos.indexOf('gestora_asignada');
igual('se guardó el nombre canónico', 'Andrea Ramírez',
      LIBROS.emp.Pedidos.filter(f => f[0] === 'd0')[0][cG]);
igual('y por correo también funciona', true,
      F.apiAsignar(DUENA, { ids: ['d1'], a: 'andrea@nutrea.com', tienda: 'ec' }).ok);
igual('guardando otra vez el nombre', 'Andrea Ramírez',
      LIBROS.emp.Pedidos.filter(f => f[0] === 'd1')[0][cG]);

console.log('\n── Si la plataforma exportó el correo, igual lo ve ──');
sembrar();
LIBROS.emp.Pedidos.filter(f => f[0] === 'd5')[0][cG] = 'andrea@nutrea.com';
igual('el filtro acepta el correo, no solo el nombre', 1, cuantosVe(HERMANA, 'Pedidos'));
/** Pero NO adivina por parecido: enseñar de más es peor que de menos. */
sembrar();
LIBROS.emp.Pedidos.filter(f => f[0] === 'd5')[0][cG] = 'Andrea';
igual('un nombre a medias NO cuenta como suyo', 0, cuantosVe(HERMANA, 'Pedidos'));

console.log('\n── Desasignar se puede ──');
sembrar();
F.apiAsignar(DUENA, { ids: ['d0','d1'], a: 'e2', tienda: 'ec' });
igual('tiene dos', 2, cuantosVe(HERMANA, 'Pedidos'));
r = F.apiAsignar(DUENA, { ids: ['d0'], a: '', tienda: 'ec' });
igual('se quita uno', 1, r.asignados);
igual('avisa que fue una desasignación', true, r.desasignado);
igual('y le queda uno', 1, cuantosVe(HERMANA, 'Pedidos'));

console.log('\n── Las tiendas no se cruzan ──');
sembrar();
r = F.apiAsignar(DUENA, { ids: ['gt1'], a: 'e2', tienda: 'gt' });
igual('no se le puede asignar GT a quien solo tiene EC', false, r.ok);
ok('y se explica', /no tiene acceso a esta tienda/.test(r.error), r.error);
igual('a quien sí tiene GT, sí', 1,
      F.apiAsignar(DUENA, { ids: ['gt1'], a: 'e3', tienda: 'gt' }).asignados);

console.log('\n── Quién puede repartir ──');
sembrar();
igual('una gestora NO reparte', false,
      F.apiAsignar(HERMANA, { ids: ['d0'], a: 'e2' }).ok);
igual('ni ve el reparto', false, F.apiReparto(HERMANA, { tienda: 'ec' }).ok);
igual('una admin sí', true, F.apiAsignar(ADMIN, { ids: ['d0'], a: 'e2', tienda: 'ec' }).ok);

console.log('\n── A quién se le puede asignar ──');
sembrar();
const gente = F.asignablesDe_(libroStub('emp'), 'ec');
igual('solo activos y de esa tienda', ['Manuela Ramírez', 'Andrea Ramírez'],
      gente.map(g => g.nombre));
ok('la inactiva no está', gente.filter(g => g.nombre === 'Ya no está').length === 0);
ok('la de GT tampoco', gente.filter(g => g.nombre === 'Solo de GT').length === 0);
igual('a alguien que no existe, no', false,
      F.apiAsignar(DUENA, { ids: ['d0'], a: 'Fulana', tienda: 'ec' }).ok);
igual('ni a alguien inactivo', false,
      F.apiAsignar(DUENA, { ids: ['d0'], a: 'Ya no está', tienda: 'ec' }).ok);

console.log('\n── El reparto: el número que de verdad importa ──');
sembrar();
let rep = F.apiReparto(DUENA, { tienda: 'ec' });
igual('10 pedidos', 10, rep.total);
/** Mientras esto sea el total, tener equipo no sirve de nada. */
igual('los 10 sin asignar', 10, rep.sinAsignar);
F.apiAsignar(DUENA, { ids: ['d0','d1','d2','d3','d4'], a: 'e2', tienda: 'ec' });
rep = F.apiReparto(DUENA, { tienda: 'ec' });
igual('ahora 5 sin asignar', 5, rep.sinAsignar);
const andrea = rep.gente.filter(g => g.nombre === 'Andrea Ramírez')[0];
igual('Andrea tiene 5', 5, andrea.total);
/** d0,d1,d2 son novedad y d3,d4 en tránsito: los 5 están abiertos. */
igual('y los 5 están abiertos', 5, andrea.abiertos);

console.log('\n── Un nombre que no es de nadie se DENUNCIA ──');
/**
 * Pasa cuando la plataforma exporta a alguien que ya no está, o escrito
 * distinto. Son pedidos que NADIE ve —su supuesta dueña no puede
 * entrar— y desaparecerían en silencio si no se contaran.
 */
sembrar();
LIBROS.emp.Pedidos.filter(f => f[0] === 'd7')[0][cG] = 'Karen (la de antes)';
LIBROS.emp.Pedidos.filter(f => f[0] === 'd8')[0][cG] = 'Karen (la de antes)';
rep = F.apiReparto(DUENA, { tienda: 'ec' });
igual('los cuenta aparte', [{ nombre: 'Karen (la de antes)', pedidos: 2 }],
      rep.aNadieConocido);
ok('y no los cuenta como sin asignar', rep.sinAsignar === 8);

console.log('\n── No repite trabajo ni escribe de más ──');
sembrar();
F.apiAsignar(DUENA, { ids: ['d0'], a: 'e2', tienda: 'ec' });
const movs = LIBROS.emp.Movimientos.length;
r = F.apiAsignar(DUENA, { ids: ['d0'], a: 'e2', tienda: 'ec' });
igual('asignar lo mismo otra vez no hace nada', 0, r.asignados);
igual('lo dice', 1, r.iguales);
igual('y no ensucia el rastro', movs, LIBROS.emp.Movimientos.length);

console.log('\n── El rastro queda ──');
sembrar();
F.apiAsignar(DUENA, { ids: ['d0'], a: 'e2', tienda: 'ec' });
const m = LIBROS.emp.Movimientos.filter(x => x[4] === 'gestora_asignada');
igual('un movimiento por pedido', 1, m.length);
igual('con el antes y el después', ['', 'Andrea Ramírez'], [m[0][5], m[0][6]]);

console.log('\n── Muchos de una sola vez ──');
sembrar();
r = F.apiAsignar(DUENA, { ids: ['d0','d1','d2','d3','d4','d5','d6','d7','d8','d9'],
                          a: 'e2', tienda: 'ec' });
igual('los diez', 10, r.asignados);
igual('y los ve todos', 10, cuantosVe(HERMANA, 'Pedidos'));
igual('más de 500 se rechaza', false,
      F.apiAsignar(DUENA, { ids: new Array(501).fill('x'), a: 'e2' }).ok);

console.log('\n── Una gestora sigue sin poder escribir lo ajeno ──');
sembrar();
F.apiAsignar(DUENA, { ids: ['d0'], a: 'e2', tienda: 'ec' });
/** d1 NO es suyo: no puede tocarlo aunque sepa el id. */
igual('no puede editar un pedido que no es suyo', false,
      globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'd1',
                                        campos: { nota: 'mío' } }).ok);
igual('el suyo sí', true,
      globalThis.apiEscribir(HERMANA, { entidad: 'Pedidos', id: 'd0',
                                        campos: { nota: 'llamé' } }).ok);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
