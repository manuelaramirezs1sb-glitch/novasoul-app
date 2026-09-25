/**
 * LA BITÁCORA DE INTENTOS · que una nota no borre a la otra.
 *
 * ┌─ LO QUE ESTA PRUEBA EXISTE PARA QUE NO VUELVA A PASAR ─────┐
 * │                                                            │
 * │ «hay novedad con nota pero dice guardar y esa nota la puse  │
 * │  como el miércoles. Recuerda que deben dejar la nota por    │
 * │  cada intento de contacto, deben aparecer todas, con fecha  │
 * │  y nombre de quien puso la nota».                          │
 * │                                                            │
 * │ Había UNA columna `nota` y una caja de texto que se abría   │
 * │ con la nota anterior dentro. Quien anotaba el jueves        │
 * │ escribía encima del miércoles. Sin error, sin aviso: el     │
 * │ intento del miércoles dejaba de existir.                   │
 * │                                                            │
 * │ La primera aserción de este archivo es esa y no otra: que   │
 * │ la segunda nota NO borre la primera.                       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
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
        setValues: (v) => { v.forEach((fila, i) => {
          const d = m[f - 1 + i] || (m[f - 1 + i] = []);
          fila.forEach((val, j) => { d[c - 1 + j] = val; }); }); },
        setValue: (val) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = val; },
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-25T12:00:00Z';
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


(0, eval)(src + '\n;globalThis.__F = { apiNotas, apiNotaAgregar, apiEscribir,' +
  ' libroOlvidar_, soulOlvidar_, ESQUEMA_EMPRESARIAL, NOTA_LARGO };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_NOT = F.ESQUEMA_EMPRESARIAL.Notas;
const C_PED = F.ESQUEMA_EMPRESARIAL.Pedidos;
const C_NOV = F.ESQUEMA_EMPRESARIAL.Novedades;
const C_TIE = F.ESQUEMA_EMPRESARIAL.Tiendas;
const C_MOV = F.ESQUEMA_EMPRESARIAL.Movimientos;
function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

function sembrar() {
  UUID = 0;
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Tiendas: [C_TIE, f(C_TIE, { id: 'ec', nombre: 'Nutrea Ecuador' }),
                     f(C_TIE, { id: 'gt', nombre: 'Nutrea Guatemala' })],
    Pedidos: [C_PED,
      f(C_PED, { id: 'p1', tienda: 'ec', cliente: 'Alicia Venus Pantoja',
                 id_externo: '6963795', estado_nova: 'novedad' }),
      f(C_PED, { id: 'p2', tienda: 'ec', cliente: 'Ronald Meneses',
                 nota: 'la escribió alguien antes de que existiera la bitácora' }),
      f(C_PED, { id: 'pg', tienda: 'gt', cliente: 'De Guatemala' }),
    ],
    Novedades: [C_NOV, f(C_NOV, { id: 'n1', tienda: 'ec', pedido_id: 'p1',
                                  motivo: 'No contesta' })],
    Notas: [C_NOT],
    Movimientos: [C_MOV],
  };
}

const S = {
  duena: { email: 'm@nova.com', nombre: 'Manuela Ramírez', rol: 'dueno',
           sheetId: 'emp', tiendas: ['ec', 'gt'] },
  zulay: { email: 'zulay@x.com', nombre: 'Zulay', rol: 'gestora',
           sheetId: 'emp', tiendas: ['ec'] },
  jaime: { email: 'jaime@x.com', nombre: 'Jaime', rol: 'gestora',
           sheetId: 'emp', tiendas: ['gt'] },
};

console.log('\n── 1 · LA SEGUNDA NOTA NO BORRA LA PRIMERA ──');
sembrar();
let r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'p1',
  texto: 'se llamó por wpp, se dejó mensaje con doña Jasinta' });
ok('se anota el primer intento', r.ok, JSON.stringify(r));

r = F.apiNotaAgregar(S.duena, { entidad: 'Pedidos', id: 'p1',
  texto: 'segundo mensaje por chat, no contesta. Dice que Alicia está de viaje' });
ok('y el segundo también', r.ok, JSON.stringify(r));

let v = F.apiNotas(S.zulay, { tienda: 'ec' });
const hilo = v.notas['Pedidos:p1'] || [];
igual('quedan LAS DOS', 2, hilo.length);
igual('en orden, la primera primero',
      'se llamó por wpp, se dejó mensaje con doña Jasinta', hilo[0].texto);
igual('y la segunda después',
      'segundo mensaje por chat, no contesta. Dice que Alicia está de viaje',
      hilo[1].texto);

console.log('\n── 2 · con fecha y nombre de quien la puso ──');
igual('la primera la firmó la gestora', 'Zulay', hilo[0].autor);
igual('la segunda, la dueña', 'Manuela Ramírez', hilo[1].autor);
ok('las dos con fecha y hora', /^2026-09-25 \d\d:\d\d/.test(hilo[0].cuando) &&
   /^2026-09-25 \d\d:\d\d/.test(hilo[1].cuando),
   hilo[0].cuando + ' / ' + hilo[1].cuando);

console.log('\n── 3 · la columna `nota` sigue con la última ──');
/**
 * No es duplicar por comodidad: la auditoría levanta «sin_nota» leyendo
 * esa columna y el export a Excel la lleva. Si dejara de escribirse,
 * todo eso diría que nadie anotó nada justo cuando el equipo por fin
 * anota.
 */
const p1 = LIBROS.emp.Pedidos[1];
igual('la columna tiene la nota más reciente',
      'segundo mensaje por chat, no contesta. Dice que Alicia está de viaje',
      p1[C_PED.indexOf('nota')]);
igual('y quien anotó queda como quien lo trabajó', 'Zulay',
      p1[C_PED.indexOf('gestionado_por')]);

console.log('\n── 4 · lo que ya estaba escrito no se pierde ──');
// p2 trae una nota vieja en su columna y ninguna entrada en la bitácora.
v = F.apiNotas(S.zulay, { tienda: 'ec' });
ok('el pedido viejo no tiene entradas todavía', !v.notas['Pedidos:p2']);
igual('pero su nota sigue en la fila',
      'la escribió alguien antes de que existiera la bitácora',
      LIBROS.emp.Pedidos[2][C_PED.indexOf('nota')]);
// Es la pantalla la que la muestra como primera entrada; aquí lo que
// importa es que agregar una nota nueva NO la destruya.
r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'p2', texto: 'la llamé hoy' });
ok('se puede anotar encima', r.ok, JSON.stringify(r));
igual('y la bitácora arranca con la nueva', 1,
      (F.apiNotas(S.zulay, { tienda: 'ec' }).notas['Pedidos:p2'] || []).length);

console.log('\n── 5 · no se puede editar ni borrar lo anotado ──');
/**
 * No hay acción para ello, a propósito. Una bitácora de intentos que el
 * auditado puede reescribir no sirve para auditar nada. Esta aserción
 * existe para que agregar esa acción «por comodidad» rompa una prueba.
 */
ok('no existe apiNotaEditar', typeof globalThis.apiNotaEditar === 'undefined');
ok('no existe apiNotaBorrar', typeof globalThis.apiNotaBorrar === 'undefined');

console.log('\n── 6 · cada quien ve las de su tienda ──');
r = F.apiNotaAgregar(S.jaime, { entidad: 'Pedidos', id: 'pg', texto: 'de Guatemala' });
ok('Jaime anota en la suya', r.ok, JSON.stringify(r));
r = F.apiNotaAgregar(S.jaime, { entidad: 'Pedidos', id: 'p1', texto: 'no debería' });
ok('pero NO en un pedido de Ecuador', r.ok === false, JSON.stringify(r));
v = F.apiNotas(S.jaime, { tienda: 'gt' });
ok('y al leer solo ve las suyas',
   !!v.notas['Pedidos:pg'] && !v.notas['Pedidos:p1'],
   JSON.stringify(Object.keys(v.notas)));
r = F.apiNotas(S.jaime, { tienda: 'ec' });
ok('no puede pedir las de la otra tienda', r.ok === false, JSON.stringify(r));

console.log('\n── 7 · también en novedades ──');
r = F.apiNotaAgregar(S.zulay, { entidad: 'Novedades', id: 'n1', texto: 'reprogramada' });
ok('se anota en una novedad', r.ok, JSON.stringify(r));
v = F.apiNotas(S.zulay, { tienda: 'ec' });
igual('y va a su propio hilo, no al del pedido', 1,
      (v.notas['Novedades:n1'] || []).length);
igual('el hilo del pedido sigue con las suyas', 2, (v.notas['Pedidos:p1'] || []).length);

console.log('\n── 8 · los bordes ──');
r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'p1', texto: '   ' });
ok('una nota en blanco no se guarda', r.ok === false, JSON.stringify(r));
r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'nohay', texto: 'x' });
ok('un caso que no existe avisa', r.ok === false, JSON.stringify(r));
r = F.apiNotaAgregar(S.zulay, { entidad: 'Equipo', id: 'e1', texto: 'x' });
ok('no se anota sobre cualquier hoja', r.ok === false, JSON.stringify(r));
r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'p1',
                                texto: 'x'.repeat(F.NOTA_LARGO + 1) });
ok('una nota kilométrica se rechaza', r.ok === false, JSON.stringify(r));
igual('y nada de eso ensució la bitácora', 2,
      (F.apiNotas(S.zulay, { tienda: 'ec' }).notas['Pedidos:p1'] || []).length);

delete LIBROS.emp.Notas;
F.libroOlvidar_();
v = F.apiNotas(S.zulay, { tienda: 'ec' });
ok('sin la hoja Notas la pantalla abre igual', v.ok === true, JSON.stringify(v));
r = F.apiNotaAgregar(S.zulay, { entidad: 'Pedidos', id: 'p1', texto: 'x' });
ok('y anotar dice qué falta', r.ok === false && /bootstrapTodo/.test(r.error || ''),
   JSON.stringify(r));

console.log('\n── 9 · la hoja está declarada ──');
['id','tienda','entidad','entidad_id','texto','autor','autor_nombre','creado_en']
  .forEach(c => ok('columna ' + c, (C_NOT || []).indexOf(c) !== -1));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
