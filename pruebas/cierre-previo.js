/**
 * `cierre_previo` EN EL SERVIDOR: qué falta para un cierre limpio.
 *
 * La pantalla ya está probada en `cierre-revision.js`. Aquí se prueba
 * lo de abajo: que la lista de pedidos que devuelve sea LA CORRECTA.
 *
 * ┌─ LO QUE TIENE QUE ACERTAR ─────────────────────────────────┐
 * │                                                            │
 * │ Solo de ESA tienda, solo de ESE mes, y solo los que no      │
 * │ tienen desenlace. Cualquiera de las tres mal y la dueña     │
 * │ congela un mes con cifras que no son las suyas — y un mes   │
 * │ cerrado solo se reabre editando la hoja a mano.             │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Y que NO ESCRIBA NADA. Es una pantalla para mirar: si escribiera,
 * abrirla sería un acto y no se podría abrir por curiosidad.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const HOJAS = {};
const ESCRITURAS = [];
function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  const anota = (q) => { ESCRITURAS.push(nombre + ':' + q); };
  return {
    getName: () => nombre,
    getLastRow: () => m.length,
    getLastColumn: () => (m[0] || []).length,
    getDataRange: () => ({ getValues: () => m }),
    appendRow: () => anota('appendRow'),
    getRange: (f, c, nf, nc) => ({
      getValues: () => m.slice(f - 1, f - 1 + (nf || 1))
        .map((x) => x.slice(c - 1, c - 1 + (nc || (x || []).length))),
      setValues: () => anota('setValues'),
      setValue: () => anota('setValue'),
    }),
  };
}
const SS = { getSheetByName: hoja, getId: () => 'emp' };
global.SpreadsheetApp = { openById: () => SS, flush: () => {} };
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS }) };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [] };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
global.UrlFetchApp = {};
global.MailApp = { sendEmail: () => {} };
global.Utilities = { sleep: () => {}, getUuid: () => 'u',
  formatDate: (d, tz, pat) => {
    const iso = new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    return iso;
  } };

(0, eval)(src + '\n;globalThis.__F = { apiCierrePrevio, apiCerrarMes, libroOlvidar_, ' +
  'ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const cols = (t) => F.ESQUEMA_EMPRESARIAL[t];
const filaDe = (t, v) => cols(t).map((c) => (v[c] === undefined ? '' : v[c]));

/** El mes que se cierra es el anterior; se calcula para que no caduque. */
const hoy = new Date();
const ant = new Date(hoy.getFullYear(), hoy.getMonth() - 1, 15);
const MES = ant.toISOString().slice(0, 7);
const OTRO = new Date(hoy.getFullYear(), hoy.getMonth() - 2, 15).toISOString().slice(0, 7);
const HOY = hoy.toISOString().slice(0, 10);

const ped = (v) => filaDe('Pedidos', Object.assign({
  tienda: 'ec', fecha: MES + '-10', cliente: 'X', producto: 'TÓNICO',
  ciudad: 'Quito', valor: 50, ultimo_movimiento: MES + '-12',
}, v));

function montar() {
  Object.keys(HOJAS).forEach((k) => delete HOJAS[k]);
  ESCRITURAS.length = 0;
  HOJAS.Parametros = [cols('Parametros'),
    filaDe('Parametros', { clave: 'moneda_reporte', valor: 'USD' })];
  HOJAS.Tiendas = [cols('Tiendas'),
    filaDe('Tiendas', { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD',
                        zona_horaria: 'UTC', estado: 'activa' }),
    filaDe('Tiendas', { id: 'gt', nombre: 'Nutrea GT', moneda: 'GTQ',
                        zona_horaria: 'UTC', estado: 'activa' })];
  HOJAS.Pedidos = [cols('Pedidos'),
    // Los tres que faltan, cada uno en un estado distinto
    ped({ id: 'p1', cliente: 'ANA QUIETA', estado_nova: 'en_transito',
          ultimo_movimiento: MES + '-02' }),
    ped({ id: 'p2', cliente: 'BEA', estado_nova: 'pendiente', valor: 70 }),
    ped({ id: 'p3', cliente: 'CARO', estado_nova: 'novedad', valor: 30 }),
    // Los que YA tienen desenlace: no cuentan
    ped({ id: 'p4', cliente: 'ENTREGADA', estado_nova: 'entregado', valor: 200 }),
    ped({ id: 'p5', cliente: 'DEVUELTA', estado_nova: 'devolucion' }),
    ped({ id: 'p6', cliente: 'CANCELADA', estado_nova: 'cancelado' }),
    // Uno abierto de OTRO MES
    ped({ id: 'p7', cliente: 'DEL MES PASADO', estado_nova: 'pendiente',
          fecha: OTRO + '-10', ultimo_movimiento: OTRO + '-11' }),
    // Y uno abierto de OTRA TIENDA
    ped({ id: 'p8', cliente: 'DE GUATEMALA', estado_nova: 'pendiente', tienda: 'gt' }),
  ];
  ['Pauta', 'Gastos', 'Cartera', 'Facturacion', 'Novedades', 'Cierres',
   'Tasas', 'CAS', 'Inventario', 'Equipo'].forEach((t) => { HOJAS[t] = [cols(t)]; });
  F.libroOlvidar_();
}

const dueña = { sheetId: 'emp', rol: 'dueno', email: 'm@n.com', nombre: 'Manuela',
                tiendas: ['ec', 'gt'], permisos: [], modulos: ['empresarial'] };

console.log('\n══ 1 · SOLO LOS QUE FALTAN, DE ESA TIENDA Y ESE MES ══');
montar();
let r = F.apiCierrePrevio(dueña, { tienda: 'ec', mes: MES });
ok('contesta', r.ok === true, JSON.stringify(r.error));
igual('los tres sin desenlace, y nadie más',
      ['p1', 'p2', 'p3'], (r.abiertos || []).map((x) => x.id).sort());
ok('el de la otra tienda no está',
   (r.abiertos || []).every((x) => x.id !== 'p8'));
ok('el de otro mes tampoco',
   (r.abiertos || []).every((x) => x.id !== 'p7'));
ok('ni los que ya tienen desenlace',
   (r.abiertos || []).every((x) => ['p4', 'p5', 'p6'].indexOf(x.id) === -1));
igual('suma la plata que se quedaría fuera', 150, r.valorAbierto);
igual('y no dice que esté limpio', false, r.limpio);

console.log('\n══ 2 · ORDENADOS POR LO QUE LLEVA MÁS QUIETO ══');
/**
 * Lo más quieto primero: es lo que menos probable es que se resuelva
 * solo, así que es lo primero que hay que tocar.
 */
igual('el más quieto va de primero', 'p1', (r.abiertos || [])[0].id);
ok('y se dice cuántos días lleva', (r.abiertos || [])[0].diasQuieto > 0,
   String((r.abiertos || [])[0].diasQuieto));
ok('cada uno trae con qué reconocerlo',
   (r.abiertos || []).every((x) => x.id && x.cliente && x.estado),
   JSON.stringify((r.abiertos || [])[0]));
igual('y se agrupan por estado, con su plata',
      [['en_transito', 1, 50], ['novedad', 1, 30], ['pendiente', 1, 70]],
      (r.porEstado || []).map((g) => [g.estado, g.n, g.valor])
        .sort((a, b) => (a[0] < b[0] ? -1 : 1)));

console.log('\n══ 3 · TRAE EL ANÁLISIS, NO SOLO EL AVISO ══');
/**
 * «sin la opción de editar o revisar o mirar y analizar». El mismo
 * análisis del cierre, no una segunda cuenta: dos cuentas del mismo mes
 * acaban diciendo cosas distintas y no hay a cuál creerle.
 */
ok('las cifras del mes', r.actual && r.actual.ventas === 200,
   JSON.stringify(r.actual && r.actual.ventas));
ok('y las del mes anterior, para comparar', !!r.anterior);
igual('dice que no está cerrado', false, r.cerrado);

console.log('\n══ 4 · MIRAR NO ES UN ACTO: NO ESCRIBE NADA ══');
/**
 * Si abrirlo escribiera, no se podría abrir por curiosidad — y esta
 * pantalla existe justamente para mirar antes de decidir.
 */
montar();
F.apiCierrePrevio(dueña, { tienda: 'ec', mes: MES });
igual('ni una escritura', [], ESCRITURAS);

console.log('\n══ 5 · UN MES LIMPIO LO DICE ══');
montar();
// Se les da desenlace a los tres
[1, 2, 3].forEach((i) => {
  HOJAS.Pedidos[i][cols('Pedidos').indexOf('estado_nova')] = 'entregado';
});
F.libroOlvidar_();
r = F.apiCierrePrevio(dueña, { tienda: 'ec', mes: MES });
igual('sin pedidos colgando', [], r.abiertos);
igual('y lo dice', true, r.limpio);
igual('sin plata fuera', 0, r.valorAbierto);

console.log('\n══ 6 · QUIÉN PUEDE, Y QUÉ NO SE ACEPTA ══');
montar();
const gestora = { sheetId: 'emp', rol: 'gestora', email: 's@x.com', nombre: 'Sara',
                  tiendas: ['ec'], permisos: [], modulos: ['empresarial'] };
igual('una gestora no revisa un cierre', false,
      F.apiCierrePrevio(gestora, { tienda: 'ec', mes: MES }).ok);
igual('ni la dueña en una tienda que no es suya', false,
      F.apiCierrePrevio(dueña, { tienda: 'mx', mes: MES }).ok);
igual('un mes mal escrito se rechaza en vez de adivinarse', false,
      F.apiCierrePrevio(dueña, { tienda: 'ec', mes: 'septiembre' }).ok);

console.log('\n══ 7 · Y AVISA SI EL MES TODAVÍA CORRE ══');
/**
 * Cerrar el mes en curso congela unas cifras a las que les faltan días.
 * Casi siempre es un error de dedo, así que se dice antes.
 */
montar();
const enCurso = F.apiCierrePrevio(dueña, { tienda: 'ec', mes: HOY.slice(0, 7) });
igual('el mes de hoy se marca como en curso', true, enCurso.mesEnCurso);
igual('y el anterior no',
      false, F.apiCierrePrevio(dueña, { tienda: 'ec', mes: MES }).mesEnCurso);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
