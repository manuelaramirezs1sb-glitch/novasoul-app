/**
 * META, todo en una pantalla — y el semáforo de las tres revisiones.
 *
 * Ella pidió dos cosas que aquí se comprueban:
 *
 * 1· «No veo dónde se ven los datos que lee Nova de Meta […] todo está
 *    regado». Una sola respuesta con la conexión, los periodos y el
 *    desglose, para que la pantalla no tenga que encadenar llamadas.
 *
 * 2· «Sacar un semáforo diario, son 3 revisiones por día, al mediodía, a
 *    las 5pm y a las 11:30pm». Y lo que importa de esas tres: que NO
 *    digan lo mismo. A mediodía todavía se corrige; a las 11:30 ya no, y
 *    dar un consejo que no se puede seguir es peor que callarse.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, cli: {} };
function libro(id) {
  const hojas = LIBROS[id] || {};
  return {
    getSheetByName: (nombre) => {
      const m = hojas[nombre];
      if (!m) return null;
      return {
        getName: () => nombre,
        getLastRow: () => m.length,
        getLastColumn: () => (m[0] ? m[0].length : 0),
        getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
        appendRow: (f) => m.push(f.slice()),
        getRange: (f, c, nf, nc) => ({
          getValues: () => {
            const out = [];
            for (let i = 0; i < (nf || 1); i++) {
              const fila = m[f - 1 + i] || [];
              out.push(fila.slice(c - 1, c - 1 + (nc || fila.length)));
            }
            return out;
          },
          setValues: (v) => {
            v.forEach((fila, i) => {
              if (!m[f - 1 + i]) m[f - 1 + i] = [];
              fila.forEach((val, j) => { m[f - 1 + i][c - 1 + j] = val; });
            });
          },
          setValue: (v) => { if (!m[f - 1]) m[f - 1] = []; m[f - 1][c - 1] = v; },
        }),
      };
    },
    insertSheet: (n) => { hojas[n] = []; return libro(id).getSheetByName(n); },
  };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
let HORA = '14:00';
const HOY = '2026-09-23T19:00:00Z';
const GUARDADAS = {};
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => GUARDADAS[k] !== undefined ? GUARDADAS[k] : (PROPS[k] || ''),
  getProperties: () => PROPS,
  setProperty: (k, v) => { GUARDADAS[k] = v; }, deleteProperty: (k) => { delete GUARDADAS[k]; } }) };
global.SpreadsheetApp = { openById: (id) => libro(id), flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'America/Bogota',
                   getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{"data":[]}' }) };
let UUID = 0;
global.Utilities = { sleep: () => {}, getUuid: () => 'u' + (++UUID) + '0000000',
  formatDate: (d, tz, pat) => {
    if (pat === 'HH:mm') return HORA;
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

(0, eval)(src + '\n;globalThis.__F = { apiMetaPanel, metaRevisiones_, metaNumeros_,' +
  ' metaFraccionDia_, META_REVISIONES };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PED = ['id','tienda','fecha','producto','estado_canonico','estado_nova',
               'valor','costo_producto','costo_envio'];
const C_PAU = ['id','fecha','tienda','plataforma','cuenta','campana','conjunto',
               'gasto','moneda_gasto','impresiones','alcance','clics','resultados'];
const C_ANU = ['id','fecha','tienda','campana','anuncio','gasto','moneda_gasto',
               'impresiones','clics','resultados'];

function ped(o) { return C_PED.map(c => (o[c] === undefined ? '' : o[c])); }
function pau(o) { return C_PAU.map(c => (o[c] === undefined ? '' : o[c])); }

function sembrar(opts) {
  /**
   * Cada caso de prueba es una PETICIÓN nueva, y una petición real
   * empieza soltando lo que Nova tenga en memoria de la anterior
   * (`manejar()` lo hace). Aquí hay que decirlo a mano porque las
   * pruebas llaman a las funciones directamente, sin pasar por ahí.
   *
   * Sin esto, el manejador del libro y las hojas ya leídas seguirían
   * apuntando a los datos del caso anterior, y la prueba mediría un
   * Nova que no existe.
   */
  libroOlvidar_(); soulOlvidar_();

  opts = opts || {};
  Object.keys(GUARDADAS).forEach(k => delete GUARDADAS[k]);
  if (opts.llave !== false) GUARDADAS['META_TOKEN_cli'] = 'xxx';
  LIBROS.cli = {
    Pedidos: [C_PED].concat(opts.pedidos || []),
    Pauta: [C_PAU].concat(opts.pauta || []),
    Anuncios: [C_ANU].concat(opts.anuncios || []),
    Tiendas: [['id','nombre','moneda','estado'], ['ec','Nutrea EC','USD','activa']],
    Parametros: [['tienda','clave','valor','nota','tipo']].concat(opts.parametros || [
      ['ec', 'meta_cuenta', '11122233', '', ''],
      ['ec', 'cpa_verde_pct', '70', '', ''],
      ['ec', 'presupuesto_diario_pauta', '100', '', ''],
    ]),
    Tasas: [['fecha','de','a','tasa']],
    Cierres: [['tienda','mes','cerrado_en']],
  };
}

const DUENA = { sheetId: 'cli', rol: 'dueno', email: 'manuela@nova.com',
                nombre: 'Manuela', tiendas: ['ec'] };

/** Un día con 10 pedidos, 6 entregados, y 60 de gasto. */
function diaNormal(fecha) {
  const out = [];
  for (let i = 0; i < 10; i++) {
    out.push(ped({ id: 'p' + fecha + i, tienda: 'ec', fecha: fecha,
      producto: 'TAG RECEDE',
      estado_canonico: i < 6 ? 'entregado' : (i < 8 ? 'devolucion' : 'en_transito'),
      valor: 50, costo_producto: 12, costo_envio: 5 }));
  }
  return out;
}

// ─────────────────────────────────────────────────────────────
console.log('\n── Todo de Meta, en una sola respuesta ──');
sembrar({
  pedidos: diaNormal('2026-09-23').concat(diaNormal('2026-09-22'), diaNormal('2026-09-20')),
  pauta: [
    pau({ id: 'a1', fecha: '2026-09-23', tienda: 'ec', campana: 'TAG · frío',
          gasto: 40, moneda_gasto: 'USD', impresiones: 9000, alcance: 6000,
          clics: 180, resultados: 7 }),
    pau({ id: 'a2', fecha: '2026-09-23', tienda: 'ec', campana: 'TAG · retargeting',
          gasto: 20, moneda_gasto: 'USD', impresiones: 3000, alcance: 2000,
          clics: 90, resultados: 3 }),
    pau({ id: 'a3', fecha: '2026-09-22', tienda: 'ec', campana: 'TAG · frío',
          gasto: 55, moneda_gasto: 'USD', impresiones: 8000, alcance: 5500,
          clics: 150, resultados: 6 }),
    pau({ id: 'a4', fecha: '2026-09-20', tienda: 'ec', campana: 'TAG · frío',
          gasto: 50, moneda_gasto: 'USD', impresiones: 7000, alcance: 5000,
          clics: 140, resultados: 5 }),
  ],
  anuncios: [
    C_ANU.map(c => ({ id: 'n1', fecha: '2026-09-23', tienda: 'ec',
      campana: 'TAG · frío', anuncio: 'Video testimonio', gasto: 30,
      moneda_gasto: 'USD', impresiones: 6000, clics: 120, resultados: 5 })[c] || ''),
    C_ANU.map(c => ({ id: 'n2', fecha: '2026-09-23', tienda: 'ec',
      campana: 'TAG · frío', anuncio: 'Carrusel antes y después', gasto: 10,
      moneda_gasto: 'USD', impresiones: 3000, clics: 60, resultados: 2 })[c] || ''),
  ],
});
let r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
ok('responde', r.ok === true, JSON.stringify(r).slice(0, 200));
igual('con el nombre de pantalla de la tienda', 'Nutrea EC', r.nombreTienda);
igual('y su moneda', 'USD', r.moneda);

console.log('\n── La conexión, en la misma pantalla ──');
igual('dice que está activo', true, r.conexion.activo);
igual('con su cuenta publicitaria', '11122233', r.conexion.cuenta);
ok('y cuándo entró la última fila de pauta',
   r.conexion.ultimaLectura === '2026-09-23', r.conexion.ultimaLectura);
ok('sin nada que explicar, porque está bien', r.conexion.porque === '');

sembrar({ llave: false, pauta: [] });
let sinLlave = F.apiMetaPanel(DUENA, { tienda: 'ec' });
igual('sin llave, no está activo', false, sinLlave.conexion.activo);
igual('y dice qué falta', 'llave', sinLlave.conexion.falta);
ok('con palabras, no con un código',
   /no has guardado la llave/.test(sinLlave.conexion.porque), sinLlave.conexion.porque);

sembrar({ parametros: [['ec', 'cpa_verde_pct', '70', '', '']] });
let sinCuenta = F.apiMetaPanel(DUENA, { tienda: 'ec' });
igual('con llave pero sin cuenta, tampoco está activo', false, sinCuenta.conexion.activo);
igual('y lo dice', 'cuenta', sinCuenta.conexion.falta);

console.log('\n── Los cuatro periodos ──');
sembrar({
  pedidos: diaNormal('2026-09-23').concat(diaNormal('2026-09-22'), diaNormal('2026-09-20')),
  pauta: [
    pau({ id: 'a1', fecha: '2026-09-23', tienda: 'ec', campana: 'TAG · frío',
          gasto: 60, moneda_gasto: 'USD', impresiones: 12000, alcance: 8000, clics: 270 }),
    pau({ id: 'a3', fecha: '2026-09-22', tienda: 'ec', campana: 'TAG · frío',
          gasto: 55, moneda_gasto: 'USD', impresiones: 8000, alcance: 5500, clics: 150 }),
    pau({ id: 'a4', fecha: '2026-09-20', tienda: 'ec', campana: 'TAG · frío',
          gasto: 50, moneda_gasto: 'USD', impresiones: 7000, alcance: 5000, clics: 140 }),
  ],
});
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
igual('son cuatro', ['hoy', 'semana', 'mes', 'trimestre'], r.periodos.map(x => x.id));
const P = {}; r.periodos.forEach(x => { P[x.id] = x; });
igual('hoy: 10 pedidos y 60 de gasto', [10, 60], [P.hoy.pedidos, P.hoy.gasto]);
igual('y su CPA', 6, P.hoy.cpa);
igual('los 7 días acumulan los tres días', [30, 165], [P.semana.pedidos, P.semana.gasto]);
ok('el mes arranca el día 1', P.mes.desde === '2026-09-01', P.mes.desde);
ok('el trimestre son 90 días', P.trimestre.desde === '2026-06-26', P.trimestre.desde);
igual('el ROAS sale de las ventas entregadas', 5, P.hoy.roas);
ok('y la utilidad después de pauta',
   P.hoy.utilidadReal === P.hoy.utilidadAntesPauta - P.hoy.gasto,
   JSON.stringify([P.hoy.utilidadAntesPauta, P.hoy.gasto, P.hoy.utilidadReal]));
igual('el alcance viene de Meta', [12000, 8000, 270],
      [P.hoy.alcance.impresiones, P.hoy.alcance.alcance, P.hoy.alcance.clics]);
ok('con su CTR', P.hoy.alcance.ctr === 2.25, P.hoy.alcance.ctr);

console.log('\n── Las tres revisiones del día ──');
igual('son tres', ['mediodia', 'tarde', 'cierre'], r.revisiones.map(x => x.id));
igual('a sus horas', ['12:00', '17:00', '23:30'], r.revisiones.map(x => x.hora));
/** Son las 14:00: mediodía ya pasó, las otras dos no. */
igual('a las 2 de la tarde solo ha pasado la del mediodía',
      [true, false, false], r.revisiones.map(x => x.yaPaso));
ok('las que no han llegado lo dicen, no se pintan de un color',
   r.revisiones[1].estado === 'pendiente' && r.revisiones[2].estado === 'pendiente');
ok('la del mediodía sí tiene color',
   ['verde', 'amarillo', 'rojo'].indexOf(r.revisiones[0].estado) !== -1,
   r.revisiones[0].estado);
ok('y explica por qué, con números',
   /CPA del día/.test(r.revisiones[0].porque) && /techo/.test(r.revisiones[0].porque),
   r.revisiones[0].porque);
/**
 * Lo importante: el presupuesto se juzga contra la PARTE del día que va.
 * A mediodía toca la mitad de los 100, no los 100.
 */
ok('el presupuesto se compara con la mitad del día, no con el día entero',
   /de los 50 que tocarían a esta hora/.test(r.revisiones[0].porque),
   r.revisiones[0].porque);

console.log('\n── Y a las 11:40 de la noche ──');
HORA = '23:40';
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
igual('ya pasaron las tres', [true, true, true], r.revisiones.map(x => x.yaPaso));
ok('la de la noche dice que ya no hay nada que mover',
   /no es para corregir|ya cerró/.test(r.revisiones[2].porque), r.revisiones[2].porque);
ok('y el consejo de la noche no manda mover presupuesto',
   !/baja ese presupuesto|sube.*presupuesto/i.test(r.queHacer), r.queHacer);

HORA = '14:00';

console.log('\n── Qué hacer con lo que hay ──');
sembrar({ pedidos: diaNormal('2026-09-23'),
  pauta: [pau({ id: 'x', fecha: '2026-09-23', tienda: 'ec', campana: 'c',
                gasto: 500, moneda_gasto: 'USD' })] });
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
ok('con el CPA por las nubes, dice que baje presupuesto',
   /por encima del techo/.test(r.queHacer), r.queHacer);
ok('y el semáforo del mediodía está en rojo', r.revisiones[0].estado === 'rojo');

sembrar({ pedidos: [], pauta: [pau({ id: 'x', fecha: '2026-09-23', tienda: 'ec',
  campana: 'c', gasto: 80, moneda_gasto: 'USD' })] });
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
ok('con gasto y cero pedidos NO pinta rojo: dice que no hay con qué saberlo',
   r.revisiones[0].estado === 'sin_datos', r.revisiones[0].estado);
ok('y propone pausar si sigue igual', /pausar/.test(r.queHacer), r.queHacer);

sembrar({ pedidos: [], pauta: [] });
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
ok('un día sin nada dice que falta traer la lectura',
   /falte traer la lectura/.test(r.queHacer), r.queHacer);

console.log('\n── La pauta sin convertir NO se esconde ──');
/**
 * Si Meta cobró en otra moneda y falta la tasa, ese gasto no entra al
 * CPA. Dar el CPA sin decirlo es dar por bueno un número al que le falta
 * plata — y es justo el que decide si se escala o se frena.
 */
sembrar({ pedidos: diaNormal('2026-09-23'),
  pauta: [pau({ id: 'x', fecha: '2026-09-23', tienda: 'ec', campana: 'c',
                gasto: 200000, moneda_gasto: 'COP' })] });
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
const hoyP = r.periodos.filter(x => x.id === 'hoy')[0];
igual('el gasto sin convertir va aparte', 200000, hoyP.gastoSinConvertir);
igual('y NO se suma al gasto bueno', 0, hoyP.gasto);
ok('y es lo primero que se dice',
   /sin convertir/.test(r.queHacer) && /NO está dentro del CPA/.test(r.queHacer),
   r.queHacer);

console.log('\n── El desglose ──');
sembrar({
  pedidos: diaNormal('2026-09-23'),
  pauta: [
    pau({ id: 'a1', fecha: '2026-09-23', tienda: 'ec', campana: 'TAG · frío',
          gasto: 40, moneda_gasto: 'USD', resultados: 7 }),
    pau({ id: 'a2', fecha: '2026-09-23', tienda: 'ec', campana: 'TAG · retargeting',
          gasto: 20, moneda_gasto: 'USD', resultados: 3 }),
  ],
  anuncios: [
    C_ANU.map(c => ({ id: 'n1', fecha: '2026-09-23', tienda: 'ec',
      campana: 'TAG · frío', anuncio: 'Video testimonio', gasto: 30,
      moneda_gasto: 'USD', resultados: 5 })[c] || ''),
    C_ANU.map(c => ({ id: 'n2', fecha: '2026-09-23', tienda: 'ec',
      campana: 'TAG · frío', anuncio: 'Carrusel', gasto: 10,
      moneda_gasto: 'USD', resultados: 2 })[c] || ''),
  ],
});
r = F.apiMetaPanel(DUENA, { tienda: 'ec' });
igual('las campañas, de la que más gasta a la que menos',
      ['TAG · frío', 'TAG · retargeting'], r.desglose.campanas.map(x => x.nombre));
igual('con su gasto y su CPA', [40, 5.71],
      [r.desglose.campanas[0].gasto, r.desglose.campanas[0].cpa]);
igual('y los anuncios, uno por uno',
      ['Video testimonio', 'Carrusel'], r.desglose.anuncios.map(x => x.nombre));

console.log('\n── Quién puede ──');
const GESTORA = { sheetId: 'cli', rol: 'gestora', email: 'k@nova.com',
                  nombre: 'Kat', tiendas: ['ec'] };
ok('una gestora no ve la pauta', F.apiMetaPanel(GESTORA, { tienda: 'ec' }).ok === false);
ok('y nadie ve una tienda que no es suya',
   F.apiMetaPanel(DUENA, { tienda: 'gt' }).ok === false);
ok('la llave NUNCA vuelve en la respuesta',
   JSON.stringify(F.apiMetaPanel(DUENA, { tienda: 'ec' })).indexOf('xxx') === -1);

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
