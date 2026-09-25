/**
 * ¿LLEGA A NOVA LO QUE YA ESTABA ESCRITO EN EL ARCHIVO?
 *
 * ┌─ LA PREGUNTA ──────────────────────────────────────────────┐
 * │                                                            │
 * │ «si en el archivo ya están las notas de las gestoras y     │
 * │  cuánto se contactó al cliente, ¿eso aparece reflejado en  │
 * │  Nova, o solo la parte de los clientes, estados y el       │
 * │  resumen general, y nada de los comentarios — comenzaría   │
 * │  de 0?»                                                    │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE SE ENCONTRÓ AL IR A MIRAR ─────────────────────────┐
 * │                                                            │
 * │ Empezaba de cero, y por dos motivos distintos:             │
 * │                                                            │
 * │ 1· Las NOTAS del pedido. Dropi las llama «notas». El        │
 * │    importador las leía y las mandaba a un campo             │
 * │    `observacion` que NO EXISTÍA como columna. Se usaban    │
 * │    para pescar un segundo teléfono y se tiraban.           │
 * │                                                            │
 * │ 2· Las SOLUCIONES de las novedades. El alias las leía      │
 * │    siempre, pero `derivarNovedades` —la función que arma   │
 * │    la fila de la novedad— nunca las copiaba. Se importaban │
 * │    setenta novedades resueltas y las setenta llegaban sin  │
 * │    una palabra de lo que ya se había hecho.                │
 * │                                                            │
 * │ Las dos son pérdidas silenciosas: el import decía «72      │
 * │ filas» y nadie podía saber que faltaba la mitad de cada    │
 * │ una.                                                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Y la otra mitad de la pregunta —«cuánto se contactó al cliente»— se
 * comprueba aquí abajo: NO viene de los archivos de pedidos. Viene de
 * IRIS, la central telefónica, y solo si se importa.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libroStub(id) {
  const hojas = LIBROS[id] || {};
  return {
    getSheetByName: (n) => {
      const m = hojas[n];
      if (!m) return null;
      return {
        getName: () => n, getLastRow: () => m.length,
        getLastColumn: () => (m[0] ? m[0].length : 0),
        getMaxColumns: () => Math.max(m[0] ? m[0].length : 0, 26),
        getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
        appendRow: (f) => m.push(f.slice()),
        deleteRow: (k) => m.splice(k - 1, 1),
        clear: () => { m.length = 1; },
        setFrozenRows: () => {}, autoResizeColumns: () => {}, setTabColor: () => {},
        deleteColumns: () => {},
        getRange: (f, c, nf, nc) => {
          const api = {
            getValues: () => {
              const o = [];
              for (let i = 0; i < (nf || 1); i++) {
                const r = m[f - 1 + i] || [];
                o.push(r.slice(c - 1, c - 1 + (nc || r.length)));
              }
              return o;
            },
            setValues: (v) => { v.forEach((r, i) => {
              const d = m[f - 1 + i] || (m[f - 1 + i] = []);
              r.forEach((x, j) => { d[c - 1 + j] = x; }); }); return api; },
            setValue: (x) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = x; return api; },
            setFontWeight: () => api, setBackground: () => api, setFontColor: () => api,
            setNumberFormat: () => api, setNote: () => api, clearContent: () => api,
          };
          return api;
        },
      };
    },
    insertSheet: (n) => { hojas[n] = [[]]; return libroStub(id).getSheetByName(n); },
    getSheets: () => Object.keys(hojas).map(n => libroStub(id).getSheetByName(n)),
    deleteSheet: () => {}, getSpreadsheetTimeZone: () => 'UTC',
  };
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

(0, eval)(src + '\n;globalThis.__F = { importar, FUENTES, COLUMNAS_DEL_EQUIPO,' +
  ' libroOlvidar_, soulOlvidar_, ESQUEMA_EMPRESARIAL };' +
  '\n;globalThis.diccionarioDe_ = diccionarioDe_;' +
  '\n;globalThis.analizarFilas = analizarFilas;' +
  '\n;globalThis.normalizarCas_ = normalizarCas_;' +
  '\n;globalThis.FUENTES = FUENTES;');
const F = globalThis.__F;
const E = F.ESQUEMA_EMPRESARIAL;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** Leer una fila ya importada, como objeto. */
function filaDe(hoja, id) {
  const m = LIBROS.emp[hoja];
  const enc = m[0].map(x => String(x || '').toLowerCase().trim());
  for (let i = 1; i < m.length; i++) {
    if (String(m[i][enc.indexOf('id')]).indexOf(id) !== -1) {
      const o = {};
      enc.forEach((c, j) => { o[c] = m[i][j]; });
      return o;
    }
  }
  return null;
}

/**
 * El export de Dropi, con las columnas que de verdad trae — incluidas
 * las dos que interesan: «NOTAS» (lo que escribía la gestora sobre el
 * pedido) y «SOLUCION» (lo que escribió para cerrar la novedad).
 */
const CRUDO_DROPI = [
  ['ID', 'FECHA', 'NOMBRE CLIENTE', 'TELEFONO', 'CIUDAD', 'ESTATUS', 'PRODUCTO',
   'TOTAL DE LA ORDEN', 'NOVEDAD', 'SOLUCION', 'FUE SOLUCIONADA LA NOVEDAD',
   'FECHA DE SOLUCION', 'NOTAS', 'ULTIMO MOVIMIENTO', 'DIRECCION'],
  ['7781', '2026-09-10', 'Marcela Andrade', '0991110011', 'Quito', 'ENTREGADO',
   'TAG RECEDE', '50', 'DIRECCION ERRADA',
   'Hablé con ella, me dio la referencia del parque. Reprogramada para el jueves.',
   'SI', '2026-09-12',
   'Clienta difícil de ubicar en la mañana. Llamar después de las 2pm.',
   '2026-09-12', 'Av. Siempre Viva 123'],
  ['7790', '2026-09-11', 'Rosa Benítez', '0992220022', 'Guayaquil', 'EN TRANSITO',
   'TAG RECEDE', '50', '', '', '', '',
   'Pidió que le llamen al 0993334455 que es el de la hija.',
   '2026-09-14', 'Calle 10 #4-55'],
];

function sembrar() {
  UUID = 0; F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Pedidos: [E.Pedidos.slice()], Novedades: [E.Novedades.slice()],
    Llamadas: [E.Llamadas.slice()],
    Tiendas: [['id','nombre','moneda','estado','modalidad','pais','zona_horaria'],
              ['ec','Nutrea EC','USD','activa','marca_propia','EC','America/Guayaquil']],
    Estados: [['fuente','texto','estado_nova','origen','pedidos','primera_vez',
               'ultima_vez','decidido_por','nota']],
    Fuentes: [['id','nombre','tipo','estado','tienda','fuente','ultima_importacion','filas_ultima']],
    Parametros: [['tienda','clave','valor','nota','tipo']],
    Movimientos: [['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo']],
    _Import_Dropi: CRUDO_DROPI.map(f => f.slice()),
  };
  LIBROS.cen = { Clientes: [['id','empresa','correo','plan','x','x','estado','x','x','x',
                             'x','x','x','sheet_id'],
                            ['c1','Nutrea','m@nova.com','pro','','','activo','','','','','','','emp']] };
}

// ══════════════════════════════════════════════════════════════
console.log('\n── Se importa el archivo tal como lo baja ella ──');
sembrar();
let msg = '';
try { msg = F.importar('dropi', 'ec'); }
catch (e) { msg = 'ERROR: ' + e.message; }
ok('el import corre', !/^ERROR/.test(msg), msg);
igual('entran los dos pedidos', 2, LIBROS.emp.Pedidos.length - 1);
igual('y la novedad que traía uno', 1, LIBROS.emp.Novedades.length - 1);

console.log('\n── LO QUE ESCRIBIÓ LA GESTORA SOBRE EL PEDIDO ──');
const p1 = filaDe('Pedidos', '7781');
ok('el pedido llegó', !!p1);
/**
 * ESTA es la pregunta. Antes esta nota se leía del archivo, se usaba
 * para pescar un segundo teléfono, y se tiraba — porque el campo al
 * que iba no tenía columna en la hoja.
 */
ok('la nota del archivo SÍ queda guardada',
   /difícil de ubicar en la mañana/.test(String(p1.observacion || '')),
   'observacion = ' + JSON.stringify(p1.observacion));
ok('y no se disfraza de nota del equipo', !String(p1.nota || '').trim(),
   'nota = ' + JSON.stringify(p1.nota));

console.log('\n── LO QUE ESCRIBIÓ PARA CERRAR LA NOVEDAD ──');
const n1 = filaDe('Novedades', '7781');
ok('la novedad llegó', !!n1);
igual('con su motivo', 'DIRECCION ERRADA', n1.motivo);
igual('y marcada como resuelta', 'resuelta', n1.estado);
/**
 * Y esto es lo que `derivarNovedades` nunca copiaba: el texto de lo
 * que se hizo. Sin él, setenta novedades resueltas llegaban sin una
 * palabra de por qué.
 */
ok('el histórico de soluciones SÍ queda',
   /referencia del parque/.test(String(n1.solucion_plataforma || '')),
   'solucion_plataforma = ' + JSON.stringify(n1.solucion_plataforma));
ok('y NO se mete en la columna del equipo', !String(n1.solucion || '').trim(),
   'solucion = ' + JSON.stringify(n1.solucion));

console.log('\n── Lo que el equipo escriba en Nova no se pisa ──');
/**
 * El equipo escribe su nota y su solución DENTRO de Nova. Una segunda
 * importación del mismo archivo no puede borrarlas: es la regla que ya
 * existía y que hay que sostener ahora que hay dos columnas parecidas.
 */
const encP = LIBROS.emp.Pedidos[0].map(x => String(x).toLowerCase());
const encN = LIBROS.emp.Novedades[0].map(x => String(x).toLowerCase());
LIBROS.emp.Pedidos[1][encP.indexOf('nota')] = 'Llamé yo el 25, quedó de confirmar.';
LIBROS.emp.Novedades[1][encN.indexOf('solucion')] = 'Cerrada por mí en Nova.';
LIBROS.emp.Pedidos[1][encP.indexOf('estado_nova')] = 'entregado';
LIBROS.emp._Import_Dropi = CRUDO_DROPI.map(f => f.slice());
F.libroOlvidar_();
F.importar('dropi', 'ec');

const p2 = filaDe('Pedidos', '7781');
const n2 = filaDe('Novedades', '7781');
igual('la nota del equipo sigue intacta', 'Llamé yo el 25, quedó de confirmar.', p2.nota);
igual('su estado en Nova también', 'entregado', p2.estado_nova);
igual('y la solución que escribió aquí', 'Cerrada por mí en Nova.', n2.solucion);
ok('mientras la del archivo se actualiza aparte',
   /referencia del parque/.test(String(n2.solucion_plataforma || '')));
igual('no se duplicaron filas', 2, LIBROS.emp.Pedidos.length - 1);

console.log('\n── Y lo que el archivo NO trae, se dice claro ──');
/**
 * «Cuánto se contactó al cliente» NO viene en ningún export de
 * pedidos. Ni Dropi ni Mastershop ni Effi traen esa cuenta.
 */
const traeIntentos = Object.keys(F.FUENTES).filter(function (k) {
  return F.FUENTES[k].alias && F.FUENTES[k].alias.intentos;
});
igual('ninguna fuente de pedidos trae «intentos»', [], traeIntentos);
igual('así que arranca vacío, no en un número inventado', '',
      String(p2.intentos || ''));

/** La única fuente que SÍ sabe cuántas veces se llamó es la central. */
ok('IRIS sí trae el historial de llamadas',
   !!F.FUENTES.iris && F.FUENTES.iris.tipo === 'llamadas');
ok('con la duración de cada una',
   !!(F.FUENTES.iris.alias.seg_conversado),
   JSON.stringify(Object.keys(F.FUENTES.iris.alias)));
ok('y con quién la hizo', !!(F.FUENTES.iris.alias.agente));

console.log('\n── Las columnas del equipo, listadas ──');
/**
 * Si alguien agrega una columna nueva donde escriba el equipo y se
 * olvida de esta lista, la siguiente importación se la borra. Por eso
 * se comprueba a la vista y no de memoria.
 */
['nota', 'solucion', 'estado_nova', 'gestora_asignada', 'intentos'].forEach(function (c) {
  ok(c + ' está protegida del importador', F.COLUMNAS_DEL_EQUIPO.indexOf(c) !== -1);
});
['observacion', 'solucion_plataforma'].forEach(function (c) {
  ok(c + ' NO lo está: es espejo del archivo y debe actualizarse',
     F.COLUMNAS_DEL_EQUIPO.indexOf(c) === -1);
});


// ══════════════════════════════════════════════════════════════
console.log('\n══ EL CONTROL DIARIO DE UNA TIENDA · sus otras dos hojas ══');
/**
 * Los encabezados son los REALES de un archivo de gestión logística que
 * ella mandó: diecisiete hojas, dos mil setecientas filas. No inventados.
 *
 *   NOVEDADES  FECHA DE GESTION · ID · FECHA NOVEDAD · CLIENTE ·
 *              SOLUCION · SOLUCIONADA · GESTIONA · NOTAS
 *
 *   CAS        FECHA RADICACION · ID · FECHA DE ENVIO ORDEN · CLIENTE ·
 *              TELEFONO · NUMERO DE GUIA · ESTATUS · TRANSPORTADORA ·
 *              FECHA DE ULTIMO MOVIMIENTO · GESTION · RESPUESTA · NOTAS
 */
const CRUDO_NOV = [
  ['FECHA DE GESTION','ID','FECHA NOVEDAD','CLIENTE','SOLUCION','SOLUCIONADA',
   'GESTIONA','NOTAS'],
  ['2026-07-31','81459518','2026-07-28','Jorge Luis Gutiérrez',
   'SE LLAMA, DICE QUE HACE DOS MESES SE MUDÓ. SE ACTUALIZA DIRECCIÓN.','SI',
   'ZULAY','CHAT'],
  ['2026-08-03','83958327','2026-07-31','Carlos Gallego',
   'Cll 40 Sur # 25 A-40, se corrige','','JAIME','CHAT - va para reenvío'],
];
const CRUDO_CAS = [
  ['FECHA RADICACION','ID','FECHA DE ENVIO ORDEN','CLIENTE','TELEFONO',
   'NUMERO DE GUIA','ESTATUS','TRANSPORTADORA','FECHA DE ULTIMO MOVIMIENTO',
   'GESTION','RESPUESTA','NOTAS'],
  ['2026-08-14','85675923','2026-08-12','Yoleisi Estacio','3228303763',
   '616297259','EN TRASLADO NACIONAL','TRANSPORTADORA','2026-08-12',
   'NO SE PUEDE RADICAR','','pendiente'],
  ['2026-08-14','85543990','2026-08-11','Yisela Carvajal','3143404365',
   '240058888893','PREPARADO PARA TRANSPORTADORA','TCC','2026-08-13',
   'SE RADICA PQR','Reponen el producto','cerrado'],
];

function analizar(tipo, crudo) {
  const dic = globalThis.diccionarioDe_(tipo);
  return globalThis.analizarFilas(crudo, dic.dicc, dic.formas);
}

console.log('\n── Nova entiende la hoja de NOVEDADES ──');
let an = analizar('novedades', CRUDO_NOV);
const puestos = {};
// El campo se llama `columna`, no `encabezado`. (Mi primera versión usó
// el nombre equivocado y las doce aserciones dieron `undefined` — que se
// lee igual que «no lo entendió» y no lo era.)
an.propuestas.forEach(function (x) { puestos[x.columna] = x.campo; });
igual('FECHA NOVEDAD → fecha', 'fecha', puestos['FECHA NOVEDAD']);
igual('SOLUCION → solucion', 'solucion', puestos['SOLUCION']);
igual('SOLUCIONADA → solucionada', 'solucionada', puestos['SOLUCIONADA']);
igual('NOTAS → nota', 'nota', puestos['NOTAS']);
/**
 * Y la que importa: GESTIONA es una ETIQUETA, no una cuenta. En su
 * archivo esa columna dice JAIME, ZULAY, APOYO — gente sin usuario en
 * Nova. Si cayera en `gestora_asignada`, que es control de acceso,
 * esos pedidos quedarían asignados a alguien que no puede entrar y no
 * los vería NADIE.
 */
igual('GESTIONA → gestionado_por, NO gestora_asignada',
      'gestionado_por', puestos['GESTIONA']);
ok('y ningún campo apunta al control de acceso',
   an.propuestas.filter(x => x.campo === 'gestora_asignada').length === 0,
   JSON.stringify(an.propuestas.map(x => x.campo)));

console.log('\n── Y la hoja de CAS ──');
an = analizar('cas', CRUDO_CAS);
const pc = {};
an.propuestas.forEach(function (x) { pc[x.columna] = x.campo; });
igual('FECHA RADICACION → abierto_en', 'abierto_en', pc['FECHA RADICACION']);
igual('NUMERO DE GUIA → guia', 'guia', pc['NUMERO DE GUIA']);
igual('ESTATUS → estado', 'estado', pc['ESTATUS']);
igual('RESPUESTA → respuesta', 'respuesta', pc['RESPUESTA']);
igual('FECHA DE ULTIMO MOVIMIENTO → ultima_gestion',
      'ultima_gestion', pc['FECHA DE ULTIMO MOVIMIENTO']);
igual('GESTION → gestion', 'gestion', pc['GESTION']);

console.log('\n── Un CAS con respuesta llega CERRADO ──');
/**
 * Sin esto, sus ciento sesenta reclamos viejos aparecerían todos
 * abiertos el primer día — y nadie vuelve a mirar una pantalla que
 * abre con ciento sesenta alarmas falsas.
 */
const caso = globalThis.normalizarCas_({
  id_externo: '85543990', abierto_en: '2026-08-14',
  ultima_gestion: '2026-08-13', respuesta: 'Reponen el producto',
}, 'propio_cas', 'col', { '85543990': 'dropi-85543990' });
ok('se cierra solo', !!caso.cerrado_en, JSON.stringify(caso.cerrado_en));
igual('y queda amarrado a su pedido', 'dropi-85543990', caso.pedido_id);

const abierto = globalThis.normalizarCas_({
  id_externo: '85675923', abierto_en: '2026-08-14', ultima_gestion: '2026-08-12',
  respuesta: '',
}, 'propio_cas', 'col', {});
ok('el que no tiene respuesta sigue abierto', !abierto.cerrado_en,
   JSON.stringify(abierto.cerrado_en));
ok('y los días quietos se CALCULAN, no se copian', abierto.dias_quieto > 0,
   'dias_quieto = ' + abierto.dias_quieto);
igual('un caso sin pedido entra igual, no se pierde', '', abierto.pedido_id);

console.log('\n── Las tres hojas van a tres sitios distintos ──');
igual('pedidos', 'pedidos', globalThis.FUENTES.propio.tipo);
igual('novedades', 'novedades', globalThis.FUENTES.propio_novedades.tipo);
igual('cas', 'cas', globalThis.FUENTES.propio_cas.tipo);
ok('y las tres piden confirmación antes de escribir',
   ['propio','propio_novedades','propio_cas']
     .every(function (f) { return globalThis.FUENTES[f].propio === true; }));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
