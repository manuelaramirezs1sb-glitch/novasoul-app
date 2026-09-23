/**
 * Cuando la dueña dice qué significa un estado.
 *
 * ── EL FALLO QUE ESTA PRUEBA EXISTE PARA ATRAPAR ──
 *
 * La pantalla decía «Listo · las cifras se están rehaciendo» y las
 * cifras no se rehacían. La traducción se guardaba en la hoja Estados
 * —el diccionario— pero los pedidos YA IMPORTADOS conservaban el
 * `estado_canonico` que se les escribió el día que entraron. Volvían a
 * contarse bien solo si ella reimportaba el archivo entero.
 *
 * Y el código afirmaba lo contrario en un comentario: «cada pantalla
 * recalcula desde Pedidos, y el estado se vuelve a traducir al leer».
 * La función que hacía esa retraducción —`estadoCanonico`— no la
 * llamaba nadie. Era código muerto que servía de coartada.
 *
 * Eso es lo peor que puede hacer un programa: decir que hizo algo,
 * no hacerlo, y tener un comentario que respalde la mentira.
 *
 * Lo que se comprueba aquí:
 *   · que clasificar reescriba los pedidos que ya estaban;
 *   · que NO pise el `estado_nova` de un pedido corregido a mano;
 *   · que solo toque los de esa fuente y ese texto;
 *   · que diga CUÁNTOS pedidos movió, para que el mensaje en pantalla
 *     pueda ser verdad;
 *   · que los meses ya cerrados se avisen y no se toquen en silencio.
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
          /**
           * Escribe DENTRO del rango, no la fila entera.
           *
           * El primer intento pisaba la fila completa con las celdas del
           * rango, y una escritura de una sola columna borraba el resto
           * del pedido. La prueba fallaba y el código estaba bien: un
           * stub que miente sobre la hoja inventa fallos y, peor, puede
           * esconder los de verdad.
           */
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
    insertSheet: (nombre) => { hojas[nombre] = []; return libro(id).getSheetByName(nombre); },
  };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-23T12:00:00Z';
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
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{"data":[]}' }) };
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

(0, eval)(src + '\n;globalThis.__F = { apiEstados, apiEstadoClasificar,' +
  ' estadosAprendidos, OPCIONES_ESTADO };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_EST = ['fuente','texto','estado_nova','origen','pedidos',
               'primera_vez','ultima_vez','decidido_por','nota'];
const C_PED = ['id','fuente','id_externo','fecha','tienda','cliente','estado',
               'estado_canonico','estado_nova','valor','actualizado_en','actualizado_por'];

function ped(o) {
  return C_PED.map(c => (o[c] === undefined ? '' : o[c]));
}

function sembrar(opts) {
  opts = opts || {};
  LIBROS.cli = {
    Estados: [C_EST,
      ['dropi', 'rechazado',   '',            'nuevo',    40, '2026-08-01', '2026-09-20', '', ''],
      ['dropi', 'entregado',   'entregado',   'catalogo', 90, '2026-08-01', '2026-09-20', '', ''],
      ['effi',  'rechazado',   '',            'nuevo',    12, '2026-08-01', '2026-09-20', '', '']],
    Pedidos: [C_PED].concat(opts.pedidos || [
      // Los 3 de dropi que ella va a clasificar
      ped({ id: 'dropi-1', fuente: 'dropi', tienda: 'ec', fecha: '2026-09-10',
            estado: 'RECHAZADO', estado_canonico: 'sin_clasificar', valor: 100 }),
      ped({ id: 'dropi-2', fuente: 'dropi', tienda: 'ec', fecha: '2026-09-11',
            estado: 'Rechazado', estado_canonico: 'sin_clasificar', valor: 100 }),
      // Uno que ella YA corrigió a mano: su decisión no se toca
      ped({ id: 'dropi-3', fuente: 'dropi', tienda: 'ec', fecha: '2026-09-12',
            estado: 'RECHAZADO', estado_canonico: 'sin_clasificar',
            estado_nova: 'entregado', valor: 100 }),
      // De otra fuente, mismo texto: no se debe tocar
      ped({ id: 'effi-9', fuente: 'effi', tienda: 'ec', fecha: '2026-09-12',
            estado: 'RECHAZADO', estado_canonico: 'sin_clasificar', valor: 100 }),
      // De otro estado, misma fuente: tampoco
      ped({ id: 'dropi-4', fuente: 'dropi', tienda: 'ec', fecha: '2026-09-13',
            estado: 'ENTREGADO', estado_canonico: 'entregado', valor: 100 }),
    ]),
    Cierres: [['tienda','mes','cerrado_en']].concat(opts.cierres || []),
    Movimientos: [['cuando','quien','entidad','entidad_id','campo','antes','ahora']],
  };
}

const DUENA = { sheetId: 'cli', rol: 'dueno', email: 'manuela@nova.com', tiendas: ['ec'] };
const GESTORA = { sheetId: 'cli', rol: 'gestora', email: 'kat@nova.com', tiendas: ['ec'] };

/** Cómo lee un pedido cualquier pantalla: el override manual y si no, lo canónico. */
function comoLoLee(id) {
  const d = LIBROS.cli.Pedidos;
  const e = d[0];
  const f = d.slice(1).filter(x => x[e.indexOf('id')] === id)[0];
  if (!f) return null;
  return String(f[e.indexOf('estado_nova')] || f[e.indexOf('estado_canonico')]).toLowerCase();
}

// ─────────────────────────────────────────────────────────────
console.log('\n── Antes de clasificar ──');
sembrar();
let r = F.apiEstados(DUENA, {});
igual('hay dos estados sin clasificar', ['rechazado', 'rechazado'],
      r.sinClasificar.map(x => x.texto));
igual('y los 40 pedidos de dropi pesan más, así que van primero',
      'dropi', r.sinClasificar[0].fuente);
igual('los pedidos no cuentan como nada', 'sin_clasificar', comoLoLee('dropi-1'));

// ─────────────────────────────────────────────────────────────
console.log('\n── Ella dice que RECHAZADO es una devolución ──');
const g = F.apiEstadoClasificar(DUENA, { fuente: 'dropi', texto: 'rechazado',
                                         estado: 'devolucion' });
ok('se guarda', g.ok === true, JSON.stringify(g));

/**
 * ── LA PRUEBA QUE IMPORTA ──
 * El diccionario cambió. Si los pedidos que ya estaban no cambian con
 * él, la pantalla dijo «las cifras se están rehaciendo» y mintió.
 */
igual('el pedido que ya estaba AHORA cuenta como devolución',
      'devolucion', comoLoLee('dropi-1'));
igual('y el que venía escrito distinto, también',
      'devolucion', comoLoLee('dropi-2'));
ok('se dice cuántos pedidos se movieron, para poder decirlo en pantalla',
   g.pedidosActualizados === 2, JSON.stringify(g.pedidosActualizados));

console.log('\n── Lo que NO se debe tocar ──');
/**
 * `estado_nova` es la corrección de ESE pedido, hecha a mano. El
 * diccionario es una regla general, y una regla general no pisa una
 * decisión concreta.
 */
igual('el pedido corregido a mano sigue como ella lo dejó',
      'entregado', comoLoLee('dropi-3'));
igual('el mismo texto en OTRA fuente no se toca',
      'sin_clasificar', comoLoLee('effi-9'));
igual('y otro estado de la misma fuente, tampoco',
      'entregado', comoLoLee('dropi-4'));

console.log('\n── Y queda dicho quién lo decidió ──');
const est = F.apiEstados(DUENA, {});
const rech = est.conocidos.filter(x => x.fuente === 'dropi' && x.texto === 'rechazado')[0];
ok('ya no está entre los sin clasificar',
   !est.sinClasificar.filter(x => x.fuente === 'dropi').length);
igual('quedó como devolución', 'devolucion', rech.estado);
igual('y marcado como decisión suya, no del catálogo', 'manual', rech.origen);
ok('el movimiento quedó registrado', LIBROS.cli.Movimientos.length === 2,
   JSON.stringify(LIBROS.cli.Movimientos));

// ─────────────────────────────────────────────────────────────
console.log('\n── Desclasificar vuelve a dejar la pregunta abierta ──');
sembrar();
F.apiEstadoClasificar(DUENA, { fuente: 'dropi', texto: 'rechazado', estado: 'devolucion' });
const v = F.apiEstadoClasificar(DUENA, { fuente: 'dropi', texto: 'rechazado', estado: '' });
ok('se puede deshacer', v.ok === true, JSON.stringify(v));
igual('el pedido vuelve a no contar como nada',
      'sin_clasificar', comoLoLee('dropi-1'));
igual('y el estado vuelve a la lista de preguntas', 1,
      F.apiEstados(DUENA, {}).sinClasificar.filter(x => x.fuente === 'dropi').length);

// ─────────────────────────────────────────────────────────────
console.log('\n── Un mes ya cerrado se avisa, no se esconde ──');
sembrar({ cierres: [['ec', '2026-09', '2026-09-30']] });
const c = F.apiEstadoClasificar(DUENA, { fuente: 'dropi', texto: 'rechazado',
                                         estado: 'devolucion' });
igual('dice qué mes quedó congelado antes de saber esto',
      ['2026-09'], (c.cerradosAfectados || []).map(x => x.mes));
/**
 * DOS, no tres. El tercero de ese mes tiene corrección manual, así que
 * su cifra no se movió — avisar por él sería mandarla a rehacer un
 * cierre por un pedido que quedó igual.
 */
igual('y cuántos pedidos de ese mes cambian DE VERDAD', 2,
      (c.cerradosAfectados || [])[0].pedidos);

// ─────────────────────────────────────────────────────────────
console.log('\n── Quién puede ──');
sembrar();
ok('una gestora no decide qué significa un estado',
   F.apiEstadoClasificar(GESTORA, { fuente: 'dropi', texto: 'rechazado',
                                    estado: 'devolucion' }).ok === false);
igual('y sus pedidos siguen intactos', 'sin_clasificar', comoLoLee('dropi-1'));
ok('pero sí puede VER la lista', F.apiEstados(GESTORA, {}).ok === true);
ok('solo que sin poder cambiarla', F.apiEstados(GESTORA, {}).puedeEditar === false);

console.log('\n── Un estado que no existe no se acepta ──');
sembrar();
ok('no se guarda cualquier cosa',
   F.apiEstadoClasificar(DUENA, { fuente: 'dropi', texto: 'rechazado',
                                  estado: 'inventado' }).ok === false);
igual('y el pedido no se movió', 'sin_clasificar', comoLoLee('dropi-1'));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
