/**
 * LEER UNA VEZ POR PETICIÓN, Y QUE ESCRIBIR LO TIRE.
 *
 * ┌─ QUÉ SE ESTÁ PROTEGIENDO ──────────────────────────────────┐
 * │                                                            │
 * │ «cuenta todos los llamados que necesita Nova Empresarial    │
 * │  para funcionar y no ser lenta, simplifícalo lo más         │
 * │  posible sin perder todo lo que hemos avanzado».            │
 * │                                                            │
 * │ El arranque quedó en UN viaje, pero por dentro ese viaje    │
 * │ leía 75 veces las pestañas (la de Pedidos, 12 veces). Con   │
 * │ el libro envuelto son 8. El ahorro es real y el riesgo      │
 * │ también, así que esta prueba no mide velocidad: mide que    │
 * │ NO se pueda leer algo viejo.                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ EL ÚNICO ERROR QUE IMPORTA AQUÍ ──────────────────────────┐
 * │                                                            │
 * │ Guardar, volver a leer, y que devuelva lo de antes. En      │
 * │ pantalla eso se ve como «no se guardó» — el peor síntoma    │
 * │ posible, porque invita a guardar otra vez encima.           │
 * │                                                            │
 * │ Por eso cada caso de abajo es: ESCRIBIR de verdad por la    │
 * │ API, y después LEER por la API, en la misma ejecución, sin  │
 * │ olvidar nada a mano.                                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
/** Cuántas veces se leyó de verdad cada pestaña. */
let LECTURAS_POR = {};
const LECTURAS_HOJA = () =>
  Object.keys(LECTURAS_POR).reduce((a, k) => a + LECTURAS_POR[k], 0);
function libro(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (nombre) => {
    const m = hojas[nombre];
    if (!m) return null;
    return {
      getName: () => nombre,
      getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => {
        LECTURAS_POR[nombre] = (LECTURAS_POR[nombre] || 0) + 1;
        return m.map(f => f.slice());
      } }),
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



(0, eval)(src + '\n;globalThis.__F = { libro_, libroOlvidar_, soulOlvidar_,' +
  ' apiNotaAgregar, apiNotas, apiEscribir, apiChatEnviar, apiChat, apiChatVisto,' +
  ' apiAccesosGuardar, apiAccesos, apiCrear, apiListar, apiArranque,' +
  ' ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const E = F.ESQUEMA_EMPRESARIAL;
function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

function sembrar() {
  UUID = 0; LECTURAS_POR = {};
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Tiendas: [E.Tiendas, f(E.Tiendas, { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD',
                                        zona_horaria: 'UTC' })],
    Pedidos: [E.Pedidos, f(E.Pedidos, { id: 'p1', tienda: 'ec', cliente: 'Alicia',
      valor: 30, estado_nova: 'novedad', nota: '' })],
    Novedades: [E.Novedades], Notas: [E.Notas], Mensajes: [E.Mensajes],
    Accesos: [E.Accesos], Movimientos: [E.Movimientos],
    Equipo: [E.Equipo, f(E.Equipo, { id: 'e1', nombre: 'Manuela', correo: 'm@n.com',
      rol: 'dueno', tienda: '*', estado: 'activo' }),
      f(E.Equipo, { id: 'e2', nombre: 'Zulay', correo: 'z@n.com',
        rol: 'gestora', tienda: 'ec', estado: 'activo' })],
    Inventario: [E.Inventario], Gastos: [E.Gastos], Pauta: [E.Pauta],
    CAS: [E.CAS], Llamadas: [E.Llamadas], Estados: [E.Estados],
    Cartera: [E.Cartera], Parametros: [E.Parametros], Cierres: [E.Cierres],
    Fuentes: [E.Fuentes], Tasas: [E.Tasas],
  };
}

const S = { email: 'm@n.com', nombre: 'Manuela', rol: 'dueno',
            sheetId: 'emp', tiendas: ['ec'] };

console.log('\n── 1 · anotar y volver a leer, en la misma ejecución ──');
sembrar();
// Se lee ANTES: así queda algo guardado en el caché que la escritura
// tiene que tirar. Sin esta primera lectura la prueba no probaría nada.
igual('al principio no hay notas', 0,
      Object.keys(F.apiNotas(S, { tienda: 'ec' }).notas).length);
let r = F.apiNotaAgregar(S, { entidad: 'Pedidos', id: 'p1', texto: 'la llamé' });
ok('se anota', r.ok, JSON.stringify(r));
let v = F.apiNotas(S, { tienda: 'ec' });
igual('y al leer OTRA VEZ aparece', 1, (v.notas['Pedidos:p1'] || []).length);
igual('con su texto', 'la llamé', v.notas['Pedidos:p1'][0].texto);

console.log('\n── 2 · la fila que se tocó al anotar también se relee ──');
let lista = F.apiListar(S, { entidad: 'Pedidos', tienda: 'ec' });
igual('la columna nota del pedido está al día', 'la llamé',
      (lista.filas[0] || {}).nota);
igual('y quien lo trabajó, también', 'Manuela', (lista.filas[0] || {}).gestionado_por);

console.log('\n── 3 · escribir un campo y releerlo ──');
r = F.apiEscribir(S, { entidad: 'Pedidos', id: 'p1',
                       campos: { estado_nova: 'entregado' } });
ok('se escribe', r.ok, JSON.stringify(r));
lista = F.apiListar(S, { entidad: 'Pedidos', tienda: 'ec' });
igual('y se lee lo nuevo, no lo de antes', 'entregado',
      (lista.filas[0] || {}).estado_nova);

console.log('\n── 4 · el chat: mandar y ver ──');
sembrar();
igual('el chat arranca vacío', 0, F.apiChat(S, { con: 'g:ec' }).mensajes.length);
r = F.apiChatEnviar(S, { con: 'g:ec', texto: 'buenos días' });
ok('se manda', r.ok, JSON.stringify(r));
igual('y se ve enseguida', 1, F.apiChat(S, { con: 'g:ec' }).mensajes.length);

// `apiChatVisto` escribe la columna entera con setValues: es el caso que
// más fácil se rompería si el caché no se tirara al escribir.
const Z = { email: 'z@n.com', nombre: 'Zulay', rol: 'gestora',
            sheetId: 'emp', tiendas: ['ec'] };
igual('a la gestora le cuenta 1 sin leer', 1, F.apiChat(Z, {}).sinLeer);
r = F.apiChatVisto(Z, { con: 'g:ec' });
igual('se marca', 1, r.marcados);
igual('y al releer ya no cuenta', 0, F.apiChat(Z, {}).sinLeer);

console.log('\n── 5 · los accesos: guardar y releer ──');
sembrar();
igual('arrancan vacíos', '', F.apiAccesos(S, { tienda: 'ec' }).accesos.plataforma);
r = F.apiAccesosGuardar(S, { tienda: 'ec', cambios: { plataforma: 'Dropi' } });
ok('se guardan', r.ok, JSON.stringify(r));
igual('y se releen', 'Dropi', F.apiAccesos(S, { tienda: 'ec' }).accesos.plataforma);

console.log('\n── 6 · crear una fila nueva y verla en la lista ──');
// Un gasto y no un pedido: los pedidos no se crean a mano, entran por
// importación. (Mi primera versión de esta prueba lo intentaba y el
// servidor la paró, que es exactamente lo que tiene que hacer.)
sembrar();
igual('no hay gastos', 0, F.apiListar(S, { entidad: 'Gastos', tienda: 'ec' }).filas.length);
r = F.apiCrear(S, { entidad: 'Gastos', datos: { tienda: 'ec', nombre: 'Shopify',
                                                valor: 35, tipo: 'fijo' } });
ok('se crea', r.ok, JSON.stringify(r));
igual('y aparece enseguida', 1, F.apiListar(S, { entidad: 'Gastos', tienda: 'ec' }).filas.length);
igual('con su nombre', 'Shopify',
      (F.apiListar(S, { entidad: 'Gastos', tienda: 'ec' }).filas[0] || {}).nombre);

console.log('\n── 7 · nadie se queda con una copia viva de la hoja ──');
/**
 * `apiChatVisto` escribe sobre el arreglo que le devolvió `getValues()`
 * antes de mandarlo de vuelta. Si el caché devolviera SU arreglo en vez
 * de una copia, esa escritura en memoria se le aparecería al siguiente
 * que leyera la misma hoja en la misma petición — un dato cambiado que
 * nunca llegó a Google.
 */
sembrar();
const sh = F.libro_('emp').getSheetByName('Pedidos');
const a = sh.getDataRange().getValues();
a[1][0] = 'ROTO';
const b = sh.getDataRange().getValues();
igual('tocar lo leído no ensucia la siguiente lectura', 'p1', b[1][0]);
ok('y tampoco la hoja de verdad', LIBROS.emp.Pedidos[1][0] === 'p1',
   String(LIBROS.emp.Pedidos[1][0]));

console.log('\n── 8 · una hoja que todavía no existe ──');
/**
 * No se guarda el «no existe»: bootstrapTodo crea pestañas en mitad de
 * una ejecución, y recordar que faltaban las dejaría invisibles hasta la
 * siguiente petición.
 */
sembrar();
const libro1 = F.libro_('emp');
igual('al principio no está', null, libro1.getSheetByName('Recien'));
LIBROS.emp.Recien = [['a', 'b'], [1, 2]];
ok('y en cuanto existe, se ve', libro1.getSheetByName('Recien') !== null);
igual('con su contenido', [['a', 'b'], [1, 2]],
      libro1.getSheetByName('Recien').getDataRange().getValues());

console.log('\n── 9 · y con todo eso, el arranque lee poco ──');
sembrar();
const antes = LECTURAS_HOJA();
F.apiArranque(S, { tienda: 'ec', mes: '2026-09', limite: 300 });
const leidas = LECTURAS_HOJA() - antes;
ok('el arranque entero lee menos de 15 veces (eran 75)', leidas < 15,
   String(leidas) + ' lecturas');
ok('y la hoja de Pedidos, una sola vez', LECTURAS_POR.Pedidos === 1,
   JSON.stringify(LECTURAS_POR));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
