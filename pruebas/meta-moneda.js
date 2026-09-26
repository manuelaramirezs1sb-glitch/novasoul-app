/**
 * EL EXPORT DE META, EN CUALQUIER MONEDA.
 *
 * ┌─ QUÉ PASÓ ─────────────────────────────────────────────────┐
 * │                                                            │
 * │ «la información de Meta no la está guardando, ¿por qué?»    │
 * │                                                            │
 * │ Porque Meta no llama igual a la misma columna según en qué  │
 * │ moneda le cobre a esa cuenta publicitaria:                  │
 * │                                                            │
 * │     Importe gastado (COP)   ← la cuenta de Ecuador          │
 * │     Importe gastado (GTQ)   ← la de Guatemala               │
 * │     Importe gastado (USD)   ← una cuenta en dólares         │
 * │                                                            │
 * │ Y los alias del código solo conocían `(cop)`, escrito a     │
 * │ mano. Con cualquier otra moneda la columna del GASTO no se  │
 * │ reconocía: las filas entraban con el importe vacío y el mes │
 * │ aparecía «sin pauta» DESPUÉS de haber subido el archivo     │
 * │ correcto. El peor tipo de fallo — el que parece que uno no  │
 * │ hizo el trabajo.                                            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { emp: {} };
function libro(id) { const h = LIBROS[id] || {}; return { getSheetByName: (n) => {
  const m = h[n]; if (!m) return null;
  return { getName: () => n, getLastRow: () => m.length,
    getLastColumn: () => (m[0] ? m[0].length : 0),
    getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
    appendRow: f => m.push(f.slice()),
    getRange: (f,c,nf,nc) => ({ getValues: () => { const o=[];
      for (let i=0;i<(nf||1);i++){const fl=m[f-1+i]||[];o.push(fl.slice(c-1,c-1+(nc||fl.length)));}
      return o; }, setValues: () => {}, setValue: () => {} }) };
} }; }
const PROPS = { ID_EMPRESARIAL:'emp' };
global.PropertiesService = { getScriptProperties: () => ({ getProperty: k => PROPS[k]||'',
  getProperties: () => PROPS, setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: id => libro(id), flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType:{CLOCK:'CLOCK'}, WeekDay:{MONDAY:'MONDAY'} };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({getEmail:()=>''}),
                   getEffectiveUser: () => ({getEmail:()=>''}) };
global.LockService = { getScriptLock: () => ({ tryLock:()=>true, releaseLock:()=>{} }) };
global.CacheService = { getScriptCache: () => ({ get:()=>null, put:()=>{}, remove:()=>{} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{}' }) };
global.Utilities = { sleep: () => {}, getUuid: () => 'u1', formatDate: (d,tz,pat) => {
  const iso = new Date(d).toISOString();
  if (pat==='yyyy-MM-dd') return iso.slice(0,10); if (pat==='yyyy-MM') return iso.slice(0,7);
  if (/HH:mm/.test(pat)) return iso.slice(0,19).replace('T',' '); return iso; } };

(0, eval)(src + '\n;globalThis.__F = { leerCrudo, FUENTES, libroOlvidar_, ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const E = F.ESQUEMA_EMPRESARIAL;
const f = (c,o) => c.map(k => (o[k]!==undefined?o[k]:''));

/** El export de Meta tal cual, con la moneda dentro del encabezado. */
function exportMeta(moneda) {
  return [
    ['Inicio del informe', 'Fin del informe', 'Nombre de la campaña',
     'Importe gastado (' + moneda + ')', 'Resultados',
     'Coste por compra (' + moneda + ')',
     'CPM (coste por 1000 impresiones) (' + moneda + ')'],
    ['2026-08-01', '2026-08-31', 'VIDEOS | TAG RECEDE', '244,50', '49', '5,00', '12,30'],
    ['2026-08-01', '2026-08-31', 'IMAGENES | TAG RECEDE', '188,00', '32', '5,90', '11,10'],
  ];
}

function montar(tabla) {
  F.libroOlvidar_();
  const emp = {};
  Object.keys(E).forEach(k => { emp[k] = [E[k]]; });
  emp.Tiendas.push(f(E.Tiendas, { id:'gt', nombre:'Nutrea GT', moneda:'GTQ',
                                  zona_horaria:'UTC', estado:'activa' }));
  emp._Import_Meta = tabla;
  LIBROS.emp = emp;
  // `leerCrudo` lee la pestaña de staging, igual que al subir el archivo.
  return F.leerCrudo(libro('emp'), 'meta', 'gt');
}

console.log('\n── 1 · EL CASO DE SU FOTO: una cuenta que NO cobra en pesos ──');
for (const moneda of ['GTQ', 'USD', 'MXN', 'COP']) {
  const r = montar(exportMeta(moneda));
  ok(moneda + ': la columna del gasto se reconoce',
     r.sinMapear.indexOf('gasto') === -1, JSON.stringify(r.sinMapear));
  ok(moneda + ': y el importe llega con su número',
     r.filas[0] && r.filas[0].gasto === 244.5,
     JSON.stringify(r.filas[0] && r.filas[0].gasto));
  ok(moneda + ': el CPA también', r.sinMapear.indexOf('cpa') === -1,
     JSON.stringify(r.sinMapear));
}

console.log('\n── 2 · lo que ya funcionaba, igual ──');
let r = montar(exportMeta('COP'));
igual('la campaña', 'VIDEOS | TAG RECEDE', r.filas[0].campana);
igual('la fecha de inicio', '2026-08-01', r.filas[0].fecha);
igual('y la de fin, que es la que hace el periodo', '2026-08-31', r.filas[0].fecha_fin);
igual('los resultados, ya como número', 49, r.filas[0].resultados);
igual('dos filas', 2, r.filas.length);
igual('con la tienda que se estaba mirando', 'gt', r.filas[0].tienda);

console.log('\n── 3 · dos columnas no se pelean la misma casilla ──');
/**
 * La segunda pasada solo mira columnas que nadie tomó en la primera.
 * Sin eso, «Importe gastado» e «Importe gastado (USD)» en el mismo
 * archivo podrían acabar los dos apuntando a la misma.
 */
r = montar([
  ['Inicio del informe', 'Importe gastado', 'Importe gastado (USD)', 'Resultados'],
  ['2026-08-01', '100', '999', '10'],
]);
igual('gana la coincidencia exacta', 100, r.filas[0].gasto);

console.log('\n── 4 · una columna que no existe se sigue diciendo ──');
r = montar([
  ['Inicio del informe', 'Nombre de la campaña', 'Resultados'],
  ['2026-08-01', 'C1', '10'],
]);
ok('el gasto aparece como no mapeado', r.sinMapear.indexOf('gasto') !== -1,
   JSON.stringify(r.sinMapear));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
