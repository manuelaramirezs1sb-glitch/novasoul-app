/**
 * «PERO SI YA LO SUBÍ» · dónde está lo que falta.
 *
 * ┌─ QUÉ PASÓ ─────────────────────────────────────────────────┐
 * │                                                            │
 * │ Ella, mirando Nutrea GT: «dice que le falta algo pero ya    │
 * │ subí toda la info». Agosto, julio y junio decían «sin       │
 * │ pauta · sin gastos fijos».                                  │
 * │                                                            │
 * │ Y era cierto para GT. El reporte estaba subido con Ecuador  │
 * │ abierto, así que las filas quedaron con `tienda = ec`.      │
 * │                                                            │
 * │ El histórico mira una tienda a la vez —y así debe ser— pero │
 * │ decir «sin pauta, súbela» sobre algo que YA está, en el     │
 * │ sitio de al lado, manda a hacer dos veces un trabajo y a    │
 * │ dudar de si Nova lee bien los archivos.                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Ahora Nova cuenta también lo de las otras tiendas del mismo mes y
 * dice dónde está. La diferencia entre un callejón sin salida y un
 * «cámbiate de tienda».
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



(0, eval)(src + '\n;globalThis.__F = { apiHistorial, libroOlvidar_, soulOlvidar_,' +
  ' ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const E = F.ESQUEMA_EMPRESARIAL;
const f = (c, o) => c.map(k => (o[k] !== undefined ? o[k] : ''));
const MES = new Date().toISOString().slice(0, 7);

function sembrar() {
  F.libroOlvidar_(); F.soulOlvidar_();
  const emp = {};
  Object.keys(E).forEach(k => { emp[k] = [E[k]]; });
  emp.Tiendas.push(f(E.Tiendas, { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD',
                                  zona_horaria: 'UTC', estado: 'activa' }));
  emp.Tiendas.push(f(E.Tiendas, { id: 'gt', nombre: 'Nutrea GT', moneda: 'GTQ',
                                  zona_horaria: 'UTC', estado: 'activa' }));
  // Pedidos en GT: el mes existe.
  for (let i = 0; i < 20; i++) {
    emp.Pedidos.push(f(E.Pedidos, { id: 'p' + i, tienda: 'gt', fecha: MES + '-10',
      producto: 'TAG', valor: 30, cantidad: 1, estado: 'ENTREGADO',
      estado_canonico: 'entregado', estado_nova: 'entregado', fecha_entrega: MES + '-12' }));
  }
  // Pero la pauta y los gastos se subieron con ECUADOR abierto.
  emp.Pauta.push(f(E.Pauta, { id: 'a1', fecha: MES + '-01', tienda: 'ec',
    plataforma: 'Meta', gasto: 500, moneda_gasto: 'USD', gasto_normalizado: 500 }));
  emp.Pauta.push(f(E.Pauta, { id: 'a2', fecha: MES + '-02', tienda: 'ec',
    plataforma: 'Meta', gasto: 300, moneda_gasto: 'USD', gasto_normalizado: 300 }));
  emp.Gastos.push(f(E.Gastos, { id: 'g1', tienda: 'ec', nombre: 'Shopify',
                                valor: 35, activo: 'si' }));
  LIBROS.emp = emp;
}

const S = { email: 'm@n.com', nombre: 'M', rol: 'dueno', sheetId: 'emp',
            tiendas: ['ec', 'gt'] };

console.log('\n── 1 · el caso de su foto ──');
sembrar();
let r = F.apiHistorial(S, { tienda: 'gt', meses: 3 });
ok('contesta', r.ok, JSON.stringify(r).slice(0, 120));
const m = r.meses.filter(x => x.mes === MES)[0];
ok('el mes tiene pedidos de GT', m.pedidos === 20, JSON.stringify(m).slice(0, 140));
igual('y en GT no hay pauta, que es cierto', 0, m.pautaFilas);
ok('pero dice que hay dos filas en Ecuador',
   m.pautaOtras && m.pautaOtras.ec === 2, JSON.stringify(m.pautaOtras));
ok('y que el gasto fijo también está allá',
   m.fijosOtras && m.fijosOtras.ec === 1, JSON.stringify(m.fijosOtras));

console.log('\n── 2 · mirando Ecuador, no sobra nada ──');
/**
 * El aviso es para lo que FALTA aquí y está al lado. Desde Ecuador la
 * pauta está donde debe, así que no hay nada que señalar: un aviso que
 * sale siempre no lo lee nadie.
 */
r = F.apiHistorial(S, { tienda: 'ec', meses: 3 });
const e = r.meses.filter(x => x.mes === MES)[0];
igual('la pauta se cuenta como suya', 2, e.pautaFilas);
igual('y el gasto fijo también', 1, e.nFijos);
ok('sin señalar a ninguna otra tienda', !e.pautaOtras, JSON.stringify(e.pautaOtras));

console.log('\n── 3 · cuando de verdad no está en ninguna parte ──');
F.libroOlvidar_();
LIBROS.emp.Pauta = [E.Pauta];
LIBROS.emp.Gastos = [E.Gastos];
r = F.apiHistorial(S, { tienda: 'gt', meses: 3 });
const v = r.meses.filter(x => x.mes === MES)[0];
igual('sigue sin pauta', 0, v.pautaFilas);
ok('y ahora NO se manda a buscar a otra tienda', !v.pautaOtras,
   JSON.stringify(v.pautaOtras));
ok('ni el gasto fijo', !v.fijosOtras, JSON.stringify(v.fijosOtras));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
