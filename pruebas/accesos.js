/**
 * LOS ACCESOS DE LA TIENDA · y la contraseña que NO va a la bitácora.
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «los links de acceso, colocarlos en los lugares donde van   │
 * │  automáticamente». Y cuáles son: «ahí está el link de la    │
 * │  plataforma, la contraseña y correo, está el link de chat    │
 * │  center o wpp: el canal que use el cliente para             │
 * │  comunicarse con los compradores de su tienda».             │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LA ASERCIÓN QUE SOSTIENE ESTE ARCHIVO ────────────────────┐
 * │                                                            │
 * │ Levanté la mano por guardar la contraseña y ella lo         │
 * │ resolvió con hechos: es una subcuenta que solo confirma y   │
 * │ gestiona, no saca dinero, y pide un código que llega a un   │
 * │ correo que la gestora no tiene. Entonces se guarda.         │
 * │                                                            │
 * │ Lo que prometí a cambio son tres cosas, y las tres se       │
 * │ miden aquí:                                                │
 * │                                                            │
 * │  1. NO ENTRA EN MOVIMIENTOS. `apiParametros` registra cada │
 * │     cambio con el valor viejo y el nuevo; si los accesos   │
 * │     fueran un parámetro más, la contraseña quedaría en     │
 * │     texto plano en la bitácora para siempre. Esa es la      │
 * │     razón entera de que exista 97-accesos.gs.              │
 * │  2. SOLO LA VE QUIEN TIENE ESA TIENDA.                     │
 * │  3. SOLO LA DUEÑA LA CAMBIA.                               │
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

(0, eval)(src + '\n;globalThis.__F = { apiAccesos, apiAccesosGuardar, accesosDe_,' +
  ' libroOlvidar_, soulOlvidar_, ESQUEMA_EMPRESARIAL, ACCESOS_CAMPOS };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_ACC = F.ESQUEMA_EMPRESARIAL.Accesos;
const C_TIE = F.ESQUEMA_EMPRESARIAL.Tiendas;
const C_MOV = F.ESQUEMA_EMPRESARIAL.Movimientos;
function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

function sembrar() {
  UUID = 0;
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.emp = {
    Tiendas: [C_TIE, f(C_TIE, { id: 'ec', nombre: 'Nutrea Ecuador' }),
                     f(C_TIE, { id: 'gt', nombre: 'Nutrea Guatemala' })],
    Accesos: [C_ACC],
    Movimientos: [C_MOV],
  };
}

const S = {
  duena:   { email: 'sara@x.com',  nombre: 'Sara',  rol: 'dueno',   sheetId: 'emp', tiendas: ['ec', 'gt'] },
  admin:   { email: 'kat@x.com',   nombre: 'Kat',   rol: 'admin',   sheetId: 'emp', tiendas: ['ec', 'gt'] },
  zulay:   { email: 'zulay@x.com', nombre: 'Zulay', rol: 'gestora', sheetId: 'emp', tiendas: ['ec'] },
  jaime:   { email: 'jaime@x.com', nombre: 'Jaime', rol: 'gestora', sheetId: 'emp', tiendas: ['gt'] },
};

/** Los accesos de verdad del archivo que ella subió, sin la contraseña real. */
const REALES = {
  plataforma: 'Dropi',
  url: 'app.dropi.co/dashboard/home',        // sin esquema, como se copia
  usuario: 'gestion.vip@correodealejandro.com',
  clave: 'LaQueSeaAqui2026',
  correo_codigo: 'alejandro.escobar@correodealejandro.com',
  canal_nombre: 'Chat Center',
  canal_url: 'https://chateapro.app/login',
  nota: 'La subcuenta solo confirma y gestiona novedades. No saca dinero.',
};

console.log('\n── 1 · vacío es vacío, y lo dice ──');
sembrar();
let r = F.apiAccesos(S.duena, { tienda: 'ec' });
ok('la pantalla abre sin accesos puestos', r.ok, JSON.stringify(r));
igual('y dice que no hay nada', false, r.hay);
igual('con el nombre de la tienda para el mensaje', 'Nutrea Ecuador', r.nombreTienda);
igual('la dueña puede editar', true, r.puedeEditar);
igual('todos los campos en blanco',
      F.ACCESOS_CAMPOS.map(() => ''), F.ACCESOS_CAMPOS.map(k => r.accesos[k]));

console.log('\n── 2 · guardar, y que el enlace quede usable ──');
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec', cambios: REALES });
ok('la dueña los guarda', r.ok, JSON.stringify(r));
igual('una sola fila en la hoja', 2, LIBROS.emp.Accesos.length);
// Un dominio sin https:// pasa la validación pero no sirve en un href:
// el navegador lo lee como ruta relativa y el botón no lleva a ninguna parte.
igual('el enlace se completa con https', 'https://app.dropi.co/dashboard/home',
      r.accesos.url);
igual('el del chat se respeta tal cual', 'https://chateapro.app/login', r.accesos.canal_url);
igual('la plataforma queda con su nombre', 'Dropi', r.accesos.plataforma);
igual('y queda dicho quién y cuándo', 'sara@x.com', r.accesos.actualizado_por);
ok('con fecha', /^2026-09-25/.test(r.accesos.actualizado_en), r.accesos.actualizado_en);

console.log('\n── 2b · por dónde le paga la compradora ──');
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec', cambios: {
  pago_nombre: 'PayPal', pago_url: 'paypal.me/nutrea',
  pago_datos: 'Bancolombia ahorros 123-456789-00 a nombre de Nutrea SAS' } });
ok('se guardan', r.ok, JSON.stringify(r));
igual('el enlace de pago también se completa', 'https://paypal.me/nutrea',
      r.accesos.pago_url);
igual('y los datos que se dictan por teléfono quedan tal cual',
      'Bancolombia ahorros 123-456789-00 a nombre de Nutrea SAS', r.accesos.pago_datos);
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec',
  cambios: { pago_url: 'javascript:alert(1)' } });
ok('un enlace de pago peligroso se rechaza igual', r.ok === false, JSON.stringify(r));
igual('y el bueno sigue', 'https://paypal.me/nutrea',
      F.apiAccesos(S.duena, { tienda: 'ec' }).accesos.pago_url);

console.log('\n── 3 · LA CONTRASEÑA NO ENTRA EN LA BITÁCORA ──');
const mov = LIBROS.emp.Movimientos.slice(1);
const texto = JSON.stringify(mov);
ok('no aparece la contraseña en ninguna celda',
   texto.indexOf(REALES.clave) === -1, texto);
ok('pero SÍ queda registrado que cambió',
   mov.some(x => String(x[4]) === 'clave'), texto);
const fClave = mov.filter(x => String(x[4]) === 'clave')[0];
igual('diciendo que antes estaba vacía', '(vacía)', fClave && fClave[5]);
igual('y que ahora hay una', '(cambiada)', fClave && fClave[6]);
// El enlace SÍ se registra: un enlace equivocado hay que poder rastrearlo,
// y un enlace no es una llave.
ok('el enlace sí queda anotado',
   mov.some(x => String(x[4]) === 'url' && /dropi/.test(String(x[6]))), texto);

console.log('\n── 4 · quién los ve ──');
r = F.apiAccesos(S.zulay, { tienda: 'ec' });
ok('la gestora de la tienda los ve', r.ok, JSON.stringify(r));
igual('incluida la contraseña, que es para eso', REALES.clave, r.accesos.clave);
igual('pero no los puede cambiar', false, r.puedeEditar);
igual('y ve el correo del código', REALES.correo_codigo, r.accesos.correo_codigo);

r = F.apiAccesos(S.jaime, { tienda: 'ec' });
ok('la gestora de OTRA tienda no los ve', r.ok === false, JSON.stringify(r));
ok('y se le dice por qué', /acceso a esa tienda/i.test(r.error || ''), r.error);

r = F.apiAccesos(S.admin, { tienda: 'ec' });
ok('la admin los ve', r.ok, JSON.stringify(r));
igual('pero tampoco los cambia', false, r.puedeEditar);

console.log('\n── 5 · quién los cambia ──');
r = F.apiAccesosGuardar(S.zulay, { tienda: 'ec', cambios: { clave: 'otra' } });
ok('la gestora no los cambia', r.ok === false, JSON.stringify(r));
igual('y la contraseña sigue siendo la de antes', REALES.clave,
      F.apiAccesos(S.duena, { tienda: 'ec' }).accesos.clave);

r = F.apiAccesosGuardar(S.admin, { tienda: 'ec', cambios: { clave: 'otra' } });
ok('la admin tampoco', r.ok === false, JSON.stringify(r));

r = F.apiAccesosGuardar(S.duena, { tienda: 'xx', cambios: { clave: 'otra' } });
ok('nadie guarda en una tienda que no es suya', r.ok === false, JSON.stringify(r));

console.log('\n── 6 · un enlace peligroso no se guarda, ni a medias ──');
const antesUrl = F.apiAccesos(S.duena, { tienda: 'ec' }).accesos.url;
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec',
  cambios: { plataforma: 'Otra', url: 'javascript:fetch("/roba?t="+TOKEN)' } });
ok('un javascript: se rechaza', r.ok === false, JSON.stringify(r));
const ahora = F.apiAccesos(S.duena, { tienda: 'ec' }).accesos;
igual('el enlace bueno sigue ahí', antesUrl, ahora.url);
igual('y NO se guardó el nombre nuevo tampoco', 'Dropi', ahora.plataforma);

console.log('\n── 7 · guardar de nuevo actualiza, no duplica ──');
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec', cambios: { nota: 'Sara la revisa los lunes' } });
ok('se guarda el cambio', r.ok, JSON.stringify(r));
igual('sigue habiendo una sola fila', 2, LIBROS.emp.Accesos.length);
igual('la nota nueva', 'Sara la revisa los lunes', r.accesos.nota);
// Lo que no se manda no se pisa: un formulario que manda solo un campo no
// puede borrar los otros siete.
igual('y lo que no se mandó queda intacto', REALES.clave, r.accesos.clave);
igual('la plataforma también', 'Dropi', r.accesos.plataforma);

console.log('\n── 8 · cada tienda tiene las suyas ──');
r = F.apiAccesosGuardar(S.duena, { tienda: 'gt',
  cambios: { plataforma: 'Dropi GT', url: 'https://app.dropi.gt/home', clave: 'otraClave' } });
ok('la segunda tienda se guarda aparte', r.ok, JSON.stringify(r));
igual('dos filas, una por tienda', 3, LIBROS.emp.Accesos.length);
igual('Ecuador no cambió', 'Dropi', F.apiAccesos(S.duena, { tienda: 'ec' }).accesos.plataforma);
igual('Guatemala tiene lo suyo', 'Dropi GT',
      F.apiAccesos(S.duena, { tienda: 'gt' }).accesos.plataforma);
r = F.apiAccesos(S.jaime, { tienda: 'gt' });
ok('la gestora de Guatemala ve las suyas', r.ok && r.accesos.clave === 'otraClave',
   JSON.stringify(r));
ok('y sigue sin ver las de Ecuador',
   F.apiAccesos(S.jaime, { tienda: 'ec' }).ok === false);

console.log('\n── 9 · sin la hoja todavía ──');
delete LIBROS.emp.Accesos;
F.libroOlvidar_();
r = F.apiAccesos(S.duena, { tienda: 'ec' });
ok('la pantalla abre igual', r.ok === true, JSON.stringify(r));
igual('sin inventar nada', false, r.hay);
r = F.apiAccesosGuardar(S.duena, { tienda: 'ec', cambios: { clave: 'x' } });
ok('y guardar dice qué falta', r.ok === false && /bootstrapTodo/.test(r.error || ''),
   JSON.stringify(r));

console.log('\n── 10 · la hoja está declarada ──');
ok('Accesos está en el esquema', !!C_ACC, JSON.stringify(C_ACC));
['tienda','plataforma','url','usuario','clave','correo_codigo',
 'canal_nombre','canal_url','pago_nombre','pago_url','pago_datos',
 'nota','actualizado_en','actualizado_por']
  .forEach(c => ok('columna ' + c, (C_ACC || []).indexOf(c) !== -1));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
