/**
 * El inventario: crear una ficha, y sobre todo EDITARLA.
 *
 * Ella lleva semanas diciendo lo mismo: «lo del inventario sigue sin
 * funcionar, sin poder editar lo del TAG RECEDE, no he podido montar la
 * información, ni el link de la landing page».
 *
 * Esta prueba manda exactamente lo que manda la pantalla —los mismos
 * campos, con los mismos nombres— contra la hoja tal como la tiene ella,
 * y mira qué se guarda y qué se rechaza. Si algo se cae, tiene que
 * caerse aquí y no en su tarde.
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

(0, eval)(src + '\n;globalThis.__F = { apiEscribir, apiCrear, apiBorrar, apiProductos,' +
  ' validarFicha, validarCanal, normalizarEnlace, ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** La hoja Inventario tal como la define el esquema de hoy. */
const C_INV = F.ESQUEMA_EMPRESARIAL.Inventario;

function fila(o) { return C_INV.map(c => (o[c] === undefined ? '' : o[c])); }

function sembrar(columnas) {
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

  const cols = columnas || C_INV;
  LIBROS.cli = {
    Inventario: [cols.slice(),
      cols.map(c => ({
        id: 'inv-tag', sku: 'TR-001', producto: 'TAG RECEDE', tienda: 'ec',
        fuente: '', origen: 'manual', categoria: 'estrella', proveedor: 'Laboratorio',
        landing: '', stock: 40, costo_unitario: 6, precio: 26,
        precio_2: '', precio_3: '', minimo: 10, activo: 'si',
      })[c] === undefined ? '' : ({
        id: 'inv-tag', sku: 'TR-001', producto: 'TAG RECEDE', tienda: 'ec',
        fuente: '', origen: 'manual', categoria: 'estrella', proveedor: 'Laboratorio',
        landing: '', stock: 40, costo_unitario: 6, precio: 26,
        precio_2: '', precio_3: '', minimo: 10, activo: 'si',
      })[c])],
    Pedidos: [['id', 'tienda', 'producto', 'sku', 'estado_canonico', 'fecha', 'valor', 'cantidad']],
    Parametros: [['tienda', 'clave', 'valor']],
    Tiendas: [['id', 'nombre', 'moneda', 'estado'], ['ec', 'Nutrea EC', 'USD', 'activa']],
    Movimientos: [['cuando', 'quien', 'entidad', 'entidad_id', 'campo', 'antes', 'ahora']],
  };
}

const DUENA = { sheetId: 'cli', rol: 'dueno', email: 'manuela@nova.com',
                nombre: 'Manuela', tiendas: ['ec'] };

/** Exactamente lo que arma `leerFicha()` en la pantalla. */
function comoLaPantalla(extra) {
  const d = {
    producto: 'TAG RECEDE', sku: 'TR-001', stock: '38', minimo: '10',
    costo_unitario: '6', precio: '26', precio_2: '46', precio_3: '63',
    proveedor: 'Laboratorio', landing: 'https://nutrea.co/tag-recede',
    nota: 'El que más sale', resp_1: 'Para las manchas', resp_2: 'Dos veces al día',
    resp_3: 'A las cuatro semanas', resp_4: 'Embarazo no',
    categoria: 'estrella',
  };
  Object.keys(extra || {}).forEach(k => { d[k] = extra[k]; });
  return d;
}

function valor(campo) {
  const d = LIBROS.cli.Inventario;
  const i = d[0].indexOf(campo);
  return i === -1 ? '(no existe la columna)' : d[1][i];
}

// ─────────────────────────────────────────────────────────────
console.log('\n── Editar el TAG RECEDE, con la hoja al día ──');
sembrar();
let r = F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
                               campos: comoLaPantalla() });
ok('se guarda', r.ok === true, JSON.stringify(r));
igual('sin rechazar nada', [], r.rechazados || []);
igual('el stock quedó', 38, Number(valor('stock')));
igual('los precios del combo también', [46, 63],
      [Number(valor('precio_2')), Number(valor('precio_3'))]);
igual('LA LANDING QUEDÓ', 'https://nutrea.co/tag-recede', valor('landing'));
igual('y las cuatro respuestas', 'Para las manchas', valor('resp_1'));
igual('tocar el stock a mano deja fechado el conteo', '2026-09-23', valor('ultimo_conteo'));

// ─────────────────────────────────────────────────────────────
console.log('\n── La landing, escrita como la escribe cualquiera ──');
[['nutrea.co/tag-recede', 'https://nutrea.co/tag-recede', 'sin https, se completa'],
 ['https://nutrea.co/x', 'https://nutrea.co/x', 'completa, se respeta'],
 ['', '', 'vacía, se acepta'],
].forEach(function (caso) {
  sembrar();
  const g = F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
                                   campos: { landing: caso[0] } });
  ok(caso[2], g.ok === true && valor('landing') === caso[1],
     JSON.stringify({ error: g.error, quedó: valor('landing') }));
});
sembrar();
let mal = F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
                                 campos: { landing: 'no es un enlace' } });
ok('una landing que no es un enlace se rechaza con explicación',
   mal.ok === false && /enlace/i.test(mal.error), JSON.stringify(mal));

// ─────────────────────────────────────────────────────────────
console.log('\n── Y AHORA: la hoja SIN migrar ──');
/**
 * Este es el caso de ella. La hoja Inventario de una cuenta vieja no
 * tiene las columnas nuevas. La pantalla filtra por `columnas`, pero si
 * el filtro falla o la pantalla es vieja, llega todo — y hay que ver qué
 * pasa entonces, porque es lo que ella está viviendo.
 */
const VIEJA = ['id', 'sku', 'producto', 'tienda', 'stock', 'costo_unitario',
               'precio', 'minimo', 'nota', 'activo', 'actualizado_en', 'actualizado_por'];
sembrar(VIEJA);
r = F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
                           campos: comoLaPantalla() });
ok('lo que sí cabe se guarda igual', r.ok === true && Number(valor('stock')) === 38,
   JSON.stringify(r));
ok('y dice CUÁLES no entraron', (r.rechazados || []).length > 0,
   JSON.stringify(r.rechazados));
ok('nombrando la columna y qué hacer',
   (r.rechazados || []).some(x => /landing/.test(x) && /bootstrapTodo/.test(x)),
   JSON.stringify(r.rechazados));

console.log('\n── Qué columnas dice el servidor que hay ──');
sembrar(VIEJA);
const prod = F.apiProductos(DUENA, { tienda: 'ec' });
ok('la respuesta trae la lista de columnas', Array.isArray(prod.columnas),
   JSON.stringify(Object.keys(prod)));
ok('y son las de la hoja de verdad, no las del esquema',
   prod.columnas.indexOf('landing') === -1 && prod.columnas.indexOf('stock') !== -1,
   JSON.stringify(prod.columnas));

// ─────────────────────────────────────────────────────────────
console.log('\n── Crear una ficha nueva ──');
sembrar();
const c = F.apiCrear(DUENA, { entidad: 'Inventario', tienda: 'ec',
                              datos: comoLaPantalla({ producto: 'DR MELAXIN',
                                                      sku: 'DM-1', tienda: 'ec' }) });
ok('se crea', c.ok === true, JSON.stringify(c));
const nueva = LIBROS.cli.Inventario[2] || [];
const cx = f => nueva[LIBROS.cli.Inventario[0].indexOf(f)];
igual('con su nombre', 'DR MELAXIN', cx('producto'));
igual('y su landing', 'https://nutrea.co/tag-recede', cx('landing'));
igual('nace como conteo manual', 'manual', cx('origen'));

console.log('\n── Lo que NO se debe poder ──');
sembrar();
ok('sin nombre no hay ficha',
   F.apiCrear(DUENA, { entidad: 'Inventario', tienda: 'ec',
                       datos: { stock: 5 } }).ok === false);
ok('un combo de 2 más barato que 1 se avisa',
   /menor que/.test(F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
     campos: { precio: '26', precio_2: '20' } }).error || ''));
ok('números negativos, no',
   /negativos/.test(F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
     campos: { stock: '-3' } }).error || ''));
ok('una categoría inventada, tampoco',
   /categoría/.test(F.apiEscribir(DUENA, { entidad: 'Inventario', id: 'inv-tag',
     campos: { categoria: 'lo que sea' } }).error || ''));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
