/**
 * UNA NOVEDAD ES DE UNA SOLA TIENDA.
 *
 * ┌─ LO QUE PASÓ ──────────────────────────────────────────────┐
 * │                                                            │
 * │ `Novedades` era la única hoja con filas de una tienda y     │
 * │ sin columna para decir de cuál. La tienda se deducía del    │
 * │ pedido… en teoría. En la práctica, en ninguno de los        │
 * │ cuatro sitios que la leen.                                  │
 * │                                                            │
 * │ Y `apiListar` filtra por tienda SOLO si la hoja tiene la    │
 * │ columna. Sin ella no filtraba nada:                         │
 * │                                                            │
 * │   · la bandeja de novedades de Ecuador traía las de         │
 * │     Guatemala,                                              │
 * │   · el KPI «NOVEDADES» del mes sumaba las dos —en Hoy y en  │
 * │     el cierre del mes,                                      │
 * │   · la alarma «novedades sin gestionar» contaba las dos,    │
 * │   · «qué quedó de ayer» decía lo mismo en las dos tiendas,  │
 * │   · y a una gestora de Ecuador se le contaban las de        │
 * │     Guatemala, que ni puede ver.                            │
 * │                                                            │
 * │ «creo que está combinando datos». Aquí sí era el servidor.  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ LA PRUEBA MONTA DOS TIENDAS DE VERDAD ────────────┐
 * │                                                            │
 * │ Porque el error no era que un filtro estuviera mal escrito: │
 * │ era que NO ESTABA. Una prueba de una sola tienda pasa igual │
 * │ con filtro y sin filtro, que es exactamente cómo esto vivió │
 * │ meses sin que nada avisara.                                 │
 * │                                                            │
 * │ Así que hay dos tiendas, cada una con su novedad, y se      │
 * │ pregunta por cada una a las cinco puertas.                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

// ── Los servicios de Google, remedados ──
const HOJAS = {};
const ESCRITO = [];
function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  return {
    getName: () => nombre,
    getLastRow: () => m.length,
    getLastColumn: () => (m[0] || []).length,
    getDataRange: () => ({ getValues: () => m }),
    getRange: (fila, col, nf, nc) => ({
      getValues: () => m.slice(fila - 1, fila - 1 + (nf || 1))
        .map((f) => f.slice(col - 1, col - 1 + (nc || 1))),
      setValues: (v) => {
        ESCRITO.push({ hoja: nombre, fila: fila, col: col, filas: v.length });
        v.forEach((f, i) => {
          const r = m[fila - 1 + i] || (m[fila - 1 + i] = []);
          f.forEach((c, j) => { r[col - 1 + j] = c; });
        });
      },
      setValue: () => {},
    }),
  };
}
const SS = { getSheetByName: hoja, getId: () => 'emp', insertSheet: () => { throw new Error('no'); } };
global.SpreadsheetApp = { openById: () => SS, flush: () => {} };
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 'soul', ID_ACADEMY: 'aca' };
global.PropertiesService = {
  getScriptProperties: () => ({ getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS }),
};
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [] };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }) };
global.UrlFetchApp = {};
global.MailApp = { sendEmail: () => {} };
global.Utilities = {
  sleep: () => {},
  formatDate: (d, tz, patron) => {
    const iso = new Date(d).toISOString();
    if (patron === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (/HH:mm/.test(patron)) return iso.slice(0, 19);
    return iso;
  },
  getUuid: () => 'uuid-' + Math.random().toString(36).slice(2),
};
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };

(0, eval)(src + '\n;globalThis.__F = { apiListar, agregarMes, reporteDelDia, ' +
  'evaluarAlarmas, apiEquipo, derivarNovedades, rellenarTiendaNovedades_, libroOlvidar_, ' +
  'ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const hoy = new Date().toISOString().slice(0, 10);
const ayer = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
/**
 * EL MES ES EL DE LA NOVEDAD, NO EL DE HOY.
 *
 * Esta prueba se puso roja sola el 1 de octubre: `ayer` era el 30 de
 * septiembre y `agregarMes` se le pedía a octubre, así que la novedad
 * quedaba fuera del mes y el KPI contaba 0. El producto estaba bien;
 * la prueba asumía que ayer y hoy son del mismo mes, y una vez al mes
 * eso es falso.
 */
const mes = ayer.slice(0, 7);

/** Una fila con encabezados de verdad: los del esquema, no los que yo recuerde. */
function fila(entidad, valores) {
  const cols = F.ESQUEMA_EMPRESARIAL[entidad];
  return cols.map((c) => (valores[c] === undefined ? '' : valores[c]));
}
function montar(conColumnaTienda) {
  Object.keys(HOJAS).forEach((k) => delete HOJAS[k]);
  ESCRITO.length = 0;

  HOJAS.Parametros = [F.ESQUEMA_EMPRESARIAL.Parametros,
    fila('Parametros', { clave: 'moneda_reporte', valor: 'COP' }),
    fila('Parametros', { tienda: 'ec', clave: 'horas_novedad', valor: 1 }),
    fila('Parametros', { tienda: 'gt', clave: 'horas_novedad', valor: 1 })];

  HOJAS.Tiendas = [F.ESQUEMA_EMPRESARIAL.Tiendas,
    fila('Tiendas', { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD',
                      zona_horaria: 'UTC', estado: 'activa' }),
    fila('Tiendas', { id: 'gt', nombre: 'Nutrea GT', moneda: 'GTQ',
                      zona_horaria: 'UTC', estado: 'activa' })];

  HOJAS.Pedidos = [F.ESQUEMA_EMPRESARIAL.Pedidos,
    fila('Pedidos', { id: 'p-ec', tienda: 'ec', fecha: ayer, cliente: 'CLIENTA EC',
                      producto: 'PROD EC', valor: 50, estado_nova: 'novedad',
                      gestora_asignada: 'Sara', ultimo_movimiento: ayer }),
    fila('Pedidos', { id: 'p-gt', tienda: 'gt', fecha: ayer, cliente: 'CLIENTA GT',
                      producto: 'PROD GT', valor: 70, estado_nova: 'novedad',
                      gestora_asignada: 'Sara', ultimo_movimiento: ayer })];

  // La novedad de cada tienda. `conColumnaTienda` decide si la fila lo dice.
  const colsN = F.ESQUEMA_EMPRESARIAL.Novedades
    .filter((c) => conColumnaTienda || c !== 'tienda');
  const filaN = (v) => colsN.map((c) => (v[c] === undefined ? '' : v[c]));
  HOJAS.Novedades = [colsN,
    filaN({ id: 'n-ec', pedido_id: 'p-ec', tienda: 'ec', fecha: ayer,
            motivo: 'DIRECCION ERRADA EC', grupo: 'direccion', estado: 'abierta',
            gestora: 'Sara' }),
    filaN({ id: 'n-gt', pedido_id: 'p-gt', tienda: 'gt', fecha: ayer,
            motivo: 'RECHAZO EN PUERTA GT', grupo: 'rechazo', estado: 'abierta',
            gestora: 'Sara' })];

  HOJAS.Equipo = [F.ESQUEMA_EMPRESARIAL.Equipo,
    fila('Equipo', { id: 'e1', nombre: 'Sara', correo: 's@x.com', rol: 'gestora',
                     tienda: 'ec', estado: 'activo' })];

  ['Pauta', 'Gastos', 'Cartera', 'Facturacion', 'CAS', 'Inventario', 'Auditorias',
   'Cierres', 'Fuentes', 'Estados', 'Tasas', 'Llamadas', 'Anuncios',
   'Mensajes', 'Accesos', 'Notas'].forEach((t) => {
    HOJAS[t] = [F.ESQUEMA_EMPRESARIAL[t]];
  });
  HOJAS.Movimientos = [['id', 'fecha_hora', 'quien', 'entidad', 'tienda', 'fila_id',
                        'campo', 'antes', 'despues']];
  F.libroOlvidar_();
}

const dueña = { sheetId: 'emp', rol: 'dueno', nombre: 'Manuela', correo: 'm@n.com',
                tiendas: ['ec', 'gt'], permisos: [] };

console.log('\n══ 1 · CON LA COLUMNA, CADA PUERTA VE SOLO SU TIENDA ══');
montar(true);

const lista = (t) => F.apiListar(dueña, { entidad: 'Novedades', tienda: t, limite: 50 });
const idsDe = (r) => (r.filas || []).map((f) => f.id).sort();
igual('la bandeja de Ecuador trae solo la de Ecuador', ['n-ec'], idsDe(lista('ec')));
igual('y la de Guatemala solo la de Guatemala', ['n-gt'], idsDe(lista('gt')));

montar(true);
igual('el KPI del mes cuenta 1 en Ecuador', 1,
      F.agregarMes(SS, 'ec', mes, dueña).novedades);
montar(true);
igual('y 1 en Guatemala', 1, F.agregarMes(SS, 'gt', mes, dueña).novedades);

montar(true);
let r = F.reporteDelDia('emp', 'ec', ayer);
igual('«qué quedó de ayer» en Ecuador ve 1 novedad abierta', 1, r.novedades.abiertas);
igual('y su motivo es el de Ecuador', ['direccion'], r.motivos.map((m) => m.grupo));
montar(true);
r = F.reporteDelDia('emp', 'gt', ayer);
igual('en Guatemala ve 1 y es la suya', ['rechazo'], r.motivos.map((m) => m.grupo));

montar(true);
const alarmaNov = (t) => ((F.evaluarAlarmas(SS, t) || {}).alarmas || [])
  .filter((a) => a.id === 'novedad_vieja')
  .reduce((n, a) => n + (a.casos || []).length, 0);
igual('la alarma de novedades viejas en Ecuador cuenta 1', 1, alarmaNov('ec'));
montar(true);
igual('y en Guatemala también 1 (la suya)', 1, alarmaNov('gt'));

montar(true);
const eq = F.apiEquipo(dueña, { tienda: 'ec', mes: mes });
igual('a la gestora de Ecuador se le cuenta 1 novedad, no 2', [1],
      (eq.personas || []).map((x) => x.novedades));

console.log('\n══ 2 · SIN LA COLUMNA TODAVÍA, NADA SE ROMPE ══');
/**
 * Entre pegar el código y correr `bootstrapTodo()` la columna no
 * existe. Un filtro que la exigiera dejaría la bandeja vacía en las dos
 * tiendas, que es el error contrario y molesta igual.
 */
montar(false);
igual('la bandeja no se queda vacía', 2, (lista('ec').filas || []).length);
montar(false);
ok('el reporte del día no revienta',
   F.reporteDelDia('emp', 'ec', ayer).ok === true);

console.log('\n══ 3 · EL RELLENO SACA LA TIENDA DEL PEDIDO ══');
/**
 * La columna nueva nace vacía. Si el relleno no funciona, el filtro deja
 * cada novedad vieja fuera de TODAS las tiendas.
 */
montar(true);
// Se les borra la tienda, como quedan al estrenar la columna
const cT = HOJAS.Novedades[0].indexOf('tienda');
HOJAS.Novedades[1][cT] = '';
HOJAS.Novedades[2][cT] = '';
// Y una huérfana: su pedido no existe
HOJAS.Novedades.push(HOJAS.Novedades[0].map((c) =>
  c === 'id' ? 'n-huerfana' : c === 'pedido_id' ? 'p-que-no-existe'
  : c === 'estado' ? 'abierta' : c === 'fecha' ? ayer : ''));
F.libroOlvidar_();

const dicho = F.rellenarTiendaNovedades_('emp');
ok('dice qué hizo y cuántas quedaron sin dueño',
   /tienda puesta en 2/.test(dicho) && /1 sin pedido/.test(dicho), dicho);
igual('la de Ecuador quedó en ec', 'ec', HOJAS.Novedades[1][cT]);
igual('la de Guatemala quedó en gt', 'gt', HOJAS.Novedades[2][cT]);
igual('la huérfana se queda sin tienda, no se le inventa una', '', HOJAS.Novedades[3][cT]);
/**
 * Cuatro mil novedades a una llamada por celda se pasan del tiempo que
 * Google da. Tiene que ser UNA escritura de columna.
 */
const escrituras = ESCRITO.filter((x) => x.hoja === 'Novedades');
igual('escribe la columna de una sola vez', 1, escrituras.length);

/**
 * Y es idempotente EN LA ESCRITURA, que es lo que cuesta.
 *
 * La huérfana no se puede resolver nunca, así que el contador de
 * «faltan» no baja a cero. La primera versión escribía igual, y cada
 * `bootstrapTodo()` reescribía cuatro mil filas para no cambiar una
 * celda. Lo que se afirma es que no VUELVE A ESCRIBIR, no lo que dice.
 */
ESCRITO.length = 0;
F.libroOlvidar_();
F.rellenarTiendaNovedades_('emp');
igual('correrlo dos veces no vuelve a escribir', [],
      ESCRITO.filter((x) => x.hoja === 'Novedades'));

console.log('\n══ 4 · EL IMPORTADOR YA ESCRIBE LA TIENDA ══');
/**
 * `derivarNovedades` recibía la tienda como parámetro desde el primer
 * día y no la escribía. Aquí se afirma que sí.
 */
const derivadas = F.derivarNovedades([
  { id: 'p1', id_externo: 'x1', fecha: ayer, motivo_novedad: 'DIRECCION ERRADA',
    estado_canonico: 'novedad' },
], 'dropi_1', 'gt');
igual('la novedad derivada nace con su tienda', ['gt'], derivadas.map((n) => n.tienda));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
