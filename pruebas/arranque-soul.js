/**
 * POR QUÉ NOVASOUL SE DEMORA. Medido, no opinado.
 *
 * ┌─ LA PREGUNTA ──────────────────────────────────────────────┐
 * │                                                            │
 * │ «en novasoul está muy lento para cargar todo. no sé qué    │
 * │  podamos hacer para optimizar, el api será que sí puede    │
 * │  con toda esta información en un solo api? en eso no se    │
 * │  podría reemplazar con n8n? para que no sea tan lento o    │
 * │  no tiene nada que ver?»                                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE ESTA PRUEBA PUEDE MEDIR Y LO QUE NO ───────────────┐
 * │                                                            │
 * │ SÍ puede: cuántas veces se abre el libro y cuántas celdas  │
 * │ se mueven. Eso es trabajo real y se cuenta exacto.         │
 * │                                                            │
 * │ NO puede: los segundos de red. Cada `nc(...)` de la        │
 * │ pantalla es una petición HTTPS entera a Apps Script, con   │
 * │ su arranque de motor y su comprobación de sesión. Eso son  │
 * │ entre medio segundo y dos segundos CADA UNA, y no depende  │
 * │ de cuántos datos tenga: depende de cuántas peticiones se   │
 * │ hagan.                                                     │
 * │                                                            │
 * │ Por eso la prueba cuenta las dos cosas por separado: las   │
 * │ peticiones (que es donde está el problema de ella) y las   │
 * │ celdas (que es donde NO está).                             │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LECTURAS = {};
const CELDAS = {};
let ABIERTAS = 0;

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libro(id) {
  ABIERTAS++;
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (nombre) => {
    const m = hojas[nombre];
    if (!m) return null;
    const ancho = m[0] ? m[0].length : 0;
    const k = id + '·' + nombre;
    return {
      getLastRow: () => m.length,
      getLastColumn: () => ancho,
      getDataRange: () => ({ getValues: () => {
        LECTURAS[k] = (LECTURAS[k] || 0) + 1;
        CELDAS[k] = (CELDAS[k] || 0) + m.length * ancho;
        return m.map(f => f.slice());
      } }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (n) => m.splice(n - 1, 1),
      getRange: (f, c, nf, nc) => ({
        getValues: () => {
          CELDAS[k] = (CELDAS[k] || 0) + (nf || 1) * (nc || ancho);
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
const HOY = '2026-09-24T12:00:00Z';
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

(0, eval)(src + '\n;globalThis.__F = { soulHoy, soulPlataOrdenada, soulCielo, soulRutina,' +
  ' soulMaterias, soulArranque, libroOlvidar_, soulOlvidar_, SOUL_HOJAS };');
const F = globalThis.__F;

const YO = 'manuela@nova.com';
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const H = F.SOUL_HOJAS;

function fila(cols, o) {
  const v = Object.assign({ usuario_id: YO }, o);
  return cols.map(c => (v[c] !== undefined ? v[c] : ''));
}

function sembrar() {
  UUID = 0;
  /**
   * Soltar los manejadores de libro, como hace cada petición real. Sin
   * esto, el manejador memorizado seguiría apuntando a los arreglos del
   * caso anterior y la medición saldría más bonita de lo que es.
   */
  F.libroOlvidar_();
  F.soulOlvidar_();
  Object.keys(LECTURAS).forEach(k => delete LECTURAS[k]);
  Object.keys(CELDAS).forEach(k => delete CELDAS[k]);
  ABIERTAS = 0;

  LIBROS.cen = {
    Finanzas: [['id','fecha','flujo','categoria','monto','moneda','concepto']],
    Trabajos: [['id','nombre','tipo','estado','horas_semana','fecha_entrega','padre_id',
                'periodicidad','dias_pago']],
    Cobros: [['id','trabajo_id','fecha','monto','moneda','estado']],
    Metas: [['id','nombre','valor']],
    Proyectos: [['id','trabajo_id','nombre','estado']],
  };
  LIBROS.s = {
    Pendientes: [H.Pendientes.slice()], Materias: [H.Materias.slice()],
    Mindlab: [H.Mindlab.slice()], Fijos: [H.Fijos.slice()],
    Rutina: [H.Rutina.slice()], Turnos: [H.Turnos.slice()], Pensum: [H.Pensum.slice()],
    Carta: [['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota']],
    Transitos: [['usuario_id','fecha','casa','casa_placidus','tema','intensidad_pct',
      'texto_transito','por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto',
      'a_natal','desde','hasta','fuente']],
    Revolucion: [['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota']],
    Usuarios: [['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
                'lugar_nacimiento','zona_horaria','acento','modo','idioma'],
               ['u1','Manuela',YO,'1995-09-20','13:20','Palmira','America/Bogota','','','']],
    Horas: [['usuario_id','dia_semana','horas_libres','nota']],
  };
  // Una carga como la suya: los pendientes de un semestre, la rutina de
  // la semana, los turnos de tres meses.
  for (let i = 0; i < 120; i++) {
    LIBROS.s.Pendientes.push(fila(H.Pendientes, { id: 'p' + i, texto: 'Pendiente ' + i,
      fecha: '2026-09-' + (10 + (i % 20)), estado: i % 3 ? 'pendiente' : 'hecho',
      horas_estimadas: 2, tipo: 'trabajo' }));
  }
  for (let i = 0; i < 12; i++) {
    LIBROS.s.Rutina.push(fila(H.Rutina, { id: 'r' + i, tipo: i < 4 ? 'turno' : 'clase',
      nombre: 'Bloque ' + i, dia_semana: 'LMXJVSD'[i % 7], hora_inicio: '08:00',
      hora_fin: '12:00', paga_fija: i < 4 ? 80000 : '', moneda: 'COP', activo: 'si' }));
  }
  for (let i = 0; i < 90; i++) {
    LIBROS.s.Turnos.push(fila(H.Turnos, { id: 't' + i, rutina_id: 'r' + (i % 4),
      fecha: '2026-0' + (7 + Math.floor(i / 31)) + '-' + (1 + i % 28), paga: 80000,
      propinas: 15000, moneda: 'COP', estado: 'hecho' }));
  }
  for (let i = 0; i < 9; i++) {
    LIBROS.s.Fijos.push(fila(H.Fijos, { id: 'f' + i, categoria: 'varios',
      concepto: 'Gasto ' + i, monto: 100000, moneda: 'COP', activo: 'si',
      tipo_pago: 'mensual' }));
  }
  for (let i = 0; i < 6; i++) {
    LIBROS.s.Materias.push(fila(H.Materias, { id: 'm' + i, nombre: 'Materia ' + i,
      estado: 'activa' }));
  }
}

const col = (s, n, der) => {
  s = String(s);
  const h = Math.max(0, n - s.length);
  return der ? ' '.repeat(h) + s : s + ' '.repeat(h);
};

function medir(nombre, fn) {
  sembrar();
  const t0 = process.hrtime.bigint();
  let err = '';
  try { fn(); } catch (e) { err = e.message; }
  const ms = Number(process.hrtime.bigint() - t0) / 1e6;
  const celdas = Object.keys(CELDAS).reduce((a, k) => a + CELDAS[k], 0);
  const lecturas = Object.keys(LECTURAS).reduce((a, k) => a + LECTURAS[k], 0);
  return { nombre, ms, celdas, lecturas, abiertas: ABIERTAS, err };
}

console.log('\nNOVASOUL AL ENTRAR · una carga como la suya\n');
console.log(col('PETICIÓN', 22) + col('ms', 8, 1) + col('abre libro', 12, 1) +
            col('hojas leídas', 14, 1) + col('celdas', 10, 1));
console.log('─'.repeat(66));

/** Las que dispara la pantalla, una por sección que ella toca. */
const ANTES = [
  ['nc_soul (Hoy)',      () => F.soulHoy(SOCIA, {})],
  ['nc_soul_plata',      () => F.soulPlataOrdenada(SOCIA, {})],
  ['nc_soul_cielo',      () => F.soulCielo(SOCIA, {})],
  ['nc_soul_rutina',     () => F.soulRutina(SOCIA, {})],
  ['nc_soul_materias',   () => F.soulMaterias(SOCIA, {})],
];

let tMs = 0, tCeldas = 0, tLect = 0, tAbre = 0;
ANTES.forEach(function (par) {
  const r = medir(par[0], par[1]);
  tMs += r.ms; tCeldas += r.celdas; tLect += r.lecturas; tAbre += r.abiertas;
  console.log(col(r.nombre, 22) + col(r.ms.toFixed(1), 8, 1) +
              col(r.abiertas, 12, 1) + col(r.lecturas, 14, 1) +
              col(r.celdas.toLocaleString('es'), 10, 1) + (r.err ? '  ⚠ ' + r.err : ''));
});
console.log('─'.repeat(66));
console.log(col('TOTAL · ' + ANTES.length + ' peticiones', 22) + col(tMs.toFixed(1), 8, 1) +
            col(tAbre, 12, 1) + col(tLect, 14, 1) + col(tCeldas.toLocaleString('es'), 10, 1));

/**
 * Y lo mismo en UNA sola petición. El trabajo de hojas es casi el
 * mismo —los datos son los que son—; lo que se ahorra son cuatro
 * idas y vueltas por internet, que es de donde salen los segundos.
 */
const uno = medir('nc_soul_arranque', () => F.soulArranque(SOCIA, {}));
console.log('\n' + col('EN UNA SOLA', 22) + col(uno.ms.toFixed(1), 8, 1) +
            col(uno.abiertas, 12, 1) + col(uno.lecturas, 14, 1) +
            col(uno.celdas.toLocaleString('es'), 10, 1) + (uno.err ? '  ⚠ ' + uno.err : ''));

console.log('\n' + '─'.repeat(66));
console.log('Peticiones HTTPS a Apps Script:  ' + ANTES.length + '  →  1');
console.log('Cada una cuesta entre 0,5 y 2 segundos de ida y vuelta,');
console.log('gaste lo que gaste por dentro. AHÍ está la demora, no en las celdas.');
console.log('Libro abierto:  ' + tAbre + '  →  ' + uno.abiertas);
console.log('Hojas leídas enteras:  ' + tLect + '  →  ' + uno.lecturas);
console.log('─'.repeat(66) + '\n');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

ok('la llamada única no falla', !uno.err, uno.err);
ok('abre el libro menos veces que las cinco juntas', uno.abiertas < tAbre,
   uno.abiertas + ' vs ' + tAbre);
ok('lee menos hojas enteras que las cinco juntas', uno.lecturas < tLect,
   uno.lecturas + ' vs ' + tLect);

/** Y que de verdad traiga las cinco secciones, no una con buen nombre. */
sembrar();
const r = F.soulArranque(SOCIA, {});
ok('trae Hoy', !!r.hoy_);
ok('trae la plata', !!r.plata);
ok('trae el cielo', !!r.cielo);
ok('trae la rutina', !!r.rutina);
ok('trae la universidad', !!r.materias);
ok('una operadora no entra',
   F.soulArranque({ correo: 'k@n.com', rol: 'operadora' }, {}).ok === false);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
