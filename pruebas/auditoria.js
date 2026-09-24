/**
 * AUDITORÍA · que el veredicto se guarde de verdad.
 *
 * ┌─ LO QUE ESTA PRUEBA EXISTE PARA QUE NO VUELVA A PASAR ─────┐
 * │                                                            │
 * │ «auditorías tampoco está guardando los hallazgos».          │
 * │                                                            │
 * │ La pantalla era una maqueta entera: tres clientas           │
 * │ inventadas y un botón «Guardar veredicto» SIN `onclick`.    │
 * │ Se marcaba, se escribía la nota, se guardaba, y no pasaba   │
 * │ nada. Ni un error — no había nada que pudiera fallar.       │
 * │                                                            │
 * │ Por eso la primera aserción de este archivo no es sobre el  │
 * │ cruce ni sobre las señales: es que guardar, guarde, y que   │
 * │ al volver a abrir el veredicto siga ahí. Lo demás es        │
 * │ detalle al lado de eso.                                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Y de paso la otra que ella encontró en la misma pantalla: «he
 * entrado desde el correo de Nova Soul y no se registra ahí».
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

(0, eval)(src + '\n;globalThis.__F = { apiAuditoriaCasos, apiAuditoriaGuardar, audSenales_,' +
  ' apiAuditoria, buscarPersona, apiVerificar, apiEquipo, libroOlvidar_, soulOlvidar_,' +
  ' ESQUEMA_EMPRESARIAL, AUD_VEREDICTOS };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_NOV = F.ESQUEMA_EMPRESARIAL.Novedades;
const C_PED = F.ESQUEMA_EMPRESARIAL.Pedidos;
const C_LLA = F.ESQUEMA_EMPRESARIAL.Llamadas;
const C_AUD = F.ESQUEMA_EMPRESARIAL.Auditorias;
const C_EQU = F.ESQUEMA_EMPRESARIAL.Equipo;

function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

function sembrar() {
  UUID = 0;
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.cen = {
    Clientes: [['id','nombre','correo','plan','x5','x6','estado','x8','x9','x10','x11',
                'x12','x13','sheet_id'],
               ['c1','Nutrea','m@nova.com','pro','','','activo','','','','','','','emp']],
  };
  LIBROS.emp = {
    Novedades: [C_NOV.slice()], Pedidos: [C_PED.slice()], Llamadas: [C_LLA.slice()],
    Auditorias: [C_AUD.slice()], Equipo: [C_EQU.slice()],
    Movimientos: [['fecha','usuario','entidad','entidad_id','campo',
                   'valor_anterior','valor_nuevo']],
    Tiendas: [['id','nombre','moneda','estado','modalidad'],
              ['ec','Nutrea EC','USD','activa','marca_propia']],
  };
  LIBROS.emp.Equipo.push(f(C_EQU, { id: 'e1', nombre: 'Manuela', correo: 'm@nova.com',
    rol: 'dueno', tienda: '*', estado: 'activo' }));
  LIBROS.emp.Equipo.push(f(C_EQU, { id: 'e2', nombre: 'Camila', correo: 'c@nova.com',
    rol: 'gestora', tienda: 'ec', estado: 'activo' }));

  // Tres pedidos con clienta y teléfono.
  [['p1', 'Andrea Morales', '593991111111', 'Quito'],
   ['p2', 'Luisa Fernández', '593992222222', 'Guayaquil'],
   ['p3', 'Carolina Jiménez', '593993333333', 'Cuenca'],
  ].forEach(function (x) {
    LIBROS.emp.Pedidos.push(f(C_PED, { id: x[0], tienda: 'ec', cliente: x[1],
      telefono: x[2], telefono_norm: x[2], ciudad: x[3], fecha: '2026-09-18' }));
  });
}

const DUENA = { sheetId: 'emp', rol: 'dueno', email: 'm@nova.com', nombre: 'Manuela',
                tiendas: ['ec'], permisos: [], modulos: ['empresarial'] };
const GESTORA = { sheetId: 'emp', rol: 'gestora', email: 'c@nova.com', nombre: 'Camila',
                  tiendas: ['ec'], permisos: [], modulos: ['empresarial'] };

// ══════════════════════════════════════════════════════════════
console.log('\n── Lo primero: que guardar, guarde ──');
sembrar();
LIBROS.emp.Novedades.push(f(C_NOV, { id: 'n1', pedido_id: 'p1', fecha: '2026-09-18',
  motivo: 'No contesta', estado: 'resuelta', solucionada: 'si',
  fecha_solucion: '2026-09-19', gestora: 'Camila',
  solucion: 'La clienta confirmó recepción. Entregué a las 14:20.' }));

let r = F.apiAuditoriaCasos(DUENA, { tienda: 'ec' });
igual('el caso aparece para revisar', 1, r.casos.length);
igual('y arranca sin veredicto', 'pendiente', r.casos[0].veredicto);

const g = F.apiAuditoriaGuardar(DUENA, { novedadId: 'n1', pedidoId: 'p1',
  gestora: 'Camila', veredicto: 'ok', nota: 'Revisado, todo en orden.', tienda: 'ec' });
igual('el veredicto se guarda', true, g.ok);
igual('y queda una fila en la hoja', 1, LIBROS.emp.Auditorias.length - 1);

/** Y —lo que de verdad falló— que al volver a abrir SIGA ahí. */
r = F.apiAuditoriaCasos(DUENA, { tienda: 'ec' });
igual('al volver a abrir, el veredicto sigue puesto', 'ok', r.casos[0].veredicto);
igual('con su nota', 'Revisado, todo en orden.', r.casos[0].notaAuditoria);
igual('y con quién lo puso', 'm@nova.com', r.casos[0].auditadoPor);
igual('el recuento dice que hay 1 revisado', 1, r.recuento.revisados);
igual('y 0 sin revisar', 0, r.recuento.sinRevisar);

console.log('\n── Cambiar de opinión corrige, no apila ──');
const g2 = F.apiAuditoriaGuardar(DUENA, { novedadId: 'n1', veredicto: 'hallazgo',
  nota: 'Revisando de nuevo: falta la hora de entrega real.', tienda: 'ec' });
igual('se puede cambiar', true, g2.ok);
igual('avisa que reemplazó', true, g2.reemplazo);
igual('y sigue habiendo UNA sola fila', 1, LIBROS.emp.Auditorias.length - 1);
igual('con el veredicto nuevo', 'hallazgo',
      F.apiAuditoriaCasos(DUENA, { tienda: 'ec' }).casos[0].veredicto);
/** El cambio no se pierde: queda en Movimientos, que es el sitio del rastro. */
ok('y el cambio quedó en el rastro',
   LIBROS.emp.Movimientos.filter(m => m[2] === 'Auditorias').length === 2,
   JSON.stringify(LIBROS.emp.Movimientos.slice(1)));

console.log('\n── Un hallazgo sin nota no entra ──');
sembrar();
LIBROS.emp.Novedades.push(f(C_NOV, { id: 'n1', pedido_id: 'p1', fecha: '2026-09-18',
  estado: 'resuelta', fecha_solucion: '2026-09-19', gestora: 'Camila', solucion: 'ok' }));
const sinNota = F.apiAuditoriaGuardar(DUENA, { novedadId: 'n1', veredicto: 'hallazgo',
  nota: '' });
igual('se rechaza', false, sinNota.ok);
ok('y dice por qué importa', /[Qq]ueda en el registro de esa persona/.test(sinNota.error),
   sinNota.error);
igual('no se guardó nada', 0, LIBROS.emp.Auditorias.length - 1);
/** Un OK sí puede ir sin nota: cien casos buenos no pueden costar cien párrafos. */
igual('un OK sin nota sí entra', true,
      F.apiAuditoriaGuardar(DUENA, { novedadId: 'n1', veredicto: 'ok' }).ok);

console.log('\n── Un veredicto inventado no entra ──');
const inv = F.apiAuditoriaGuardar(DUENA, { novedadId: 'n1', veredicto: 'más o menos' });
igual('se rechaza', false, inv.ok);
ok('y dice cuáles valen', /OK, Hallazgo o Pendiente/.test(inv.error), inv.error);

console.log('\n── Quién puede auditar ──');
igual('una gestora no ve la auditoría', false,
      F.apiAuditoriaCasos(GESTORA, { tienda: 'ec' }).ok);
igual('ni pone veredictos', false,
      F.apiAuditoriaGuardar(GESTORA, { novedadId: 'n1', veredicto: 'ok' }).ok);

// ══════════════════════════════════════════════════════════════
console.log('\n── El cruce con la central: dos relatos del mismo hecho ──');
sembrar();
LIBROS.emp.Novedades.push(f(C_NOV, { id: 'n2', pedido_id: 'p2', fecha: '2026-09-18',
  motivo: 'Dirección incorrecta', estado: 'resuelta', fecha_solucion: '2026-09-19',
  gestora: 'Camila', solucion: 'Dirección incorrecta. Llamé 3 veces, no contesta.' }));
// La central solo registra DOS.
[['2026-09-19 10:30', 0], ['2026-09-19 12:15', 0]].forEach(function (x, i) {
  LIBROS.emp.Llamadas.push(f(C_LLA, { id: 'l' + i, fecha_hora: x[0], tienda: 'ec',
    sentido: 'saliente', estado: 'sin_respuesta', agente: 'Camila',
    telefono_norm: '593992222222', pedido_id: 'p2', seg_conversado: x[1] }));
});

r = F.apiAuditoriaCasos(DUENA, { tienda: 'ec' });
const c2 = r.casos.filter(c => c.novedadId === 'n2')[0];
igual('trae lo que escribió la gestora',
      'Dirección incorrecta. Llamé 3 veces, no contesta.', c2.dijoGestora);
igual('y las llamadas que sí hay', 2, c2.llamadas.length);
ok('Nova levanta la discrepancia',
   c2.senales.filter(x => x.tipo === 'llamadas_de_menos').length === 1,
   JSON.stringify(c2.senales));
igual('con los dos números a la vista', { dijo: 3, hay: 2 },
      c2.senales.filter(x => x.tipo === 'llamadas_de_menos')[0].dato);
/** Y NO pone el veredicto sola: eso es de una persona. */
igual('pero el veredicto sigue siendo de una persona', 'pendiente', c2.veredicto);

console.log('\n── «Llamé tres veces», con letra ──');
igual('lo entiende escrito con letra', 1,
  F.audSenales_({ solucion: 'Llamé tres veces y nada.' },
                [{ sentido: 'saliente', seg_conversado: 0 }])
   .filter(x => x.tipo === 'llamadas_de_menos').length);

console.log('\n── Las otras señales ──');
igual('dice que acordó y no hay ni una llamada', 1,
  F.audSenales_({ solucion: 'Reprogramé con acuerdo de la clienta.' }, [])
   .filter(x => x.tipo === 'sin_llamada').length);
igual('dice que confirmó pero nadie contestó', 1,
  F.audSenales_({ solucion: 'La clienta confirmó.' },
                [{ sentido: 'saliente', seg_conversado: 0 }])
   .filter(x => x.tipo === 'nadie_contesto').length);
igual('cerró sin escribir qué hizo', 1,
  F.audSenales_({ solucion: '', estado: 'resuelta' }, [])
   .filter(x => x.tipo === 'sin_nota').length);
ok('y cuando todo calza, lo dice también',
  F.audSenales_({ solucion: 'Llamé 1 vez y quedó entregado.' },
                [{ sentido: 'saliente', seg_conversado: 90 }])
   .filter(x => x.nivel === 'bien').length === 1);

console.log('\n── Una novedad abierta no se audita ──');
sembrar();
LIBROS.emp.Novedades.push(f(C_NOV, { id: 'n3', pedido_id: 'p3', fecha: '2026-09-22',
  estado: 'abierta', gestora: 'Camila' }));
igual('no aparece en la lista', 0, F.apiAuditoriaCasos(DUENA, { tienda: 'ec' }).casos.length);

console.log('\n── Sin llamadas importadas, se dice ──');
sembrar();
LIBROS.emp.Novedades.push(f(C_NOV, { id: 'n1', pedido_id: 'p1', fecha: '2026-09-18',
  estado: 'resuelta', fecha_solucion: '2026-09-19', gestora: 'Camila', solucion: 'ok' }));
r = F.apiAuditoriaCasos(DUENA, { tienda: 'ec' });
igual('avisa que falta la mitad del cruce', true, r.sinLlamadas);
ok('y explica qué falta importar', /IRIS/.test(r.porqueSinLlamadas), r.porqueSinLlamadas);

console.log('\n── Sin novedades, no finge ──');
sembrar();
r = F.apiAuditoriaCasos(DUENA, { tienda: 'ec' });
igual('no hay casos', 0, r.casos.length);
ok('y dice qué tiene que pasar para que los haya',
   /cierre la primera/.test(r.porque), r.porque);

// ══════════════════════════════════════════════════════════════
console.log('\n── Y el otro fallo de la misma pantalla: la conexión ──');
/**
 * «he entrado desde el correo de Nova Soul y no se registra ahí; si
 *  así estará cuando se tenga un equipo, estamos mal».
 *
 * Entrar anotaba el movimiento en el rastro pero nunca escribía la
 * celda `ultima_conexion` de la hoja Equipo, que es la que lee la
 * pantalla. Resultado: «Nunca ha entrado», para siempre.
 */
sembrar();
const colCon = C_EQU.indexOf('ultima_conexion');
igual('antes de entrar, la celda está vacía', '', LIBROS.emp.Equipo[1][colCon]);

const persona = F.buscarPersona('m@nova.com');
ok('se encuentra a la persona', !!persona);
igual('y se sabe en qué fila está', 2, persona.fila);

/**
 * Entrar de verdad, por la misma puerta que ella: código al correo y
 * canje del código. Se pone el código en el caché a mano porque el
 * correo no sale de una prueba.
 */
global.CacheService = { getScriptCache: () => ({
  get: (k) => (k === 'cod_m@nova.com' ? '123456' : null),
  put: () => {}, remove: () => {} }) };
const entrada = F.apiVerificar({ email: 'm@nova.com', codigo: '123456' });
igual('entra', true, entrada.ok);
ok('ahora la celda tiene fecha', !!LIBROS.emp.Equipo[1][colCon],
   JSON.stringify(LIBROS.emp.Equipo[1][colCon]));
ok('y sigue quedando en el rastro',
   LIBROS.emp.Movimientos.filter(m => m[4] === 'ultima_conexion').length === 1);

const eq = F.apiEquipo(DUENA, { tienda: 'ec' });
const yo = (eq.personas || eq.equipo || []).filter(x => x.correo === 'm@nova.com')[0];
ok('y la pantalla ya no dice «nunca ha entrado»', !!(yo && yo.ultima_conexion),
   JSON.stringify(yo && yo.ultima_conexion));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
