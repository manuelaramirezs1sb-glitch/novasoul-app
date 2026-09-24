/**
 * MI PLATA, REAGRUPADA. Con sus propias cifras.
 *
 * ┌─ QUÉ PRUEBA ESTO Y QUÉ NO ─────────────────────────────────┐
 * │                                                            │
 * │ Ella dijo: «como lo veo no entiendo nada de lo que me      │
 * │ quiere transmitir la pantalla más que mi plata entrante    │
 * │ y ya». Eso no se prueba con una aserción — lo prueba ella  │
 * │ abriendo la pantalla.                                      │
 * │                                                            │
 * │ Lo que SÍ se puede probar, y es donde estaría el daño de   │
 * │ verdad, son las cuentas que hay debajo del reagrupamiento: │
 * │                                                            │
 * │   · que ninguna línea suya se caiga de todos los bloques   │
 * │     al reordenar — una deuda invisible es peor que una     │
 * │     deuda mal puesta;                                      │
 * │   · que las cuotas digan CUÁNDO terminan, que es el dato   │
 * │     que ella reclamó y el que separa deber de estar        │
 * │     pagando;                                               │
 * │   · que nunca se sumen dos monedas.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Los números son los de su pantalla real, los de la captura que
 * mandó: arriendo 788.000, mercado 230.000, servicios 200.000,
 * internet 93.000, crédito 223.000, deudas 690.000.
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

(0, eval)(src + '\n;globalThis.__F = { plataOrdenada_, plataLinea_, plataMasMeses_,' +
  ' soulPlataOrdenada, soulFijoGuardar, PLATA_TIPOS, PLATA_BLOQUES, SOUL_HOJAS };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_FIJ = F.SOUL_HOJAS.Fijos;
const YO = 'manuela@nova.com';
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };

function fila(o) {
  const v = Object.assign({ usuario_id: YO, activo: 'si', moneda: 'COP' }, o);
  return C_FIJ.map(c => (v[c] !== undefined ? v[c] : ''));
}

function sembrar() {
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

  UUID = 0;
  LIBROS.cen = { Finanzas: [['id','fecha','flujo','categoria','monto','moneda','concepto']],
                 Trabajos: [['id','nombre','tipo','estado']], Cobros: [['id']], Metas: [['id']] };
  LIBROS.s = {
    Fijos: [C_FIJ.slice()],
    Rutina: [['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin','lugar',
              'trabajo_id','materia_id','paga_fija','moneda','desde','hasta','activo','nota']],
    Turnos: [['id','usuario_id','rutina_id','fecha','paga','propinas','moneda','estado',
              'finanza_id','nota']],
    Pendientes: [['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
      'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
      'riesgo','nota','materia_id']],
    Usuarios: [['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
                'lugar_nacimiento','zona_horaria','acento','modo','idioma']],
  };
  // Su plan real, el de la captura.
  LIBROS.s.Fijos.push(fila({ id: 'f1', categoria: 'arriendo', concepto: 'apto',
    monto: 788000, dia_del_mes: 15, tipo_pago: 'mensual' }));
  LIBROS.s.Fijos.push(fila({ id: 'f2', categoria: 'mercado', concepto: '—',
    monto: 230000, dia_del_mes: 20, tipo_pago: 'mensual' }));
  LIBROS.s.Fijos.push(fila({ id: 'f3', categoria: 'servicios', concepto: 'luz, gas y agua',
    monto: 200000, dia_del_mes: 25, tipo_pago: 'mensual' }));
  LIBROS.s.Fijos.push(fila({ id: 'f4', categoria: 'internet', concepto: 'tigo',
    monto: 93000, tipo_pago: 'mensual' }));
  // Las dos que ella tenía apiladas bajo «Deudas», ahora con nombre y plazo.
  LIBROS.s.Fijos.push(fila({ id: 'f5', categoria: 'credito', concepto: 'nu',
    monto: 223000, dia_del_mes: 18, tipo_pago: 'cuotas', acreedor: 'Nu',
    cuotas_total: 12, cuota_desde: '2026-03' }));
  LIBROS.s.Fijos.push(fila({ id: 'f6', categoria: 'deudas', concepto: 'sistecredito',
    monto: 690000, dia_del_mes: 25, tipo_pago: 'cuotas', acreedor: 'Sistecrédito',
    cuotas_total: 6, cuota_desde: '2026-08' }));
}

// ══════════════════════════════════════════════════════════════
console.log('\n── Nada se pierde al reagrupar ──');
sembrar();
let o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
const enBloques = o.bloques.entra.lineas.length + o.bloques.fijo.lineas.length +
                  o.bloques.cuotas.lineas.length + o.bloques.unico.lineas.length +
                  o.bloques.ahorro.lineas.length;
igual('las 6 líneas siguen todas en algún bloque', 6, enBloques);

console.log('\n── Lo fijo y lo que es deuda quedan separados ──');
igual('4 gastos de todos los meses', 4, o.bloques.fijo.lineas.length);
igual('2 en cuotas', 2, o.bloques.cuotas.lineas.length);
igual('el fijo suma 1.311.000', { COP: 1311000 }, o.bloques.fijo.total);
igual('las cuotas suman 913.000 al mes', { COP: 913000 }, o.bloques.cuotas.total);
/**
 * Este es EL número que la pantalla vieja no podía dar: cuánto falta
 * en total, no cuánto toca este mes. Nu lleva 6 de 12 (marzo→sept), o
 * sea faltan 6 × 223.000 = 1.338.000. Sistecrédito lleva 1 de 6
 * (agosto→sept), faltan 5 × 690.000 = 3.450.000. Total 4.788.000.
 */
igual('y falta pagar 4.788.000 en total', { COP: 4788000 }, o.bloques.cuotas.faltaTodo);

console.log('\n── Las cuotas dicen cuándo se acaban ──');
const nu = o.bloques.cuotas.lineas.filter(l => l.concepto === 'nu')[0];
igual('Nu: van 6 de 12', [6, 12], [nu.cuotas.pagadas, nu.cuotas.total]);
igual('y se acaba en febrero de 2027', '2027-02', nu.cuotas.termina);
ok('las contó solas desde el mes de arranque', nu.cuotas.contadasSolas);
const sis = o.bloques.cuotas.lineas.filter(l => l.concepto === 'sistecredito')[0];
igual('Sistecrédito: van 1 de 6', [1, 6], [sis.cuotas.pagadas, sis.cuotas.total]);
igual('y se acaba en enero de 2027', '2027-01', sis.cuotas.termina);
igual('la última en acabarse es febrero', '2027-02', o.bloques.cuotas.ultima);

/**
 * Un mes DESPUÉS, las cuotas pagadas suben solas. Si hubiera que
 * actualizar una columna a mano cada mes, en tres meses estaría
 * mintiendo — y una deuda que miente es exactamente lo que ella no
 * quiere ver en esta pantalla.
 */
o = F.plataOrdenada_(YO, '2026-12', '2026-12-24');
igual('en diciembre, Nu ya va en 9 de 12', 9,
      o.bloques.cuotas.lineas.filter(l => l.concepto === 'nu')[0].cuotas.pagadas);

console.log('\n── Una deuda sin número de cuotas lo dice ──');
sembrar();
LIBROS.s.Fijos.push(fila({ id: 'f7', categoria: 'deudas', concepto: 'mama',
  monto: 100000, tipo_pago: 'cuotas', acreedor: 'Mamá' }));
o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
const mama = o.bloques.cuotas.lineas.filter(l => l.concepto === 'mama')[0];
igual('no inventa una fecha de fin', '', mama.cuotas.termina);
ok('y explica qué falta para poder darla',
   /cuántas cuotas son/.test(mama.cuotas.porque), mama.cuotas.porque);
igual('la cuenta la excluye en vez de inventarla', { COP: 4788000 },
      o.bloques.cuotas.faltaTodo);
igual('y avisa que hay una sin fecha', 1, o.bloques.cuotas.sinFecha);

console.log('\n── Lo que entra va aparte de lo que sale ──');
sembrar();
LIBROS.s.Fijos.push(fila({ id: 'f8', categoria: 'varios', concepto: 'mesada',
  monto: 500000, flujo: 'ingreso', tipo_pago: 'mensual' }));
o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
igual('la mesada entra al bloque de lo que entra', 1, o.bloques.entra.lineas.length);
ok('y NO aparece entre lo que sale',
   o.bloques.fijo.lineas.filter(l => l.concepto === 'mesada').length === 0);
igual('lo que sale sigue siendo 2.224.000', { COP: 2224000 }, o.sale);

console.log('\n── Nunca se suman dos monedas ──');
sembrar();
LIBROS.s.Fijos.push(fila({ id: 'f9', categoria: 'varios', concepto: 'hosting',
  monto: 12, moneda: 'USD', tipo_pago: 'mensual' }));
o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
igual('el fijo trae las dos monedas, separadas', { COP: 1311000, USD: 12 },
      o.bloques.fijo.total);
igual('y el resultado también', ['COP', 'USD'], o.resultado.map(r => r.moneda));

console.log('\n── Lo que Nova dedujo se puede corregir ──');
sembrar();
LIBROS.s.Fijos.push(fila({ id: 'f10', categoria: 'deudas', concepto: 'vieja sin tipo',
  monto: 50000 }));
o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
ok('una línea vieja sin tipo de pago NO desaparece',
   o.bloques.cuotas.lineas.filter(l => l.concepto === 'vieja sin tipo').length === 1);
igual('y queda en la lista de lo que hay que confirmar', 1, o.porClasificar.length);
igual('diciendo qué supuso', 'Por cuotas', o.porClasificar[0].supuesto);

console.log('\n── El formulario no acepta un tipo inventado ──');
sembrar();
const mal = F.soulFijoGuardar(SOCIA, { datos: { categoria: 'varios', concepto: 'x',
  monto: 1, tipoPago: 'cuando pueda' } });
igual('lo rechaza', false, mal.ok);
ok('y dice cuáles valen', /mensual, cuotas o único/.test(mal.error), mal.error);

const bien = F.soulFijoGuardar(SOCIA, { datos: { categoria: 'deudas', concepto: 'tarjeta',
  monto: 300000, tipoPago: 'cuotas', cuotasTotal: 9, cuotaDesde: '2026-09',
  acreedor: 'Bancolombia' } });
igual('una buena se guarda', true, bien.ok);
o = F.plataOrdenada_(YO, '2026-09', '2026-09-24');
const tar = o.bloques.cuotas.lineas.filter(l => l.concepto === 'tarjeta')[0];
ok('y llega completa a la pantalla', !!tar && tar.acreedor === 'Bancolombia' &&
   tar.cuotas.total === 9, JSON.stringify(tar && tar.cuotas));
igual('con su fecha de salida', '2027-05', tar.cuotas.termina);

console.log('\n── Sumar meses no se equivoca en el cambio de año ──');
igual('2026-09 + 5 = 2027-02', '2027-02', F.plataMasMeses_('2026-09', 5));
igual('2026-12 + 1 = 2027-01', '2027-01', F.plataMasMeses_('2026-12', 1));
igual('2026-01 + 0 = 2026-01', '2026-01', F.plataMasMeses_('2026-01', 0));
igual('2026-01 + 23 = 2027-12', '2027-12', F.plataMasMeses_('2026-01', 23));

console.log('\n── La llamada de la pantalla ──');
sembrar();
const r = F.soulPlataOrdenada(SOCIA, { mes: '2026-09' });
ok('responde', r.ok);
ok('trae el orden nuevo', !!r.orden && r.orden.bloques.cuotas.lineas.length === 2);
ok('y conserva lo que la pantalla vieja ya usaba', !!r.categorias && !!r.resumen);
igual('una operadora no entra', false,
      F.soulPlataOrdenada({ correo: 'k@n.com', rol: 'operadora' }, {}).ok);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
