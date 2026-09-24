/**
 * «Qué quedó de ayer»: la lista de trabajo del día.
 *
 * Ella lo pidió con una referencia concreta —una tabla de categorías de
 * gestión con su cantidad— y eligió dos cosas como pendiente: lo que
 * sigue esperando a una persona, y lo que se venció y nadie movió.
 *
 * Lo que esta prueba cuida, que es lo que hace útil o inútil el reporte:
 *
 * 1· Que un pedido de esta mañana y uno de hace tres días NO se cuenten
 *    igual. El segundo casi siempre termina en devolución, y si la lista
 *    no los separa, nadie ve cuál es la parte que se está pudriendo.
 * 2· Que lo que está EN CAMINO no cuente como pendiente. No espera a
 *    nadie: espera al courier. Meterlo infla la lista con cosas que no
 *    se pueden hacer, y una lista así deja de leerse.
 * 3· Que un pedido de ayer que YA se entregó esta mañana no aparezca.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cli: {} };
function libro(id) {
  const hojas = LIBROS[id] || {};
  return {
    getSheetByName: (nombre) => {
      const m = hojas[nombre];
      if (!m) return null;
      return {
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
          setValues: () => {}, setValue: () => {},
        }),
      };
    },
  };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T14:00:00Z';
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
global.Utilities = { sleep: () => {}, getUuid: () => 'u1',
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

(0, eval)(src + '\n;globalThis.__F = { reporteDelDia, reporteDelDiaTexto, apiReporteDia,' +
  ' DIA_LIMITES, DIA_CATEGORIAS };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PED = ['id','id_externo','tienda','fecha','cliente','telefono','ciudad',
               'producto','valor','estado_canonico','estado_nova','gestora_asignada'];
const C_NOV = ['id','pedido_id','fecha','motivo','grupo','estado'];
function ped(o) { return C_PED.map(c => (o[c] === undefined ? '' : o[c])); }
function nov(o) { return C_NOV.map(c => (o[c] === undefined ? '' : o[c])); }

function sembrar(pedidos, novedades) {
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

  LIBROS.cli = {
    Pedidos: [C_PED].concat(pedidos || []),
    Novedades: [C_NOV].concat(novedades || []),
    Tiendas: [['id','nombre','moneda','estado'], ['ec','Nutrea EC','USD','activa']],
    Parametros: [['tienda','clave','valor']],
  };
}
const DUENA = { sheetId: 'cli', rol: 'dueno', email: 'm@nova.com', tiendas: ['ec'] };

// ─────────────────────────────────────────────────────────────
console.log('\n── Lo que quedó esperando ──');
sembrar([
  // De ayer, sin confirmar: pendiente fresco
  ped({ id: 'p1', id_externo: '1', tienda: 'ec', fecha: '2026-09-23',
        cliente: 'Lucía', telefono: '0991', valor: 50, estado_canonico: 'pendiente' }),
  // De hace cuatro días, sin confirmar: pendiente PODRIDO
  ped({ id: 'p2', id_externo: '2', tienda: 'ec', fecha: '2026-09-20',
        cliente: 'Andrés', telefono: '0992', valor: 50, estado_canonico: 'pendiente' }),
  // Novedad abierta de hace tres días
  ped({ id: 'p3', id_externo: '3', tienda: 'ec', fecha: '2026-09-21',
        cliente: 'Sofía', telefono: '0993', valor: 60, estado_canonico: 'novedad' }),
  // En oficina, dentro de plazo
  ped({ id: 'p4', id_externo: '4', tienda: 'ec', fecha: '2026-09-23',
        cliente: 'Karla', valor: 40, estado_canonico: 'en_oficina' }),
  // EN CAMINO: no espera a nadie
  ped({ id: 'p5', id_externo: '5', tienda: 'ec', fecha: '2026-09-23',
        cliente: 'Pablo', valor: 30, estado_canonico: 'en_transito' }),
  // De ayer y YA entregado hoy: no es pendiente, es trabajo hecho
  ped({ id: 'p6', id_externo: '6', tienda: 'ec', fecha: '2026-09-23',
        cliente: 'Rosa', valor: 70, estado_canonico: 'entregado' }),
  // De otra tienda: no se mezcla
  ped({ id: 'p7', id_externo: '7', tienda: 'gt', fecha: '2026-09-23',
        cliente: 'Otra', valor: 99, estado_canonico: 'pendiente' }),
], [
  nov({ id: 'n1', pedido_id: '3', fecha: '2026-09-21', grupo: 'no_contacta', estado: 'abierta' }),
  nov({ id: 'n2', pedido_id: '8', fecha: '2026-09-24', grupo: 'direccion', estado: 'abierta' }),
  nov({ id: 'n3', pedido_id: '9', fecha: '2026-09-01', grupo: 'no_contacta', estado: 'resuelta' }),
]);

let r = F.reporteDelDia('cli', 'ec');
ok('responde', r.ok === true);
igual('mira ayer por defecto', '2026-09-23', r.dia);
igual('con el nombre de pantalla', 'Nutrea EC', r.nombreTienda);

const C = {}; r.categorias.forEach(c => { C[c.id] = c; });
igual('dos sin confirmar', 2, C.pendiente.n);
/** La que importa: uno de los dos lleva cuatro días. */
igual('y UNO de ellos lleva demasiado', 1, C.pendiente.viejos);
igual('una novedad abierta', 1, C.novedad.n);
igual('que además está vencida', 1, C.novedad.viejos);
igual('uno en oficina, dentro de plazo', [1, 0],
      [C.en_oficina.n, C.en_oficina.viejos]);

console.log('\n── Lo que NO cuenta como pendiente ──');
igual('lo que va en camino va aparte', 1, r.enCamino);
ok('y no está en ninguna categoría de trabajo',
   r.categorias.every(c => c.id !== 'en_transito'));
igual('el total esperando son cuatro, no seis', 4, r.total);
ok('el entregado no aparece',
   !r.categorias.some(c => c.casos.some(p => p.cliente === 'Rosa')));
ok('ni el de la otra tienda',
   !r.categorias.some(c => c.casos.some(p => p.cliente === 'Otra')));
igual('pero sí se dice cuántos se cerraron ese día', 1, r.cerradosEseDia);
igual('y cuántos entraron', 4, r.entraronEseDia);

console.log('\n── Los motivos de las novedades ──');
igual('solo las abiertas', 2, r.novedades.abiertas);
igual('y cuántas llevan más de un día', 1, r.novedades.viejas);
igual('agrupadas por motivo', ['no_contacta', 'direccion'], r.motivos.map(m => m.grupo));

console.log('\n── Los casos, en el orden en que hay que llamarlos ──');
igual('el más viejo primero', 'Andrés', C.pendiente.casos[0].cliente);
igual('con sus días quieto', 4, C.pendiente.casos[0].dias);
ok('y con su teléfono, para poder llamar', C.pendiente.casos[0].telefono === '0992');
igual('la plata detenida en lo que espera', 200, r.valorPendiente);

console.log('\n── El texto, el que va al correo y a imprimir ──');
const txt = F.reporteDelDiaTexto(r);
ok('dice cuántos esperan', /HAY 4 PEDIDOS ESPERANDO A ALGUIEN/.test(txt), txt);
ok('y cuántos llevan demasiado', /2 de esos llevan más tiempo/.test(txt), txt);
ok('con la tabla de categorías', /Sin confirmar/.test(txt) && /CATEGORÍA/.test(txt));
ok('lo que va en camino, dicho aparte',
   /En camino, sin nada que hacer: 1/.test(txt), txt);
ok('los motivos de las novedades', /POR QUÉ ESTÁN ABIERTAS/.test(txt));
/** Un total no se puede trabajar; una lista de personas sí. */
ok('y los más quietos con nombre y teléfono',
   /LOS QUE LLEVAN MÁS QUIETOS/.test(txt) && /Andrés · 4 d/.test(txt) &&
   /0992/.test(txt), txt);

console.log('\n── Un día limpio ──');
sembrar([
  ped({ id: 'p1', tienda: 'ec', fecha: '2026-09-23', cliente: 'Ya', valor: 10,
        estado_canonico: 'entregado' }),
]);
r = F.reporteDelDia('cli', 'ec');
igual('nada esperando', 0, r.total);
igual('y se dice', false, r.hayAlgo);
ok('el texto no inventa trabajo',
   /No quedó nada esperando/.test(F.reporteDelDiaTexto(r)));

console.log('\n── Lo que ella corrigió a mano manda ──');
/**
 * Si la transportadora dice «en tránsito» y ella marcó «entregado», el
 * reporte tiene que creerle a ella. Es la misma precedencia del resto
 * de Nova.
 */
sembrar([
  ped({ id: 'p1', tienda: 'ec', fecha: '2026-09-23', cliente: 'Sofía', valor: 50,
        estado_canonico: 'novedad', estado_nova: 'entregado' }),
]);
r = F.reporteDelDia('cli', 'ec');
igual('no queda como novedad abierta', 0, r.total);

console.log('\n── Quién puede ──');
sembrar([]);
ok('nadie pide una tienda que no es suya',
   F.apiReporteDia(DUENA, { tienda: 'gt' }).ok === false);
ok('y se puede pedir un día concreto',
   F.apiReporteDia(DUENA, { tienda: 'ec', dia: '2026-09-01' }).dia === '2026-09-01');
ok('una fecha inventada no rompe nada: cae en ayer',
   F.apiReporteDia(DUENA, { tienda: 'ec', dia: 'ayer' }).dia === '2026-09-23');

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
