/**
 * Lo automático del lado de ELLA.
 *
 * Lo que se comprueba, en orden de lo que importa:
 *
 *   1· que un correo VACÍO no se mande — un aviso diario que dice
 *      «todo bien» deja de leerse, y entonces tampoco se lee el día
 *      que traía algo;
 *   2· que lo de PHH no salga de NovaSoul: del proyecto confidencial
 *      puede salir el nombre y las horas, nunca el texto de la tarea,
 *      porque un correo se reenvía y se imprime;
 *   3· que el lunes diga si la semana cabe, y si no cabe, qué se puede
 *      mover de verdad;
 *   4· que lo tarde y lo que está por vencerse no se junten en un solo
 *      total, que es la forma más fácil de esconder lo primero.
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
        setValues: (v) => { v.forEach((fila, i) => { m[f - 1 + i] = fila.slice(); }); },
        setValue: () => {},
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-21T12:00:00Z';   // un LUNES
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
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{"data":[]}' }) };

/** El buzón. Nada sale de aquí sin quedar anotado. */
let BUZON = [];
global.MailApp = { sendEmail: (m) => BUZON.push(m) };

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

(0, eval)(src + '\n;globalThis.__F = { sociasPlataforma_, soulLunes, soulLunesTexto,' +
  ' soulLunesHayAlgo_, centralDiario, centralDiarioArmar, centralDiarioTexto,' +
  ' soulHoy, centralMio, TRABAJOS, estadoAutomatico, autoFecha_ };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const YO = 'manuela@nova.com';
const C_TRAB = ['id','nombre','contraparte','tipo','estado','moneda','valor_acordado',
  'forma_cobro','fecha_inicio','fecha_entrega','horas_semana','especificacion',
  'documento','nota','mi_rol','modalidad','porcentaje','base_porcentaje',
  'cliente_id','tienda_id','confidencial'];
const C_COB = ['id','trabajo_id','concepto','monto','moneda','fecha_esperada',
  'fecha_cobrada','estado','nota'];
const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];

/** Una fila, con las columnas en su sitio y sin repetir el esquema cada vez. */
function fila(cols, o) { return cols.map(c => (o[c] === undefined ? '' : o[c])); }

function sembrar(opts) {
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

  opts = opts || {};
  BUZON = []; UUID = 0;
  LIBROS.cen = {
    Plataforma: [['id','nombre','correo','rol','estado','ultima_conexion','nota'],
      ['p1', 'Manuela', YO, 'socia', 'activa', '', ''],
      ['p2', 'Katherine', 'kat@nova.com', 'operadora', 'activa', '', ''],
      ['p3', 'Vieja socia', 'vieja@nova.com', 'socia', 'inactivo', '', '']],
    Trabajos: [C_TRAB].concat(opts.trabajos || []),
    Cobros: [C_COB].concat(opts.cobros || []),
    Finanzas: [['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta',
                'recurrente','trabajo_id','nota']],
    Metas: [['id','tipo','nombre','con_quien','monto_meta','saldo','moneda','cuota',
             'dia_del_mes','fecha_meta','estado','nota']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND].concat(opts.pendientes || []),
    Horas: [['usuario_id','dia_semana','horas_libres','nota']].concat(
      opts.horas === null ? []
        : [1,2,3,4,5,6,7].map(d => [YO, d, opts.horas === undefined ? 4 : opts.horas, ''])),
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas',
               'desde','hasta','estado','nota']],
    Rutina: [['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin','lugar',
              'trabajo_id','materia_id','paga_fija','moneda','desde','hasta','activo','nota']]
      .concat(opts.rutina || []),
    Turnos: [['id','usuario_id','rutina_id','fecha','paga','propinas','moneda','estado',
              'finanza_id','nota']],
    Carta: [['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota']],
    Transitos: [['usuario_id','fecha','casa','tema','intensidad_pct','texto_transito',
      'por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto','a_natal',
      'desde','hasta','fuente']],
    Pensum: [['id','usuario_id','desde','hasta','titulo','cuerpo','casa','momento',
              'que_pide','que_evitar','nota']],
    Revolucion: [['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota']],
    Usuarios: [['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
                'lugar_nacimiento','zona_horaria','acento','modo','idioma']],
    Materias: [['id','usuario_id','nombre','trabajo_id','estado']],
    Fijos: [['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes',
             'activo','nota']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}

// ─────────────────────────────────────────────────────────────
console.log('\n── A quién le escribe Nova ──');
sembrar({});
const socias = F.sociasPlataforma_();
igual('solo a las socias activas', [YO], socias.map(s => s.correo));
ok('y con rol de socia, que es lo que abre NovaSoul',
   socias.every(s => s.rol === 'socia'));

// ─────────────────────────────────────────────────────────────
console.log('\n── Correo vacío no se manda ──');
sembrar({});
F.soulLunes();
igual('una semana sin nada no produce correo del lunes', 0, BUZON.length);
F.centralDiario();
igual('y un día sin cobros ni entregas tampoco', 0, BUZON.length);

// ─────────────────────────────────────────────────────────────
console.log('\n── El lunes: ¿cabe la semana? ──');
/**
 * 28 h libres (4 por día). PHH se lleva 12 h fijas. Dos entregas
 * sueltas de 6 h cada una. 12 + 12 = 24 ≤ 28: cabe, sobran 4.
 */
sembrar({
  trabajos: [
    fila(C_TRAB, { id: 'phh', nombre: 'PHH', tipo: 'cliente', estado: 'activo',
                   horas_semana: 12, moneda: 'COP', confidencial: 'si',
                   fecha_entrega: '2026-09-23' }),
    fila(C_TRAB, { id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo',
                   moneda: 'COP', valor_acordado: 2000000, confidencial: 'no' }),
  ],
  pendientes: [
    fila(C_PEND, { id: 't1', usuario_id: YO, texto: 'Carta de cocteles',
                   fecha: '2026-09-24', estado: 'pendiente', horas_estimadas: 6,
                   trabajo_id: 'sky', riesgo: 'se_puede_mover', prioridad: 'media' }),
    fila(C_PEND, { id: 't2', usuario_id: YO, texto: 'Reels de octubre',
                   fecha: '2026-09-25', estado: 'pendiente', horas_estimadas: 6,
                   trabajo_id: 'sky', riesgo: 'se_puede_mover', prioridad: 'baja' }),
    /** Lo de PHH: entra en las 12 h fijas, NO se suma otra vez. */
    fila(C_PEND, { id: 't3', usuario_id: YO, texto: 'Informe reservado de PHH',
                   fecha: '2026-09-23', estado: 'pendiente', horas_estimadas: 8,
                   trabajo_id: 'phh', riesgo: 'se_puede_mover', prioridad: 'alta' }),
  ],
});
let h = F.soulHoy({ correo: YO, rol: 'socia' }, {});
let txt = F.soulLunesTexto(h, null);
ok('dice que cabe', /✓ CABE/.test(txt), txt);
ok('con las horas libres a la vista', /28 h libres/.test(txt), txt);
ok('y dice cuántas son horas fijas de proyectos', /12 h son horas fijas/.test(txt), txt);
ok('sí se manda: la semana tiene entregas', F.soulLunesHayAlgo_(h) === true);
F.soulLunes();
igual('llega un correo', 1, BUZON.length);
igual('a ella', YO, BUZON[0].to);
ok('con la fecha del lunes en el asunto', /2026-09-21/.test(BUZON[0].subject), BUZON[0].subject);

// ── Ahora que NO cabe ──
sembrar({
  horas: 2,   // 14 h libres
  trabajos: [
    fila(C_TRAB, { id: 'phh', nombre: 'PHH', tipo: 'cliente', estado: 'activo',
                   horas_semana: 12, moneda: 'COP', confidencial: 'si' }),
    fila(C_TRAB, { id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo',
                   moneda: 'COP' }),
    fila(C_TRAB, { id: 'uni', nombre: 'Universidad', tipo: 'estudio', estado: 'activo',
                   moneda: 'COP' }),
  ],
  pendientes: [
    fila(C_PEND, { id: 't1', usuario_id: YO, texto: 'Carta de cocteles',
                   fecha: '2026-09-24', estado: 'pendiente', horas_estimadas: 6,
                   trabajo_id: 'sky', riesgo: 'se_puede_mover', prioridad: 'baja' }),
    fila(C_PEND, { id: 't2', usuario_id: YO, texto: 'Parcial de Estadística',
                   fecha: '2026-09-25', estado: 'pendiente', horas_estimadas: 5,
                   trabajo_id: 'uni', riesgo: 'inamovible', prioridad: 'alta' }),
    fila(C_PEND, { id: 't3', usuario_id: YO, texto: 'Informe reservado de PHH',
                   fecha: '2026-09-23', estado: 'pendiente', horas_estimadas: 8,
                   trabajo_id: 'phh', riesgo: 'se_puede_mover', prioridad: 'alta' }),
  ],
});
h = F.soulHoy({ correo: YO, rol: 'socia' }, {});
txt = F.soulLunesTexto(h, null);
ok('dice que NO cabe', /✕ NO CABE/.test(txt), txt);
ok('ofrece mover lo que de verdad libera horas', /Carta de cocteles/.test(txt), txt);
/**
 * ── LAS DOS QUE NO SE PUEDEN OFRECER ──
 *
 * El parcial es inamovible: correrlo no es una opción, así que
 * ofrecerlo es ruido. Y lo de PHH ya está dentro de las 12 h fijas:
 * correrlo no libera ni una hora, pero le haría creer que la semana
 * ya cabe. Las dos son formas distintas de mentir con una lista.
 */
ok('NO ofrece correr el parcial', !/Parcial de Estad/.test(
   txt.split('Lo que se puede mover')[1] || ''), txt);
ok('NI lo que ya está dentro de las horas fijas',
   !/Informe reservado/.test(txt.split('Lo que se puede mover')[1] || ''), txt);

console.log('\n── Y lo de PHH no sale de NovaSoul ──');
/**
 * Un correo se reenvía y se imprime. Del proyecto confidencial puede
 * salir el nombre —eso es de Central— pero jamás el texto de la tarea.
 */
F.soulLunes();
ok('el correo del lunes no dice qué es lo de PHH',
   !/Informe reservado/.test(BUZON[0].body), BUZON[0].body);

// ── Sin horas no se finge una respuesta ──
sembrar({ horas: null,
  pendientes: [fila(C_PEND, { id: 't1', usuario_id: YO, texto: 'Algo',
    fecha: '2026-09-24', estado: 'pendiente', horas_estimadas: 6 })] });
h = F.soulHoy({ correo: YO, rol: 'socia' }, {});
txt = F.soulLunesTexto(h, null);
ok('sin horas libres dice NO SÉ, no un «cabe» inventado',
   /NO SÉ SI CABE/.test(txt) && !/CABE\./.test(txt.replace('NO SÉ SI CABE', '')), txt);
ok('y eso SÍ se manda: es lo que hay que arreglar', F.soulLunesHayAlgo_(h) === true);

// ─────────────────────────────────────────────────────────────
console.log('\n── Qué momento es, dentro del correo ──');
/**
 * El cielo es lo último del correo y es de adorno: si falla, el correo
 * del lunes tiene que salir igual. Una semana que no cabe no se puede
 * quedar sin avisar porque falte una carta astral.
 */
sembrar({
  pendientes: [fila(C_PEND, { id: 't1', usuario_id: YO, texto: 'Entregar algo',
    fecha: '2026-09-24', estado: 'pendiente', horas_estimadas: 3 })],
});
LIBROS.s.Carta.push([YO, 'ascendente', 'capricornio', 15, '', 'no', '']);
LIBROS.s.Carta.push([YO, 'sol', 'virgo', 27, 9, 'no', '']);
LIBROS.s.Usuarios.push(['u1', 'Manuela', YO, '1995-09-20', '13:20', 'Palmira',
                        'America/Bogota', '', '', '']);
F.soulLunes();
let cuerpo = BUZON[0].body;
ok('el correo dice qué momento es hoy', /QUÉ MOMENTO ES/.test(cuerpo), cuerpo);
ok('con la casa que rige el mes', /El mes: casa \d+/.test(cuerpo), cuerpo);
ok('y la que rige el año', /El año: casa \d+/.test(cuerpo), cuerpo);
/**
 * Contada, no inventada: nació el 20/09/1995 y hoy es 21/09/2026, o
 * sea que ya cumplió 31. (31 % 12) + 1 = casa 8, y con Ascendente
 * Capricornio la 8 cae en Leo. Es la misma cuenta que se comprueba
 * aparte en pruebas/cielo.js; aquí lo que se comprueba es que llegue
 * igual por correo.
 */
ok('contada de verdad, no inventada', /El año: casa 8 · Leo/.test(cuerpo), cuerpo);

BUZON = [];
delete LIBROS.s.Carta;
/**
 * Tocar la hoja por debajo es simular una edición FUERA de la app —
 * ella abriendo el Google Sheet a mano. Eso, en la vida real, pasa
 * entre dos peticiones, así que aquí hay que marcar el corte: Nova
 * suelta lo que tenía leído en memoria, igual que al empezar una
 * petición nueva.
 */
soulOlvidar_();
F.soulLunes();
igual('y sin carta el correo sale igual', 1, BUZON.length);
ok('con lo que importa, que es la semana', /TU SEMANA/.test(BUZON[0].body));
/**
 * Sin carta, la luna sigue hablando —no necesita carta— pero las
 * profecciones no, porque sin Ascendente no hay casas. Se calla esa
 * parte en vez de dar una casa cualquiera.
 */
cuerpo = BUZON[0].body;
ok('la luna sigue diciendo qué momento es', /Hoy: tiempo de/.test(cuerpo), cuerpo);
ok('pero sin Ascendente no se inventa la casa del año',
   !/El año: casa/.test(cuerpo), cuerpo);

// ─────────────────────────────────────────────────────────────
console.log('\n── Central, todos los días ──');
sembrar({
  trabajos: [
    fila(C_TRAB, { id: 'phh', nombre: 'PHH', tipo: 'cliente', estado: 'activo',
                   moneda: 'COP', horas_semana: 12, confidencial: 'si',
                   fecha_entrega: '2026-09-22' }),
    fila(C_TRAB, { id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo',
                   moneda: 'COP', fecha_entrega: '2026-12-01' }),
    fila(C_TRAB, { id: 'viejo', nombre: 'Cliente cerrado', tipo: 'cliente',
                   estado: 'cerrado', moneda: 'COP', fecha_entrega: '2026-09-22' }),
  ],
  pendientes: [
    fila(C_PEND, { id: 't3', usuario_id: YO, texto: 'Informe reservado de PHH',
                   fecha: '2026-09-22', estado: 'pendiente', horas_estimadas: 8,
                   trabajo_id: 'phh', prioridad: 'alta' }),
  ],
  cobros: [
    fila(C_COB, { id: 'c1', trabajo_id: 'sky', concepto: 'Primera cuota',
                  monto: 1000000, moneda: 'COP', fecha_esperada: '2026-09-10',
                  estado: 'pendiente' }),
    fila(C_COB, { id: 'c2', trabajo_id: 'phh', concepto: 'Mes de septiembre',
                  monto: 500, moneda: 'USD', fecha_esperada: '2026-09-23',
                  estado: 'pendiente' }),
    fila(C_COB, { id: 'c3', trabajo_id: 'sky', concepto: 'Segunda cuota',
                  monto: 1000000, moneda: 'COP', fecha_esperada: '2026-11-30',
                  estado: 'pendiente' }),
    fila(C_COB, { id: 'c4', trabajo_id: 'sky', concepto: 'Anticipo',
                  monto: 300000, moneda: 'COP', fecha_esperada: '2026-08-01',
                  fecha_cobrada: '2026-08-02', estado: 'cobrado' }),
  ],
});
const m = F.centralMio({ correo: YO, rol: 'socia' }, {});
const a = F.centralDiarioArmar(m);
igual('lo tarde va aparte de lo que está por vencerse',
      [['c1'], ['c2']], [a.atrasados.map(c => c.id), a.porVencer.map(c => c.id)]);
ok('lo que vence en noviembre no aparece todavía',
   !a.porVencer.filter(c => c.id === 'c3').length);
ok('lo ya cobrado no aparece nunca',
   !a.atrasados.concat(a.porVencer).filter(c => c.id === 'c4').length);
igual('de las entregas, solo la de mañana y no la de diciembre',
      ['phh'], a.entregas.map(t => t.id));
ok('un proyecto cerrado no entrega nada',
   !a.entregas.filter(t => t.id === 'viejo').length);

const ct = F.centralDiarioTexto(a);
ok('dice cuánto está atrasado, por moneda y sin sumarlas',
   /Total atrasado: 1\.000\.000 COP/.test(ct), ct);
ok('y cuántos días lleva tarde', /11 días tarde/.test(ct), ct);
/**
 * La pared, otra vez. El nombre del proyecto es de Central y sale; el
 * texto de la tarea es de NovaSoul y no sale — pero el número sí, para
 * que ella sepa que hay algo sin que el correo lo cuente.
 */
ok('del proyecto confidencial sale el nombre', /PHH/.test(ct), ct);
ok('y cuántas tareas abiertas tiene', /1 tarea abierta/.test(ct), ct);
ok('pero NUNCA qué dicen', !/Informe reservado/.test(ct), ct);
ok('y se dice por qué no están', /el detalle solo en NovaSoul/.test(ct), ct);

BUZON = [];
F.centralDiario();
igual('se manda un solo correo, a ella', [1, YO], [BUZON.length, BUZON[0].to]);
ok('el asunto dice cuántos cobros están tarde',
   /1 cobro\(s\) tarde/.test(BUZON[0].subject), BUZON[0].subject);

// ─────────────────────────────────────────────────────────────
console.log('\n── Quedan registrados como lo automático ──');
const reg = {};
F.TRABAJOS.forEach(t => { reg[t.fn] = t; });
ok('soulLunes está en la lista', !!reg.soulLunes);
igual('y corre los lunes', 'MONDAY', reg.soulLunes.dia);
ok('centralDiario está en la lista', !!reg.centralDiario);
ok('y es diario', !reg.centralDiario.dia);
ok('los dos dicen qué se pierde si no corren',
   !!reg.soulLunes.porque && !!reg.centralDiario.porque);
/** La consola de Central los muestra sin que nadie los agregue ahí. */
sembrar({});
const est = F.estadoAutomatico();
ok('la tarjeta de «lo automático» los muestra',
   est.trabajos.filter(t => /NovaSoul|Nova Central/.test(t.nombre)).length === 2,
   est.trabajos.map(t => t.nombre).join(' | '));
ok('y apagados, porque todavía nadie los prendió',
   est.trabajos.every(t => t.prendido === false) && est.todoPrendido === false);

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
