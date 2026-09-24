/**
 * EL PUENTE: los proyectos de Nova Central y los pendientes de NovaSoul.
 *
 * Ella preguntó: «a cada pendiente le puse el proyecto pero no sé si ya
 * hilan y toman info de Nova Central o qué».
 *
 * Es una pregunta justa, porque el puente es invisible: un proyecto vive
 * en el libro de Central y un pendiente en el de Soul, y lo único que
 * los une es un `trabajo_id` escrito en una celda. Si ese hilo se corta,
 * nada falla: simplemente el pendiente sale sin nombre de proyecto y las
 * horas fijas dejan de descontarse — y las dos pantallas siguen viéndose
 * perfectas mientras mienten.
 *
 * Esta prueba recorre el puente en LOS DOS SENTIDOS:
 *
 *   CENTRAL → SOUL   el pendiente muestra el nombre del proyecto, su
 *                    tipo, y sus horas fijas entran al cálculo de si la
 *                    semana cabe.
 *   SOUL → CENTRAL   el proyecto muestra cuántas tareas abiertas tiene y
 *                    cuántas horas — NÚMEROS, nunca el texto de la tarea.
 *
 * Y comprueba los tres modos de romperse en silencio:
 *   · un trabajo_id que ya no existe
 *   · un proyecto cerrado que sigue descontando horas
 *   · Central caída, con Soul teniendo que seguir funcionando
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libro(id) {
  const hojas = LIBROS[id];
  if (!hojas) throw new Error('No existe ese libro: ' + id);
  return {
    getSheetByName: (nombre) => {
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
  };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T12:00:00Z';   // jueves
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

(0, eval)(src + '\n;globalThis.__F = { soulHoy, soulPendienteGuardar, centralMio,' +
  ' soulTrabajos_, soulCargaPorTrabajo_, familiaDe_, PROY_FAMILIAS };');
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
const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];
function trab(o) { return C_TRAB.map(c => (o[c] === undefined ? '' : o[c])); }
function pend(o) { return C_PEND.map(c => (o[c] === undefined ? '' : o[c])); }

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
  UUID = 0;
  LIBROS.cen = {
    Trabajos: [C_TRAB].concat(opts.trabajos || []),
    Cobros: [['id','trabajo_id','concepto','monto','moneda','fecha_esperada',
              'fecha_cobrada','estado','nota']],
    Finanzas: [['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta',
                'recurrente','trabajo_id','nota']],
    Metas: [['id','tipo','nombre','con_quien','monto_meta','saldo','moneda','cuota',
             'dia_del_mes','fecha_meta','estado','nota']],
    Plataforma: [['id','nombre','correo','rol','estado','ultima_conexion','nota'],
                 ['p1','Manuela',YO,'socia','activa','','']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND].concat(opts.pendientes || []),
    Horas: [['usuario_id','dia_semana','horas_libres','nota']].concat(
      [1,2,3,4,5,6,7].map(d => [YO, d, opts.horas === undefined ? 4 : opts.horas, ''])),
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas',
               'desde','hasta','estado','nota']],
    Rutina: [['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin',
              'lugar','trabajo_id','materia_id','paga_fija','moneda','desde','hasta',
              'activo','nota']].concat(opts.rutina || []),
    Turnos: [['id','usuario_id','rutina_id','fecha','paga','propinas','moneda','estado',
              'finanza_id','nota']],
    Carta: [['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota']],
    Transitos: [['usuario_id','fecha','casa','casa_placidus','tema','intensidad_pct',
      'texto_transito','por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto',
      'a_natal','desde','hasta','fuente']],
    Pensum: [['id','usuario_id','desde','hasta','titulo','cuerpo','casa','casa_alterna',
              'momento','que_pide','que_evitar','nota']],
    Revolucion: [['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota']],
    Usuarios: [['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
                'lugar_nacimiento','zona_horaria','acento','modo','idioma']],
    Materias: [['id','usuario_id','nombre','trabajo_id','estado']],
    Fijos: [['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes',
             'activo','nota']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia',
                permisos: ['facturacion'] };

// ═════════════════════════════════════════════════════════════
console.log('\n── CENTRAL → SOUL: el pendiente sabe de qué proyecto es ──');
sembrar({
  trabajos: [
    trab({ id: 'phh', nombre: 'PHH', tipo: 'cliente', estado: 'activo',
           horas_semana: 12, moneda: 'COP', confidencial: 'si' }),
    trab({ id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo',
           moneda: 'COP', valor_acordado: 2000000, fecha_entrega: '2026-10-15' }),
    trab({ id: 'uni', nombre: 'Universidad', tipo: 'estudio', estado: 'activo',
           moneda: 'COP' }),
  ],
  pendientes: [
    pend({ id: 't1', usuario_id: YO, texto: 'Carta de cocteles', fecha: '2026-09-25',
           estado: 'pendiente', horas_estimadas: 6, trabajo_id: 'sky',
           prioridad: 'alta', riesgo: 'se_puede_mover' }),
    pend({ id: 't2', usuario_id: YO, texto: 'Parcial de Estadística', fecha: '2026-09-26',
           estado: 'pendiente', horas_estimadas: 5, trabajo_id: 'uni',
           prioridad: 'alta', riesgo: 'inamovible' }),
    pend({ id: 't3', usuario_id: YO, texto: 'Informe reservado', fecha: '2026-09-25',
           estado: 'pendiente', horas_estimadas: 8, trabajo_id: 'phh',
           prioridad: 'alta' }),
    pend({ id: 't4', usuario_id: YO, texto: 'Suelta, sin proyecto', fecha: '2026-09-25',
           estado: 'pendiente', horas_estimadas: 2 }),
  ],
});

let h = F.soulHoy(SOCIA, {});
ok('NovaSoul abre', h.ok === true, JSON.stringify(h).slice(0, 200));

const P = {}; h.pendientes.forEach(x => { P[x.id] = x; });
igual('el pendiente trae el NOMBRE del proyecto, no su id',
      'Son de Sky', P.t1.proyecto);
igual('y su tipo, que viene de Central', 'cliente', P.t1.tipo);
igual('el de la universidad también', ['Universidad', 'estudio'],
      [P.t2.proyecto, P.t2.tipo]);
igual('el suelto no inventa proyecto', '', P.t4.proyecto);

console.log('\n── La lista de proyectos que ve NovaSoul ──');
igual('son los tres de Central', ['PHH', 'Son de Sky', 'Universidad'],
      h.trabajos.map(t => t.nombre).sort());
ok('con sus horas fijas', h.trabajos.filter(t => t.id === 'phh')[0].horasSemana === 12);
ok('y su fecha de entrega',
   h.trabajos.filter(t => t.id === 'sky')[0].entrega === '2026-10-15');

console.log('\n── Las horas fijas de Central entran al cálculo ──');
/**
 * 28 h libres. PHH tiene 12 h fijas en Central. Las tareas sueltas de la
 * semana suman: cocteles 6 + parcial 5 + suelta 2 = 13. Lo de PHH NO se
 * suma: ya está dentro de sus 12 h fijas.
 */
igual('las horas fijas salen de los proyectos activos', 12, h.riesgo.fijas);
igual('lo demás se suma aparte', 13, h.riesgo.extra);
igual('comprometidas = fijas + extra', 25, h.riesgo.comprometidas);
igual('y la tarea de PHH NO se contó dos veces', 1, h.riesgo.dentroDeFijas);
ok('la semana cabe', h.riesgo.sobra === 3, JSON.stringify(h.riesgo));

console.log('\n── El mes, agrupado por proyecto ──');
const porP = {}; h.resumen.mes.porProyecto.forEach(x => { porP[x.nombre] = x; });
igual('cada proyecto con sus horas', 6, porP['Son de Sky'].horas);
igual('y el suelto dice que no tiene proyecto', 2, porP['Sin proyecto'].horas);

// ═════════════════════════════════════════════════════════════
console.log('\n── SOUL → CENTRAL: el proyecto sabe cuánto le deben ──');
const m = F.centralMio(SOCIA, {});
ok('Nova Central abre', m.ok === true, JSON.stringify(m).slice(0, 200));
const T = {}; m.trabajos.forEach(x => { T[x.id] = x; });
igual('Son de Sky tiene 1 tarea abierta y 6 horas',
      [1, 6], [T.sky.tareas.abiertas, T.sky.tareas.horas]);
igual('la universidad, 1 y 5', [1, 5], [T.uni.tareas.abiertas, T.uni.tareas.horas]);
igual('PHH, 1 y 8', [1, 8], [T.phh.tareas.abiertas, T.phh.tareas.horas]);
igual('y con la próxima entrega', '2026-09-25', T.sky.tareas.proxima);

console.log('\n── La pared: Central NUNCA ve el texto de la tarea ──');
/**
 * Lo de PHH es confidencial. Central puede saber que hay una tarea
 * abierta y cuántas horas cuesta; jamás qué dice. Un correo se reenvía
 * y Central se abre delante de un contador.
 */
const crudo = JSON.stringify(m);
ok('«Informe reservado» no está en ninguna parte de Central',
   crudo.indexOf('Informe reservado') === -1);
ok('ni el texto de las otras tareas',
   crudo.indexOf('Carta de cocteles') === -1 &&
   crudo.indexOf('Parcial de Estadística') === -1);
ok('pero sí se sabe que PHH es confidencial', T.phh.confidencial === true);

// ═════════════════════════════════════════════════════════════
console.log('\n── Las tres formas de romperse en silencio ──');

/** 1 · Un trabajo_id que ya no existe en Central. */
sembrar({
  trabajos: [trab({ id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo' })],
  pendientes: [
    pend({ id: 't1', usuario_id: YO, texto: 'De un proyecto borrado',
           fecha: '2026-09-25', estado: 'pendiente', horas_estimadas: 3,
           trabajo_id: 'ya-no-existe' }),
  ],
});
h = F.soulHoy(SOCIA, {});
ok('un proyecto borrado no tumba la pantalla', h.ok === true);
igual('el pendiente queda sin nombre de proyecto, no con un id crudo',
      '', h.pendientes[0].proyecto);
igual('pero conserva el id, para poder arreglarlo', 'ya-no-existe',
      h.pendientes[0].trabajoId);
igual('y sus horas SÍ se cuentan: el trabajo existe aunque el proyecto no',
      3, h.riesgo.extra);

/** 2 · Un proyecto cerrado no puede seguir descontando horas. */
sembrar({
  trabajos: [
    trab({ id: 'viejo', nombre: 'Cliente que se fue', tipo: 'cliente',
           estado: 'cerrado', horas_semana: 20 }),
    trab({ id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo',
           horas_semana: 5 }),
  ],
  pendientes: [],
});
h = F.soulHoy(SOCIA, {});
igual('solo el activo aporta horas fijas', 5, h.riesgo.fijas);
ok('el cerrado no se cuenta',
   h.riesgo.fijas !== 25, 'fijas=' + h.riesgo.fijas);

/** 3 · Central caída: Soul tiene que seguir funcionando. */
sembrar({ pendientes: [
  pend({ id: 't1', usuario_id: YO, texto: 'Algo mío', fecha: '2026-09-25',
         estado: 'pendiente', horas_estimadas: 3 }) ] });
delete LIBROS.cen.Trabajos;
h = F.soulHoy(SOCIA, {});
ok('sin la hoja de Trabajos, NovaSoul abre igual', h.ok === true);
igual('sin proyectos, pero con los pendientes', 1, h.pendientes.length);
igual('y sin horas fijas inventadas', 0, h.riesgo.fijas);

// ═════════════════════════════════════════════════════════════
console.log('\n── Y al guardar un pendiente desde la pantalla ──');
sembrar({
  trabajos: [trab({ id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo' })],
});
let g = F.soulPendienteGuardar(SOCIA, { datos: {
  texto: 'Nueva tarea', fecha: '2026-09-28', horas: 4, trabajo_id: 'sky',
  prioridad: 'alta' } });
ok('se guarda', g.ok === true, JSON.stringify(g));
h = F.soulHoy(SOCIA, {});
igual('y ya sale con su proyecto', 'Son de Sky', h.pendientes[0].proyecto);
igual('Central lo ve al instante', 1,
      F.centralMio(SOCIA, {}).trabajos[0].tareas.abiertas);

/**
 * ── LA TRAMPA DE LOS DOS NOMBRES ──
 *
 * La hoja guarda `trabajo_id`, pero todo lo que Nova DEVUELVE lo llama
 * `trabajoId`. Escribiendo esta prueba caí yo mismo: mandé `trabajoId`
 * y la tarea quedó sin proyecto sin que nada fallara. Ahora se aceptan
 * los dos, porque un enlace que se pierde en silencio no se nota hasta
 * que las horas fijas dejan de cuadrar.
 */
sembrar({
  trabajos: [trab({ id: 'sky', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo' })],
});
F.soulPendienteGuardar(SOCIA, { datos: {
  texto: 'Mandada con trabajoId', fecha: '2026-09-28', horas: 4,
  trabajoId: 'sky' } });
igual('mandar trabajoId también enlaza', 'Son de Sky',
      F.soulHoy(SOCIA, {}).pendientes[0].proyecto);

/** Un id que no existe se guarda igual: se podrá arreglar, no se pierde. */
sembrar({ trabajos: [] });
F.soulPendienteGuardar(SOCIA, { datos: {
  texto: 'De un proyecto por crear', fecha: '2026-09-28', trabajo_id: 'futuro' } });
igual('un proyecto que todavía no existe no se descarta', 'futuro',
      F.soulHoy(SOCIA, {}).pendientes[0].trabajoId);


// ═════════════════════════════════════════════════════════════
console.log('\n── Repartir lo que ya está cargado: trabajo o proyecto ──');
/**
 * Su regla, con sus palabras: «entre proyectos pueden entrar los que no
 * son pagos, como los de Nova; en trabajos todo lo que me da ingresos».
 * La línea es una sola: ¿entra plata o no?
 */
[[{ tipo: 'estudio', nombre: 'Universidad' }, 'proyecto', true, 'la universidad'],
 [{ tipo: 'propio', nombre: 'Nova' }, 'proyecto', true, 'lo propio'],
 [{ tipo: 'empleo', nombre: 'Un empleo' }, 'trabajo', true, 'un empleo'],
 [{ tipo: 'cliente', valor_acordado: 2000000 }, 'trabajo', true, 'un cliente con valor'],
 [{ tipo: 'cliente', porcentaje: 50 }, 'trabajo', true, 'un cliente por porcentaje'],
 [{ tipo: 'cliente', modalidad: 'por_hora' }, 'trabajo', true, 'un cliente por hora'],
 [{ tipo: 'cliente', modalidad: 'sin_cobro' }, 'proyecto', true, 'un cliente que no cobra'],
].forEach(function (caso) {
  const f = F.familiaDe_(caso[0]);
  igual(caso[3] + ' → ' + caso[1], [caso[1], caso[2]], [f.familia, f.claro]);
});

console.log('\n── Y lo que NO está claro, se pregunta ──');
/**
 * Un «cliente» sin valor, sin porcentaje y sin forma de cobro puede ser
 * alguien que todavía no negoció precio, o un favor que nunca va a
 * pagar. Nova no puede saberlo, y adivinar mal cambia si esa fila suma
 * a lo que le deben. Lo fácil sería mandarlo a Trabajos y que ella lo
 * descubra un día mirando una cifra rara.
 */
const dudoso = F.familiaDe_({ tipo: 'cliente', nombre: 'Alguien' });
igual('un cliente sin nada de plata queda por preguntar', false, dudoso.claro);
ok('y se dice por qué, con palabras',
   /no sé si te va a pagar/.test(dudoso.porque), dudoso.porque);
const sinTipo = F.familiaDe_({ nombre: 'Sin tipo' });
igual('sin tipo, tampoco está claro', false, sinTipo.claro);
ok('mientras tanto NO desaparece de la lista', dudoso.familia === 'trabajo');

sembrar({
  trabajos: [
    trab({ id: 'phh', nombre: 'PHH', tipo: 'cliente', estado: 'activo',
           modalidad: 'por_hora', horas_semana: 12 }),
    trab({ id: 'nova', nombre: 'Nova', tipo: 'propio', estado: 'activo' }),
    trab({ id: 'uni', nombre: 'Universidad', tipo: 'estudio', estado: 'activo' }),
    trab({ id: 'raro', nombre: 'Alguien sin definir', tipo: 'cliente', estado: 'activo' }),
  ],
});
const mm = F.centralMio(SOCIA, {});
const fam = {}; mm.trabajos.forEach(x => { fam[x.id] = x.familia; });
igual('en la pantalla, cada uno en su familia',
      ['trabajo', 'proyecto', 'proyecto', 'trabajo'],
      [fam.phh, fam.nova, fam.uni, fam.raro]);
igual('la universidad va marcada aparte', true,
      mm.trabajos.filter(x => x.id === 'uni')[0].universidad);
igual('y solo el dudoso entra en la lista de preguntas',
      ['Alguien sin definir'], mm.porClasificar.map(x => x.nombre));
ok('con su porqué, para poder responder sin adivinar',
   /no sé si te va a pagar/.test(mm.porClasificar[0].porque));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
