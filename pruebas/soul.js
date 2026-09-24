/**
 * NovaSoul, sobre hojas falsas.
 *
 * Lo que se comprueba no es que corra: es que no mienta. Que una
 * operadora no entre. Que sin horas libres diga que faltan en vez de
 * pintar una semana holgada. Que las horas de un proyecto con horas
 * fijas no se cuenten dos veces. Y —la que más importa— que lo que
 * Soul le pasa a Central sean números y nunca el texto de una tarea.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

// ── Hojas falsas, por archivo ────────────────────────────────
const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
let ESCRITURAS = [];

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
        appendRow: (f) => { ESCRITURAS.push({ hoja: nombre, tipo: 'nueva', fila: f }); m.push(f); },
        deleteRow: (n) => { ESCRITURAS.push({ hoja: nombre, tipo: 'borrar', n }); m.splice(n - 1, 1); },
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
            ESCRITURAS.push({ hoja: nombre, tipo: 'editar', fila: v[0] });
            v.forEach((fila, i) => { m[f - 1 + i] = fila.slice(); });
          },
          setValue: () => {},
        }),
      };
    },
  };
}

const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
let HOY = '2026-09-23T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: (k, v) => { PROPS[k] = v; }, deleteProperty: (k) => { delete PROPS[k]; } }) };
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
global.Utilities = { sleep: () => {}, getUuid: () => 'uuid' + (++UUID) + '0000000',
  formatDate: (d, tz, pat) => {
    const iso = (d && d.getTime && d.getTime() === new Date(HOY).getTime())
      ? HOY : new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(pat)) return iso.slice(0, 19).replace('T', ' ');
    return iso;
  } };
// `ahoraISO` usa new Date(): se fija el reloj para que las pruebas no
// cambien de resultado al día siguiente.
const RealDate = Date;
global.Date = class extends RealDate {
  constructor(...a) { return a.length ? new RealDate(...a) : new RealDate(HOY); }
  static now() { return new RealDate(HOY).getTime(); }
};

(0, eval)(src + '\n;globalThis.__F = { soulHoy, soulPendienteGuardar, soulPendienteBorrar,' +
  ' soulHorasGuardar, soulMindlabGuardar, soulMindlabAPendientes, soulFinanzas,' +
  ' soulFijoGuardar, soulCargaPorTrabajo_, soulUrgencia_, centralMio, manejarCentral,' +
  ' MINDLAB_PLAN, MINDLAB_INICIO, SOUL_CATEGORIAS };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

// ── Las hojas ────────────────────────────────────────────────
const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota'];
const C_TRAB = ['id','nombre','contraparte','tipo','estado','moneda','valor_acordado',
  'forma_cobro','fecha_inicio','fecha_entrega','horas_semana','especificacion',
  'documento','nota'];
const C_COBR = ['id','trabajo_id','concepto','monto','moneda','fecha_esperada',
  'fecha_cobrada','estado','nota'];
const C_FIN = ['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta',
  'recurrente','trabajo_id','nota'];
const C_META = ['id','tipo','nombre','con_quien','monto_meta','saldo','moneda','cuota',
  'dia_del_mes','fecha_meta','estado','nota'];
const C_HORAS = ['usuario_id','dia_semana','horas_libres','nota'];
const C_ML = ['id','usuario_id','semana','mes','tema','tarea','horas_estimadas',
  'desde','hasta','estado','nota'];
const C_FIJ = ['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes',
  'activo','nota'];

const YO = 'manuela@nova.com';
const p = (id, texto, o) => {
  const f = C_PEND.map(() => '');
  f[0] = id; f[1] = YO; f[2] = texto;
  f[5] = o.fecha || ''; f[9] = o.trabajo || ''; f[10] = o.estado || 'pendiente';
  f[11] = o.prioridad || 'media'; f[12] = o.horas || 0; f[14] = o.riesgo || 'corrible';
  f[7] = o.hechoEn || '';
  return f;
};

function sembrarHojas() {
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

  ESCRITURAS = [];
  LIBROS.cen = {
    Trabajos: [C_TRAB,
      // PHH: 12 horas fijas por semana
      ['t1','PHH','Upwork','empleo','activo','USD',0,'','', '',12,'','',''],
      // Son de Sky: proyecto con entrega, sin horas fijas
      ['t2','Son de Sky','Salsabor','cliente','activo','COP',2400000,'','','2026-09-30',0,'','',''],
      ['t3','Universidad','U','estudio','activo','',0,'','','2026-09-20',6,'','',''],
    ],
    Cobros: [C_COBR,
      ['c1','t2','Segundo pago',800000,'COP','2026-09-05','','pendiente',''],
    ],
    Finanzas: [C_FIN,
      ['f1','2026-09-02','gasto','arriendo','Arriendo',1250000,'COP','Bancolombia','','',''],
      ['f2','2026-09-10','gasto','mercado','Mercado',420000,'COP','Bancolombia','','',''],
    ],
    Metas: [C_META],
    Movimientos: [['fecha','quien','entidad','id','campo','antes','ahora']],
    Clientes: [['id','empresa','pais','plan','tarifa','costo','estado','fecha_alta',
                'fecha_corte','ultimo_pago','tickets_mes','usuarios','tiendas','sheet_id']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND,
      p('p1', 'Entregar la carta de cocteles', { fecha: '2026-09-25', trabajo: 't2', horas: 6, riesgo: 'acordado' }),
      p('p2', 'Parcial de Estadística', { fecha: '2026-09-20', trabajo: 't3', horas: 4, riesgo: 'inamovible', prioridad: 'alta' }),
      p('p3', 'Encargo reservado de PHH', { fecha: '2026-09-24', trabajo: 't1', horas: 5 }),
      p('p4', 'Comprar Omega 3', { fecha: '2026-09-26', horas: 1 }),
      p('p5', 'Llamar a mamá', { fecha: '2026-09-21', estado: 'hecho', hechoEn: '2026-09-21', horas: 0 }),
      p('p6', 'Algo viejo ya hecho', { fecha: '2026-01-02', estado: 'hecho', hechoEn: '2026-01-02' }),
    ],
    Horas: [C_HORAS],
    Mindlab: [C_ML],
    Fijos: [C_FIJ,
      ['x1', YO, 'arriendo', 'Apartamento', 1250000, 'COP', 5, 'si', ''],
      ['x2', YO, 'ahorro', 'Colchón', 300000, 'COP', 30, 'si', ''],
    ],
    Usuarios: [['id','nombre','correo']],
    Dias: [['usuario_id','fecha']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}

const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const OPERADORA = { correo: 'kat@nova.com', nombre: 'Kat', rol: 'operadora' };

console.log('\n── La puerta ──');
sembrarHojas();
ok('una operadora no entra a NovaSoul', F.soulHoy(OPERADORA, {}).ok === false);
ok('y el mensaje no le inventa otra razón',
   /de Manuela/.test(F.soulHoy(OPERADORA, {}).error));
ok('tampoco entra por el enrutador con la acción a mano',
   F.manejarCentral('nc_soul', { token: 'x' }).ok === false);
ok('la socia sí entra', F.soulHoy(SOCIA, {}).ok === true);

console.log('\n── El día, la semana y el mes ──');
let r = F.soulHoy(SOCIA, {});
igual('hoy', '2026-09-23', r.hoy);
igual('la semana va de lunes a domingo', ['2026-09-21', '2026-09-27'], [r.semana.lunes, r.semana.domingo]);
igual('siete días', 7, r.semana.dias.length);
ok('marca cuál es hoy', r.semana.dias.filter(d => d.esHoy).length === 1);
igual('el parcial está vencido', 3, r.pendientes.filter(x => x.id === 'p2')[0].dias);
igual('una entrega futura da días negativos', -2, r.pendientes.filter(x => x.id === 'p1')[0].dias);
igual('un vencido se cuenta una sola vez', 1, r.resumen.dia.vencidas);
igual('lo de esta semana: 3 entregas con fecha', 3, r.resumen.semana.entregas);
ok('el parcial vencido NO se cuela en la semana: vencio la pasada',
   !r.semana.dias.filter(d => d.fecha === '2026-09-20').length);

ok('lo hecho hace nueve meses no viaja a la pantalla',
   !r.pendientes.filter(x => x.id === 'p6').length);
ok('lo hecho esta semana sí',
   !!r.pendientes.filter(x => x.id === 'p5').length);

console.log('\n── Lo más urgente ──');
igual('gana lo vencido, no lo de mañana', 'p2', r.urgente.id);
ok('una tarea sin fecha nunca gana el primer puesto',
   F.soulUrgencia_({ dias: null, prioridad: 'alta', riesgo: 'inamovible' }) <
   F.soulUrgencia_({ dias: 0, prioridad: 'baja', riesgo: 'corrible' }));

console.log('\n── El riesgo de la semana ──');
ok('sin horas libres NO calcula: dice que faltan', r.riesgo.sinHoras === true);
igual('y no inventa un sobrante', null, r.riesgo.sobra);
ok('explica por qué falta, en vez de dejar el hueco mudo',
   /cuántas horas libres/.test(r.riesgo.porque));
igual('sin horas no propone candidatas a caerse', 0, r.riesgo.candidatas.length);

// Ahora sí: 4 horas libres de lunes a viernes, nada el fin de semana
let g = F.soulHorasGuardar(SOCIA, { dias: { 1: 4, 2: 4, 3: 4, 4: 4, 5: 4, 6: 0, 7: 0 } });
ok('guardar las horas devuelve ok', g.ok === true);
r = F.soulHoy(SOCIA, {});
igual('20 horas libres esta semana', 20, r.riesgo.libres);
igual('18 horas fijas: PHH 12 + universidad 6', 18, r.riesgo.fijas);
/**
 * La carta (6 h) y el Omega (1 h) suman 7: el encargo de PHH NO suma,
 * porque sus horas ya están dentro de las 12 fijas del proyecto. El
 * parcial tampoco: la universidad tiene 6 horas fijas.
 */
igual('las horas de un proyecto con horas fijas no se cuentan dos veces', 7, r.riesgo.extra);
igual('dos entregas quedaron dentro de horas ya contadas', 2, r.riesgo.dentroDeFijas);
igual('comprometidas = 18 + 7', 25, r.riesgo.comprometidas);
igual('faltan 5 horas', -5, r.riesgo.sobra);
ok('propone qué se puede correr', r.riesgo.candidatas.length > 0);
ok('y NUNCA propone correr el parcial',
   !r.riesgo.candidatas.filter(c => c.id === 'p2').length,
   JSON.stringify(r.riesgo.candidatas.map(c => c.id)));
ok('empieza por lo menos urgente', r.riesgo.candidatas[0].id === 'p4',
   JSON.stringify(r.riesgo.candidatas.map(c => c.id)));
/**
 * Y no ofrece correr lo de PHH: sus horas ya estaban dentro de las doce
 * fijas, así que moverlo no libera nada. Cuando lo que se puede correr
 * no alcanza, lo dice en vez de dejarla creer que la semana ya cabe.
 */
ok('no propone correr algo cuyas horas nunca se sumaron',
   !r.riesgo.candidatas.filter(c => c.id === 'p3').length,
   JSON.stringify(r.riesgo.candidatas.map(c => c.id)));
igual('las dos candidatas alcanzan a cubrir el hueco', ['p4', 'p1'],
      r.riesgo.candidatas.map(c => c.id));
igual('así que no queda nada por explicar', 0, r.riesgo.noAlcanza);

/**
 * Y la semana en que ni corriéndolo todo alcanza: 2 horas libres al
 * día contra 18 fijas. Ahí correr entregas no arregla nada, y decir
 * «corre estas dos» sería mandarla a hacer un sacrificio inútil.
 */
F.soulHorasGuardar(SOCIA, { dias: { 1: 2, 2: 2, 3: 2, 4: 2, 5: 2, 6: 0, 7: 0 } });
r = F.soulHoy(SOCIA, {});
igual('10 horas libres', 10, r.riesgo.libres);
igual('faltan 15', -15, r.riesgo.sobra);
igual('corriendo todo lo corrible siguen faltando 8', 8, r.riesgo.noAlcanza);
ok('y la razón apunta a las horas fijas, no a una entrega',
   /horas fijas/.test(r.riesgo.porque), r.riesgo.porque);

console.log('\n── La frontera con Central ──');
const carga = F.soulCargaPorTrabajo_(YO);
igual('Central sabe cuántas entregas abiertas tiene PHH', 1, carga.t1.abiertas);
igual('y cuántas horas suman', 5, carga.t1.horas);
const texto = JSON.stringify(carga);
ok('pero NO ve el texto de una sola tarea',
   texto.indexOf('reservado') === -1 && texto.indexOf('Parcial') === -1, texto);
const mio = F.centralMio(SOCIA, {});
ok('Central pinta el peso en cada proyecto',
   mio.trabajos.filter(t => t.id === 't1')[0].tareas.horas === 5);
ok('y el texto tampoco llega por ahí',
   JSON.stringify(mio).indexOf('reservado') === -1 &&
   JSON.stringify(mio).indexOf('Parcial') === -1);

console.log('\n── Crear, mover y borrar ──');
sembrarHojas();
let n = F.soulPendienteGuardar(SOCIA, { datos: {
  texto: 'Cotizar la web de Son de Sky', trabajo_id: 't2', fecha: '2026-09-29',
  horas: 3, prioridad: 'alta', riesgo: 'acordado' } });
ok('crear devuelve el id nuevo', n.ok === true && !!n.id);
ok('una tarea sin texto no se crea',
   F.soulPendienteGuardar(SOCIA, { datos: { horas: 2 } }).ok === false);
r = F.soulHoy(SOCIA, {});
const nueva = r.pendientes.filter(x => x.id === n.id)[0];
ok('y aparece con su proyecto ya resuelto', nueva && nueva.proyecto === 'Son de Sky',
   JSON.stringify(nueva));

F.soulPendienteGuardar(SOCIA, { datos: { id: 'p4', estado: 'haciendo' } });
r = F.soulHoy(SOCIA, {});
igual('mover a «haciendo» se guarda', 'haciendo', r.pendientes.filter(x => x.id === 'p4')[0].estado);

F.soulPendienteGuardar(SOCIA, { datos: { id: 'p4', estado: 'hecho' } });
r = F.soulHoy(SOCIA, {});
const p4 = r.pendientes.filter(x => x.id === 'p4')[0];
igual('marcar hecho pone la fecha sola', '2026-09-23', p4.hechoEn);
igual('y la saca de lo abierto', 2, r.resumen.semana.entregas);

F.soulPendienteGuardar(SOCIA, { datos: { id: 'p4', estado: 'pendiente' } });
r = F.soulHoy(SOCIA, {});
igual('y al devolverla, la fecha de hecho se limpia', '',
      r.pendientes.filter(x => x.id === 'p4')[0].hechoEn);

ok('editar el texto no borra las horas',
   F.soulPendienteGuardar(SOCIA, { datos: { id: 'p1', texto: 'Carta de cocteles v2' } }).ok &&
   F.soulHoy(SOCIA, {}).pendientes.filter(x => x.id === 'p1')[0].horas === 6);

ok('borrar borra', F.soulPendienteBorrar(SOCIA, { id: 'p4' }).ok === true);
ok('borrar dos veces avisa que ya no está',
   F.soulPendienteBorrar(SOCIA, { id: 'p4' }).ok === false);
ok('una operadora no borra nada',
   F.soulPendienteBorrar(OPERADORA, { id: 'p1' }).ok === false);

console.log('\n── Mindlab, recortado al trimestre ──');
sembrarHojas();
r = F.soulHoy(SOCIA, {});
igual('doce semanas', 12, r.mindlab.plan.length);
igual('arranca el lunes 28 de septiembre', '2026-09-28', r.mindlab.plan[0].desde);
igual('y cierra el domingo 20 de diciembre', '2026-12-20', r.mindlab.plan[11].hasta);
ok('deja libres las dos últimas semanas del año',
   r.mindlab.fin < '2026-12-21', r.mindlab.fin);
igual('37 horas en total: tres por semana', 37, r.mindlab.horasTotales);
ok('ninguna semana pasa de cuatro horas',
   r.mindlab.plan.every(m => m.horas <= 4));
ok('las dos de diciembre son las más livianas',
   r.mindlab.plan[10].horas === 2 && r.mindlab.plan[11].horas === 2);
ok('cada semana trae UNA tarea', r.mindlab.plan.every(m => !!m.tarea && !/·/.test(m.tarea)));

ok('marcar una semana se guarda',
   F.soulMindlabGuardar(SOCIA, { semana: 1, estado: 'hecho' }).ok === true);
igual('y se cuenta', 1, F.soulHoy(SOCIA, {}).mindlab.hechas);
ok('una semana que no existe se rechaza',
   F.soulMindlabGuardar(SOCIA, { semana: 99, estado: 'hecho' }).ok === false);

ok('bajar una semana a pendientes funciona',
   F.soulMindlabAPendientes(SOCIA, { semana: 2 }).ok === true);
ok('y no se puede bajar dos veces',
   F.soulMindlabAPendientes(SOCIA, { semana: 2 }).ok === false);
const ml = F.soulHoy(SOCIA, {}).pendientes.filter(x => x.origen === 'mindlab')[0];
ok('la tarea de Mindlab entra con sus horas y su fecha',
   ml && ml.horas === 4 && ml.fecha === '2026-10-11', JSON.stringify(ml));

console.log('\n── Finanzas: el plan contra lo que pasó ──');
sembrarHojas();
const fin = F.soulFinanzas(SOCIA, {});
igual('están las nueve categorías que ella nombró', 9, fin.categorias.length);
igual('en su orden', ['arriendo','mercado','servicios','internet','credito','deudas',
  'movil','varios','ahorro'], fin.categorias.map(c => c.id));
const arr = fin.categorias.filter(c => c.id === 'arriendo')[0];
igual('el arriendo tiene plan', { COP: 1250000 }, arr.planeado);
igual('y lo que de verdad salió', { COP: 1250000 }, arr.gastado);
const mer = fin.categorias.filter(c => c.id === 'mercado')[0];
igual('el mercado no tiene plan: se dice, no se pone cero', null, mer.planeado);
igual('pero sí salió plata', { COP: 420000 }, mer.gastado);
igual('el total del plan NO incluye el ahorro', { COP: 1250000 }, fin.totales.plan);
igual('el ahorro va aparte', { COP: 300000 }, fin.totales.ahorro);

ok('se puede cambiar un gasto fijo',
   F.soulFijoGuardar(SOCIA, { datos: { id: 'x1', monto: 1400000 } }).ok === true);
igual('y queda cambiado', { COP: 1400000 },
      F.soulFinanzas(SOCIA, {}).categorias.filter(c => c.id === 'arriendo')[0].planeado);
ok('se puede agregar uno nuevo',
   F.soulFijoGuardar(SOCIA, { datos: { categoria: 'internet', concepto: 'Claro', monto: 95000 } }).ok);
ok('una categoría inventada se rechaza',
   F.soulFijoGuardar(SOCIA, { datos: { categoria: 'cripto', monto: 1 } }).ok === false);
ok('una operadora no ve ni toca las finanzas',
   F.soulFinanzas(OPERADORA, {}).ok === false &&
   F.soulFijoGuardar(OPERADORA, { datos: { categoria: 'movil', monto: 1 } }).ok === false);

console.log('\n── Cuando Central todavía no está ──');
sembrarHojas();
delete LIBROS.cen.Trabajos;
r = F.soulHoy(SOCIA, {});
ok('Soul abre igual, sin proyectos', r.ok === true && r.trabajos.length === 0);
ok('y las tareas siguen ahí', r.pendientes.length > 0);

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
