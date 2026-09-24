/**
 * Turnos, clases y la plata chiquita.
 *
 * Lo que se comprueba: que un turno de 18:00 a 02:00 dure ocho horas y
 * no menos dieciséis; que una clase del semestre pasado deje de ocupar
 * la semana; que las horas libres sean las útiles MENOS lo ocupado; que
 * corregir las propinas no cree un ingreso nuevo; y que a fin de mes lo
 * que ya pasó no se mezcle con lo que falta por pasar.
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
// Miércoles 23 de septiembre de 2026
const HOY = '2026-09-23T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: (id) => libro(id), flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' }, WeekDay: { MONDAY: 'MONDAY' } };
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

(0, eval)(src + '\n;globalThis.__F = { duracionHoras_, horaNum_, rutinaCorre_, rutinaDe_,' +
  ' diasDeRutina_,' +
  ' soulRutina, soulRutinaGuardar, soulRutinaBorrar, soulTurnoGuardar, turnosPendientes_,' +
  ' soulHormigaGuardar, soulPlata, soulHoy, soulHorasGuardar, soulMes_, SOUL_HORMIGA };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];
const C_RUT = ['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin',
  'lugar','trabajo_id','materia_id','paga_fija','moneda','desde','hasta','activo','nota'];
const C_TUR = ['id','usuario_id','rutina_id','fecha','paga','propinas','moneda',
  'estado','finanza_id','nota'];
const C_FIN = ['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta',
  'recurrente','trabajo_id','nota'];
const C_FIJ = ['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes','activo','nota'];
const C_TRAB = ['id','nombre','contraparte','tipo','estado','moneda','valor_acordado',
  'forma_cobro','fecha_inicio','fecha_entrega','horas_semana','especificacion','documento','nota'];
const C_HOR = ['usuario_id','dia_semana','horas_libres','nota'];
const YO = 'manuela@nova.com';

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

  UUID = 0;
  LIBROS.cen = {
    Trabajos: [C_TRAB, ['t9','Salsabor','Bar','empleo','activo','COP',0,'','','',0,'','','']],
    Cobros: [['id','trabajo_id','concepto','monto','moneda','fecha_esperada','fecha_cobrada','estado','nota']],
    Finanzas: [C_FIN],
    Metas: [['id','tipo','nombre']],
    Movimientos: [['fecha','quien','entidad','id','campo','antes','ahora']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND],
    Materias: [['id','usuario_id','nombre','codigo','profesor','carpeta','semestre','trabajo_id','estado','nota']],
    Horas: [C_HOR],
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas','desde','hasta','estado','nota']],
    Fijos: [C_FIJ, ['x1', YO, 'arriendo', 'Apto', 1200000, 'COP', 5, 'si', '']],
    Rutina: [C_RUT,
      // Viernes y sábado en Salsabor: 18:00 a 02:00, 80.000 el turno
      ['r1', YO, 'turno', 'Salsabor', 5, '18:00', '02:00', 'Salsabor', 't9', '', 80000, 'COP', '', '', 'si', ''],
      ['r2', YO, 'turno', 'Salsabor', 6, '18:00', '02:00', 'Salsabor', 't9', '', 80000, 'COP', '', '', 'si', ''],
      // Clase de estadística: lunes 7 a 9, hasta el 5 de diciembre
      ['r3', YO, 'clase', 'Estadística', 1, '07:00', '09:00', 'Bloque 3', '', '', 0, '', '2026-08-03', '2026-12-05', 'si', ''],
      // Una clase del semestre PASADO, que ya terminó
      ['r4', YO, 'clase', 'Cálculo (pasado)', 3, '07:00', '11:00', '', '', '', 0, '', '2026-02-01', '2026-06-10', 'si', ''],
    ],
    Turnos: [C_TUR],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const OPERADORA = { correo: 'k@nova.com', nombre: 'Kat', rol: 'operadora' };

console.log('\n── Las horas de un bloque ──');
igual('un turno normal', 4, F.duracionHoras_('14:00', '18:00'));
/**
 * El de Salsabor cruza la medianoche. Restar a secas daba −16, y la
 * semana salía con horas de SOBRA justo los días que trabaja de noche.
 */
igual('uno que cruza la medianoche dura 8, no menos 16', 8, F.duracionHoras_('18:00', '02:00'));
igual('con media hora', 2.5, F.duracionHoras_('09:00', '11:30'));
igual('una hora escrita mal no inventa duración', 0, F.duracionHoras_('', ''));
igual('25:00 no existe', 0, F.horaNum_('25:00'));

console.log('\n── Se repite solo ──');
sembrarHojas();
let r = F.soulRutina(SOCIA, {});
igual('cuatro bloques guardados', 4, r.rutinas.length);
igual('el viernes ocupa 8 horas', 8, r.ocupadasPorDia[5]);
igual('y el lunes 2, de clase', 2, r.ocupadasPorDia[1]);

const vie = r.rutinas.filter(x => x.dia === 5)[0];
ok('el viernes 25 de septiembre corre', F.rutinaCorre_(vie, '2026-09-25'));
ok('y el viernes 2 de octubre también, sin escribir nada',
   F.rutinaCorre_(vie, '2026-10-02'));
ok('pero el jueves no', !F.rutinaCorre_(vie, '2026-10-01'));

const clase = r.rutinas.filter(x => x.id === 'r3')[0];
ok('la clase de este semestre corre el lunes 28', F.rutinaCorre_(clase, '2026-09-28'));
ok('y deja de correr después del 5 de diciembre',
   !F.rutinaCorre_(clase, '2026-12-14'));
const vieja = r.rutinas.filter(x => x.id === 'r4')[0];
ok('la clase del semestre pasado YA NO ocupa la semana',
   !F.rutinaCorre_(vieja, '2026-09-23'));

console.log('\n── Las horas libres se calculan, no se escriben ──');
// 10 horas útiles de lunes a sábado, 6 el domingo
F.soulHorasGuardar(SOCIA, { dias: { 1: 10, 2: 10, 3: 10, 4: 10, 5: 10, 6: 10, 7: 6 } });
let h = F.soulHoy(SOCIA, {});
igual('el lunes: 10 útiles − 2 de clase = 8', 8, h.semana.dias[0].libres);
igual('y lo dice separado', [10, 2], [h.semana.dias[0].utiles, h.semana.dias[0].ocupadas]);
igual('el miércoles no descuenta la clase que ya terminó', 10, h.semana.dias[2].libres);
igual('el viernes: 10 − 8 del turno = 2', 2, h.semana.dias[4].libres);
igual('el domingo, sin nada, queda igual', 6, h.semana.dias[6].libres);
igual('la semana suma 48 libres: 66 útiles − 18 ocupadas', 48, h.riesgo.libres);
igual('y 18 horas ocupadas por turnos y clases', 18, h.riesgo.ocupadas);
ok('cada día trae sus bloques, con nombre y hora',
   h.semana.dias[4].bloques.length === 1 &&
   h.semana.dias[4].bloques[0].nombre === 'Salsabor' &&
   h.semana.dias[4].bloques[0].inicio === '18:00',
   JSON.stringify(h.semana.dias[4].bloques));

/**
 * Un día donde lo ocupado se come todas las horas no da negativo: da
 * cero. Una semana con horas negativas se sumaría mal y diría que cabe.
 */
F.soulHorasGuardar(SOCIA, { dias: { 5: 4 } });
h = F.soulHoy(SOCIA, {});
igual('si el turno se come el día, quedan 0 y no −4', 0, h.semana.dias[4].libres);
F.soulHorasGuardar(SOCIA, { dias: { 5: 10 } });

console.log('\n── Las propinas, al día siguiente ──');
sembrarHojas();
let pend = F.turnosPendientes_(YO, '2026-09-23');
// Hacia atrás: viernes 18, sábado 19; y viernes 11, sábado 12...
ok('aparecen los turnos que ya pasaron', pend.length >= 2, pend.length + '');
ok('el más reciente primero', pend[0].fecha === '2026-09-19', JSON.stringify(pend.slice(0, 2)));
ok('NO pide propinas de un turno que no ha pasado',
   !pend.filter(t => t.fecha > '2026-09-23').length);
ok('cada uno trae su paga fija', pend[0].paga === 80000 && pend[0].horas === 8,
   JSON.stringify(pend[0]));

let g = F.soulTurnoGuardar(SOCIA, { rutina: 'r2', fecha: '2026-09-19', propinas: 34000 });
igual('guardar suma base + propinas', 114000, g.total);
igual('y queda un solo movimiento en Nova Central', 1, LIBROS.cen.Finanzas.length - 1);
const mov = LIBROS.cen.Finanzas[1];
igual('como ingreso, con la fecha del turno',
      ['ingreso', '2026-09-19', 114000], [mov[2], mov[1], mov[5]]);
ok('la nota deja ver la base y las propinas por separado',
   /Base 80000 \+ propinas 34000/.test(mov[10]), mov[10]);

pend = F.turnosPendientes_(YO, '2026-09-23');
ok('ese turno ya no aparece pendiente',
   !pend.filter(t => t.fecha === '2026-09-19' && t.rutinaId === 'r2').length);

/**
 * Corregir las propinas NO puede crear un ingreso nuevo: sería cobrarse
 * el mismo turno dos veces y el mes cuadraría de más.
 */
g = F.soulTurnoGuardar(SOCIA, { rutina: 'r2', fecha: '2026-09-19', propinas: 40000 });
igual('corregir las propinas no crea otro ingreso', 1, LIBROS.cen.Finanzas.length - 1);
igual('actualiza el que ya estaba', 120000, LIBROS.cen.Finanzas[1][5]);
igual('y tampoco duplica la fila del turno', 1, LIBROS.s.Turnos.length - 1);

ok('un día sin propinas también se puede cerrar',
   F.soulTurnoGuardar(SOCIA, { rutina: 'r1', fecha: '2026-09-18', propinas: 0 }).ok);
ok('y deja de aparecer pendiente',
   !F.turnosPendientes_(YO, '2026-09-23')
     .filter(t => t.fecha === '2026-09-18' && t.rutinaId === 'r1').length);

ok('no se puede cerrar un turno del futuro',
   /todavía no ha pasado/.test(
     F.soulTurnoGuardar(SOCIA, { rutina: 'r1', fecha: '2026-10-30', propinas: 10 }).error));
ok('ni uno que no está en la rutina',
   F.soulTurnoGuardar(SOCIA, { rutina: 'zz', fecha: '2026-09-18', propinas: 10 }).ok === false);
ok('ni propinas negativas',
   F.soulTurnoGuardar(SOCIA, { rutina: 'r1', fecha: '2026-09-18', propinas: -5 }).ok === false);
ok('una operadora no toca nada de esto',
   F.soulTurnoGuardar(OPERADORA, { rutina: 'r1', fecha: '2026-09-18', propinas: 1 }).ok === false &&
   F.soulRutinaGuardar(OPERADORA, { datos: { nombre: 'x' } }).ok === false &&
   F.soulHormigaGuardar(OPERADORA, { categoria: 'bus', monto: 1 }).ok === false);

console.log('\n── Gastos hormiga ──');
sembrarHojas();
igual('las seis categorías que nombró',
      ['bus','transporte','uber','antojos','salidas','hormiga'],
      F.SOUL_HORMIGA.map(c => c.id));
ok('un bus se guarda en dos toques',
   F.soulHormigaGuardar(SOCIA, { categoria: 'bus', monto: 2900 }).ok);
ok('y toma la fecha de hoy sin preguntar',
   LIBROS.cen.Finanzas[1][1] === '2026-09-23', LIBROS.cen.Finanzas[1][1]);
ok('con un concepto por defecto, no vacío',
   LIBROS.cen.Finanzas[1][4] === 'Buses', LIBROS.cen.Finanzas[1][4]);
ok('sin monto no se guarda',
   /Cuánto fue/.test(F.soulHormigaGuardar(SOCIA, { categoria: 'uber', monto: 0 }).error));
ok('una categoría inventada se rechaza',
   F.soulHormigaGuardar(SOCIA, { categoria: 'cripto', monto: 1 }).ok === false);

F.soulHormigaGuardar(SOCIA, { categoria: 'antojos', concepto: 'Café', monto: 12000 });
F.soulHormigaGuardar(SOCIA, { categoria: 'salidas', monto: 65000 });
F.soulHormigaGuardar(SOCIA, { categoria: 'uber', monto: 18000, fecha: '2026-09-20' });

console.log('\n── ¿Sobra o falta a fin de mes? ──');
F.soulTurnoGuardar(SOCIA, { rutina: 'r2', fecha: '2026-09-19', propinas: 34000 });
let plata = F.soulPlata(SOCIA, {});
ok('la pantalla de plata abre', plata.ok === true);
const cop = plata.resumen.resultado.filter(x => x.moneda === 'COP')[0];
igual('entró la plata del turno', 114000, cop.entro);
igual('salió lo hormiga: 2.900 + 12.000 + 65.000 + 18.000', 97900, cop.salio);
igual('queda lo que queda, sin adornos', 16100, cop.queda);
/**
 * Y aparte lo que FALTA del mes: el arriendo que se paga el 5 ya pasó,
 * así que no cuenta; los turnos del viernes 25, sábado 26 —el 19 ya
 * está cobrado— sí.
 */
igual('el arriendo del 5 ya pasó: no se cuenta como pendiente', 0, cop.fijosPendientes);
igual('faltan dos turnos por trabajar', 2, plata.resumen.turnosPorVenir);
igual('que valen 160.000', 160000, cop.porVenir);
igual('proyectado a fin de mes: 16.100 + 160.000', 176100, cop.proyectado);

const horm = plata.resumen.hormigaPorCategoria;
ok('lo hormiga se desglosa por categoría',
   horm.filter(c => c.id === 'salidas')[0].monto.COP === 65000,
   JSON.stringify(horm.filter(c => c.monto)));
ok('y una categoría sin gastos dice null, no cero',
   horm.filter(c => c.id === 'transporte')[0].monto === null);
igual('el total hormiga del mes', { COP: 97900 }, plata.resumen.hormiga);

/**
 * El arriendo pendiente sí tiene que contar cuando todavía no se ha
 * pagado: se mueve al día 28 y reaparece.
 */
LIBROS.s.Fijos[1][6] = 28;
/**
 * Tocar la hoja por debajo es simular una edición FUERA de la app —
 * ella abriendo el Google Sheet a mano. Eso, en la vida real, pasa
 * entre dos peticiones, así que aquí hay que marcar el corte: Nova
 * suelta lo que tenía leído en memoria, igual que al empezar una
 * petición nueva.
 */
soulOlvidar_();
plata = F.soulPlata(SOCIA, {});
igual('un fijo que aún no se paga sí cuenta como pendiente', 1200000,
      plata.resumen.resultado.filter(x => x.moneda === 'COP')[0].fijosPendientes);
igual('y la proyección se pone en rojo', -1023900,
      plata.resumen.resultado.filter(x => x.moneda === 'COP')[0].proyectado);

console.log('\n── Borrar la rutina ──');
sembrarHojas();
F.soulTurnoGuardar(SOCIA, { rutina: 'r1', fecha: '2026-09-18', propinas: 20000 });
const b = F.soulRutinaBorrar(SOCIA, { id: 'r1' });
ok('avisa cuántos turnos ya trabajados quedan', b.ok && b.trabajados === 1, JSON.stringify(b));
igual('y NO borra la plata que ya entró', 1, LIBROS.cen.Finanzas.length - 1);
igual('ni el turno trabajado', 1, LIBROS.s.Turnos.length - 1);

console.log('\n── Validaciones al crear ──');
sembrarHojas();
ok('sin nombre no se crea',
   /nombre/.test(F.soulRutinaGuardar(SOCIA, { datos: { tipo: 'turno' } }).error));
ok('sin tipo tampoco',
   /turno, clase/.test(F.soulRutinaGuardar(SOCIA, { datos: { nombre: 'x' } }).error));
ok('sin día de la semana tampoco',
   /día de la semana/.test(F.soulRutinaGuardar(SOCIA,
     { datos: { nombre: 'x', tipo: 'turno' } }).error));
ok('sin horas tampoco',
   /cuánto ocupa/.test(F.soulRutinaGuardar(SOCIA,
     { datos: { nombre: 'x', tipo: 'turno', dia_semana: 2 } }).error));
ok('con las fechas al revés, tampoco',
   /va después/.test(F.soulRutinaGuardar(SOCIA, { datos: { nombre: 'x', tipo: 'clase',
     dia_semana: 2, hora_inicio: '08:00', hora_fin: '10:00',
     desde: '2026-12-01', hasta: '2026-10-01' } }).error));
ok('con todo, sí',
   F.soulRutinaGuardar(SOCIA, { datos: { nombre: 'Gimnasio', tipo: 'otro', dia_semana: 2,
     hora_inicio: '06:00', hora_fin: '07:00' } }).ok);
igual('y ya son cinco bloques', 5, F.soulRutina(SOCIA, {}).rutinas.length);


console.log('\n── Varias veces por semana, en UNA sola fila ──');
/**
 * Ella lo pidió así: «hay horarios que son varias veces por semana […]
 * las 6 veces que debo de hacer ejercicio en la semana».
 *
 * Antes cada fila era un día: seis veces eran seis filas, y cambiar la
 * hora había que cambiarla seis veces.
 */
[['1,3,5', [1, 3, 5], 'separados por coma'],
 ['1-5', [1, 2, 3, 4, 5], 'un rango, de lunes a viernes'],
 ['1-5,7', [1, 2, 3, 4, 5, 7], 'un rango y un suelto'],
 ['1-6', [1, 2, 3, 4, 5, 6], 'las seis veces de ejercicio'],
 ['L,X,V', [1, 3, 5], 'con letras, como lo dice cualquiera'],
 ['lunes y jueves', [1, 4], 'escrito como se habla'],
 ['diario', [1, 2, 3, 4, 5, 6, 7], 'todos los días'],
 ['3', [3], 'uno solo, como siempre'],
].forEach(function (caso) {
  igual(caso[2], caso[1], F.diasDeRutina_(caso[0]));
});

/** Lo que no se entiende se descarta, no se convierte en lunes. */
igual('un día que no se entiende NO se vuelve lunes', [], F.diasDeRutina_('cuando pueda'));
igual('ni un número fuera de rango', [], F.diasDeRutina_('9'));
igual('y de una lista a medias se salva lo que sí se entiende',
      [2, 5], F.diasDeRutina_('2, cuando pueda, 5'));
igual('sin repetidos, y en orden', [1, 3], F.diasDeRutina_('3,1,3'));

console.log('\n── Y corre los días que toca ──');
const gym = { activo: true, dias: [1, 2, 3, 4, 5, 6], desde: '', hasta: '' };
igual('el lunes sí, el domingo no',
      [true, false],
      [F.rutinaCorre_(gym, '2026-09-21'), F.rutinaCorre_(gym, '2026-09-27')]);
igual('y los seis días de la semana corren',
      6, [21, 22, 23, 24, 25, 26, 27]
        .filter(d => F.rutinaCorre_(gym, '2026-09-' + d)).length);

/** Una fila vieja, con un solo día en la columna, tiene que seguir corriendo. */
const deAntes = { activo: true, dia: 4, desde: '', hasta: '' };
ok('una rutina vieja de un solo día sigue funcionando',
   F.rutinaCorre_(deAntes, '2026-09-24') === true &&
   F.rutinaCorre_(deAntes, '2026-09-25') === false);


console.log('\n── Y las horas se cuentan en TODOS sus días ──');
/**
 * El fallo que esto atrapa: con una fila por día daba igual sumar en
 * `r.dia`. Ahora que el gimnasio de seis veces es UNA fila, contarlo
 * solo el lunes dejaría cinco días pareciendo libres — y esa es justo
 * la cuenta que decide si la semana cabe.
 */
sembrarHojas();
// Solo el encabezado: se mide el gimnasio, no lo que ya hubiera sembrado.
LIBROS.s.Rutina = [LIBROS.s.Rutina[0]];
LIBROS.s.Rutina.push(['g1', YO, 'otro', 'Gimnasio', '1-6', '06:00', '07:00',
                      'Smart Fit', '', '', '', '', '', '', 'si', '']);
const rr = F.soulRutina(SOCIA, {});
igual('una hora cada uno de los seis días',
      [1, 1, 1, 1, 1, 1, 0],
      [1, 2, 3, 4, 5, 6, 7].map(d => rr.ocupadasPorDia[d]));
igual('y es UN bloque, no seis', 1, rr.rutinas.length);
igual('que sabe sus días', [1, 2, 3, 4, 5, 6], rr.rutinas[0].dias);

console.log('\n── Si todavía no se corrió bootstrapTodo ──');
sembrarHojas();
delete LIBROS.s.Rutina;
delete LIBROS.s.Turnos;
h = F.soulHoy(SOCIA, {});
ok('NovaSoul abre igual, no se cae entera', h.ok === true);
ok('y dice qué hojas faltan, en vez de fingir que está vacío',
   h.faltanHojas.indexOf('Rutina') !== -1, JSON.stringify(h.faltanHojas));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
