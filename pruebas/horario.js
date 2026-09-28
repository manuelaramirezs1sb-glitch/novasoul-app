/**
 * QUE NOVA ORGANICE LA SEMANA.
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «que Nova organice mis horarios de trabajo, Nutrea siempre  │
 * │  es en la mañana y en la tarde, reparte las horas de        │
 * │  trabajo durante el día sin cruzarse con otros proyectos».  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE ESTA PRUEBA VIGILA DE VERDAD ──────────────────────┐
 * │                                                            │
 * │ No que el horario sea bonito: que sea POSIBLE.              │
 * │                                                            │
 * │  · nada encima de una clase o un turno                     │
 * │  · nada encima de otro bloque                              │
 * │  · nada fuera de la ventana del día                        │
 * │  · Nutrea en dos trozos, mañana y tarde                    │
 * │  · y —la más importante— que lo que NO cabe se diga con su  │
 * │    nombre en vez de meterse a la fuerza                    │
 * │                                                            │
 * │ Un horario que lo acomoda todo siempre existe: basta con    │
 * │ poner bloques encima de las clases. Pero esa semana es una  │
 * │ mentira que se descubre el martes, y entonces ya no se      │
 * │ vuelve a abrir la pantalla.                                 │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
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


(0, eval)(src + '\n;globalThis.__F = { soulHorario, bloquesDe_, diasDe_, huecosDelDia_, semanaComprometida_,' +
  ' horaNum_, libroOlvidar_, soulOlvidar_, sembrarBloques, HORARIO_DESDE, HORARIO_HASTA };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const YO = 'novasoul959@gmail.com';
const S = { correo: YO, nombre: 'Manuela', rol: 'socia' };

const C_BLO = ['id','usuario_id','nombre','tipo','trabajo_id','horas_min','horas_max',
               'cada','veces','franja','dias','partes','orden','activo','nota'];
const C_RUT = ['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin',
               'lugar','trabajo_id','materia_id','paga_fija','moneda','desde','hasta',
               'activo','nota'];
const C_HOR = ['usuario_id','dia_semana','horas_libres','nota'];

function bloque(o) {
  return C_BLO.map(function (c) {
    const d = { usuario_id: YO, activo: 'si', partes: 1, orden: 50,
                cada: 'dia', veces: 0, franja: 'cualquiera', dias: '1-7' };
    return o[c] !== undefined ? o[c] : (d[c] !== undefined ? d[c] : '');
  });
}
/**
 * Una clase o un turno.
 *
 * El `id` NO es decorativo: `rutinaDe_` descarta toda fila sin id. Mi
 * primera versión de este arnés no lo ponía, así que el servidor tiraba
 * las seis clases y la prueba comprobaba «no se pisa con nada» sobre una
 * semana SIN nada. Pasaba en verde sin probar absolutamente nada — el
 * peor tipo de prueba, porque da tranquilidad falsa.
 */
let nCla = 0;
function clase(nombre, dia, ini, fin) {
  nCla++;
  return C_RUT.map(function (c) {
    const d = { id: 'r' + nCla, usuario_id: YO, tipo: 'clase', nombre: nombre,
                dia_semana: dia, hora_inicio: ini, hora_fin: fin, activo: 'si' };
    return d[c] !== undefined ? d[c] : '';
  });
}

/** Su semana real: las clases y el turno de doce horas del sábado. */
function sembrar(bloques, horas) {
  F.libroOlvidar_(); F.soulOlvidar_();
  LIBROS.s = {
    Bloques: [C_BLO].concat(bloques || []),
    Rutina: [C_RUT,
      clase('QCA AMBIENTAL', 2, '14:00', '18:00'),
      clase('SALSABOR', 3, '16:00', '01:00'),
      clase('PROCESOS', 4, '16:00', '18:00'),
      clase('PROCESOS', 5, '14:00', '16:00'),
      clase('LABORATORIO', 5, '16:00', '19:00'),
      clase('SALSABOR', 6, '15:00', '03:00'),
    ],
    Horas: [C_HOR].concat([1,2,3,4,5,6,7].map(function (d) {
      const h = horas || { 1: 10, 2: 10, 3: 15, 4: 10, 5: 10, 6: 15, 7: 0 };
      return [YO, d, h[d], ''];
    })),
    Pendientes: [['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
                  'trabajo_id','materia_id','horas','estado','prioridad','riesgo','nota',
                  'materia_id2']],
    Trabajos: [['id','nombre','estado','horas_semana']],
    Turnos: [['id','usuario_id','fecha','rutina_id','base','propinas','nota']],
  };
}

const num = (h) => Number(String(h).split(':')[0]) + Number(String(h).split(':')[1]) / 60;

/** ¿Se pisan dos tramos? */
function chocan(a, b) {
  return num(a.inicio) < num(b.fin) && num(b.inicio) < num(a.fin);
}

console.log('\n══ 1 · NADA SE PISA CON NADA ══');
sembrar([
  bloque({ id:'b1', nombre:'Nutrea · tiendas', horas_min:2, horas_max:3,
           cada:'dia', dias:'1-6', partes:2, orden:10 }),
  bloque({ id:'b2', nombre:'Nova', horas_min:2, horas_max:3,
           cada:'semana', veces:3, franja:'tarde', dias:'1-6', orden:20 }),
  bloque({ id:'b3', nombre:'Universidad · estudio', horas_min:1, horas_max:2,
           cada:'dia', dias:'1-6', orden:30 }),
  bloque({ id:'b4', nombre:'Gym', horas_min:1.5, horas_max:1.5,
           cada:'semana', veces:3, franja:'manana', dias:'1-6', orden:60 }),
]);
let r = F.soulHorario(S, {});
ok('arma la semana', r.ok, JSON.stringify(r).slice(0, 160));
igual('siete días', 7, r.dias.length);
// Sin esto, todo lo de abajo pasaría en verde sobre una semana sin clases.
igual('las seis clases y turnos llegaron', 6,
      r.dias.reduce(function (a, d) { return a + d.fijos.length; }, 0));

let choques = [];
r.dias.forEach(function (d) {
  const todo = d.puestos.concat(d.fijos);
  for (let i = 0; i < todo.length; i++) {
    for (let j = i + 1; j < todo.length; j++) {
      // El turno que cruza medianoche se recorta a la ventana del día
      if (num(todo[i].fin) <= num(todo[i].inicio)) continue;
      if (num(todo[j].fin) <= num(todo[j].inicio)) continue;
      if (chocan(todo[i], todo[j])) {
        choques.push(d.dia + ': ' + todo[i].nombre + ' vs ' + todo[j].nombre);
      }
    }
  }
});
igual('ningún bloque encima de otro ni de una clase', [], choques);

console.log('\n══ 2 · TODO DENTRO DE LA VENTANA DEL DÍA ══');
let fuera = [];
r.dias.forEach(function (d) {
  d.puestos.forEach(function (b) {
    if (num(b.inicio) < F.HORARIO_DESDE - 0.001 || num(b.fin) > F.HORARIO_HASTA + 0.001) {
      fuera.push(d.dia + ' ' + b.nombre + ' ' + b.inicio + '-' + b.fin);
    }
  });
});
igual('nada antes de las 8 ni después de las 23', [], fuera);

console.log('\n══ 3 · NUTREA: MAÑANA Y TARDE, NO DOS VECES SEGUIDAS ══');
/**
 * «Nutrea siempre es en la mañana y en la tarde». Sin partirlo en dos
 * franjas el algoritmo pondría las dos sesiones seguidas a las ocho y
 * habría cumplido la letra sin cumplir nada.
 */
let malNutrea = [];
r.dias.forEach(function (d) {
  const n = d.puestos.filter(function (b) { return b.nombre === 'Nutrea · tiendas'; });
  if (!n.length) return;
  /**
   * Un día puede quedarse con UNA sola sesión: el viernes tiene Procesos
   * y Laboratorio ocupando toda la tarde. Lo que NO puede pasar es que
   * las dos caigan por la mañana, seguidas — eso cumple las horas e
   * incumple lo único que ella pidió de Nutrea.
   */
  if (n.length > 2) { malNutrea.push('día ' + d.dia + ': ' + n.length + ' bloques'); return; }
  if (num(n[0].inicio) >= 13) malNutrea.push('día ' + d.dia + ': el primero no es de mañana');
  if (n.length === 2 && num(n[1].inicio) < 13) {
    malNutrea.push('día ' + d.dia + ': las dos por la mañana');
  }
});
igual('nunca dos sesiones de Nutrea seguidas por la mañana', [], malNutrea);
// Y en los días con tarde libre, sí tiene que estar partida de verdad.
const conDos = r.dias.filter(function (d) {
  return d.puestos.filter(function (b) { return b.nombre === 'Nutrea · tiendas'; }).length === 2;
}).length;
ok('en la mayoría de los días queda partida mañana/tarde', conDos >= 3,
   conDos + ' días de 6');
ok('y está de lunes a sábado, no el domingo',
   r.dias[6].puestos.every(function (b) { return b.nombre !== 'Nutrea · tiendas'; }),
   JSON.stringify(r.dias[6].puestos.map(function (b) { return b.nombre; })));

console.log('\n══ 4 · NO SE PASA DE SUS HORAS ÚTILES ══');
/**
 * El reloj dice dónde CABE; sus horas útiles dicen cuánto QUIERE. Gana
 * la segunda: puede tener la tarde libre y no querer trabajar diez
 * horas ese día.
 */
let excedidos = [];
r.dias.forEach(function (d) {
  if (d.utiles === null) return;
  if (d.horasPuestas + d.horasFijas > d.utiles + 0.001) {
    excedidos.push('día ' + d.dia + ': ' + (d.horasPuestas + d.horasFijas) +
                   ' sobre ' + d.utiles);
  }
});
igual('ningún día se pasa de lo que ella dijo', [], excedidos);

console.log('\n══ 5 · LO QUE NO CABE SE DICE, NO SE MACHACA ══');
/**
 * Esta es la aserción que sostiene el archivo. Un horario que lo acomoda
 * todo siempre existe —basta con pisar las clases— y es exactamente el
 * que no sirve.
 */
sembrar([
  // Cuarenta horas de un solo proyecto: imposible de cuajar.
  bloque({ id:'x1', nombre:'Imposible', horas_min:8, horas_max:8,
           cada:'dia', dias:'1-7', orden:10 }),
]);
r = F.soulHorario(S, {});
ok('lo que no cupo se nombra', r.fuera.length > 0, JSON.stringify(r.fuera));
ok('con las horas que faltaron', r.faltaron > 0, String(r.faltaron));
ok('y aun así no pisó ninguna clase', (function () {
  let mal = false;
  r.dias.forEach(function (d) {
    d.puestos.forEach(function (b) {
      d.fijos.forEach(function (f) {
        if (num(f.fin) > num(f.inicio) && chocan(b, f)) mal = true;
      });
    });
  });
  return !mal;
})());

console.log('\n══ 6 · EL SÁBADO, CON SUS DOCE HORAS DE SALSABOR ══');
/**
 * «el sábado pueden ser 15, teniendo en cuenta que trabajo 12 h en
 *  Salsabor». Quedan tres, y ahí no cabe una sesión de tres horas más
 *  las dos de Nutrea: tiene que sobrar algo, y decirse.
 */
sembrar([
  bloque({ id:'n1', nombre:'Nutrea · tiendas', horas_min:2, horas_max:3,
           cada:'dia', dias:'6', partes:2, orden:10 }),
]);
r = F.soulHorario(S, {});
const sab = r.dias[5];
ok('el sábado tiene su turno largo', sab.horasFijas >= 8, String(sab.horasFijas));
ok('y lo que se puso cabe en lo que queda',
   sab.horasPuestas + sab.horasFijas <= 15.001,
   sab.horasPuestas + ' + ' + sab.horasFijas);

console.log('\n══ 7 · SIN REGLAS NO SE INVENTA UNA SEMANA ══');
sembrar([]);
r = F.soulHorario(S, {});
igual('lo dice en vez de devolver una semana vacía', true, r.sinReglas);
ok('y explica qué hacer', /Bloques/.test(r.porque || ''), r.porque);

console.log('\n══ 8 · los días se leen como ella los escribe ══');
igual("'1-6' son lunes a sábado", [1,2,3,4,5,6], F.diasDe_('1-6'));
igual("'1,2,4' son tres días sueltos", [1,2,4], F.diasDe_('1,2,4'));
igual('vacío son todos', [1,2,3,4,5,6,7], F.diasDe_(''));

console.log('\n══ 9 · SIN CONTAR LAS MISMAS HORAS DOS VECES ══');
/**
 * Ella, antes de pegar nada: «¿cruzas la info que ya tienes con la que
 * te di si hay algo repetido como las horas de Nutrea o Nova?».
 *
 * No lo hacía. `Trabajos.horas_semana` (Central) y `Bloques` (lo que le
 * dijo a Nova) son dos listas, y Nutrea está en las dos. Sin cruzarlas,
 * la tarjeta de la semana sumaba las de Central y el organizador
 * colocaba las de Bloques: dos números distintos para lo mismo, en dos
 * pantallas que se miran seguidas.
 */
sembrar([
  bloque({ id:'b1', nombre:'Nutrea · tiendas', trabajo_id:'t1',
           horas_min:2, horas_max:3, cada:'dia', dias:'1-6', partes:2, orden:10 }),
  bloque({ id:'b2', nombre:'Universidad · estudio',
           horas_min:1, horas_max:2, cada:'dia', dias:'1-6', orden:30 }),
]);
const ACTIVOS = [
  { id:'t1', nombre:'Nutrea', estado:'activo', horasSemana:20 },
  { id:'t2', nombre:'PHH',    estado:'activo', horasSemana:5 },
];
let c = F.semanaComprometida_(YO, ACTIVOS);

// Nutrea: 2 h × 6 días = 12 por el bloque. Las 20 de Central NO se suman.
const nutrea = c.detalle.filter(function (x) { return /Nutrea/.test(x.nombre); });
igual('Nutrea aparece UNA sola vez', 1, nutrea.length);
igual('y con las horas del bloque, no las de Central', 12, nutrea[0].horas);
ok('marcada como que viene del bloque', nutrea[0].deBloque === true);

// PHH no tiene bloque: conserva sus horas de Central, no desaparece.
const phh = c.detalle.filter(function (x) { return x.nombre === 'PHH'; });
igual('un proyecto sin bloque sigue contando', 1, phh.length);
igual('con sus horas de Central', 5, phh[0].horas);

// Y lo que no es proyecto —estudio, gym— suma lo suyo.
ok('el estudio también cuenta',
   c.detalle.some(function (x) { return /estudio/i.test(x.nombre) && x.horas === 6; }),
   JSON.stringify(c.detalle));

igual('el total es 12 + 5 + 6, sin duplicar', 23, c.total);

// Sin bloques, todo sigue como antes: las horas de Central mandan.
sembrar([]);
c = F.semanaComprometida_(YO, ACTIVOS);
igual('sin bloques, manda Central', 25, c.total);
igual('y se sabe que no hay bloques', false, c.hayBloques);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
