/**
 * El lector de cronogramas.
 *
 * Lo que se comprueba no es que encuentre fechas: es que NO se invente
 * ninguna, que diga cuándo adivinó el año, que muestre lo que ignoró, y
 * que un parcial entre como inamovible — porque si entrara corrible, el
 * repartidor de la semana le propondría correrlo.
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
      appendRow: (f) => m.push(f),
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

(0, eval)(src + '\n;globalThis.__F = { silaboLeer_, silaboFechas_, silaboAnio_,' +
  ' soulSilaboLeer, soulSilaboGuardar, soulMaterias, soulMateriaGuardar,' +
  ' soulMateriaBorrar, soulHoy };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];
const C_MAT = ['id','usuario_id','nombre','codigo','profesor','carpeta','semestre',
  'trabajo_id','estado','nota'];
const C_TRAB = ['id','nombre','contraparte','tipo','estado','moneda','valor_acordado',
  'forma_cobro','fecha_inicio','fecha_entrega','horas_semana','especificacion','documento','nota'];
const YO = 'manuela@nova.com';

function sembrarHojas() {
  LIBROS.cen = {
    Trabajos: [C_TRAB, ['t3','Universidad','U','estudio','activo','',0,'','','',6,'','','']],
    Cobros: [['id','trabajo_id','concepto','monto','moneda','fecha_esperada','fecha_cobrada','estado','nota']],
    Finanzas: [['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta','recurrente','trabajo_id','nota']],
    Metas: [['id','tipo','nombre']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND],
    Materias: [C_MAT, ['m1', YO, 'Estadística', 'EST-301', 'Prof. Gómez',
                        'https://drive.google.com/drive/folders/abc', '2026-2', 't3', 'activa', '']],
    Horas: [['usuario_id','dia_semana','horas_libres','nota']],
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas','desde','hasta','estado','nota']],
    Fijos: [['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes','activo','nota']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const OPERADORA = { correo: 'k@nova.com', nombre: 'Kat', rol: 'operadora' };

// ── Un cronograma como los de verdad ─────────────────────────
const SILABO = `
CRONOGRAMA · ESTADÍSTICA INFERENCIAL 2026-2
Profesor: Gómez

Semana 1 | Presentación del curso | 5 de agosto
Quiz 1 - 30 de septiembre - 10%
Parcial 1 — 15 de octubre — 25%
Entrega del proyecto: semana del 12 al 16 de octubre  20%
Taller en clase 03/11
Parcial 2, 19 de noviembre, 25%
Sustentación final  5 de diciembre de 2026   20%
Bibliografía: Walpole, 9a edición
La nota mínima para aprobar es 3.0
`;

console.log('\n── La puerta ──');
sembrarHojas();
ok('una operadora no lee sílabos', F.soulSilaboLeer(OPERADORA, { texto: SILABO }).ok === false);
ok('ni guarda materias', F.soulMateriaGuardar(OPERADORA, { datos: { nombre: 'x' } }).ok === false);
ok('un texto vacío se rechaza con razón',
   /Pega el cronograma/.test(F.soulSilaboLeer(SOCIA, { texto: '  ' }).error));

console.log('\n── Lo que encuentra ──');
let r = F.soulSilaboLeer(SOCIA, { texto: SILABO });
const t = r.encontradas.map(e => e.titulo + ' | ' + e.fecha + ' | ' + e.peso);
ok('encuentra las siete líneas con fecha', r.encontradas.length === 7,
   r.encontradas.length + ':\n         ' + t.join('\n         '));
igual('y las ordena por fecha',
      r.encontradas.map(e => e.fecha).slice().sort(), r.encontradas.map(e => e.fecha));

const porFecha = {};
r.encontradas.forEach(e => { porFecha[e.fecha] = e; });
ok('el parcial del 15 de octubre, con su 25%',
   porFecha['2026-10-15'] && porFecha['2026-10-15'].peso === 25 &&
   /Parcial 1/.test(porFecha['2026-10-15'].titulo),
   JSON.stringify(porFecha['2026-10-15']));
ok('el quiz en formato con guiones',
   porFecha['2026-09-30'] && /Quiz 1/.test(porFecha['2026-09-30'].titulo),
   JSON.stringify(porFecha['2026-09-30']));
ok('la fecha en 03/11 se lee día primero: 3 de noviembre',
   !!porFecha['2026-11-03'] && /Taller/.test(porFecha['2026-11-03'].titulo),
   JSON.stringify(porFecha['2026-11-03']));
ok('el parcial 2 separado por comas',
   porFecha['2026-11-19'] && porFecha['2026-11-19'].peso === 25);
ok('la sustentación, que sí trae año',
   porFecha['2026-12-05'] && porFecha['2026-12-05'].anioInferido === false,
   JSON.stringify(porFecha['2026-12-05']));

/**
 * «Semana del 12 al 16 de octubre» tiene dos fechas. Manda la última:
 * es cuando se entrega, no cuando empieza. Y se marca como rango para
 * que la pantalla lo pueda decir.
 */
ok('de un rango toma la fecha de entrega, no la de inicio',
   porFecha['2026-10-16'] && porFecha['2026-10-16'].rango === true,
   JSON.stringify(r.encontradas.filter(e => e.rango)));

console.log('\n── Lo que NO se inventa ──');
ok('cuando adivina el año, lo dice',
   porFecha['2026-10-15'].anioInferido === true);
/**
 * El 5 de agosto ya pasó, pero hace mes y medio: es de ESTE semestre,
 * no del año entrante. Correrlo a 2027 le habría metido en la lista una
 * clase que ya dio.
 */
igual('una fecha de hace mes y medio se queda en este año', '2026-08-05',
      r.encontradas.filter(e => /Presentación/.test(e.titulo))[0].fecha);

const ign = r.ignoradas.map(i => i.linea);
ok('muestra lo que ignoró, no lo tira en silencio', r.ignoradas.length > 0,
   JSON.stringify(ign));
ok('y la bibliografía queda fuera',
   !r.encontradas.filter(e => /Walpole/.test(e.titulo)).length);
ok('«nota mínima 3.0» NO se lee como una fecha',
   !r.encontradas.filter(e => /mínima/.test(e.titulo)).length,
   JSON.stringify(r.encontradas.map(e => e.titulo)));
ok('cada línea ignorada dice por qué',
   r.ignoradas.every(i => !!i.porque), JSON.stringify(r.ignoradas));

console.log('\n── Fechas imposibles ──');
r = F.soulSilaboLeer(SOCIA, { texto: '31 de febrero entrega final\n45/13/2026 otra cosa' });
ok('el 31 de febrero no se corre solo al 3 de marzo',
   !r.encontradas.filter(e => /03-0[123]/.test(e.fecha)).length,
   JSON.stringify(r.encontradas));
ok('y una fecha inválida se ignora con razón', r.ignoradas.length === 2,
   JSON.stringify(r.ignoradas));

console.log('\n── Guardar solo lo confirmado ──');
sembrarHojas();
r = F.soulSilaboLeer(SOCIA, { texto: SILABO });
ok('leer NO escribió nada', LIBROS.s.Pendientes.length === 1,
   (LIBROS.s.Pendientes.length - 1) + ' filas');

const dos = r.encontradas.filter(e => e.fecha === '2026-10-15' || e.fecha === '2026-11-19');
let g = F.soulSilaboGuardar(SOCIA, { materia: 'm1', items: dos });
igual('guarda las dos que marcó, no las siete', 2, g.creadas);
igual('y solo esas están', 2, LIBROS.s.Pendientes.length - 1);

let hoy = F.soulHoy(SOCIA, {});
const parcial = hoy.pendientes.filter(p => /Parcial 1/.test(p.texto))[0];
ok('el pendiente lleva el nombre de la materia adelante',
   /^Estadística · /.test(parcial.texto), parcial.texto);
igual('un parcial entra INAMOVIBLE: nadie va a proponer correrlo',
      'inamovible', parcial.riesgo);
igual('y con la materia y el proyecto de Central enganchados',
      ['t3'], [parcial.trabajoId]);
igual('lo que vale 25% entra como prioridad alta', 'alta', parcial.prioridad);
ok('y la nota dice cuánto vale', /25%/.test(parcial.nota), parcial.nota);
igual('sin horas estimadas: eso lo pone ella', 0, parcial.horas);

console.log('\n── Pegar el mismo sílabo otra vez ──');
g = F.soulSilaboGuardar(SOCIA, { materia: 'm1', items: dos });
igual('no duplica nada', 0, g.creadas);
igual('y dice cuántas ya estaban', 2, g.repetidas);
igual('la hoja sigue con dos', 2, LIBROS.s.Pendientes.length - 1);

ok('una materia que no existe se rechaza',
   F.soulSilaboGuardar(SOCIA, { materia: 'zz', items: dos }).ok === false);
ok('sin nada marcado, lo dice',
   /No marcaste/.test(F.soulSilaboGuardar(SOCIA, { materia: 'm1', items: [] }).error));

console.log('\n── Las materias ──');
let m = F.soulMaterias(SOCIA, {});
igual('lista la materia', 1, m.materias.length);
igual('con sus dos entregas abiertas', 2, m.materias[0].abiertas);
igual('y la más próxima', '2026-10-15', m.materias[0].proxima);
ok('guarda el enlace de la carpeta, no el archivo',
   /^https:\/\/drive/.test(m.materias[0].carpeta), m.materias[0].carpeta);
ok('solo ofrece proyectos de estudio para enganchar',
   m.trabajos.length === 1 && m.trabajos[0].id === 't3');

ok('un enlace que no es enlace se rechaza',
   /empezar por http/.test(F.soulMateriaGuardar(SOCIA,
     { datos: { nombre: 'Física', carpeta: 'mi carpeta' } }).error));
ok('una materia sin nombre se rechaza',
   F.soulMateriaGuardar(SOCIA, { datos: { carpeta: 'https://x.com' } }).ok === false);
ok('se puede crear una materia nueva',
   F.soulMateriaGuardar(SOCIA, { datos: { nombre: 'Química Ambiental', trabajo_id: 't3' } }).ok);
igual('y ya son dos', 2, F.soulMaterias(SOCIA, {}).materias.length);

/**
 * Borrar la materia no puede llevarse los parciales por delante. Se
 * borra la carpeta, no la fecha del examen.
 */
const b = F.soulMateriaBorrar(SOCIA, { id: 'm1' });
ok('borrar la materia avisa cuántas entregas quedan sueltas',
   b.ok === true && b.colgadas === 2, JSON.stringify(b));
igual('y NO borra los parciales', 2, LIBROS.s.Pendientes.length - 1);
hoy = F.soulHoy(SOCIA, {});
ok('que siguen contando en la semana',
   hoy.pendientes.filter(p => /Parcial/.test(p.texto)).length === 2,
   JSON.stringify(hoy.pendientes.map(p => p.texto)));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
