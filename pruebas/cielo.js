/**
 * El cielo de NovaSoul.
 *
 * Lo que se comprueba: que la fase de la Luna sea la de verdad —contra
 * lunas llenas y nuevas conocidas—; que la carta se lea de un texto
 * pegado y diga qué NO entendió; que el pensum de ella mande sobre la
 * lectura común de la luna; y —la que importa— que la medición NO
 * hable sin muestra.
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

(0, eval)(src + '\n;globalThis.__F = { faseLunar_, revolucionVentana_, cartaLeer_,'
  + ' revolucionLectura_, pensumProponer_, soulPensumDesdeTransitos, CIELO_CASAS,' +
  ' soulCielo, soulCartaLeer, soulCartaGuardar, soulNacimientoGuardar,' +
  ' soulPensumGuardar, soulTransitoGuardar, soulRevolucionGuardar, cieloMedir_,' +
  ' CIELO_CUERPOS, CIELO_MINIMO, LUNA_FASES };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];
const C_CAR = ['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota'];
const C_TRA = ['usuario_id','fecha','casa','tema','intensidad_pct','texto_transito',
  'por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto','a_natal','desde','hasta','fuente'];
const C_PEN = ['id','usuario_id','desde','hasta','titulo','cuerpo','casa','momento',
  'que_pide','que_evitar','nota'];
const C_REV = ['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota'];
const C_USU = ['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
  'lugar_nacimiento','zona_horaria','acento','modo','idioma'];
const YO = 'manuela@nova.com';

function sembrarHojas() {
  UUID = 0;
  LIBROS.cen = { Trabajos: [['id','nombre','tipo','estado','horas_semana','fecha_entrega']],
                 Cobros: [['id']], Finanzas: [['id','fecha','flujo','categoria','monto','moneda']],
                 Metas: [['id']] };
  LIBROS.s = {
    Pendientes: [C_PEND], Carta: [C_CAR], Transitos: [C_TRA], Pensum: [C_PEN],
    Revolucion: [C_REV], Usuarios: [C_USU],
    Materias: [['id','usuario_id','nombre','trabajo_id','estado']],
    Horas: [['usuario_id','dia_semana','horas_libres','nota']],
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas','desde','hasta','estado','nota']],
    Fijos: [['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes','activo','nota']],
    Rutina: [['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin','lugar',
              'trabajo_id','materia_id','paga_fija','moneda','desde','hasta','activo','nota']],
    Turnos: [['id','usuario_id','rutina_id','fecha','paga','propinas','moneda','estado','finanza_id','nota']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const OPERADORA = { correo: 'k@nova.com', nombre: 'Kat', rol: 'operadora' };

console.log('\n── La Luna, contra fechas conocidas ──');
/**
 * Lunas llenas y nuevas reales. Si el cálculo se corre, esto falla —
 * que es exactamente para lo que está.
 */
[['2026-01-03', 'llena'], ['2026-03-03', 'llena'], ['2026-08-28', 'llena'],
 ['2026-01-18', 'nueva'], ['2026-09-11', 'nueva'], ['2000-01-06', 'nueva'],
].forEach(function (par) {
  const f = F.faseLunar_(par[0]);
  ok(par[0] + ' es ' + par[1], f.id === par[1], 'dio ' + f.id + ' (' + f.edadDias + ' días)');
});
ok('en la llena la iluminación pasa de 90%', F.faseLunar_('2026-01-03').iluminacion > 90,
   F.faseLunar_('2026-01-03').iluminacion + '%');
ok('en la nueva no llega a 10%', F.faseLunar_('2026-01-18').iluminacion < 10,
   F.faseLunar_('2026-01-18').iluminacion + '%');
/**
 * Entre la llena del 3 y la nueva del 18, la luna MENGUA. Entre la
 * nueva del 18 y la llena de febrero, crece.
 */
ok('el 10 de enero mengua (viene de la llena del 3)',
   F.faseLunar_('2026-01-10').creciendo === false,
   F.faseLunar_('2026-01-10').nombre);
ok('y el 24 crece (viene de la nueva del 18)',
   F.faseLunar_('2026-01-24').creciendo === true,
   F.faseLunar_('2026-01-24').nombre);
ok('cada fase trae qué momento es',
   F.LUNA_FASES.every(f => ['aprender','descansar','cambiar'].indexOf(f.momento) !== -1));

console.log('\n── Su año solar ──');
let r = F.revolucionVentana_('1998-09-14', '2026-09-23');
igual('el cumpleaños ya pasó: el año va de 2026 a 2027',
      ['2026-09-14', '2027-09-13'], [r.desde, r.hasta]);
igual('cumplió 28', 28, r.edad);
r = F.revolucionVentana_('1998-12-20', '2026-09-23');
igual('si el cumpleaños no ha llegado, el año arrancó el anterior',
      ['2025-12-20', '2026-12-19'], [r.desde, r.hasta]);
ok('sin fecha de nacimiento no inventa un año', F.revolucionVentana_('', '2026-09-23') === null);

console.log('\n── Leer la carta pegada ──');
const CARTA = `
Carta natal de Manuela
Sol en Virgo 21°34' Casa 2
Luna Capricornio 8°12 Casa 6
Mercurio en Libra 3° Casa 3
Venus | Leo | 27°05 | Casa 1
Marte en Escorpio 14° Casa 4 R
Júpiter en Piscis 2°
Saturno en Aries 29°48 Casa 9
Urano Acuario 12
Neptuno en Acuario 1°
Plutón Sagitario 7° R
Ascendente Leo 15°
Medio Cielo en Tauro 4°
Aspectos: Sol trígono Luna
Generado por Horus
`;
sembrarHojas();
let c = F.soulCartaLeer(SOCIA, { texto: CARTA });
ok('lee los doce cuerpos', c.encontradas.length === 12,
   c.encontradas.length + ': ' + c.encontradas.map(x => x.nombre).join(', '));
const porC = {};
c.encontradas.forEach(x => { porC[x.cuerpo] = x; });
igual('el Sol, con signo, grado y casa',
      ['virgo', 2, true], [porC.sol.signo, porC.sol.casa, porC.sol.grado > 21 && porC.sol.grado < 22]);
igual('Venus separado por barras', 'leo', porC.venus.signo);
ok('Marte marcado retrógrado', porC.marte.retrogrado === true);
ok('y el Sol no', porC.sol.retrogrado === false);
igual('«Urano Acuario 12» no inventa casa: 12 es el grado', null, porC.urano.casa);
igual('Medio Cielo, que son dos palabras', 'tauro', porC.medio_cielo.signo);
igual('cada cuerpo sabe a qué grupo pertenece',
      ['personal', 'social', 'generacional', 'angulo'],
      [porC.sol.grupo, porC.saturno.grupo, porC.pluton.grupo, porC.ascendente.grupo]);
ok('muestra lo que NO entendió', c.ignoradas.length >= 2,
   JSON.stringify(c.ignoradas.map(x => x.linea)));
ok('y «Generado por Horus» queda fuera',
   !c.encontradas.filter(x => /Horus/.test(x.linea)).length);

c = F.soulCartaLeer(SOCIA, { texto: 'Sol en Virgo\nLuna en Cáncer' });
igual('avisa qué cuerpos faltan', true, c.faltan.length > 0 && c.faltan.indexOf('Saturno') !== -1);

/**
 * ── SU CARTA DE VERDAD, COMO LA ESCRIBE HORUS ──
 *
 * Copiada de las capturas que mandó. Si el lector deja de entender
 * ESTA, da igual que entienda una inventada: esta es la que va a pegar.
 */
const HORUS = `Manuela Ramirez Sepúlveda
Sol	Vir 27	Casa 9
Luna	Leo 10	Casa 7
Mercurio	Lib 20	Casa 10
Venus	Lib 5	Casa 9
Marte	Esc 9	Casa 10
Júpiter	Sag 8	Casa 11
Saturno Rx	Pis 20	Casa 3
Urano Rx	Cap 26	Casa 1
Neptuno Rx	Cap 22	Casa 1
Plutón	Esc 28	Casa 11
Nodo Norte Rx	Lib 27	Casa 10
Nodo Sur Rx	Ari 27	Casa 4
Lilith	Gém 29	Casa 6
Quirón	Lib 1	Casa 9
Folo	Vir 22	Casa 9
Ceres	Lib 7	Casa 9
Pallas	Vir 14	Casa 8
Juno	Sag 22	Casa 12
Vesta	Vir 14	Casa 8
Parte de la Fortuna	Esc 28	Casa 11
Vertex	Vir 21	Casa 9
Ascendente	Cap 15
Fondo del cielo	Ari 19
Descendente	Cán 15
Mediocielo	Lib 19`;

console.log('\n── La carta REAL de Manuela, de Horus ──');
sembrarHojas();
let H = F.soulCartaLeer(SOCIA, { texto: HORUS });
const hp = {};
H.encontradas.forEach(x => { hp[x.cuerpo] = x; });
igual('los veinticinco cuerpos', 25, H.encontradas.length);
igual('Sol en Virgo 27, casa 9', ['virgo', 27, 9],
      [hp.sol.signo, hp.sol.grado, hp.sol.casa]);
igual('Luna en Leo 10, casa 7', ['leo', 10, 7],
      [hp.luna.signo, hp.luna.grado, hp.luna.casa]);
igual('Ascendente en Capricornio 15', ['capricornio', 15],
      [hp.ascendente.signo, hp.ascendente.grado]);
/**
 * «Vir 27 Casa 9» tiene dos números. El grado es 27, no 9. Confundirlos
 * pondría su Sol a 9 grados y movería medio análisis.
 */
ok('el grado NO se confunde con la casa', hp.sol.grado === 27 && hp.sol.casa === 9,
   JSON.stringify(hp.sol));
ok('los tres Rx de Horus se marcan',
   hp.saturno.retrogrado && hp.urano.retrogrado && hp.neptuno.retrogrado);
ok('y el Sol no', hp.sol.retrogrado === false);
igual('Nodo Sur NO se confunde con Nodo Norte',
      ['aries', 'libra'], [hp.nodo_sur.signo, hp.nodo_norte.signo]);
igual('Mediocielo en Libra 19', ['libra', 19],
      [hp.medio_cielo.signo, hp.medio_cielo.grado]);
igual('Fondo del cielo en Aries', 'aries', hp.fondo_cielo.signo);
igual('Descendente en Cáncer', 'cancer', hp.descendente.signo);
ok('los asteroides entran: Ceres, Pallas, Juno, Vesta, Folo, Vertex',
   !!(hp.ceres && hp.pallas && hp.juno && hp.vesta && hp.folo && hp.vertex));
igual('la Parte de la Fortuna también', 'escorpio', hp.fortuna.signo);

/** Su marco se respeta: los asteroides NO se cuelan entre los suyos. */
igual('los personales son sus cinco', 5,
      H.encontradas.filter(x => x.grupo === 'personal').length);
igual('los sociales, Júpiter y Saturno', 2,
      H.encontradas.filter(x => x.grupo === 'social').length);
igual('los generacionales, los tres de siempre', 3,
      H.encontradas.filter(x => x.grupo === 'generacional').length);
ok('Quirón NO se cuela entre los generacionales', hp.quiron.grupo === 'punto');
igual('no falta ningún planeta', 0, H.faltan.length);
ok('y la línea del nombre queda fuera',
   H.ignoradas.filter(x => /Manuela/.test(x.linea)).length === 1,
   JSON.stringify(H.ignoradas.map(x => x.linea)));

const gH = F.soulCartaGuardar(SOCIA, { filas: H.encontradas });
igual('se guardan los veinticinco', 25, gH.guardadas);
const cl0 = F.soulCielo(SOCIA, {});
igual('y Nova los agrupa como ella los piensa',
      [5, 2, 3, 4, 11],
      ['personal','social','generacional','angulo','punto']
        .map(gp => (cl0.cartaPorGrupo[gp] || []).length));

console.log('\n── Guardar la carta ──');
sembrarHojas();
let g = F.soulCartaGuardar(SOCIA, { filas: F.cartaLeer_(CARTA).encontradas });
igual('guarda los doce', 12, g.guardadas);
igual('y quedan doce filas, no más', 12, LIBROS.s.Carta.length - 1);
F.soulCartaGuardar(SOCIA, { filas: F.cartaLeer_(CARTA).encontradas });
igual('volver a cargarla la ACTUALIZA, no la duplica', 12, LIBROS.s.Carta.length - 1);
ok('una operadora no toca la carta',
   F.soulCartaGuardar(OPERADORA, { filas: [{ cuerpo: 'sol', signo: 'aries' }] }).ok === false &&
   F.soulCielo(OPERADORA, {}).ok === false);

console.log('\n── El cielo completo ──');
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14', hora: '04:20', lugar: 'Medellín', zona: 'America/Bogota' });
let cielo = F.soulCielo(SOCIA, {});
ok('abre', cielo.ok === true);
ok('dice que YA hay carta', cielo.tieneCarta === true);
igual('agrupa la carta como ella la piensa',
      [5, 2, 3, 2],
      ['personal','social','generacional','angulo'].map(gp => cielo.cartaPorGrupo[gp].length));
igual('los personales miran lo inmediato',
      'Lo inmediato, lo mío, lo propio.', cielo.grupos.personal.que);
igual('los sociales, expansión y estructura',
      'Expansión y estructura.', cielo.grupos.social.que);
igual('los generacionales, la época',
      'Marcan época. Van más allá del ego.', cielo.grupos.generacional.que);
igual('la semana tiene siete días', 7, cielo.semana.dias.length);
ok('cada día trae su luna y su momento',
   cielo.semana.dias.every(d => d.luna && d.momento));
igual('su año solar sale solo', '2026-09-14', cielo.revolucion.desde);
igual('y su edad', 28, cielo.revolucion.edad);

console.log('\n── Su pensum manda sobre la luna ──');
sembrarHojas();
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14' });
F.soulPensumGuardar(SOCIA, { datos: { titulo: 'Saturno por casa 10', desde: '2026-09-01',
  hasta: '2026-12-31', cuerpo: 'saturno', casa: 10, momento: 'cambiar',
  quePide: 'Cerrar lo que ya no sostiene.' } });
cielo = F.soulCielo(SOCIA, {});
igual('la temporada está abierta', 1, cielo.pensumAbierto.length);
ok('y todos los días toman SU momento, no el de la luna',
   cielo.semana.dias.every(d => d.momento === 'cambiar' && d.deDonde === 'pensum'),
   JSON.stringify(cielo.semana.dias.map(d => d.deDonde + ':' + d.momento)));
ok('cuando no hay temporada, habla la luna',
   (function () {
     LIBROS.s.Pensum = [C_PEN];
     const x = F.soulCielo(SOCIA, {});
     return x.semana.dias.every(d => d.deDonde === 'luna');
   })());
ok('una temporada sin título no se guarda',
   F.soulPensumGuardar(SOCIA, { datos: { desde: '2026-01-01' } }).ok === false);
ok('ni una sin fecha de inicio',
   /desde cuándo/.test(F.soulPensumGuardar(SOCIA, { datos: { titulo: 'x' } }).error));
ok('ni un momento inventado',
   F.soulPensumGuardar(SOCIA, { datos: { titulo: 'x', desde: '2026-01-01',
     momento: 'volar' } }).ok === false);

console.log('\n── Los tránsitos llevan duración ──');
sembrarHojas();
ok('un tránsito con rango se guarda',
   F.soulTransitoGuardar(SOCIA, { datos: { cuerpo: 'saturno', desde: '2026-09-01',
     hasta: '2026-11-30', tema: 'Estructura', aspecto: 'cuadratura', a_natal: 'luna' } }).ok);
ok('uno SIN fecha de inicio se rechaza',
   /no se puede cruzar/.test(F.soulTransitoGuardar(SOCIA,
     { datos: { cuerpo: 'luna', tema: 'x' } }).error));
ok('y uno con las fechas al revés también',
   F.soulTransitoGuardar(SOCIA, { datos: { cuerpo: 'luna', desde: '2026-12-01',
     hasta: '2026-10-01' } }).ok === false);
F.soulTransitoGuardar(SOCIA, { datos: { cuerpo: 'luna', desde: '2026-09-22',
  hasta: '2026-09-24', tema: 'Luna por casa 4' } });
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14' });
cielo = F.soulCielo(SOCIA, {});
igual('los dos están guardados', 2, cielo.transitos.length);
igual('y los dos corren hoy', 2, cielo.transitosHoy.length);
igual('Saturno es social y la Luna personal',
      ['social', 'personal'],
      cielo.transitos.map(t => t.grupo).sort().reverse());

console.log('\n── Si le funcionó a ELLA: no se habla sin muestra ──');
sembrarHojas();
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14' });
const tarea = (id, fecha, hecha) => {
  const f = C_PEND.map(() => '');
  f[0] = id; f[1] = YO; f[2] = 'Tarea ' + id; f[5] = fecha;
  f[10] = hecha ? 'hecho' : 'pendiente'; f[7] = hecha ? fecha : '';
  return f;
};
LIBROS.s.Pendientes.push(tarea('a1', '2026-08-03', true));
LIBROS.s.Pendientes.push(tarea('a2', '2026-08-04', false));
cielo = F.soulCielo(SOCIA, {});
igual('con dos tareas no dice nada', null, cielo.medicion.promedio === null ? null : cielo.medicion.fases.filter(f => f.pct !== null).length || null);
ok('y explica qué falta para poder hablar',
   /muestra suficiente/.test(cielo.medicion.porque), cielo.medicion.porque);
ok('ninguna fase da porcentaje todavía',
   cielo.medicion.fases.every(f => f.pct === null));
ok('pero dice cuántas faltan en cada una',
   cielo.medicion.fases.every(f => f.faltan > 0));

// Ahora sí: 12 entregas en menguante cumplidas, 12 en llena a medias
for (let i = 0; i < 12; i++) {
  // 2026-08-09 es cuarto menguante; se avanzan meses para no repetir fecha
  LIBROS.s.Pendientes.push(tarea('m' + i, '2026-0' + (1 + i % 5) + '-09', true));
}
cielo = F.soulCielo(SOCIA, {});
ok('con muestra sí aparece un porcentaje',
   cielo.medicion.fases.filter(f => f.pct !== null).length >= 1,
   JSON.stringify(cielo.medicion.fases.map(f => f.nombre + ':' + f.n + '/' + f.pct)));
ok('el promedio general existe', cielo.medicion.promedio !== null);
ok('una entrega del FUTURO no cuenta como incumplida',
   (function () {
     const antes = F.soulCielo(SOCIA, {}).medicion.conFecha;
     LIBROS.s.Pendientes.push(tarea('fut', '2027-01-15', false));
     return F.soulCielo(SOCIA, {}).medicion.conFecha === antes;
   })());

console.log('\n── La lectura de la revolución, compuesta por Nova ──');
sembrarHojas();
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14' });
let cl = F.soulCielo(SOCIA, {});
ok('sin los dos datos no hay lectura, y se dice',
   cl.revolucion.lectura.hay === false && /dos datos de Horus/.test(cl.revolucion.lectura.porque));

F.soulRevolucionGuardar(SOCIA, { anio: 2026, desde: '2026-09-14', hasta: '2027-09-13',
  ascendente: 'capricornio', casaSol: 10,
  planetas: [{ cuerpo: 'marte', casa: 1 }, { cuerpo: 'saturno', casa: 6 },
             { cuerpo: 'pluton', casa: 12 }, { cuerpo: 'inventado', casa: 3 },
             { cuerpo: 'venus', casa: 99 }] });
cl = F.soulCielo(SOCIA, {});
const L = cl.revolucion.lectura;
ok('ahora sí hay lectura', L.hay === true);
igual('dos partes: cómo entras y dónde va la atención', 2, L.partes.length);
ok('el ascendente del año dice CÓMO se entra',
   /estructura/.test(L.partes[0].texto), L.partes[0].texto);
ok('y la casa del Sol dice DÓNDE va',
   /Carrera, visibilidad/.test(L.partes[1].texto), L.partes[1].texto);
/**
 * Casa 10 es angular, y angular es acción: el año pide cambiar. No es
 * una regla inventada para que cuadrara — angulares, sucedentes y
 * cadentes son sus tres momentos desde hace siglos.
 */
igual('la casa 10 es angular, así que el año pide cambiar', 'cambiar', L.momento);
igual('los planetas van por SUS tres grupos',
      ['personal', 'social', 'generacional'], L.grupos.map(g => g.grupo));
igual('cada grupo dice qué mira',
      'Lo inmediato, lo mío, lo propio.', L.grupos[0].que);
ok('Marte en casa 1 sale con su área',
   L.grupos[0].cuerpos[0].area === 'Tú', JSON.stringify(L.grupos[0].cuerpos[0]));
ok('Saturno en casa 6 es el social',
   L.grupos[1].cuerpos[0].casa === 6 && L.grupos[1].cuerpos[0].nombre === 'Saturno');
ok('un planeta inventado se descarta en silencio',
   !JSON.stringify(L.grupos).match(/inventado/));
ok('y una casa 99 también', !JSON.stringify(L.grupos).match(/"casa":99/));
igual('no falta nada por decir', 0, L.faltan.length);

F.soulRevolucionGuardar(SOCIA, { anio: 2026, ascendente: 'capricornio', casaSol: 0, planetas: [] });
cl = F.soulCielo(SOCIA, {});
ok('sin la casa del Sol lo dice, y da la mitad que sí puede',
   cl.revolucion.lectura.partes.length === 1 &&
   /casa cae tu Sol/.test(cl.revolucion.lectura.faltan.join(' ')),
   JSON.stringify(cl.revolucion.lectura.faltan));

console.log('\n── Temporadas propuestas desde los tránsitos ──');
sembrarHojas();
F.soulNacimientoGuardar(SOCIA, { fecha: '1998-09-14' });
F.soulTransitoGuardar(SOCIA, { datos: { cuerpo: 'saturno', desde: '2026-09-01',
  hasta: '2026-11-30', casa: 10, tema: 'Estructura en lo público' } });
F.soulTransitoGuardar(SOCIA, { datos: { cuerpo: 'luna', desde: '2026-09-22',
  hasta: '2026-09-24', casa: 4, tema: 'Luna por casa 4' } });
cl = F.soulCielo(SOCIA, {});
igual('propone solo el que dura: la Luna de dos días no es una temporada',
      1, cl.pensumPropuesto.length);
igual('con su título armado', 'Saturno por casa 10', cl.pensumPropuesto[0].titulo);
igual('y el momento sacado de la casa', 'cambiar', cl.pensumPropuesto[0].momento);
igual('sin fechas inventadas: las de Horus',
      ['2026-09-01', '2026-11-30'],
      [cl.pensumPropuesto[0].desde, cl.pensumPropuesto[0].hasta]);
ok('proponer NO guardó nada', LIBROS.s.Pensum.length === 1);

g = F.soulPensumDesdeTransitos(SOCIA, { items: cl.pensumPropuesto });
igual('al confirmar, se guarda', 1, g.creadas);
igual('y ya está abierta', 1, F.soulCielo(SOCIA, {}).pensumAbierto.length);
igual('confirmarla otra vez no duplica', 0,
      F.soulPensumDesdeTransitos(SOCIA, { items: cl.pensumPropuesto }).creadas);
ok('y la propuesta ya la marca como puesta',
   F.soulCielo(SOCIA, {}).pensumPropuesto[0].yaEsta === true);
ok('sin marcar nada, lo dice',
   /No marcaste/.test(F.soulPensumDesdeTransitos(SOCIA, { items: [] }).error));
ok('una operadora no propone ni guarda',
   F.soulPensumDesdeTransitos(OPERADORA, { items: [{ titulo: 'x', desde: '2026-01-01' }] }).ok === false);

console.log('\n── Si todavía no hay carta ──');
sembrarHojas();
cielo = F.soulCielo(SOCIA, {});
ok('lo dice en vez de fingir una', cielo.tieneCarta === false);
ok('pero la luna de hoy sí se calcula', !!cielo.lunaHoy.nombre);
ok('y la pantalla abre igual', cielo.ok === true);
delete LIBROS.s.Carta;
delete LIBROS.s.Pensum;
ok('si faltan las hojas, tampoco se cae', F.soulCielo(SOCIA, {}).ok === true);

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
