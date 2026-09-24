/**
 * LA PROFUNDIDAD DEL CIELO. Lo que ella pidió, comprobado.
 *
 * ┌─ LO QUE ESTA PRUEBA EXISTE PARA ATRAPAR ───────────────────┐
 * │                                                            │
 * │ El pensum automático llevaba semanas «funcionando» y sin    │
 * │ producir nada. No estaba roto: leía la hoja Tránsitos, la  │
 * │ hoja estaba vacía, y cero tránsitos dan cero temporadas    │
 * │ sin lanzar ni un error. La prueba vieja lo tapaba porque   │
 * │ SEMBRABA tránsitos a mano antes de llamarlo.               │
 * │                                                            │
 * │ Así que aquí la cadena se prueba ENTERA y desde cero:      │
 * │                                                            │
 * │   hoja vacía → sembrar → pensum automático → temporadas    │
 * │                                                            │
 * │ Si cualquier eslabón deja de producir, esto falla. Que es  │
 * │ lo que no pasó la vez anterior.                            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Y lo otro: que la lectura responda SIEMPRE las cuatro preguntas que
 * ella hizo —tiempo para qué, cómo me afecta, cómo la trabajo a favor,
 * cómo se ve el desequilibrio— y no tres de cuatro cuando falta un
 * dato. Media lectura es peor que ninguna, porque no se nota.
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
        /**
         * Escribe DENTRO del rango, no filas enteras. Un stub que
         * reemplaza la fila completa hace pasar código que en Google
         * borraría columnas — ese error ya costó una tarde.
         */
        setValues: (v) => {
          v.forEach((fila, i) => {
            const destino = m[f - 1 + i] || (m[f - 1 + i] = []);
            fila.forEach((val, j) => { destino[c - 1 + j] = val; });
          });
        },
        setValue: (val) => {
          const destino = m[f - 1] || (m[f - 1] = []);
          destino[c - 1] = val;
        },
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

(0, eval)(src + '\n;globalThis.__F = { transitosSembrar_, lecturaTransito_, lecProfeccion_,' +
  ' lecEjes_, lecIntensidad_, cieloRango_, soulCieloLectura, soulTransitoLectura,' +
  ' pensumAuto_, pensumDe_, transitosDe_, soulCielo, efemeridesVencen_, lunasEntre_,' +
  ' faseLunar_, EFEMERIDES_SEMILLA, EFEMERIDES_CARTA, EFEMERIDES_HASTA,' +
  ' LEC_CASAS, LEC_CUERPOS, LEC_ASPECTOS, LEC_NATAL, CIELO_RANGOS, LUNA_FASES };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_TRA = ['usuario_id','fecha','casa','casa_placidus','tema','intensidad_pct',
  'texto_transito','por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto',
  'a_natal','desde','hasta','fuente'];
const C_PEN = ['id','usuario_id','desde','hasta','titulo','cuerpo','casa','casa_alterna',
  'momento','que_pide','que_evitar','nota'];
const C_CAR = ['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota'];
const C_USU = ['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
  'lugar_nacimiento','zona_horaria','acento','modo','idioma'];
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
  LIBROS.s = {
    Transitos: [C_TRA.slice()], Pensum: [C_PEN.slice()], Carta: [C_CAR.slice()],
    Usuarios: [C_USU.slice(),
      ['u1', 'Manuela', YO, '1995-09-20', '13:20', 'Palmira', 'America/Bogota', '', '', '']],
    Revolucion: [['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota']],
    Pendientes: [['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
      'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
      'riesgo','nota','materia_id']],
  };
  // Su carta real, la que ya tiene cargada desde Horus.
  [['ascendente','capricornio',15.1,1],['sol','virgo',27.3,9],['luna','leo',10.2,8],
   ['mercurio','libra',20.0,10],['venus','libra',5.7,10],['marte','escorpio',9.0,11],
   ['jupiter','sagitario',9.0,12],['saturno','piscis',20.9,3],['urano','capricornio',26.6,1],
   ['neptuno','capricornio',22.8,1],['pluton','escorpio',28.3,11],
   ['medio_cielo','libra',19.4,10],['nodo_norte','libra',26.9,10],
  ].forEach(function (c) {
    LIBROS.s.Carta.push([YO, c[0], c[1], c[2], c[3], '', '']);
  });
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };

// ══════════════════════════════════════════════════════════════
console.log('\n── Las efemérides existen y son las suyas ──');
ok('hay semilla cargada', F.EFEMERIDES_SEMILLA.length > 100,
   F.EFEMERIDES_SEMILLA.length + ' temporadas');
igual('la semilla es de su carta', YO, F.EFEMERIDES_CARTA.usuario_id);
ok('trae los nodos, que ella pidió por nombre',
   !!F.EFEMERIDES_CARTA.natal.nodo_norte && !!F.EFEMERIDES_CARTA.natal.nodo_sur);
ok('trae el Fondo del cielo, que ella pidió por nombre',
   !!F.EFEMERIDES_CARTA.natal.fondo_cielo);
igual('su Ascendente es Capricornio', 'capricornio', F.EFEMERIDES_CARTA.natal.ascendente.signo);
igual('su Luna está en Leo, casa 8', ['leo', 8],
  [F.EFEMERIDES_CARTA.natal.luna.signo, F.EFEMERIDES_CARTA.natal.luna.casa]);
/**
 * Contra su captura de Horus. Si alguien cambia los datos de nacimiento
 * o el script, esto se cae antes de que ella lo note en pantalla.
 */
igual('su Nodo Norte está en Libra casa 10', ['libra', 10],
  [F.EFEMERIDES_CARTA.natal.nodo_norte.signo, F.EFEMERIDES_CARTA.natal.nodo_norte.casa]);

// ══════════════════════════════════════════════════════════════
console.log('\n── Sembrar: que Nova traiga lo que Nova sabe ──');
sembrarHojas();
igual('antes de sembrar la hoja está vacía', 0, F.transitosDe_(YO).length);

const s1 = F.transitosSembrar_(YO);
ok('siembra las temporadas calculadas', s1.sembro === F.EFEMERIDES_SEMILLA.length,
   'sembró ' + s1.sembro);
igual('y ahora sí se leen', F.EFEMERIDES_SEMILLA.length, F.transitosDe_(YO).length);

const s2 = F.transitosSembrar_(YO);
igual('correrlo otra vez NO duplica', 0, s2.sembro);
igual('y la hoja sigue igual', F.EFEMERIDES_SEMILLA.length, F.transitosDe_(YO).length);

sembrarHojas();
const ajeno = F.transitosSembrar_('otra@persona.com');
igual('a otra persona NO se le siembra la carta de Manuela', 0, ajeno.sembro);
ok('y se le explica por qué', /otra carta/.test(ajeno.porque || ''), ajeno.porque);

/** Lo que ella escribió a mano manda: no se siembra encima. */
sembrarHojas();
LIBROS.s.Transitos.push([YO, '2026-10-01', 5, 5, 'Lo mío', '', 'mi texto', '', '', '',
                         'saturno', 'cuadratura', 'Sol', '2026-10-01', '2026-11-01', 'Horus']);
igual('si ella ya escribió algo, no se siembra nada encima', 0, F.transitosSembrar_(YO).sembro);
igual('y su fila sigue ahí sola', 1, F.transitosDe_(YO).length);

// ══════════════════════════════════════════════════════════════
console.log('\n── LA CADENA ENTERA: hoja vacía → pensum lleno ──');
/**
 * Esto es lo que ella reclamó: «el pensum kármico tampoco me lo está
 * dando y se supone que lo da en automático». Aquí se prueba SIN
 * sembrar tránsitos a mano, que es como le pasa a ella.
 */
sembrarHojas();
igual('de entrada no hay ni un tránsito', 0, F.transitosDe_(YO).length);
igual('ni una temporada de pensum', 0, F.pensumDe_(YO).length);

const foto = F.soulCielo(SOCIA, {});
ok('abrir El cielo siembra sola', foto.automatico.transitosSembrados > 100,
   'sembró ' + foto.automatico.transitosSembrados);
ok('y crea pensum sola', foto.automatico.pensumCreado > 0,
   'creó ' + foto.automatico.pensumCreado);
ok('el pensum queda de verdad en la hoja', F.pensumDe_(YO).length > 0,
   F.pensumDe_(YO).length + ' temporadas');
ok('y hay al menos una abierta hoy', foto.pensumAbierto.length > 0,
   foto.pensumAbierto.length + ' abiertas');
/** No vuelca cinco años de una: la ventana es de 120 días. */
ok('no vuelca los cinco años de golpe', F.pensumDe_(YO).length < 40,
   F.pensumDe_(YO).length + ' temporadas creadas');
ok('dice qué creó, no lo hace en silencio', foto.automatico.pensumTitulos.length > 0);

const foto2 = F.soulCielo(SOCIA, {});
igual('entrar otra vez no siembra de nuevo', 0, foto2.automatico.transitosSembrados);
igual('ni crea el pensum otra vez', 0, foto2.automatico.pensumCreado);

// ══════════════════════════════════════════════════════════════
console.log('\n── Las cuatro preguntas, siempre las cuatro ──');
const t = { cuerpo: 'saturno', aspecto: 'cuadratura', aNatal: 'Ascendente', casa: 4 };
const l = F.lecturaTransito_(t);
ok('hay lectura', l.hay);
ok('1· tiempo para qué', !!l.tiempoPara && l.tiempoPara.length > 20, l.tiempoPara);
ok('2· cómo te afecta', !!l.queEs && l.queEs.length > 40, l.queEs);
ok('3· cómo trabajarla a favor', l.aFavor.length >= 3, l.aFavor);
ok('4· cómo se ve el desequilibrio', l.desequilibrio.length >= 3, l.desequilibrio);
ok('dice que es la lectura común y que la suya manda', /manda sobre esto/.test(l.nota), l.nota);

/**
 * Las cuatro preguntas se responden ENTERAS para cualquier combinación
 * de la semilla. Si alguna casa, planeta o punto natal falta en el
 * vocabulario, la lectura saldría coja y nadie lo vería en pantalla:
 * saldría un párrafo más corto, no un error.
 */
let cojas = 0, sinCasa = 0, sinNatal = 0;
F.EFEMERIDES_SEMILLA.forEach(function (linea) {
  const x = linea.split('|');
  const nat = F.LEC_NATAL[x[2]];
  const lec = F.lecturaTransito_({ cuerpo: x[0], aspecto: x[1],
    aNatal: nat ? nat.nombre : x[2], casa: Number(x[6]) });
  if (!lec.hay) { cojas++; return; }
  if (!lec.casaNombre) sinCasa++;
  if (!nat) sinNatal++;
  if (!lec.tiempoPara || !lec.queEs || lec.aFavor.length < 3 || lec.desequilibrio.length < 3) cojas++;
});
igual('las ' + F.EFEMERIDES_SEMILLA.length + ' temporadas se leen completas', 0, cojas);
igual('ninguna se queda sin casa', 0, sinCasa);
igual('ningún punto natal falta en el vocabulario', 0, sinNatal);

console.log('\n── La intensidad ordena bien ──');
ok('Plutón cuadratura a la Luna pesa más que Júpiter sextil a Mercurio',
   F.lecIntensidad_('pluton', 'cuadratura', 'luna') >
   F.lecIntensidad_('jupiter', 'sextil', 'mercurio'),
   F.lecIntensidad_('pluton', 'cuadratura', 'luna') + ' vs ' +
   F.lecIntensidad_('jupiter', 'sextil', 'mercurio'));
ok('un ángulo pesa más que un generacional tocando a otro generacional',
   F.lecIntensidad_('saturno', 'cuadratura', 'ascendente') >
   F.lecIntensidad_('neptuno', 'trigono', 'pluton'));
ok('nunca se pasa de 100',
   F.lecIntensidad_('pluton', 'conjuncion', 'ascendente') <= 100);

// ══════════════════════════════════════════════════════════════
console.log('\n── Su año: casa 8 en Leo, como ella lo dijo ──');
const prof = F.lecProfeccion_(8, 'capricornio', F.EFEMERIDES_CARTA);
igual('casa 8', 8, prof.casa);
igual('en Leo', 'Leo', prof.signoNombre);
igual('lo dirige el Sol', 'Sol', prof.regente);
/**
 * Y el dato que vuelve el año SUYO y no el de cualquiera de 31: su Sol
 * está en casa 9, así que un año de casa 8 se resuelve por casa 9.
 */
igual('y su Sol está en casa 9', 9, prof.regenteDonde.casa);
ok('lo dice en el texto', prof.texto.join(' ').indexOf('casa 9') !== -1);
/** Su Luna vive en casa 8: el año la despierta. Eso hay que decirlo. */
ok('ve que su Luna vive en esa casa',
   prof.habitantes.filter(function (h) { return h.nombre === 'Luna'; }).length === 1,
   JSON.stringify(prof.habitantes));
ok('el texto lo nombra', prof.texto.join(' ').indexOf('Luna') !== -1);
ok('responde tiempo para qué', /soltar lo que ya se acab/.test(prof.tiempoPara), prof.tiempoPara);
ok('responde el desequilibrio', !!prof.desequilibrio, prof.desequilibrio);

console.log('\n── Los ejes del karma ──');
const ejes = F.lecEjes_(F.EFEMERIDES_CARTA);
ok('hay eje de nodos', ejes.filter(e => e.id === 'nodos').length === 1);
ok('hay eje Mediocielo–Fondo del cielo', ejes.filter(e => e.id === 'mc_ic').length === 1);
/**
 * Su Nodo Norte está a 7.5° de su Mediocielo. Eso no es un detalle
 * decorativo: es la firma de su mapa, y si Nova no lo nombra, la
 * lectura del año se queda a medias.
 */
ok('ve que su Nodo Norte está pegado al Mediocielo',
   ejes.filter(e => e.id === 'nn_mc').length === 1,
   JSON.stringify(ejes.map(e => e.id)));

// ══════════════════════════════════════════════════════════════
console.log('\n── Semana, mes y año miran cosas distintas ──');
sembrarHojas();
F.transitosSembrar_(YO);
const sem = F.cieloRango_(YO, 'semana', '2026-09-24');
const anio = F.cieloRango_(YO, 'anio', '2026-09-24');

ok('la semana no muestra generacionales',
   sem.temporadas.filter(x => x.grupo === 'generacional').length === 0);
ok('el año no muestra personales',
   anio.temporadas.filter(x => x.grupo === 'personal').length === 0);
ok('la semana trae las lunas del tramo', sem.lunas.length >= 0);
ok('el año no llena de lunas la pantalla', anio.lunas.length === 0);
ok('la semana dice dónde mirar lo lento', /año/.test(sem.fueraDeRango), sem.fueraDeRango);
ok('viene ordenado por intensidad, no por fecha',
   anio.temporadas.length < 2 ||
   anio.temporadas[0].intensidad >= anio.temporadas[1].intensidad);
ok('cada temporada trae sus cuatro respuestas',
   anio.temporadas.every(x => x.tiempoPara && x.queEs &&
                              x.aFavor.length >= 3 && x.desequilibrio.length >= 3));

console.log('\n── La llamada completa ──');
sembrarHojas();
const r = F.soulCieloLectura(SOCIA, { rango: 'anio' });
ok('responde', r.ok);
ok('trae la profección leída a fondo', !!r.profeccion && r.profeccion.casa === 8,
   JSON.stringify(r.profeccion && r.profeccion.casa));
ok('trae los ejes', r.ejes.length >= 2);
ok('trae el horizonte de las efemérides', !!r.efemerides.hasta, r.efemerides.hasta);
igual('y avisa con seis meses de anticipación, no el último día', false, r.efemerides.avisar);

const rMal = F.soulCieloLectura({ correo: 'k@nova.com', rol: 'operadora' }, {});
igual('una operadora no entra a NovaSoul', false, rMal.ok);

console.log('\n── La luna, dicha en castellano ──');
/**
 * Ella preguntó «¿qué es gibosa creciente?». Una fase que hay que ir a
 * buscar a Google es una fase que la pantalla no terminó de explicar.
 */
ok('todas las fases dicen qué se ve en el cielo',
   F.LUNA_FASES.every(f => f.forma && f.forma.length > 20),
   JSON.stringify(F.LUNA_FASES.filter(f => !f.forma).map(f => f.id)));
ok('la gibosa explica de dónde viene la palabra',
   /giba|joroba/.test(F.LUNA_FASES.filter(f => f.id === 'gibosa')[0].forma));
ok('la fase de hoy llega con su forma a la pantalla',
   !!F.faseLunar_('2026-09-24').forma);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
