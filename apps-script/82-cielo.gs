/**
 * ═══════════════════════════════════════════════════════════
 *  NOVASOUL · EL CIELO
 * ═══════════════════════════════════════════════════════════
 *
 * Su marco, con sus palabras, porque manda sobre todo lo de abajo:
 *
 *   «Los planetas marcan el tiempo: hay tiempo para aprender, tiempo
 *    para descansar y tiempo para cambiar. Los personales afectan lo
 *    inmediato, lo mío, lo propio. Los sociales hablan de expansión y
 *    estructura. Los generacionales marcan época, van más allá del ego.
 *    El planeta me debe mostrar qué momento es.»
 *
 * ┌─ QUÉ CALCULA NOVA Y QUÉ TRAE ELLA ─────────────────────────┐
 * │                                                            │
 * │ NOVA CALCULA, sin internet y sin inventar:                 │
 * │   · la fase de la Luna de cualquier día                    │
 * │   · la ventana de su revolución solar: de cumpleaños a     │
 * │     cumpleaños, que es aritmética de calendario            │
 * │   · qué ventana del pensum está abierta hoy                │
 * │   · y —esto no lo hace ninguna app— SI LE FUNCIONÓ A ELLA  │
 * │                                                            │
 * │ ELLA TRAE, de Horus, que ya se lo da bien:                 │
 * │   · su carta natal                                         │
 * │   · los tránsitos de la semana, el mes y el año            │
 * │   · la carta de su revolución solar                        │
 * │                                                            │
 * │ Calcular efemérides aquí sería rehacer mal algo que ya     │
 * │ está bien hecho, y con una precisión que no se puede       │
 * │ verificar desde una hoja de cálculo.                       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * LA REGLA DE SIEMPRE, APLICADA AQUÍ
 *
 * Nova puede decir «hoy tienes Luna menguante y tu pensum dice que esta
 * temporada pide cerrar». Eso es aritmética sobre lo que ella cargó, y
 * es verdad dentro de su marco. Lo que Nova NO va a decir es que por eso
 * algo vaya a salir bien.
 *
 * Lo que sí hace, y es lo único que aquí es un dato y no una creencia:
 * cada tarea queda con la ventana en que se hizo, y a los tres meses
 * Nova puede decirle «las entregas que pusiste en menguante las
 * terminaste el 80% de las veces». Si el patrón no aparece, también se
 * lo dice.
 */

/**
 * Los cuerpos, por grupo. El grupo es lo que decide QUÉ tipo de momento
 * marca cada uno, con la división que ella misma hizo.
 */
const CIELO_CUERPOS = {
  sol:         { nombre: 'Sol',         grupo: 'personal' },
  luna:        { nombre: 'Luna',        grupo: 'personal' },
  mercurio:    { nombre: 'Mercurio',    grupo: 'personal' },
  venus:       { nombre: 'Venus',       grupo: 'personal' },
  marte:       { nombre: 'Marte',       grupo: 'personal' },
  jupiter:     { nombre: 'Júpiter',     grupo: 'social' },
  saturno:     { nombre: 'Saturno',     grupo: 'social' },
  urano:       { nombre: 'Urano',       grupo: 'generacional' },
  neptuno:     { nombre: 'Neptuno',     grupo: 'generacional' },
  pluton:      { nombre: 'Plutón',      grupo: 'generacional' },
  ascendente:  { nombre: 'Ascendente',  grupo: 'angulo' },
  medio_cielo: { nombre: 'Mediocielo',  grupo: 'angulo' },
  descendente: { nombre: 'Descendente', grupo: 'angulo' },
  fondo_cielo: { nombre: 'Fondo del cielo', grupo: 'angulo' },
  /**
   * Los puntos y asteroides van APARTE de sus tres grupos.
   *
   * Ella fue explícita: personales son Sol a Marte, sociales Júpiter y
   * Saturno, generacionales Urano, Neptuno y Plutón. Meter a Quirón o a
   * Ceres entre los generacionales diluiría su marco, y el marco es
   * suyo. Aquí están porque Horus se los da y no se van a perder, pero
   * en su propio cajón.
   */
  nodo_norte:  { nombre: 'Nodo Norte',  grupo: 'punto' },
  nodo_sur:    { nombre: 'Nodo Sur',    grupo: 'punto' },
  lilith:      { nombre: 'Lilith',      grupo: 'punto' },
  quiron:      { nombre: 'Quirón',      grupo: 'punto' },
  folo:        { nombre: 'Folo',        grupo: 'punto' },
  ceres:       { nombre: 'Ceres',       grupo: 'punto' },
  pallas:      { nombre: 'Pallas',      grupo: 'punto' },
  juno:        { nombre: 'Juno',        grupo: 'punto' },
  vesta:       { nombre: 'Vesta',       grupo: 'punto' },
  fortuna:     { nombre: 'Parte de la Fortuna', grupo: 'punto' },
  vertex:      { nombre: 'Vertex',      grupo: 'punto' },
};

/** Qué mira cada grupo. Son sus palabras, no las mías. */
const CIELO_GRUPOS = {
  personal: {
    nombre: 'Personales',
    que: 'Lo inmediato, lo mío, lo propio.',
    cuerpos: 'Sol, Luna, Mercurio, Venus, Marte' },
  social: {
    nombre: 'Sociales',
    que: 'Expansión y estructura.',
    cuerpos: 'Júpiter y Saturno' },
  generacional: {
    nombre: 'Generacionales',
    que: 'Marcan época. Van más allá del ego.',
    cuerpos: 'Urano, Neptuno, Plutón' },
  angulo: {
    nombre: 'Ángulos',
    que: 'Por dónde entras y hacia dónde apuntas.',
    cuerpos: 'Ascendente, Mediocielo, Descendente y Fondo del cielo' },
  punto: {
    nombre: 'Puntos y asteroides',
    que: 'Matices. No son el marco, lo afinan.',
    cuerpos: 'Nodos, Lilith, Quirón y los asteroides' },
};

/** Los tres momentos que ella nombró. */
const CIELO_MOMENTOS = {
  aprender:  { nombre: 'Aprender',  que: 'Entra información. Se estudia, se prueba, se pregunta.' },
  descansar: { nombre: 'Descansar', que: 'Se sostiene lo que hay. No se abre nada nuevo.' },
  cambiar:   { nombre: 'Cambiar',   que: 'Se cierra, se suelta, se mueve de sitio.' },
};

const CIELO_SIGNOS = ['aries','tauro','geminis','cancer','leo','virgo','libra',
                      'escorpio','sagitario','capricornio','acuario','piscis'];
const CIELO_SIGNOS_NOMBRE = ['Aries','Tauro','Géminis','Cáncer','Leo','Virgo','Libra',
                             'Escorpio','Sagitario','Capricornio','Acuario','Piscis'];

// ─── LA LUNA ─────────────────────────────────────────────────

/**
 * La fase de la Luna, calculada y no adivinada.
 *
 * Desde una luna nueva conocida —6 de enero de 2000, 18:14 UTC— y el
 * mes sinódico. Es la única pieza de astronomía que Nova hace sola,
 * porque es una cuenta de dos líneas que no se puede equivocar mucho:
 * el error acumulado en treinta años es de horas, y una fase dura días.
 *
 * El significado de cada fase es la lectura común y está aquí como
 * PUNTO DE PARTIDA. Lo que mande de verdad es lo que ella escriba en su
 * pensum: su marco antes que el de nadie.
 */
const LUNA_EPOCA_MS = Date.UTC(2000, 0, 6, 18, 14, 0);
const LUNA_SINODICO = 29.530588853;

const LUNA_FASES = [
  { id: 'nueva',            nombre: 'Luna nueva',        momento: 'aprender',
    que: 'Se siembra. Es para empezar algo, no para mostrarlo.' },
  { id: 'creciente',        nombre: 'Creciente',         momento: 'aprender',
    que: 'Lo que empezó toma cuerpo. Se construye.' },
  { id: 'cuarto_creciente', nombre: 'Cuarto creciente',  momento: 'cambiar',
    que: 'Aparece la resistencia. Es el punto donde se decide seguir o no.' },
  { id: 'gibosa',           nombre: 'Gibosa creciente',  momento: 'aprender',
    que: 'Se afina. Se corrige antes de mostrar.' },
  { id: 'llena',            nombre: 'Luna llena',        momento: 'cambiar',
    que: 'Culmina y se ve. Es para mostrar y para entregar.' },
  { id: 'diseminadora',     nombre: 'Gibosa menguante',  momento: 'descansar',
    que: 'Se comparte lo que salió. Se cuenta, se enseña.' },
  { id: 'cuarto_menguante', nombre: 'Cuarto menguante',  momento: 'cambiar',
    que: 'Se corta lo que no sirvió. Cierre con decisión.' },
  { id: 'balsamica',        nombre: 'Balsámica',         momento: 'descansar',
    que: 'Se suelta y se descansa. No es día de arrancar nada.' },
];

function faseLunar_(fechaISO) {
  const t = new Date(fechaISO + 'T12:00:00Z').getTime();
  let edad = ((t - LUNA_EPOCA_MS) / 86400000) % LUNA_SINODICO;
  if (edad < 0) edad += LUNA_SINODICO;

  // Ocho fases del mismo ancho, centradas en la nueva.
  const ancho = LUNA_SINODICO / 8;
  let i = Math.floor((edad + ancho / 2) / ancho);
  if (i >= 8) i = 0;
  const f = LUNA_FASES[i];

  return {
    id: f.id, nombre: f.nombre, momento: f.momento, que: f.que,
    edadDias: Math.round(edad * 10) / 10,
    // Iluminación: 0 en la nueva, 100 en la llena.
    iluminacion: Math.round((1 - Math.cos(2 * Math.PI * edad / LUNA_SINODICO)) / 2 * 100),
    creciendo: edad < LUNA_SINODICO / 2,
  };
}

// ─── LA REVOLUCIÓN SOLAR ─────────────────────────────────────

/**
 * Su año solar: de cumpleaños a cumpleaños.
 *
 * Esto sí es aritmética de calendario y se puede hacer sin efemérides.
 * El MOMENTO exacto del retorno del Sol se corre unas horas cada año y
 * ese dato lo trae ella de Horus; la ventana del año no depende de esas
 * horas, y es lo que sirve para repartir doce meses.
 */
function revolucionVentana_(nacimientoISO, hoyISO) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(nacimientoISO || ''))) return null;
  const mmdd = nacimientoISO.slice(5);
  const anioHoy = Number(hoyISO.slice(0, 4));
  // Si el cumpleaños de este año todavía no llega, el año solar vigente
  // arrancó el año pasado.
  const arranca = hoyISO.slice(5) >= mmdd ? anioHoy : anioHoy - 1;
  const desde = arranca + '-' + mmdd;
  const hasta = masDias_((arranca + 1) + '-' + mmdd, -1);
  const dias = Math.round((new Date(hasta + 'T00:00:00Z') -
                           new Date(hoyISO + 'T00:00:00Z')) / 86400000) + 1;
  return {
    anio: arranca, desde: desde, hasta: hasta,
    edad: arranca - Number(nacimientoISO.slice(0, 4)),
    diasRestantes: dias,
    transcurrido: Math.round(
      (new Date(hoyISO + 'T00:00:00Z') - new Date(desde + 'T00:00:00Z')) / 86400000 /
      ((new Date(masDias_(hasta, 1) + 'T00:00:00Z') -
        new Date(desde + 'T00:00:00Z')) / 86400000) * 100),
  };
}

// ─── LEER LO QUE PEGA DE HORUS ───────────────────────────────

function cuerpoDeTexto_(s) {
  const t = norm(s);
  if (!t) return '';
  /**
   * Los nombres tal como los escribe Horus, que es de donde vienen.
   *
   * «Nodo Sur» y «Nodo Norte» son dos palabras y empiezan igual: el
   * orden importa, porque buscar «nodo» primero convertiría el Nodo Sur
   * en Nodo Norte sin que nadie lo notara.
   */
  const alias = {
    'sol': 'sol', 'luna': 'luna', 'mercurio': 'mercurio', 'venus': 'venus',
    'marte': 'marte', 'jupiter': 'jupiter', 'saturno': 'saturno',
    'urano': 'urano', 'neptuno': 'neptuno', 'pluton': 'pluton',
    'quiron': 'quiron', 'chiron': 'quiron',
    'lilith': 'lilith', 'luna negra': 'lilith',
    'nodo norte': 'nodo_norte', 'nodo lunar': 'nodo_norte', 'nodo n': 'nodo_norte',
    'nodo sur': 'nodo_sur', 'nodo s': 'nodo_sur',
    'nodo': 'nodo_norte',
    'folo': 'folo', 'pholus': 'folo',
    'ceres': 'ceres', 'pallas': 'pallas', 'palas': 'pallas',
    'juno': 'juno', 'vesta': 'vesta', 'vertex': 'vertex',
    'parte de la fortuna': 'fortuna', 'fortuna': 'fortuna', 'rueda de la fortuna': 'fortuna',
    'ascendente': 'ascendente', 'asc': 'ascendente', 'as': 'ascendente',
    'medio cielo': 'medio_cielo', 'mediocielo': 'medio_cielo', 'mc': 'medio_cielo',
    'descendente': 'descendente', 'ds': 'descendente', 'dsc': 'descendente',
    'fondo del cielo': 'fondo_cielo', 'fondo de cielo': 'fondo_cielo',
    'ic': 'fondo_cielo', 'imum coeli': 'fondo_cielo',
  };
  return alias[t] || '';
}

/**
 * El signo, incluyendo como lo abrevia Horus: Vir, Lib, Esc, Cán…
 *
 * Sin las abreviaturas, pegar la pantalla de Horus tal cual no
 * encontraba NI UN signo, que es justo lo que ella iba a hacer.
 */
const CIELO_SIGNOS_ALIAS = {
  ari: 'aries', tau: 'tauro', gem: 'geminis', can: 'cancer', leo: 'leo',
  vir: 'virgo', lib: 'libra', esc: 'escorpio', sco: 'escorpio',
  sag: 'sagitario', cap: 'capricornio', acu: 'acuario', pis: 'piscis',
  escorpion: 'escorpio', geminis: 'geminis',
};

function signoDeTexto_(s) {
  const t = norm(s);
  if (!t) return '';
  if (CIELO_SIGNOS.indexOf(t) !== -1) return t;
  return CIELO_SIGNOS_ALIAS[t] || '';
}

/**
 * Busca un cuerpo dentro de un texto, DE LO LARGO A LO CORTO.
 *
 * El orden no es un detalle: «Nodo Sur Rx» empieza por «Nodo», y
 * buscando palabra por palabra el Nodo Sur se convertía en Nodo Norte
 * sin que nadie lo notara — y después se descartaba por repetido, así
 * que el Nodo Sur simplemente desaparecía de la carta. Probar primero
 * los grupos largos de palabras es lo que lo arregla, y sirve igual
 * para «Parte de la Fortuna» y «Fondo del cielo».
 */
function cuerpoEnTexto_(s) {
  const palabras = String(s || '').replace(/[^a-záéíóúñ ]/gi, ' ')
    .split(/\s+/).filter(function (w) { return w; });
  for (let largo = Math.min(4, palabras.length); largo >= 1; largo--) {
    for (let i = 0; i + largo <= palabras.length; i++) {
      const c = cuerpoDeTexto_(palabras.slice(i, i + largo).join(' '));
      if (c) return c;
    }
  }
  return '';
}

function signoEnTexto_(s) {
  const palabras = String(s || '').replace(/[^a-záéíóúñ ]/gi, ' ')
    .split(/\s+/).filter(function (w) { return w; });
  for (let i = 0; i < palabras.length; i++) {
    const g = signoDeTexto_(palabras[i]);
    if (g) return g;
  }
  return '';
}

/**
 * Lee una carta pegada y PROPONE las filas.
 *
 * Forgiving a propósito: cada app escribe la carta distinto y no vale
 * la pena pelear con el formato. Lo que se busca en cada línea es un
 * cuerpo y un signo; lo demás —grado, casa, retrógrado— se toma si está.
 *
 * Y como en todo lo demás: se muestra lo que NO se entendió. Una carta
 * a la que le falta Saturno en silencio es peor que una carta vacía.
 */
function cartaLeer_(texto) {
  const lineas = String(texto || '').split(/\r?\n/);
  const encontradas = [], ignoradas = [];
  const vistos = {};

  lineas.forEach(function (cruda) {
    const linea = String(cruda).replace(/\s+/g, ' ').trim();
    if (!linea) return;

    // Se parte por lo que suele separar: | , ; tabulador, «en», «·»
    const piezas = linea.split(/[|;,\t·]+|\ben\b/i).map(function (x) { return x.trim(); });
    let cuerpo = '', signo = '';
    piezas.forEach(function (x) {
      if (!cuerpo) cuerpo = cuerpoEnTexto_(x);
      if (!signo) signo = signoEnTexto_(x);
    });
    // Y si venían pegados en la misma pieza: «Sol Virgo 21°»
    if (!cuerpo) cuerpo = cuerpoEnTexto_(linea);
    if (!signo) signo = signoEnTexto_(linea);

    if (!cuerpo || !signo) {
      ignoradas.push({ linea: linea,
        porque: !cuerpo && !signo ? 'No encontré ni planeta ni signo.'
              : !cuerpo ? 'Encontré el signo pero no el planeta.'
              : 'Encontré el planeta pero no el signo.' });
      return;
    }
    if (vistos[cuerpo]) {
      ignoradas.push({ linea: linea, porque: 'Ya tenía ' +
        CIELO_CUERPOS[cuerpo].nombre + ' de una línea anterior.' });
      return;
    }
    vistos[cuerpo] = 1;

    /**
     * Horus escribe «Vir 27», sin el símbolo de grado. Así que si no
     * hay «°», se toma el primer número que NO sea el de la casa —
     * porque «Casa 9» también es un número y confundirlos pondría el
     * Sol a 9 grados en vez de a 27.
     */
    let gr = linea.match(/(\d{1,2})\s*[°º]\s*(\d{1,2})?/);
    if (!gr) {
      const sinCasa = linea.replace(/casa\s*\d{1,2}/ig, ' ');
      gr = sinCasa.match(/(?:^|[^\d])(\d{1,2})(?![\d])/);
      if (gr) gr = [gr[0], gr[1], null];
    }
    // La casa solo cuenta si la palabra «casa» está: un número suelto
    // puede ser el grado, y una casa inventada mueve el tema entero.
    const casa = linea.match(/casa\s*(\d{1,2})/i);
    encontradas.push({
      cuerpo: cuerpo, nombre: CIELO_CUERPOS[cuerpo].nombre,
      grupo: CIELO_CUERPOS[cuerpo].grupo,
      signo: signo, signoNombre: CIELO_SIGNOS_NOMBRE[CIELO_SIGNOS.indexOf(signo)],
      grado: gr ? Number(gr[1]) + (gr[2] ? Number(gr[2]) / 60 : 0) : null,
      casa: casa ? Number(casa[1]) : null,
      retrogrado: /\bR\b|\bRx\b|retr[oó]grad/i.test(linea),
      linea: linea,
    });
  });

  // El orden de siempre: personales, sociales, generacionales, ángulos.
  const orden = Object.keys(CIELO_CUERPOS);
  encontradas.sort(function (a, b) {
    return orden.indexOf(a.cuerpo) - orden.indexOf(b.cuerpo);
  });

  /**
   * Qué se esperaba y no llegó. Solo los diez planetas y el Ascendente:
   * los asteroides son opcionales y avisar de que falta Vesta sería
   * ruido que tapa el aviso de que falta Saturno.
   */
  const ESPERADOS = ['sol','luna','mercurio','venus','marte','jupiter','saturno',
                     'urano','neptuno','pluton','ascendente'];
  const faltan = ESPERADOS.filter(function (c) { return !vistos[c]; });

  return { encontradas: encontradas, ignoradas: ignoradas, faltan: faltan };
}

// ─── LA API ──────────────────────────────────────────────────

function cartaDe_(uid) {
  return soulLeerSuave_('Carta', uid, []).map(function (f) {
    const c = norm(f.cuerpo);
    const def = CIELO_CUERPOS[c] || { nombre: String(f.cuerpo || ''), grupo: 'generacional' };
    const s = norm(f.signo);
    return {
      cuerpo: c, nombre: def.nombre, grupo: def.grupo,
      signo: s, signoNombre: CIELO_SIGNOS.indexOf(s) !== -1
        ? CIELO_SIGNOS_NOMBRE[CIELO_SIGNOS.indexOf(s)] : String(f.signo || ''),
      grado: f.grado === '' || f.grado === null ? null : num(f.grado),
      casa: f.casa === '' || f.casa === null ? null : num(f.casa),
      retrogrado: norm(f.retrogrado) === 'si' || String(f.retrogrado).toLowerCase() === 'true',
      nota: String(f.nota || ''),
    };
  }).filter(function (x) { return x.cuerpo; });
}

/** La hoja Carta no tiene columna `id`: la llave es el cuerpo. */
function cartaGuardarFilas_(uid, filas) {
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Hay otro cambio guardándose.');
  try {
    const sh = soulSheet_('Carta');
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];
    const cU = enc.indexOf('usuario_id'), cC = enc.indexOf('cuerpo');
    let puestas = 0;

    filas.forEach(function (f) {
      const cuerpo = norm(f.cuerpo);
      if (!CIELO_CUERPOS[cuerpo]) return;
      const fila = enc.map(function (col) {
        switch (col) {
          case 'usuario_id': return uid;
          case 'cuerpo': return cuerpo;
          case 'signo': return norm(f.signo);
          case 'grado': return f.grado === null || f.grado === undefined ? '' : num(f.grado);
          case 'casa': return f.casa === null || f.casa === undefined ? '' : num(f.casa);
          case 'retrogrado': return f.retrogrado ? 'si' : 'no';
          case 'nota': return String(f.nota || '');
          default: return '';
        }
      });
      let en = -1;
      for (let i = 1; i < d.length; i++) {
        const suyo = String(d[i][cU] || '').toLowerCase().trim();
        if ((!suyo || suyo === uid) && norm(d[i][cC]) === cuerpo) { en = i; break; }
      }
      if (en === -1) { sh.appendRow(fila); d.push(fila); }
      else { d[en] = fila; sh.getRange(en + 1, 1, 1, enc.length).setValues([fila]); }
      puestas++;
    });
    return puestas;
  } finally { lock.releaseLock(); }
}

function soulCartaLeer(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const texto = String(p.texto || '');
  if (!texto.trim()) return { ok: false, error: 'Pega tu carta primero.' };
  const r = cartaLeer_(texto);
  return { ok: true, encontradas: r.encontradas, ignoradas: r.ignoradas,
           faltan: r.faltan.map(function (c) { return CIELO_CUERPOS[c].nombre; }),
           cuerpos: CIELO_CUERPOS, grupos: CIELO_GRUPOS,
           signos: CIELO_SIGNOS.map(function (x, i) {
             return { id: x, nombre: CIELO_SIGNOS_NOMBRE[i] }; }) };
}

function soulCartaGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const filas = p.filas || [];
  if (!filas.length) return { ok: false, error: 'No marcaste nada.' };
  try {
    const n = cartaGuardarFilas_(soulUsuario_(s), filas);
    return { ok: true, guardadas: n };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/** Los datos de nacimiento, en la hoja Usuarios. */
function soulNacimientoGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const fecha = String(p.fecha || '').trim();
  if (fecha && !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
    return { ok: false, error: 'La fecha va como AAAA-MM-DD.' };
  }
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro cambio guardándose.' };
  try {
    const sh = soulSheet_('Usuarios');
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];
    const cC = enc.indexOf('correo');
    let en = -1;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cC] || '').toLowerCase().trim() === uid) { en = i; break; }
    }
    const pon = function (fila) {
      const set = function (col, v) {
        const j = enc.indexOf(col);
        if (j !== -1 && v !== undefined) fila[j] = v;
      };
      set('id', fila[0] || uid);
      set('correo', uid);
      set('nombre', p.nombre !== undefined ? String(p.nombre) : undefined);
      set('fecha_nacimiento', fecha || undefined);
      set('hora_nacimiento', p.hora !== undefined ? String(p.hora) : undefined);
      set('lugar_nacimiento', p.lugar !== undefined ? String(p.lugar) : undefined);
      set('zona_horaria', p.zona !== undefined ? String(p.zona) : undefined);
      return fila;
    };
    if (en === -1) sh.appendRow(pon(enc.map(function () { return ''; })));
    else sh.getRange(en + 1, 1, 1, enc.length).setValues([pon(d[en])]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

function nacimientoDe_(uid) {
  const filas = soulLeerSuave_('Usuarios', uid, []).filter(function (f) {
    return String(f.correo || '').toLowerCase().trim() === uid;
  });
  const f = filas[0] || {};
  return {
    nombre: String(f.nombre || ''),
    fecha: aISO(f.fecha_nacimiento, 'UTC') || '',
    hora: typeof f.hora_nacimiento === 'string' ? f.hora_nacimiento
          : (f.hora_nacimiento ? hhmm_(horaNum_(f.hora_nacimiento)) : ''),
    lugar: String(f.lugar_nacimiento || ''),
    zona: String(f.zona_horaria || ''),
  };
}

// ─── TRÁNSITOS, PENSUM Y REVOLUCIÓN ──────────────────────────

function transitosDe_(uid) {
  return soulLeerSuave_('Transitos', uid, []).map(function (f) {
    const c = norm(f.cuerpo);
    const def = CIELO_CUERPOS[c] || { nombre: String(f.cuerpo || ''), grupo: 'generacional' };
    // `fecha` es lo viejo; `desde`/`hasta` es lo que sirve. Una fila sin
    // rango se trata como de un solo día en vez de descartarla.
    const desde = aISO(f.desde, 'UTC') || aISO(f.fecha, 'UTC') || '';
    return {
      cuerpo: c, nombre: def.nombre, grupo: def.grupo,
      aspecto: String(f.aspecto || ''), aNatal: String(f.a_natal || ''),
      casa: f.casa === '' ? null : num(f.casa),
      tema: String(f.tema || ''),
      intensidad: f.intensidad_pct === '' ? null : num(f.intensidad_pct),
      desde: desde, hasta: aISO(f.hasta, 'UTC') || desde,
      texto: String(f.texto_transito || ''),
      porQue: String(f.por_que || ''),
      como: String(f.como_trabajarlo || ''),
      elOtroLado: String(f.el_otro_lado || ''),
      fuente: String(f.fuente || ''),
    };
  }).filter(function (t) { return t.desde; });
}

function soulTransitoGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const d = p.datos || {};
  const cuerpo = norm(d.cuerpo);
  if (!CIELO_CUERPOS[cuerpo]) return { ok: false, error: 'Ese cuerpo no está en la lista.' };
  const desde = String(d.desde || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(desde)) {
    return { ok: false, error: 'Un tránsito sin fecha de inicio no se puede cruzar con nada.' };
  }
  const hasta = String(d.hasta || '').trim() || desde;
  if (hasta < desde) return { ok: false, error: 'La fecha de fin va después de la de inicio.' };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro cambio guardándose.' };
  try {
    const sh = soulSheet_('Transitos');
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const v = { usuario_id: soulUsuario_(s), fecha: desde, desde: desde, hasta: hasta,
                cuerpo: cuerpo, aspecto: String(d.aspecto || ''),
                a_natal: String(d.a_natal || ''),
                casa: d.casa ? num(d.casa) : '', tema: String(d.tema || ''),
                intensidad_pct: d.intensidad ? num(d.intensidad) : '',
                texto_transito: String(d.texto || ''), por_que: String(d.porQue || ''),
                como_trabajarlo: String(d.como || ''),
                el_otro_lado: String(d.elOtroLado || ''),
                fuente: String(d.fuente || 'Horus') };
    sh.appendRow(enc.map(function (c) { return v[c] !== undefined ? v[c] : ''; }));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

function pensumDe_(uid) {
  return soulLeerSuave_('Pensum', uid, [])
    .filter(function (f) { return String(f.id || '').trim(); })
    .map(function (f) {
      const c = norm(f.cuerpo);
      const def = CIELO_CUERPOS[c] || { nombre: String(f.cuerpo || ''), grupo: '' };
      const m = norm(f.momento);
      return {
        id: String(f.id), desde: aISO(f.desde, 'UTC') || '', hasta: aISO(f.hasta, 'UTC') || '',
        titulo: String(f.titulo || ''), cuerpo: c, cuerpoNombre: def.nombre,
        grupo: def.grupo, casa: f.casa === '' ? null : num(f.casa),
        momento: CIELO_MOMENTOS[m] ? m : '',
        quePide: String(f.que_pide || ''), queEvitar: String(f.que_evitar || ''),
        nota: String(f.nota || ''),
      };
    })
    .sort(function (a, b) { return (a.desde || '9') < (b.desde || '9') ? -1 : 1; });
}

function soulPensumGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const d = p.datos || {};
  const esNuevo = !String(d.id || '').trim();
  if (esNuevo && !String(d.titulo || '').trim()) {
    return { ok: false, error: 'Ponle un título a la temporada.' };
  }
  if (esNuevo && !/^\d{4}-\d{2}-\d{2}$/.test(String(d.desde || ''))) {
    return { ok: false, error: 'Falta desde cuándo va.' };
  }
  if (d.desde && d.hasta && String(d.hasta) < String(d.desde)) {
    return { ok: false, error: 'La fecha de fin va después de la de inicio.' };
  }
  const m = norm(d.momento);
  if (m && !CIELO_MOMENTOS[m]) return { ok: false, error: 'Ese momento no existe.' };
  try {
    soulGuardar_('Pensum', {
      id: String(d.id || ''),
      desde: d.desde !== undefined ? String(d.desde) : undefined,
      hasta: d.hasta !== undefined ? String(d.hasta) : undefined,
      titulo: d.titulo !== undefined ? String(d.titulo).trim() : undefined,
      cuerpo: d.cuerpo !== undefined ? norm(d.cuerpo) : undefined,
      casa: d.casa !== undefined ? num(d.casa) : undefined,
      momento: m || undefined,
      que_pide: d.quePide !== undefined ? String(d.quePide) : undefined,
      que_evitar: d.queEvitar !== undefined ? String(d.queEvitar) : undefined,
      nota: d.nota !== undefined ? String(d.nota) : undefined,
    }, soulUsuario_(s));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulPensumBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  try {
    const fue = soulBorrar_('Pensum', p.id, soulUsuario_(s));
    return { ok: fue, error: fue ? '' : 'No lo encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulRevolucionGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const anio = num(p.anio);
  if (!anio) return { ok: false, error: '¿De qué año es la revolución?' };
  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro cambio guardándose.' };
  try {
    const sh = soulSheet_('Revolucion');
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];
    const cU = enc.indexOf('usuario_id'), cA = enc.indexOf('anio');
    const v = { usuario_id: uid, anio: anio, desde: String(p.desde || ''),
                hasta: String(p.hasta || ''), ascendente: norm(p.ascendente),
                casa_sol: p.casaSol ? num(p.casaSol) : '',
                tema: String(p.tema || ''), texto: String(p.texto || ''),
                nota: (p.planetas || []).filter(function (x) {
                  return CIELO_CUERPOS[norm(x.cuerpo)] && num(x.casa) >= 1 && num(x.casa) <= 12;
                }).map(function (x) {
                  return norm(x.cuerpo) + ':' + num(x.casa);
                }).join(',') || String(p.nota || '') };
    const fila = enc.map(function (c) { return v[c] !== undefined ? v[c] : ''; });
    let en = -1;
    for (let i = 1; i < d.length; i++) {
      const suyo = String(d[i][cU] || '').toLowerCase().trim();
      if ((!suyo || suyo === uid) && num(d[i][cA]) === anio) { en = i; break; }
    }
    if (en === -1) sh.appendRow(fila);
    else sh.getRange(en + 1, 1, 1, enc.length).setValues([fila]);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

// ─── LA LECTURA DE LA REVOLUCIÓN ─────────────────────────────

/**
 * ┌─ POR QUÉ ESTO SE PUEDE ESCRIBIR Y LA IMAGEN NO ────────────┐
 * │                                                            │
 * │ Nova NO puede leer una captura de Horus. Sacar texto de    │
 * │ una imagen y entenderlo necesita un modelo con visión, y   │
 * │ Apps Script no tiene ninguno. Pegar la foto no va a        │
 * │ funcionar, y prefiero decirlo antes que dejar un botón     │
 * │ que no hace nada.                                          │
 * │                                                            │
 * │ Lo que SÍ se puede: ella teclea DOS datos —el ascendente   │
 * │ del año y en qué casa cae su Sol— y Nova compone la        │
 * │ lectura desde una tabla. No es una interpretación          │
 * │ inventada por encargo: es la regla escrita, la misma       │
 * │ siempre, y ella la puede leer, discutir y corregir.        │
 * │                                                            │
 * │ Una lectura generada de nuevo cada vez diría algo distinto │
 * │ el martes que el jueves con los mismos datos. Esta no.     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * Las doce casas, y qué momento marca cada una.
 *
 * El momento sale de la división tradicional, que encaja con la suya:
 *
 *   ANGULARES   (1, 4, 7, 10)  se actúa y se mueve  → CAMBIAR
 *   SUCEDENTES  (2, 5, 8, 11)  se sostiene y cuaja  → DESCANSAR
 *   CADENTES    (3, 6, 9, 12)  se ajusta y se sabe  → APRENDER
 *
 * No es una regla que me inventé para que cuadrara: es como se leen las
 * casas desde hace siglos, y da la casualidad de que sus tres momentos
 * son exactamente esas tres clases.
 */
const CIELO_CASAS = [
  { n: 1,  clase: 'angular',   momento: 'cambiar',
    area: 'Tú',              que: 'Cómo te presentas y qué cuerpo le pones al año.' },
  { n: 2,  clase: 'sucedente', momento: 'descansar',
    area: 'Lo que tienes',   que: 'Plata propia, recursos, lo que sostiene.' },
  { n: 3,  clase: 'cadente',   momento: 'aprender',
    area: 'Lo que aprendes', que: 'Estudio, escritura, lo cercano, los hermanos.' },
  { n: 4,  clase: 'angular',   momento: 'cambiar',
    area: 'La casa',         que: 'Familia, raíz, dónde vives y de dónde vienes.' },
  { n: 5,  clase: 'sucedente', momento: 'descansar',
    area: 'Lo que creas',    que: 'Creación, juego, lo que sale de ti y te gusta.' },
  { n: 6,  clase: 'cadente',   momento: 'aprender',
    area: 'El trabajo diario', que: 'Rutina, salud, el oficio de todos los días.' },
  { n: 7,  clase: 'angular',   momento: 'cambiar',
    area: 'Los otros',       que: 'Sociedades, contratos, el uno a uno.' },
  { n: 8,  clase: 'sucedente', momento: 'descansar',
    area: 'Lo compartido',   que: 'Plata de otros, deudas, lo que se transforma.' },
  { n: 9,  clase: 'cadente',   momento: 'aprender',
    area: 'Lo que amplía',   que: 'Estudio mayor, viajes, publicar, lo que abre mundo.' },
  { n: 10, clase: 'angular',   momento: 'cambiar',
    area: 'Lo público',      que: 'Carrera, visibilidad, lo que te reconocen.' },
  { n: 11, clase: 'sucedente', momento: 'descansar',
    area: 'La gente',        que: 'Red, comunidad, proyectos con otros, lo que viene.' },
  { n: 12, clase: 'cadente',   momento: 'aprender',
    area: 'Lo de adentro',   que: 'Retiro, cierre de ciclo, lo que se trabaja a solas.' },
];

/** El ascendente del año dice CÓMO se entra, no qué pasa. */
const REV_ASCENDENTE = {
  aries:       'Se entra empujando. El año premia arrancar, no preparar.',
  tauro:       'Se entra despacio y con cuerpo. Pide construir algo que dure, sin apuro.',
  geminis:     'Se entra preguntando. Año de moverse, escribir, hablar con mucha gente.',
  cancer:      'Se entra hacia adentro. Casa, gente cercana, lo que te sostiene.',
  leo:         'Se entra mostrándose. Lo que hagas este año se va a ver.',
  virgo:       'Se entra ordenando. Año de afinar el método, no de inventar de cero.',
  libra:       'Se entra de a dos. Lo que pase, pasa con alguien más.',
  escorpio:    'Se entra removiendo. Año de cerrar en serio lo que estaba a medias.',
  sagitario:   'Se entra abriendo. Pide más mundo: estudiar, viajar, publicar.',
  capricornio: 'Se entra con estructura. Año de responsabilidad y de resultados medibles.',
  acuario:     'Se entra rompiendo. Lo que cambies este año no se devuelve.',
  piscis:      'Se entra soltando. Año de cerrar ciclo antes de empezar otro.',
};

/**
 * La lectura, compuesta.
 *
 * Devuelve las partes por separado —el cómo, el dónde, los grupos— para
 * que la pantalla las muestre como bloques y no como un párrafo. Y
 * devuelve SIEMPRE qué le faltó para poder decir más.
 */
function revolucionLectura_(asc, casaSol, planetas) {
  const partes = [], faltan = [];
  const a = norm(asc);
  const casa = CIELO_CASAS.filter(function (c) { return c.n === num(casaSol); })[0] || null;

  if (a && REV_ASCENDENTE[a]) {
    partes.push({ clave: 'ascendente', titulo: 'Cómo entras al año',
      valor: CIELO_SIGNOS_NOMBRE[CIELO_SIGNOS.indexOf(a)],
      texto: REV_ASCENDENTE[a] });
  } else {
    faltan.push('el ascendente del año');
  }

  if (casa) {
    partes.push({ clave: 'sol', titulo: 'Dónde va tu atención',
      valor: 'Casa ' + casa.n + ' · ' + casa.area,
      texto: casa.que + ' Es una casa ' + casa.clase + ', así que el año pide ' +
        (CIELO_MOMENTOS[casa.momento] || {}).nombre.toLowerCase() + '.' });
  } else {
    faltan.push('en qué casa cae tu Sol');
  }

  /**
   * Los planetas por casa, agrupados con SU división. Cada grupo dice
   * qué mira, que es lo que ella pidió: los personales lo inmediato,
   * los sociales la expansión y la estructura, los generacionales la
   * época.
   */
  const grupos = [];
  ['personal', 'social', 'generacional'].forEach(function (g) {
    const suyos = (planetas || []).filter(function (x) {
      const def = CIELO_CUERPOS[norm(x.cuerpo)];
      return def && def.grupo === g && num(x.casa) >= 1 && num(x.casa) <= 12;
    });
    if (!suyos.length) return;
    grupos.push({
      grupo: g, nombre: CIELO_GRUPOS[g].nombre, que: CIELO_GRUPOS[g].que,
      cuerpos: suyos.map(function (x) {
        const c = CIELO_CASAS.filter(function (k) { return k.n === num(x.casa); })[0];
        return { cuerpo: norm(x.cuerpo), nombre: CIELO_CUERPOS[norm(x.cuerpo)].nombre,
                 casa: num(x.casa), area: c ? c.area : '', que: c ? c.que : '',
                 momento: c ? c.momento : '' };
      }),
    });
  });
  if (!grupos.length) faltan.push('en qué casa cae cada planeta (opcional, pero es lo que más dice)');

  /**
   * El momento del año: el que más se repite entre el Sol y los
   * planetas que tengan casa. Si hay empate, manda el del Sol — es su
   * año, y el Sol es de quien es el año.
   */
  const votos = {};
  if (casa) votos[casa.momento] = 2;
  grupos.forEach(function (g) {
    g.cuerpos.forEach(function (c) {
      if (c.momento) votos[c.momento] = (votos[c.momento] || 0) + 1;
    });
  });
  let momento = '';
  Object.keys(votos).forEach(function (k) {
    if (!momento || votos[k] > votos[momento]) momento = k;
  });

  return {
    partes: partes, grupos: grupos, momento: momento,
    momentoNombre: (CIELO_MOMENTOS[momento] || {}).nombre || '',
    faltan: faltan,
    // Sin los dos datos base no hay lectura. Se dice, no se rellena.
    hay: partes.length > 0,
    porque: partes.length ? '' :
      'Para darte la lectura necesito dos datos de Horus: el ascendente del año y en qué ' +
      'casa cae tu Sol. Con eso compongo el resto.',
  };
}

/**
 * Propone temporadas de pensum desde los tránsitos que ella ya cargó.
 *
 * No inventa fechas: usa las que trajo de Horus. Lo que agrega es el
 * momento, sacado de la casa por la que pasa. Y propone — ella confirma,
 * como en todo lo demás.
 */
function pensumProponer_(uid) {
  const transitos = transitosDe_(uid);
  const yaEstan = {};
  pensumDe_(uid).forEach(function (x) {
    yaEstan[norm(x.titulo) + '|' + x.desde] = 1;
  });

  return transitos.filter(function (t) {
    /**
     * Solo los que duran. Un tránsito de la Luna dura dos días y medio
     * y no es una temporada: meterlo en el pensum llenaría la lista de
     * ruido y taparía a Saturno, que es el que de verdad marca meses.
     */
    const dias = Math.round((new Date(t.hasta + 'T00:00:00Z') -
                             new Date(t.desde + 'T00:00:00Z')) / 86400000) + 1;
    return dias >= 14;
  }).map(function (t) {
    const c = CIELO_CASAS.filter(function (k) { return k.n === num(t.casa); })[0];
    const titulo = t.nombre + (t.casa ? ' por casa ' + t.casa : '') +
                   (t.aNatal ? ' a ' + t.aNatal : '');
    return {
      titulo: titulo, desde: t.desde, hasta: t.hasta,
      cuerpo: t.cuerpo, casa: t.casa,
      momento: c ? c.momento : '',
      quePide: t.tema || (c ? c.que : ''),
      grupo: t.grupo,
      dias: Math.round((new Date(t.hasta + 'T00:00:00Z') -
                        new Date(t.desde + 'T00:00:00Z')) / 86400000) + 1,
      yaEsta: !!yaEstan[norm(titulo) + '|' + t.desde],
    };
  }).sort(function (a, b) { return a.desde < b.desde ? -1 : 1; });
}

/** Guarda las temporadas que ella confirmó de la propuesta. */
function soulPensumDesdeTransitos(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const items = p.items || [];
  if (!items.length) return { ok: false, error: 'No marcaste ninguna temporada.' };

  const yaEstan = {};
  pensumDe_(uid).forEach(function (x) { yaEstan[norm(x.titulo) + '|' + x.desde] = 1; });

  let creadas = 0, repetidas = 0;
  items.forEach(function (it) {
    const titulo = String(it.titulo || '').trim();
    const desde = String(it.desde || '').trim();
    if (!titulo || !/^\d{4}-\d{2}-\d{2}$/.test(desde)) return;
    if (yaEstan[norm(titulo) + '|' + desde]) { repetidas++; return; }
    try {
      soulGuardar_('Pensum', {
        titulo: titulo, desde: desde, hasta: String(it.hasta || ''),
        cuerpo: norm(it.cuerpo), casa: it.casa ? num(it.casa) : '',
        momento: CIELO_MOMENTOS[norm(it.momento)] ? norm(it.momento) : '',
        que_pide: String(it.quePide || ''),
        nota: 'Propuesta desde un tránsito',
      }, uid);
      yaEstan[norm(titulo) + '|' + desde] = 1;
      creadas++;
    } catch (e) { /* una fila mala no tumba las demás */ }
  });
  return { ok: true, creadas: creadas, repetidas: repetidas };
}

// ─── LO QUE RIGE EL AÑO Y EL MES ─────────────────────────────

/**
 * ┌─ POR QUÉ LAS PROFECCIONES Y NO OTRA COSA ──────────────────┐
 * │                                                            │
 * │ Ella pidió «algo que rija mi año», en tiempos distintos.   │
 * │ Las profecciones son la única técnica que da eso y que     │
 * │ además es ARITMÉTICA PURA: una casa por año de vida, una   │
 * │ casa por mes. No hace falta ninguna efeméride, no hay nada │
 * │ que adivinar, y con los mismos datos da siempre lo mismo.  │
 * │                                                            │
 * │ Así queda su marco en tres escalas:                        │
 * │   EL AÑO   la casa de profección + su revolución solar     │
 * │   EL MES   la casa que le toca a ese mes                   │
 * │   EL DÍA   la luna y la temporada de su pensum             │
 * │                                                            │
 * │ Se usan CASAS ENTERAS —un signo, una casa— porque es como  │
 * │ se hacen las profecciones desde siempre, y porque no       │
 * │ depende de la hora exacta de nacimiento.                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Los regentes tradicionales. Donde el moderno difiere, se dice. */
const CIELO_REGENTES = {
  aries:       { cuerpo: 'marte',    moderno: '' },
  tauro:       { cuerpo: 'venus',    moderno: '' },
  geminis:     { cuerpo: 'mercurio', moderno: '' },
  cancer:      { cuerpo: 'luna',     moderno: '' },
  leo:         { cuerpo: 'sol',      moderno: '' },
  virgo:       { cuerpo: 'mercurio', moderno: '' },
  libra:       { cuerpo: 'venus',    moderno: '' },
  escorpio:    { cuerpo: 'marte',    moderno: 'pluton' },
  sagitario:   { cuerpo: 'jupiter',  moderno: '' },
  capricornio: { cuerpo: 'saturno',  moderno: '' },
  acuario:     { cuerpo: 'saturno',  moderno: 'urano' },
  piscis:      { cuerpo: 'jupiter',  moderno: 'neptuno' },
};

/** Qué signo cae en cada casa, contando de a uno desde el Ascendente. */
function signoDeCasa_(signoAsc, casa) {
  const i = CIELO_SIGNOS.indexOf(norm(signoAsc));
  if (i === -1 || !(casa >= 1 && casa <= 12)) return '';
  return CIELO_SIGNOS[(i + casa - 1) % 12];
}

function casaInfo_(n) {
  return CIELO_CASAS.filter(function (c) { return c.n === n; })[0] || null;
}

/**
 * La casa que rige el año y la que rige el mes.
 *
 * Una casa por año cumplido: a los 0 manda la 1, a los 12 vuelve a la 1.
 * Dentro del año, una casa por mes solar, empezando por la del año.
 */
function profecciones_(carta, nacimientoISO, hoyISO, ventana) {
  const asc = (carta || []).filter(function (c) { return c.cuerpo === 'ascendente'; })[0];
  if (!asc || !asc.signo) {
    return { hay: false,
      porque: 'Para saber qué casa rige tu año necesito tu Ascendente, y todavía no ' +
              'está en tu carta.' };
  }
  if (!ventana) {
    return { hay: false,
      porque: 'Necesito tu fecha de nacimiento para contar los años cumplidos.' };
  }

  const edad = ventana.edad;
  const casaAnual = (edad % 12) + 1;
  const signoAnual = signoDeCasa_(asc.signo, casaAnual);
  const reg = CIELO_REGENTES[signoAnual] || null;
  const regenteEn = reg
    ? (carta || []).filter(function (c) { return c.cuerpo === reg.cuerpo; })[0] || null
    : null;

  /**
   * El mes solar. Doce tramos iguales entre cumpleaños y cumpleaños —
   * no meses de calendario, que no empiezan el día de su cumpleaños y
   * darían un corte distinto cada año.
   */
  const largo = Math.round((new Date(masDias_(ventana.hasta, 1) + 'T00:00:00Z') -
                            new Date(ventana.desde + 'T00:00:00Z')) / 86400000);
  const tramo = largo / 12;
  const pasados = Math.round((new Date(hoyISO + 'T00:00:00Z') -
                              new Date(ventana.desde + 'T00:00:00Z')) / 86400000);
  let idx = Math.floor(pasados / tramo);
  if (idx < 0) idx = 0;
  if (idx > 11) idx = 11;

  const casaMes = ((casaAnual - 1 + idx) % 12) + 1;
  const desdeMes = masDias_(ventana.desde, Math.round(idx * tramo));
  const hastaMes = masDias_(ventana.desde, Math.round((idx + 1) * tramo) - 1);

  const cA = casaInfo_(casaAnual), cM = casaInfo_(casaMes);
  const nombreSigno = function (s) {
    const i = CIELO_SIGNOS.indexOf(s);
    return i === -1 ? '' : CIELO_SIGNOS_NOMBRE[i];
  };

  return {
    hay: true, edad: edad, ascendente: asc.signo, ascendenteNombre: nombreSigno(asc.signo),
    anual: {
      casa: casaAnual, signo: signoAnual, signoNombre: nombreSigno(signoAnual),
      area: cA ? cA.area : '', que: cA ? cA.que : '',
      momento: cA ? cA.momento : '', clase: cA ? cA.clase : '',
      regente: reg ? reg.cuerpo : '',
      regenteNombre: reg ? CIELO_CUERPOS[reg.cuerpo].nombre : '',
      regenteModerno: reg && reg.moderno ? CIELO_CUERPOS[reg.moderno].nombre : '',
      // Dónde está ese regente en SU carta: es lo que vuelve el año suyo
      // y no el de cualquiera que tenga la misma edad.
      regenteEn: regenteEn
        ? { signo: regenteEn.signo, signoNombre: regenteEn.signoNombre,
            casa: regenteEn.casa, retrogrado: regenteEn.retrogrado }
        : null,
      desde: ventana.desde, hasta: ventana.hasta,
    },
    mes: {
      indice: idx + 1, casa: casaMes,
      signo: signoDeCasa_(asc.signo, casaMes),
      signoNombre: nombreSigno(signoDeCasa_(asc.signo, casaMes)),
      area: cM ? cM.area : '', que: cM ? cM.que : '',
      momento: cM ? cM.momento : '', clase: cM ? cM.clase : '',
      desde: desdeMes, hasta: hastaMes,
      diasRestantes: Math.round((new Date(hastaMes + 'T00:00:00Z') -
                                 new Date(hoyISO + 'T00:00:00Z')) / 86400000) + 1,
    },
    // Los doce meses del año, para poder ver el año entero de una.
    calendario: (function () {
      const out = [];
      for (let k = 0; k < 12; k++) {
        const c = ((casaAnual - 1 + k) % 12) + 1;
        const info = casaInfo_(c);
        out.push({
          indice: k + 1, casa: c,
          desde: masDias_(ventana.desde, Math.round(k * tramo)),
          hasta: masDias_(ventana.desde, Math.round((k + 1) * tramo) - 1),
          area: info ? info.area : '', momento: info ? info.momento : '',
          esAhora: k === idx,
        });
      }
      return out;
    })(),
    porque: '',
  };
}

// ─── SI LE FUNCIONÓ A ELLA ───────────────────────────────────

/**
 * Lo único que aquí es un dato y no una creencia.
 *
 * Cada pendiente hecho tiene fecha de entrega y fecha de cierre. Con
 * eso se puede contar, fase por fase, cuántos se cumplieron. A los tres
 * meses eso deja de ser una impresión y pasa a ser su propio número.
 *
 * NO SE HABLA SIN MUESTRA. Con menos de `MINIMO` tareas en una fase, se
 * dice cuántas faltan en vez de dar un porcentaje. Un 100% sobre dos
 * casos no es un patrón: es una casualidad con decimales.
 */
const CIELO_MINIMO = 8;

function cieloMedir_(uid, hoyISO) {
  const pend = soulLeerSuave_('Pendientes', uid, []);
  const por = {};
  LUNA_FASES.forEach(function (f) { por[f.id] = { id: f.id, nombre: f.nombre, n: 0, hechas: 0 }; });

  let conFecha = 0;
  pend.forEach(function (f) {
    const fecha = aISO(f.fecha, 'UTC') || '';
    if (!fecha) return;
    // Solo lo que ya pasó: una entrega de la semana que viene no se ha
    // cumplido ni incumplido todavía.
    if (fecha > hoyISO) return;
    conFecha++;
    const fase = faseLunar_(fecha);
    por[fase.id].n++;
    if (soulHecho_(f)) por[fase.id].hechas++;
  });

  const fases = LUNA_FASES.map(function (f) {
    const x = por[f.id];
    return {
      id: f.id, nombre: f.nombre, momento: f.momento,
      n: x.n, hechas: x.hechas,
      // Sin muestra no hay porcentaje. Se dice cuántas faltan.
      pct: x.n >= CIELO_MINIMO ? Math.round(x.hechas / x.n * 100) : null,
      faltan: x.n >= CIELO_MINIMO ? 0 : CIELO_MINIMO - x.n,
    };
  });

  const conMuestra = fases.filter(function (f) { return f.pct !== null; });
  const total = fases.reduce(function (a, f) { return a + f.n; }, 0);
  const hechasTotal = fases.reduce(function (a, f) { return a + f.hechas; }, 0);
  const promedio = total ? Math.round(hechasTotal / total * 100) : null;

  /**
   * Un patrón solo se nombra si se separa del promedio de verdad. Diez
   * puntos sobre su propia media, no sobre cero: si ella cumple el 70%
   * de todo, un 72% en menguante no es nada.
   */
  let patron = '';
  if (conMuestra.length >= 2 && promedio !== null) {
    const mejor = conMuestra.slice().sort(function (a, b) { return b.pct - a.pct; })[0];
    const peor = conMuestra.slice().sort(function (a, b) { return a.pct - b.pct; })[0];
    if (mejor.pct - promedio >= 10) {
      patron = 'Las entregas que pusiste en ' + mejor.nombre.toLowerCase() +
        ' las terminaste el ' + mejor.pct + '% de las veces, contra un ' + promedio +
        '% en general.';
    }
    if (promedio - peor.pct >= 10) {
      patron += (patron ? ' ' : '') + 'En ' + peor.nombre.toLowerCase() +
        ' bajas al ' + peor.pct + '%.';
    }
  }

  return {
    minimo: CIELO_MINIMO, conFecha: conFecha, promedio: promedio,
    fases: fases, conMuestra: conMuestra.length, patron: patron,
    // Cuando no hay para hablar, se dice qué falta para poder hablar.
    porque: conMuestra.length < 2
      ? 'Todavía no hay muestra suficiente. Necesito al menos ' + CIELO_MINIMO +
        ' entregas con fecha en dos fases distintas antes de decirte nada — ' +
        'un patrón sobre tres casos no es un patrón.'
      : (patron ? '' : 'Miré tus entregas fase por fase y no encontré diferencia ' +
         'que valga la pena nombrar. Eso también es un resultado.'),
  };
}

// ─── LA FOTO DEL CIELO ───────────────────────────────────────

/**
 * Todo el cielo en una sola llamada: la carta, la luna, el pensum
 * abierto, los tránsitos vigentes, el año solar y la medición.
 */
function soulCielo(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);
  const lunes = lunesDe_(hoy);

  const nac = nacimientoDe_(uid);
  const carta = cartaDe_(uid);
  const transitos = transitosDe_(uid);
  const pensum = pensumDe_(uid);

  const vigente = function (x, f) {
    return x.desde && x.desde <= f && (!x.hasta || x.hasta >= f);
  };

  // La semana, día por día: la luna, lo que está abierto, y qué momento es.
  const dias = [];
  for (let i = 0; i < 7; i++) {
    const f = masDias_(lunes, i);
    const luna = faseLunar_(f);
    const abiertos = pensum.filter(function (x) { return vigente(x, f); });
    const trans = transitos.filter(function (x) { return vigente(x, f); });
    /**
     * Qué momento es. Manda SU pensum; la luna solo habla cuando ella
     * no escribió nada para esa temporada. Su marco antes que el común.
     */
    const delPensum = abiertos.filter(function (x) { return x.momento; })[0];
    dias.push({
      fecha: f, esHoy: f === hoy,
      luna: luna,
      momento: delPensum ? delPensum.momento : luna.momento,
      deDonde: delPensum ? 'pensum' : 'luna',
      porque: delPensum ? delPensum.titulo : luna.nombre,
      pensum: abiertos.map(function (x) { return x.id; }),
      transitos: trans.length,
    });
  }

  const rev = revolucionVentana_(nac.fecha, hoy);
  let revGuardada = null;
  try {
    revGuardada = soulLeerSuave_('Revolucion', uid, [])
      .filter(function (f) { return rev && num(f.anio) === rev.anio; })
      .map(function (f) {
        /**
         * Las casas de los planetas viajan en `nota` como «marte:7,venus:3».
         * Van ahí y no en columnas nuevas porque son hasta trece valores
         * y una hoja con trece columnas más se vuelve ilegible a mano —
         * y esta hoja ella la va a abrir a mano.
         */
        const casas = [];
        String(f.nota || '').split(',').forEach(function (par) {
          const x = par.split(':');
          const c = norm(x[0]);
          if (CIELO_CUERPOS[c] && num(x[1]) >= 1 && num(x[1]) <= 12) {
            casas.push({ cuerpo: c, casa: num(x[1]) });
          }
        });
        return { anio: num(f.anio), ascendente: String(f.ascendente || ''),
                 casaSol: f.casa_sol === '' ? null : num(f.casa_sol),
                 tema: String(f.tema || ''), texto: String(f.texto || ''),
                 nota: String(f.nota || ''), planetas: casas };
      })[0] || null;
  } catch (e) { /* la hoja puede no estar */ }

  const porGrupo = {};
  Object.keys(CIELO_GRUPOS).forEach(function (g) {
    porGrupo[g] = carta.filter(function (c) { return c.grupo === g; });
  });

  return {
    ok: true, hoy: hoy,
    nacimiento: nac,
    cuerpos: CIELO_CUERPOS, grupos: CIELO_GRUPOS, momentos: CIELO_MOMENTOS,
    signos: CIELO_SIGNOS.map(function (x, i) {
      return { id: x, nombre: CIELO_SIGNOS_NOMBRE[i] }; }),
    fases: LUNA_FASES,
    carta: carta, cartaPorGrupo: porGrupo,
    // Sin carta no se finge una: se dice qué falta y cómo se carga.
    tieneCarta: carta.length > 0,
    lunaHoy: faseLunar_(hoy),
    semana: { lunes: lunes, dias: dias },
    pensum: pensum,
    pensumAbierto: pensum.filter(function (x) { return vigente(x, hoy); }),
    transitos: transitos.sort(function (a, b) { return a.desde < b.desde ? -1 : 1; }),
    transitosHoy: transitos.filter(function (x) { return vigente(x, hoy); }),
    revolucion: rev ? Object.assign({}, rev, {
      carta: revGuardada,
      lectura: revolucionLectura_(
        revGuardada ? revGuardada.ascendente : '',
        revGuardada ? revGuardada.casaSol : 0,
        revGuardada && revGuardada.planetas ? revGuardada.planetas : []),
    }) : null,
    casas: CIELO_CASAS,
    momentos_: CIELO_MOMENTOS,
    /**
     * Lo que rige el año y el mes. Aritmética pura: una casa por año
     * cumplido y una por mes solar, sin efemérides de por medio.
     */
    profecciones: profecciones_(carta, nac.fecha, hoy, rev),
    // Las temporadas que se podrían armar solas desde sus tránsitos.
    pensumPropuesto: pensumProponer_(uid),
    medicion: cieloMedir_(uid, hoy),
  };
}
