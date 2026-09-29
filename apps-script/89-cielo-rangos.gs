/**
 * ═══════════════════════════════════════════════════════════════
 *  EL CIELO POR SEMANA, POR MES Y POR AÑO
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LAS DOS COSAS QUE ELLA RECLAMÓ, Y SON LA MISMA ───────────┐
 * │                                                            │
 * │ «en ese cuadro me pides que te dé toda la info que no       │
 * │  tengo y deberías tener tú o Nova, no yo»                  │
 * │                                                            │
 * │ «el pensum kármico tampoco me lo está dando y se supone    │
 * │  que lo da en automático»                                  │
 * │                                                            │
 * │ Las dos salen del mismo hueco: el pensum automático ya      │
 * │ estaba escrito y funcionando, pero leía la hoja Tránsitos, │
 * │ y esa hoja estaba VACÍA porque esperaba que los pegara     │
 * │ ella. Un motor correcto sin combustible no da ninguna      │
 * │ señal de estar roto: simplemente no produce nada.          │
 * │                                                            │
 * │ Sus tránsitos no son un dato suyo que haya que pedirle.    │
 * │ Son una cuenta, y la cuenta ya está hecha en               │
 * │ 88-efemerides.gs. Aquí se siembra sola la primera vez que  │
 * │ ella abre El cielo, y el pensum arranca detrás.            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ TRES RANGOS Y NO UNO ─────────────────────────────┐
 * │                                                            │
 * │ «necesito que lea los tránsitos de los planetas, por       │
 * │  semana, mes y año, dependiendo del planeta y su energía»  │
 * │                                                            │
 * │ Y tiene razón técnica, no solo de gusto: la Luna cambia    │
 * │ de fase cada tres días, Júpiter dura tres meses y Plutón   │
 * │ cuatro años. Mostrar los tres en la misma lista es lo que  │
 * │ hace que una pantalla de astrología no sirva para decidir  │
 * │ nada — lo de hoy queda enterrado debajo de lo de 2029.     │
 * │                                                            │
 * │   SEMANA → la Luna manda. Lo que se hace esta semana.      │
 * │   MES    → el mes solar profectado y lo social (Júpiter,   │
 * │            Saturno). Lo que se ordena este mes.            │
 * │   AÑO    → la profección anual, los ejes del karma y lo    │
 * │            generacional. De qué va el capítulo.            │
 * │                                                            │
 * │ Cada rango filtra por la VELOCIDAD del planeta, no por     │
 * │ gusto: un tránsito solo aparece donde de verdad se nota.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

// ─── SEMBRAR LO QUE NOVA YA SABE ─────────────────────────────

/**
 * Pone los tránsitos calculados en la hoja, si la hoja está vacía.
 *
 * Tres reglas, y las tres importan:
 *
 * 1· SOLO SI ESTÁ VACÍA. Si ella ya escribió o corrigió tránsitos, no
 *    se toca nada. Sembrar encima de lo que escribió alguien es la
 *    forma más rápida de que deje de confiar en la pantalla.
 *
 * 2· SOLO A QUIEN LE CORRESPONDE. La semilla es de SU carta y lleva su
 *    usuario_id dentro. A otra persona no se le siembra la carta de
 *    Manuela: se le dice que hay que calcular la suya.
 *
 * 3· EN UN SOLO SETVALUES. Ciento setenta y ocho appendRow son ciento
 *    setenta y ocho escrituras y unos veinte segundos; el mismo bloque
 *    de una vez es una.
 */
/**
 * ═══════════════════════════════════════════════════════════════
 *   ¿ESTAS EFEMÉRIDES SON DE QUIEN PREGUNTA?
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ EL ERROR QUE ESTO ARREGLA ────────────────────────────────┐
 * │                                                            │
 * │ «en el cielo sigue pidiéndome tránsitos cuando eso ya lo    │
 * │  habíamos organizado (…) ¡y sigue con los tránsitos y el    │
 * │  pensum kármico!»                                           │
 * │                                                            │
 * │ Cuando calculé sus efemérides con Swiss Ephemeris guardé    │
 * │ su carta con un correo DE RELLENO —`manuela@nova.com`— que  │
 * │ me inventé y nunca le pregunté. Y el código decía: siembra  │
 * │ solo si el correo de quien entra coincide con ese.          │
 * │                                                            │
 * │ Su usuario en NovaSoul es el correo con el que entra a Nova │
 * │ Central, que no es ese. Así que no coincidía NUNCA, y de    │
 * │ esa comparación de texto colgaban CUATRO cosas:             │
 * │                                                            │
 * │   · las temporadas de tránsito (la hoja quedaba vacía)      │
 * │   · el pensum kármico (se arma desde esas temporadas)       │
 * │   · los ejes: nodos, Mediocielo, Fondo del cielo            │
 * │   · la lectura a fondo del año                              │
 * │                                                            │
 * │ Y lo peor: Nova SABÍA por qué no sembraba —lo guardaba en   │
 * │ `transitosPorque`— y la pantalla nunca lo mostraba. Solo    │
 * │ cuadros vacíos, días seguidos.                              │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ SE COMPARA LA CARTA Y NO EL CORREO ───────────────┐
 * │                                                            │
 * │ Una carta natal no es de un correo: es de una persona. Ella │
 * │ puede cambiar de correo mañana y su Ascendente seguirá en   │
 * │ 15° de Capricornio.                                        │
 * │                                                            │
 * │ Así que la pregunta correcta no es «¿te llamas igual?» sino │
 * │ «¿es tu cielo?». Y eso sí se puede comprobar: si el         │
 * │ Ascendente que ella tiene cargado coincide con el de estas  │
 * │ efemérides, están calculadas para su nacimiento. Un grado   │
 * │ de tolerancia porque Horus y Swiss Ephemeris redondean      │
 * │ distinto, y un grado no alcanza para confundir dos cartas.  │
 * │                                                            │
 * │ El correo se sigue aceptando para no romper nada de antes.  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
function efemeridesSonDe_(uid) {
  if (typeof EFEMERIDES_CARTA === 'undefined') return false;

  /**
   * Aquí había un atajo: si el correo calzaba con
   * `EFEMERIDES_CARTA.usuario_id`, listo. Lo quité.
   *
   * Ese campo vale 'manuela@nova.com', que es un marcador de posición
   * que nunca fue el correo de nadie — o sea que el atajo no ayudaba a
   * la persona a la que pretendía ayudar. Y sí podía hacer daño: es un
   * correo escrito a mano dentro de un producto que ahora usan tres
   * cuentas, y cualquiera que llegara a tenerlo heredaría la carta
   * natal de otra persona. Es la misma clase de error que dejó a Sara
   * sin poder entrar: un dato de una cuenta escrito dentro del código.
   *
   * La prueba de verdad es el Ascendente, que es evidencia y no una
   * etiqueta.
   */

  /**
   * Y si no, ¿es su cielo? El Ascendente lo dice: depende de la hora y
   * el lugar EXACTOS del nacimiento, así que dos personas distintas no
   * lo comparten por casualidad.
   *
   * Se compara el signo y el grado DENTRO del signo, y no el grado a
   * secas, porque las dos cosas se guardan distinto según de dónde
   * vengan: la hoja Carta trae «capricornio 15,1» y estas efemérides
   * traen 285,1045 —los mismos 15,1 contados desde Aries—. El `% 30`
   * las deja hablando el mismo idioma sin tener que saber cuál es cuál.
   *
   * (Mi primera versión comparaba los grados crudos: 15,1 contra
   * 285,1045. No habría casado nunca, y habría dejado el problema
   * exactamente igual que estaba.)
   */
  const asc = (cartaDe_(uid) || []).filter(function (c) {
    return c.cuerpo === 'ascendente' && c.grado !== null;
  })[0];
  if (!asc) return false;

  const suyo = EFEMERIDES_CARTA.natal && EFEMERIDES_CARTA.natal.ascendente;
  if (!suyo) return false;
  if (norm(asc.signo) !== norm(suyo.signo)) return false;

  const enSigno = function (g) { return ((Number(g) % 30) + 30) % 30; };
  const a = enSigno(asc.grado), b = enSigno(suyo.grado);
  if (!isFinite(a) || !isFinite(b)) return false;
  // Un grado de tolerancia: Horus y Swiss Ephemeris redondean distinto,
  // y un grado no alcanza para confundir dos cartas.
  const sep = Math.min(Math.abs(a - b), 30 - Math.abs(a - b));
  return sep <= 1;
}

/**
 * ═══════════════════════════════════════════════════════════════
 *   POR QUÉ EL CIELO NO ESTÁ CARGANDO
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ POR QUÉ EXISTE ───────────────────────────────────────────┐
 * │                                                            │
 * │ «en el cielo sigue pidiéndome tránsitos cuando eso ya lo    │
 * │  habíamos organizado (…) ¡¡y sigue con los tránsitos y el   │
 * │  pensum kármico!! ayñ ya me está estresando esto».          │
 * │                                                            │
 * │ El cielo es una CADENA, y cuando se corta, la pantalla solo │
 * │ puede decir el síntoma del final: «faltan tránsitos». Pero  │
 * │ el corte puede estar tres eslabones antes, y desde fuera no │
 * │ hay forma de saber en cuál.                                 │
 * │                                                            │
 * │     Carta natal → ¿las efemérides son tuyas? → se siembran  │
 * │     los 178 tránsitos → el pensum se propone desde ellos    │
 * │                                                            │
 * │ Esto recorre la cadena y dice DÓNDE se cortó y qué hacer.   │
 * │ Sin esto, cada vuelta cuesta un mensaje mío preguntando y   │
 * │ uno suyo mirando la hoja.                                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Se corre a mano desde el editor, sin argumentos, como todo lo demás.
 * No escribe nada: solo mira y cuenta.
 */
function porQueElCielo(uid) {
  const L = [];
  let yo;
  try { yo = soulDeQuien_(uid); }
  catch (e) { return soulDecir_('No pude saber de quién: ' + e.message); }
  L.push('EL CIELO DE ' + yo);
  L.push('');

  // 1 · La carta natal
  let carta = [];
  try { carta = cartaDe_(yo) || []; } catch (e) { /* hoja sin crear */ }
  const asc = carta.filter(function (c) {
    return c.cuerpo === 'ascendente' && c.grado !== null;
  })[0];
  L.push('1· Carta natal: ' + carta.length + ' cuerpos cargados.');
  if (!carta.length) {
    L.push('   ✗ AQUÍ SE CORTA. Sin carta, Nova no puede comprobar que las');
    L.push('     efemérides que trae sean las tuyas, y no siembra nada.');
    L.push('     Carga tu carta en la hoja Carta de Nova_Soul.');
    return soulDecir_(L.join('\n'));
  }
  if (!asc) {
    L.push('   ✗ AQUÍ SE CORTA. Hay carta pero le falta el ASCENDENTE con');
    L.push('     su grado, que es justo lo que identifica tu cielo: depende');
    L.push('     de la hora y el lugar exactos, así que no se repite por');
    L.push('     casualidad. Agrega la fila «ascendente» con signo y grado.');
    return soulDecir_(L.join('\n'));
  }
  L.push('   Ascendente: ' + asc.signo + ' ' + asc.grado + '°');

  // 2 · ¿Son suyas las efemérides que trae este Nova?
  const suyas = efemeridesSonDe_(yo);
  const ref = (typeof EFEMERIDES_CARTA !== 'undefined' && EFEMERIDES_CARTA.natal)
    ? EFEMERIDES_CARTA.natal.ascendente : null;
  L.push('');
  L.push('2· Las efemérides que trae Nova son de un ascendente ' +
    (ref ? ref.signo + ' ' + (Math.round((ref.grado % 30) * 10) / 10) + '°' : '(no hay)') + '.');
  if (!suyas) {
    L.push('   ✗ AQUÍ SE CORTA. No coinciden con el tuyo, así que Nova NO');
    L.push('     te las va a poner: serían los tránsitos de otra persona.');
    L.push('     O el ascendente de tu carta está mal escrito, o estas');
    L.push('     efemérides hay que recalcularlas con tus datos.');
    return soulDecir_(L.join('\n'));
  }
  L.push('   ✓ Coinciden: son tuyas.');

  // 3 · Los tránsitos
  let transitos = [];
  try { transitos = transitosDe_(yo) || []; } catch (e) { /* hoja sin crear */ }
  L.push('');
  L.push('3· Tránsitos en tu hoja: ' + transitos.length +
    (typeof EFEMERIDES_SEMILLA !== 'undefined'
      ? ' (Nova trae ' + EFEMERIDES_SEMILLA.length + ' listos para sembrar)' : ''));
  if (!transitos.length) {
    L.push('   Todavía ninguno. Se siembran solos la próxima vez que abras');
    L.push('   el Cielo en NovaSoul. Si no pasa, dime qué dice la pantalla.');
    return soulDecir_(L.join('\n'));
  }
  const hoy = ahoraISO().slice(0, 10);
  const abiertos = transitos.filter(function (t) {
    return t.desde <= hoy && (t.hasta || t.desde) >= hoy;
  });
  L.push('   ✓ ' + abiertos.length + ' abiertos hoy.');

  // 4 · El pensum
  let pensum = [];
  try { pensum = pensumDe_(yo) || []; } catch (e) { /* hoja sin crear */ }
  L.push('');
  L.push('4· Pensum kármico: ' + pensum.length + ' temporadas.');
  if (!pensum.length) {
    L.push('   Ninguna todavía. El pensum NO se llena solo: sale de los');
    L.push('   tránsitos, pero eres tú quien decide cuáles son temporada.');
    L.push('   En el Cielo, «Proponer desde mis tránsitos» te arma');
    L.push('   candidatas y tú aceptas las que sean.');
  } else {
    const deNova = pensum.filter(function (x) { return x.laPusoNova; }).length;
    L.push('   ✓ ' + (pensum.length - deNova) + ' tuyas, ' + deNova + ' propuestas por Nova.');
  }
  L.push('');
  L.push('La cadena está entera.');
  return soulDecir_(L.join('\n'));
}

function transitosSembrar_(uid) {
  if (typeof EFEMERIDES_SEMILLA === 'undefined') {
    return { sembro: 0, porque: 'No hay efemérides calculadas en este Nova.' };
  }
  if (!efemeridesSonDe_(uid)) {
    return { sembro: 0, porque: 'Las efemérides cargadas son de otra carta. ' +
             'Para sembrar las tuyas hay que correr el script con tus datos de nacimiento.' };
  }

  let sh;
  try { sh = soulSheet_('Transitos'); }
  catch (e) { return { sembro: 0, porque: 'Todavía no existe la hoja Tránsitos. ' +
                                          'Corre bootstrapTodo() una vez.' }; }

  const mias = soulLeerSuave_('Transitos', uid, []).length;
  if (mias > 0) return { sembro: 0, yaHabia: mias };

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { sembro: 0, porque: 'Hay otro cambio guardándose.' };
  try {
    // Otra vez dentro del candado: dos pestañas abiertas a la vez
    // sembrarían dos veces, y la comprobación de arriba no sirve para eso.
    if (soulLeerSuave_('Transitos', uid, []).length > 0) return { sembro: 0 };

    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const filas = EFEMERIDES_SEMILLA.map(function (linea) {
      const x = linea.split('|');
      const cuerpoId = x[0], aspecto = x[1], natalId = x[2];
      const desde = x[3], hasta = x[4], pico = x[5];
      const casa = Number(x[6]), casaP = Number(x[7]), retro = x[8] === '1';

      const nat = LEC_NATAL[natalId];
      const cue = LEC_CUERPOS[cuerpoId];
      const nombreNatal = nat ? nat.nombre : natalId;
      const tema = (cue ? cue.nombre : cuerpoId) + ' ' + aspecto + ' a mi ' + nombreNatal;

      const v = {
        usuario_id: uid, fecha: pico, desde: desde, hasta: hasta,
        cuerpo: cuerpoId, aspecto: aspecto, a_natal: nombreNatal,
        casa: casa || '', casa_placidus: casaP || '',
        tema: tema,
        intensidad_pct: lecIntensidad_(cuerpoId, norm(aspecto), natalId),
        // Los textos se dejan EN BLANCO a propósito: la lectura la
        // compone Nova al leer, no se congela en la hoja. Así, cuando
        // el vocabulario mejore, mejoran también las filas viejas.
        texto_transito: '', por_que: '', como_trabajarlo: '', el_otro_lado: '',
        fuente: 'Swiss Ephemeris' + (retro ? ' · retrógrado' : ''),
      };
      return enc.map(function (c) { return v[c] !== undefined ? v[c] : ''; });
    });

    if (!filas.length) return { sembro: 0 };
    sh.getRange(sh.getLastRow() + 1, 1, filas.length, enc.length).setValues(filas);
    // Se acaba de escribir en bloque: lo que hubiera en memoria ya no
    // sirve, y justo después el pensum va a leer esta misma hoja.
    soulOlvidar_('Transitos');
    return { sembro: filas.length, hasta: EFEMERIDES_HASTA };
  } catch (e) {
    return { sembro: 0, porque: e.message };
  } finally { lock.releaseLock(); }
}

/**
 * Cuándo hay que volver a correr el script de efemérides.
 *
 * Se avisa con seis meses, no el día que se acaba. Un horizonte que se
 * agota en silencio se ve igual que un pensum que dejó de funcionar, y
 * ese error ya pasó una vez.
 */
function efemeridesVencen_(hoyISO) {
  if (typeof EFEMERIDES_HASTA === 'undefined') return null;
  const dias = Math.round((new Date(EFEMERIDES_HASTA + 'T00:00:00Z') -
                           new Date(hoyISO + 'T00:00:00Z')) / 86400000);
  return {
    hasta: EFEMERIDES_HASTA, diasRestantes: dias,
    avisar: dias < 180,
    porque: dias < 180
      ? 'Las efemérides llegan hasta ' + EFEMERIDES_HASTA + '. Cuando falten menos de ' +
        'seis meses hay que volver a correr efemerides/generar-transitos.py para ' +
        'extender el horizonte.'
      : '',
  };
}

// ─── LAS LUNAS DE UN TRAMO ───────────────────────────────────

/**
 * Las lunas nuevas y llenas que caen dentro de un rango.
 *
 * Se buscan recorriendo días y mirando cuándo cambia la fase, en vez de
 * calcular el instante exacto. La diferencia con el instante real son
 * horas; para «qué día entrego» eso no cambia nada, y para «a qué hora
 * exacta» ella ya tiene Horus.
 */
function lunasEntre_(desdeISO, hastaISO) {
  const out = [];
  let f = desdeISO, anterior = '';
  let guarda = 0;
  while (f <= hastaISO && guarda++ < 800) {
    const fase = faseLunar_(f);
    if (fase.id !== anterior && (fase.id === 'nueva' || fase.id === 'llena')) {
      out.push({ fecha: f, id: fase.id, nombre: fase.nombre, forma: fase.forma,
                 que: fase.que, momento: fase.momento });
    }
    anterior = fase.id;
    f = masDias_(f, 1);
  }
  return out;
}

// ─── LOS TRES RANGOS ─────────────────────────────────────────

/**
 * Qué planetas tienen sentido mirar en cada ventana.
 *
 * No es una preferencia: es la velocidad. En una semana Plutón no se
 * movió, así que ponerlo en la lista de la semana es ruido; y Mercurio
 * ya pasó tres veces en un año, así que ponerlo en la del año también.
 */
/**
 * ═══════════════════════════════════════════════════════════════
 *   LOS TRES RANGOS, QUE AHORA SÍ SON DE TIEMPO
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE ESTABA MAL ────────────────────────────────────────┐
 * │                                                            │
 * │ «me lo tiraste todo de un solo golpe, mucha información me  │
 * │  sobre-estimula (…) me aparece todo lo de la semana, el mes │
 * │  y el año, y no veo cuáles son los de la semana, cuáles los │
 * │  del mes y cuáles los del año, no hay distinción».          │
 * │                                                            │
 * │ Tenía toda la razón, y la causa era de diseño: las tres     │
 * │ pestañas NO FILTRABAN POR TIEMPO. Filtraban por planeta —   │
 * │ «semana» era personales+sociales, «mes» esos MISMOS más     │
 * │ generacionales—. Por eso se repetían entre pestañas y el    │
 * │ nombre prometía tiempo mientras el filtro hablaba de        │
 * │ planetas.                                                   │
 * │                                                            │
 * │ Se ve en sus números: las 8 temporadas que le salieron en   │
 * │ «Esta semana» duraban entre 25 y 175 días. NINGUNA era de   │
 * │ esa semana.                                                 │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ CÓMO QUEDÓ ───────────────────────────────────────────────┐
 * │                                                            │
 * │ Cada temporada cae en UN SOLO rango, por lo que DURA. No    │
 * │ hay repetición posible, y la distinción es la que el        │
 * │ nombre promete.                                             │
 * │                                                            │
 * │   Esta semana  ·  hasta 21 días  ·  lo que se decide ya     │
 * │   Este mes     ·  22 a 120 días  ·  lo que se está armando  │
 * │   Este año     ·  más de 120     ·  el fondo que no se mueve│
 * │                                                            │
 * │ Los cortes no son astrología: son la escala de tiempo con   │
 * │ la que una persona decide. Tres semanas es «esta semana o   │
 * │ la que viene»; cuatro meses ya es «este año».               │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const CIELO_RANGOS = {
  semana: { nombre: 'Esta semana', dias: 7, duraHasta: 21,
            que: 'Lo que se decide en días. Aquí manda la Luna.' },
  mes:    { nombre: 'Este mes', dias: 30, duraDesde: 22, duraHasta: 120,
            que: 'Lo que se está armando: semanas, no días. Júpiter y Saturno ordenando.' },
  anio:   { nombre: 'Este año', dias: 365, duraDesde: 121,
            que: 'El fondo que no se mueve. Energía larga y constante: es la que se puede aprovechar a propósito.' },
};

/** Cuántos días dura una temporada, contando el primero y el último. */
function duracionTransito_(t) {
  const fin = t.hasta || t.desde;
  if (!t.desde || !fin) return 1;
  return Math.round((new Date(fin + 'T00:00:00Z') -
                     new Date(t.desde + 'T00:00:00Z')) / 86400000) + 1;
}

/** ¿A cuál de los tres rangos pertenece, por lo que dura? */
function rangoDeTransito_(t) {
  const d = duracionTransito_(t);
  if (d <= CIELO_RANGOS.semana.duraHasta) return 'semana';
  if (d <= CIELO_RANGOS.mes.duraHasta) return 'mes';
  return 'anio';
}

/**
 * La lectura de un rango, con las cuatro preguntas respondidas.
 *
 * Devuelve las temporadas ORDENADAS POR INTENSIDAD, no por fecha. En
 * una semana cualquiera hay seis cosas abiertas y cinco no importan;
 * ordenar por fecha pone arriba la que empezó primero, que no tiene
 * ninguna razón para ser la que más pesa.
 */
function cieloRango_(uid, rango, hoyISO, carta, nac) {
  const R = CIELO_RANGOS[rango] || CIELO_RANGOS.semana;
  const desde = rango === 'semana' ? lunesDe_(hoyISO) : hoyISO;
  const hasta = masDias_(desde, R.dias - 1);

  const transitos = transitosDe_(uid);
  const dentro = transitos.filter(function (t) {
    // Cruza el tramo: empieza antes y termina después, o cae adentro.
    if (t.desde > hasta) return false;
    if (t.hasta && t.hasta < desde) return false;
    // Y es de ESTE rango por lo que dura. Una temporada cae en uno solo:
    // es lo que hace que las tres pestañas se distingan.
    return rangoDeTransito_(t) === rango;
  });

  const leidos = dentro.map(function (t) {
    const l = lecturaTransito_(t);
    if (!l.hay) return null;
    const dias = duracionTransito_(t);
    return Object.assign({}, l, {
      desde: t.desde, hasta: t.hasta, dias: dias,
      activaHoy: t.desde <= hoyISO && (!t.hasta || t.hasta >= hoyISO),
      pico: t.desde <= hoyISO && (!t.hasta || t.hasta >= hoyISO),
      casa: t.casa, casaPlacidus: t.casaPlacidus, casasDifieren: t.casasDifieren,
      // Lo que ella haya escrito a mano manda sobre lo compuesto.
      suyo: t.texto || t.porQue || t.como || t.elOtroLado ? {
        texto: t.texto, porQue: t.porQue, como: t.como, elOtroLado: t.elOtroLado,
      } : null,
    });
  }).filter(function (x) { return x; })
    .sort(function (a, b) { return b.intensidad - a.intensidad; });

  const fuera = {
    semana: 'Lo que dura meses no cabe en una semana. Está en Este mes y en Este año.',
    mes: '',
    anio: 'Lo que se decide en días no es del año. Está en Esta semana.',
  }[rango];

  /**
   * ── ACTIVAS HOY, PRIMERO ──
   *
   * «para la parte de temporadas abiertas, que estén como principal,
   *  activas hoy. Eso me interesa leerlo».
   *
   * Lo que ya está abierto es sobre lo que se puede hacer algo hoy. Lo
   * que empieza en dos semanas es una nota al pie, no una lectura.
   */
  const activas = leidos.filter(function (x) { return x.activaHoy; });
  const porVenir = leidos.filter(function (x) { return !x.activaHoy; });

  /**
   * ── Y EL FONDO LARGO, QUE NO PUEDE DESAPARECER ──
   *
   * «me importa lo de Saturno porque es una energía constante y
   *  duradera y debo aprovecharla».
   *
   * Con el filtro por duración, Saturno de 175 días se va a «Este año»
   * — que es donde va, porque no es noticia de esta semana. Pero
   * tampoco puede desaparecer de la vista: es la energía de fondo sobre
   * la que se construye todo lo demás.
   *
   * Así que en los rangos cortos aparece nombrada, sin su lectura
   * completa y sin competir. Una línea, no un bloque.
   */
  let deFondo = [];
  if (rango !== 'anio') {
    deFondo = transitos.filter(function (t) {
      return rangoDeTransito_(t) === 'anio' &&
             t.desde <= hoyISO && (!t.hasta || t.hasta >= hoyISO);
    }).map(function (t) {
      const l = lecturaTransito_(t);
      return l.hay ? { titulo: l.titulo, hasta: t.hasta,
                       intensidad: l.intensidad, dias: duracionTransito_(t) } : null;
    }).filter(function (x) { return x; })
      .sort(function (a, b) { return b.intensidad - a.intensidad; })
      .slice(0, 3);
  }

  return {
    rango: rango, nombre: R.nombre, que: R.que,
    desde: desde, hasta: hasta,
    temporadas: leidos,
    // Lo que se lee primero, y lo que puede esperar
    activas: activas,
    porVenir: porVenir,
    deFondo: deFondo,
    lunas: rango === 'anio' ? [] : lunasEntre_(desde, hasta),
    lunaHoy: faseLunar_(hoyISO),
    fueraDeRango: fuera,
    // Un tramo sin nada no es un error. Se dice.
    vacio: leidos.length === 0,
    porqueVacio: leidos.length ? '' : (
      rango === 'semana'
        ? 'Ninguna temporada corta abierta esta semana: nada que se decida en días. ' +
          'Un cielo tranquilo también es información — es tiempo para sostener lo que ya está' +
          (deFondo.length ? ', y lo de fondo sigue corriendo.' : '.')
        : 'Nada de esta escala abierto ahora mismo. Mira los otros dos tramos.'),
  };
}

// ─── LA LLAMADA ──────────────────────────────────────────────

/**
 * `nc_soul_cielo_lectura` · la profundidad, en su propia llamada.
 *
 * Va aparte de `nc_soul_cielo` a propósito: la foto del cielo se pinta
 * al entrar, y la lectura de fondo solo cuando ella toca una pestaña.
 * Componer ciento setenta y ocho lecturas para pintar una que se ve
 * sería pagar el precio entero por la parte que se mira.
 */
function soulCieloLectura(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);
  const rango = CIELO_RANGOS[String(p.rango || '')] ? String(p.rango) : 'semana';

  // Si nunca se sembró, se siembra aquí también: ella puede llegar por
  // esta pantalla antes que por la otra.
  const semilla = transitosSembrar_(uid);

  const carta = cartaDe_(uid);
  const nac = nacimientoDe_(uid);
  const rev = revolucionVentana_(nac.fecha, hoy);

  const out = {
    ok: true, hoy: hoy,
    rangos: CIELO_RANGOS,
    lectura: cieloRango_(uid, rango, hoy, carta, nac),
    sembro: semilla.sembro || 0,
    efemerides: efemeridesVencen_(hoy),
  };

  // El año trae además la profección leída a fondo y los ejes: es lo
  // que ella pidió con el ejemplo de «casa 8 de Leo por mi año».
  if (rango === 'anio') {
    const asc = (carta || []).filter(function (c) { return c.cuerpo === 'ascendente'; })[0];
    if (asc && asc.signo && rev) {
      out.profeccion = lecProfeccion_((rev.edad % 12) + 1, asc.signo,
        efemeridesSonDe_(uid) ? EFEMERIDES_CARTA : null);
      out.profeccion.desde = rev.desde;
      out.profeccion.hasta = rev.hasta;
      out.profeccion.edad = rev.edad;
    } else {
      out.profeccionPorque = !asc || !asc.signo
        ? 'Para leer tu año necesito tu Ascendente, y todavía no está en tu carta.'
        : 'Necesito tu fecha de nacimiento para contar los años cumplidos.';
    }
    out.ejes = efemeridesSonDe_(uid) ? lecEjes_(EFEMERIDES_CARTA) : [];
  }

  return out;
}

/**
 * La lectura de UNA temporada, a fondo. Para cuando toca una de la lista.
 *
 * Se recompone en vez de guardarse: el vocabulario de 87-lectura.gs va
 * a seguir creciendo, y una lectura congelada en la hoja el día que se
 * creó envejecería sin que nadie se entere.
 */
function soulTransitoLectura(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const clave = String(p.clave || '').trim();
  if (!clave) return { ok: false, error: 'No sé qué temporada leer.' };

  const t = transitosDe_(uid).filter(function (x) {
    return x.cuerpo + '|' + norm(x.aspecto) + '|' + norm(x.aNatal) + '|' + x.desde === clave;
  })[0];
  if (!t) return { ok: false, error: 'No encuentro esa temporada.' };

  const l = lecturaTransito_(t);
  if (!l.hay) return { ok: false, error: l.porQue };

  // Las dos casas, cuando discrepan, con las dos lecturas enteras.
  let otraCasa = null;
  if (t.casasDifieren) {
    const otra = lecturaTransito_(Object.assign({}, t, { casa: t.casaPlacidus }));
    otraCasa = {
      casa: t.casaPlacidus, sistema: 'Placidus (el de Horus)',
      tiempoPara: otra.tiempoPara, aFavor: otra.aFavor, desequilibrio: otra.desequilibrio,
    };
  }

  return { ok: true, lectura: l, desde: t.desde, hasta: t.hasta,
           casaEntera: t.casa, casaPlacidus: t.casaPlacidus, otraCasa: otraCasa,
           suyo: { texto: t.texto, porQue: t.porQue, como: t.como, elOtroLado: t.elOtroLado } };
}
