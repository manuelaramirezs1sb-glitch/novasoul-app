/**
 * ═══════════════════════════════════════════════════════════════
 *  LA LECTURA DEL CIELO · el vocabulario con el que Nova interpreta
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ QUÉ PIDIÓ ELLA, CON SUS PALABRAS ─────────────────────────┐
 * │                                                            │
 * │ «quiero más análisis, más profundo, por ejemplo en casa 8  │
 * │  de Leo por mi año, necesito profundidad y análisis,       │
 * │  tiempo para qué, cómo me afecta la energía que me toca,   │
 * │  cómo la puedo trabajar a mi favor, y cómo se puede ver    │
 * │  si no se integra correctamente la energía, como un        │
 * │  desequilibrio».                                           │
 * │                                                            │
 * │ Esas son CUATRO preguntas fijas, y esta pantalla las       │
 * │ responde siempre las cuatro. No es una frase de galleta    │
 * │ de la fortuna: es una lectura compuesta de piezas que se   │
 * │ pueden mirar por separado.                                 │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ CÓMO SE COMPONE UNA LECTURA ──────────────────────────────┐
 * │                                                            │
 * │ Un tránsito son cuatro datos: QUÉ planeta, CÓMO toca       │
 * │ (el aspecto), A QUÉ de tu carta, y POR DÓNDE (la casa).    │
 * │                                                            │
 * │   Júpiter  ·  sextil  ·  a tu Mediocielo  ·  por casa 8    │
 * │   ───────     ──────     ───────────────     ─────────     │
 * │   el verbo    el tono     a quién le pasa    el escenario  │
 * │                                                            │
 * │ Cada pieza tiene aquí su propio vocabulario, y la lectura  │
 * │ se arma cruzándolas. Por eso una temporada que nunca       │
 * │ estuvo escrita en ningún lado sale leída igual de fondo    │
 * │ que una famosa: no hay lista de textos, hay una gramática. │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE ESTO NO ES ────────────────────────────────────────┐
 * │                                                            │
 * │ Esto es la lectura COMÚN: lo que la tradición dice de esa  │
 * │ combinación. Nova no sabe lo que a ella le pasa por        │
 * │ dentro, y no va a fingir que sí.                           │
 * │                                                            │
 * │ Por eso todo lo que sale de aquí va marcado como lectura   │
 * │ de partida, y CUALQUIER cosa que ella escriba en su        │
 * │ pensum manda sobre esto. Es la misma regla que ya estaba   │
 * │ en el resto de El cielo, sostenida.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

// ─── LOS PLANETAS QUE TRANSITAN ──────────────────────────────

/**
 * Qué HACE cada planeta cuando pasa por encima de algo tuyo.
 *
 * `regala` y `tuerce` son la misma energía: la que se integra y la que
 * no. Ella pidió las dos caras, y la segunda no es un castigo — es la
 * primera sin digerir.
 */
const LEC_CUERPOS = {
  sol: {
    nombre: 'Sol', grupo: 'personal', ritmo: 'un mes por signo', dura: 'unos días',
    hace: 'alumbra y pone en el centro',
    regala: 'claridad sobre lo que de verdad es tuyo, y ganas de mostrarlo',
    tuerce: 'volverlo todo sobre ti, y confundir ser visto con ser querido',
  },
  mercurio: {
    nombre: 'Mercurio', grupo: 'personal', ritmo: 'unas semanas por signo', dura: 'unos días',
    hace: 'nombra, pregunta y pone en palabras',
    regala: 'la conversación que estaba pendiente y el papel que había que firmar',
    tuerce: 'hablar de más, decidir sobre datos a medias, o quedarte analizando en vez de mover',
  },
  venus: {
    nombre: 'Venus', grupo: 'personal', ritmo: 'un mes por signo', dura: 'una o dos semanas',
    hace: 'acerca, endulza y pone precio',
    regala: 'que te lean bien, que te paguen bien, y que lo bonito valga por sí solo',
    tuerce: 'comprar la paz con complacencia, y gastar para tapar un vacío',
  },
  marte: {
    nombre: 'Marte', grupo: 'personal', ritmo: 'unos dos meses por signo', dura: 'de una a tres semanas',
    hace: 'empuja, corta y defiende',
    regala: 'la fuerza para arrancar lo que llevaba meses en la lista',
    tuerce: 'pelear con quien no era, quemarte en una semana, o apurar algo que no estaba listo',
  },
  jupiter: {
    nombre: 'Júpiter', grupo: 'social', ritmo: 'un año por signo', dura: 'de dos a cuatro meses',
    hace: 'abre, agranda y da permiso',
    regala: 'la puerta que se abre sola y el sí que llega sin pelearlo',
    tuerce: 'agrandar también el gasto, el riesgo y la promesa — decir que sí a todo y no caber',
  },
  saturno: {
    nombre: 'Saturno', grupo: 'social', ritmo: 'dos años y medio por signo', dura: 'de tres a nueve meses',
    hace: 'aprieta, ordena y cobra',
    regala: 'una estructura que aguanta, y respeto por lo que sí sostuviste',
    tuerce: 'miedo, frío, y creer que lo que cuesta trabajo es que no sirves',
  },
  urano: {
    nombre: 'Urano', grupo: 'generacional', ritmo: 'siete años por signo', dura: 'de seis meses a dos años',
    hace: 'rompe el molde y suelta de golpe',
    regala: 'la libertad que no te atrevías a pedir, y una idea que cambia el tablero',
    tuerce: 'romper por romper, e irte de algo que solo había que cambiar de forma',
  },
  neptuno: {
    nombre: 'Neptuno', grupo: 'generacional', ritmo: 'catorce años por signo', dura: 'de uno a tres años',
    hace: 'disuelve los bordes',
    regala: 'compasión, intuición fina y arte — y soltar lo que ya no se puede sostener',
    tuerce: 'niebla: idealizar a alguien, no ver la cuenta, o cansarte sin causa clara',
  },
  pluton: {
    nombre: 'Plutón', grupo: 'generacional', ritmo: 'de doce a treinta años por signo', dura: 'de uno a cuatro años',
    hace: 'saca a la superficie y transforma de raíz',
    regala: 'poder real sobre algo donde antes solo aguantabas',
    tuerce: 'control, obsesión y pulsos de poder donde bastaba con hablar',
  },
};

// ─── CÓMO TOCA: EL ASPECTO ───────────────────────────────────

/**
 * El aspecto es el TONO. No dice qué pasa, dice cómo llega.
 *
 * `tension` de 0 a 3 sirve para dos cosas: ordenar la pantalla y
 * calcular la intensidad. Un trígono no es «bueno» y una cuadratura no
 * es «mala»: lo fácil se desaprovecha y lo difícil es lo que construye.
 * Eso está escrito en los textos a propósito.
 */
const LEC_ASPECTOS = {
  conjuncion: {
    nombre: 'Conjunción', simbolo: '☌', grados: 0, tension: 2,
    que: 'se funden: el planeta se sienta encima y por un rato no se distingue qué es tuyo y qué es suyo',
    pide: 'arrancar algo con esa energía, sabiendo que la vas a sentir en carne propia',
    riesgo: 'no ver de afuera lo que estás viviendo por dentro',
  },
  sextil: {
    nombre: 'Sextil', simbolo: '⚹', grados: 60, tension: 0,
    que: 'se hacen guiños: hay una puerta abierta, pero no se abre sola',
    pide: 'hacer el movimiento — esta es la ayuda que solo sirve si la usas',
    riesgo: 'que pase entero y ni te enteres, porque nada dolió lo suficiente para mirarlo',
  },
  cuadratura: {
    nombre: 'Cuadratura', simbolo: '□', grados: 90, tension: 3,
    que: 'chocan: dos cosas tuyas quieren lo contrario al mismo tiempo',
    pide: 'elegir, y pagar el precio de elegir — no hay salida donde ganen las dos',
    riesgo: 'aguantar la fricción en vez de decidir, y que se vuelva síntoma',
  },
  trigono: {
    nombre: 'Trígono', simbolo: '△', grados: 120, tension: 0,
    que: 'fluyen: la energía corre sin resistencia, como algo que siempre supiste hacer',
    pide: 'apoyarte ahí para mover lo pesado de otra parte de la carta',
    riesgo: 'acomodarte — lo fácil rara vez enseña algo nuevo',
  },
  oposicion: {
    nombre: 'Oposición', simbolo: '☍', grados: 180, tension: 3,
    que: 'se miran de frente: lo que te toca llega por fuera, casi siempre con cara de otra persona',
    pide: 'ver que lo de enfrente también es tuyo, y negociar en vez de ganar',
    riesgo: 'echarle la culpa al de enfrente y perder el aprendizaje entero',
  },
};

/** Los nombres vienen con tilde desde las efemérides. */
function lecAspecto_(a) {
  const k = norm(a).replace(/ó/g, 'o').replace(/í/g, 'i');
  return LEC_ASPECTOS[k] || LEC_ASPECTOS[k.replace('conjucion', 'conjuncion')] || null;
}

// ─── A QUÉ TE TOCA: EL PUNTO NATAL ───────────────────────────

/**
 * Qué parte de ella es cada punto de la carta.
 *
 * Los nodos, el Mediocielo y el Fondo del cielo están aquí porque ella
 * los pidió por nombre: «ten en cuenta mis nodos del karma, mi medio
 * cielo y mi fondo cielo».
 */
const LEC_NATAL = {
  sol: { nombre: 'Sol', quien: 'quién eres cuando nadie te está pidiendo nada',
         aFavor: 'decidir desde lo que sí eres, no desde lo que se espera',
         mal: 'te cuesta reconocerte en lo que estás haciendo' },
  luna: { nombre: 'Luna', quien: 'lo que sientes antes de pensarlo, y lo que necesitas para estar en paz',
          aFavor: 'cuidar el cuerpo, el descanso y la casa como si fuera trabajo — porque lo es',
          mal: 'el cuerpo habla primero: sueño, digestión, ganas de llorar sin motivo' },
  mercurio: { nombre: 'Mercurio', quien: 'cómo piensas, hablas, escribes y cierras acuerdos',
              aFavor: 'poner por escrito lo que estaba de palabra',
              mal: 'malentendidos, mensajes que no llegan, decisiones tomadas a medias' },
  venus: { nombre: 'Venus', quien: 'lo que te gusta, lo que vales y cómo te vinculas',
           aFavor: 'revisar precios, acuerdos y a quién le estás dando tu tiempo gratis',
           mal: 'aceptar menos de lo que vale tu trabajo, o pagar por ser querida' },
  marte: { nombre: 'Marte', quien: 'tu empuje, tu rabia y tu manera de defenderte',
           aFavor: 'usar la rabia como información: te está diciendo dónde te pasaron por encima',
           mal: 'explotas con quien no era, o te quedas sin fuerza para lo que sí importaba' },
  jupiter: { nombre: 'Júpiter', quien: 'tu fe, tu apetito de más y por dónde crece tu vida',
             aFavor: 'apostar por lo que ya demostró que funciona, no por lo que suena bonito',
             mal: 'exceso: más gasto, más promesas y más cosas abiertas de las que caben' },
  saturno: { nombre: 'Saturno', quien: 'tu disciplina, tus miedos viejos y lo que sostienes aunque nadie mire',
             aFavor: 'terminar UNA cosa hasta el final, aunque sea pequeña',
             mal: 'parálisis, autoexigencia sin descanso, y sensación de ir atrasada en la vida' },
  urano: { nombre: 'Urano', quien: 'lo que en ti no se deja domesticar',
           aFavor: 'darte el permiso tú, antes de que te lo quiten a la fuerza',
           mal: 'cambios bruscos de los que después te arrepientes' },
  neptuno: { nombre: 'Neptuno', quien: 'tu imaginación, tu fe y por dónde te engañas',
             aFavor: 'crear, descansar y dejar de pelear con lo que ya se acabó',
             mal: 'confusión, cansancio raro y ver en alguien lo que querías ver' },
  pluton: { nombre: 'Plutón', quien: 'lo que te transforma y lo que no sueltas',
            aFavor: 'mirar de frente eso que llevas tiempo rodeando',
            mal: 'control, celos, y quedarte en algo solo por no perder lo invertido' },
  ascendente: { nombre: 'Ascendente', quien: 'tu cara, tu cuerpo y cómo entras a los lugares',
                aFavor: 'cambiar algo visible: la forma de presentarte, el cuerpo, el nombre de lo que haces',
                mal: 'no te reconoces en el espejo ni en cómo te describen los demás' },
  descendente: { nombre: 'Descendente', quien: 'con quién te asocias y qué buscas en el otro',
                 aFavor: 'revisar sociedades y contratos: quién pone qué',
                 mal: 'atraes justo lo que dijiste que no querías' },
  medio_cielo: { nombre: 'Mediocielo', quien: 'tu lugar en el mundo: la carrera, el nombre público, lo que se ve de ti',
                 aFavor: 'mover fichas donde te ven — mostrar trabajo, cobrar, tomar el puesto',
                 mal: 'sientes que trabajas mucho y nadie sabe qué haces' },
  fondo_cielo: { nombre: 'Fondo del cielo', quien: 'tu raíz: la casa, la familia, de dónde vienes y dónde descansas',
                 aFavor: 'ordenar la base — mudanza, familia, el cuarto donde duermes',
                 mal: 'la casa pesa, o cargas algo familiar que no es tuyo' },
  nodo_norte: { nombre: 'Nodo Norte', quien: 'hacia dónde va tu karma: lo que viniste a aprender y todavía te incomoda',
                aFavor: 'hacer justo lo que te da pereza porque no lo dominas todavía',
                mal: 'te quedas en lo que ya sabes hacer y la vida se te repite igual' },
  nodo_sur: { nombre: 'Nodo Sur', quien: 'lo que ya traías sabido: tu don fácil, y también tu escondite',
              aFavor: 'usar ese don al servicio de lo otro, sin vivir ahí',
              mal: 'vuelves al refugio de siempre justo cuando tocaba crecer' },
};

function lecNatal_(nombre) {
  const t = norm(nombre);
  const alias = {
    'sol': 'sol', 'luna': 'luna', 'mercurio': 'mercurio', 'venus': 'venus',
    'marte': 'marte', 'jupiter': 'jupiter', 'saturno': 'saturno', 'urano': 'urano',
    'neptuno': 'neptuno', 'pluton': 'pluton', 'ascendente': 'ascendente',
    'descendente': 'descendente', 'mediocielo': 'medio_cielo', 'medio cielo': 'medio_cielo',
    'medio_cielo': 'medio_cielo', 'fondo del cielo': 'fondo_cielo',
    'fondo cielo': 'fondo_cielo', 'fondo_cielo': 'fondo_cielo',
    'nodo norte': 'nodo_norte', 'nodo_norte': 'nodo_norte',
    'nodo sur': 'nodo_sur', 'nodo_sur': 'nodo_sur',
  };
  return LEC_NATAL[alias[t] || t] || null;
}

// ─── POR DÓNDE PASA: LA CASA ─────────────────────────────────

/**
 * Las doce casas: el ESCENARIO donde se nota.
 *
 * `tiempoPara` es literal de lo que ella pidió: «tiempo para qué».
 */
const LEC_CASAS = {
  1:  { nombre: 'Casa 1', area: 'tu cuerpo, tu cara y cómo arrancas',
        tiempoPara: 'volver a ti: empezar de cero, cambiar de forma, ocupar tu propio espacio',
        aFavor: 'cambia algo visible y deja que el resto se acomode a eso',
        mal: 'te disuelves en lo que los demás necesitan y te pierdes de vista' },
  2:  { nombre: 'Casa 2', area: 'tu plata propia, tus cosas y lo que crees que vales',
        tiempoPara: 'ordenar ingresos, subir precios y quedarte con lo que sí usas',
        aFavor: 'mira los números reales una vez por semana, aunque no te gusten',
        mal: 'gastas para sentirte segura, o cobras poco para que no te dejen' },
  3:  { nombre: 'Casa 3', area: 'lo que dices y escribes, los hermanos, lo cercano',
        tiempoPara: 'estudiar, escribir, negociar y arreglar lo que quedó dicho a medias',
        aFavor: 'pon por escrito lo importante y manda el mensaje que estás evitando',
        mal: 'ruido: mucha conversación, mucha información, ninguna decisión' },
  4:  { nombre: 'Casa 4', area: 'tu casa, tu familia y tu raíz',
        tiempoPara: 'arreglar la base: dónde vives, con quién, y qué cargas de tu familia',
        aFavor: 'cuida el lugar donde duermes como si fuera la oficina — decide desde ahí',
        mal: 'llevas el trabajo a la cama y la casa deja de ser refugio' },
  5:  { nombre: 'Casa 5', area: 'lo que creas, lo que disfrutas, los hijos y el riesgo',
        tiempoPara: 'crear algo tuyo y mostrarlo, sin que tenga que ser rentable todavía',
        aFavor: 'guárdate un rato de la semana para algo que hagas solo porque te gusta',
        mal: 'lo conviertes todo en producto y se te apaga el gusto' },
  6:  { nombre: 'Casa 6', area: 'tu rutina, tu salud y el trabajo del día a día',
        tiempoPara: 'ajustar el sistema: horarios, cuerpo, y lo que haces todos los días',
        aFavor: 'cambia una costumbre pequeña y sostenla tres semanas antes de juzgarla',
        mal: 'el cuerpo pasa la cuenta de lo que la agenda no quiso mirar' },
  7:  { nombre: 'Casa 7', area: 'tus sociedades, tu pareja y los contratos',
        tiempoPara: 'revisar con quién estás asociada y bajo qué condiciones',
        aFavor: 'di en voz alta el acuerdo que estabas dando por supuesto',
        mal: 'te quedas en un trato desigual por no abrir la conversación' },
  8:  { nombre: 'Casa 8', area: 'la plata de otros, lo que se comparte, lo que muere y lo que te transforma',
        tiempoPara: 'soltar lo que ya se acabó y mirar de frente deudas, sociedades y acuerdos de dinero ajeno',
        aFavor: 'nombra la cifra y la fecha — aquí lo que no se dice es lo que cobra intereses',
        mal: 'controlas para no sentir miedo, y lo que no soltaste se te vuelve deuda' },
  9:  { nombre: 'Casa 9', area: 'lo lejano: estudios largos, viajes, otro país, tu manera de creer',
        tiempoPara: 'aprender en serio, publicar, irte lejos o cambiar de marco',
        aFavor: 'estudia algo que no te sirva de inmediato: aquí eso sí paga',
        mal: 'predicas lo que todavía no practicas, o huyes llamándolo búsqueda' },
  10: { nombre: 'Casa 10', area: 'tu carrera, tu nombre público y lo que se ve de ti',
        tiempoPara: 'tomar el puesto, cobrar lo que vales y dejar que se vea tu trabajo',
        aFavor: 'muestra lo que ya hiciste antes de esperar a que esté perfecto',
        mal: 'trabajas todo el día y nadie sabe decir a qué te dedicas' },
  11: { nombre: 'Casa 11', area: 'tu gente, tus redes y lo que quieres a futuro',
        tiempoPara: 'construir comunidad y pedir ayuda a quien ya te ve',
        aFavor: 'escribe a tres personas que podrían abrirte algo, sin pedirles nada todavía',
        mal: 'estás rodeada y sola, o te pierdes en el grupo para no decidir' },
  12: { nombre: 'Casa 12', area: 'lo que no se ve: el descanso, lo inconsciente, lo que se cierra en silencio',
        tiempoPara: 'descansar, cerrar puertas viejas y preparar en privado lo que aún no se muestra',
        aFavor: 'no arranques nada público — usa este tiempo para sanar y para escribir en borrador',
        mal: 'te agotas sin causa visible y saboteas justo antes de llegar' },
};

// ─── LOS SIGNOS: EL ESTILO ───────────────────────────────────

const LEC_SIGNOS = {
  aries:       { nombre: 'Aries',       elemento: 'fuego',  modo: 'cardinal', estilo: 'arranca de golpe y pregunta después' },
  tauro:       { nombre: 'Tauro',       elemento: 'tierra', modo: 'fijo',     estilo: 'va lento y no se mueve de donde se planta' },
  geminis:     { nombre: 'Géminis',     elemento: 'aire',   modo: 'mutable',  estilo: 'prueba de todo y se aburre rápido' },
  cancer:      { nombre: 'Cáncer',      elemento: 'agua',   modo: 'cardinal', estilo: 'se mueve por lo que siente y por los suyos' },
  leo:         { nombre: 'Leo',         elemento: 'fuego',  modo: 'fijo',     estilo: 'necesita que se vea, y da el pecho por lo suyo' },
  virgo:       { nombre: 'Virgo',       elemento: 'tierra', modo: 'mutable',  estilo: 'afina, corrige y no entrega hasta que sirva' },
  libra:       { nombre: 'Libra',       elemento: 'aire',   modo: 'cardinal', estilo: 'mide, compara y busca el acuerdo' },
  escorpio:    { nombre: 'Escorpio',    elemento: 'agua',   modo: 'fijo',     estilo: 'va al fondo y no suelta' },
  sagitario:   { nombre: 'Sagitario',   elemento: 'fuego',  modo: 'mutable',  estilo: 'apunta lejos y se aburre de lo pequeño' },
  capricornio: { nombre: 'Capricornio', elemento: 'tierra', modo: 'cardinal', estilo: 'construye despacio y aguanta lo que sea' },
  acuario:     { nombre: 'Acuario',     elemento: 'aire',   modo: 'fijo',     estilo: 'rompe la regla y se va por su lado' },
  piscis:      { nombre: 'Piscis',      elemento: 'agua',   modo: 'mutable',  estilo: 'se adapta, se disuelve y siente de más' },
};

/** Quién manda en cada signo. El regente del año sale de aquí. */
const LEC_REGENTES = {
  aries: 'marte', tauro: 'venus', geminis: 'mercurio', cancer: 'luna',
  leo: 'sol', virgo: 'mercurio', libra: 'venus', escorpio: 'pluton',
  sagitario: 'jupiter', capricornio: 'saturno', acuario: 'urano', piscis: 'neptuno',
};

/** El regente tradicional, para quien lee a la antigua. Se dicen los dos. */
const LEC_REGENTES_VIEJOS = {
  escorpio: 'marte', acuario: 'saturno', piscis: 'jupiter',
};

// ─── EL COMPOSITOR ───────────────────────────────────────────

/**
 * La intensidad de una temporada, de 0 a 100.
 *
 * No es un adorno: es lo que ordena la pantalla cuando hay ocho cosas
 * abiertas a la vez y ella tiene veinte minutos. Pesa tres cosas:
 *
 *   · el aspecto — una cuadratura se nota más que un sextil;
 *   · el planeta — Plutón deja marca, Mercurio pasa;
 *   · a quién le toca — los ángulos y las luminarias se sienten en el
 *     cuerpo; un generacional tocando a otro generacional, mucho menos.
 *
 * La cuenta está a la vista a propósito. Si algún día ella dice «esto
 * me pesó más de lo que decía Nova», se mueve un número y ya.
 */
function lecIntensidad_(cuerpoId, aspectoId, natalId) {
  const c = LEC_CUERPOS[cuerpoId];
  const a = LEC_ASPECTOS[aspectoId];
  if (!c || !a) return 40;

  const porAspecto = [22, 22, 34, 40][a.tension];       // 0..3
  const porCuerpo = { personal: 8, social: 26, generacional: 34 }[c.grupo] || 15;
  const angulos = ['ascendente', 'descendente', 'medio_cielo', 'fondo_cielo'];
  const luminarias = ['sol', 'luna'];
  const nodos = ['nodo_norte', 'nodo_sur'];
  const lentos = ['urano', 'neptuno', 'pluton'];

  let porNatal = 14;
  if (angulos.indexOf(natalId) !== -1) porNatal = 26;
  else if (luminarias.indexOf(natalId) !== -1) porNatal = 24;
  else if (nodos.indexOf(natalId) !== -1) porNatal = 20;
  else if (lentos.indexOf(natalId) !== -1) porNatal = 8;

  return Math.max(5, Math.min(100, porAspecto + porCuerpo + porNatal));
}

/**
 * La lectura completa de UNA temporada.
 *
 * Devuelve las cuatro respuestas que ella pidió, cada una por separado
 * para que la pantalla las pueda pintar como cuatro bloques y no como
 * un párrafo que nadie lee.
 */
function lecturaTransito_(t) {
  const cuerpoId = norm(t.cuerpo).replace(/ó/g, 'o').replace(/ú/g, 'u');
  const c = LEC_CUERPOS[cuerpoId];
  const a = lecAspecto_(t.aspecto);
  const n = lecNatal_(t.aNatal || t.a_natal || '');
  const casa = LEC_CASAS[Number(t.casa)] || null;

  // Sin planeta o sin aspecto no hay lectura posible, y decirlo es mejor
  // que inventar una genérica que sirva para cualquier cosa.
  if (!c || !a) {
    return { hay: false,
             porQue: 'Falta el planeta o el aspecto para poder leer esta temporada.' };
  }

  const titulo = c.nombre + ' ' + a.nombre.toLowerCase() +
                 (n ? ' a tu ' + n.nombre : '') +
                 (casa ? ' · por tu ' + casa.nombre.toLowerCase() : '');

  // 1 · TIEMPO PARA QUÉ
  const tiempoPara = casa
    ? 'Tiempo para ' + casa.tiempoPara + '.'
    : 'Tiempo para lo que ' + c.nombre + ' ' + c.hace.split(',')[0] + '.';

  // 2 · QUÉ ES / CÓMO TE AFECTA
  const queEs = c.nombre + ' ' + c.hace + ', y durante esta temporada lo hace sobre ' +
    (n ? n.quien : 'lo que estés moviendo') + '. ' +
    'El aspecto es ' + a.nombre.toLowerCase() + ': ' + a.que + '. ' +
    (casa ? 'Y se nota sobre todo en ' + casa.area + '.' : '');

  // 3 · CÓMO TRABAJARLA A FAVOR
  const aFavor = [];
  aFavor.push(a.pide.charAt(0).toUpperCase() + a.pide.slice(1) + '.');
  if (n) aFavor.push('Con ' + n.nombre + ' de por medio, lo concreto es: ' + n.aFavor + '.');
  if (casa) aFavor.push(casa.aFavor.charAt(0).toUpperCase() + casa.aFavor.slice(1) + '.');
  aFavor.push('Si sale bien, lo que deja es ' + c.regala + '.');

  // 4 · CÓMO SE VE EL DESEQUILIBRIO
  const mal = [];
  mal.push('La misma energía sin digerir se ve así: ' + c.tuerce + '.');
  mal.push('Del aspecto, lo que hay que vigilar es ' + a.riesgo + '.');
  if (n) mal.push('Y en la parte tuya que toca: ' + n.mal + '.');
  if (casa) mal.push('En el terreno de la casa: ' + casa.mal + '.');

  return {
    hay: true,
    titulo: titulo,
    cuerpo: c.nombre, aspecto: a.nombre, simbolo: a.simbolo,
    aNatal: n ? n.nombre : String(t.aNatal || t.a_natal || ''),
    casaNombre: casa ? casa.nombre : '',
    casaArea: casa ? casa.area : '',
    grupo: c.grupo,
    ritmo: c.ritmo,
    duraTipico: c.dura,
    tension: a.tension,
    intensidad: lecIntensidad_(cuerpoId, norm(a.nombre), (function () {
      const k = norm(t.aNatal || t.a_natal || '');
      const al = { 'mediocielo': 'medio_cielo', 'fondo del cielo': 'fondo_cielo',
                   'nodo norte': 'nodo_norte', 'nodo sur': 'nodo_sur' };
      return al[k] || k;
    })()),
    tiempoPara: tiempoPara,
    queEs: queEs,
    aFavor: aFavor,
    desequilibrio: mal,
    // La marca de siempre: esto es el punto de partida, no la verdad.
    nota: 'Lectura común de esta combinación. Lo que tú escribas en tu pensum manda sobre esto.',
  };
}

// ─── LA PROFECCIÓN, LEÍDA DE VERDAD ──────────────────────────

/**
 * El año profectado, con la profundidad que ella pidió con un ejemplo:
 * «por ejemplo en casa 8 de Leo por mi año».
 *
 * Una profección son tres datos encadenados, y los tres importan:
 *
 *   la CASA   → de qué va el año
 *   el SIGNO  → con qué estilo
 *   el REGENTE→ qué planeta lo dirige, y dónde está ÉSE en tu carta
 *
 * El tercero es el que casi nadie mira y el que más dice: el año de
 * casa 8 en Leo lo dirige el Sol, y si tu Sol está en casa 9, el año
 * de casa 8 se va a resolver por asuntos de casa 9.
 */
function lecProfeccion_(casa, ascendenteSigno, carta) {
  const signoIdx = (CIELO_SIGNOS.indexOf(norm(ascendenteSigno)) + casa - 1 + 12) % 12;
  const signo = CIELO_SIGNOS[signoIdx];
  const sg = LEC_SIGNOS[signo];
  const c = LEC_CASAS[casa];
  const regenteId = LEC_REGENTES[signo];
  const viejo = LEC_REGENTES_VIEJOS[signo];
  const reg = LEC_CUERPOS[regenteId];

  // Dónde está el regente en SU carta, si la tenemos sembrada.
  const natal = carta && carta.natal ? carta.natal : {};
  const donde = natal[regenteId] || null;
  const dondeViejo = viejo && natal[viejo] ? natal[viejo] : null;

  const partes = [];
  partes.push('Este año se lee por tu ' + c.nombre.toLowerCase() + ', en ' + sg.nombre +
              '. Va de ' + c.area + '.');
  partes.push('Tiempo para ' + c.tiempoPara + '.');
  partes.push('El estilo lo pone ' + sg.nombre + ', que ' + sg.estilo +
              ' — signo de ' + sg.elemento + ' y ' + sg.modo + '.');

  if (reg) {
    let f = 'Quien dirige el año es ' + reg.nombre + ', regente de ' + sg.nombre + ': ' +
            reg.hace + '.';
    if (donde) {
      const dc = LEC_CASAS[donde.casa];
      f += ' En tu carta, ' + reg.nombre + ' está en ' + (LEC_SIGNOS[donde.signo] || {}).nombre +
           ', casa ' + donde.casa + (dc ? ' — ' + dc.area : '') + '. ' +
           'Eso quiere decir que los asuntos de ' + c.nombre.toLowerCase() +
           ' se te van a resolver por ahí: ' + (dc ? dc.tiempoPara : 'por esa área') + '.';
    }
    partes.push(f);
  }
  if (dondeViejo && LEC_CUERPOS[viejo]) {
    partes.push('Si lees a la antigua, ' + sg.nombre + ' lo rige ' + LEC_CUERPOS[viejo].nombre +
                ', que tienes en casa ' + dondeViejo.casa + '. Los dos sirven; el moderno ' +
                'habla de la transformación y el tradicional de la acción concreta.');
  }

  // Si hay algo natal EN esa casa, el año lo despierta. Es el dato más
  // personal de toda la lectura y no cuesta nada mirarlo.
  const habitantes = [];
  Object.keys(natal).forEach(function (k) {
    if (['ascendente', 'descendente', 'medio_cielo', 'fondo_cielo'].indexOf(k) !== -1) return;
    if (Number(natal[k].casa) === casa) habitantes.push(natal[k]);
  });
  if (habitantes.length) {
    partes.push('Y hay algo tuyo viviendo ahí: ' +
      habitantes.map(function (h) {
        return h.nombre + ' en ' + (LEC_SIGNOS[h.signo] || {}).nombre;
      }).join(', ') + '. Un año de esa casa despierta lo que ya tenías puesto en ella, ' +
      'así que no es un tema nuevo: es uno tuyo de siempre, subido de volumen.');
  }

  return {
    casa: casa, signo: signo, signoNombre: sg.nombre,
    area: c.area, tiempoPara: c.tiempoPara,
    aFavor: c.aFavor, desequilibrio: c.mal,
    regente: reg ? reg.nombre : '', regenteId: regenteId,
    regenteDonde: donde ? { signo: donde.signo, casa: donde.casa } : null,
    habitantes: habitantes.map(function (h) {
      return { nombre: h.nombre, signo: h.signo, casa: h.casa };
    }),
    texto: partes,
  };
}

// ─── EL EJE DEL KARMA ────────────────────────────────────────

/**
 * Los nodos, el Mediocielo y el Fondo del cielo leídos como un eje.
 *
 * Ella los pidió para el pensum, y tiene sentido: un pensum kármico sin
 * los nodos es un calendario de tránsitos con otro nombre.
 *
 * Un eje no se lee por puntas sueltas. El Nodo Norte dice hacia dónde,
 * el Sur de dónde vienes, y la gracia está en el movimiento entre los
 * dos — no en «ser» el Norte y «dejar» el Sur.
 */
function lecEjes_(carta) {
  const natal = (carta && carta.natal) || {};
  const ejes = [];

  if (natal.nodo_norte && natal.nodo_sur) {
    const nn = natal.nodo_norte, ns = natal.nodo_sur;
    const cn = LEC_CASAS[nn.casa], cs = LEC_CASAS[ns.casa];
    ejes.push({
      id: 'nodos',
      nombre: 'Tu eje del karma',
      puntas: 'Nodo Sur en ' + (LEC_SIGNOS[ns.signo] || {}).nombre + ', casa ' + ns.casa +
              '  ·  Nodo Norte en ' + (LEC_SIGNOS[nn.signo] || {}).nombre + ', casa ' + nn.casa,
      deDonde: cs ? 'Lo que ya traes sabido está en ' + cs.area + '. Ahí te sale fácil, ' +
                    'y por eso mismo es donde te escondes: ' + cs.mal + '.' : '',
      haciaDonde: cn ? 'Hacia donde vas es ' + cn.area + '. Ahí te sientes torpe, y es ' +
                       'exactamente la señal de que es por ahí. ' + cn.aFavor + '.' : '',
      elMovimiento: cs && cn
        ? 'El movimiento del año: usar lo de casa ' + ns.casa + ' como herramienta, ' +
          'no como casa. Trabajas desde ' + cs.area.split(',')[0] + ' para construir en ' +
          cn.area.split(',')[0] + '.'
        : '',
    });
  }

  if (natal.medio_cielo && natal.fondo_cielo) {
    const mc = natal.medio_cielo, ic = natal.fondo_cielo;
    ejes.push({
      id: 'mc_ic',
      nombre: 'Tu eje de casa y carrera',
      puntas: 'Fondo del cielo en ' + (LEC_SIGNOS[ic.signo] || {}).nombre +
              '  ·  Mediocielo en ' + (LEC_SIGNOS[mc.signo] || {}).nombre,
      deDonde: 'La raíz: ' + LEC_CASAS[4].area + '. ' + LEC_CASAS[4].aFavor + '.',
      haciaDonde: 'Lo público: ' + LEC_CASAS[10].area + '. ' + LEC_CASAS[10].aFavor + '.',
      elMovimiento: 'Este eje es una balanza, no una carrera. Cuando el Mediocielo se ' +
                    'lleva todo, el Fondo del cielo pasa la cuenta en la casa y en el ' +
                    'cuerpo — y al revés.',
    });
  }

  // Si el Nodo Norte cae cerca de un ángulo, eso no es un detalle: es
  // la firma del mapa. Se dice, con el número, para que se pueda dudar.
  if (natal.nodo_norte && natal.medio_cielo) {
    /**
     * La separación angular entre dos grados del zodíaco, de 0 a 180.
     * El +540 y el -180 son el truco de siempre para que el cruce por
     * 0° Aries no dé 359 en vez de 1.
     */
    const sep = Math.abs(((natal.nodo_norte.grado - natal.medio_cielo.grado + 540) % 360) - 180);
    if (sep <= 10) {
      ejes.push({
        id: 'nn_mc',
        nombre: 'Tu Nodo Norte está pegado a tu Mediocielo',
        puntas: 'Separación: ' + sep.toFixed(1) + '°',
        deDonde: '',
        haciaDonde: 'Tu dirección kármica y tu carrera pública son la misma cosa. ' +
                    'No es que el trabajo «te quite tiempo» de lo espiritual: el ' +
                    'trabajo visible ES el camino.',
        elMovimiento: 'Lo que en otra carta sería ambición, en la tuya es tarea. ' +
                      'Esconder el trabajo pesa más de lo que pesaría en otra persona.',
      });
    }
  }

  return ejes;
}
