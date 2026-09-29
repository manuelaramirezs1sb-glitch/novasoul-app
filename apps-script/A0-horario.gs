/**
 * ═══════════════════════════════════════════════════════════════
 *  QUE NOVA ORGANICE LA SEMANA
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «organiza las horas para tener de 10 a 12 horas de trabajo  │
 * │  y organización (…) que Nova organice mis horarios de       │
 * │  trabajo, Nutrea siempre es en la mañana y en la tarde,     │
 * │  reparte las horas de trabajo durante el día sin cruzarse   │
 * │  con otros proyectos».                                      │
 * │                                                            │
 * │ Con su esquema, y sus respuestas a lo que no podía inventar:│
 * │                                                            │
 * │   franja del día ····· 8:00 a 23:00                        │
 * │   Nutrea ············· lunes a sábado, 2-3 h, mañana Y tarde│
 * │   Nova ··············· 2-3 h por sesión                     │
 * │   carta de Sky ······· 2-4 h a la semana                    │
 * │   estudio ············ 1-2 h al día, aparte de las clases   │
 * │   ejercicio ·········· 30 min-1 h en casa · 1 h 30 en gym,  │
 * │                        gym 2-3 veces por semana             │
 * │   PHH ················ horas reservadas aunque falten pasos │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE ESTO NO HACE, Y ES LO MÁS IMPORTANTE ──────────────┐
 * │                                                            │
 * │ No mete a la fuerza lo que no cabe.                        │
 * │                                                            │
 * │ Sumando lo que pidió salen más de cuarenta horas de         │
 * │ proyectos sobre una semana que ya tiene clases y un turno   │
 * │ de doce horas el sábado. Un horario que lo acomoda todo     │
 * │ existe: basta con poner bloques encima de las clases, o     │
 * │ jornadas de dieciocho horas. Pero ese horario es una        │
 * │ mentira que se descubre el martes, y entonces ya no se      │
 * │ vuelve a abrir la pantalla.                                 │
 * │                                                            │
 * │ Así que Nova coloca lo que cabe, EN ORDEN, y lo que no      │
 * │ cabe lo deja fuera con su nombre y las horas que faltaron.  │
 * │ Una semana honesta con tres cosas fuera se puede negociar;  │
 * │ una semana perfecta que no se cumple, no.                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ NO ESCRIBE NADA ──────────────────────────────────┐
 * │                                                            │
 * │ Devuelve una PROPUESTA. No toca la rutina.                 │
 * │                                                            │
 * │ Un organizador que además guarda convierte cada prueba en   │
 * │ una limpieza: se mira, no gusta, y hay que borrar veinte    │
 * │ bloques a mano. Primero se ve; guardarlo es otra decisión,  │
 * │ y suya.                                                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** La ventana del día en la que se puede poner trabajo. */
const HORARIO_DESDE = 8;    // 8:00
const HORARIO_HASTA = 23;   // 23:00

/** Nada más corto que esto: media hora de proyecto no es una sesión. */
const HORARIO_MINIMO = 0.5;

/** Un respiro entre bloques, para que el día sea posible y no un tetris. */
const HORARIO_AIRE = 0.25;

/**
 * Las franjas, en horas del reloj.
 *
 * «Nutrea siempre es en la mañana y en la tarde» — por eso existen: sin
 * franjas, el algoritmo pondría las dos sesiones de Nutrea seguidas a
 * las ocho de la mañana y habría cumplido la letra sin cumplir nada.
 */
const HORARIO_FRANJAS = {
  manana:     [8, 13],
  tarde:      [13, 19],
  noche:      [19, 23],
  cualquiera: [HORARIO_DESDE, HORARIO_HASTA],
};

/**
 * Lo que ella quiere que quepa cada semana.
 *
 * Vive en la hoja `Bloques` para que lo pueda cambiar sin tocar código:
 * subir Nova a cuatro horas o apagar PHH una semana es algo que va a
 * querer hacer, y no debería costar un despliegue.
 *
 * `orden` decide quién entra primero cuando no cabe todo. No es un
 * capricho: lo que sostiene la plata y los compromisos con otros va
 * antes que lo que se puede correr una semana sin consecuencia.
 */
function bloquesDe_(uid) {
  const filas = soulLeerSuave_('Bloques', uid, []);
  return filas.filter(function (f) {
    return norm(f.activo) !== 'no' && String(f.nombre || '').trim();
  }).map(function (f) {
    const min = num(f.horas_min) || 0;
    const max = num(f.horas_max) || min;
    return {
      id: String(f.id || ''),
      nombre: String(f.nombre || '').trim(),
      trabajoId: String(f.trabajo_id || '').trim(),
      min: min, max: Math.max(min, max),
      cada: norm(f.cada) === 'semana' ? 'semana' : 'dia',
      veces: num(f.veces) || 0,
      franja: HORARIO_FRANJAS[norm(f.franja)] ? norm(f.franja) : 'cualquiera',
      dias: diasDe_(f.dias),
      partes: Math.max(1, num(f.partes) || 1),
      orden: num(f.orden) || 50,
      nota: String(f.nota || ''),
    };
  }).sort(function (a, b) { return a.orden - b.orden; });
}

/**
 * ═══════════════════════════════════════════════════════════════
 *   UN SOLO NÚMERO, NO DOS
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE ELLA PREGUNTÓ ANTES DE PEGAR NADA ─────────────────┐
 * │                                                            │
 * │ «¿está el horario conectado con los proyectos de Central y  │
 * │  lo que hay en mi horario, sin copias ni duplicados? ¿es    │
 * │  decir, cruzas la info que ya tienes con la que te di si    │
 * │  hay algo repetido como las horas de Nutrea o Nova?»        │
 * │                                                            │
 * │ No lo estaba, y tenía razón en preguntar. Había DOS listas  │
 * │ viviendo en paralelo:                                       │
 * │                                                            │
 * │   Trabajos.horas_semana ··· lo que puso en Central. De ahí  │
 * │                             salía el «44 h comprometidas».  │
 * │   Bloques ················· lo que le dijo a Nova que       │
 * │                             organizara.                     │
 * │                                                            │
 * │ Nutrea y Nova están en las dos. Sin cruzarlas, la tarjeta   │
 * │ de la semana sumaba las de Central y el organizador          │
 * │ colocaba las de Bloques: dos números distintos para lo      │
 * │ mismo, en dos pantallas que se miran seguidas. Exactamente  │
 * │ el error que arreglamos en el punto de equilibrio.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ CÓMO SE CRUZAN ───────────────────────────────────────────┐
 * │                                                            │
 * │ Un bloque puede apuntar a un proyecto (`trabajo_id`). Si    │
 * │ apunta, MANDA el bloque: es lo más detallado —cuántas       │
 * │ horas, qué días, en qué franja— y lo que de verdad se va a  │
 * │ colocar en la semana. Las `horas_semana` de Central siguen  │
 * │ ahí para todo lo demás, pero ya no se suman otra vez.      │
 * │                                                            │
 * │ Un proyecto SIN bloque conserva sus horas de Central: no    │
 * │ desaparece de la cuenta por no estar en la lista nueva.     │
 * │                                                            │
 * │ Y un bloque sin proyecto —estudio, gym, yoga— suma lo suyo. │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
function semanaComprometida_(uid, activos) {
  const bloques = bloquesDe_(uid);
  const porTrabajo = {};
  const detalle = [];

  bloques.forEach(function (b) {
    // Cuántas horas pide a la semana este bloque, en su forma mínima:
    // lo que se promete es el mínimo; el máximo es lo que se aprovecha
    // si el día da. Prometer el máximo hace que la semana no cuadre casi
    // nunca, y una alarma que salta siempre no se lee.
    const horas = b.cada === 'dia'
      ? b.min * b.dias.length
      : b.min * Math.max(1, b.veces || 1);
    if (b.trabajoId) porTrabajo[b.trabajoId] = (porTrabajo[b.trabajoId] || 0) + horas;
    detalle.push({ id: b.trabajoId || b.id, nombre: b.nombre, horas: horas,
                   deBloque: true });
  });

  (activos || []).forEach(function (t) {
    // Si ya hay un bloque para este proyecto, el bloque manda: no se
    // suma dos veces lo mismo.
    if (porTrabajo[t.id] !== undefined) return;
    if (!(t.horasSemana > 0)) return;
    detalle.push({ id: t.id, nombre: t.nombre, horas: t.horasSemana, deBloque: false });
  });

  return {
    total: detalle.reduce(function (a, x) { return a + x.horas; }, 0),
    detalle: detalle.sort(function (a, b) { return b.horas - a.horas; }),
    hayBloques: bloques.length > 0,
  };
}

/** "1-6", "1,2,4" o vacío (todos). */
function diasDe_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return [1, 2, 3, 4, 5, 6, 7];
  const out = [];
  s.split(/[,;]/).forEach(function (t) {
    const r = t.trim().match(/^(\d)\s*-\s*(\d)$/);
    if (r) {
      for (let i = Number(r[1]); i <= Number(r[2]); i++) out.push(i);
    } else if (/^\d$/.test(t.trim())) {
      out.push(Number(t.trim()));
    }
  });
  return out.length ? out : [1, 2, 3, 4, 5, 6, 7];
}

/**
 * Los huecos de un día: lo que queda entre los bloques fijos.
 *
 * Los fijos son las clases y los turnos. No se tocan ni se mueven: son
 * compromisos con otras personas, y un organizador que los pisa no es
 * un organizador.
 */
function huecosDelDia_(fijos, tope) {
  const ocupados = fijos.map(function (b) {
    const a = horaNum_(b.inicio);
    return { a: a, b: a + duracionHoras_(b.inicio, b.fin) };
  }).sort(function (x, y) { return x.a - y.a; });

  const libres = [];
  let cursor = HORARIO_DESDE;
  ocupados.forEach(function (o) {
    // Un turno que cruza la medianoche se recorta a la ventana del día.
    const ini = Math.max(o.a, HORARIO_DESDE), fin = Math.min(o.b, HORARIO_HASTA);
    if (fin <= HORARIO_DESDE || ini >= HORARIO_HASTA) return;
    if (ini - cursor >= HORARIO_MINIMO) libres.push({ a: cursor, b: ini });
    cursor = Math.max(cursor, fin);
  });
  if (HORARIO_HASTA - cursor >= HORARIO_MINIMO) libres.push({ a: cursor, b: HORARIO_HASTA });

  /**
   * Y el tope de horas útiles del día manda sobre el reloj.
   *
   * Ella puede tener la tarde libre y aun así no querer trabajar diez
   * horas ese día. El reloj dice dónde CABE; sus horas útiles dicen
   * cuánto QUIERE. Gana la segunda.
   *
   * ── Y LAS CLASES CUENTAN DENTRO DE ESE TOPE ──
   *
   * Ella: «el sábado pueden ser 15, teniendo en cuenta que trabajo 12 h
   * en Salsabor». O sea: quince EN TOTAL, turno incluido — quedan tres
   * para todo lo demás.
   *
   * La primera versión no restaba lo fijo y dejaba poner quince horas
   * de proyectos ENCIMA de las doce del turno: un sábado de veintisiete
   * horas. Lo encontró la prueba, no yo.
   */
  if (tope !== null && tope !== undefined) {
    const yaFijas = fijos.reduce(function (a, b) {
      return a + duracionHoras_(b.inicio, b.fin);
    }, 0);
    let queda = Math.max(0, tope - yaFijas);
    const cortados = [];
    libres.forEach(function (h) {
      if (queda <= 0) return;
      const largo = Math.min(h.b - h.a, queda);
      if (largo >= HORARIO_MINIMO) { cortados.push({ a: h.a, b: h.a + largo }); queda -= largo; }
    });
    return cortados;
  }
  return libres;
}

/** El primer hueco donde cabe `horas` dentro de la franja pedida. */
function cabeEn_(huecos, horas, franja) {
  const f = HORARIO_FRANJAS[franja] || HORARIO_FRANJAS.cualquiera;
  for (let i = 0; i < huecos.length; i++) {
    const h = huecos[i];
    const ini = Math.max(h.a, f[0]);
    const fin = Math.min(h.b, f[1]);
    if (fin - ini >= horas) return { i: i, a: ini, b: ini + horas };
  }
  return null;
}

/** Quita del hueco lo que se acaba de ocupar, dejando aire después. */
function ocupar_(huecos, puesto) {
  const h = huecos[puesto.i];
  const nuevos = [];
  if (puesto.a - h.a >= HORARIO_MINIMO) nuevos.push({ a: h.a, b: puesto.a });
  const resto = puesto.b + HORARIO_AIRE;
  if (h.b - resto >= HORARIO_MINIMO) nuevos.push({ a: resto, b: h.b });
  huecos.splice.apply(huecos, [puesto.i, 1].concat(nuevos));
}

/**
 * La semana organizada.
 *
 * Devuelve los bloques colocados día por día, y —tan importante como
 * eso— lo que NO cupo, con su nombre y las horas que le faltaron.
 */
function soulHorario(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);
  const lunes = lunesDe_(String(p.lunes || hoy));

  const bloques = bloquesDe_(uid);
  if (!bloques.length) {
    return { ok: true, lunes: lunes, dias: [], fuera: [], sinReglas: true,
             porque: 'Todavía no has dicho qué quieres que quepa en tu semana. ' +
                     'Corre sembrarBloques() una vez y después ajústalo en la hoja Bloques.' };
  }

  const rutinas = rutinaDe_(uid);
  const utiles = horasUtilesDe_(uid);

  // Los huecos de cada día, ya descontadas las clases y los turnos
  const dias = [];
  const huecos = {};
  for (let d = 1; d <= 7; d++) {
    const fecha = masDias_(lunes, d - 1);
    const fijos = rutinaDelDia_(rutinas, fecha);
    const tope = utiles ? (utiles[d] === undefined ? null : utiles[d]) : null;
    huecos[d] = huecosDelDia_(fijos, tope);
    dias.push({
      dia: d, fecha: fecha,
      fijos: fijos.map(function (b) {
        return { nombre: b.nombre, inicio: hhmm_(horaNum_(b.inicio)),
                 fin: hhmm_(horaNum_(b.inicio) + duracionHoras_(b.inicio, b.fin)),
                 horas: duracionHoras_(b.inicio, b.fin), tipo: b.tipo };
      }),
      puestos: [],
      utiles: tope,
    });
  }

  const fuera = [];
  const poner = function (b, dia, horas, franja) {
    const puesto = cabeEn_(huecos[dia], horas, franja);
    if (!puesto) return false;
    ocupar_(huecos[dia], puesto);
    dias[dia - 1].puestos.push({
      bloque: b.id, nombre: b.nombre, trabajoId: b.trabajoId,
      inicio: hhmm_(puesto.a), fin: hhmm_(puesto.b), horas: horas,
      franja: franja,
    });
    return true;
  };

  bloques.forEach(function (b) {
    /**
     * Se intenta con el MÁXIMO y se baja hasta el mínimo.
     *
     * Pedir «2 a 3 horas» y recibir siempre 2 sería cumplir la letra y
     * fallar la intención: el rango existe para aprovechar los días
     * buenos, no para redondear a la baja todos.
     */
    const escalones = [];
    for (let h = b.max; h >= b.min - 0.001; h -= 0.5) escalones.push(Math.round(h * 2) / 2);
    if (!escalones.length) escalones.push(b.min);

    if (b.cada === 'dia') {
      b.dias.forEach(function (d) {
        let faltan = 0;
        for (let parte = 0; parte < b.partes; parte++) {
          /**
           * Dos partes, dos franjas. «Nutrea siempre es en la mañana y
           * en la tarde»: sin esto el algoritmo pondría las dos seguidas
           * a las ocho y habría cumplido la letra sin cumplir nada.
           */
          const franja = b.partes === 2
            ? (parte === 0 ? 'manana' : 'tarde')
            : b.franja;
          /**
           * Si en su franja no cabe, se prueba MÁS TARDE — nunca antes.
           *
           * La primera versión caía a «cualquiera», y el viernes —con
           * Procesos y Laboratorio ocupando toda la tarde— acababa
           * poniendo las DOS sesiones de Nutrea por la mañana, seguidas.
           * Eso cumple el número de horas e incumple lo único que ella
           * pidió de Nutrea: que sea mañana y tarde.
           *
           * Así que la tarde puede correrse a la noche, pero no volver a
           * la mañana. Y si tampoco cabe de noche, no se mete a empujones:
           * se dice que ese día faltó.
           */
          const alternativas = franja === 'manana' ? ['manana']
            : franja === 'tarde' ? ['tarde', 'noche']
            : [franja];
          let puesto = false;
          for (let a = 0; a < alternativas.length && !puesto; a++) {
            for (let e = 0; e < escalones.length && !puesto; e++) {
              const horas = Math.max(HORARIO_MINIMO, escalones[e] / b.partes);
              puesto = poner(b, d, horas, alternativas[a]);
            }
          }
          if (!puesto) faltan += b.min / b.partes;
        }
        if (faltan > 0) fuera.push({ nombre: b.nombre, dia: d, horas: faltan });
      });
    } else {
      // Por semana: se reparte en `veces` sesiones, en los días permitidos
      const veces = Math.max(1, b.veces || 1);
      const porVez = Math.max(HORARIO_MINIMO, Math.round((b.max / veces) * 2) / 2);
      let hechas = 0;
      /**
       * Se empieza por los días con MÁS hueco.
       *
       * Meterlo en el primer día que aparezca llena el lunes y deja el
       * jueves vacío. Repartir por holgura da una semana que se parece
       * a una semana.
       */
      const orden = b.dias.slice().sort(function (x, y) {
        const lx = huecos[x].reduce(function (a, h) { return a + (h.b - h.a); }, 0);
        const ly = huecos[y].reduce(function (a, h) { return a + (h.b - h.a); }, 0);
        return ly - lx;
      });
      orden.forEach(function (d) {
        if (hechas >= veces) return;
        if (poner(b, d, porVez, b.franja)) hechas++;
      });
      if (hechas < veces) {
        fuera.push({ nombre: b.nombre, sesiones: veces - hechas,
                     horas: (veces - hechas) * porVez });
      }
    }
  });

  // Cada día, ordenado por hora
  dias.forEach(function (x) {
    x.puestos.sort(function (a, b) { return horaNum_(a.inicio) - horaNum_(b.inicio); });
    x.horasPuestas = x.puestos.reduce(function (a, b) { return a + b.horas; }, 0);
    x.horasFijas = x.fijos.reduce(function (a, b) { return a + b.horas; }, 0);
    x.libre = huecos[x.dia].reduce(function (a, h) { return a + (h.b - h.a); }, 0);
  });

  return {
    ok: true, lunes: lunes, desde: hhmm_(HORARIO_DESDE), hasta: hhmm_(HORARIO_HASTA),
    dias: dias,
    fuera: fuera,
    puestas: dias.reduce(function (a, d) { return a + d.horasPuestas; }, 0),
    faltaron: fuera.reduce(function (a, f) { return a + f.horas; }, 0),
    // Que quede dicho: esto es una propuesta, no la rutina.
    propuesta: true,
  };
}

/** Las horas útiles que ella declaró, por día de la semana. */
function horasUtilesDe_(uid) {
  const filas = soulLeerSuave_('Horas', uid, []);
  if (!filas.length) return null;
  const out = {};
  filas.forEach(function (f) {
    const d = num(f.dia_semana);
    if (d >= 1 && d <= 7) out[d] = num(f.horas_libres);
  });
  return Object.keys(out).length ? out : null;
}

/**
 * ★ Siembra el esquema de trabajo que ella describió. Se corre UNA vez.
 *
 * Después se edita en la hoja `Bloques`: subir Nova a cuatro horas,
 * apagar PHH una semana, mover el estudio a la noche. Nada de eso
 * debería costar un despliegue.
 *
 * El `orden` es lo que decide quién entra cuando no cabe todo, y por eso
 * no es alfabético: primero lo que sostiene la plata y los compromisos
 * con otras personas, al final lo que se puede correr sin que nadie se
 * entere.
 */
/**
 * ═══════════════════════════════════════════════════════════════
 *   DE QUIÉN SON LAS COSAS DE SOUL, SIN QUE HAYA QUE ESCRIBIRLO
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE PASÓ ──────────────────────────────────────────────┐
 * │                                                            │
 * │   Error: sembrarBloques("tucorreo@…") — di de quién son.    │
 * │                                                            │
 * │ No es un fallo del código: es un fallo de diseño mío. Esta  │
 * │ función se corre UNA VEZ, a mano, desde el botón            │
 * │ «Ejecutar» del editor de Apps Script — y ese botón llama    │
 * │ sin argumentos. O sea que la escribí de forma que no se     │
 * │ puede usar por el único camino por el que se iba a usar.    │
 * │                                                            │
 * │ `bootstrapTodo()` no tiene ese problema y por eso nunca lo  │
 * │ noté: no le hace falta saber quién eres.                    │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ DE DÓNDE SALE EL CORREO, Y EN QUÉ ORDEN ──────────────────┐
 * │                                                            │
 * │ 1. El que se pase como argumento. Siempre manda.            │
 * │ 2. Quien esté corriendo el script. Es quien tocó el botón.  │
 * │ 3. El correo que YA está escrito en las hojas de Soul —     │
 * │    Rutina, Trabajos, Horas—. Si hay exactamente UNO, es     │
 * │    ella; si hay varios, NO SE ADIVINA.                      │
 * │                                                            │
 * │ El tercero existe porque los dos primeros pueden fallar:    │
 * │ Google a veces devuelve vacío en `getActiveUser`, y un      │
 * │ correo de Google puede no ser el mismo con el que entra a   │
 * │ Nova. Las hojas dicen la verdad de este producto.           │
 * │                                                            │
 * │ Con dos correos distintos se rinde y lo dice, en vez de     │
 * │ sembrarle a una persona los bloques de otra. Un horario     │
 * │ puesto en la cuenta equivocada es peor que un error.        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
function soulDeQuien_(uid) {
  const dado = String(uid || '').toLowerCase().trim();
  if (dado) return dado;

  // 2 · Quien tocó el botón
  try {
    const s = Session.getActiveUser().getEmail() ||
              Session.getEffectiveUser().getEmail();
    if (s) return String(s).toLowerCase().trim();
  } catch (e) { /* en algunos contextos Google no lo da; seguimos */ }

  // 3 · Lo que ya dicen las hojas
  const vistos = {};
  ['Rutina', 'Trabajos', 'Horas', 'Materias', 'Pendientes'].forEach(function (tab) {
    let d;
    try { d = soulCrudo_(tab); } catch (e) { return; }
    if (!d || !d.length) return;
    const cU = d[0].map(norm).indexOf('usuario_id');
    if (cU === -1) return;
    d.slice(1).forEach(function (f) {
      const c = String(f[cU] || '').toLowerCase().trim();
      if (c) vistos[c] = (vistos[c] || 0) + 1;
    });
  });
  const correos = Object.keys(vistos);
  if (correos.length === 1) return correos[0];
  if (correos.length > 1) {
    throw new Error('En estas hojas hay ' + correos.length + ' personas (' +
      correos.join(', ') + '). Dime de quién: sembrarBloques("elcorreo@…").');
  }
  throw new Error('No pude saber de quién son. Córrelo así: ' +
    'sembrarBloques("tucorreo@…") — el mismo con el que entras a NovaSoul.');
}

function sembrarBloques(uid) {
  const yo = soulDeQuien_(uid);

  const sh = soulSheet_('Bloques');
  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  const ya = soulLeerSuave_('Bloques', yo, []).length;
  // También por Logger: si no, correrlo dos veces no dice nada y parece
  // que la segunda vez tampoco hizo nada… que es cierto, pero hay que
  // poder saberlo.
  if (ya) return soulDecir_('Ya tienes ' + ya + ' bloques. No toco nada: ' +
    'edítalos en la hoja Bloques.');

  const filas = [
    // nombre, tipo, min, max, cada, veces, franja, dias, partes, orden
    ['Nutrea · tiendas', 'trabajo', 2, 3, 'dia', 0, 'cualquiera', '1-6', 2, 10],
    ['Nova', 'trabajo', 2, 3, 'semana', 3, 'tarde', '1-6', 1, 20],
    ['Universidad · estudio', 'estudio', 1, 2, 'dia', 0, 'cualquiera', '1-6', 1, 30],
    ['Carta de Sky', 'trabajo', 2, 4, 'semana', 2, 'cualquiera', '1,2,4', 1, 40],
    ['PHH', 'trabajo', 2, 3, 'semana', 1, 'cualquiera', '1,2,4', 1, 50],
    ['Gym', 'ejercicio', 1.5, 1.5, 'semana', 3, 'manana', '1-6', 1, 60],
    ['Yoga o pilates en casa', 'ejercicio', 0.5, 1, 'semana', 3, 'cualquiera', '1-7', 1, 70],
  ];

  /**
   * ── Y SE ENLAZAN CON LOS PROYECTOS QUE YA EXISTEN ──
   *
   * Ella preguntó justo esto antes de pegar nada: «¿cruzas la info que
   * ya tienes con la que te di si hay algo repetido como las horas de
   * Nutrea o Nova?».
   *
   * Sin el enlace, «Nutrea» aquí y «Nutrea» en Central son dos cosas
   * para Nova, y sus horas se suman dos veces. Con él, el bloque manda
   * y las `horas_semana` de Central dejan de contarse aparte.
   *
   * El emparejado es por nombre, y es flojo a propósito: «Nutrea ·
   * tiendas» tiene que encontrar a «Nutrea». Lo que no encuentre queda
   * sin enlazar y se puede arreglar a mano en la hoja — mejor eso que
   * enlazarlo al proyecto equivocado por parecerse un poco.
   */
  const proyectos = soulTrabajos_().filter(function (x) { return x.estado === 'activo'; });
  const enlazar = function (nombre) {
    const n = norm(nombre);
    let mejor = null;
    proyectos.forEach(function (pr) {
      const pn = norm(pr.nombre);
      if (!pn) return;
      if (n === pn || n.indexOf(pn) !== -1 || pn.indexOf(n) !== -1) {
        if (!mejor || pn.length > norm(mejor.nombre).length) mejor = pr;
      }
    });
    return mejor ? mejor.id : '';
  };

  const enlazados = [];
  filas.forEach(function (f, i) {
    const trabajoId = enlazar(f[0]);
    if (trabajoId) enlazados.push(f[0]);
    const o = {
      id: 'b' + (i + 1) + Utilities.getUuid().slice(0, 4),
      usuario_id: yo, nombre: f[0], tipo: f[1], trabajo_id: trabajoId,
      horas_min: f[2], horas_max: f[3], cada: f[4], veces: f[5],
      franja: f[6], dias: f[7], partes: f[8], orden: f[9], activo: 'si', nota: '',
    };
    sh.appendRow(enc.map(function (c) { return o[c] !== undefined ? o[c] : ''; }));
  });
  soulOlvidar_('Bloques');
  /**
   * Y SE LOGUEA, no solo se devuelve.
   *
   * «ya esto me salió en lo de sembrar» — y lo que le salió fue
   * «Se ha completado la ejecución» y nada más. El registro del editor
   * de Apps Script solo enseña lo que pasa por `Logger.log`: el valor
   * que una función DEVUELVE no se ve por ningún lado.
   *
   * O sea que sembró siete bloques, los enlazó con sus proyectos de
   * Central —que es justo lo que ella había preguntado antes de pegar
   * nada— y no se enteró de ninguna de las dos cosas. Un mensaje que
   * nadie puede leer es lo mismo que no escribirlo.
   */
  return soulDecir_('Listo: ' + filas.length + ' bloques sembrados.\n' +
    (enlazados.length
      ? 'Enlazados con tus proyectos de Central: ' + enlazados.join(', ') +
        '. Sus horas ya NO se cuentan dos veces.'
      : 'Ninguno calzó con un proyecto de Central: si alguno debería, ponle el ' +
        'trabajo_id a mano en la hoja Bloques.') +
    '\nAjusta lo que quieras en la hoja Bloques.');
}
