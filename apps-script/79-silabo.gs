/**
 * ═══════════════════════════════════════════════════════════
 *  NOVASOUL · LA UNIVERSIDAD
 * ═══════════════════════════════════════════════════════════
 *
 * ┌─ DÓNDE SE LEEN LOS ARCHIVOS, QUE ES LA PREGUNTA ───────────┐
 * │                                                            │
 * │ Los archivos NO se leen aquí, y no se copian a ninguna     │
 * │ hoja. Viven en su Drive, en la carpeta de cada materia, y  │
 * │ Nova guarda el ENLACE. Misma regla que con PHH: Nova sabe  │
 * │ dónde están las cosas, no qué dicen.                       │
 * │                                                            │
 * │ Lo que sí se lee es el TEXTO que ella pegue. Un cronograma │
 * │ de sílabo casi siempre es texto —«Parcial 1 · 15 de        │
 * │ octubre · 25%»— y de ahí se pueden sacar fechas con        │
 * │ aritmética, no con adivinanza.                             │
 * │                                                            │
 * │ LO QUE ESTO NO HACE: entender un PDF. Subir el archivo y   │
 * │ que Nova lo comprenda sola necesita un modelo de lenguaje, │
 * │ que cuesta plata por cada lectura. Esa decisión está       │
 * │ aparcada a propósito, no olvidada.                         │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * LA REGLA QUE HACE ESTO CONFIABLE: nunca escribe sola.
 *
 * El lector propone y ella confirma. Y no solo muestra lo que
 * encontró: muestra también **lo que ignoró y por qué**. Un lector que
 * solo enseña sus aciertos parece infalible y no lo es — lo peligroso
 * de un parcial no es que salga mal escrito, es que no salga.
 */

const MESES_ES = {
  enero: 1, ene: 1, febrero: 2, feb: 2, marzo: 3, mar: 3, abril: 4, abr: 4,
  mayo: 5, may: 5, junio: 6, jun: 6, julio: 7, jul: 7, agosto: 8, ago: 8,
  septiembre: 9, setiembre: 9, sep: 9, sept: 9, set: 9,
  octubre: 10, oct: 10, noviembre: 11, nov: 11, diciembre: 12, dic: 12,
};

/** Cuántas líneas se miran. Pegar un libro entero no es un cronograma. */
const SILABO_LINEAS_MAX = 300;

function dosDig_(n) { return (n < 10 ? '0' : '') + n; }

/**
 * El año, cuando el sílabo no lo dice.
 *
 * Un cronograma casi nunca escribe el año: dice «15 de octubre» porque
 * quien lo lee ya sabe de qué semestre habla. Se elige el año que pone
 * esa fecha lo más cerca posible de hoy sin quedar muy atrás —hasta
 * 60 días antes, para que un parcial de la semana pasada no salte al
 * año entrante.
 *
 * Y se devuelve `anioInferido` para poder decirlo en pantalla. Una
 * fecha adivinada que no se anuncia es una fecha inventada.
 */
function silaboAnio_(dia, mes, hoyISO) {
  const hoy = new Date(hoyISO + 'T00:00:00Z');
  const base = hoy.getUTCFullYear();
  let mejor = null, mejorDist = null;
  [base - 1, base, base + 1].forEach(function (a) {
    const f = new Date(Date.UTC(a, mes - 1, dia));
    if (isNaN(f.getTime()) || f.getUTCMonth() !== mes - 1) return;   // 31 de febrero
    const dist = Math.round((f - hoy) / 86400000);
    if (dist < -60) return;
    if (mejorDist === null || Math.abs(dist) < Math.abs(mejorDist)) {
      mejorDist = dist; mejor = a;
    }
  });
  return mejor === null ? base : mejor;
}

/**
 * Todas las fechas de una línea, en el orden en que aparecen.
 *
 * Devuelve también el trozo de texto que ocupó cada una, para poder
 * quitarlo del título: si no, el pendiente se llamaría «Parcial 1 15 de
 * octubre 25%», que es la línea entera y no el nombre de nada.
 */
function silaboFechas_(linea, hoyISO) {
  const out = [];
  const marcar = function (crudo, d, m, a, esRango) {
    if (!(d >= 1 && d <= 31) || !(m >= 1 && m <= 12)) return;
    const inferido = !a;
    const anio = a || silaboAnio_(d, m, hoyISO);
    const iso = anio + '-' + dosDig_(m) + '-' + dosDig_(d);
    // Una fecha imposible (31 de abril) se descarta en vez de correrse
    // sola al 1 de mayo, que es lo que hace new Date y nadie lo ve.
    const f = new Date(iso + 'T00:00:00Z');
    if (isNaN(f.getTime()) || f.getUTCDate() !== d) return;
    out.push({ iso: iso, crudo: crudo, anioInferido: inferido, rango: !!esRango });
  };

  /**
   * Primero los rangos: «del 12 al 16 de octubre».
   *
   * Van antes que todo porque mandan sobre el trozo entero de texto. De
   * un rango importa la ÚLTIMA fecha —es cuando se entrega, no cuando
   * empieza—, y tomando el trozo completo el título queda limpio en vez
   * de terminar en «Entrega del proyecto: semana del 12».
   */
  let re = /(?:del?\s+)?(\d{1,2})\s+al\s+(\d{1,2})\s+(?:de\s+)?([a-záéíóúñ]{3,10})\.?(?:\s*(?:de|del)?\s*(\d{4}))?/gi;
  let m;
  while ((m = re.exec(linea)) !== null) {
    const mes = MESES_ES[String(m[3]).toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')];
    if (!mes) continue;
    marcar(m[0], parseInt(m[2], 10), mes, m[4] ? parseInt(m[4], 10) : 0, true);
  }

  // 15 de octubre [de 2026]  ·  15 de oct
  re = /(\d{1,2})\s*(?:de\s+)?([a-záéíóúñ]{3,10})\.?(?:\s*(?:de|del)?\s*(\d{4}))?/gi;
  while ((m = re.exec(linea)) !== null) {
    const mes = MESES_ES[String(m[2]).toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')];
    if (!mes) continue;
    marcar(m[0], parseInt(m[1], 10), mes, m[3] ? parseInt(m[3], 10) : 0);
  }

  // octubre 15
  re = /([a-záéíóúñ]{3,10})\.?\s+(\d{1,2})(?![\d%])/gi;
  while ((m = re.exec(linea)) !== null) {
    const mes = MESES_ES[String(m[1]).toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '')];
    if (!mes) continue;
    marcar(m[0], parseInt(m[2], 10), mes, 0);
  }

  /**
   * 15/10/2026 · 15-10 · 15.10.26
   *
   * Con un dígito pegado a cada lado NO cuenta, y esa guarda no es
   * paranoia: sin ella el encabezado «ESTADÍSTICA 2026-2» entraba como
   * el 26 de febrero. Una fecha inventada a partir del título del
   * documento es justo el error que nadie revisa.
   */
  re = /(^|[^\d])(\d{1,2})[\/\-.](\d{1,2})(?:[\/\-.](\d{2,4}))?(?!\d)/g;
  while ((m = re.exec(linea)) !== null) {
    let a = m[4] ? parseInt(m[4], 10) : 0;
    if (a && a < 100) a += 2000;
    // Día primero: en toda América Latina 03/09 es 3 de septiembre.
    marcar(m[0].slice(m[1].length), parseInt(m[2], 10), parseInt(m[3], 10), a);
  }

  // Sin repetidas, y en el orden del texto
  const vistas = {}, limpio = [];
  out.sort(function (a, b) { return linea.indexOf(a.crudo) - linea.indexOf(b.crudo); });
  out.forEach(function (f) {
    if (vistas[f.iso]) return;
    vistas[f.iso] = 1;
    limpio.push(f);
  });
  return limpio;
}

/**
 * Lee un cronograma pegado y PROPONE entregas.
 *
 * No escribe nada. Devuelve dos listas: lo que entendió y lo que no,
 * con el motivo. La segunda importa igual que la primera.
 */
function silaboLeer_(texto, hoyISO) {
  const lineas = String(texto || '').split(/\r?\n/);
  const encontradas = [], ignoradas = [];
  let miradas = 0;

  lineas.forEach(function (cruda) {
    const linea = String(cruda).replace(/\s+/g, ' ').trim();
    if (!linea) return;
    if (miradas >= SILABO_LINEAS_MAX) return;
    miradas++;

    const fechas = silaboFechas_(linea, hoyISO);
    const pct = linea.match(/(\d{1,3})\s*%/);

    if (!fechas.length) {
      /**
       * Sin fecha no hay entrega. Se muestra igual, con el porqué: si
       * el cronograma traía algo que sí importaba y no tenía fecha,
       * ella lo ve y lo agrega a mano. Lo que no se puede hacer es
       * inventarle una fecha para que la lista quede bonita.
       */
      ignoradas.push({
        linea: linea,
        porque: pct ? 'Tiene un porcentaje pero ninguna fecha.' : 'No encontré ninguna fecha.',
      });
      return;
    }

    // De un rango («del 12 al 16 de octubre») manda la última: es
    // cuando se entrega, no cuando empieza.
    const elegida = fechas[fechas.length - 1];
    const hubieronVarias = fechas.length > 1 || fechas.some(function (f) { return f.rango; });

    let titulo = linea;
    fechas.forEach(function (f) { titulo = titulo.split(f.crudo).join(' '); });
    if (pct) titulo = titulo.split(pct[0]).join(' ');
    titulo = titulo
      .replace(/\b(lunes|martes|mi[eé]rcoles|jueves|viernes|s[áa]bado|domingo)\b/gi, ' ')
      .replace(/\b(del?|al|hasta|entrega|fecha|semana)\b\s*$/gi, ' ')
      .replace(/^[\s\-–—•*|,.;:>\d)]+/, '')
      .replace(/[\s\-–—•*|,.;:]+$/, '')
      .replace(/\s+/g, ' ')
      .trim();

    encontradas.push({
      titulo: titulo || linea,
      // Cuando no quedó título, se guarda la línea entera y se dice.
      sinTitulo: !titulo,
      fecha: elegida.iso,
      anioInferido: elegida.anioInferido,
      rango: hubieronVarias,
      peso: pct ? Number(pct[1]) : null,
      linea: linea,
    });
  });

  encontradas.sort(function (a, b) { return a.fecha < b.fecha ? -1 : 1; });
  return {
    encontradas: encontradas,
    ignoradas: ignoradas,
    truncado: lineas.filter(function (l) { return String(l).trim(); }).length > SILABO_LINEAS_MAX,
  };
}

// ─── LA API ──────────────────────────────────────────────────

/** Las materias, con cuántas entregas abiertas tiene cada una. */
function soulMaterias(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);

  const pend = soulLeer_('Pendientes', uid);
  const cuenta = {};
  pend.forEach(function (f) {
    const m = String(f.materia_id || '').trim();
    if (!m || soulHecho_(f)) return;
    if (!cuenta[m]) cuenta[m] = { abiertas: 0, vencidas: 0, proxima: '' };
    cuenta[m].abiertas++;
    const fe = aISO(f.fecha, 'UTC') || '';
    if (fe && fe < hoy) cuenta[m].vencidas++;
    if (fe && fe >= hoy && (!cuenta[m].proxima || fe < cuenta[m].proxima)) cuenta[m].proxima = fe;
  });

  return {
    ok: true, hoy: hoy,
    trabajos: soulTrabajos_().filter(function (t) { return t.tipo === 'estudio'; }),
    materias: soulLeer_('Materias', uid)
      .filter(function (f) { return String(f.id || '').trim(); })
      .map(function (f) {
        const c = cuenta[String(f.id)] || { abiertas: 0, vencidas: 0, proxima: '' };
        return {
          id: String(f.id), nombre: String(f.nombre || ''),
          codigo: String(f.codigo || ''), profesor: String(f.profesor || ''),
          // El enlace se devuelve tal cual quedó guardado. Nunca se
          // abre el archivo ni se copia su contenido a ninguna hoja.
          carpeta: String(f.carpeta || ''),
          semestre: String(f.semestre || ''),
          trabajoId: String(f.trabajo_id || ''),
          estado: norm(f.estado) || 'activa',
          nota: String(f.nota || ''),
          abiertas: c.abiertas, vencidas: c.vencidas, proxima: c.proxima,
        };
      }),
  };
}

function soulMateriaGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const d = p.datos || {};
  const esNueva = !String(d.id || '').trim();
  if (esNueva && !String(d.nombre || '').trim()) {
    return { ok: false, error: 'La materia necesita un nombre.' };
  }
  const carpeta = d.carpeta !== undefined ? String(d.carpeta).trim() : undefined;
  if (carpeta && !/^https?:\/\//i.test(carpeta)) {
    return { ok: false, error: 'El enlace de la carpeta tiene que empezar por http.' };
  }
  try {
    soulGuardar_('Materias', {
      id: String(d.id || ''),
      nombre: d.nombre !== undefined ? String(d.nombre).trim() : undefined,
      codigo: d.codigo !== undefined ? String(d.codigo).trim() : undefined,
      profesor: d.profesor !== undefined ? String(d.profesor).trim() : undefined,
      carpeta: carpeta,
      semestre: d.semestre !== undefined ? String(d.semestre).trim() : undefined,
      trabajo_id: d.trabajo_id !== undefined ? String(d.trabajo_id) : undefined,
      estado: d.estado !== undefined ? norm(d.estado) : (esNueva ? 'activa' : undefined),
      nota: d.nota !== undefined ? String(d.nota) : undefined,
    }, soulUsuario_(s));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulMateriaBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  /**
   * Borrar la materia NO borra sus entregas.
   *
   * Un parcial con fecha sigue teniendo fecha aunque se borre la
   * carpeta que lo mencionaba. Si se fueran juntos, un clic de limpieza
   * haría desaparecer en silencio lo único que no se puede incumplir.
   */
  const colgadas = soulLeer_('Pendientes', uid).filter(function (f) {
    return String(f.materia_id || '') === String(p.id) && !soulHecho_(f);
  }).length;
  try {
    const fue = soulBorrar_('Materias', p.id, uid);
    return { ok: fue, error: fue ? '' : 'No la encontré.', colgadas: colgadas };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/** Lee el cronograma pegado y devuelve la propuesta. NO escribe nada. */
function soulSilaboLeer(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const texto = String(p.texto || '');
  if (!texto.trim()) return { ok: false, error: 'Pega el cronograma primero.' };

  const r = silaboLeer_(texto, ahoraISO().slice(0, 10));
  return {
    ok: true,
    encontradas: r.encontradas, ignoradas: r.ignoradas, truncado: r.truncado,
    maximo: SILABO_LINEAS_MAX,
  };
}

/** Guarda SOLO lo que ella confirmó. */
function soulSilaboGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const materiaId = String(p.materia || '').trim();
  const items = p.items || [];
  if (!items.length) return { ok: false, error: 'No marcaste ninguna entrega.' };

  const materia = soulLeer_('Materias', uid)
    .filter(function (f) { return String(f.id) === materiaId; })[0];
  if (!materia) return { ok: false, error: 'Esa materia no existe.' };

  // Lo que ya está, no entra dos veces. Pegar el mismo sílabo otra vez
  // es lo más normal del mundo, y duplicar un parcial lo vuelve ruido.
  /**
   * La clave se arma con el texto COMPLETO, el mismo que se escribe.
   *
   * La primera versión comparaba «Parcial 1» contra «Estadística ·
   * Parcial 1» y nunca coincidía: pegar el sílabo dos veces metía el
   * parcial dos veces, y en el tablero eso se ve como dos parciales.
   */
  const nombreMateria = String(materia.nombre || 'Universidad');
  const clave = function (texto, fecha) { return norm(texto) + '|' + fecha; };
  const yaEstan = {};
  soulLeer_('Pendientes', uid).forEach(function (f) {
    if (String(f.materia_id || '') !== materiaId) return;
    yaEstan[clave(f.texto, aISO(f.fecha, 'UTC') || '')] = 1;
  });

  let creadas = 0, repetidas = 0;
  const errores = [];
  items.forEach(function (it) {
    const texto = String(it.titulo || '').trim();
    const fecha = String(it.fecha || '').trim();
    if (!texto || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
      errores.push(texto || '(sin título)');
      return;
    }
    const completo = nombreMateria + ' · ' + texto;
    if (yaEstan[clave(completo, fecha)]) { repetidas++; return; }
    try {
      soulGuardar_('Pendientes', {
        texto: completo,
        origen: 'silabo',
        materia_id: materiaId,
        trabajo_id: String(materia.trabajo_id || ''),
        fecha: fecha,
        estado: 'pendiente',
        // La fecha la pone la universidad, no ella: por eso inamovible.
        // Es lo que impide que el repartidor proponga correrlo.
        riesgo: 'inamovible',
        prioridad: it.peso && Number(it.peso) >= 20 ? 'alta' : 'media',
        horas_estimadas: it.horas !== undefined ? num(it.horas) : '',
        nota: it.peso ? 'Vale el ' + it.peso + '% de la materia' : '',
      }, uid);
      yaEstan[clave(completo, fecha)] = 1;
      creadas++;
    } catch (e) {
      errores.push(texto + ': ' + e.message);
    }
  });

  return { ok: true, creadas: creadas, repetidas: repetidas, errores: errores };
}
