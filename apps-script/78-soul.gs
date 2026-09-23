/**
 * ═══════════════════════════════════════════════════════════
 *  NOVASOUL · EL DÍA A DÍA
 * ═══════════════════════════════════════════════════════════
 *
 * Nova Central y NovaSoul comparten un solo objeto —el PROYECTO— y lo
 * miran por lados distintos. Central sabe cuánto vale, qué se cobró y
 * qué falta cobrar. Soul sabe qué hay que entregar, cuándo, y cuánto
 * cuesta en horas.
 *
 * ┌─ LA FRONTERA, QUE ES DE PRIVACIDAD ────────────────────────┐
 * │                                                            │
 * │ De Central BAJA a Soul la lista de proyectos: su nombre,   │
 * │ su tipo y sus horas fijas. Nada de plata: Soul no necesita │
 * │ saber cuánto paga PHH para repartir una semana.            │
 * │                                                            │
 * │ De Soul SUBE a Central solo el peso: cuántas tareas        │
 * │ abiertas tiene un proyecto y cuántas horas suman. EL TEXTO │
 * │ NO SUBE. Manuela dejó dicho que lo de PHH es confidencial  │
 * │ y que lo comparte con Nova solo para organizarse; Central  │
 * │ se abre delante de una socia o un contador, y Soul no se   │
 * │ abre delante de nadie.                                     │
 * │                                                            │
 * │ Esa regla no es una nota en un documento: es la razón por  │
 * │ la que `soulCargaPorTrabajo_` devuelve números y no filas. │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * SE ENTRA CON LA SESIÓN DE CENTRAL. Ella lo decidió así: Soul es
 * suya y solo suya, y pedir un segundo código para la misma persona en
 * la misma máquina no protege nada — enseña a saltarse la puerta.
 */

/** Las hojas de Soul con id propio, y sus columnas. */
const SOUL_HOJAS = {
  Pendientes: ['id','usuario_id','texto','tipo','origen','fecha','hecho',
               'hecho_en','plataforma_id','trabajo_id','estado','prioridad',
               'horas_estimadas','horas_reales','riesgo','nota','materia_id'],
  Materias:   ['id','usuario_id','nombre','codigo','profesor','carpeta',
               'semestre','trabajo_id','estado','nota'],
  Mindlab:    ['id','usuario_id','semana','mes','tema','tarea','horas_estimadas',
               'desde','hasta','estado','nota'],
  Fijos:      ['id','usuario_id','categoria','concepto','monto','moneda',
               'dia_del_mes','activo','nota'],
};

/** Las tres columnas del tablero. Un pendiente está en una sola. */
const SOUL_ESTADOS = ['pendiente', 'haciendo', 'hecho'];

/**
 * Qué pasa si NO se entrega. No es lo mismo que la prioridad.
 *
 * La prioridad dice qué quiere hacer primero; el riesgo dice qué puede
 * correrse sin romper nada. Cuando la semana no cabe, lo que decide
 * cuál se cae es esta columna, no la otra.
 */
const SOUL_RIESGOS = {
  inamovible: { nombre: 'No se puede mover', orden: 3,
                ayuda: 'Un parcial, una entrega con fecha de ellos.' },
  acordado:   { nombre: 'Acordado con alguien', orden: 2,
                ayuda: 'Se puede correr, pero hay que avisar.' },
  corrible:   { nombre: 'Se puede correr', orden: 1,
                ayuda: 'Nadie está esperando una fecha exacta.' },
};

const SOUL_PRIORIDADES = { alta: 3, media: 2, baja: 1 };

/** Los gastos fijos que ella nombró, en el orden en que los nombró. */
const SOUL_CATEGORIAS = [
  { id: 'arriendo',  nombre: 'Arriendo',      tipo: 'gasto'  },
  { id: 'mercado',   nombre: 'Mercado',       tipo: 'gasto'  },
  { id: 'servicios', nombre: 'Servicios',     tipo: 'gasto'  },
  { id: 'internet',  nombre: 'Internet',      tipo: 'gasto'  },
  { id: 'credito',   nombre: 'Crédito',       tipo: 'gasto'  },
  { id: 'deudas',    nombre: 'Deudas',        tipo: 'gasto'  },
  { id: 'movil',     nombre: 'Móvil',         tipo: 'gasto'  },
  { id: 'varios',    nombre: 'Gastos varios', tipo: 'gasto'  },
  { id: 'ahorro',    nombre: 'Ahorro',        tipo: 'ahorro' },
];

/**
 * MINDLAB · doce semanas, recortado para este último trimestre.
 *
 * El programa es suyo y los doce temas son los suyos. Lo que se ajustó
 * aquí es el calendario y el peso, porque me pidió que no la
 * sobrecargara:
 *
 *   · Arranca el lunes 28 de septiembre y cierra el domingo 20 de
 *     diciembre. Las dos últimas semanas del año quedan libres: son
 *     finales en la universidad y fiestas, y un plan que las ocupa se
 *     incumple el primer día.
 *   · Cada semana trae UNA tarea y un estimado honesto. Suman 37 horas
 *     en doce semanas — unas tres por semana, media hora al día.
 *   · Las semanas 11 y 12 caen en diciembre y van a dos horas.
 *
 * Es el único compromiso sin cliente que reclame. Por eso está marcado
 * como meta: para que NovaSoul pueda decirle «esta semana no tocaste lo
 * único que es tuyo», que es lo que ninguna lista de tareas hace.
 */
const MINDLAB_INICIO = '2026-09-28';
const MINDLAB_PLAN = [
  { semana: 1,  mes: 1, tema: 'Fundamentos de marca',    horas: 3,
    tarea: 'Optimizar el perfil y definir el posicionamiento.' },
  { semana: 2,  mes: 1, tema: 'Investigación y estrategia', horas: 4,
    tarea: 'Crear 30 ideas y 15 guiones.' },
  { semana: 3,  mes: 1, tema: 'Producción',               horas: 4,
    tarea: 'Grabar entre 8 y 12 piezas.' },
  { semana: 4,  mes: 1, tema: 'Edición',                  horas: 3,
    tarea: 'Editar y dejar programado lo grabado.' },
  { semana: 5,  mes: 2, tema: 'Historias que conectan',   horas: 3,
    tarea: 'Escribir las tres historias que explican por qué haces esto.' },
  { semana: 6,  mes: 2, tema: 'Ecosistema de confianza',  horas: 3,
    tarea: 'Ordenar dónde te encuentran y qué ven primero.' },
  { semana: 7,  mes: 2, tema: 'Sistemas de conversión',   horas: 3,
    tarea: 'Lanzar tu sistema de adquisición.' },
  { semana: 8,  mes: 2, tema: 'Ventas desde contenido',   horas: 3,
    tarea: 'Convertir dos piezas que ya funcionaron en piezas de venta.' },
  { semana: 9,  mes: 3, tema: 'Diseño de oferta',         horas: 3,
    tarea: 'Definir qué vendes, a quién y por cuánto.' },
  { semana: 10, mes: 3, tema: 'Creación de producto',     horas: 4,
    tarea: 'Armar la primera versión de lo que vas a vender.' },
  { semana: 11, mes: 3, tema: 'Infraestructura',          horas: 2,
    tarea: 'Cobro, entrega y correo: que funcione sin ti.' },
  { semana: 12, mes: 3, tema: 'Landing pages',            horas: 2,
    tarea: 'Una página que venda sola.' },
];

// ─── PUERTA ──────────────────────────────────────────────────

/**
 * Soul es de ella. Punto.
 *
 * Una operadora crea clientes y los acompaña; no tiene por qué ver el
 * ciclo, las comidas ni los encargos de PHH de nadie. Y si algún día
 * hay una segunda socia, cada una ve lo suyo: las filas van con su
 * correo y se filtran por él, no por quién tiene la llave.
 */
function soulPuede_(s) {
  return s && s.rol === 'socia';
}

function soulUsuario_(s) {
  return String((s && s.correo) || '').toLowerCase().trim();
}

// ─── HOJAS ───────────────────────────────────────────────────

function soulSheet_(nombre) {
  const ss = SpreadsheetApp.openById(IDS_().soul);
  const sh = ss.getSheetByName(nombre);
  if (!sh) {
    throw new Error('Falta la hoja ' + nombre + ' en Nova_Soul. ' +
                    'Corre bootstrapTodo() una vez.');
  }
  return sh;
}

/** Lee una hoja de Soul, ya filtrada por quién es. */
function soulLeer_(nombre, uid) {
  const sh = soulSheet_(nombre);
  if (sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const enc = d[0].map(norm);
  const cU = enc.indexOf('usuario_id');
  return d.slice(1).map(function (f) {
    const o = {};
    enc.forEach(function (c, i) { o[c] = f[i]; });
    return o;
  }).filter(function (o) {
    if (cU === -1) return true;
    const suyo = String(o.usuario_id || '').toLowerCase().trim();
    // Una fila sin dueño es de quien la está mirando: son las que
    // quedaron de antes de que existiera la columna.
    return !suyo || suyo === uid;
  });
}

/**
 * Escribe o crea una fila por id, estampando de quién es.
 *
 * El usuario_id lo pone el servidor, nunca la pantalla. Si lo mandara
 * la pantalla, bastaría cambiarlo en la consola del navegador para
 * escribir en las filas de otra persona.
 */
function soulGuardar_(nombre, datos, uid) {
  const cols = SOUL_HOJAS[nombre];
  if (!cols) throw new Error('No se puede escribir en ' + nombre + '.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Hay otro cambio guardándose. Un segundo.');
  try {
    const sh = soulSheet_(nombre);
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const id = String(datos.id || '').trim();
    datos.usuario_id = uid;

    if (!id) {
      datos.id = 's' + Utilities.getUuid().slice(0, 8);
      sh.appendRow(enc.map(function (c) {
        return datos[c] !== undefined ? datos[c] : '';
      }));
      return datos;
    }

    const d = sh.getDataRange().getValues();
    const cU = enc.indexOf('usuario_id');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim() !== id) continue;
      const duenoFila = cU === -1 ? '' : String(d[i][cU] || '').toLowerCase().trim();
      if (duenoFila && duenoFila !== uid) throw new Error('Esa fila no es tuya.');
      enc.forEach(function (c, j) {
        if (datos[c] !== undefined) d[i][j] = datos[c];
      });
      sh.getRange(i + 1, 1, 1, enc.length).setValues([d[i]]);
      const o = {};
      enc.forEach(function (c, j) { o[c] = d[i][j]; });
      return o;
    }
    throw new Error('No existe ' + nombre + ' con id ' + id + '.');
  } finally {
    lock.releaseLock();
  }
}

function soulBorrar_(nombre, id, uid) {
  if (!SOUL_HOJAS[nombre]) throw new Error('No se puede borrar en ' + nombre + '.');
  const sh = soulSheet_(nombre);
  const d = sh.getDataRange().getValues();
  const enc = d[0].map(norm);
  const cU = enc.indexOf('usuario_id');
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() !== String(id).trim()) continue;
    const dueno = cU === -1 ? '' : String(d[i][cU] || '').toLowerCase().trim();
    if (dueno && dueno !== uid) throw new Error('Esa fila no es tuya.');
    sh.deleteRow(i + 1);
    return true;
  }
  return false;
}

// ─── LOS PROYECTOS, QUE VIVEN EN CENTRAL ─────────────────────

/**
 * La lista de proyectos, leída de Nova_Central.
 *
 * Sin plata a propósito: Soul reparte horas, y para repartir horas el
 * valor del contrato no aporta nada. Lo que sí baja es `horas_semana`,
 * que es el peso fijo del compromiso.
 */
function soulTrabajos_() {
  try {
    const filas = mioLeer_('Trabajos');
    return filas.map(function (t) {
      return {
        id: String(t.id || ''),
        nombre: String(t.nombre || ''),
        tipo: norm(t.tipo) || 'cliente',
        estado: norm(t.estado) || 'activo',
        horasSemana: num(t.horas_semana),
        entrega: aISO(t.fecha_entrega, 'UTC') || '',
      };
    }).filter(function (t) { return t.id && t.nombre; });
  } catch (e) {
    // Central puede no estar lista todavía. Soul sigue funcionando sin
    // proyectos: las tareas sueltas no dependen de ellos.
    return [];
  }
}

/**
 * Lo que Soul le cuenta a Central de cada proyecto: CUÁNTO, nunca QUÉ.
 *
 * Esta función es la frontera hecha código. Devuelve conteos y horas.
 * No devuelve el texto de una sola tarea, y no debe devolverlo nunca.
 */
function soulCargaPorTrabajo_(uid) {
  const out = {};
  try {
    const hoy = ahoraISO().slice(0, 10);
    soulLeer_('Pendientes', uid).forEach(function (f) {
      const t = String(f.trabajo_id || '').trim();
      if (!t) return;
      if (soulHecho_(f)) return;
      if (!out[t]) out[t] = { abiertas: 0, horas: 0, vencidas: 0, proxima: '' };
      out[t].abiertas++;
      out[t].horas += num(f.horas_estimadas);
      const fecha = aISO(f.fecha, 'UTC') || '';
      if (fecha && fecha < hoy) out[t].vencidas++;
      if (fecha && (!out[t].proxima || fecha < out[t].proxima)) out[t].proxima = fecha;
    });
  } catch (e) { /* si Soul no está, Central no se cae por eso */ }
  return out;
}

// ─── UN PENDIENTE ────────────────────────────────────────────

function soulHecho_(f) {
  return norm(f.estado) === 'hecho' || String(f.hecho).toLowerCase() === 'true' ||
         norm(f.hecho) === 'si' || !!f.hecho_en;
}

function soulPendiente_(f, hoy, trabajos) {
  const fecha = aISO(f.fecha, 'UTC') || '';
  const hecho = soulHecho_(f);
  const t = trabajos.filter(function (x) { return x.id === String(f.trabajo_id || ''); })[0];
  return {
    id: String(f.id || ''),
    texto: String(f.texto || ''),
    trabajoId: String(f.trabajo_id || ''),
    proyecto: t ? t.nombre : '',
    tipo: t ? t.tipo : (norm(f.tipo) || ''),
    origen: norm(f.origen) || '',
    fecha: fecha,
    // Días que faltan (negativo) o que ya pasaron (positivo).
    dias: fecha ? Math.round((new Date(hoy + 'T00:00:00Z') -
                              new Date(fecha + 'T00:00:00Z')) / 86400000) : null,
    estado: hecho ? 'hecho' : (SOUL_ESTADOS.indexOf(norm(f.estado)) !== -1
                               ? norm(f.estado) : 'pendiente'),
    prioridad: SOUL_PRIORIDADES[norm(f.prioridad)] ? norm(f.prioridad) : 'media',
    riesgo: SOUL_RIESGOS[norm(f.riesgo)] ? norm(f.riesgo) : 'corrible',
    horas: num(f.horas_estimadas),
    horasReales: num(f.horas_reales),
    hechoEn: aISO(f.hecho_en, 'UTC') || '',
    nota: String(f.nota || ''),
  };
}

/**
 * Qué tan urgente es, en un número.
 *
 * Existe para poder ORDENAR, que es distinto de juzgar: lo vencido pesa
 * más que lo de hoy, lo de hoy más que lo de la semana entrante, y a
 * igualdad de fecha pesa más lo que no se puede mover. Una tarea sin
 * fecha nunca gana: no se puede decir que algo está tarde si nadie dijo
 * cuándo.
 */
function soulUrgencia_(p) {
  let base;
  if (p.dias === null) base = 60;
  else if (p.dias > 0) base = 900 + Math.min(p.dias, 60) * 2;
  else if (p.dias === 0) base = 800;
  else base = Math.max(100, 700 + p.dias * 12);
  base += (SOUL_PRIORIDADES[p.prioridad] || 2) * 15;
  base += (SOUL_RIESGOS[p.riesgo] || { orden: 1 }).orden * 20;
  return base;
}

// ─── LAS HORAS LIBRES ────────────────────────────────────────

/**
 * Cuántas horas libres tiene cada día, de lunes a domingo.
 *
 * Devuelve `null` cuando no hay ni una fila: es el techo de toda la
 * pantalla y no se puede suponer. Una semana de 24 horas por día daría
 * siempre verde, que es peor que no decir nada.
 */
function soulHorasLibres_(uid) {
  const sh = soulSheet_('Horas');
  if (sh.getLastRow() < 2) return null;
  const d = sh.getDataRange().getValues();
  const enc = d[0].map(norm);
  const cU = enc.indexOf('usuario_id'), cD = enc.indexOf('dia_semana'),
        cH = enc.indexOf('horas_libres');
  const out = {}; let hubo = false;
  for (let i = 1; i < d.length; i++) {
    const suyo = String(d[i][cU] || '').toLowerCase().trim();
    if (suyo && suyo !== uid) continue;
    const dia = num(d[i][cD]);
    if (dia < 1 || dia > 7) continue;
    out[dia] = num(d[i][cH]);
    hubo = true;
  }
  return hubo ? out : null;
}

function soulHorasGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const dias = p.dias || {};

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro cambio guardándose.' };
  try {
    const sh = soulSheet_('Horas');
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const cU = enc.indexOf('usuario_id'), cD = enc.indexOf('dia_semana'),
          cH = enc.indexOf('horas_libres');
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];

    Object.keys(dias).forEach(function (k) {
      const dia = Number(k);
      if (!(dia >= 1 && dia <= 7)) return;
      const horas = num(dias[k]);
      let fila = -1;
      for (let i = 1; i < d.length; i++) {
        const suyo = String(d[i][cU] || '').toLowerCase().trim();
        if ((!suyo || suyo === uid) && num(d[i][cD]) === dia) { fila = i; break; }
      }
      if (fila === -1) {
        const nueva = enc.map(function () { return ''; });
        nueva[cU] = uid; nueva[cD] = dia; nueva[cH] = horas;
        sh.appendRow(nueva);
        d.push(nueva);
      } else {
        d[fila][cU] = uid; d[fila][cH] = horas;
        sh.getRange(fila + 1, 1, 1, enc.length).setValues([d[fila]]);
      }
    });
    return { ok: true, horas: soulHorasLibres_(uid) };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally {
    lock.releaseLock();
  }
}

// ─── MINDLAB ─────────────────────────────────────────────────

/**
 * El plan con sus fechas y lo que ella haya marcado encima.
 *
 * Las doce semanas NO se escriben en la hoja al abrirla. El plan es una
 * constante con fechas calculadas; la hoja guarda solo lo que cambió —
 * el estado de cada semana. Así el plan se puede corregir en una línea
 * sin dejar doce filas viejas contando otra historia.
 */
function soulMindlabPlan_(uid) {
  let marcas = {};
  try {
    soulLeer_('Mindlab', uid).forEach(function (f) {
      marcas[String(f.semana)] = f;
    });
  } catch (e) { marcas = {}; }

  return MINDLAB_PLAN.map(function (m) {
    const desde = masDias_(MINDLAB_INICIO, (m.semana - 1) * 7);
    const marca = marcas[String(m.semana)] || {};
    return {
      id: 'ml' + m.semana,
      semana: m.semana, mes: m.mes, tema: m.tema, tarea: m.tarea,
      horas: m.horas,
      desde: desde, hasta: masDias_(desde, 6),
      estado: SOUL_ESTADOS.indexOf(norm(marca.estado)) !== -1
              ? norm(marca.estado) : 'pendiente',
      nota: String(marca.nota || ''),
    };
  });
}

function soulMindlabGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const semana = Number(p.semana);
  const base = MINDLAB_PLAN.filter(function (m) { return m.semana === semana; })[0];
  if (!base) return { ok: false, error: 'Esa semana no está en el plan.' };

  try {
    const existe = soulLeer_('Mindlab', uid)
      .filter(function (f) { return Number(f.semana) === semana; })[0];
    const desde = masDias_(MINDLAB_INICIO, (semana - 1) * 7);
    soulGuardar_('Mindlab', {
      id: existe ? existe.id : '',
      semana: semana, mes: base.mes, tema: base.tema, tarea: base.tarea,
      horas_estimadas: base.horas, desde: desde, hasta: masDias_(desde, 6),
      estado: SOUL_ESTADOS.indexOf(norm(p.estado)) !== -1 ? norm(p.estado) : 'pendiente',
      nota: p.nota !== undefined ? String(p.nota) : undefined,
    }, uid);
    return { ok: true, plan: soulMindlabPlan_(uid) };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/** Bajar la semana de Mindlab a los pendientes, como una tarea más. */
function soulMindlabAPendientes(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const semana = Number(p.semana);
  const m = soulMindlabPlan_(uid).filter(function (x) { return x.semana === semana; })[0];
  if (!m) return { ok: false, error: 'Esa semana no está en el plan.' };

  const yaEsta = soulLeer_('Pendientes', uid).filter(function (f) {
    return norm(f.origen) === 'mindlab' && String(f.plataforma_id) === 'ml' + semana;
  })[0];
  if (yaEsta) return { ok: false, error: 'Esa semana ya está en tus pendientes.' };

  try {
    soulGuardar_('Pendientes', {
      texto: 'Mindlab ' + semana + ' · ' + m.tarea,
      origen: 'mindlab', plataforma_id: 'ml' + semana,
      fecha: m.hasta, estado: 'pendiente', prioridad: 'media',
      horas_estimadas: m.horas,
      // Es lo único suyo y no tiene cliente que reclame: por eso corrible.
      // Marcarlo inamovible sería mentirle a la cuenta de la semana.
      riesgo: 'corrible',
      nota: m.tema,
    }, uid);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── LA FOTO DEL DÍA, LA SEMANA Y EL MES ─────────────────────

/**
 * Todo lo que pinta NovaSoul, en una sola llamada.
 *
 * Una sola a propósito, por lo mismo que en Central: pedir el día, la
 * semana, el mes y los proyectos por separado son cuatro viajes para
 * dibujar una pantalla.
 */
function soulHoy(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);

  const hoy = ahoraISO().slice(0, 10);
  const lunes = lunesDe_(hoy);
  const domingo = masDias_(lunes, 6);
  const mes = hoy.slice(0, 7);

  const trabajos = soulTrabajos_();
  const crudos = soulLeer_('Pendientes', uid);
  const todos = crudos.map(function (f) { return soulPendiente_(f, hoy, trabajos); });

  /**
   * Lo hecho se recorta a catorce días.
   *
   * La columna «Hecho» del tablero es para mirar la semana, no el
   * archivo: sin el corte, en tres meses la pantalla carga cientos de
   * filas que nadie lee para dibujar tres que sí.
   */
  const corte = masDias_(hoy, -14);
  const pendientes = todos.filter(function (t) {
    return t.estado !== 'hecho' || !t.hechoEn || t.hechoEn >= corte;
  });
  const abiertos = todos.filter(function (t) { return t.estado !== 'hecho'; });

  const enSemana = function (t) { return t.fecha && t.fecha >= lunes && t.fecha <= domingo; };
  const vencidas = abiertos.filter(function (t) { return t.dias !== null && t.dias > 0; });
  const deHoy = abiertos.filter(function (t) { return t.fecha === hoy; });

  // ── El riesgo de la semana ──
  const libresPorDia = soulHorasLibres_(uid);
  const libres = libresPorDia
    ? [1, 2, 3, 4, 5, 6, 7].reduce(function (a, d) { return a + (libresPorDia[d] || 0); }, 0)
    : null;

  const activos = trabajos.filter(function (t) { return t.estado === 'activo'; });
  const fijas = activos.reduce(function (a, t) { return a + t.horasSemana; }, 0);

  /**
   * Las horas de una entrega NO se suman si su proyecto ya tiene horas
   * fijas: las doce horas semanales de PHH ya incluyen la tarea de PHH.
   * Sumar las dos contaría lo mismo dos veces y daría una semana
   * imposible que no existe — y una alarma falsa se apaga sola en una
   * semana, para siempre.
   */
  const conFijas = {};
  activos.forEach(function (t) { if (t.horasSemana > 0) conFijas[t.id] = 1; });
  const candidatasSemana = vencidas.concat(abiertos.filter(enSemana));
  const extra = candidatasSemana.reduce(function (a, t) {
    return a + (conFijas[t.trabajoId] ? 0 : t.horas);
  }, 0);
  const dentroDeFijas = candidatasSemana.filter(function (t) { return !!conFijas[t.trabajoId]; }).length;

  const comprometidas = fijas + extra;
  const sobra = libres === null ? null : libres - comprometidas;

  /**
   * Si no cabe, cuáles son las candidatas a caerse.
   *
   * De la menos urgente hacia arriba, hasta cubrir lo que sobra de más,
   * y sin tocar lo que no se puede mover: correr un parcial no es una
   * opción, así que ofrecerlo como opción es ruido.
   */
  const candidatas = [];
  let noAlcanza = 0;
  if (sobra !== null && sobra < 0) {
    let falta = -sobra;
    candidatasSemana
      .filter(function (t) {
        /**
         * Solo puede caerse lo que de verdad SUMÓ.
         *
         * Correr el encargo de PHH no libera ni una hora: sus horas
         * nunca se sumaron, porque ya estaban dentro de las doce fijas
         * del proyecto. Ofrecerlo como candidata sería proponerle un
         * sacrificio que no arregla nada, y peor: le haría creer que la
         * semana ya cabe.
         */
        return t.riesgo !== 'inamovible' && t.horas > 0 && !conFijas[t.trabajoId];
      })
      .sort(function (a, b) { return soulUrgencia_(a) - soulUrgencia_(b); })
      .forEach(function (t) {
        if (falta <= 0) return;
        candidatas.push({ id: t.id, texto: t.texto, horas: t.horas,
                          proyecto: t.proyecto, riesgo: t.riesgo, fecha: t.fecha });
        falta -= t.horas;
      });
    noAlcanza = Math.max(0, Math.round(falta * 10) / 10);
  }

  // ── Lo más urgente, que es UNA sola cosa ──
  const urgente = abiertos.slice()
    .sort(function (a, b) { return soulUrgencia_(b) - soulUrgencia_(a); })[0] || null;

  // ── La semana, día por día ──
  const dias = [];
  for (let i = 0; i < 7; i++) {
    const f = masDias_(lunes, i);
    const delDia = abiertos.filter(function (t) { return t.fecha === f; });
    dias.push({
      fecha: f, dow: i + 1, esHoy: f === hoy,
      libres: libresPorDia ? (libresPorDia[i + 1] || 0) : null,
      entregas: delDia.length,
      horas: delDia.reduce(function (a, t) { return a + t.horas; }, 0),
      textos: delDia.sort(function (a, b) { return soulUrgencia_(b) - soulUrgencia_(a); })
                    .slice(0, 4).map(function (t) {
        return { id: t.id, texto: t.texto, proyecto: t.proyecto, horas: t.horas };
      }),
    });
  }

  // ── El mes ──
  const delMes = todos.filter(function (t) {
    return (t.fecha && t.fecha.slice(0, 7) === mes) ||
           (t.hechoEn && t.hechoEn.slice(0, 7) === mes);
  });
  const porProyecto = {};
  abiertos.forEach(function (t) {
    const k = t.trabajoId || '·sueltas';
    if (!porProyecto[k]) {
      porProyecto[k] = { id: k, nombre: t.proyecto || 'Sin proyecto',
                         abiertas: 0, horas: 0, vencidas: 0 };
    }
    porProyecto[k].abiertas++;
    porProyecto[k].horas += t.horas;
    if (t.dias !== null && t.dias > 0) porProyecto[k].vencidas++;
  });

  const ml = soulMindlabPlan_(uid);
  const mlSemana = ml.filter(function (m) { return hoy >= m.desde && hoy <= m.hasta; })[0] || null;

  return {
    ok: true,
    hoy: hoy, mes: mes,
    semana: { lunes: lunes, domingo: domingo, dias: dias },
    trabajos: trabajos,
    tipos: TIPOS_TRABAJO,
    riesgos: SOUL_RIESGOS,
    pendientes: pendientes,
    urgente: urgente,
    resumen: {
      dia: {
        entregas: deHoy.length,
        horas: deHoy.reduce(function (a, t) { return a + t.horas; }, 0),
        vencidas: vencidas.length,
        hechasHoy: todos.filter(function (t) { return t.hechoEn === hoy; }).length,
        haciendo: abiertos.filter(function (t) { return t.estado === 'haciendo'; }).length,
      },
      semana: {
        entregas: abiertos.filter(enSemana).length,
        horasEntregas: abiertos.filter(enSemana)
          .reduce(function (a, t) { return a + t.horas; }, 0),
        hechas: todos.filter(function (t) {
          return t.hechoEn && t.hechoEn >= lunes && t.hechoEn <= domingo;
        }).length,
      },
      mes: {
        entregas: delMes.length,
        hechas: delMes.filter(function (t) { return t.estado === 'hecho'; }).length,
        abiertas: abiertos.filter(function (t) {
          return t.fecha && t.fecha.slice(0, 7) === mes;
        }).length,
        porProyecto: Object.keys(porProyecto).map(function (k) { return porProyecto[k]; })
          .sort(function (a, b) { return b.horas - a.horas; }),
      },
    },
    riesgo: {
      libres: libres, fijas: fijas, extra: extra,
      comprometidas: comprometidas, sobra: sobra,
      dentroDeFijas: dentroDeFijas,
      candidatas: candidatas,
      noAlcanza: noAlcanza,
      // La pantalla no adivina por qué falta: se lo decimos.
      sinHoras: libres === null,
      porque: libres === null
        ? 'Todavía no me has dicho cuántas horas libres tienes cada día. ' +
          'Sin ese número no puedo decirte si la semana cabe: lo demás sería un adorno.'
        : noAlcanza > 0
          ? 'Aunque corras todo lo que se puede correr, siguen faltando ' +
            noAlcanza + ' h. Lo que sobra viene de las horas fijas de tus ' +
            'proyectos, y eso no se arregla moviendo una entrega: se arregla ' +
            'hablando con alguien.'
          : '',
    },
    horas: libresPorDia,
    mindlab: {
      inicio: MINDLAB_INICIO, fin: masDias_(MINDLAB_INICIO, 12 * 7 - 1),
      plan: ml, semanaActual: mlSemana,
      hechas: ml.filter(function (m) { return m.estado === 'hecho'; }).length,
      total: ml.length,
      horasTotales: ml.reduce(function (a, m) { return a + m.horas; }, 0),
    },
  };
}

// ─── ESCRIBIR UN PENDIENTE ───────────────────────────────────

function soulPendienteGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const d = p.datos || {};

  if (!String(d.texto || '').trim() && !String(d.id || '').trim()) {
    return { ok: false, error: 'Escribe qué hay que hacer.' };
  }

  const estado = SOUL_ESTADOS.indexOf(norm(d.estado)) !== -1 ? norm(d.estado) : undefined;
  const fila = {
    id: String(d.id || ''),
    texto: d.texto !== undefined ? String(d.texto).trim() : undefined,
    trabajo_id: d.trabajo_id !== undefined ? String(d.trabajo_id) : undefined,
    fecha: d.fecha !== undefined ? String(d.fecha) : undefined,
    prioridad: SOUL_PRIORIDADES[norm(d.prioridad)] ? norm(d.prioridad) : undefined,
    riesgo: SOUL_RIESGOS[norm(d.riesgo)] ? norm(d.riesgo) : undefined,
    horas_estimadas: d.horas !== undefined ? num(d.horas) : undefined,
    nota: d.nota !== undefined ? String(d.nota) : undefined,
    origen: d.origen !== undefined ? String(d.origen) : undefined,
  };

  if (estado) {
    fila.estado = estado;
    /**
     * `hecho` y `hecho_en` se llenan aquí y no en la pantalla, y las
     * tres columnas se mueven juntas. Si una fila quedara en «hecho»
     * sin fecha, el resumen del mes no sabría en qué mes contarla.
     */
    fila.hecho = estado === 'hecho';
    fila.hecho_en = estado === 'hecho' ? ahoraISO().slice(0, 10) : '';
    if (estado === 'hecho' && d.horasReales !== undefined) {
      fila.horas_reales = num(d.horasReales);
    }
  }

  try {
    const guardada = soulGuardar_('Pendientes', fila, uid);
    return { ok: true, id: guardada.id };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulPendienteBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  try {
    const fue = soulBorrar_('Pendientes', p.id, soulUsuario_(s));
    return { ok: fue, error: fue ? '' : 'No lo encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── FINANZAS: EL PLAN CONTRA LO QUE PASÓ ────────────────────

/**
 * Los gastos fijos del mes, y cuánto salió de verdad.
 *
 * El plan vive en Soul y los movimientos en Central. Se comparan aquí y
 * no se mezclan nunca: si el presupuesto se reescribiera solo con lo
 * que se gastó, siempre cuadraría y nunca serviría para nada.
 *
 * Las monedas NO se suman entre sí. Ella vive en COP y cobra en USD;
 * un total mezclado cambiaría solo mañana con la tasa.
 */
function soulFinanzas(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const mes = String(p.mes || ahoraISO().slice(0, 7));

  const fijos = soulLeer_('Fijos', uid).map(function (f) {
    return {
      id: String(f.id || ''), categoria: norm(f.categoria) || 'varios',
      concepto: String(f.concepto || ''), monto: num(f.monto),
      moneda: String(f.moneda || 'COP').toUpperCase(),
      dia: num(f.dia_del_mes),
      activo: norm(f.activo) !== 'no',
      nota: String(f.nota || ''),
    };
  });

  // Lo que de verdad salió, de la hoja Finanzas de Central
  const real = {}; let errorReal = '';
  try {
    mioLeer_('Finanzas').forEach(function (f) {
      if (String(aISO(f.fecha, 'UTC') || '').slice(0, 7) !== mes) return;
      const cat = norm(f.categoria) || 'varios';
      const mon = String(f.moneda || 'COP').toUpperCase();
      const k = cat + '|' + mon;
      if (!real[k]) real[k] = { categoria: cat, moneda: mon, monto: 0, n: 0 };
      real[k].monto += Math.abs(num(f.monto));
      real[k].n++;
    });
  } catch (e) {
    errorReal = 'No pude leer los movimientos de Nova Central: ' + e.message;
  }

  const categorias = SOUL_CATEGORIAS.map(function (c) {
    const suyos = fijos.filter(function (f) { return f.categoria === c.id && f.activo; });
    const monedas = {};
    suyos.forEach(function (f) {
      monedas[f.moneda] = (monedas[f.moneda] || 0) + f.monto;
    });
    const gastado = {};
    Object.keys(real).forEach(function (k) {
      if (real[k].categoria !== c.id) return;
      gastado[real[k].moneda] = (gastado[real[k].moneda] || 0) + real[k].monto;
    });
    return {
      id: c.id, nombre: c.nombre, tipo: c.tipo,
      lineas: suyos,
      // Sin plan no hay comparación: se dice «sin definir», no cero.
      planeado: Object.keys(monedas).length ? monedas : null,
      gastado: Object.keys(gastado).length ? gastado : null,
    };
  });

  const planTotal = {};
  fijos.filter(function (f) { return f.activo && f.categoria !== 'ahorro'; })
    .forEach(function (f) { planTotal[f.moneda] = (planTotal[f.moneda] || 0) + f.monto; });
  const ahorroTotal = {};
  fijos.filter(function (f) { return f.activo && f.categoria === 'ahorro'; })
    .forEach(function (f) { ahorroTotal[f.moneda] = (ahorroTotal[f.moneda] || 0) + f.monto; });

  return {
    ok: true, mes: mes, hoy: ahoraISO().slice(0, 10),
    categorias: categorias,
    totales: { plan: planTotal, ahorro: ahorroTotal },
    error: errorReal,
  };
}

function soulFijoGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const d = p.datos || {};
  const cat = norm(d.categoria);
  /**
   * Al editar, la categoría puede no venir: cambiar el monto del
   * arriendo no es decir otra vez que es arriendo. Pero si viene, tiene
   * que ser una de las nueve — una inventada crearía una fila que
   * ninguna pantalla vuelve a mostrar.
   */
  if (cat && !SOUL_CATEGORIAS.filter(function (c) { return c.id === cat; })[0]) {
    return { ok: false, error: 'Esa categoría no existe.' };
  }
  const esNuevo = !String(d.id || '').trim();
  if (!cat && esNuevo) {
    return { ok: false, error: 'Falta decir de qué categoría es.' };
  }
  try {
    soulGuardar_('Fijos', {
      id: String(d.id || ''),
      categoria: cat || undefined,
      concepto: d.concepto !== undefined ? String(d.concepto).trim() : undefined,
      monto: d.monto !== undefined ? num(d.monto) : undefined,
      /**
       * En una edición, lo que no venga NO se toca. Poner un valor por
       * defecto aquí haría que cambiar el monto de una línea apagada la
       * volviera a encender sola, y nadie entendería por qué.
       */
      moneda: d.moneda !== undefined ? String(d.moneda).toUpperCase()
              : (esNuevo ? 'COP' : undefined),
      dia_del_mes: d.dia !== undefined ? num(d.dia) : undefined,
      activo: d.activo !== undefined ? (d.activo === false || norm(d.activo) === 'no' ? 'no' : 'si')
              : (esNuevo ? 'si' : undefined),
      nota: d.nota !== undefined ? String(d.nota) : undefined,
    }, soulUsuario_(s));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulFijoBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  try {
    const fue = soulBorrar_('Fijos', p.id, soulUsuario_(s));
    return { ok: fue, error: fue ? '' : 'No lo encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── NOVA FAMILY, VISTA DESDE SOUL ───────────────────────────

/**
 * El resumen de las otras tres pantallas.
 *
 * Va aparte de `soulHoy` porque abre hojas de cliente y calcula el
 * semáforo: son segundos, no milisegundos, y el día a día no puede
 * esperarlos. Se carga cuando ella entra a esta sección.
 *
 * UNA TIENDA A LA VEZ. Es la regla de toda Nova y aquí también: dos
 * tiendas en la misma pantalla invitan a compararlas, y son negocios
 * distintos en países distintos con monedas distintas.
 */
function soulFamily(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);

  const out = {
    ok: true,
    clientes: [], cliente: null, tiendas: [], tienda: '',
    alarmas: [], semaforo: null, errorTienda: '',
    central: null, academy: null,
  };

  // ── Empresarial ──
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
    if (sh && sh.getLastRow() > 1) {
      const d = sh.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const empresa = String(d[i][c('empresa')] || '').trim();
        const sid = String(d[i][c('sheet_id')] || '').trim();
        if (!empresa || !sid) continue;
        if (norm(d[i][c('estado')]) === 'inactivo') continue;
        out.clientes.push({ id: String(d[i][c('id')] || ''), empresa: empresa, sheetId: sid });
      }
    }
  } catch (err) {
    out.errorTienda = 'No pude leer la lista de clientes: ' + err.message;
  }

  const elegido = String(p.cliente || '');
  const cl = out.clientes.filter(function (x) { return x.id === elegido; })[0] ||
             out.clientes[0] || null;
  out.cliente = cl;

  if (cl) {
    try {
      const cs = SpreadsheetApp.openById(cl.sheetId);
      out.tiendas = tiendasActivas_(cs).map(function (t) {
        return { id: t, nombre: nombreTienda(cs, t) };
      });
      const tienda = out.tiendas.filter(function (t) { return t.id === String(p.tienda || ''); })[0] ||
                     out.tiendas[0] || null;
      if (tienda) {
        out.tienda = tienda.id;
        const ev = evaluarAlarmas(cs, tienda.id);
        out.alarmas = (ev.alarmas || []).map(function (a) {
          return { id: a.id, nivel: a.nivel, nombre: a.nombre,
                   titulo: a.titulo, detalle: a.detalle, casos: (a.casos || []).length };
        });
        /**
         * El mismo semáforo del lunes, sin volver a calcularlo de otra
         * forma. Si fueran dos cuentas distintas, el correo y esta
         * pantalla podrían decirle cosas distintas el mismo día y no
         * habría manera de saber cuál tiene razón.
         */
        const sem = semaforoSemanal(cl.sheetId, tienda.id, '');
        out.semaforo = {
          tienda: sem.tienda, moneda: sem.moneda, semana: sem.semana,
          luces: sem.luces, alertas: sem.alertas, hayDatos: sem.hayDatos,
        };
      } else {
        out.errorTienda = 'Ese cliente no tiene tiendas activas.';
      }
    } catch (err) {
      out.errorTienda = 'No pude leer la tienda: ' + err.message;
    }
  }

  // ── Central: lo suyo, contado ──
  try {
    const hoy = ahoraISO().slice(0, 10);
    const trabajos = mioLeer_('Trabajos');
    const cobros = mioLeer_('Cobros');
    const atrasados = cobros.filter(function (co) {
      if (norm(co.estado) === 'cobrado' || co.fecha_cobrada) return false;
      const f = aISO(co.fecha_esperada, 'UTC') || '';
      return f && f < hoy;
    });
    const porMoneda = {};
    atrasados.forEach(function (co) {
      const m = String(co.moneda || '?').toUpperCase();
      porMoneda[m] = (porMoneda[m] || 0) + num(co.monto);
    });
    out.central = {
      proyectos: trabajos.filter(function (t) { return norm(t.estado) !== 'cerrado'; }).length,
      entregasVencidas: trabajos.filter(function (t) {
        const f = aISO(t.fecha_entrega, 'UTC') || '';
        return f && f < hoy && norm(t.estado) !== 'cerrado';
      }).length,
      cobrosAtrasados: atrasados.length,
      atrasadoPorMoneda: porMoneda,
      carga: soulCargaPorTrabajo_(uid),
    };
  } catch (err) {
    out.central = { error: 'No pude leer Nova Central: ' + err.message };
  }

  // ── Academy ──
  // No hay estudiantes todavía. Decirlo es más útil que un cero que
  // parece un dato.
  try {
    const sh = SpreadsheetApp.openById(IDS_().academy).getSheetByName('Estudiantes');
    const n = sh && sh.getLastRow() > 1 ? sh.getLastRow() - 1 : 0;
    out.academy = { estudiantes: n,
      porque: n ? '' : 'Todavía no has dado de alta a nadie en novAcademy.' };
  } catch (err) {
    out.academy = { estudiantes: 0, porque: 'novAcademy todavía no está montada.' };
  }

  return out;
}
