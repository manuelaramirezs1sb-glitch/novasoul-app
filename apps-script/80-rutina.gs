/**
 * ═══════════════════════════════════════════════════════════
 *  NOVASOUL · LO QUE SE REPITE, Y LA PLATA CHIQUITA
 * ═══════════════════════════════════════════════════════════
 *
 * Tres cosas que pidió y que están conectadas entre sí:
 *
 * 1. LOS TURNOS Y LAS CLASES SE ESCRIBEN UNA VEZ. Se repiten solos
 *    todas las semanas hasta la fecha en que se acaban.
 *
 * 2. EL TURNO VALE LO MISMO SIEMPRE, LAS PROPINAS NO. La paga es fija y
 *    vive en la rutina; las propinas se escriben al día siguiente, y
 *    hasta que no se escriban el turno aparece pendiente.
 *
 * 3. LOS GASTOS HORMIGA. Buses, Uber, antojos, salidas. Sueltos y
 *    chiquitos, pero a fin de mes son la diferencia entre que sobre y
 *    que falte.
 *
 * ┌─ POR QUÉ ESTO CAMBIA EL CÁLCULO DE LA SEMANA ──────────────┐
 * │                                                            │
 * │ Antes ella escribía a mano cuántas horas libres tenía cada │
 * │ día, ya descontados turnos y clases. Ahora los turnos y    │
 * │ las clases los sabe Nova, así que el número que escribe es │
 * │ otro: las horas ÚTILES del día, antes de descontar nada.   │
 * │                                                            │
 * │   libres = útiles − (turnos + clases de ese día)           │
 * │                                                            │
 * │ El cambio de significado no puede ser silencioso: la       │
 * │ pantalla muestra la resta día por día mientras ella        │
 * │ escribe, para que vea el número nuevo en el momento.       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * UNA SOLA CONTABILIDAD. La paga del turno, las propinas y los gastos
 * hormiga se escriben en la hoja Finanzas de Nova_Central, que es donde
 * ya viven sus movimientos. Si Soul llevara su propia caja, a fin de
 * mes habría dos respuestas a «¿cuánto me queda?» y ninguna forma de
 * saber cuál es la buena.
 */

const RUTINA_TIPOS = {
  turno:  { nombre: 'Turno',  paga: true,  ayuda: 'Salsabor, o cualquier trabajo por turnos.' },
  clase:  { nombre: 'Clase',  paga: false, ayuda: 'Una materia, a su hora, todas las semanas.' },
  otro:   { nombre: 'Otro',   paga: false, ayuda: 'Gimnasio, terapia, lo que ocupe tiempo fijo.' },
};

/**
 * Los gastos hormiga, con los nombres que ella usó.
 *
 * Están aparte de los fijos a propósito: un arriendo se planea y un
 * antojo se registra. Mezclarlos haría que el presupuesto pareciera
 * cumplido justo los meses en que no lo fue.
 */
const SOUL_HORMIGA = [
  { id: 'bus',        nombre: 'Buses' },
  { id: 'transporte', nombre: 'Transporte' },
  { id: 'uber',       nombre: 'Uber' },
  { id: 'antojos',    nombre: 'Antojos' },
  { id: 'salidas',    nombre: 'Salidas' },
  { id: 'hormiga',    nombre: 'Otros sueltos' },
];

/** Cuántos días para atrás se buscan turnos sin propinas. */
const TURNOS_DIAS_ATRAS = 21;

// ─── HORAS ───────────────────────────────────────────────────

/** "18:30" → 18.5. Lo que no se entienda vale 0, no adivina. */
function horaNum_(v) {
  const s = String(v == null ? '' : v).trim();
  if (!s) return 0;
  // Una celda con formato de hora llega como Date
  if (v instanceof Date) return v.getHours() + v.getMinutes() / 60;
  const m = s.match(/^(\d{1,2})(?:[:.](\d{1,2}))?/);
  if (!m) return 0;
  const h = parseInt(m[1], 10), min = m[2] ? parseInt(m[2], 10) : 0;
  if (h > 23 || min > 59) return 0;
  return h + min / 60;
}

/**
 * Cuánto dura un bloque, contando los que cruzan la medianoche.
 *
 * Un turno de 18:00 a 02:00 dura ocho horas, no menos dieciséis. Restar
 * a secas daba negativo y la semana salía con horas de sobra
 * justamente los días en que ella trabaja hasta la madrugada.
 */
function duracionHoras_(inicio, fin) {
  const a = horaNum_(inicio), b = horaNum_(fin);
  if (!a && !b) return 0;
  const d = b - a;
  return d > 0 ? d : (d + 24);
}

function hhmm_(n) {
  const h = Math.floor(n), m = Math.round((n - h) * 60);
  return dosDig_(h) + ':' + dosDig_(m);
}

// ─── LA RUTINA ───────────────────────────────────────────────

function rutinaDe_(uid, faltan) {
  return soulLeerSuave_('Rutina', uid, faltan)
    .filter(function (f) { return String(f.id || '').trim(); })
    .map(function (f) {
      return {
        id: String(f.id), tipo: RUTINA_TIPOS[norm(f.tipo)] ? norm(f.tipo) : 'otro',
        nombre: String(f.nombre || ''),
        dia: num(f.dia_semana),
        inicio: typeof f.hora_inicio === 'string' ? f.hora_inicio.trim() : hhmm_(horaNum_(f.hora_inicio)),
        fin: typeof f.hora_fin === 'string' ? f.hora_fin.trim() : hhmm_(horaNum_(f.hora_fin)),
        horas: duracionHoras_(f.hora_inicio, f.hora_fin),
        lugar: String(f.lugar || ''),
        trabajoId: String(f.trabajo_id || ''),
        materiaId: String(f.materia_id || ''),
        paga: num(f.paga_fija),
        moneda: String(f.moneda || 'COP').toUpperCase(),
        desde: aISO(f.desde, 'UTC') || '',
        hasta: aISO(f.hasta, 'UTC') || '',
        activo: norm(f.activo) !== 'no',
      };
    })
    .filter(function (r) { return r.dia >= 1 && r.dia <= 7; });
}

/**
 * ¿Esta rutina corre en esta fecha?
 *
 * Fuera de `desde`–`hasta` no corre, y eso es lo que impide que una
 * clase del semestre pasado siga ocupando la semana para siempre.
 */
function rutinaCorre_(r, fechaISO) {
  if (!r.activo) return false;
  const dow = new Date(fechaISO + 'T00:00:00Z').getUTCDay();
  const dia = dow === 0 ? 7 : dow;
  if (r.dia !== dia) return false;
  if (r.desde && fechaISO < r.desde) return false;
  if (r.hasta && fechaISO > r.hasta) return false;
  return true;
}

/** Los bloques de un día, ordenados por hora. */
function rutinaDelDia_(rutinas, fechaISO) {
  return rutinas.filter(function (r) { return rutinaCorre_(r, fechaISO); })
    .sort(function (a, b) { return horaNum_(a.inicio) - horaNum_(b.inicio); });
}

function soulRutina(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const rutinas = rutinaDe_(uid);
  const porDia = {};
  for (let d = 1; d <= 7; d++) porDia[d] = 0;
  rutinas.forEach(function (r) { if (r.activo) porDia[r.dia] += r.horas; });

  return {
    ok: true, hoy: ahoraISO().slice(0, 10),
    tipos: RUTINA_TIPOS,
    rutinas: rutinas.sort(function (a, b) {
      return a.dia - b.dia || horaNum_(a.inicio) - horaNum_(b.inicio);
    }),
    ocupadasPorDia: porDia,
    trabajos: soulTrabajos_(),
    materias: soulLeerSuave_('Materias', uid, []).map(function (f) {
      return { id: String(f.id || ''), nombre: String(f.nombre || '') };
    }).filter(function (m) { return m.id; }),
  };
}

function soulRutinaGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const d = p.datos || {};
  const esNueva = !String(d.id || '').trim();
  const tipo = RUTINA_TIPOS[norm(d.tipo)] ? norm(d.tipo) : '';

  if (esNueva) {
    if (!String(d.nombre || '').trim()) return { ok: false, error: 'Ponle un nombre.' };
    if (!tipo) return { ok: false, error: 'Dime si es turno, clase u otra cosa.' };
    if (!(num(d.dia_semana) >= 1 && num(d.dia_semana) <= 7)) {
      return { ok: false, error: 'Falta el día de la semana.' };
    }
    if (!duracionHoras_(d.hora_inicio, d.hora_fin)) {
      return { ok: false, error: 'Sin hora de inicio y de fin no sé cuánto ocupa.' };
    }
  }
  if (d.desde && d.hasta && String(d.hasta) < String(d.desde)) {
    return { ok: false, error: 'La fecha de fin va después de la de inicio.' };
  }

  try {
    soulGuardar_('Rutina', {
      id: String(d.id || ''),
      tipo: tipo || undefined,
      nombre: d.nombre !== undefined ? String(d.nombre).trim() : undefined,
      dia_semana: d.dia_semana !== undefined ? num(d.dia_semana) : undefined,
      hora_inicio: d.hora_inicio !== undefined ? String(d.hora_inicio) : undefined,
      hora_fin: d.hora_fin !== undefined ? String(d.hora_fin) : undefined,
      lugar: d.lugar !== undefined ? String(d.lugar) : undefined,
      trabajo_id: d.trabajo_id !== undefined ? String(d.trabajo_id) : undefined,
      materia_id: d.materia_id !== undefined ? String(d.materia_id) : undefined,
      paga_fija: d.paga_fija !== undefined ? num(d.paga_fija) : undefined,
      moneda: d.moneda !== undefined ? String(d.moneda).toUpperCase() : (esNueva ? 'COP' : undefined),
      desde: d.desde !== undefined ? String(d.desde) : undefined,
      hasta: d.hasta !== undefined ? String(d.hasta) : undefined,
      activo: d.activo !== undefined ? (d.activo === false || norm(d.activo) === 'no' ? 'no' : 'si')
              : (esNueva ? 'si' : undefined),
      nota: d.nota !== undefined ? String(d.nota) : undefined,
    }, soulUsuario_(s));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulRutinaBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  /**
   * Borrar la rutina NO borra los turnos ya trabajados.
   *
   * Dejar de trabajar los viernes no deshace los viernes que ya
   * trabajó, y esa plata ya entró. Si se fueran juntos, cambiar de
   * horario le borraría ingresos del mes pasado.
   */
  const trabajados = soulLeerSuave_('Turnos', uid, []).filter(function (f) {
    return String(f.rutina_id || '') === String(p.id);
  }).length;
  try {
    const fue = soulBorrar_('Rutina', p.id, uid);
    return { ok: fue, error: fue ? '' : 'No la encontré.', trabajados: trabajados };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── LOS TURNOS TRABAJADOS ───────────────────────────────────

function turnosDe_(uid, faltan) {
  const por = {};
  soulLeerSuave_('Turnos', uid, faltan).forEach(function (f) {
    const fecha = aISO(f.fecha, 'UTC') || '';
    if (!fecha) return;
    por[String(f.rutina_id || '') + '|' + fecha] = {
      id: String(f.id || ''), rutinaId: String(f.rutina_id || ''), fecha: fecha,
      paga: num(f.paga), propinas: num(f.propinas),
      moneda: String(f.moneda || 'COP').toUpperCase(),
      estado: norm(f.estado) || 'pendiente',
      finanzaId: String(f.finanza_id || ''),
      nota: String(f.nota || ''),
    };
  });
  return por;
}

/**
 * Los turnos que ya pasaron y todavía no tienen propinas escritas.
 *
 * Solo hacia atrás: el turno de mañana no tiene propinas porque no ha
 * pasado, y pedirlas hoy sería pedir un número inventado. El de hoy
 * tampoco, hasta que termine el día.
 */
function turnosPendientes_(uid, hoyISO, rutinas, faltan) {
  const hechos = turnosDe_(uid, faltan);
  const turnos = (rutinas || rutinaDe_(uid, faltan))
    .filter(function (r) { return r.tipo === 'turno'; });
  const out = [];
  for (let i = 1; i <= TURNOS_DIAS_ATRAS; i++) {
    const f = masDias_(hoyISO, -i);
    turnosDelDia_(turnos, f, hechos).forEach(function (t) {
      if (t.escrito) return;
      out.push(t);
    });
  }
  return out.sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
}

function turnosDelDia_(turnos, fechaISO, hechos) {
  return turnos.filter(function (r) { return rutinaCorre_(r, fechaISO); })
    .map(function (r) {
      const y = hechos[r.id + '|' + fechaISO];
      return {
        rutinaId: r.id, nombre: r.nombre, lugar: r.lugar, fecha: fechaISO,
        inicio: r.inicio, fin: r.fin, horas: r.horas,
        paga: y ? y.paga : r.paga,
        pagaFija: r.paga,
        moneda: y ? y.moneda : r.moneda,
        propinas: y ? y.propinas : 0,
        // «Escrito» es haber pasado por aquí, aunque las propinas fueran
        // cero: un día malo es un dato, no un olvido.
        escrito: !!y,
        turnoId: y ? y.id : '',
      };
    });
}

/**
 * Guarda el turno con sus propinas, y lo mete en la contabilidad.
 *
 * El movimiento se escribe en Nova_Central UNA vez: si ella corrige las
 * propinas mañana, se actualiza el mismo movimiento en vez de aparecer
 * un ingreso nuevo. Por eso existe `finanza_id`.
 */
function soulTurnoGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const rutinaId = String(p.rutina || '').trim();
  const fecha = String(p.fecha || '').trim();
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return { ok: false, error: 'Falta la fecha del turno.' };
  if (fecha > ahoraISO().slice(0, 10)) {
    return { ok: false, error: 'Ese turno todavía no ha pasado.' };
  }

  const r = rutinaDe_(uid).filter(function (x) { return x.id === rutinaId; })[0];
  if (!r) return { ok: false, error: 'Ese turno no está en tu rutina.' };

  const propinas = num(p.propinas);
  const paga = p.paga !== undefined && p.paga !== '' ? num(p.paga) : r.paga;
  if (propinas < 0 || paga < 0) return { ok: false, error: 'No puede ser negativo.' };

  const yaEsta = turnosDe_(uid)[rutinaId + '|' + fecha];
  const total = paga + propinas;

  try {
    let finanzaId = yaEsta ? yaEsta.finanzaId : '';
    if (total > 0) {
      const fila = mioGuardar_('Finanzas', {
        id: finanzaId,
        fecha: fecha, flujo: 'ingreso', categoria: 'turno',
        concepto: r.nombre + (r.lugar ? ' · ' + r.lugar : ''),
        monto: total, moneda: r.moneda,
        trabajo_id: r.trabajoId,
        nota: 'Base ' + paga + ' + propinas ' + propinas,
      });
      finanzaId = String(fila.id || '');
    }
    soulGuardar_('Turnos', {
      id: yaEsta ? yaEsta.id : '',
      rutina_id: rutinaId, fecha: fecha,
      paga: paga, propinas: propinas, moneda: r.moneda,
      estado: 'cerrado', finanza_id: finanzaId,
      nota: p.nota !== undefined ? String(p.nota) : undefined,
    }, uid);
    return { ok: true, total: total, moneda: r.moneda };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── GASTOS HORMIGA ──────────────────────────────────────────

/**
 * Un gasto chiquito, escrito en dos toques.
 *
 * Va directo a Finanzas de Nova_Central con su categoría. No se guarda
 * una copia en Soul: dos copias de la misma plata terminan siempre en
 * dos respuestas distintas a fin de mes.
 */
function soulHormigaGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const cat = norm(p.categoria);
  if (!SOUL_HORMIGA.filter(function (c) { return c.id === cat; })[0]) {
    return { ok: false, error: 'Esa categoría no está en la lista.' };
  }
  const monto = num(p.monto);
  if (!monto || monto <= 0) return { ok: false, error: '¿Cuánto fue?' };

  const fecha = String(p.fecha || '').trim() || ahoraISO().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) return { ok: false, error: 'La fecha no sirve.' };

  try {
    mioGuardar_('Finanzas', {
      fecha: fecha, flujo: 'gasto', categoria: cat,
      concepto: String(p.concepto || '').trim() ||
                (SOUL_HORMIGA.filter(function (c) { return c.id === cat; })[0].nombre),
      monto: monto,
      moneda: String(p.moneda || 'COP').toUpperCase(),
      nota: 'Hormiga, desde NovaSoul',
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulHormigaBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  try {
    const fue = mioBorrar_('Finanzas', p.id);
    return { ok: fue, error: fue ? '' : 'No lo encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── EL MES: ¿SOBRA O FALTA? ─────────────────────────────────

/**
 * La cuenta que ella pidió, con sus dos mitades separadas.
 *
 * Lo que YA pasó es un hecho: lo que entró y lo que salió, leído de
 * Finanzas. Lo que FALTA del mes es una previsión: los gastos fijos que
 * todavía no se han pagado y los turnos que faltan por trabajar.
 *
 * Van separadas porque son cosas distintas. Un solo número que mezcla
 * hechos con previsiones se lee como un hecho, y decide compras.
 */
function soulMes_(uid, mes, hoyISO) {
  const rutinas = rutinaDe_(uid);
  const fijos = soulLeerSuave_('Fijos', uid, []).filter(function (f) {
    return norm(f.activo) !== 'no';
  });

  let movs = [], errorMovs = '';
  try { movs = mioLeer_('Finanzas'); }
  catch (e) { errorMovs = 'No pude leer los movimientos de Nova Central: ' + e.message; }

  const delMes = movs.filter(function (f) {
    return String(aISO(f.fecha, 'UTC') || '').slice(0, 7) === mes;
  });

  const suma = function (filas) {
    const m = {};
    filas.forEach(function (f) {
      const k = String(f.moneda || 'COP').toUpperCase();
      m[k] = (m[k] || 0) + Math.abs(num(f.monto));
    });
    return m;
  };
  const esHormiga = function (f) {
    return SOUL_HORMIGA.filter(function (c) { return c.id === norm(f.categoria); }).length > 0;
  };

  const entro = suma(delMes.filter(function (f) { return norm(f.flujo) === 'ingreso'; }));
  const gastos = delMes.filter(function (f) { return norm(f.flujo) !== 'ingreso'; });
  const salio = suma(gastos);
  const hormiga = suma(gastos.filter(esHormiga));

  // Lo hormiga por categoría, para saber cuál se está comiendo el mes
  const porCategoria = SOUL_HORMIGA.map(function (c) {
    const suyos = gastos.filter(function (f) { return norm(f.categoria) === c.id; });
    return { id: c.id, nombre: c.nombre, n: suyos.length,
             monto: Object.keys(suma(suyos)).length ? suma(suyos) : null };
  });

  // ── Lo que falta del mes ──
  const ultimo = new Date(Date.UTC(Number(mes.slice(0, 4)), Number(mes.slice(5, 7)), 0))
    .toISOString().slice(0, 10);
  const diaHoy = hoyISO.slice(0, 7) === mes ? Number(hoyISO.slice(8, 10)) : 32;

  const fijosPendientes = {};
  fijos.forEach(function (f) {
    if (norm(f.categoria) === 'ahorro') return;
    const dia = num(f.dia_del_mes);
    // Sin día del mes no se puede saber si ya se pagó: se cuenta como
    // pendiente y se dice, en vez de darlo por hecho.
    if (dia && dia < diaHoy) return;
    const k = String(f.moneda || 'COP').toUpperCase();
    fijosPendientes[k] = (fijosPendientes[k] || 0) + num(f.monto);
  });

  const turnosHechos = turnosDe_(uid);
  const porVenir = {};
  let turnosPorVenir = 0;
  if (hoyISO.slice(0, 7) === mes) {
    let f = hoyISO;
    while (f <= ultimo) {
      rutinas.filter(function (r) { return r.tipo === 'turno' && r.paga > 0; })
        .forEach(function (r) {
          if (!rutinaCorre_(r, f)) return;
          if (turnosHechos[r.id + '|' + f]) return;   // ya está contado en «entró»
          porVenir[r.moneda] = (porVenir[r.moneda] || 0) + r.paga;
          turnosPorVenir++;
        });
      f = masDias_(f, 1);
    }
  }

  /**
   * El resultado, moneda por moneda. Nunca sumadas entre sí: ella cobra
   * en dólares y vive en pesos, y un total mezclado cambia solo mañana.
   */
  const monedas = {};
  [entro, salio, fijosPendientes, porVenir].forEach(function (o) {
    Object.keys(o).forEach(function (k) { monedas[k] = 1; });
  });
  const resultado = Object.keys(monedas).map(function (k) {
    const hoyQueda = (entro[k] || 0) - (salio[k] || 0);
    return {
      moneda: k,
      entro: entro[k] || 0, salio: salio[k] || 0,
      queda: hoyQueda,
      fijosPendientes: fijosPendientes[k] || 0,
      porVenir: porVenir[k] || 0,
      // La proyección a fin de mes: lo que hay, más lo que falta por
      // entrar, menos lo que falta por salir.
      proyectado: hoyQueda + (porVenir[k] || 0) - (fijosPendientes[k] || 0),
    };
  }).sort(function (a, b) { return a.moneda === 'COP' ? -1 : 1; });

  return {
    mes: mes, ultimoDia: ultimo,
    entro: entro, salio: salio, hormiga: hormiga,
    hormigaPorCategoria: porCategoria,
    resultado: resultado,
    turnosPorVenir: turnosPorVenir,
    movimientosHormiga: gastos.filter(esHormiga).map(function (f) {
      return { id: String(f.id || ''), fecha: aISO(f.fecha, 'UTC') || '',
               categoria: norm(f.categoria), concepto: String(f.concepto || ''),
               monto: num(f.monto), moneda: String(f.moneda || 'COP').toUpperCase() };
    }).sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; }),
    error: errorMovs,
  };
}

/** La pantalla de plata: el plan, lo hormiga, los turnos y el resultado. */
function soulPlata(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);
  const mes = String(p.mes || hoy.slice(0, 7));

  const base = soulFinanzas(s, { mes: mes });
  if (!base.ok) return base;

  return {
    ok: true, hoy: hoy, mes: mes,
    categorias: base.categorias, totales: base.totales, error: base.error,
    hormigaCategorias: SOUL_HORMIGA,
    turnosPendientes: turnosPendientes_(uid, hoy),
    resumen: soulMes_(uid, mes, hoy),
  };
}
