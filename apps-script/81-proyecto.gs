/**
 * ═══════════════════════════════════════════════════════════
 *  NOVA CENTRAL · UN PROYECTO A FONDO
 * ═══════════════════════════════════════════════════════════
 *
 * La lista de proyectos dice que existen. Esta pantalla dice qué son:
 * su especificación, de dónde sale su información, cómo cobra, cuánto
 * ha entrado, cuánto falta, y qué hay que entregar esta semana.
 *
 * ┌─ DÓNDE VA LA INFORMACIÓN PROFUNDA ─────────────────────────┐
 * │                                                            │
 * │ En FUENTES, y son enlaces. El Excel, el PDF, el Word, la   │
 * │ presentación o el artefacto de Claude se quedan donde      │
 * │ están; Nova guarda dónde y qué es. Misma regla que con las │
 * │ materias: Nova sabe dónde están las cosas, no guarda una   │
 * │ copia de lo que dicen.                                     │
 * │                                                            │
 * │ De lo que se PEGA como texto sí se sacan fechas, con el    │
 * │ mismo lector del sílabo: propone entregas y ella confirma. │
 * │                                                            │
 * │ Entender un PDF o una presentación por su cuenta sigue     │
 * │ necesitando un modelo de lenguaje. Un .xlsx sí se puede    │
 * │ leer hoy —Nova ya lo hace con los reportes de las          │
 * │ tiendas—, pero eso es otra cosa que leer un contrato.      │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ EL TEXTO DE LAS TAREAS PUEDE NO SUBIR ────────────┐
 * │                                                            │
 * │ Un proyecto marcado `confidencial` le manda a Central los  │
 * │ NÚMEROS —cuántas entregas, cuántas horas— y deja el texto  │
 * │ en NovaSoul. Nace encendido para los empleos, porque lo de │
 * │ PHH es confidencial y ella lo dejó dicho. Un olvido no     │
 * │ puede ser lo único que proteja eso.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Qué es ELLA dentro del proyecto. Distinto de qué ES el proyecto. */
const PROY_ROLES = {
  socia:       { nombre: 'Socia',       ayuda: 'Te llevas una parte de lo que produce.' },
  trabajadora: { nombre: 'Trabajadora', ayuda: 'Te pagan por hacerlo.' },
  propio:      { nombre: 'Propio',      ayuda: 'Es tuyo. No hay a quién cobrarle.' },
  estudio:     { nombre: 'Estudio',     ayuda: 'La universidad. Ocupa horas y no paga.' },
};

/** Cómo entra la plata. */
const PROY_MODALIDADES = {
  fijo:       { nombre: 'Precio fijo',   cobra: true,  ayuda: 'Un valor acordado por todo el trabajo.' },
  porcentaje: { nombre: 'Porcentaje',    cobra: true,  ayuda: 'Una parte de lo que produce una tienda.' },
  por_hora:   { nombre: 'Por hora',      cobra: true,  ayuda: 'Se paga lo que se trabaje.' },
  sin_cobro:  { nombre: 'No se cobra',   cobra: false, ayuda: 'Propio, o de la universidad.' },
};

/**
 * Sobre qué se calcula el porcentaje.
 *
 * Las tres dan números MUY distintos sobre el mismo mes, así que no
 * puede haber una por defecto escondida: se elige y se muestra siempre
 * al lado de la cifra.
 */
const PROY_BASES = {
  utilidad_neta: {
    nombre: 'Utilidad neta del mes',
    ayuda: 'Ganancia − devoluciones − pauta − costos fijos. La más baja y la más honesta.' },
  utilidad_post_pauta: {
    nombre: 'Utilidad después de pauta',
    ayuda: 'Ganancia − devoluciones − pauta. No descuenta bodega, sueldos ni herramientas.' },
  utilidad_antes_pauta: {
    nombre: 'Utilidad antes de pauta',
    ayuda: 'Ganancia − devoluciones. Se reparte antes de pagar la publicidad.' },
  ventas: {
    nombre: 'Ventas',
    ayuda: 'Sobre lo facturado, sin descontar nada. Casi nunca es lo que se quiere.' },
};

/** Los tipos de fuente que puede tener un proyecto. */
const PROY_FUENTES = [
  { id: 'excel',  nombre: 'Excel o CSV',   lee: true,
    ayuda: 'Nova ya sabe leer hojas de cálculo: se puede importar de verdad.' },
  { id: 'pdf',    nombre: 'PDF',           lee: false,
    ayuda: 'Se guarda el enlace. Para entenderlo solo falta el modelo de lenguaje.' },
  { id: 'word',   nombre: 'Word',          lee: false, ayuda: 'Se guarda el enlace.' },
  { id: 'ppt',    nombre: 'Presentación',  lee: false, ayuda: 'Se guarda el enlace.' },
  { id: 'claude', nombre: 'Artefacto de Claude', lee: false,
    ayuda: 'Se guarda el enlace. Nova no puede abrirlo: pega el texto si quieres las fechas.' },
  { id: 'drive',  nombre: 'Carpeta de Drive', lee: false, ayuda: 'Toda la carpeta, de una.' },
  { id: 'otro',   nombre: 'Otro',          lee: false, ayuda: '' },
];

function proyRol_(t) {
  const r = norm(t.mi_rol);
  if (PROY_ROLES[r]) return r;
  // Sin rol escrito se deduce del tipo, y se dice en pantalla que se
  // dedujo: un valor supuesto que parece escrito es el que nadie revisa.
  const tipo = norm(t.tipo);
  if (tipo === 'estudio') return 'estudio';
  if (tipo === 'propio') return 'propio';
  if (tipo === 'empleo') return 'trabajadora';
  return 'trabajadora';
}

function proyModalidad_(t) {
  const m = norm(t.modalidad);
  if (PROY_MODALIDADES[m]) return m;
  if (num(t.porcentaje) > 0) return 'porcentaje';
  if (num(t.valor_acordado) > 0) return 'fijo';
  const tipo = norm(t.tipo);
  return (tipo === 'propio' || tipo === 'estudio') ? 'sin_cobro' : 'fijo';
}

/**
 * ¿El texto de las tareas de este proyecto sube a Central?
 *
 * Vacío NO es «no». Para un empleo, vacío es «sí, es confidencial»:
 * lo de PHH lo es, y el valor por defecto tiene que proteger en vez de
 * exponer.
 */
function proyConfidencial_(t) {
  const v = norm(t.confidencial);
  if (v === 'si' || v === 'sí' || v === 'true') return true;
  if (v === 'no' || v === 'false') return false;
  return norm(t.tipo) === 'empleo';
}

// ─── LA UTILIDAD DE UNA TIENDA, MES A MES ────────────────────

/**
 * Un parámetro suelto de la hoja Parametros.
 *
 * `ajustes()` no sirve para esto: solo devuelve las claves que ya
 * conoce, y `costos_fijos_mes` no es una de ellas. Buscarlo por ahí
 * devolvía vacío siempre, y vacío aquí significa «no los desconté» —
 * que es justo lo que no puede pasar sin avisar.
 *
 * La fila de la tienda manda sobre la fila global `*`.
 */
function parametroDe_(ss, tienda, clave) {
  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return '';
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda'), cK = e.indexOf('clave'), cV = e.indexOf('valor');
  if (cK === -1 || cV === -1) return '';
  let global = '';
  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][cK]) !== norm(clave)) continue;
    const t = String(cT === -1 ? '' : d[i][cT] || '').trim();
    const v = String(d[i][cV] == null ? '' : d[i][cV]).trim();
    if (t === tienda) return v;
    if (!t || t === '*') global = v;
  }
  return global;
}

/**
 * La utilidad neta de un mes de una tienda, con todo lo que se le resta.
 *
 * ┌────────────────────────────────────────────────────────────┐
 * │  ganancia            lo que dejaron los pedidos entregados │
 * │  − devoluciones      el flete que se pagó y no volvió      │
 * │  − pauta             lo que se gastó en publicidad         │
 * │  − costos fijos      bodega, sueldos, herramientas         │
 * │  = utilidad neta                                           │
 * └────────────────────────────────────────────────────────────┘
 *
 * Cada resta que NO se pudo hacer se anota en `faltan`, y la pantalla
 * lo dice al lado de la cifra. Un costo fijo vacío tratado como cero
 * sube la utilidad, y sobre esa utilidad se reparte plata de verdad:
 * es exactamente el error que nadie revisa porque da un número alto.
 */
function utilidadMes_(sheetId, tienda, mes) {
  const ss = SpreadsheetApp.openById(sheetId);
  const moneda = monedaDeTienda(ss, tienda) || '';
  const desde = mes + '-01';
  const hasta = mes + '-31';
  const d = datosParaSemaforo_(ss, tienda, desde, hasta);

  let ganancia = 0, costoDevoluciones = 0, ventas = 0;
  let entregados = 0, devoluciones = 0, pedidos = 0;
  d.pedidos.forEach(function (p) {
    pedidos++;
    const est = p.estado;
    if (est === ESTADOS.ENTREGADO) {
      entregados++;
      ventas += p.valor;
      ganancia += p.valor - p.costoProducto - p.costoEnvio;
    } else if (est === ESTADOS.DEVOLUCION) {
      devoluciones++;
      costoDevoluciones += p.costoEnvio;
    }
  });

  const faltan = [];
  let pauta = 0, pautaSinConvertir = 0;
  d.pauta.forEach(function (g) {
    if (g.convertido === null) { pautaSinConvertir += g.gasto; return; }
    pauta += g.convertido;
  });
  if (pautaSinConvertir > 0) faltan.push('tasa_de_cambio');
  if (!d.pauta.length) faltan.push('pauta');

  // Los costos fijos los pone la dueña en Parametros. Vacío no es cero.
  const crudo = parametroDe_(ss, tienda, 'costos_fijos_mes');
  const hayFijos = crudo !== '' && crudo !== null && crudo !== undefined && !isNaN(Number(crudo));
  const costosFijos = hayFijos ? Number(crudo) : 0;
  if (!hayFijos) faltan.push('costos_fijos_mes');

  const antesPauta = ganancia - costoDevoluciones;
  const postPauta = antesPauta - pauta;
  const neta = postPauta - costosFijos;

  return {
    tienda: tienda, mes: mes, moneda: moneda,
    pedidos: pedidos, entregados: entregados, devoluciones: devoluciones,
    ventas: ventas, ganancia: ganancia, costoDevoluciones: costoDevoluciones,
    pauta: pauta, pautaSinConvertir: pautaSinConvertir,
    costosFijos: costosFijos, hayCostosFijos: hayFijos,
    utilidad_antes_pauta: antesPauta,
    utilidad_post_pauta: postPauta,
    utilidad_neta: neta,
    faltan: faltan,
    // Sin un solo pedido no hay mes: es distinto de un mes malo.
    hayDatos: pedidos > 0,
  };
}

/** La hoja y el nombre de una tienda, buscando entre los clientes. */
function tiendaDeCentral_(tiendaId) {
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!sh || sh.getLastRow() < 2) return null;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cSheet = e.indexOf('sheet_id'), cEmp = e.indexOf('empresa'), cId = e.indexOf('id');
  for (let i = 1; i < d.length; i++) {
    const sid = String(d[i][cSheet] || '').trim();
    if (!sid) continue;
    try {
      const cs = SpreadsheetApp.openById(sid);
      if (tiendasActivas_(cs).indexOf(tiendaId) === -1) continue;
      return { sheetId: sid, clienteId: String(d[i][cId] || ''),
               empresa: String(d[i][cEmp] || ''), nombre: nombreTienda(cs, tiendaId) };
    } catch (err) { /* una hoja que no abre no rompe la búsqueda */ }
  }
  return null;
}

/** Todas las tiendas que Central puede ver, para elegir en un proyecto. */
function tiendasParaProyecto_() {
  const out = [];
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
    if (!sh || sh.getLastRow() < 2) return out;
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const sid = String(d[i][c('sheet_id')] || '').trim();
      if (!sid) continue;
      if (norm(d[i][c('estado')]) === 'inactivo') continue;
      try {
        const cs = SpreadsheetApp.openById(sid);
        tiendasActivas_(cs).forEach(function (t) {
          out.push({ id: t, nombre: nombreTienda(cs, t),
                     empresa: String(d[i][c('empresa')] || ''),
                     clienteId: String(d[i][c('id')] || ''),
                     moneda: monedaDeTienda(cs, t) || '' });
        });
      } catch (err) { /* idem */ }
    }
  } catch (err) { /* Central sin clientes todavía */ }
  return out;
}

/**
 * Lo que le toca a ella este mes por un proyecto de porcentaje.
 *
 * Devuelve la cuenta ENTERA —la base, el porcentaje y el resultado— y
 * no solo el resultado. Un número sin su cuenta no se puede discutir, y
 * este es el número con el que se decide si el mes alcanza.
 */
function parteDelMes_(t, mes) {
  const tiendaId = String(t.tienda_id || '').trim();
  if (!tiendaId) return { hay: false, porque: 'Este proyecto no está enlazado a ninguna tienda.' };
  const pct = num(t.porcentaje);
  if (!pct) return { hay: false, porque: 'Falta decir qué porcentaje te toca.' };

  const ref = tiendaDeCentral_(tiendaId);
  if (!ref) return { hay: false, porque: 'No encontré la tienda ' + tiendaId + ' en ningún cliente.' };

  let u;
  try { u = utilidadMes_(ref.sheetId, tiendaId, mes); }
  catch (e) { return { hay: false, porque: 'No pude leer la tienda: ' + e.message }; }

  const base = PROY_BASES[norm(t.base_porcentaje)] ? norm(t.base_porcentaje) : 'utilidad_neta';
  const valorBase = u[base];
  return {
    hay: true,
    tienda: tiendaId, tiendaNombre: ref.nombre, empresa: ref.empresa,
    mes: mes, moneda: u.moneda,
    base: base, baseNombre: PROY_BASES[base].nombre, valorBase: valorBase,
    porcentaje: pct,
    // Una utilidad negativa no se reparte: no se le cobra a nadie por
    // un mes malo, y mostrar un número negativo como «tu parte» confunde.
    parte: valorBase > 0 ? valorBase * pct / 100 : 0,
    enPerdida: valorBase < 0,
    detalle: u,
    faltan: u.faltan,
    hayDatos: u.hayDatos,
  };
}

// ─── LAS FUENTES ─────────────────────────────────────────────

function fuentesDe_(trabajoId) {
  try {
    return mioLeer_('Fuentes')
      .filter(function (f) { return String(f.trabajo_id || '') === String(trabajoId); })
      .map(function (f) {
        const tipo = norm(f.tipo);
        const def = PROY_FUENTES.filter(function (x) { return x.id === tipo; })[0] ||
                    { nombre: 'Otro', lee: false, ayuda: '' };
        return { id: String(f.id || ''), nombre: String(f.nombre || ''),
                 tipo: tipo || 'otro', tipoNombre: def.nombre, lee: !!def.lee,
                 enlace: String(f.enlace || ''), nota: String(f.nota || ''),
                 agregado: aISO(f.agregado_en, 'UTC') || '' };
      });
  } catch (e) { return []; }
}

function centralFuenteGuardar(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const d = p.datos || {};
  const esNueva = !String(d.id || '').trim();
  if (esNueva && !String(d.trabajo_id || '').trim()) {
    return { ok: false, error: 'Falta decir de qué proyecto es.' };
  }
  if (esNueva && !String(d.nombre || '').trim()) {
    return { ok: false, error: 'Ponle un nombre a la fuente.' };
  }
  const enlace = d.enlace !== undefined ? String(d.enlace).trim() : undefined;
  if (enlace && !/^https?:\/\//i.test(enlace)) {
    return { ok: false, error: 'El enlace tiene que empezar por http.' };
  }
  const tipo = norm(d.tipo);
  if (tipo && !PROY_FUENTES.filter(function (x) { return x.id === tipo; })[0]) {
    return { ok: false, error: 'Ese tipo de fuente no está en la lista.' };
  }
  try {
    mioGuardar_('Fuentes', {
      id: String(d.id || ''),
      trabajo_id: d.trabajo_id !== undefined ? String(d.trabajo_id) : undefined,
      nombre: d.nombre !== undefined ? String(d.nombre).trim() : undefined,
      tipo: tipo || (esNueva ? 'otro' : undefined),
      enlace: enlace,
      nota: d.nota !== undefined ? String(d.nota) : undefined,
      agregado_en: esNueva ? ahoraISO().slice(0, 10) : undefined,
    });
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function centralFuenteBorrar(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  try {
    const fue = mioBorrar_('Fuentes', p.id);
    return { ok: fue, error: fue ? '' : 'No la encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

// ─── EL PROYECTO, A FONDO ────────────────────────────────────

function centralProyecto(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const id = String(p.id || '').trim();
  const t = mioLeer_('Trabajos').filter(function (x) { return String(x.id) === id; })[0];
  if (!t) return { ok: false, error: 'No encontré ese proyecto.' };

  const hoy = ahoraISO().slice(0, 10);
  const mes = String(p.mes || hoy.slice(0, 7));
  const lunes = lunesDe_(hoy), domingo = masDias_(lunes, 6);
  const rol = proyRol_(t), modalidad = proyModalidad_(t);
  const confidencial = proyConfidencial_(t);

  // ── Cobros ──
  const cobros = mioLeer_('Cobros')
    .filter(function (c) { return String(c.trabajo_id || '') === id; })
    .map(function (c) {
      const esperada = aISO(c.fecha_esperada, 'UTC') || '';
      const cobrado = norm(c.estado) === 'cobrado' || !!c.fecha_cobrada;
      return { id: String(c.id || ''), concepto: String(c.concepto || ''),
               monto: num(c.monto), moneda: String(c.moneda || '').toUpperCase(),
               esperada: esperada, cobrada: aISO(c.fecha_cobrada, 'UTC') || '',
               cobrado: cobrado,
               dias: esperada && !cobrado
                 ? Math.round((new Date(hoy + 'T00:00:00Z') -
                               new Date(esperada + 'T00:00:00Z')) / 86400000) : null };
    })
    .sort(function (a, b) { return (a.esperada || '9') < (b.esperada || '9') ? -1 : 1; });

  const porMoneda = function (filas) {
    const m = {};
    filas.forEach(function (f) { m[f.moneda || '?'] = (m[f.moneda || '?'] || 0) + f.monto; });
    return m;
  };

  // ── Movimientos de este proyecto ──
  let movimientos = [];
  try {
    movimientos = mioLeer_('Finanzas')
      .filter(function (f) { return String(f.trabajo_id || '') === id; })
      .map(function (f) {
        return { id: String(f.id || ''), fecha: aISO(f.fecha, 'UTC') || '',
                 flujo: norm(f.flujo), categoria: String(f.categoria || ''),
                 concepto: String(f.concepto || ''), monto: num(f.monto),
                 moneda: String(f.moneda || '').toUpperCase() };
      })
      .sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });
  } catch (e) { /* Finanzas puede no estar */ }

  // ── Las tareas, con la frontera puesta ──
  const uid = String(s.correo || '').toLowerCase();
  let tareas = [], carga = { abiertas: 0, horas: 0, vencidas: 0 };
  try {
    const todas = soulLeerSuave_('Pendientes', uid, [])
      .filter(function (f) { return String(f.trabajo_id || '') === id && !soulHecho_(f); });
    carga = {
      abiertas: todas.length,
      horas: todas.reduce(function (a, f) { return a + num(f.horas_estimadas); }, 0),
      vencidas: todas.filter(function (f) {
        const fe = aISO(f.fecha, 'UTC') || '';
        return fe && fe < hoy;
      }).length,
    };
    /**
     * El texto solo sube si el proyecto no es confidencial. Para PHH
     * Central ve «3 entregas, 14 horas» y nada más — que es justo lo
     * que ella pidió que se pudiera compartir.
     */
    if (!confidencial) {
      tareas = todas.map(function (f) {
        const fe = aISO(f.fecha, 'UTC') || '';
        return {
          id: String(f.id || ''), texto: String(f.texto || ''), fecha: fe,
          horas: num(f.horas_estimadas),
          estado: norm(f.estado) || 'pendiente',
          riesgo: norm(f.riesgo) || 'corrible',
          estaSemana: !!fe && fe >= lunes && fe <= domingo,
          vencida: !!fe && fe < hoy,
        };
      }).sort(function (a, b) { return (a.fecha || '9') < (b.fecha || '9') ? -1 : 1; });
    }
  } catch (e) { /* Soul puede no estar */ }

  const cobrado = cobros.filter(function (c) { return c.cobrado; });
  const pendiente = cobros.filter(function (c) { return !c.cobrado; });

  return {
    ok: true, hoy: hoy, mes: mes, semana: { lunes: lunes, domingo: domingo },
    roles: PROY_ROLES, modalidades: PROY_MODALIDADES, bases: PROY_BASES,
    tiposFuente: PROY_FUENTES,
    tiendas: tiendasParaProyecto_(),
    proyecto: {
      id: id, nombre: String(t.nombre || ''), contraparte: String(t.contraparte || ''),
      tipo: norm(t.tipo) || 'cliente', estado: norm(t.estado) || 'activo',
      moneda: String(t.moneda || '').toUpperCase(),
      valor: num(t.valor_acordado), formaCobro: String(t.forma_cobro || ''),
      inicio: aISO(t.fecha_inicio, 'UTC') || '', entrega: aISO(t.fecha_entrega, 'UTC') || '',
      horasSemana: num(t.horas_semana),
      especificacion: String(t.especificacion || ''),
      documento: String(t.documento || ''), nota: String(t.nota || ''),
      rol: rol, rolDeducido: !PROY_ROLES[norm(t.mi_rol)],
      modalidad: modalidad, modalidadDeducida: !PROY_MODALIDADES[norm(t.modalidad)],
      porcentaje: num(t.porcentaje),
      base: PROY_BASES[norm(t.base_porcentaje)] ? norm(t.base_porcentaje) : 'utilidad_neta',
      tiendaId: String(t.tienda_id || ''), clienteId: String(t.cliente_id || ''),
      confidencial: confidencial,
    },
    fuentes: fuentesDe_(id),
    cobros: cobros,
    plata: {
      cobrado: porMoneda(cobrado), porCobrar: porMoneda(pendiente),
      atrasado: porMoneda(pendiente.filter(function (c) { return c.dias !== null && c.dias > 0; })),
      // Lo que falta contra el valor acordado: vacío si no hay valor,
      // en vez de restar contra cero y prometer un número.
      falta: num(t.valor_acordado)
        ? num(t.valor_acordado) - cobrado.reduce(function (a, c) { return a + c.monto; }, 0)
        : null,
    },
    parte: modalidad === 'porcentaje' ? parteDelMes_(t, mes) : null,
    tareas: tareas,
    carga: carga,
    movimientos: movimientos,
  };
}

/**
 * Pegar el alcance de un proyecto y sacarle las fechas.
 *
 * El mismo lector del sílabo, por lo mismo: propone y ella confirma,
 * muestra lo que ignoró, y avisa cuando adivinó el año. Un contrato y
 * un cronograma de universidad se parecen más de lo que uno creería —
 * los dos son una lista de entregas con fecha.
 */
function centralProyectoLeer(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const texto = String(p.texto || '');
  if (!texto.trim()) return { ok: false, error: 'Pega el texto primero.' };
  const r = silaboLeer_(texto, ahoraISO().slice(0, 10));
  return { ok: true, encontradas: r.encontradas, ignoradas: r.ignoradas,
           truncado: r.truncado, maximo: SILABO_LINEAS_MAX };
}

/**
 * Guarda las entregas confirmadas COMO PENDIENTES DE NOVASOUL.
 *
 * No se guardan en Central aunque se creen desde aquí: las entregas con
 * fecha son del día a día, y el día a día vive en Soul. Si vivieran en
 * las dos, la semana se contaría dos veces.
 */
function centralProyectoGuardarTareas(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const id = String(p.trabajo || '').trim();
  const t = mioLeer_('Trabajos').filter(function (x) { return String(x.id) === id; })[0];
  if (!t) return { ok: false, error: 'No encontré ese proyecto.' };
  const items = p.items || [];
  if (!items.length) return { ok: false, error: 'No marcaste ninguna entrega.' };

  const uid = String(s.correo || '').toLowerCase();
  const yaEstan = {};
  soulLeerSuave_('Pendientes', uid, []).forEach(function (f) {
    if (String(f.trabajo_id || '') !== id) return;
    yaEstan[norm(f.texto) + '|' + (aISO(f.fecha, 'UTC') || '')] = 1;
  });

  let creadas = 0, repetidas = 0;
  const errores = [];
  items.forEach(function (it) {
    const texto = String(it.titulo || '').trim();
    const fecha = String(it.fecha || '').trim();
    if (!texto || !/^\d{4}-\d{2}-\d{2}$/.test(fecha)) { errores.push(texto || '(sin título)'); return; }
    if (yaEstan[norm(texto) + '|' + fecha]) { repetidas++; return; }
    try {
      soulGuardar_('Pendientes', {
        texto: texto, origen: 'proyecto', trabajo_id: id, fecha: fecha,
        estado: 'pendiente', prioridad: 'media',
        // Lo acordado con un cliente se puede correr avisando; no es
        // inamovible como un parcial, ni libre como algo propio.
        riesgo: norm(t.tipo) === 'cliente' ? 'acordado' : 'corrible',
        horas_estimadas: it.horas !== undefined ? num(it.horas) : '',
        nota: it.peso ? String(it.peso) + '%' : '',
      }, uid);
      yaEstan[norm(texto) + '|' + fecha] = 1;
      creadas++;
    } catch (e) { errores.push(texto + ': ' + e.message); }
  });

  return { ok: true, creadas: creadas, repetidas: repetidas, errores: errores };
}
