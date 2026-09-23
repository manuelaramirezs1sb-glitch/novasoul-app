/**
 * ═══════════════════════════════════════════════════════════
 *  META, TODO EN UNA SOLA PANTALLA
 * ═══════════════════════════════════════════════════════════
 *
 * Ella lo dijo así: «no veo dónde se ven los datos que lee Nova de Meta
 * […] todo está regado […] que cargue la información en la misma
 * pantalla, no que me toque ir a configuración».
 *
 * Y tenía razón: la llave y el botón de traer estaban en Configuración,
 * los números en «Pauta y gastos», y el semáforo en otra parte. Tres
 * sitios para una sola pregunta —¿cómo va la pauta?— y ninguno que la
 * respondiera solo.
 *
 * Esto devuelve TODO lo de Meta de una vez:
 *   · si está conectado, y si no, qué falta exactamente
 *   · lo de HOY, con las tres revisiones del día
 *   · los periodos: hoy, 7 días, mes y trimestre
 *   · el desglose por campaña y por anuncio
 *
 * ── LAS TRES REVISIONES DEL DÍA ──
 *
 * También suyas: «sacar un semáforo diario, son 3 revisiones por día, al
 * mediodía, a las 5pm y a las 11:30pm».
 *
 * No son tres horas caprichosas: son los tres momentos en que todavía se
 * puede hacer algo distinto. Al mediodía se ve si el día arrancó torcido
 * y aún queda tarde para corregir; a las 5 es la última hora útil para
 * mover presupuesto; a las 11:30 ya no se corrige nada, se cierra y se
 * aprende.
 *
 * Por eso cada revisión NO dice lo mismo. La de la noche no propone
 * mover nada: sería un consejo que no se puede seguir.
 */

const META_REVISIONES = [
  { id: 'mediodia', hora: 12, minuto: 0, nombre: 'Mediodía',
    que: 'Todavía queda media jornada. Si el CPA va alto, aquí se corrige.' },
  { id: 'tarde', hora: 17, minuto: 0, nombre: '5:00 p.m.',
    que: 'Última hora útil para mover presupuesto y que alcance a notarse.' },
  { id: 'cierre', hora: 23, minuto: 30, nombre: '11:30 p.m.',
    que: 'El día ya está. Esto no es para corregir: es para saber con qué cerró.' },
];

/**
 * Cuánto del día lleva corrido en cada revisión.
 *
 * Sirve para lo único que hace justa la comparación: a mediodía no se
 * puede exigir el gasto de un día entero. Se compara contra la PARTE del
 * día que va, no contra el día completo.
 */
function metaFraccionDia_(hora, minuto) {
  return Math.min(1, Math.max(0, (hora * 60 + minuto) / (24 * 60)));
}

/** La hora local de la tienda, en minutos desde medianoche. */
function metaMinutosAhora_(ss, tienda) {
  const z = String(ajustes(ss, tienda).zona_horaria || '').trim() ||
            Session.getScriptTimeZone() || 'UTC';
  try {
    const hm = Utilities.formatDate(new Date(), z, 'HH:mm').split(':');
    return Number(hm[0]) * 60 + Number(hm[1]);
  } catch (e) {
    const hm = Utilities.formatDate(new Date(), 'UTC', 'HH:mm').split(':');
    return Number(hm[0]) * 60 + Number(hm[1]);
  }
}

/**
 * Los números de un rango: lo que se gastó y lo que entró.
 *
 * `gastoSinConvertir` no es un detalle: si Meta cobra en otra moneda y
 * falta la tasa de ese día, ese gasto NO está dentro del CPA. Decir el
 * CPA sin decir eso es dar por bueno un número al que le falta plata.
 */
function metaNumeros_(d, desde, hasta) {
  const enRango = function (f) { return f >= desde && f <= hasta; };
  const peds = d.pedidos.filter(function (p) { return enRango(p.fecha); });
  const gastos = d.pauta.filter(function (g) { return enRango(g.fecha); });

  let gasto = 0, sinConvertir = 0;
  gastos.forEach(function (g) {
    if (g.convertido === null) sinConvertir += g.gasto;
    else gasto += g.convertido;
  });

  const entregados = peds.filter(function (p) { return p.estado === 'entregado'; });
  const devueltos = peds.filter(function (p) { return p.estado === 'devolucion'; });
  const ventas = entregados.reduce(function (a, p) { return a + p.valor; }, 0);
  const costos = entregados.reduce(function (a, p) {
    return a + p.costoProducto + p.costoEnvio; }, 0);
  // Una devolución no vende y su flete se paga igual.
  const costoDev = devueltos.reduce(function (a, p) { return a + p.costoEnvio; }, 0);
  const utilidadAntes = ventas - costos - costoDev;
  const resueltos = entregados.length + devueltos.length;

  return {
    desde: desde, hasta: hasta,
    pedidos: peds.length,
    entregados: entregados.length,
    devoluciones: devueltos.length,
    ventas: Math.round(ventas * 100) / 100,
    gasto: Math.round(gasto * 100) / 100,
    gastoSinConvertir: Math.round(sinConvertir * 100) / 100,
    // Sin pedidos no hay CPA. Dividir entre cero y mandar Infinity, o
    // mandar cero, son dos formas de decir algo que no se sabe.
    cpa: peds.length ? Math.round(gasto / peds.length * 100) / 100 : null,
    cpaEntrega: entregados.length
      ? Math.round(gasto / entregados.length * 100) / 100 : null,
    utilidadAntesPauta: Math.round(utilidadAntes * 100) / 100,
    utilidadReal: Math.round((utilidadAntes - gasto) * 100) / 100,
    roas: gasto ? Math.round(ventas / gasto * 100) / 100 : null,
    ticket: entregados.length ? Math.round(ventas / entregados.length * 100) / 100 : null,
    tasaEntrega: resueltos
      ? Math.round(entregados.length / resueltos * 1000) / 10 : null,
    enTransito: peds.length - resueltos -
                peds.filter(function (p) { return p.estado === 'cancelado'; }).length,
  };
}

/** Lo que Meta trajo de impresiones, alcance y clics en un rango. */
function metaAlcance_(ss, tienda, desde, hasta) {
  const sh = ss.getSheetByName('Pauta');
  const out = { impresiones: 0, alcance: 0, clics: 0, ctr: null, cpm: null, hay: false };
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  if (c('impresiones') === -1) return out;

  let gasto = 0;
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')]).trim() !== tienda) continue;
    const f = aISO(d[i][c('fecha')], 'UTC');
    if (!f || f < desde || f > hasta) continue;
    out.impresiones += num(d[i][c('impresiones')]);
    out.alcance += num(d[i][c('alcance')]);
    out.clics += num(d[i][c('clics')]);
    gasto += num(d[i][c('gasto')]);
    out.hay = true;
  }
  if (out.impresiones) {
    out.ctr = Math.round(out.clics / out.impresiones * 10000) / 100;
    out.cpm = Math.round(gasto / out.impresiones * 1000 * 100) / 100;
  }
  return out;
}

/** El desglose por campaña, y por anuncio si la hoja Anuncios existe. */
function metaDesglose_(ss, tienda, desde, hasta, monTienda) {
  const junta = function (hoja, claveCol) {
    const sh = ss.getSheetByName(hoja);
    if (!sh || sh.getLastRow() < 2) return [];
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    if (c(claveCol) === -1) return [];

    const acc = {};
    for (let i = 1; i < d.length; i++) {
      if (c('tienda') !== -1 && String(d[i][c('tienda')]).trim() !== tienda) continue;
      const f = aISO(d[i][c('fecha')], 'UTC');
      if (!f || f < desde || f > hasta) continue;
      const k = String(d[i][c(claveCol)] || '').trim() || '(sin nombre)';
      if (!acc[k]) acc[k] = { nombre: k, gasto: 0, impresiones: 0, clics: 0,
                              resultados: 0, sinConvertir: 0 };
      const g = num(d[i][c('gasto')]);
      const mon = String(d[i][c('moneda_gasto')] || '').toUpperCase();
      if (mon && monTienda && mon !== monTienda) {
        const t = buscarTasa(ss, f, mon, monTienda);
        if (t && t.tasa) acc[k].gasto += g * t.tasa;
        else acc[k].sinConvertir += g;
      } else {
        acc[k].gasto += g;
      }
      if (c('impresiones') !== -1) acc[k].impresiones += num(d[i][c('impresiones')]);
      if (c('clics') !== -1) acc[k].clics += num(d[i][c('clics')]);
      if (c('resultados') !== -1) acc[k].resultados += num(d[i][c('resultados')]);
    }
    return Object.keys(acc).map(function (k) {
      const x = acc[k];
      x.gasto = Math.round(x.gasto * 100) / 100;
      x.sinConvertir = Math.round(x.sinConvertir * 100) / 100;
      x.cpa = x.resultados ? Math.round(x.gasto / x.resultados * 100) / 100 : null;
      return x;
    }).sort(function (a, b) { return b.gasto - a.gasto; });
  };

  return { campanas: junta('Pauta', 'campana'), anuncios: junta('Anuncios', 'anuncio') };
}

/**
 * Las tres revisiones del día.
 *
 * Cada una se juzga contra la PARTE del día que va, no contra el día
 * entero: a mediodía el presupuesto todavía no se gastó, y llamarlo
 * «vas bien de gasto» sería felicitarla por no haber llegado aún.
 */
function metaRevisiones_(hoy, techo, presupuestoDia, minutosAhora, moneda) {
  return META_REVISIONES.map(function (r) {
    const minuto = r.hora * 60 + r.minuto;
    const yaPaso = minutosAhora >= minuto;
    const frac = metaFraccionDia_(r.hora, r.minuto);

    const base = { id: r.id, nombre: r.nombre, que: r.que, yaPaso: yaPaso,
                   hora: (r.hora < 10 ? '0' : '') + r.hora + ':' +
                         (r.minuto < 10 ? '0' : '') + r.minuto };

    if (!yaPaso) {
      return Object.assign(base, { estado: 'pendiente',
        porque: 'Todavía no llega. A esa hora se vuelve a mirar.' });
    }

    /**
     * Sin pedidos del día no se pinta un color.
     *
     * Con gasto y cero pedidos a mediodía la respuesta no es «rojo»: es
     * que puede ser normal y todavía no hay con qué saberlo. Un rojo a
     * las doce por un día que termina bien enseña a no mirar el semáforo.
     */
    if (!hoy.pedidos && !hoy.gasto) {
      return Object.assign(base, { estado: 'sin_datos',
        porque: 'Ni gasto ni pedidos todavía. Puede ser que Meta no haya ' +
                'reportado aún: la lectura entra por la mañana.' });
    }
    if (!hoy.pedidos) {
      return Object.assign(base, { estado: 'sin_datos',
        porque: 'Hay ' + redondear_(hoy.gasto) + ' ' + moneda +
                ' de gasto y ningún pedido todavía.' });
    }

    // El CPA del día contra el techo que ella misma definió.
    if (techo === null) {
      return Object.assign(base, { estado: 'sin_umbral',
        porque: 'No hay techo de CPA definido para esta tienda, así que ' +
                'no hay contra qué comparar.' });
    }
    const cpa = hoy.cpa;
    const estado = cpa <= techo * 0.8 ? 'verde'
                 : cpa <= techo ? 'amarillo' : 'rojo';
    let porque = 'CPA del día: ' + redondear_(cpa) + ' ' + moneda +
                 '. El techo es ' + redondear_(techo) + '.';

    // El presupuesto, contra la parte del día que va.
    if (presupuestoDia) {
      const esperado = presupuestoDia * frac;
      const pct = Math.round(hoy.gasto / esperado * 100);
      porque += ' Llevas ' + redondear_(hoy.gasto) + ' de los ' +
                redondear_(esperado) + ' que tocarían a esta hora (' + pct + '%).';
    }
    if (r.id === 'cierre') {
      porque += ' El día ya cerró: esto es para el registro, no para mover nada.';
    }
    return Object.assign(base, { estado: estado, porque: porque });
  });
}

/** Qué hacer con lo que hay. Nunca un consejo que ya no se puede seguir. */
function metaQueHacer_(hoy, techo, revisiones, moneda) {
  const ultima = revisiones.filter(function (r) { return r.yaPaso; }).pop();
  const esNoche = ultima && ultima.id === 'cierre';

  if (hoy.gastoSinConvertir) {
    return 'Hay ' + redondear_(hoy.gastoSinConvertir) + ' de pauta en otra moneda ' +
           'sin convertir: ese gasto NO está dentro del CPA de arriba. ' +
           'Corre las tasas antes de decidir nada con este número.';
  }
  if (!hoy.gasto && !hoy.pedidos) {
    return 'Todavía no hay nada del día. Si ya pasó la mañana y sigue vacío, ' +
           'lo más probable es que falte traer la lectura de Meta.';
  }
  if (!hoy.pedidos) {
    return esNoche
      ? 'El día cerró con gasto y sin pedidos. Vale la pena mirar mañana el ' +
        'anuncio y la landing antes de volver a invertir ahí.'
      : 'Gasto sin pedidos todavía. Si a la próxima revisión sigue igual, ' +
        'conviene pausar ese conjunto.';
  }
  if (techo === null) {
    return 'Define el techo de CPA de esta tienda en Ajustes y el semáforo ' +
           'empieza a significar algo.';
  }
  if (hoy.cpa > techo) {
    return esNoche
      ? 'El día cerró con el CPA por encima del techo. No hay nada que mover ' +
        'hoy; mañana temprano, revisa qué campaña lo subió.'
      : 'El CPA va por encima del techo (' + redondear_(hoy.cpa) + ' contra ' +
        redondear_(techo) + ' ' + moneda + '). Mira abajo qué campaña lo está ' +
        'subiendo y baja ese presupuesto.';
  }
  if (hoy.cpa <= techo * 0.8) {
    return esNoche
      ? 'Buen día: el CPA cerró cómodo bajo el techo. Anota qué campaña lo hizo.'
      : 'El CPA va cómodo bajo el techo. Si hay stock, es el momento de subir ' +
        'presupuesto en lo que está funcionando.';
  }
  return 'El CPA va justo debajo del techo. Ni para escalar ni para frenar: ' +
         'vale la pena mirarlo otra vez en la siguiente revisión.';
}

const META_PERIODOS = [
  { id: 'hoy', nombre: 'Hoy', dias: 0 },
  { id: 'semana', nombre: '7 días', dias: 6 },
  { id: 'mes', nombre: 'Este mes', dias: null },
  { id: 'trimestre', nombre: 'Trimestre', dias: 89 },
];

/**
 * TODO lo de Meta, en una sola respuesta.
 *
 * Una sola llamada a propósito: la pantalla tiene que poder abrirse y
 * mostrarlo todo sin encadenar peticiones. Antes había que ir a
 * Configuración a ver si estaba conectado y volver a Pauta a ver los
 * números, y entre las dos cosas nadie sabía si el número que estaba
 * mirando venía de una conexión viva o de un archivo de hace un mes.
 */
function apiMetaPanel(s, p) {
  if (!puede(s, 'leer', 'Pauta') && s.rol !== 'dueno') {
    return { ok: false, error: 'Tu rol no ve la pauta.' };
  }
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }

  const ss = SpreadsheetApp.openById(s.sheetId);
  const props = PropertiesService.getScriptProperties();
  const hoyISO = ahoraISO().slice(0, 10);
  const moneda = monedaDeTienda(ss, tienda) || '';

  // ── ¿Está conectado, y si no, qué falta? ──
  const hayLlave = !!metaToken(s);
  const cuenta = metaCuenta(ss, tienda);
  let prueba = null;
  try {
    const crudo = props.getProperty(metaClave(s) + '_PRUEBA');
    if (crudo) prueba = JSON.parse(crudo);
  } catch (e) { prueba = null; }

  // Cuándo fue la última fila de pauta que entró: es la prueba de vida.
  let ultimaLectura = '';
  const shA = ss.getSheetByName('Pauta');
  if (shA && shA.getLastRow() > 1) {
    const d = shA.getDataRange().getValues();
    const e = d[0].map(norm);
    const cT = e.indexOf('tienda'), cF = e.indexOf('fecha');
    for (let i = 1; i < d.length; i++) {
      if (cT !== -1 && String(d[i][cT]).trim() !== tienda) continue;
      const f = aISO(d[i][cF], 'UTC');
      if (f && f > ultimaLectura) ultimaLectura = f;
    }
  }

  const conexion = {
    hayLlave: hayLlave,
    cuenta: cuenta,
    /**
     * «Activo» es llave MÁS cuenta MÁS datos recientes. Con llave y sin
     * cuenta no entra nada, y decir «conectado» sería mentir despacio —
     * es el mismo criterio que usa la consola de Nova Central.
     */
    activo: !!(hayLlave && cuenta),
    alDia: !!(ultimaLectura && ultimaLectura >= masDias_(hoyISO, -2)),
    ultimaLectura: ultimaLectura,
    ultimaPrueba: prueba,
    guardadaEn: props.getProperty(metaClave(s) + '_FECHA') || '',
    falta: !hayLlave ? 'llave' : !cuenta ? 'cuenta' : '',
    porque: !hayLlave
      ? 'Todavía no has guardado la llave de Meta. Sin ella Nova no puede pedir nada.'
      : !cuenta
        ? 'Hay llave guardada, pero esta tienda no tiene número de cuenta ' +
          'publicitaria. El gasto no puede entrar hasta que lo pongas.'
        : !ultimaLectura
          ? 'Está conectado, pero todavía no ha entrado ninguna fila de pauta. ' +
            'Dale a «Traer ahora».'
          : '',
  };

  // ── Los periodos ──
  const trimestreDesde = masDias_(hoyISO, -89);
  const d = datosParaSemaforo_(ss, tienda, trimestreDesde, hoyISO);

  const periodos = META_PERIODOS.map(function (per) {
    const desde = per.id === 'mes' ? hoyISO.slice(0, 8) + '01'
                : masDias_(hoyISO, -per.dias);
    const n = metaNumeros_(d, desde, hoyISO);
    n.id = per.id;
    n.nombre = per.nombre;
    n.alcance = metaAlcance_(ss, tienda, desde, hoyISO);
    return n;
  });

  const hoy = periodos.filter(function (x) { return x.id === 'hoy'; })[0];

  // ── El techo de CPA y el presupuesto, de sus ajustes ──
  const u = ajustesSemaforo(ss, tienda);
  const semana = periodos.filter(function (x) { return x.id === 'semana'; })[0];
  /**
   * El techo sale del margen de los últimos 7 días, no del día: un solo
   * día mueve demasiado el ticket y el techo bailaría a diario.
   */
  const techo = (semana.entregados && u.cpa_verde_pct)
    ? Math.round(semana.utilidadAntesPauta / semana.entregados *
                 (u.cpa_verde_pct / 100) * 100) / 100
    : null;
  const presupuestoDia = num(parametroDe_(ss, tienda, 'presupuesto_diario_pauta'));

  const minutos = metaMinutosAhora_(ss, tienda);
  const revisiones = metaRevisiones_(hoy, techo, presupuestoDia, minutos, moneda);

  return {
    ok: true,
    tienda: tienda, nombreTienda: nombreTienda(ss, tienda),
    moneda: moneda, hoy: hoyISO,
    horaLocal: (Math.floor(minutos / 60) < 10 ? '0' : '') + Math.floor(minutos / 60) +
               ':' + (minutos % 60 < 10 ? '0' : '') + (minutos % 60),
    conexion: conexion,
    periodos: periodos,
    techo: techo,
    presupuestoDia: presupuestoDia || null,
    revisiones: revisiones,
    queHacer: metaQueHacer_(hoy, techo, revisiones, moneda),
    desglose: metaDesglose_(ss, tienda, masDias_(hoyISO, -6), hoyISO, moneda),
    puedeTraer: s.rol === 'dueno',
  };
}
