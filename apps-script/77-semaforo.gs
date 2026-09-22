/**
 * ═══════════════════════════════════════════════════════════
 *  EL SEMÁFORO SEMANAL
 * ═══════════════════════════════════════════════════════════
 *
 * Cada lunes en la mañana: qué pasó la semana que cerró, cruzando la
 * pauta de Meta con lo que de verdad se entregó.
 *
 * POR QUÉ CRUZARLOS ES EL PUNTO
 *
 * Meta cuenta una compra cuando alguien confirma en el checkout. En
 * contraentrega, una parte de esas compras nunca se entrega: el ROAS que
 * muestra Meta está inflado y con él se deciden presupuestos. Nova sabe
 * cuáles llegaron, porque tiene los dos lados.
 *
 * LAS REGLAS DE CÁLCULO, QUE NO SON NEGOCIABLES
 *
 * Son las que ya venían usándose a mano, escritas aquí para que no
 * cambien sin que alguien lo decida:
 *
 *   · Solo un pedido ENTREGADO genera ganancia.
 *   · Cada DEVOLUCIÓN cuesta el flete: el producto vuelve, el envío no.
 *   · Un CANCELADO cuesta cero — nunca salió.
 *   · La tasa de entrega es entregados ÷ (entregados + devoluciones).
 *     Lo que sigue en tránsito no entra: todavía no falló.
 *   · El techo de CPA es la utilidad antes de pauta ÷ pedidos creados.
 *     Es cuánto se puede pagar por pedido sin perder plata.
 *
 * Y LA REGLA QUE PROTEGE A TODAS LAS DEMÁS
 *
 * Donde no hay dato, dice «sin dato». Nada se extrapola, nada se rellena
 * con un promedio. Una semana sin pedidos en la hoja no significa que no
 * se vendió: significa que la hoja no se actualizó, y son dos cosas
 * distintas que llevan a decisiones opuestas.
 */

/**
 * Los umbrales de esta tienda, como números.
 *
 * Viven en AJUSTES_DEFAULT con el resto de ajustes, y cada tienda los
 * cambia en Parametros. Aquí solo se convierten: lo que vuelve de una
 * hoja es texto, y "72" no es 72 cuando se compara con un porcentaje.
 *
 * No hay una segunda lista de valores por defecto a propósito. Dos
 * listas del mismo umbral acaban diciendo cosas distintas, y entonces el
 * semáforo se contradice con la pantalla que lo configura.
 */
const SEMAFORO_CLAVES = ['ticket_minimo', 'entrega_minima',
                         'muestra_minima', 'cpa_verde_pct'];

function ajustesSemaforo(ss, tienda) {
  const a = ajustes(ss, tienda) || {};
  const out = {};
  SEMAFORO_CLAVES.forEach(function (k) {
    const v = Number(a[k]);
    out[k] = (a[k] === '' || a[k] === undefined || isNaN(v))
      ? Number(AJUSTES_DEFAULT[k]) || 0 : v;
  });
  return out;
}

// ─── SEMANAS ─────────────────────────────────────────────────

/** El lunes de la semana de una fecha ISO. */
function lunesDe_(iso) {
  const d = new Date(iso + 'T00:00:00Z');
  const dow = d.getUTCDay();                 // 0 domingo … 6 sábado
  const atras = (dow === 0) ? 6 : dow - 1;   // el domingo pertenece a la semana que cierra
  d.setUTCDate(d.getUTCDate() - atras);
  return d.toISOString().slice(0, 10);
}

function masDias_(iso, n) {
  const d = new Date(iso + 'T00:00:00Z');
  d.setUTCDate(d.getUTCDate() + n);
  return d.toISOString().slice(0, 10);
}

/** La semana cerrada más reciente: de lunes a domingo, ya terminada. */
function semanaCerrada_(hoyISO) {
  return masDias_(lunesDe_(hoyISO), -7);
}

// ─── LOS NÚMEROS DE UNA SEMANA ───────────────────────────────

/**
 * Todo lo que se puede contar de una semana, sin juzgar nada todavía.
 *
 * Separado del semáforo a propósito: los números son hechos y los
 * colores son opiniones. Poder mirar los hechos sin la opinión encima es
 * lo que permite discutir la opinión.
 */
function numerosSemana_(pedidos, pauta, lunes) {
  const domingo = masDias_(lunes, 6);
  const o = {
    lunes: lunes, domingo: domingo,
    pedidos: 0, despachados: 0, entregados: 0, devoluciones: 0, cancelados: 0,
    enTransito: 0, sinClasificar: 0,
    ventas: 0, ganancia: 0, costoDevoluciones: 0,
    gasto: 0, gastoSinConvertir: 0,
    diasConPedidos: {}, productos: {},
    sinCosto: 0,                 // entregados sin costo de producto
  };

  pedidos.forEach(function (p) {
    if (p.fecha < lunes || p.fecha > domingo) return;
    o.pedidos++;
    o.diasConPedidos[p.fecha] = (o.diasConPedidos[p.fecha] || 0) + 1;

    const est = p.estado;
    if (est !== ESTADOS.CANCELADO && est !== ESTADOS.PENDIENTE &&
        est !== ESTADOS.SIN_CLASIFICAR) o.despachados++;

    const prod = p.producto || '(sin producto)';
    if (!o.productos[prod]) {
      o.productos[prod] = { pedidos: 0, entregados: 0, devoluciones: 0, ganancia: 0 };
    }
    const pr = o.productos[prod];
    pr.pedidos++;

    if (est === ESTADOS.ENTREGADO) {
      o.entregados++; pr.entregados++;
      o.ventas += p.valor;
      /**
       * La ganancia solo se calcula si hay costo de producto.
       *
       * Sin costo, valor − flete daría una ganancia enorme y falsa. Se
       * cuenta aparte y se dice cuántos quedaron fuera: una utilidad a la
       * que le faltan pedidos es mejor que una utilidad inventada, pero
       * solo si se sabe que le faltan.
       */
      if (p.costoProducto > 0) {
        const g = p.valor - p.costoProducto - p.costoEnvio;
        o.ganancia += g;
        pr.ganancia += g;
      } else {
        o.sinCosto++;
      }
    } else if (est === ESTADOS.DEVOLUCION) {
      o.devoluciones++; pr.devoluciones++;
      // El producto vuelve; el flete ya se pagó y no vuelve.
      o.costoDevoluciones += p.costoEnvio;
    } else if (est === ESTADOS.CANCELADO) {
      o.cancelados++;
    } else if (est === ESTADOS.SIN_CLASIFICAR) {
      o.sinClasificar++;
    } else {
      o.enTransito++;
    }
  });

  pauta.forEach(function (g) {
    if (g.fecha < lunes || g.fecha > domingo) return;
    if (g.convertido === null) { o.gastoSinConvertir += g.gasto; return; }
    o.gasto += g.convertido;
  });

  const resueltos = o.entregados + o.devoluciones;
  o.tasaEntrega = resueltos ? o.entregados / resueltos * 100 : null;
  o.ticket = o.entregados ? o.ventas / o.entregados : null;
  o.utilidadAntesPauta = o.ganancia - o.costoDevoluciones;
  o.hayPauta = o.gasto > 0 || o.gastoSinConvertir > 0;
  o.utilidadReal = o.hayPauta ? o.utilidadAntesPauta - o.gasto : null;
  o.cpaPorPedido = (o.hayPauta && o.pedidos) ? o.gasto / o.pedidos : null;
  o.cpaPorEntrega = (o.hayPauta && o.entregados) ? o.gasto / o.entregados : null;
  // El techo: cuánto se puede pagar por pedido sin perder plata.
  o.techo = o.pedidos ? o.utilidadAntesPauta / o.pedidos : null;
  o.diasConDatos = Object.keys(o.diasConPedidos).length;
  return o;
}

// ─── LEER LAS HOJAS UNA SOLA VEZ ─────────────────────────────

/**
 * Pedidos y pauta de las últimas N semanas, ya normalizados.
 *
 * Se leen de una vez y se reparten por semana después. Abrir la hoja una
 * vez por semana analizada multiplicaba por ocho el trabajo para leer
 * exactamente los mismos datos.
 */
function datosParaSemaforo_(ss, tienda, desde, hasta) {
  const monTienda = monedaDeTienda(ss, tienda) || '';
  const pedidos = [];
  const pauta = [];

  const shP = ss.getSheetByName('Pedidos');
  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha || fecha < desde || fecha > hasta) continue;
      pedidos.push({
        fecha: fecha,
        estado: norm(f[c('estado_nova')] || f[c('estado_canonico')]) || ESTADOS.SIN_CLASIFICAR,
        valor: num(f[c('valor')]),
        costoProducto: num(f[c('costo_producto')]),
        costoEnvio: num(f[c('costo_envio')]),
        producto: String(f[c('producto')] || '').trim(),
      });
    }
  }

  const shA = ss.getSheetByName('Pauta');
  if (shA && shA.getLastRow() > 1) {
    const d = shA.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha || fecha < desde || fecha > hasta) continue;
      const gasto = num(f[c('gasto')]);
      const mon = String(f[c('moneda_gasto')] || '').toUpperCase();
      let convertido = gasto;
      if (mon && monTienda && mon !== monTienda) {
        const t = buscarTasa(ss, fecha, mon, monTienda);
        convertido = (t && t.tasa) ? gasto * t.tasa : null;
      }
      pauta.push({ fecha: fecha, gasto: gasto, convertido: convertido,
                   conjunto: String(f[c('conjunto')] || f[c('campana')] || '').trim() });
    }
  }

  return { pedidos: pedidos, pauta: pauta, moneda: monTienda };
}

// ─── EL SEMÁFORO ─────────────────────────────────────────────

/**
 * Un semáforo: qué mide, cuánto dio, de qué color y por qué.
 *
 * El «por qué» es obligatorio y va junto al color a propósito. Un color
 * sin su razón es una orden; con ella es un argumento, y un argumento se
 * puede discutir cuando el umbral está mal puesto.
 */
function luz_(estado, valor, etiqueta, porque, delta) {
  return { estado: estado, valor: valor, etiqueta: etiqueta, porque: porque,
           delta: (delta === undefined ? null : delta) };
}

function semaforosDe_(hoy, antes, u, moneda) {
  const luces = [];

  // 1 · CPA por entrega contra el techo
  if (hoy.cpaPorEntrega === null || hoy.techo === null) {
    luces.push(luz_('sin_medir', null, 'CPA por entrega',
      !hoy.hayPauta ? 'No hay gasto de pauta cargado en esta semana.'
                    : 'No hubo entregas, así que no hay contra qué dividir.'));
  } else {
    const pct = hoy.techo > 0 ? hoy.cpaPorEntrega / hoy.techo * 100 : 999;
    const e = pct <= u.cpa_verde_pct ? 'verde' : (pct <= 100 ? 'amarillo' : 'rojo');
    luces.push(luz_(e, hoy.cpaPorEntrega, 'CPA por entrega',
      'El techo es ' + redondear_(hoy.techo) + ' ' + moneda + '. ' +
      (e === 'rojo' ? 'Está por encima: cada entrega pierde plata.'
        : 'Queda ' + redondear_(hoy.techo - hoy.cpaPorEntrega) + ' de margen.')));
  }

  // 2 · Ticket promedio
  if (hoy.ticket === null) {
    luces.push(luz_('sin_medir', null, 'Ticket promedio', 'No hubo entregas esta semana.'));
  } else if (!u.ticket_minimo) {
    luces.push(luz_('sin_umbral', hoy.ticket, 'Ticket promedio',
      'No hay un mínimo definido. Se pone en Configuración para que esto se pueda juzgar.',
      antes && antes.ticket !== null ? hoy.ticket - antes.ticket : null));
  } else {
    const e = hoy.ticket >= u.ticket_minimo ? 'verde'
            : hoy.ticket >= u.ticket_minimo * 0.9 ? 'amarillo' : 'rojo';
    luces.push(luz_(e, hoy.ticket, 'Ticket promedio',
      'El mínimo que definiste es ' + redondear_(u.ticket_minimo) + ' ' + moneda + '.',
      antes && antes.ticket !== null ? hoy.ticket - antes.ticket : null));
  }

  // 3 · Tasa de entrega
  if (hoy.tasaEntrega === null) {
    luces.push(luz_('sin_medir', null, 'Entrega sobre resuelto',
      'Ningún pedido de la semana llegó todavía a entregado ni a devolución.'));
  } else {
    const resueltos = hoy.entregados + hoy.devoluciones;
    const chica = resueltos < u.muestra_minima;
    const e = chica ? 'muestra_chica'
      : hoy.tasaEntrega >= u.entrega_minima ? 'verde'
      : hoy.tasaEntrega >= u.entrega_minima - 8 ? 'amarillo' : 'rojo';
    luces.push(luz_(e, hoy.tasaEntrega, 'Entrega sobre resuelto',
      hoy.entregados + ' de ' + resueltos + ' resueltos' +
      (hoy.enTransito ? ' · ' + hoy.enTransito +
        (hoy.enTransito === 1 ? ' sigue' : ' siguen') + ' en tránsito, fuera de esta cuenta' : '') +
      (chica ? '. Con menos de ' + u.muestra_minima + ' resueltos esta cifra no es una señal.' : ''),
      antes && antes.tasaEntrega !== null ? hoy.tasaEntrega - antes.tasaEntrega : null));
  }

  // 4 · Utilidad
  if (!hoy.hayPauta) {
    luces.push(luz_('incompleto', hoy.utilidadAntesPauta, 'Utilidad de la semana',
      'Es ANTES de pauta. Sin el gasto de anuncios no se puede cerrar.',
      antes ? hoy.utilidadAntesPauta - antes.utilidadAntesPauta : null));
  } else if (hoy.gastoSinConvertir > 0) {
    luces.push(luz_('incompleto', hoy.utilidadReal, 'Utilidad de la semana',
      redondear_(hoy.gastoSinConvertir) + ' de pauta quedaron sin convertir por falta ' +
      'de tasa de cambio, así que no están restados aquí.',
      antes && antes.utilidadReal !== null ? hoy.utilidadReal - antes.utilidadReal : null));
  } else {
    const e = hoy.utilidadReal > 0
      ? (antes && antes.utilidadReal !== null && hoy.utilidadReal < antes.utilidadReal
          ? 'amarillo' : 'verde')
      : 'rojo';
    luces.push(luz_(e, hoy.utilidadReal, 'Utilidad de la semana',
      e === 'rojo' ? 'La semana cerró en pérdida.'
        : 'Después de producto, flete, devoluciones y pauta.',
      antes && antes.utilidadReal !== null ? hoy.utilidadReal - antes.utilidadReal : null));
  }

  return luces;
}

function redondear_(n) {
  if (n === null || n === undefined || isNaN(n)) return '—';
  return Math.round(Number(n) * 100) / 100;
}

// ─── LAS ALERTAS ─────────────────────────────────────────────

/**
 * Lo que hay que mirar, con qué hacer al respecto.
 *
 * Cada alerta trae una acción concreta. Una alerta sin acción solo
 * informa de que algo va mal, y quien la lee queda igual de atascado que
 * antes pero además preocupado.
 */
function alertasDe_(hoy, antes, u, moneda, cobertura) {
  const a = [];

  // Lo primero: ¿los datos están?
  if (cobertura.semanasVacias.length) {
    a.push({ nivel: 'rojo',
      titulo: 'Hay ' + cobertura.semanasVacias.length + ' semana(s) sin un solo pedido en la hoja',
      accion: 'Eso no significa que no vendiste: significa que la hoja no se ' +
        'actualizó. Baja de Dropi el reporte con rango completo y súbelo en ' +
        'Importar. Sin eso el semáforo no mide nada. Semanas vacías: ' +
        cobertura.semanasVacias.join(', ') + '.' });
  }
  if (hoy.diasConDatos > 0 && hoy.diasConDatos < 7) {
    a.push({ nivel: 'amarillo',
      titulo: 'La semana tiene ' + hoy.diasConDatos + ' de 7 días con registros',
      accion: 'Todo lo de abajo está calculado sobre esos días. No se extrapoló ' +
        'nada, así que las cifras son parciales, no bajas.' });
  }
  if (hoy.sinClasificar) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.sinClasificar + ' pedido(s) con un estado que Nova no reconoce',
      accion: 'No entran en ninguna cuenta: ni entregados, ni devueltos, ni ' +
        'despachados. Se clasifican en Configuración → Estados, una sola vez ' +
        'por estado nuevo.' });
  }
  if (hoy.sinCosto) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.sinCosto + ' pedido(s) entregados no tienen costo de producto',
      accion: 'Su ganancia quedó fuera de la utilidad. Se llena en Inventario, ' +
        'y desde ahí cuenta para todos los cierres.' });
  }
  if (hoy.gastoSinConvertir > 0) {
    a.push({ nivel: 'rojo',
      titulo: redondear_(hoy.gastoSinConvertir) + ' de pauta sin convertir',
      accion: 'Meta cobra en otra moneda y faltan las tasas de esos días. Ese ' +
        'gasto NO está restado de la utilidad. Se arregla en Nova Central, ' +
        'con «Prender lo automático».' });
  }

  // Y después: ¿cómo va el negocio?
  if (hoy.tasaEntrega !== null &&
      (hoy.entregados + hoy.devoluciones) >= u.muestra_minima &&
      hoy.tasaEntrega < u.entrega_minima) {
    a.push({ nivel: 'rojo',
      titulo: 'Entrega en ' + redondear_(hoy.tasaEntrega) + '%, bajo tu mínimo de ' +
        u.entrega_minima + '%',
      accion: 'Cada punto de entrega que se cae baja el techo de CPA. Mira de qué ' +
        'ciudades y de qué anuncios vienen las devoluciones antes de tocar el ' +
        'presupuesto.' });
  }
  if (hoy.cpaPorEntrega !== null && hoy.techo !== null && hoy.cpaPorEntrega > hoy.techo) {
    a.push({ nivel: 'rojo',
      titulo: 'El CPA por entrega (' + redondear_(hoy.cpaPorEntrega) + ') está sobre el techo (' +
        redondear_(hoy.techo) + ')',
      accion: 'Cada entrega de esta semana costó más de lo que deja. O sube el ' +
        'ticket, o baja el costo por entrega, o se apaga lo que menos entrega.' });
  }
  if (antes && hoy.pedidos && antes.pedidos && hoy.pedidos < antes.pedidos * 0.6) {
    a.push({ nivel: 'amarillo',
      titulo: 'Los pedidos cayeron de ' + antes.pedidos + ' a ' + hoy.pedidos,
      accion: 'Confirma si bajaste el presupuesto o si se apagó algo solo. Una ' +
        'caída de este tamaño sin decisión detrás suele ser un anuncio rechazado.' });
  }
  if (hoy.pedidos && hoy.cancelados / hoy.pedidos > 0.2) {
    a.push({ nivel: 'amarillo',
      titulo: hoy.cancelados + ' de ' + hoy.pedidos + ' pedidos cancelados (' +
        Math.round(hoy.cancelados / hoy.pedidos * 100) + '%)',
      accion: 'No cuestan flete, pero es 1 de cada ' +
        Math.round(hoy.pedidos / Math.max(hoy.cancelados, 1)) +
        ' leads pagados que nunca llega a despacho.' });
  }

  if (!a.length) {
    a.push({ nivel: 'ninguno', titulo: 'Nada que reportar esta semana',
      accion: 'Los números están dentro de lo que definiste. Es una respuesta ' +
        'válida y conviene que exista: si el semáforo siempre encontrara algo, ' +
        'dejaría de significar algo cuando lo encuentre.' });
  }
  return a;
}

// ─── COBERTURA ───────────────────────────────────────────────

/**
 * Qué días tiene la hoja, semana por semana.
 *
 * Va antes que cualquier número porque decide si los números valen. Una
 * semana entera ausente no es una semana mala: es una semana que no se
 * importó, y confundirlas lleva a apagar una campaña que estaba
 * funcionando.
 */
function coberturaDe_(pedidos, lunes, semanas) {
  const porDia = {};
  pedidos.forEach(function (p) { porDia[p.fecha] = (porDia[p.fecha] || 0) + 1; });

  const filas = [];
  const vacias = [];
  for (let s = semanas - 1; s >= 0; s--) {
    const ini = masDias_(lunes, -7 * s);
    const dias = [];
    let total = 0;
    for (let d = 0; d < 7; d++) {
      const f = masDias_(ini, d);
      const n = porDia[f] || 0;
      total += n;
      dias.push({ fecha: f, pedidos: n });
    }
    filas.push({ lunes: ini, dias: dias, total: total });
    if (!total) vacias.push(ini);
  }
  return { filas: filas, semanasVacias: vacias };
}

// ─── LO QUE SE ARMA Y SE ENTREGA ─────────────────────────────

/**
 * El semáforo completo de una tienda.
 *
 * @param {string} sheetId  hoja del cliente
 * @param {string} tienda   id de la tienda
 * @param {string} lunes    lunes de la semana a analizar; vacío = la última cerrada
 */
function semaforoSemanal(sheetId, tienda, lunes) {
  const ss = SpreadsheetApp.openById(sheetId);
  const hoyISO = ahoraISO().slice(0, 10);
  const L = lunes || semanaCerrada_(hoyISO);
  const SEMANAS = 8;

  const desde = masDias_(L, -7 * (SEMANAS - 1));
  const hasta = masDias_(L, 6);
  const u = ajustesSemaforo(ss, tienda);
  const d = datosParaSemaforo_(ss, tienda, desde, hasta);

  const hoy = numerosSemana_(d.pedidos, d.pauta, L);
  const antes = numerosSemana_(d.pedidos, d.pauta, masDias_(L, -7));

  // El promedio de las 4 semanas previas: por semana de CALENDARIO, no
  // por semana con datos. Dividir entre las que tienen datos esconde
  // justamente las que faltan.
  const previas = [];
  for (let k = 2; k <= 5; k++) previas.push(numerosSemana_(d.pedidos, d.pauta, masDias_(L, -7 * k)));
  const prom = {};
  ['pedidos', 'entregados', 'devoluciones', 'cancelados', 'ventas',
   'utilidadAntesPauta', 'gasto'].forEach(function (k) {
    prom[k] = previas.reduce(function (a, x) { return a + (x[k] || 0); }, 0) / previas.length;
  });

  const cobertura = coberturaDe_(d.pedidos, L, SEMANAS);

  return {
    tienda: tienda, moneda: d.moneda,
    semana: { lunes: L, domingo: hasta },
    generado: ahoraISO(),
    umbrales: u,
    hoy: hoy, anterior: antes, promedio4: prom,
    luces: semaforosDe_(hoy, antes, u, d.moneda),
    alertas: alertasDe_(hoy, antes, u, d.moneda, cobertura),
    cobertura: cobertura,
    productos: productosOrdenados_(hoy, u),
    hayDatos: hoy.pedidos > 0,
  };
}

/** Los productos de la semana, del que más pedidos trajo al que menos. */
function productosOrdenados_(hoy, u) {
  return Object.keys(hoy.productos).map(function (nombre) {
    const p = hoy.productos[nombre];
    const resueltos = p.entregados + p.devoluciones;
    return {
      producto: nombre, pedidos: p.pedidos,
      entregados: p.entregados, devoluciones: p.devoluciones,
      tasaEntrega: resueltos ? p.entregados / resueltos * 100 : null,
      gananciaPorPedido: p.pedidos ? p.ganancia / p.pedidos : null,
      // Por debajo de la muestra mínima no se opina: se dice cuánto falta.
      señal: resueltos >= u.muestra_minima ? 'medible'
             : 'faltan ' + (u.muestra_minima - resueltos) + ' resueltos para que signifique algo',
    };
  }).sort(function (a, b) { return b.pedidos - a.pedidos; });
}

// ─── LA API Y EL CORREO ──────────────────────────────────────

/**
 * Quién puede ver esto.
 *
 * Hacen falta las dos cosas: que el PLAN del cliente incluya pauta y
 * dinero —eso lo decide Nova Central, o sea Manuela— y que la PERSONA
 * tenga el permiso dentro de su equipo —eso lo decide su dueña.
 *
 * Son preguntas distintas y por eso se responden por separado. Mezclarlas
 * llevaría a que darle un permiso a una gestora le abriera un módulo que
 * su empresa no paga, o a que pagar un módulo se lo mostrara a todo el
 * equipo.
 */
function puedeVerSemaforo(s) {
  const modulos = s.modulos || [];
  const permisos = s.permisos || [];
  if (modulos.indexOf('pauta') === -1 || modulos.indexOf('dinero') === -1) {
    return { puede: false, porque: 'plan',
      error: 'El semáforo cruza la pauta con la plata, y tu plan no incluye ' +
             'los dos módulos. Se activa desde Nova Central.' };
  }
  if (s.rol !== 'dueno' && permisos.indexOf('ver_dinero') === -1) {
    return { puede: false, porque: 'permiso',
      error: 'El semáforo muestra utilidad y gasto. Tu dueña decide quién los ve.' };
  }
  return { puede: true };
}

function apiSemaforo(s, p) {
  const v = puedeVerSemaforo(s);
  if (!v.puede) return { ok: false, error: v.error, porque: v.porque };

  const tienda = String(p.tienda || s.tiendas[0]);
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  return { ok: true, semaforo: semaforoSemanal(s.sheetId, tienda, String(p.lunes || '')) };
}

/**
 * El correo del lunes, en texto.
 *
 * Texto plano y no HTML porque esto se lee en el teléfono a las siete de
 * la mañana, y lo que importa es que las cuatro luces se vean en la
 * primera pantalla sin cargar nada.
 */
function semaforoTexto(sem) {
  const m = sem.moneda ? ' ' + sem.moneda : '';
  const L = [];
  L.push('SEMÁFORO · ' + sem.tienda);
  L.push(sem.semana.lunes + ' a ' + sem.semana.domingo);
  L.push('');

  if (!sem.hayDatos) {
    L.push('No hay ni un pedido de esa semana en la hoja.');
    L.push('');
    L.push('Eso no significa que no se vendió:');
    L.push('significa que la hoja no se actualizó.');
    L.push('');
    L.push('Sube el reporte de Dropi con el rango completo y vuelve a mirarlo.');
    L.push('Mientras no esté, este semáforo no mide nada — y prefiero decirlo');
    L.push('a mandarte unos ceros que parecen una mala semana.');
    return L.join('\n');
  }

  const icono = { verde: '🟢', amarillo: '🟡', rojo: '🔴',
                  sin_medir: '⚪', incompleto: '🟡', muestra_chica: '⚪',
                  sin_umbral: '⚪' };
  sem.luces.forEach(function (l) {
    L.push((icono[l.estado] || '⚪') + '  ' + l.etiqueta + ': ' +
           (l.valor === null ? 'sin dato' : redondear_(l.valor)));
    if (l.delta !== null && l.delta !== undefined) {
      L.push('    ' + (l.delta >= 0 ? '▲ +' : '▼ ') + redondear_(l.delta) +
             ' vs. la semana anterior');
    }
    if (l.porque) L.push('    ' + l.porque);
    L.push('');
  });

  L.push('');
  L.push('LA SEMANA');
  L.push('  Pedidos creados      ' + sem.hoy.pedidos);
  L.push('  Entregados           ' + sem.hoy.entregados);
  L.push('  Devoluciones         ' + sem.hoy.devoluciones);
  L.push('  Cancelados           ' + sem.hoy.cancelados);
  L.push('  En tránsito          ' + sem.hoy.enTransito);
  L.push('  Utilidad antes pauta ' + redondear_(sem.hoy.utilidadAntesPauta) + m);
  L.push('  Pauta                ' + (sem.hoy.hayPauta ? redondear_(sem.hoy.gasto) + m : 'sin dato'));
  L.push('  Utilidad real        ' + (sem.hoy.utilidadReal === null ? 'sin dato'
                                      : redondear_(sem.hoy.utilidadReal) + m));
  L.push('  Techo de CPA         ' + (sem.hoy.techo === null ? 'sin dato'
                                      : redondear_(sem.hoy.techo) + m));

  L.push('');
  L.push('QUÉ MIRAR');
  sem.alertas.forEach(function (a, i) {
    L.push('  ' + (i + 1) + '. ' + a.titulo);
    L.push('     ' + a.accion);
  });

  L.push('');
  L.push('Cómo se calculó: solo los entregados generan ganancia; cada devolución');
  L.push('cuesta el flete; los cancelados cuestan cero. La tasa de entrega es');
  L.push('entregados ÷ (entregados + devoluciones). El techo de CPA es la');
  L.push('utilidad antes de pauta ÷ pedidos creados. Donde no hay dato, dice');
  L.push('sin dato: nada está extrapolado.');
  return L.join('\n');
}

/**
 * Lo que corre el lunes: un semáforo por tienda, a quien pueda verlo.
 *
 * Se le manda solo a la dueña. El semáforo lleva utilidad y gasto, y
 * quién más los ve es una decisión de ella, no de Nova.
 */
function semaforoLunes() {
  const log = [];
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try {
      const ss = SpreadsheetApp.openById(c.sheetId);
      const modulos = modulosDeCliente_(c.sheetId);
      if (modulos.indexOf('pauta') === -1 || modulos.indexOf('dinero') === -1) {
        log.push(c.empresa + ': el plan no incluye pauta y dinero. Saltado.');
        return;
      }
      duenasDe_(ss).forEach(function (due) {
        (due.tiendas === '*' ? tiendasActivas_(ss) : [due.tiendas]).forEach(function (t) {
          const sem = semaforoSemanal(c.sheetId, t, '');
          try {
            MailApp.sendEmail({
              to: due.correo,
              subject: 'Semáforo de la semana · ' + t + ' · ' + sem.semana.lunes,
              body: semaforoTexto(sem),
            });
            log.push(c.empresa + ' · ' + t + ' → ' + due.correo);
          } catch (e) {
            log.push(c.empresa + ' · ' + t + ': no se pudo enviar — ' + e.message);
          }
        });
      });
    } catch (e) {
      log.push(c.empresa + ': FALLÓ — ' + e.message);
    }
  });
  const msg = log.length ? log.join('\n') : 'Ningún cliente con plan de pauta y dinero.';
  Logger.log(msg);
  return msg;
}

/** Las dueñas activas de una cuenta, con sus tiendas. */
function duenasDe_(ss) {
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (rolCanonico(d[i][c('rol')]) !== 'dueno') continue;
    if (norm(d[i][c('estado')]) === 'inactivo') continue;
    const correo = String(d[i][c('correo')] || '').trim();
    if (!correo) continue;
    out.push({ correo: correo, tiendas: String(d[i][c('tienda')] || '*').trim() || '*' });
  }
  return out;
}

function tiendasActivas_(ss) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cEst = e.indexOf('estado');
  const out = [];
  for (let i = 1; i < d.length; i++) {
    const id = String(d[i][cId] || '').trim();
    if (!id) continue;
    if (cEst !== -1 && norm(d[i][cEst]) === 'inactiva') continue;
    out.push(id);
  }
  return out;
}

/**
 * Qué módulos incluye el plan de un cliente, por su hoja.
 *
 * Se apoya en `modulosDelPlan`, que es la misma función que usa el login:
 * si fueran dos, un cliente podría ver el semáforo por correo y no en su
 * pantalla, o al revés, y nadie sabría cuál de las dos tiene razón.
 */
function modulosDeCliente_(sheetId) {
  try {
    const shC = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
    if (!shC || shC.getLastRow() < 2) return [];
    const d = shC.getDataRange().getValues();
    const e = d[0].map(norm);
    const cSheet = e.indexOf('sheet_id'), cPlan = e.indexOf('plan');
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cSheet] || '').trim() === sheetId) {
        return modulosDelPlan(d[i][cPlan]);
      }
    }
    return [];
  } catch (e) {
    return [];
  }
}
