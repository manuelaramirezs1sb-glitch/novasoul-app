/**
 * Nova · Alarmas
 * ─────────────────────────────────────────────────────────────
 * Las seis cosas que no pueden esperar a que alguien abra la app.
 *
 * ┌─ POR QUÉ LOS UMBRALES LOS PONE EL CLIENTE ─────────────────┐
 * │                                                            │
 * │ Una tienda que sabe que vive con 40% de devolución no      │
 * │ necesita que le griten todos los días por eso. Otra que    │
 * │ vende un producto de 200 dólares se hunde con un 12%.      │
 * │                                                            │
 * │ Un umbral inventado por mí produce una de dos cosas: una   │
 * │ alarma que suena siempre —y que por eso se ignora— o una   │
 * │ que nunca suena. Las dos son igual de inútiles.            │
 * │                                                            │
 * │ Así que cada umbral vive en la hoja Parametros, por        │
 * │ tienda, y el dueño lo cambia desde la app. Los valores de  │
 * │ abajo son solo el punto de partida.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * El CPA es la excepción, y a propósito: su techo NO se escribe, se
 * calcula. Es lo que deja cada entrega después del producto y el flete,
 * porque pagar más que eso por conseguir un pedido es perder plata en
 * cada venta. Ese número cambia solo cuando cambian los costos.
 */

const ALARMAS_DEFAULT = {
  dias_sin_mover:     3,     // pedido sin movimiento
  horas_novedad:      24,    // novedad sin gestionar
  efectividad_min:    65,    // % de entrega sobre lo resuelto
  devoluciones_max:   '',    // % — lo pone el cliente; vacío = apagada
  cpa_aviso_pct:      85,    // % del techo a partir del cual avisa
  cpa_subida_pct:     25,    // % de subida contra el mes pasado que avisa
  stock_dias_min:     '',    // días de cobertura — vacío = apagada
  alarmas_a:          '',    // correos extra, separados por coma
  alarmas_hora:       7,     // hora local de la revisión diaria
};

/** Qué es cada alarma, en palabras de quien la va a leer. */
const ALARMAS = [
  { id: 'sin_mover',    nombre: 'Pedidos detenidos',
    param: 'dias_sin_mover', unidad: 'días' },
  { id: 'novedad_vieja', nombre: 'Novedades sin gestionar',
    param: 'horas_novedad', unidad: 'horas' },
  { id: 'efectividad',  nombre: 'Efectividad baja',
    param: 'efectividad_min', unidad: '%' },
  { id: 'devoluciones', nombre: 'Devoluciones altas',
    param: 'devoluciones_max', unidad: '%', opcional: true },
  { id: 'cpa',          nombre: 'CPA cerca del techo',
    param: 'cpa_aviso_pct', unidad: '% del techo' },
  { id: 'cpa_sube',     nombre: 'CPA subiendo',
    param: 'cpa_subida_pct', unidad: '% vs. mes pasado', opcional: true },
  { id: 'stock',        nombre: 'Stock por agotarse',
    param: 'stock_dias_min', unidad: 'días de cobertura', opcional: true },
];

/** Los umbrales de una tienda: lo que diga Parametros, o el de fábrica. */
function umbrales(ss, tienda) {
  const out = {};
  Object.keys(ALARMAS_DEFAULT).forEach(function (k) { out[k] = ALARMAS_DEFAULT[k]; });

  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda'), cK = e.indexOf('clave'), cV = e.indexOf('valor');
  for (let i = 1; i < d.length; i++) {
    const t = String(d[i][cT] || '').trim();
    // Un parámetro sin tienda vale para todas: sirve de valor general
    if (t && t !== tienda) continue;
    const k = norm(d[i][cK]);
    if (!(k in out)) continue;
    const v = d[i][cV];
    out[k] = (v === '' || v === null) ? '' : v;
  }
  return out;
}

/**
 * Evalúa las seis alarmas de una tienda.
 *
 * Devuelve una lista, no manda correos. Separarlo permite que la pantalla
 * las muestre en vivo y que el correo diario use exactamente lo mismo:
 * si fueran dos cálculos distintos, tarde o temprano dirían cosas
 * distintas y no habría forma de saber cuál creer.
 */
function evaluarAlarmas(ss, tienda) {
  const u = umbrales(ss, tienda);
  const tz = zonaHorariaDe(ss, tienda);
  const hoy = new Date();
  const mes = Utilities.formatDate(hoy, tz || 'UTC', 'yyyy-MM');
  const out = [];
  const moneda = monedaDeTienda(ss, tienda);

  // ── Datos del mes, una sola lectura ──
  const sesionFalsa = { rol: 'dueno' };
  const m = agregarMes(ss, tienda, mes, sesionFalsa);

  // ── 1. Pedidos detenidos ──
  const dias = Number(u.dias_sin_mover) || 0;
  if (dias > 0) {
    const detenidos = [];
    const shP = ss.getSheetByName('Pedidos');
    if (shP && shP.getLastRow() > 1) {
      const d = shP.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        const est = norm(f[c('estado_nova')] || f[c('estado_canonico')]);
        if (['entregado','devolucion','cancelado'].indexOf(est) !== -1) continue;
        const ult = aISO(f[c('ultimo_movimiento')] || f[c('actualizado_en')] ||
                         f[c('fecha')], 'UTC');
        if (!ult) continue;
        const d2 = (hoy - new Date(ult + 'T00:00:00Z')) / 86400000;
        if (d2 >= dias) {
          detenidos.push({ id: f[c('id_externo')] || f[c('id')],
                           cliente: f[c('cliente')], dias: Math.floor(d2),
                           gestora: f[c('gestora_asignada')] });
        }
      }
    }
    if (detenidos.length) {
      detenidos.sort(function (a, b) { return b.dias - a.dias; });
      out.push(alarma('sin_mover', 'mal',
        pl(detenidos.length, '1 pedido lleva', '% pedidos llevan') + ' ' +
          dias + ' días o más sin moverse',
        'El más viejo lleva ' + detenidos[0].dias + ' días. Un pedido detenido ' +
        'no avisa solo: o se gestiona, o se convierte en devolución.',
        detenidos.slice(0, 10), 'Pedidos'));
    }
  }

  // ── 2. Novedades sin gestionar ──
  const horas = Number(u.horas_novedad) || 0;
  if (horas > 0) {
    const viejas = [];
    const shN = ss.getSheetByName('Novedades');
    if (shN && shN.getLastRow() > 1) {
      const d = shN.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (norm(f[c('estado')]) !== 'abierta') continue;
        const fch = aISO(f[c('fecha')], 'UTC');
        if (!fch) continue;
        const h = (hoy - new Date(fch + 'T00:00:00Z')) / 3600000;
        if (h >= horas) {
          viejas.push({ id: f[c('pedido_id')], motivo: f[c('motivo')],
                        horas: Math.floor(h), gestora: f[c('gestora')] });
        }
      }
    }
    if (viejas.length) {
      viejas.sort(function (a, b) { return b.horas - a.horas; });
      out.push(alarma('novedad_vieja', 'mal',
        pl(viejas.length, '1 novedad lleva', '% novedades llevan') +
          ' más de ' + horas + ' horas abierta' + (viejas.length === 1 ? '' : 's'),
        'Una novedad sin contactar a la clienta en el primer día se vuelve ' +
        'devolución en la mayoría de los casos.',
        viejas.slice(0, 10), 'Novedades'));
    }
  }

  // ── 3. Efectividad baja ──
  const efMin = Number(u.efectividad_min) || 0;
  if (efMin > 0 && m.resueltos >= 10 && m.efectividad < efMin) {
    out.push(alarma('efectividad', 'mal',
      'Efectividad en ' + m.efectividad.toFixed(1) + '%, bajo tu meta de ' + efMin + '%',
      'De ' + m.resueltos + ' pedidos ya resueltos este mes llegaron ' +
      m.entregados + '. Cada punto por debajo son ventas que ya pagaste en pauta ' +
      'y no entraron.', [], 'Pedidos'));
  }

  // ── 4. Devoluciones altas (opcional) ──
  const devMax = u.devoluciones_max === '' ? null : Number(u.devoluciones_max);
  if (devMax !== null && devMax > 0 && m.resueltos >= 10) {
    const tasa = m.resueltos ? m.devueltos / m.resueltos * 100 : 0;
    if (tasa > devMax) {
      out.push(alarma('devoluciones', 'mal',
        'Devoluciones en ' + tasa.toFixed(1) + '%, sobre tu límite de ' + devMax + '%',
        m.devueltos + ' de ' + m.resueltos + ' pedidos resueltos volvieron. ' +
        'Cada uno cuesta el flete de ida y el de vuelta.', [], 'Novedades'));
    }
  }

  /**
   * ── 5. CPA cerca del techo ──
   *
   * El techo no se escribe en ninguna parte: es lo que deja cada entrega
   * después del producto y el flete. Pagar más que eso por traer un
   * pedido es perder plata en cada venta, por bien que se vea el ROAS.
   */
  /**
   * Diez entregas como mínimo, y aquí está el porqué.
   *
   * A principios de mes la pauta ya se gastó y los pedidos todavía no
   * llegan: el CPA se calcula sobre dos o tres entregas y sale disparado.
   * El 12 de septiembre esta alarma decía "pagas USD 86,41 por pedido"
   * cuando el mes cerrado anterior iba en 17,93. No era un problema de
   * la operación, era un mes que apenas empezaba.
   *
   * Una alarma que grita cada primero de mes es una alarma que se ignora
   * el resto del mes.
   */
  if (m.entregados >= 10 && m.gasto > 0) {
    const techo = (m.ventas - m.costoProducto - m.costoEnvio) / m.entregados;
    const pagado = m.gasto / m.entregados;

    /**
     * El techo dice si estás perdiendo. Esta otra dice si estás
     * empeorando, que es la que avisa a tiempo: un CPA que sube 30% en un
     * mes todavía puede estar bajo el techo, y aun así ser la señal de
     * que la campaña se está agotando o la competencia subió la puja.
     *
     * Se compara con el mes pasado completo, no con el promedio: un
     * promedio de varios meses suaviza justo lo que hay que ver.
     */
    const sub = u.cpa_subida_pct === '' ? null : Number(u.cpa_subida_pct);
    if (sub !== null && sub > 0) {
      const ant = agregarMes(ss, tienda, mesAnterior(mes), sesionFalsa);
      const cpaAnt = ant.entregados ? ant.gasto / ant.entregados : 0;
      // También el mes pasado necesita volumen: comparar contra un mes
      // de tres entregas produce porcentajes enormes que no dicen nada.
      if (cpaAnt > 0 && ant.entregados >= 10) {
        const delta = (pagado - cpaAnt) / cpaAnt * 100;
        if (delta >= sub) {
          out.push(alarma('cpa_sube', 'ojo',
            'El CPA subió ' + delta.toFixed(0) + '% contra el mes pasado',
            'Pagabas ' + moneda + ' ' + cpaAnt.toFixed(2) + ' por pedido y ahora ' +
            'pagas ' + moneda + ' ' + pagado.toFixed(2) + '. Todavía ' +
            (pagado < techo ? 'estás bajo el techo, pero la tendencia se come el colchón.'
                            : 'y además ya pasaste el techo.'),
            [], 'Dinero'));
        }
      }
    }
    const pct = techo > 0 ? pagado / techo * 100 : 999;
    const aviso = Number(u.cpa_aviso_pct) || 85;
    if (techo > 0 && pct >= aviso) {
      const grave = pct >= 100;
      out.push(alarma('cpa', grave ? 'mal' : 'ojo',
        grave
          ? 'Estás pagando más por pedido de lo que deja cada entrega'
          : 'CPA al ' + pct.toFixed(0) + '% del techo',
        'Cada entrega deja ' + moneda + ' ' + techo.toFixed(2) + ' después del ' +
        'producto y el flete, y estás pagando ' + moneda + ' ' + pagado.toFixed(2) +
        ' de pauta por conseguirla. ' +
        (grave ? 'Así, vender más es perder más.'
               : 'Queda poco colchón: si la entrega cae unos puntos, el mes se pone rojo.'),
        [], 'Dinero'));
    }
  }

  // ── 6. Stock por agotarse (opcional) ──
  const stockDias = u.stock_dias_min === '' ? null : Number(u.stock_dias_min);
  if (stockDias !== null && stockDias > 0) {
    const bajos = [];
    const shI = ss.getSheetByName('Inventario');
    if (shI && shI.getLastRow() > 1) {
      const d = shI.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      // Ritmo de venta de cada producto en el mes, para estimar cobertura
      const ritmo = {};
      Object.keys(m.productos || {}).forEach(function (k) {
        const dm = new Date(hoy.getFullYear(), hoy.getMonth() + 1, 0).getDate();
        ritmo[norm(k)] = (m.productos[k].entregados || 0) / dm;
      });
      for (let i = 1; i < d.length; i++) {
        const f = d[i];
        if (String(f[c('tienda')]).trim() !== tienda) continue;
        if (norm(f[c('activo')]) === 'no') continue;
        const nom = String(f[c('producto')] || '').trim();
        const stock = num(f[c('stock')]);
        const r = ritmo[norm(nom)] || 0;
        if (!r) continue;              // sin ventas no hay cobertura que estimar
        const cobertura = stock / r;
        if (cobertura <= stockDias) {
          bajos.push({ producto: nom, stock: stock, dias: Math.floor(cobertura) });
        }
      }
    }
    if (bajos.length) {
      bajos.sort(function (a, b) { return a.dias - b.dias; });
      out.push(alarma('stock', 'ojo',
        pl(bajos.length, '1 producto se acaba', '% productos se acaban') +
          ' en ' + stockDias + ' días o menos',
        'Al ritmo de venta de este mes. Quedarse sin stock con la pauta ' +
        'prendida es pagar por pedidos que no puedes despachar.',
        bajos.slice(0, 10), 'Inventario'));
    }
  }

  return { alarmas: out, umbrales: u, tienda: tienda, mes: mes };
}

/** Uno o varios. "1 novedades" delata que lo escribió una máquina. */
function pl(n, uno, varios) {
  return n === 1 ? uno.replace('%', n) : varios.replace('%', n);
}

function alarma(id, nivel, titulo, detalle, casos, ir) {
  const def = ALARMAS.filter(function (a) { return a.id === id; })[0] || {};
  return { id: id, nivel: nivel, nombre: def.nombre || id,
           titulo: titulo, detalle: detalle, casos: casos || [], ir: ir || '' };
}

/**
 * La revisión diaria que manda el correo.
 *
 * Manda UNA vez por alarma y por día. Sin eso, una efectividad baja que
 * dura toda la semana produce siete correos idénticos, y al tercero ya
 * nadie los abre — que es exactamente cuando deja de servir.
 */
function revisarAlarmas(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const tiendas = tiendasDeCliente(ss);
  const log = [];

  tiendas.forEach(function (tienda) {
    const r = evaluarAlarmas(ss, tienda);
    if (!r.alarmas.length) { log.push(tienda + ': sin alarmas'); return; }

    const nuevas = r.alarmas.filter(function (a) {
      return !yaAvisada(ss, a.id + '|' + tienda);
    });
    if (!nuevas.length) {
      log.push(tienda + ': ' + r.alarmas.length + ' alarmas, ya avisadas hoy');
      return;
    }

    const destinos = destinatarios(ss, r.umbrales);
    if (destinos.length) {
      MailApp.sendEmail({
        to: destinos.join(','),
        subject: 'Nova · ' + nuevas.length + ' cosas que mirar en ' + nombreTienda(ss, tienda),
        body: cuerpoCorreo(nuevas, tienda, nombreTienda(ss, tienda)),
      });
    }
    nuevas.forEach(function (a) { marcarAvisada(ss, a.id + '|' + tienda); });
    log.push(tienda + ': avisadas ' + nuevas.length + ' a ' + destinos.join(', '));
  });

  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}

/** A quién le llega: las dueñas activas, más los correos que se agreguen. */
function destinatarios(ss, u) {
  const out = [];
  const sh = ss.getSheetByName('Equipo');
  if (sh && sh.getLastRow() > 1) {
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const cC = e.indexOf('correo'), cR = e.indexOf('rol'), cE = e.indexOf('estado');
    for (let i = 1; i < d.length; i++) {
      if (norm(d[i][cE]) === 'inactivo') continue;
      if (rolCanonico(d[i][cR]) !== 'dueno') continue;
      const c = String(d[i][cC] || '').trim();
      if (c) out.push(c);
    }
  }
  String(u.alarmas_a || '').split(/[,;]/).forEach(function (c) {
    const x = c.trim();
    if (x && out.indexOf(x) === -1) out.push(x);
  });
  return out;
}

function cuerpoCorreo(alarmas, tienda, nombre) {
  const lineas = ['Hola,', '',
    'Esto es lo que Nova encontró hoy en ' + nombre + ':', ''];
  alarmas.forEach(function (a, i) {
    lineas.push((i + 1) + '. ' + a.titulo);
    lineas.push('   ' + a.detalle);
    if (a.casos && a.casos.length) {
      a.casos.slice(0, 5).forEach(function (c) {
        lineas.push('   · ' + (c.cliente || c.producto || c.id || '') +
          (c.dias !== undefined ? ' — ' + c.dias + ' días' : '') +
          (c.horas !== undefined ? ' — ' + c.horas + ' horas' : ''));
      });
      if (a.casos.length > 5) lineas.push('   · y ' + (a.casos.length - 5) + ' más');
    }
    lineas.push('');
  });
  lineas.push('Los umbrales de estas alarmas los cambias tú en Nova, ' +
              'en Configuración.');
  lineas.push('');
  lineas.push('— Nova');
  return lineas.join('\n');
}

/** Una alarma avisada hoy no se vuelve a avisar hoy. */
function yaAvisada(ss, clave) {
  const sh = ss.getSheetByName('Alertas_enviadas');
  if (!sh || sh.getLastRow() < 2) return false;
  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() !== clave) continue;
    if (String(aISO(d[i][2], 'UTC')) === hoy) return true;
  }
  return false;
}

function marcarAvisada(ss, clave) {
  const sh = ss.getSheetByName('Alertas_enviadas');
  if (!sh) return;
  sh.appendRow([clave, '', ahoraISO(), '']);
}

function nombreTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return tienda;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cN = e.indexOf('nombre');
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][cId]).trim() === tienda) return String(d[i][cN] || tienda);
  }
  return tienda;
}

/**
 * Deja la revisión diaria corriendo sola.
 *
 * Es lo que hace que Nova avise sin que nadie abra nada. Sin esto, las
 * alarmas solo existen para quien ya está mirando la pantalla — que es
 * justo quien menos las necesita.
 */
function instalarTriggerAlarmas() {
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'revisarAlarmasTodos') {
      ScriptApp.deleteTrigger(t);
    }
  });
  const hora = Number(ALARMAS_DEFAULT.alarmas_hora) || 7;
  ScriptApp.newTrigger('revisarAlarmasTodos').timeBased().atHour(hora).everyDays(1).create();
  const msg = 'Revisión diaria de alarmas instalada para las ' + hora + ':00.';
  Logger.log(msg);
  return msg;
}

/** Recorre todos los clientes registrados. */
function revisarAlarmasTodos() {
  const central = SpreadsheetApp.openById(IDS_().central).getSheetByName('Clientes');
  if (!central || central.getLastRow() < 2) return 'Sin clientes.';
  const filas = central.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('sheet_id') !== -1 ? enc.indexOf('sheet_id') : 13;
  const log = [];
  for (let i = 1; i < filas.length; i++) {
    const id = String(filas[i][cId] || '').trim();
    if (!id) continue;
    try { log.push(revisarAlarmas(id)); }
    catch (e) { log.push('Cliente ' + id + ': ' + e.message); }
  }
  const msg = log.join('\n');
  Logger.log(msg);
  return msg;
}
