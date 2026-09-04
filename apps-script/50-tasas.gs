/**
 * Nova · Tasas automáticas y efecto cambiario
 * ─────────────────────────────────────────────────────────────
 * El problema que resuelve, en concreto:
 *
 *   Operas Nutrea EC en dólares, pero vives en Colombia y tu plata es
 *   en pesos. Si el dólar pasa de 4.000 a 3.700, tu tienda puede
 *   facturar exactamente lo mismo en USD y aun así tu utilidad en COP
 *   cae 7,5%. Nada pasó en la operación.
 *
 *   El riesgo no es perder esa plata: es no saber por qué la perdiste,
 *   ver caer el número y ponerte a optimizar pauta cuando el problema
 *   no está en la pauta.
 *
 * Por eso hay tres piezas:
 *   1. actualizarTasas()   trae la tasa sola, todos los días
 *   2. alarmaTasa()        avisa cuando el movimiento ya pesa
 *   3. efectoCambiario()   separa cuánto es operación y cuánto es cambio
 */

// Celda de trabajo para GOOGLEFINANCE. Se usa y se limpia.
const TAB_FX = '_fx_tmp';

// ─── QUÉ PARES HACEN FALTA ───────────────────────────────────

/**
 * Los pares de moneda que este cliente necesita: de la moneda de cada
 * tienda a la moneda en que reporta. Si operas GT en quetzales y EC en
 * dólares, y reportas en pesos, los pares son GTQ->COP y USD->COP.
 */
function paresEnUso(ss) {
  const destino = monedaReporte(ss);
  const shT = ss.getSheetByName('Tiendas');
  if (!shT || shT.getLastRow() < 2) return [];

  const filas = shT.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cMon = enc.indexOf('moneda');
  const cEstado = enc.indexOf('estado');

  const vistos = {};
  filas.slice(1).forEach(function (f) {
    if (cEstado !== -1 && norm(f[cEstado]) === 'inactiva') return;
    const m = String(f[cMon] || '').toUpperCase();
    if (m && m !== destino) vistos[m] = 1;
  });

  return Object.keys(vistos).map(function (o) {
    return { origen: o, destino: destino };
  });
}

/** La moneda en que el dueño ve su plata. Sale de Parametros. */
function monedaReporte(ss) {
  const sh = ss.getSheetByName('Parametros');
  if (sh && sh.getLastRow() > 1) {
    const filas = sh.getDataRange().getValues().slice(1);
    for (let i = 0; i < filas.length; i++) {
      if (norm(filas[i][1]) === 'moneda_reporte' && filas[i][2]) {
        return String(filas[i][2]).toUpperCase();
      }
    }
  }
  return 'COP';
}

// ─── TRAER LAS TASAS ─────────────────────────────────────────

/**
 * Rellena los huecos de la hoja `Tasas` con GOOGLEFINANCE.
 *
 * Se usa GOOGLEFINANCE y no una API externa por tres razones: no pide
 * llave, no tiene límite de llamadas, y no se cae un domingo dejando la
 * plataforma sin tasas. La contra es que solo vive dentro de Sheets, así
 * que hay que escribir la fórmula en una celda y leer el resultado.
 *
 * @param {string} cliente  nombre del cliente, o vacío si solo hay uno
 * @param {number} dias     cuántos días hacia atrás revisar (por defecto 90)
 */
function actualizarTasas(cliente, dias) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const pares = paresEnUso(ss);
  if (!pares.length) {
    Logger.log('No hay pares que actualizar: todas las tiendas reportan en la misma moneda.');
    return 'Sin pares que actualizar.';
  }

  const shTasas = ss.getSheetByName('Tasas');
  const hoy = new Date();
  const desde = new Date(hoy.getTime() - (dias || 90) * 86400000);

  // Qué fechas ya tenemos, para no repetir
  const existentes = {};
  if (shTasas.getLastRow() > 1) {
    shTasas.getDataRange().getValues().slice(1).forEach(function (f) {
      const k = aISO(f[0], 'UTC') + '|' + String(f[1]).toUpperCase() +
                '|' + String(f[2]).toUpperCase();
      existentes[k] = 1;
    });
  }

  let sh = ss.getSheetByName(TAB_FX);
  if (!sh) { sh = ss.insertSheet(TAB_FX); sh.hideSheet(); }

  const nuevas = [];
  const errores = [];

  pares.forEach(function (p) {
    const simbolo = 'CURRENCY:' + p.origen + p.destino;
    const formula = '=GOOGLEFINANCE("' + simbolo + '","price",DATE(' +
      desde.getFullYear() + ',' + (desde.getMonth() + 1) + ',' + desde.getDate() + '),DATE(' +
      hoy.getFullYear() + ',' + (hoy.getMonth() + 1) + ',' + hoy.getDate() + '),"DAILY")';

    sh.clear();
    sh.getRange(1, 1).setFormula(formula);
    SpreadsheetApp.flush();
    Utilities.sleep(1200); // GOOGLEFINANCE tarda en resolver

    const datos = sh.getDataRange().getValues();
    // La primera fila son encabezados (Date / Close); si vino un error,
    // la celda trae el texto del error en vez de la tabla.
    if (datos.length < 2) {
      errores.push(p.origen + '->' + p.destino + ': ' + (datos[0] ? datos[0][0] : 'sin datos'));
      return;
    }

    datos.slice(1).forEach(function (r) {
      const fecha = aISO(r[0], 'UTC');
      const tasa = aNumero(r[1]);
      if (!fecha || tasa === '' || tasa <= 0) return;
      const k = fecha + '|' + p.origen + '|' + p.destino;
      if (existentes[k]) return;
      existentes[k] = 1;
      nuevas.push([fecha, p.origen, p.destino, tasa]);
    });
  });

  sh.clear();

  if (nuevas.length) {
    nuevas.sort(function (a, b) { return a[0] < b[0] ? -1 : 1; });
    shTasas.getRange(shTasas.getLastRow() + 1, 1, nuevas.length, 4).setValues(nuevas);
  }

  const msg = [
    'Tasas agregadas: ' + nuevas.length,
    'Pares: ' + pares.map(function (p) { return p.origen + '->' + p.destino; }).join(', '),
    errores.length ? 'ERRORES:\n  ' + errores.join('\n  ') : '',
  ].filter(String).join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Deja la actualización corriendo sola, todos los días a las 6 a.m.
 * Córrelo una vez por cliente.
 */
function instalarTriggerTasas() {
  const ya = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'actualizarTasasDiario';
  });
  if (ya.length) { Logger.log('El trigger ya existía.'); return 'Ya existía.'; }

  ScriptApp.newTrigger('actualizarTasasDiario')
    .timeBased().atHour(6).everyDays(1).create();
  Logger.log('Trigger diario instalado (6 a.m.).');
  return 'Trigger diario instalado.';
}

/** Lo que corre el trigger: actualiza las tasas de todos los clientes. */
function actualizarTasasDiario() {
  const clientes = listarClientes();
  clientes.forEach(function (c) {
    if (!c.sheetId) return;
    try {
      actualizarTasas(c.sheetId, 10); // ventana corta: solo tapar huecos
    } catch (e) {
      Logger.log('Fallo actualizando ' + c.empresa + ': ' + e.message);
    }
  });
}

// ─── ALARMA DE MOVIMIENTO ────────────────────────────────────

/**
 * Avisa cuando la tasa se movió lo suficiente para que importe.
 *
 * Dos umbrales, porque son dos preocupaciones distintas:
 *   · corto plazo — se movió fuerte esta semana, revisa precios
 *   · desde el inicio — cuánto se corrió respecto a cuando arrancaste,
 *     que es el caso que planteaste del dólar a 4.000
 *
 * Umbrales configurables en Parametros:
 *   tasa_alerta_pct_30d   (por defecto 5)
 *   tasa_referencia_USD   (opcional: la tasa con la que montaste el negocio)
 */
function alarmaTasa(cliente) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const pares = paresEnUso(ss);
  const destino = monedaReporte(ss);
  const hoy = Utilities.formatDate(new Date(), 'UTC', 'yyyy-MM-dd');
  const hace30 = Utilities.formatDate(new Date(Date.now() - 30 * 86400000), 'UTC', 'yyyy-MM-dd');

  const umbral = Number(parametro(ss, 'tasa_alerta_pct_30d') || 5);
  const avisos = [];

  pares.forEach(function (p) {
    const ahora = buscarTasa(ss, hoy, p.origen, p.destino);
    const antes = buscarTasa(ss, hace30, p.origen, p.destino);
    if (!ahora) {
      avisos.push({
        nivel: 'Atención',
        titulo: 'Sin tasa ' + p.origen + '→' + p.destino,
        detalle: 'No hay tasa reciente. Todos los números de dinero de esa ' +
                 'tienda están sin convertir. Corre actualizarTasas().',
      });
      return;
    }
    if (antes) {
      const varPct = ((ahora.tasa - antes.tasa) / antes.tasa) * 100;
      if (Math.abs(varPct) >= umbral) {
        const sube = varPct > 0;
        avisos.push({
          nivel: Math.abs(varPct) >= umbral * 2 ? 'Crítica' : 'Atención',
          titulo: p.origen + '→' + p.destino + ' se movió ' + varPct.toFixed(1) + '% en 30 días',
          detalle: 'Pasó de ' + antes.tasa.toFixed(2) + ' a ' + ahora.tasa.toFixed(2) + '. ' +
            (sube
              ? 'Lo que factura esa tienda vale más en ' + destino + ' que hace un mes.'
              : 'Lo que factura esa tienda vale menos en ' + destino + '. ' +
                'Tu utilidad cae aunque la operación esté igual.'),
          accion: { texto: 'Ver efecto cambiario', destino: 'dinero' },
        });
      }
    }

    // Contra la tasa con la que se montó el negocio
    const ref = Number(parametro(ss, 'tasa_referencia_' + p.origen) || 0);
    if (ref > 0) {
      const varRef = ((ahora.tasa - ref) / ref) * 100;
      if (Math.abs(varRef) >= 10) {
        avisos.push({
          nivel: 'Informativa',
          titulo: p.origen + '→' + p.destino + ': ' + varRef.toFixed(0) +
                  '% contra tu tasa de referencia',
          detalle: 'Montaste el negocio con ' + ref.toFixed(0) + ' y hoy está en ' +
                   ahora.tasa.toFixed(0) + '. Si los precios no se movieron desde ' +
                   'entonces, el margen real cambió ' + varRef.toFixed(0) + '%.',
        });
      }
    }
  });

  Logger.log(avisos.length
    ? avisos.map(function (a) { return '[' + a.nivel + '] ' + a.titulo + '\n  ' + a.detalle; }).join('\n\n')
    : 'Sin movimientos de tasa relevantes.');
  return avisos;
}

/** Lee un parámetro. Busca primero por tienda, si no el global '*'. */
function parametro(ss, clave, tienda) {
  const sh = ss.getSheetByName('Parametros');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues().slice(1);
  let global = null;
  for (let i = 0; i < filas.length; i++) {
    if (norm(filas[i][1]) !== norm(clave)) continue;
    if (tienda && String(filas[i][0]).trim() === tienda) return filas[i][2];
    if (String(filas[i][0]).trim() === '*') global = filas[i][2];
  }
  return global;
}

// ─── EFECTO CAMBIARIO ────────────────────────────────────────

/**
 * Separa cuánto de tu variación es operación y cuánto es tipo de cambio.
 *
 * Es la pieza que evita el error caro: ver caer la utilidad en pesos y
 * salir a optimizar pauta, cuando lo que se movió fue el dólar.
 *
 * La descomposición es exacta:
 *   total      = R1·T1 − R0·T0
 *   operación  = (R1 − R0) · T0     ... cuánto habrías cambiado a tasa fija
 *   cambio     = R1 · (T1 − T0)     ... cuánto movió la tasa sola
 *   y operación + cambio = total, siempre.
 *
 * @param {string} cliente   nombre del cliente
 * @param {string} tienda    id de la tienda. Ej: 'ec'
 * @param {string} mesA      mes base 'AAAA-MM'
 * @param {string} mesB      mes a comparar 'AAAA-MM'
 */
function efectoCambiario(cliente, tienda, mesA, mesB) {
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const destino = monedaReporte(ss);
  const origen = monedaDeTienda(ss, tienda);
  if (!origen) throw new Error('No encuentro la tienda "' + tienda + '" en la hoja Tiendas.');

  if (origen === destino) {
    const msg = 'La tienda ' + tienda + ' opera en ' + origen +
                ', la misma moneda en que reportas. No hay efecto cambiario.';
    Logger.log(msg);
    return { aplica: false, mensaje: msg };
  }

  const R0 = ventasDelMes(ss, tienda, mesA);
  const R1 = ventasDelMes(ss, tienda, mesB);
  const T0 = tasaPromedioMes(ss, mesA, origen, destino);
  const T1 = tasaPromedioMes(ss, mesB, origen, destino);

  if (T0 === null || T1 === null) {
    const msg = 'Faltan tasas para ' + (T0 === null ? mesA : mesB) +
                '. Corre actualizarTasas() primero.';
    Logger.log(msg);
    return { aplica: false, mensaje: msg };
  }

  const enDestino0 = R0 * T0;
  const enDestino1 = R1 * T1;
  const total = enDestino1 - enDestino0;
  const porOperacion = (R1 - R0) * T0;
  const porCambio = R1 * (T1 - T0);

  const pct = function (x) {
    return enDestino0 ? (x / enDestino0 * 100).toFixed(1) + '%' : 'n/d';
  };

  const msg = [
    'EFECTO CAMBIARIO · tienda ' + tienda + ' · ' + mesA + ' vs ' + mesB,
    '',
    'Ventas en ' + origen + ':',
    '  ' + mesA + ': ' + fmtMoneda(R0, origen),
    '  ' + mesB + ': ' + fmtMoneda(R1, origen) +
      (R0 ? '   (' + ((R1 - R0) / R0 * 100).toFixed(1) + '%)' : ''),
    '',
    'Tasa ' + origen + '→' + destino + ':',
    '  ' + mesA + ': ' + T0.toFixed(2),
    '  ' + mesB + ': ' + T1.toFixed(2) +
      (T0 ? '   (' + ((T1 - T0) / T0 * 100).toFixed(1) + '%)' : ''),
    '',
    'En ' + destino + ':',
    '  ' + mesA + ': ' + fmtMoneda(enDestino0, destino),
    '  ' + mesB + ': ' + fmtMoneda(enDestino1, destino),
    '',
    '  variación total ....... ' + fmtMoneda(total, destino) + '   ' + pct(total),
    '  por OPERACIÓN ......... ' + fmtMoneda(porOperacion, destino) + '   ' + pct(porOperacion),
    '  por TIPO DE CAMBIO .... ' + fmtMoneda(porCambio, destino) + '   ' + pct(porCambio),
    '',
    lecturaEfecto(porOperacion, porCambio, destino),
  ].join('\n');

  Logger.log(msg);
  return {
    aplica: true, origen: origen, destino: destino,
    ventas: { a: R0, b: R1 }, tasas: { a: T0, b: T1 },
    total: total, porOperacion: porOperacion, porCambio: porCambio,
    mensaje: msg,
  };
}

/** Traduce los dos números a una frase que dice qué hacer. */
function lecturaEfecto(op, fx, destino) {
  const aOp = Math.abs(op), aFx = Math.abs(fx);
  if (aOp === 0 && aFx === 0) return 'Sin variación.';

  if (aFx > aOp * 2) {
    return op >= 0 && fx < 0
      ? '→ La operación MEJORÓ, pero la tasa se comió la mejora. No es un problema ' +
        'de ventas ni de pauta: es cambiario. Revisa precios, no campañas.'
      : '→ El movimiento es principalmente cambiario. La operación explica poco.';
  }
  if (aOp > aFx * 2) {
    return '→ El movimiento es de operación. La tasa casi no influyó: ' +
           'lo que cambió está en las ventas.';
  }
  return '→ Operación y tasa pesan parecido. Mira las dos antes de decidir.';
}

/** Moneda de una tienda, según la hoja Tiendas. */
function monedaDeTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cMon = enc.indexOf('moneda');
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][cId]).trim() === tienda) {
      return String(filas[i][cMon] || '').toUpperCase();
    }
  }
  return null;
}

/** Ventas entregadas de una tienda en un mes 'AAAA-MM'. */
function ventasDelMes(ss, tienda, mes) {
  const sh = ss.getSheetByName('Pedidos');
  if (!sh || sh.getLastRow() < 2) return 0;
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cF = enc.indexOf('fecha'), cT = enc.indexOf('tienda');
  const cV = enc.indexOf('valor'), cE = enc.indexOf('estado_canonico');

  let suma = 0;
  for (let i = 1; i < filas.length; i++) {
    const f = filas[i];
    if (String(f[cT]).trim() !== tienda) continue;
    if (String(aISO(f[cF], 'UTC')).slice(0, 7) !== mes) continue;
    // Solo lo entregado cuenta como venta: lo devuelto y lo cancelado no.
    if (cE !== -1 && norm(f[cE]) !== 'entregado') continue;
    const v = aNumero(f[cV]);
    if (v !== '') suma += v;
  }
  return suma;
}

/** Tasa promedio de un mes. Promedio y no cierre: las ventas se reparten. */
function tasaPromedioMes(ss, mes, origen, destino) {
  const sh = ss.getSheetByName('Tasas');
  if (!sh || sh.getLastRow() < 2) return null;
  const filas = sh.getDataRange().getValues().slice(1);
  let suma = 0, n = 0;
  filas.forEach(function (f) {
    const fecha = aISO(f[0], 'UTC');
    if (!fecha || fecha.slice(0, 7) !== mes) return;
    const o = String(f[1]).toUpperCase(), d = String(f[2]).toUpperCase();
    const t = aNumero(f[3]);
    if (t === '' || t <= 0) return;
    if (o === origen && d === destino) { suma += t; n++; }
    else if (o === destino && d === origen) { suma += 1 / t; n++; }
  });
  return n ? suma / n : null;
}
