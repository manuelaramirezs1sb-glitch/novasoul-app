/**
 * Nova · Importadores
 * ─────────────────────────────────────────────────────────────
 * Convierte las pestañas _Import_* (crudas) en filas normalizadas.
 *
 * Regla del spec:
 *   "un cambio de formato rompe UNA función de importación, no seis pantallas"
 *
 * Por eso el mapeo es declarativo. Agregar una fuente nueva es agregar
 * un bloque a FUENTES — no escribir un importador nuevo.
 *
 * ⚠ ESTADO DE LOS MAPEOS
 *   Los alias de columna de abajo son la primera aproximación. Cada
 *   plataforma nombra distinto la misma cosa y cambia los nombres sin avisar.
 *   Marcados VERIFICADO los que ya se cotejaron contra un export real.
 *   Marcados POR VERIFICAR los que hay que confirmar con un CSV de muestra.
 */

// ─── CATÁLOGO DE FUENTES ─────────────────────────────────────
// `alias` mapea: columna_normalizada -> [posibles nombres en el export]
// El importador toma el primero que encuentre, sin distinguir may/min ni tildes.

const FUENTES = {

  // ── PEDIDOS / FULFILLMENT ──────────────────────────────────

  // VERIFICADO contra el Maestro de Nutrea EC (Dropi-Pedidos-NutreaShop.xlsx).
  // Ese archivo ya es la salida normalizada de tu motor anterior, así que
  // estos alias cubren tanto el export crudo de Dropi como tu Maestro.
  dropi: {
    tipo: 'pedidos',
    verificado: true,
    alias: {
      id_externo:       ['id', 'no. orden', 'numero orden', 'orden'],
      fecha:            ['fecha', 'fecha creacion', 'created at'],
      cliente:          ['nombre cliente', 'cliente', 'nombre'],
      telefono:         ['telefono', 'celular', 'phone'],
      guia:             ['numero guia', 'no guia', 'guia', 'tracking'],
      estado:           ['estatus', 'estado', 'status'],
      ciudad:           ['ciudad destino', 'ciudad', 'city'],
      transportadora:   ['transportadora', 'courier', 'operador'],
      valor:            ['total de la orden', 'valor', 'total', 'monto'],
      ganancia:         ['ganancia'],
      producto:         ['producto', 'productos', 'nombre producto'],
      cantidad:         ['cantidad', 'qty', 'unidades'],
      motivo_novedad:   ['novedad'],
      solucion:         ['solucion'],
      ultimo_movimiento:['ultimo movimiento'],
      fecha_ingreso:    ['fecha ingreso'],
      actualizado:      ['ultima actualizacion'],
      direccion:        ['direccion', 'address'],
      sku:              ['sku', 'codigo'],
      costo_envio:      ['flete', 'costo envio', 'envio'],
    },
  },

  // VERIFICADO contra reporte-historial-de-pedidos-*.xlsx
  mastershop: {
    tipo: 'pedidos',
    verificado: true,
    alias: {
      id_externo:       ['id pedido'],
      guia:             ['n° de guia', 'no de guia', 'n de guia'],
      fecha:            ['fecha del pedido'],
      fecha_entrega:    ['fecha de entrega'],
      cliente:          ['cliente'],
      cedula:           ['cedula'],
      correo:           ['correo electronico'],
      telefono:         ['telefono'],
      direccion:        ['direccion'],
      ciudad:           ['ciudad'],
      departamento:     ['departamento'],
      transportadora:   ['transportadora'],
      bodega:           ['nombre de la bodega'],
      costo_envio:      ['valor del envio'],
      tiene_novedad:    ['presento novedad'],
      estado:           ['estado del pedido'],
      estado_transportadora: ['estado en transportadora'],
      ultimo_movimiento:['ultima actualizacion'],
      metodo_pago:      ['metodo de pago'],
      producto:         ['producto'],
      valor:            ['valor del pedido'],
      cantidad:         ['cantidad'],
      costo_producto:   ['costo de productos'],
      nota:             ['notas'],
      razon_cancelacion:['razon de cancelacion'],
    },
  },

  // Effi usa el mismo formato de reporte que Mastershop.
  // Si un export real muestra diferencias, se separa en su propio bloque.
  effi: {
    tipo: 'pedidos',
    verificado: false, // POR VERIFICAR — se asume formato Mastershop
    hereda: 'mastershop',
  },

  // IRIS NO es una plataforma de pedidos: es la central telefónica.
  // VERIFICADO contra IRIS (1).csv — 2.782 llamadas salientes.
  // Alimenta la entidad Llamadas y cruza con Pedidos por telefono_norm.
  iris: {
    tipo: 'llamadas',
    verificado: true,
    alias: {
      id_externo:     ['id'],
      uid:            ['uniqueid'],
      fecha_hora:     ['fecha hora'],
      seg_conversado: ['tiempo conversado'],
      seg_espera:     ['tiempo espera'],
      seg_total:      ['tiempo total'],
      sentido:        ['sentido'],
      estado:         ['estado'],
      extension:      ['extension'],
      telefono:       ['numero'],
      agente:         ['usuario'],
      campana:        ['campana'],
      grabacion:      ['grabacion'],
      etiqueta:       ['etiqueta'],
      observacion:    ['observacion'],
    },
  },

  shopify: {
    tipo: 'pedidos',
    verificado: false, // POR VERIFICAR — Shopify sí está documentado,
                       // pero el export cambia según las columnas que marques
    alias: {
      id_externo:   ['name', 'id', 'order'],
      fecha:        ['created at', 'paid at', 'fecha'],
      cliente:      ['billing name', 'shipping name', 'customer name'],
      telefono:     ['phone', 'billing phone', 'shipping phone'],
      ciudad:       ['shipping city', 'billing city'],
      direccion:    ['shipping address1', 'billing address1'],
      producto:     ['lineitem name'],
      sku:          ['lineitem sku'],
      cantidad:     ['lineitem quantity'],
      valor:        ['total', 'subtotal'],
      costo_envio:  ['shipping'],
      estado:       ['fulfillment status', 'financial status'],
      guia:         ['tracking number'],
    },
  },

  // ── PAUTA ──────────────────────────────────────────────────

  meta: {
    tipo: 'pauta',
    verificado: false, // POR VERIFICAR
    // El spec avisa: el informe de Meta viene en COP. Se normaliza con
    // la tasa del DÍA DE LA TRANSACCIÓN, nunca la de hoy.
    moneda_default: 'COP',
    alias: {
      fecha:        ['dia', 'día', 'fecha', 'date', 'reporting starts'],
      campana:      ['nombre de la campaña', 'campaña', 'campaign name'],
      conjunto:     ['nombre del conjunto de anuncios', 'conjunto', 'ad set name'],
      gasto:        ['importe gastado', 'importe gastado (cop)', 'amount spent'],
      impresiones:  ['impresiones', 'impressions'],
      clics:        ['clics', 'clicks', 'clics en el enlace'],
      resultados:   ['resultados', 'results'],
      cpm:          ['cpm', 'cpm (costo por 1000 impresiones)'],
    },
  },

  tiktok: {
    tipo: 'pauta',
    verificado: false, // POR VERIFICAR
    moneda_default: 'USD',
    alias: {
      fecha:        ['date', 'fecha', 'dia'],
      campana:      ['campaign name', 'nombre de campaña', 'campaña'],
      conjunto:     ['ad group name', 'grupo de anuncios', 'conjunto'],
      gasto:        ['cost', 'spend', 'gasto', 'costo'],
      impresiones:  ['impressions', 'impresiones'],
      clics:        ['clicks', 'clics'],
      resultados:   ['conversions', 'results', 'conversiones'],
      cpm:          ['cpm'],
    },
  },
};

// ─── NORMALIZACIÓN DE TEXTO ──────────────────────────────────

/** Quita tildes, baja a minúscula, colapsa espacios. */
function norm(s) {
  return String(s == null ? '' : s)
    .normalize('NFD').replace(/[̀-ͯ]/g, '')
    .toLowerCase().trim().replace(/\s+/g, ' ');
}

/** Monto a número puro: sin $, sin puntos de miles, sin espacios. */
function aNumero(v) {
  if (typeof v === 'number') return v;
  if (v == null || v === '') return '';
  let s = String(v).replace(/[^\d,.\-]/g, '').trim();
  if (!s) return '';
  const ultimaComa = s.lastIndexOf(',');
  const ultimoPunto = s.lastIndexOf('.');
  // El separador decimal es el que aparece de último
  if (ultimaComa > ultimoPunto) {
    s = s.replace(/\./g, '').replace(',', '.');
  } else {
    s = s.replace(/,/g, '');
  }
  const n = parseFloat(s);
  return isNaN(n) ? '' : n;
}

/** Fecha a ISO AAAA-MM-DD. Causa número uno de errores al leer hojas. */
function aISO(v, zonaHoraria) {
  if (!v) return '';
  if (v instanceof Date) {
    return Utilities.formatDate(v, zonaHoraria || 'UTC', 'yyyy-MM-dd');
  }
  const s = String(v).trim();

  let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);           // ya ISO
  if (m) return m[1] + '-' + m[2] + '-' + m[3];

  m = s.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/); // DD/MM/AAAA
  if (m) {
    return m[3] + '-' + ('0' + m[2]).slice(-2) + '-' + ('0' + m[1]).slice(-2);
  }

  const d = new Date(s);
  if (!isNaN(d.getTime())) {
    return Utilities.formatDate(d, zonaHoraria || 'UTC', 'yyyy-MM-dd');
  }
  return '';
}

// ─── MOTOR DE IMPORTACIÓN ────────────────────────────────────

/**
 * Lee una pestaña _Import_* y devuelve filas normalizadas.
 * No escribe nada: separar la lectura de la escritura es lo que
 * permite probar el mapeo sin ensuciar las hojas de producción.
 */
function leerCrudo(ss, fuenteId, tienda) {
  let cfg = FUENTES[fuenteId];
  if (!cfg) throw new Error('Fuente desconocida: ' + fuenteId);
  // Fuentes que comparten formato (Effi usa el reporte de Mastershop)
  if (cfg.hereda) cfg = Object.assign({}, FUENTES[cfg.hereda], { tipo: cfg.tipo });

  const tab = '_Import_' + fuenteId.charAt(0).toUpperCase() + fuenteId.slice(1);
  const sh = ss.getSheetByName(tab);
  if (!sh) throw new Error('Falta la pestaña ' + tab);

  const datos = sh.getDataRange().getValues();
  if (datos.length < 2) return { filas: [], sinMapear: [], tab: tab };

  // La fila de encabezados es la primera que tenga 3+ celdas con texto
  let filaEnc = 0;
  for (let i = 0; i < Math.min(datos.length, 10); i++) {
    const llenas = datos[i].filter(function (c) { return String(c).trim() !== ''; });
    if (llenas.length >= 3) { filaEnc = i; break; }
  }

  const enc = datos[filaEnc].map(norm);
  const cuerpo = datos.slice(filaEnc + 1);

  // columna normalizada -> índice en el export
  const idx = {};
  Object.keys(cfg.alias).forEach(function (campo) {
    const opciones = cfg.alias[campo].map(norm);
    for (let i = 0; i < enc.length; i++) {
      if (opciones.indexOf(enc[i]) !== -1) { idx[campo] = i; return; }
    }
  });

  const sinMapear = Object.keys(cfg.alias).filter(function (c) {
    return idx[c] === undefined;
  });

  const tz = zonaHorariaDe(ss, tienda);
  const filas = cuerpo
    .filter(function (r) { return r.some(function (c) { return String(c).trim() !== ''; }); })
    .map(function (r) {
      const o = { fuente: fuenteId, tienda: tienda };
      Object.keys(idx).forEach(function (campo) {
        let v = r[idx[campo]];
        if (campo === 'fecha') v = aISO(v, tz);
        else if (['valor','costo_envio','costo_producto','cantidad','gasto',
                  'impresiones','clics','resultados','cpm'].indexOf(campo) !== -1) {
          v = aNumero(v);
        } else v = String(v == null ? '' : v).trim();
        o[campo] = v;
      });
      return o;
    });

  return { filas: filas, sinMapear: sinMapear, tab: tab, tipo: cfg.tipo };
}

/** Zona horaria de la tienda, para que las alarmas 1/2/4 evalúen en hora local. */
function zonaHorariaDe(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return 'UTC';
  const filas = sh.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cTz = enc.indexOf('zona_horaria');
  for (let i = 1; i < filas.length; i++) {
    if (String(filas[i][cId]).trim() === tienda) return filas[i][cTz] || 'UTC';
  }
  return 'UTC';
}

/**
 * Diagnóstico. Córrelo con un export pegado en la pestaña y te dice
 * exactamente qué columnas no encontró, para completar el mapeo.
 */
function diagnosticar(fuenteId, tienda) {
  const ss = SpreadsheetApp.openById(IDS.empresarial);
  const r = leerCrudo(ss, fuenteId, tienda || 'gt');
  const msg = [
    'Fuente: ' + fuenteId + '  (' + r.tipo + ')',
    'Pestaña: ' + r.tab,
    'Filas leídas: ' + r.filas.length,
    'Columnas SIN mapear: ' + (r.sinMapear.length ? r.sinMapear.join(', ') : 'ninguna ✓'),
    '',
    'Primera fila normalizada:',
    JSON.stringify(r.filas[0] || {}, null, 2),
  ].join('\n');
  Logger.log(msg);
  return msg;
}
