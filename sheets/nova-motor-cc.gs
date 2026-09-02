/**
 * NOVA Motor — Centro de Comando v3
 * Plataformas: Dropi (EC + GT), Effi, Mastershop, Shopify, Meta, TikTok, IRIS
 *
 * Pegar en Extensions → Apps Script del sheet Centro de Comando.
 * Ejecutar instalarBotones() una vez para crear el menú y los triggers.
 *
 * ESTRUCTURA DE PESTAÑAS REQUERIDAS:
 *   PARAMETROS, MAPEO-PLATAFORMA, MAPEO-ESTATUS, MAPEO-CARTERA, MAPEO-PAUTA
 *   IMP-ORDENES, IMP-CARTERA, IMP-PAUTA, IMP-TARJETA
 *   MAESTRO-PEDIDOS, MAESTRO-CARTERA, MAESTRO-PAUTA
 *   GASTOS, DASHBOARD-DUENO, HISTORICO-CAMBIOS, RESUMEN, CIERRE-MES
 */

// ── Constantes globales ──────────────────────────────────────────────────────

var HOJAS = {
  par:    'PARAMETROS',
  mapPl:  'MAPEO-PLATAFORMA',
  mapEs:  'MAPEO-ESTATUS',
  mapCa:  'MAPEO-CARTERA',
  mapPa:  'MAPEO-PAUTA',
  impOr:  'IMP-ORDENES',
  impCa:  'IMP-CARTERA',
  impPa:  'IMP-PAUTA',
  impTa:  'IMP-TARJETA',
  mPed:   'MAESTRO-PEDIDOS',
  mCar:   'MAESTRO-CARTERA',
  mPau:   'MAESTRO-PAUTA',
  gas:    'GASTOS',
  dash:   'DASHBOARD-DUENO',
  hist:   'HISTORICO-CAMBIOS',
  res:    'RESUMEN',
  cie:    'CIERRE-MES',
};

var FILA_ENC  = 4;  // headers en fila 4
var FILA_DATOS = 5; // datos desde fila 5

// Headers canónicos MAESTRO-PEDIDOS (26 cols A–Z)
var ENC_PEDIDOS = [
  'ID_ORDEN','FECHA_ORDEN','MES','PLATAFORMA',
  'NOMBRE_CLIENTE','TELEFONO','CIUDAD','ZONA',
  'TRANSPORTADORA','GUIA','PRODUCTO','CANTIDAD',
  'VALOR_ORDEN','GANANCIA_NETA','COSTO_ENVIO',
  'ESTATUS_PLATAFORMA','ESTATUS_NORM','ESTADO_ENTREGA',
  'NOVEDAD_TIPO','NOVEDAD_DETALLE','SOLUCION',
  'GESTORA_ASIGNADA','INTENTOS_CONTACTO',
  'FECHA_ENTREGA','ULTIMA_ACTUALIZACION','NOTAS_INTERNAS',
];

// Headers canónicos MAESTRO-CARTERA (15 cols A–O)
var ENC_CARTERA = [
  'TELEFONO','NOMBRE_CLIENTE','CIUDAD','ZONA','ESTATUS_ACTUAL',
  'ORDENES_ACTIVAS','ORDENES_ENTREGADAS','ORDENES_DEVUELTAS',
  'TOTAL_FACTURADO','TOTAL_COBRADO','SALDO_PENDIENTE',
  'DIAS_SIN_CONTACTO','ULTIMO_CONTACTO','GESTORA_ASIGNADA','NOTAS',
];

// Headers canónicos MAESTRO-PAUTA (14 cols A–N)
var ENC_PAUTA = [
  'FECHA_INICIO','FECHA_FIN','MES','PLATAFORMA',
  'CAMPAÑA','CONJUNTO','PRESUPUESTO','GASTO',
  'IMPRESIONES','CLICS','RESULTADOS','CPR','ROAS','VALOR_CONVERSION',
];

// ── Punto de entrada principal ───────────────────────────────────────────────

function procesarTodo() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  log_(ss, 'Iniciando procesarTodo()');
  procesarPedidos_(ss);
  procesarCartera_(ss);
  procesarPautaYGastos_(ss);
  actualizarDashboard_(ss);
  log_(ss, 'procesarTodo() completado');
  SpreadsheetApp.getUi().alert('✅ Procesamiento completo');
}

function procesarPedidos()    { procesarPedidos_(SpreadsheetApp.getActiveSpreadsheet()); }
function procesarCartera()    { procesarCartera_(SpreadsheetApp.getActiveSpreadsheet()); }
function procesarPautaYGastos(){ procesarPautaYGastos_(SpreadsheetApp.getActiveSpreadsheet()); }

// ── PEDIDOS ──────────────────────────────────────────────────────────────────

function procesarPedidos_(ss) {
  var shImp  = ss.getSheetByName(HOJAS.impOr);
  var shMaes = ss.getSheetByName(HOJAS.mPed);
  if (!shImp || !shMaes) { log_(ss,'FALTA hoja IMP-ORDENES o MAESTRO-PEDIDOS'); return; }

  garantizarHeaders_(shMaes, ENC_PEDIDOS);

  var impData = shImp.getDataRange().getValues();
  if (impData.length < 2) { log_(ss,'IMP-ORDENES vacía'); return; }

  // Leer lo escrito por gestora en NOTAS_INTERNAS (col Z = índice 25) antes de reescribir
  var notasGuardadas = leerNotas_(shMaes);

  var encReales = impData[0].map(function(v){ return String(v).trim(); });
  var plataforma = detectarPlataforma_(ss, encReales);
  log_(ss, 'Plataforma detectada: ' + plataforma);

  var mapEstatus = leerMapeoEstatus_(ss);

  var filasMaes = [];
  for (var i = 1; i < impData.length; i++) {
    var row = impData[i];
    if (row.every(function(c){ return c === '' || c === null; })) continue;
    var fila = importarFilaPedido_(row, encReales, plataforma, mapEstatus);
    if (!fila) continue;
    // Restaurar nota interna si existe
    var id = String(fila[0]);
    if (notasGuardadas[id]) fila[25] = notasGuardadas[id];
    filasMaes.push(fila);
  }

  if (filasMaes.length === 0) { log_(ss,'Sin filas válidas en IMP-ORDENES'); return; }

  // Escribir desde FILA_DATOS, borrar lo anterior
  var lastRow = shMaes.getLastRow();
  if (lastRow >= FILA_DATOS) {
    shMaes.getRange(FILA_DATOS, 1, lastRow - FILA_DATOS + 1, ENC_PEDIDOS.length).clearContent();
  }
  shMaes.getRange(FILA_DATOS, 1, filasMaes.length, ENC_PEDIDOS.length).setValues(filasMaes);
  log_(ss, 'MAESTRO-PEDIDOS: ' + filasMaes.length + ' filas escritas (' + plataforma + ')');
}

function importarFilaPedido_(row, enc, plataforma, mapEstatus) {
  var g = function(nombre) {
    var i = enc.indexOf(nombre);
    return i >= 0 ? String(row[i] || '').trim() : '';
  };

  var id, fecha, nombre, tel, ciudad, zona, transp, guia, producto, cantidad,
      valor, ganancia, costoEnvio, estatusPlat, novedad, novDet, solucion, fechaEnt, ultAct;

  switch(plataforma) {
    case 'dropi_ec':
    case 'dropi_gt':
      id         = g('ID');
      fecha      = aFecha_(g('FECHA') || g('FECHA INGRESO'));
      nombre     = g('NOMBRE CLIENTE');
      tel        = tel10_(g('TELÉFONO') || g('TELEFONO'));
      ciudad     = g('CIUDAD DESTINO');
      zona       = '';
      transp     = g('TRANSPORTADORA');
      guia       = g('NÚMERO GUIA') || g('NUMERO GUIA');
      producto   = g('PRODUCTO');
      cantidad   = aNum_(g('CANTIDAD'));
      valor      = aNum_(g('TOTAL DE LA ORDEN'));
      ganancia   = aNum_(g('GANANCIA'));
      costoEnvio = 0;
      estatusPlat= g('ESTATUS');
      novedad    = g('NOVEDAD');
      novDet     = novedad;
      solucion   = g('SOLUCIÓN') || g('SOLUCION');
      fechaEnt   = '';
      ultAct     = g('ÚLTIMA ACTUALIZACIÓN') || g('ULTIMO MOVIMIENTO');
      break;

    case 'effi':
      id         = g('Reporte ID Pedido');
      fecha      = aFecha_(g('Fecha del Pedido'));
      nombre     = g('Cliente');
      tel        = tel10_(g('Teléfono'));
      ciudad     = g('Ciudad');
      zona       = g('Departamento');
      transp     = g('Transportadora');
      guia       = g('N° de Guía');
      producto   = g('Producto');
      cantidad   = aNum_(g('Cantidad'));
      valor      = aNum_(g('Valor del pedido'));
      ganancia   = 0;
      costoEnvio = aNum_(g('Valor del Envio'));
      estatusPlat= g('Estado del Pedido');
      novedad    = g('Presento Novedad') === 'SI' ? 'NOVEDAD' : '';
      novDet     = '';
      solucion   = g('Razón de cancelación');
      fechaEnt   = aFecha_(g('Fecha de Entrega'));
      ultAct     = g('Última Actualización');
      break;

    case 'mastershop':
      id         = g('Número de Orden') || g('ID');
      fecha      = aFecha_(g('Fecha de Orden') || g('Fecha'));
      nombre     = g('Nombre Completo') || g('Cliente');
      tel        = tel10_(g('Teléfono') || g('Celular'));
      ciudad     = g('Ciudad');
      zona       = g('Departamento');
      transp     = g('Transportadora');
      guia       = g('Guía') || g('Numero de Guia');
      producto   = g('Producto');
      cantidad   = aNum_(g('Cantidad'));
      valor      = aNum_(g('Valor') || g('Total'));
      ganancia   = 0;
      costoEnvio = aNum_(g('Costo Envío') || g('Flete'));
      estatusPlat= g('Estado') || g('Estatus');
      novedad    = g('Novedad') || '';
      novDet     = novedad;
      solucion   = g('Solución') || '';
      fechaEnt   = aFecha_(g('Fecha Entrega'));
      ultAct     = g('Última Actualización') || '';
      break;

    case 'shopify':
      id         = g('Name') || g('Order');
      fecha      = aFecha_(g('Created at'));
      nombre     = (g('Billing Name') || g('Shipping Name') || '').trim();
      tel        = tel10_(g('Billing Phone') || g('Shipping Phone'));
      ciudad     = g('Shipping City');
      zona       = g('Shipping Province');
      transp     = '';
      guia       = g('Tracking Number') || '';
      producto   = g('Lineitem name') || g('Name');
      cantidad   = aNum_(g('Lineitem quantity') || '1');
      valor      = aNum_(g('Total') || g('Subtotal'));
      ganancia   = 0;
      costoEnvio = aNum_(g('Shipping'));
      estatusPlat= g('Financial Status') || g('Fulfillment Status') || '';
      novedad    = '';
      novDet     = '';
      solucion   = g('Notes') || '';
      fechaEnt   = '';
      ultAct     = '';
      break;

    case 'iris':
      // IRIS: CRM de auditoría de llamadas
      id         = g('Llamada ID') || g('ID Llamada') || g('id');
      fecha      = aFecha_(g('Fecha Llamada') || g('Fecha') || g('fecha'));
      nombre     = g('Nombre Cliente') || g('Cliente') || g('nombre');
      tel        = tel10_(g('Teléfono') || g('Celular') || g('telefono'));
      ciudad     = g('Ciudad') || '';
      zona       = g('Departamento') || '';
      transp     = '';
      guia       = g('Guía') || g('Orden') || '';
      producto   = '';
      cantidad   = 0;
      valor      = 0;
      ganancia   = 0;
      costoEnvio = 0;
      estatusPlat= g('Resultado') || g('Estado') || '';
      novedad    = g('Novedad') || '';
      novDet     = g('Observación') || g('Comentario') || '';
      solucion   = g('Solución') || '';
      fechaEnt   = '';
      ultAct     = g('Fecha') || '';
      break;

    default:
      return null;
  }

  if (!id) return null;

  var estatusNorm   = normEstatus_(estatusPlat, mapEstatus);
  var estadoEntrega = resolverEstadoEntrega_(estatusNorm, novedad);
  var mes           = fecha ? aMes_(fecha) : '';

  return [
    id, fecha, mes, plataforma,
    norm_(nombre), tel, norm_(ciudad), norm_(zona),
    norm_(transp), guia, norm_(producto), cantidad,
    valor, ganancia, costoEnvio,
    estatusPlat, estatusNorm, estadoEntrega,
    novedad, novDet, solucion,
    '', 0,
    fechaEnt, ultAct, '',
  ];
}

// ── CARTERA ──────────────────────────────────────────────────────────────────

function procesarCartera_(ss) {
  var shMaes = ss.getSheetByName(HOJAS.mPed);
  var shCar  = ss.getSheetByName(HOJAS.mCar);
  if (!shMaes || !shCar) return;

  garantizarHeaders_(shCar, ENC_CARTERA);

  var maesData = shMaes.getDataRange().getValues();
  if (maesData.length < FILA_DATOS) return;

  var encPed = maesData[FILA_ENC - 1].map(function(v){ return String(v).trim(); });
  var idx = function(n){ return encPed.indexOf(n); };

  // Leer notas guardadas en cartera
  var notasCartera = leerNotasCartera_(shCar);

  // Agrupar por teléfono
  var mapa = {};
  for (var i = FILA_DATOS - 1; i < maesData.length; i++) {
    var r = maesData[i];
    var tel   = String(r[idx('TELEFONO')] || '').trim();
    if (!tel) continue;
    var nombre = String(r[idx('NOMBRE_CLIENTE')] || '');
    var ciudad = String(r[idx('CIUDAD')] || '');
    var zona   = String(r[idx('ZONA')] || '');
    var estado = String(r[idx('ESTADO_ENTREGA')] || '');
    var valor  = aNum_(r[idx('VALOR_ORDEN')]);
    var ultAct = r[idx('ULTIMA_ACTUALIZACION')];
    var gestora= String(r[idx('GESTORA_ASIGNADA')] || '');

    if (!mapa[tel]) {
      mapa[tel] = {
        tel:tel, nombre:nombre, ciudad:ciudad, zona:zona,
        activas:0, entregadas:0, devueltas:0,
        facturado:0, cobrado:0,
        ultContact: ultAct, gestora:gestora,
      };
    }
    var c = mapa[tel];
    c.facturado += valor;
    if (estado === 'entregado')  { c.entregadas++; c.cobrado += valor; }
    else if (estado === 'devuelto') c.devueltas++;
    else if (estado === 'en_camino' || estado === 'pendiente') c.activas++;
    if (gestora && !c.gestora) c.gestora = gestora;
  }

  var hoy = new Date();
  var filas = Object.values(mapa).map(function(c){
    var diasSin = c.ultContact
      ? Math.floor((hoy - new Date(c.ultContact)) / 86400000) : 999;
    var saldo  = c.facturado - c.cobrado;
    var estatus = resolverEstatusCartera_(c.entregadas, c.devueltas, c.activas, saldo);
    return [
      c.tel, c.nombre, c.ciudad, c.zona, estatus,
      c.activas, c.entregadas, c.devueltas,
      c.facturado, c.cobrado, saldo,
      diasSin, c.ultContact || '', c.gestora,
      notasCartera[c.tel] || '',
    ];
  });

  var lastRow = shCar.getLastRow();
  if (lastRow >= FILA_DATOS) {
    shCar.getRange(FILA_DATOS, 1, lastRow - FILA_DATOS + 1, ENC_CARTERA.length).clearContent();
  }
  if (filas.length > 0) {
    shCar.getRange(FILA_DATOS, 1, filas.length, ENC_CARTERA.length).setValues(filas);
  }
  log_(ss, 'MAESTRO-CARTERA: ' + filas.length + ' clientes');
}

function resolverEstatusCartera_(ent, dev, act, saldo) {
  if (dev > 2)         return 'bloqueado';
  if (saldo > 200000)  return 'moroso';
  if (dev > 0 && act > 0) return 'dudoso';
  if (ent > 0)         return 'ok';
  return 'activo';
}

// ── PAUTA ────────────────────────────────────────────────────────────────────

function procesarPautaYGastos_(ss) {
  var shImp  = ss.getSheetByName(HOJAS.impPa);
  var shMaes = ss.getSheetByName(HOJAS.mPau);
  if (!shImp || !shMaes) return;

  garantizarHeaders_(shMaes, ENC_PAUTA);

  var impData = shImp.getDataRange().getValues();
  if (impData.length < 2) return;

  var encReales = impData[0].map(function(v){ return String(v).trim(); });
  var plataforma = detectarPlataformaPauta_(encReales);
  log_(ss, 'Pauta plataforma: ' + plataforma);

  var filas = [];
  for (var i = 1; i < impData.length; i++) {
    var row = impData[i];
    if (row.every(function(c){ return c === '' || c === null; })) continue;
    var fila = importarFilaPauta_(row, encReales, plataforma);
    if (fila) filas.push(fila);
  }

  var lastRow = shMaes.getLastRow();
  if (lastRow >= FILA_DATOS) {
    shMaes.getRange(FILA_DATOS, 1, lastRow - FILA_DATOS + 1, ENC_PAUTA.length).clearContent();
  }
  if (filas.length > 0) {
    shMaes.getRange(FILA_DATOS, 1, filas.length, ENC_PAUTA.length).setValues(filas);
  }
  log_(ss, 'MAESTRO-PAUTA: ' + filas.length + ' filas (' + plataforma + ')');
}

function importarFilaPauta_(row, enc, plataforma) {
  var g = function(n) {
    var i = enc.indexOf(n);
    return i >= 0 ? String(row[i] || '').trim() : '';
  };

  var inicio, fin, campaña, conjunto, presup, gasto, imp, clics, res, cpr, roas, valConv;

  switch(plataforma) {
    case 'meta_camp':
      inicio  = aFecha_(g('Inicio del informe'));
      fin     = aFecha_(g('Fin del informe'));
      campaña = g('Nombre de la campaña');
      conjunto= '';
      presup  = aNum_(g('Presupuesto del conjunto de anuncios'));
      gasto   = aNum_(g('Importe gastado (COP)'));
      imp     = aNum_(g('Impresiones'));
      clics   = aNum_(g('Clics en el enlace'));
      res     = aNum_(g('Resultados') || g('Compras'));
      cpr     = aNum_(g('Coste por compra (COP)'));
      roas    = aNum_(g('ROAS'));
      valConv = aNum_(g('Valor de conversión'));
      break;

    case 'meta_adset':
      inicio  = aFecha_(g('Inicio del informe'));
      fin     = aFecha_(g('Fin del informe'));
      campaña = '';
      conjunto= g('Nombre del conjunto de anuncios');
      presup  = aNum_(g('Presupuesto del conjunto de anuncios'));
      gasto   = aNum_(g('Importe gastado (COP)'));
      imp     = aNum_(g('Impresiones'));
      clics   = aNum_(g('Clics en el enlace'));
      res     = aNum_(g('Resultados') || g('Compras'));
      cpr     = aNum_(g('Coste por compra (COP)'));
      roas    = aNum_(g('ROAS'));
      valConv = aNum_(g('Valor de conversión'));
      break;

    case 'tiktok':
      inicio  = aFecha_(g('Date') || g('Start Date') || g('Fecha'));
      fin     = inicio;
      campaña = g('Campaign Name') || g('Nombre de Campaña');
      conjunto= g('Ad Group Name') || g('Nombre del Grupo de Anuncios');
      presup  = aNum_(g('Budget') || g('Presupuesto'));
      gasto   = aNum_(g('Cost') || g('Gasto') || g('Spend'));
      imp     = aNum_(g('Impressions') || g('Impresiones'));
      clics   = aNum_(g('Clicks') || g('Clics'));
      res     = aNum_(g('Conversions') || g('Conversiones') || g('Results'));
      cpr     = aNum_(g('CPA') || g('Cost per Conversion'));
      roas    = aNum_(g('ROAS'));
      valConv = aNum_(g('Conversion Value') || g('Valor de Conversión'));
      break;

    default:
      return null;
  }

  if (!inicio) return null;
  return [
    inicio, fin, aMes_(inicio), plataforma,
    campaña, conjunto, presup, gasto,
    imp, clics, res, cpr, roas, valConv,
  ];
}

// ── DASHBOARD-DUENO ──────────────────────────────────────────────────────────

function actualizarDashboard_(ss) {
  var shDash  = ss.getSheetByName(HOJAS.dash);
  var shMaes  = ss.getSheetByName(HOJAS.mPed);
  var shPauta = ss.getSheetByName(HOJAS.mPau);
  if (!shDash || !shMaes) return;

  var hoy = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
  var mesAct = hoy.substring(0, 7);

  var maesData = shMaes.getDataRange().getValues();
  var encPed   = maesData[FILA_ENC - 1].map(function(v){ return String(v).trim(); });
  var idx = function(n){ return encPed.indexOf(n); };

  var pedHoy = 0, ventasHoy = 0, ganHoy = 0, novHoy = 0;
  var pedMes = 0, ventasMes = 0, ganMes = 0, devMes = 0, canMes = 0, entMes = 0;
  var activos = [];

  for (var i = FILA_DATOS - 1; i < maesData.length; i++) {
    var r = maesData[i];
    var fechaOrd = String(r[idx('FECHA_ORDEN')] || '');
    var mes      = String(r[idx('MES')] || '');
    var estado   = String(r[idx('ESTADO_ENTREGA')] || '');
    var valor    = aNum_(r[idx('VALOR_ORDEN')]);
    var gan      = aNum_(r[idx('GANANCIA_NETA')]);
    var nov      = String(r[idx('NOVEDAD_TIPO')] || '');

    if (fechaOrd.startsWith(hoy)) {
      pedHoy++;
      ventasHoy += valor;
      ganHoy    += gan;
      if (nov)   novHoy++;
    }
    if (mes === mesAct) {
      pedMes++;
      ventasMes += valor;
      ganMes    += gan;
      if (estado === 'devuelto')   devMes++;
      if (estado === 'cancelado')  canMes++;
      if (estado === 'entregado')  entMes++;
    }
    if (estado === 'en_camino' || estado === 'pendiente' || estado === 'novedad') {
      activos.push([
        r[idx('ID_ORDEN')], r[idx('NOMBRE_CLIENTE')], r[idx('TELEFONO')],
        r[idx('CIUDAD')], r[idx('GUIA')], r[idx('ESTATUS_NORM')],
        r[idx('NOVEDAD_TIPO')], r[idx('GESTORA_ASIGNADA')],
      ]);
    }
  }

  var tasaEnt = pedMes > 0 ? (entMes / pedMes) : 0;

  // Pauta del mes
  var gastoPauta = 0, sumaRoas = 0, contRoas = 0;
  if (shPauta) {
    var pauData = shPauta.getDataRange().getValues();
    var encPau  = pauData[FILA_ENC - 1].map(function(v){ return String(v).trim(); });
    var ip = function(n){ return encPau.indexOf(n); };
    for (var j = FILA_DATOS - 1; j < pauData.length; j++) {
      var pr = pauData[j];
      if (String(pr[ip('MES')] || '') === mesAct) {
        gastoPauta += aNum_(pr[ip('GASTO')]);
        var r = aNum_(pr[ip('ROAS')]);
        if (r > 0) { sumaRoas += r; contRoas++; }
      }
    }
  }
  var roasMes = contRoas > 0 ? (sumaRoas / contRoas) : 0;

  shDash.clearContents();

  var bloqueHoy = [
    ['CLAVE', 'VALOR'],
    ['fecha_hoy',       hoy],
    ['pedidos_hoy',     pedHoy],
    ['ventas_hoy',      ventasHoy],
    ['ganancia_hoy',    ganHoy],
    ['novedades_hoy',   novHoy],
    ['tasa_entrega_mes',tasaEnt],
  ];

  var bloqueMes = [
    ['CLAVE', 'VALOR'],
    ['mes_actual',      mesAct],
    ['pedidos_mes',     pedMes],
    ['ventas_mes',      ventasMes],
    ['ganancia_mes',    ganMes],
    ['devueltos_mes',   devMes],
    ['cancelados_mes',  canMes],
    ['gasto_pauta_mes', gastoPauta],
    ['roas_mes',        roasMes],
  ];

  shDash.getRange(1, 1, bloqueHoy.length, 2).setValues(bloqueHoy);
  shDash.getRange(12, 1, bloqueMes.length, 2).setValues(bloqueMes);

  if (activos.length > 0) {
    var encActivos = [['ID_ORDEN','NOMBRE_CLIENTE','TELEFONO','CIUDAD','GUIA',
                       'ESTATUS_NORM','NOVEDAD_TIPO','GESTORA_ASIGNADA']];
    shDash.getRange(23, 1, 1, 8).setValues(encActivos);
    shDash.getRange(24, 1, activos.length, 8).setValues(activos);
  }

  // Alarmas
  var alarmas = [];
  if (roasMes > 0 && roasMes < 1.5)
    alarmas.push(['CRITICA', 'ROAS bajo: ' + roasMes.toFixed(2)]);
  if (tasaEnt > 0 && tasaEnt < 0.6)
    alarmas.push(['CRITICA', 'Tasa de entrega baja: ' + (tasaEnt*100).toFixed(1) + '%']);
  if (pedHoy === 0)
    alarmas.push(['ALERTA', 'Sin pedidos hoy']);

  if (alarmas.length > 0) {
    var rowAl = 24 + activos.length + 2;
    shDash.getRange(rowAl, 1, 1, 2).setValues([['NIVEL','ALARMA']]);
    shDash.getRange(rowAl + 1, 1, alarmas.length, 2).setValues(alarmas);
  }

  log_(ss, 'DASHBOARD-DUENO actualizado');
}

// ── Detección de plataforma ──────────────────────────────────────────────────

function detectarPlataforma_(ss, encReales) {
  var shMap = ss.getSheetByName(HOJAS.mapPl);
  if (!shMap) return detectarPorDefecto_(encReales);

  var mapData = shMap.getDataRange().getValues();
  if (mapData.length < 2) return detectarPorDefecto_(encReales);

  var plataformas = mapData[0].map(function(v){ return String(v).trim(); });
  var mejorPl = '';
  var mejorScore = 3; // mínimo 4 matches

  for (var col = 0; col < plataformas.length; col++) {
    if (!plataformas[col]) continue;
    var headers = [];
    for (var row = 1; row < mapData.length; row++) {
      var h = String(mapData[row][col] || '').trim();
      if (h) headers.push(h);
    }
    var score = 0;
    headers.forEach(function(h){
      if (encReales.indexOf(h) >= 0) score++;
    });
    if (score > mejorScore) {
      mejorScore = score;
      mejorPl    = plataformas[col];
    }
  }
  return mejorPl || detectarPorDefecto_(encReales);
}

function detectarPorDefecto_(enc) {
  var tiene = function(n){ return enc.indexOf(n) >= 0; };
  if (tiene('NÚMERO GUIA') || tiene('NUMERO GUIA'))      return 'dropi_ec';
  if (tiene('Reporte ID Pedido'))                         return 'effi';
  if (tiene('Name') && tiene('Financial Status'))         return 'shopify';
  if (tiene('Número de Orden') || tiene('Numero de Orden')) return 'mastershop';
  if (tiene('Llamada ID') || tiene('ID Llamada'))         return 'iris';
  return 'dropi_ec';
}

function detectarPlataformaPauta_(enc) {
  var tiene = function(n){ return enc.indexOf(n) >= 0; };
  if (tiene('Nombre de la campaña'))           return 'meta_camp';
  if (tiene('Nombre del conjunto de anuncios')) return 'meta_adset';
  if (tiene('Campaign Name') || tiene('Ad Group Name')) return 'tiktok';
  return 'meta_camp';
}

// ── Mapeo de estatus ─────────────────────────────────────────────────────────

function leerMapeoEstatus_(ss) {
  var sh = ss.getSheetByName(HOJAS.mapEs);
  var mapa = {};
  if (!sh) return mapa;
  var data = sh.getDataRange().getValues();
  for (var i = 1; i < data.length; i++) {
    var raw  = norm_(String(data[i][0] || ''));
    var norm = String(data[i][1] || '').trim();
    if (raw && norm) mapa[raw] = norm;
  }
  return mapa;
}

function normEstatus_(estatusPlat, mapEstatus) {
  var k = norm_(estatusPlat);
  if (mapEstatus[k]) return mapEstatus[k];
  if (k.includes('entreg'))    return 'entregado';
  if (k.includes('devuel'))    return 'devuelto';
  if (k.includes('cancel'))    return 'cancelado';
  if (k.includes('camino') || k.includes('transito')) return 'en_camino';
  if (k.includes('novedad') || k.includes('excep'))   return 'novedad';
  if (k.includes('bodeg'))     return 'en_bodega';
  return 'pendiente';
}

function resolverEstadoEntrega_(estatusNorm, novedad) {
  if (estatusNorm === 'entregado') return 'entregado';
  if (estatusNorm === 'devuelto')  return 'devuelto';
  if (estatusNorm === 'cancelado') return 'cancelado';
  if (novedad && novedad.length > 0) return 'en_camino';
  if (estatusNorm === 'en_camino' || estatusNorm === 'en_bodega') return 'en_camino';
  return 'pendiente';
}

// ── Lectura de notas guardadas ────────────────────────────────────────────────

function leerNotas_(sh) {
  var mapa = {};
  var lastRow = sh.getLastRow();
  if (lastRow < FILA_DATOS) return mapa;
  var data = sh.getRange(FILA_DATOS, 1, lastRow - FILA_DATOS + 1, ENC_PEDIDOS.length).getValues();
  data.forEach(function(r){
    var id   = String(r[0] || '').trim();
    var nota = String(r[25] || '').trim();
    if (id && nota) mapa[id] = nota;
  });
  return mapa;
}

function leerNotasCartera_(sh) {
  var mapa = {};
  var lastRow = sh.getLastRow();
  if (lastRow < FILA_DATOS) return mapa;
  var data = sh.getRange(FILA_DATOS, 1, lastRow - FILA_DATOS + 1, ENC_CARTERA.length).getValues();
  data.forEach(function(r){
    var tel  = String(r[0] || '').trim();
    var nota = String(r[14] || '').trim();
    if (tel && nota) mapa[tel] = nota;
  });
  return mapa;
}

// ── Helpers de headers ───────────────────────────────────────────────────────

function garantizarHeaders_(sh, enc) {
  var encAct = sh.getRange(FILA_ENC, 1, 1, enc.length).getValues()[0];
  var iguales = enc.every(function(h, i){ return String(encAct[i] || '').trim() === h; });
  if (!iguales) {
    sh.getRange(FILA_ENC, 1, 1, enc.length).setValues([enc]);
    sh.getRange(FILA_ENC, 1, 1, enc.length)
      .setBackground('#1a1a2e')
      .setFontColor('#ffffff')
      .setFontWeight('bold');
  }
}

// ── Cierre de mes ─────────────────────────────────────────────────────────────

function guardarCierre() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var shRes = ss.getSheetByName(HOJAS.res);
  var shCie = ss.getSheetByName(HOJAS.cie);
  if (!shRes || !shCie) return;

  var mes = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM');
  var dashData = ss.getSheetByName(HOJAS.dash).getDataRange().getValues();

  var resumen = {};
  dashData.forEach(function(r){
    if (r[0] && r[1] !== '') resumen[String(r[0])] = r[1];
  });

  var fila = [
    mes,
    resumen['pedidos_mes']    || 0,
    resumen['ventas_mes']     || 0,
    resumen['ganancia_mes']   || 0,
    resumen['devueltos_mes']  || 0,
    resumen['cancelados_mes'] || 0,
    resumen['gasto_pauta_mes']|| 0,
    resumen['roas_mes']       || 0,
    new Date(),
  ];

  var lastRow = shCie.getLastRow();
  shCie.getRange(lastRow + 1, 1, 1, fila.length).setValues([fila]);
  log_(ss, 'Cierre guardado: ' + mes);
  SpreadsheetApp.getUi().alert('✅ Cierre ' + mes + ' guardado');
}

// ── Menú e instalación ───────────────────────────────────────────────────────

function instalarBotones() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('🏠 Nova CC')
    .addItem('▶ Procesar todo',         'procesarTodo')
    .addSeparator()
    .addItem('📦 Solo pedidos',          'procesarPedidos')
    .addItem('👥 Solo cartera',          'procesarCartera')
    .addItem('📣 Solo pauta',            'procesarPautaYGastos')
    .addSeparator()
    .addItem('📅 Guardar cierre de mes', 'guardarCierre')
    .addToUi();

  SpreadsheetApp.getActiveSpreadsheet().toast(
    'Motor Nova CC instalado. Usa el menú 🏠 Nova CC para procesar.', '✅', 5
  );
}

// ── Utilidades ───────────────────────────────────────────────────────────────

function norm_(v) {
  if (!v) return '';
  return String(v).trim()
    .normalize('NFD')
    .replace(/[̀-ͯ]/g, '')
    .toLowerCase();
}

function tel10_(v) {
  if (!v) return '';
  var t = String(v).replace(/\D/g, '');
  // Ecuador: prefijo 593 + 9 dígitos → quitar 593 → queda 9 dígitos, completar a 10 con 0
  if (t.length === 12 && t.startsWith('593')) { t = '0' + t.substring(3); }
  // Colombia: prefijo 57 + 10 dígitos
  if (t.length === 12 && t.startsWith('57'))  { t = t.substring(2); }
  // Guatemala: prefijo 502 + 8 dígitos
  if (t.length === 11 && t.startsWith('502')) { t = t.substring(3); }
  // Quitar + al inicio
  if (t.startsWith('+')) t = t.substring(1);
  // Dejar solo los últimos 10 dígitos
  if (t.length > 10) t = t.slice(-10);
  return t;
}

function aFecha_(v) {
  if (!v) return '';
  if (v instanceof Date) return Utilities.formatDate(v, 'America/Bogota', 'yyyy-MM-dd');
  var s = String(v).trim();
  // dd/mm/yyyy
  var m = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m) return m[3] + '-' + m[2].padStart(2,'0') + '-' + m[1].padStart(2,'0');
  // yyyy-mm-dd already
  if (/^\d{4}-\d{2}-\d{2}/.test(s)) return s.substring(0, 10);
  // mm/dd/yyyy
  var m2 = s.match(/^(\d{1,2})\/(\d{1,2})\/(\d{4})/);
  if (m2) return m2[3] + '-' + m2[1].padStart(2,'0') + '-' + m2[2].padStart(2,'0');
  return s.substring(0, 10);
}

function aMes_(v) {
  if (!v) return '';
  var s = String(v);
  if (s.length >= 7) return s.substring(0, 7);
  return '';
}

function aNum_(v) {
  if (v === null || v === undefined || v === '') return 0;
  if (typeof v === 'number') return v;
  var s = String(v).replace(/[^\d.,-]/g, '').replace(/\./g, '').replace(',', '.');
  var n = parseFloat(s);
  return isNaN(n) ? 0 : n;
}

function log_(ss, msg) {
  var sh = ss.getSheetByName(HOJAS.hist);
  if (!sh) return;
  var ts = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd HH:mm:ss');
  sh.appendRow([ts, msg]);
}
