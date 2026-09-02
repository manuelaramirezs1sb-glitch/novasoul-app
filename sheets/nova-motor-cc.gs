/**
 * NOVA Motor — Centro de Comando v3
 *
 * Rol de este script:
 *  1. instalarHojas()     — crea todas las tabs con headers correctos si no existen
 *  2. actualizarDashboard() — recalcula DASHBOARD-DUENO / ADMIN / GESTORA
 *     (se llama automáticamente via trigger cuando Netlify escribe datos)
 *  3. limpiarDatos()      — consolida y depura Pedidos/Novedades/Pauta
 *
 * Los IMPORTADORES viven en Netlify (/import). Este motor no necesita
 * pegar archivos; procesa lo que Netlify ya escribió en las hojas normalizadas.
 *
 * ESTRUCTURA DE PESTAÑAS (el script las crea todas):
 *   PARAMETROS · Pedidos · Novedades · Pauta · Inventario · Equipo · Cartera
 *   Movimientos · alertas_enviadas
 *   DASHBOARD-DUENO · DASHBOARD-ADMIN · DASHBOARD-GESTORA
 *   (Dias · Pendientes · Estudiantes · Progreso · Ejercicios — para otros productos)
 */

// ── Esquemas de headers ──────────────────────────────────────────────────────

var SCHEMAS = {
  // ── Nova Empresarial ────────────────────────────────────────────────────
  Pedidos: [
    'id','fecha','tienda','cliente','telefono','ciudad','direccion',
    'producto','cantidad','valor','costo_producto','costo_envio',
    'estado','transportadora','guia','intentos','gestora_asignada',
    'fecha_promesa','fecha_entrega','actualizado_en','actualizado_por'
  ],
  Novedades: [
    'id','pedido_id','fecha','tipo','motivo','estado',
    'gestora','nota','intentos','resuelta_en','actualizado_en'
  ],
  Pauta: [
    'fecha','tienda','plataforma','campaña','conjunto',
    'gasto','impresiones','clics','resultados','cpm','cpa','roas'
  ],
  Inventario: [
    'sku','producto','tienda','stock','costo_unitario','precio',
    'dias_cobertura','ultimo_conteo'
  ],
  Equipo: [
    'id','nombre','correo','rol','tienda','estado',
    'casos_asignados','casos_resueltos','nota_auditoria','ultima_conexion'
  ],
  Cartera: [
    'telefono','nombre','ciudad','tienda',
    'pedidos_activos','pedidos_entregados','pedidos_devueltos',
    'total_facturado','total_cobrado','saldo_pendiente',
    'dias_sin_contacto','ultimo_contacto','gestora_asignada','notas'
  ],
  Parametros: [
    'tienda','clave','valor','actualizado_en','actualizado_por'
  ],
  // ── Nova Central ────────────────────────────────────────────────────────
  Clientes: [
    'id','empresa','pais','plan','tarifa','costo','estado',
    'fecha_alta','fecha_corte','ultimo_pago','tickets_mes','usuarios','tiendas'
  ],
  Planes: [
    'id','nombre','modulos','limite_usuarios','limite_tiendas',
    'costo_calculado','precio_sugerido','tarifa_fijada'
  ],
  Solicitudes: [
    'id','cliente_id','nombre','correo','rol_pedido','estado',
    'fecha_solicitud','resuelta_en','resuelta_por'
  ],
  Candidatas: [
    'id','nombre','pais','experiencia','equipo_disponible',
    'estado','fecha_postulacion','nota'
  ],
  // ── NovaSoul (toda la data la genera la plataforma, no hay importación) ─
  Usuarios: [
    'id','nombre','correo','fecha_nacimiento','hora_nacimiento',
    'lugar_nacimiento','zona_horaria','acento','modo','idioma'
  ],
  Pendientes: [
    'id','usuario_id','texto','tipo','origen','fecha',
    'hecho','hecho_en','plataforma_id'
  ],
  Dias: [
    'usuario_id','fecha','comidas_marcadas','movimiento_hecho',
    'puntos','cerrado','cerrado_en','perdonado'
  ],
  Recompensas: [
    'id','usuario_id','nombre','costo_puntos','canjeada','canjeada_en'
  ],
  Transitos: [
    // precalculado desde datos natales con librería de efemérides, no importado
    'usuario_id','fecha','casa','tema','intensidad_pct',
    'texto_transito','por_que','como_trabajarlo','el_otro_lado'
  ],
  // ── novAcademy (todo lo genera la plataforma, no hay importación) ───────
  Estudiantes: [
    'id','nombre','correo','empresa','perfil','estado',
    'fecha_alta','alta_por','acceso_vence','primer_ingreso'
  ],
  Permisos: [
    'estudiante_id','ver_dinero','descargar_plantillas',
    'ver_avance_equipo','emitir_certificado','entrar_practica'
  ],
  Documentos: [
    'id','estudiante_id','nombre_archivo','tipo','peso','subido_en','subido_por'
  ],
  Progreso: [
    'estudiante_id','leccion_id','modulo','estado',
    'completada_en','nota_quiz','intentos','ultima_conexion'
  ],
  Ejercicios: [
    'id','estudiante_id','perfil','fecha','respuestas','aciertos','total',
    'nota','segundos','enviado_revision','comentario_profesor','comentado_en'
  ],
  Certificados: [
    'id','estudiante_id','ruta','emitido_en','vence','url_pdf'
  ],
  // ── Comunes ─────────────────────────────────────────────────────────────
  Movimientos: [
    'fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo'
  ],
  Alertas_enviadas: [
    'alarma_id','entidad_id','enviado_en','resuelta_en'
  ],
};

var ROW_DATA = 2; // datos desde fila 2 (fila 1 = headers)

// ── INSTALACIÓN ──────────────────────────────────────────────────────────────

function instalarHojas() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();

  Object.keys(SCHEMAS).forEach(function(nombre) {
    var sh = ss.getSheetByName(nombre);
    if (!sh) {
      sh = ss.insertSheet(nombre);
    }
    var enc = SCHEMAS[nombre];
    var actual = sh.getRange(1, 1, 1, enc.length).getValues()[0]
                    .map(function(v){ return String(v).trim(); });
    var igual  = enc.every(function(h, i){ return actual[i] === h; });
    if (!igual) {
      sh.getRange(1, 1, 1, enc.length).setValues([enc]);
      sh.getRange(1, 1, 1, enc.length)
        .setBackground('#1a1a2e').setFontColor('#ffffff').setFontWeight('bold');
    }
  });

  // Crear DASHBOARD tabs vacías si no existen (uno por producto/rol)
  [
    'DASHBOARD-DUENO','DASHBOARD-ADMIN','DASHBOARD-GESTORA',
    'DASHBOARD-CENTRAL',
    'DASHBOARD-SOUL',
    'DASHBOARD-ACADEMY-PROFESOR','DASHBOARD-ACADEMY-ESTUDIANTE',
  ].forEach(function(nm) {
    if (!ss.getSheetByName(nm)) ss.insertSheet(nm);
  });

  // Valores por defecto en Parametros (tienda=ec primero, gt segundo)
  var shPar = ss.getSheetByName('Parametros');
  var defaults = [
    // tienda, clave, valor, actualizado_en, actualizado_por
    ['ec','cpa_techo',              '38000',  '', 'instalacion'],
    ['ec','costos_fijos_mes',       '0',      '', 'instalacion'],
    ['ec','entrega_minima_pct',     '0.65',   '', 'instalacion'],
    ['ec','corte_despacho_hora',    '15:00',  '', 'instalacion'],
    ['ec','dias_reposicion_proveedor','7',    '', 'instalacion'],
    ['gt','cpa_techo',              '200',    '', 'instalacion'],
    ['gt','costos_fijos_mes',       '0',      '', 'instalacion'],
    ['gt','entrega_minima_pct',     '0.65',   '', 'instalacion'],
    ['gt','corte_despacho_hora',    '15:00',  '', 'instalacion'],
    ['gt','dias_reposicion_proveedor','7',    '', 'instalacion'],
  ];
  if (shPar.getLastRow() < 2) {
    shPar.getRange(2, 1, defaults.length, defaults[0].length).setValues(defaults);
  }

  instalarTriggers_();
  SpreadsheetApp.getActiveSpreadsheet()
    .toast('Nova CC instalado. Hojas y triggers listos.', '✅', 5);
}

function instalarTriggers_() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  // Borrar triggers previos para evitar duplicados
  ScriptApp.getProjectTriggers().forEach(function(t) {
    ScriptApp.deleteTrigger(t);
  });
  ScriptApp.newTrigger('onSheetChange')
    .forSpreadsheet(ss)
    .onChange()
    .create();
}

// Se llama automáticamente cuando Netlify escribe en la hoja
function onSheetChange(e) {
  var hojas = ['Pedidos','Novedades','Pauta','Inventario'];
  if (e && e.source) {
    var sheet = e.source.getActiveSheet();
    if (hojas.indexOf(sheet.getName()) >= 0) {
      actualizarDashboard();
    }
  }
}

// ── DASHBOARD ────────────────────────────────────────────────────────────────

function actualizarDashboard() {
  var ss = SpreadsheetApp.getActiveSpreadsheet();
  var par = leerParametros_(ss);

  var hoy     = Utilities.formatDate(new Date(), 'America/Bogota', 'yyyy-MM-dd');
  var mesAct  = hoy.substring(0, 7);

  var pedidos  = leerHoja_(ss, 'Pedidos',  SCHEMAS.Pedidos);
  var noveds   = leerHoja_(ss, 'Novedades',SCHEMAS.Novedades);
  var pauta    = leerHoja_(ss, 'Pauta',    SCHEMAS.Pauta);
  var equipo   = leerHoja_(ss, 'Equipo',   SCHEMAS.Equipo);

  // ── Métricas de pedidos ──────────────────────────────────────────────────
  var pedHoy=0, ventasHoy=0, pedMes=0, ventasMes=0;
  var devMes=0, canMes=0, entMes=0;
  var activos = [];

  pedidos.forEach(function(r) {
    var fecha  = String(r.fecha || '').substring(0, 10);
    var mes    = fecha.substring(0, 7);
    var estado = String(r.estado || '');
    var valor  = aNum_(r.valor);
    var esEnt  = estado === 'entregado';
    var esDev  = estado === 'devuelto';
    var esCan  = estado === 'cancelado';
    var esAct  = ['en_camino','pendiente','novedad','en_bodega'].indexOf(estado) >= 0;

    if (fecha === hoy)     { pedHoy++; ventasHoy += valor; }
    if (mes   === mesAct)  {
      pedMes++; ventasMes += valor;
      if (esEnt) entMes++;
      if (esDev) devMes++;
      if (esCan) canMes++;
    }
    if (esAct) {
      activos.push({
        id:r.id, cliente:r.cliente, telefono:r.telefono, ciudad:r.ciudad,
        guia:r.guia, estado:r.estado, gestora:r.gestora_asignada
      });
    }
  });

  var tasaEnt = pedMes > 0 ? entMes / pedMes : 0;

  // ── Novedades abiertas ───────────────────────────────────────────────────
  var novAbiertas = noveds.filter(function(r) {
    return String(r.estado || '') !== 'resuelta';
  });
  var novHoy = novAbiertas.filter(function(r) {
    return String(r.fecha || '').substring(0,10) === hoy;
  }).length;

  // ── Pauta del mes ────────────────────────────────────────────────────────
  var gastoPauta=0, sumaRoas=0, contRoas=0;
  pauta.forEach(function(r) {
    if (String(r.fecha || '').substring(0,7) === mesAct) {
      gastoPauta += aNum_(r.gasto);
      var rv = aNum_(r.roas);
      if (rv > 0) { sumaRoas += rv; contRoas++; }
    }
  });
  var roasMes = contRoas > 0 ? sumaRoas / contRoas : 0;

  // ── PARAMETROS ───────────────────────────────────────────────────────────
  var techoCP = aNum_(par['techo_cpa'] || '38000');

  // ── Alarmas ──────────────────────────────────────────────────────────────
  var alarmas = calcularAlarmas_(pedidos, noveds, pauta, equipo, par, hoy, mesAct,
    tasaEnt, roasMes, gastoPauta, techoCP);

  // ── Escribir DASHBOARD-DUENO ─────────────────────────────────────────────
  var shD = ss.getSheetByName('DASHBOARD-DUENO');
  shD.clearContents();

  var bloq = function(titulo, filas) { return [[titulo,'']].concat(filas); };

  var bloqHoy = bloq('HOY', [
    ['fecha_hoy',       hoy],
    ['pedidos_hoy',     pedHoy],
    ['ventas_hoy',      ventasHoy],
    ['novedades_hoy',   novHoy],
    ['tasa_entrega_mes',tasaEnt],
  ]);
  var bloqMes = bloq('MES', [
    ['mes_actual',      mesAct],
    ['pedidos_mes',     pedMes],
    ['ventas_mes',      ventasMes],
    ['devueltos_mes',   devMes],
    ['cancelados_mes',  canMes],
    ['gasto_pauta_mes', gastoPauta],
    ['roas_mes',        roasMes],
  ]);

  var pos = 1;
  shD.getRange(pos, 1, bloqHoy.length, 2).setValues(bloqHoy);
  pos += bloqHoy.length + 1;
  shD.getRange(pos, 1, bloqMes.length, 2).setValues(bloqMes);
  pos += bloqMes.length + 1;

  if (alarmas.length > 0) {
    shD.getRange(pos, 1, 1, 3).setValues([['NIVEL','ALARMA','ACCION']]);
    pos++;
    shD.getRange(pos, 1, alarmas.length, 3)
       .setValues(alarmas.map(function(a){ return [a.nivel, a.titulo, a.accion]; }));
    pos += alarmas.length + 1;
  }

  if (activos.length > 0) {
    shD.getRange(pos, 1, 1, 7)
       .setValues([['ACTIVOS — id','cliente','telefono','ciudad','guia','estado','gestora']]);
    pos++;
    shD.getRange(pos, 1, activos.length, 7)
       .setValues(activos.map(function(r){
         return [r.id, r.cliente, r.telefono, r.ciudad, r.guia, r.estado, r.gestora];
       }));
  }

  // ── DASHBOARD-ADMIN ──────────────────────────────────────────────────────
  var shA = ss.getSheetByName('DASHBOARD-ADMIN');
  shA.clearContents();

  var urgentes = pedidos.filter(function(r) {
    var f = String(r.fecha_promesa || '').substring(0,10);
    return f === hoy && String(r.estado||'') !== 'entregado';
  });
  var novSinRes = novAbiertas.length;

  var bloqAdmin = [
    ['HOY',''],
    ['urgentes_hoy',   urgentes.length],
    ['novedades_sin_resolver', novSinRes],
    ['tasa_entrega_mes', tasaEnt],
    ['',''],
    ['EQUIPO',''],
  ];
  equipo.forEach(function(g) {
    bloqAdmin.push([g.nombre + ' · ' + (g.rol || 'gestora'), g.casos_asignados + ' casos']);
  });
  shA.getRange(1, 1, bloqAdmin.length, 2).setValues(bloqAdmin);

  // ── DASHBOARD-GESTORA ────────────────────────────────────────────────────
  var shG = ss.getSheetByName('DASHBOARD-GESTORA');
  shG.clearContents();
  // Cada gestora leerá filtrando por su nombre desde el cliente.
  // El dashboard guarda todos los casos activos agrupados.
  var encCasos = [['gestora','id','cliente','telefono','ciudad','guia','estado','novedad']];
  var filasCasos = activos.map(function(r) {
    var nov = novAbiertas.find(function(n){ return n.pedido_id === r.id; });
    return [r.gestora, r.id, r.cliente, r.telefono, r.ciudad, r.guia, r.estado,
            nov ? nov.tipo : ''];
  });
  if (filasCasos.length > 0) {
    shG.getRange(1, 1, 1, 8).setValues(encCasos);
    shG.getRange(2, 1, filasCasos.length, 8).setValues(filasCasos);
  }
}

// ── Alarmas ──────────────────────────────────────────────────────────────────

function calcularAlarmas_(pedidos, noveds, pauta, equipo, par, hoy, mes,
    tasaEnt, roasMes, gastoPauta, techoCP) {
  var al = [];
  var add = function(nivel, titulo, accion) {
    al.push({ nivel:nivel, titulo:titulo, accion:accion });
  };

  // 1. Pedidos sin movimiento > 3 días
  var diasMaxMov = aNum_(par['dias_sin_mov_max'] || '3');
  pedidos.forEach(function(r) {
    if (['entregado','devuelto','cancelado'].indexOf(String(r.estado||'')) < 0) {
      var ultAct = r.actualizado_en ? new Date(r.actualizado_en) : new Date(r.fecha);
      var dias = (new Date(hoy) - ultAct) / 86400000;
      if (dias > diasMaxMov) {
        add('CRITICA', 'Pedido ' + r.id + ' sin movimiento ' + Math.floor(dias) + ' días',
            'Ver pedido ' + r.id);
      }
    }
  });

  // 2. Novedades sin resolver > 24 h
  var diasNov = aNum_(par['dias_novedad_max'] || '1');
  noveds.forEach(function(r) {
    if (String(r.estado||'') !== 'resuelta') {
      var fechaNov = r.fecha ? new Date(r.fecha) : new Date(hoy);
      var horas = (new Date(hoy + 'T23:59') - fechaNov) / 3600000;
      if (horas > diasNov * 24) {
        add('CRITICA', 'Novedad ' + r.id + ' sin resolver ' + Math.round(horas) + ' h',
            'Ver novedad ' + r.id);
      }
    }
  });

  // 5. Tasa de entrega < umbral
  var tasaMin = aNum_(par['tasa_entrega_min'] || '0.65');
  if (tasaEnt > 0 && tasaEnt < tasaMin) {
    add('CRITICA', 'Tasa de entrega ' + (tasaEnt*100).toFixed(1) + '% < ' + (tasaMin*100) + '%',
        'Revisar pedidos del mes');
  }

  // 6. CPA > techo
  var cpaMes = gastoPauta > 0 ? gastoPauta / Math.max(1, pedidos.filter(function(r){
    return r.fecha && r.fecha.substring(0,7) === mes && r.estado === 'entregado';
  }).length) : 0;
  if (cpaMes > techoCP) {
    add('CRITICA', 'CPA del mes $' + Math.round(cpaMes).toLocaleString('es') +
        ' > techo $' + techoCP.toLocaleString('es'), 'Ver Pauta');
  }

  if (tasaEnt > 0 && tasaEnt < 0.75) {
    add('ATENCION', 'Tasa de entrega por debajo del 75%', 'Revisar novedades');
  }

  if (roasMes > 0 && roasMes < 2.0) {
    add('ATENCION', 'ROAS del mes ' + roasMes.toFixed(2) + ' < 2.0×', 'Ver Pauta y ROAS');
  }

  return al;
}

// ── Lectura de hojas ──────────────────────────────────────────────────────────

function leerHoja_(ss, nombre, enc) {
  var sh = ss.getSheetByName(nombre);
  if (!sh || sh.getLastRow() < 2) return [];
  var data = sh.getRange(2, 1, sh.getLastRow() - 1, enc.length).getValues();
  return data
    .filter(function(r){ return r.some(function(c){ return c !== '' && c !== null; }); })
    .map(function(r) {
      var obj = {};
      enc.forEach(function(h, i){ obj[h] = r[i]; });
      return obj;
    });
}

function leerParametros_(ss) {
  var sh = ss.getSheetByName('PARAMETROS');
  var par = {};
  if (!sh || sh.getLastRow() < 2) return par;
  var data = sh.getRange(2, 1, sh.getLastRow()-1, 2).getValues();
  data.forEach(function(r){ if (r[0]) par[String(r[0]).trim()] = String(r[1]||'').trim(); });
  return par;
}

// ── Menú ─────────────────────────────────────────────────────────────────────

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('🏠 Nova CC')
    .addItem('Instalar hojas y triggers', 'instalarHojas')
    .addSeparator()
    .addItem('Recalcular dashboards', 'actualizarDashboard')
    .addToUi();
}

// ── Utilidades ────────────────────────────────────────────────────────────────

function aNum_(v) {
  if (!v && v !== 0) return 0;
  if (typeof v === 'number') return v;
  var n = parseFloat(String(v).replace(/[^\d.,-]/g,'').replace(/\./g,'').replace(',','.'));
  return isNaN(n) ? 0 : n;
}
