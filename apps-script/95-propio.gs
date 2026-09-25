/**
 * ═══════════════════════════════════════════════════════════════
 *  EL CONTROL DIARIO DE UNA TIENDA · novedades y CAS
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ DE DÓNDE SALEN ESTOS DICCIONARIOS ────────────────────────┐
 * │                                                            │
 * │ No de imaginarse cómo escribiría alguien una columna: de    │
 * │ un archivo REAL de gestión logística de una tienda —        │
 * │ diecisiete hojas, dos mil setecientas filas— que se leyó    │
 * │ entero antes de escribir una línea de esto.                 │
 * │                                                            │
 * │ Sus encabezados, tal cual:                                 │
 * │                                                            │
 * │   NOVEDADES  FECHA DE GESTION · ID · FECHA NOVEDAD ·       │
 * │              CLIENTE · SOLUCION · SOLUCIONADA ·             │
 * │              GESTIONA · NOTAS                              │
 * │                                                            │
 * │   CAS        FECHA RADICACION · ID · FECHA DE ENVIO ORDEN ·│
 * │              CLIENTE · TELEFONO · NUMERO DE GUIA ·         │
 * │              ESTATUS · TRANSPORTADORA ·                    │
 * │              FECHA DE ULTIMO MOVIMIENTO · GESTION ·        │
 * │              RESPUESTA · NOTAS                             │
 * │                                                            │
 * │ Los sinónimos cubren esos y sus variantes. Pero, igual que │
 * │ con los pedidos, NADA se importa a ciegas: Nova propone lo │
 * │ que entendió, la persona corrige, y solo entonces se       │
 * │ escribe. Un histórico mal leído es peor que no cargarlo.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ «GESTIONA» NO ES QUIEN LO VE ─────────────────────┐
 * │                                                            │
 * │ En ese archivo la columna GESTIONA dice JAIME (399 filas), │
 * │ APOYO (68), ZULAY (48), GERALD (15), RUBEN (3). Gente sin  │
 * │ cuenta en Nova, y «APOYO», que ni siquiera es una persona. │
 * │                                                            │
 * │ Por eso cae en `gestionado_por` —una etiqueta— y nunca en  │
 * │ `gestora_asignada`, que es control de acceso. Así la       │
 * │ dueña importa su historia entera y tiene sus números por   │
 * │ agente desde el primer día, sin crear una sola cuenta.     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

const SINONIMOS_NOVEDADES = {
  id_externo:   ['id','id pedido','no orden','numero orden','orden','pedido',
                 'id de la orden','no. orden','referencia'],
  fecha:        ['fecha novedad','fecha de la novedad','fecha','dia','fecha reporte'],
  fecha_solucion:['fecha de gestion','fecha gestion','fecha de solucion','fecha solucion',
                 'fecha de cierre','gestionada el'],
  cliente:      ['cliente','clienta','destinatario','nombre cliente','nombre'],
  motivo:       ['novedad','motivo','motivo novedad','tipo de novedad','causa',
                 'razon','descripcion novedad'],
  solucion:     ['solucion','soluciones','historico de soluciones','que se hizo',
                 'gestion realizada','accion'],
  solucionada:  ['solucionada','fue solucionada','fue solucionada la novedad',
                 'resuelta','cerrada','se soluciono'],
  gestionado_por:['gestiona','gestor','gestora','agente','asesor','asesora',
                 'responsable','atendido por','quien gestiona','encargada'],
  nota:         ['notas','nota','observacion','observaciones','comentario',
                 'comentarios','detalle'],
  guia:         ['guia','numero guia','no guia','tracking','guia novedad'],
  estado:       ['estado','estatus','status','situacion'],
};

const FORMAS_NOVEDADES = {
  id_externo:    'texto',
  fecha:         'fecha',
  fecha_solucion:'fecha',
  cliente:       'texto',
  motivo:        'texto_repetido',
  solucion:      'texto',
  solucionada:   'texto_constante',
  gestionado_por:'texto_repetido',
  nota:          'texto',
  guia:          'texto',
  estado:        'texto_repetido',
};

/**
 * El CAS: el reclamo que se le radica a la transportadora.
 *
 * Es la parte del trabajo que más se pierde cuando alguien cambia de
 * herramienta, porque vive en un Excel aparte y nadie la migra. Son
 * ciento sesenta reclamos con su respuesta — meses de gestión.
 */
const SINONIMOS_CAS = {
  id_externo:   ['id','id pedido','orden','no orden','numero orden','pedido'],
  abierto_en:   ['fecha radicacion','fecha de radicacion','radicado el','fecha',
                 'fecha caso','fecha del caso'],
  fecha_envio:  ['fecha de envio orden','fecha envio','fecha de envio','despachado'],
  cliente:      ['cliente','clienta','destinatario','nombre cliente'],
  telefono:     ['telefono','celular','tel','contacto','numero'],
  guia:         ['numero de guia','numero guia','no guia','guia','tracking'],
  estado:       ['estatus','estado','status','situacion','estado del caso'],
  transportadora:['transportadora','courier','operador','empresa envio','mensajeria'],
  ultima_gestion:['fecha de ultimo movimiento','ultimo movimiento','ultima gestion',
                 'ultima actualizacion','fecha ultimo movimiento'],
  gestion:      ['gestion','que se hizo','accion','gestion realizada','seguimiento'],
  respuesta:    ['respuesta','respuesta transportadora','resultado','desenlace'],
  ticket:       ['ticket','radicado','numero de caso','caso','pqr','numero pqr'],
  nota:         ['notas','nota','observacion','observaciones','comentario','detalle'],
  gestionado_por:['gestiona','gestor','gestora','agente','asesor','responsable',
                 'atendido por','radicado por'],
};

const FORMAS_CAS = {
  id_externo:    'texto',
  abierto_en:    'fecha',
  fecha_envio:   'fecha',
  cliente:       'texto',
  telefono:      'texto',
  guia:          'texto',
  estado:        'texto_repetido',
  transportadora:'texto_repetido',
  ultima_gestion:'fecha',
  gestion:       'texto',
  respuesta:     'texto',
  ticket:        'texto',
  nota:          'texto',
  gestionado_por:'texto_repetido',
};

/**
 * Qué diccionario usar según lo que se esté subiendo.
 *
 * Existe para que `apiSubirArchivo` deje de tener «pedidos» escrito a
 * mano. Cuando se agregó la tercera cosa importable, esa constante
 * enterrada en una llamada era justo lo que había que encontrar y
 * nadie sabía que estaba ahí.
 */
function diccionarioDe_(tipo) {
  if (tipo === 'novedades') {
    return { dicc: SINONIMOS_NOVEDADES, formas: FORMAS_NOVEDADES };
  }
  if (tipo === 'cas') {
    return { dicc: SINONIMOS_CAS, formas: FORMAS_CAS };
  }
  return { dicc: SINONIMOS_PEDIDOS, formas: FORMAS_PEDIDOS };
}

/**
 * Normaliza una fila de CAS recién leída del archivo.
 *
 * Lo único que no puede salir del archivo es `pedido_id`: el CAS habla
 * del mismo pedido que ya está en Nova, y la única forma de amarrarlos
 * es por el id externo. Si no calza, el caso entra igual y se dice —
 * un reclamo sin pedido sigue siendo un reclamo, y esconderlo sería
 * perder la gestión que costó hacerlo.
 */
function normalizarCas_(o, fuenteId, tienda, porExterno) {
  const ext = String(o.id_externo || '').trim();
  o.id = fuenteId + '-cas-' + (ext || Utilities.getUuid().slice(0, 8));
  o.tienda = tienda;
  o.pedido_id = (ext && porExterno && porExterno[ext.toLowerCase()]) || '';

  /**
   * «Días quieto» no se importa: se calcula. Un número copiado de un
   * archivo de hace dos meses dice «3 días» para siempre, y eso es
   * peor que no decir nada — parece fresco y no lo está.
   */
  const ult = String(o.ultima_gestion || o.abierto_en || '').slice(0, 10);
  if (/^\d{4}-\d{2}-\d{2}$/.test(ult)) {
    o.dias_quieto = Math.max(0, Math.round(
      (new Date(ahoraISO().slice(0, 10) + 'T00:00:00Z') -
       new Date(ult + 'T00:00:00Z')) / 86400000));
  }

  /**
   * Si trae respuesta, el caso ya se cerró. Es lo que permite que los
   * ciento sesenta reclamos viejos no aparezcan todos como abiertos el
   * primer día — que es exactamente lo que haría que nadie vuelva a
   * mirar esta pantalla.
   */
  if (String(o.respuesta || '').trim() && !o.cerrado_en) {
    o.cerrado_en = String(o.ultima_gestion || o.abierto_en || '').slice(0, 10);
  }
  return o;
}

/**
 * Los pedidos ya importados, indexados por su id externo.
 *
 * Sirve para amarrar novedades y CAS al pedido del que hablan sin
 * recorrer la hoja de Pedidos una vez por fila.
 */
function pedidosPorExterno_(ss) {
  const sh = ss.getSheetByName('Pedidos');
  if (!sh || sh.getLastRow() < 2) return {};
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cExt = e.indexOf('id_externo');
  if (cId === -1) return {};
  const out = {};
  for (let i = 1; i < d.length; i++) {
    const ext = cExt === -1 ? '' : String(d[i][cExt] || '').trim();
    if (ext) out[ext.toLowerCase()] = String(d[i][cId] || '');
  }
  return out;
}
