/**
 * ═══════════════════════════════════════════════════════════
 *  META · LA LECTURA DIARIA
 * ═══════════════════════════════════════════════════════════
 *
 * Aquí es donde empieza a pasar el agua. 75-meta.gs deja la llave puesta
 * y comprobada; este archivo la usa para pedirle a Meta el gasto de cada
 * día y escribirlo en Pauta, sin que nadie exporte un Excel.
 *
 * TRES COSAS QUE SE ARREGLAN SOLAS AL PEDIRLO ASÍ
 *
 * 1. `time_increment=1` obliga a Meta a devolver UNA FILA POR DÍA. El
 *    export manual trae una sola fila por todo el periodo, y repartirla
 *    entre los días dibuja una curva que nunca existió.
 *
 * 2. La moneda viene declarada por Meta (`account_currency`), no
 *    adivinada por el nombre de una columna. El export dice "Importe
 *    gastado (COP)" solo a veces, y cuando no lo dice, el importador
 *    tiene que suponer.
 *
 * 3. Las fechas vienen en ISO. El Excel las trae en el formato de quien
 *    lo exportó, y 03/04 es marzo o abril según el país.
 *
 * LO QUE NO SE INVENTA
 *
 * Meta no entrega presupuesto ni estado de entrega en este informe. Esas
 * columnas quedan VACÍAS, no en cero. Un cero dice "no había
 * presupuesto"; vacío dice "no lo sé", que es la verdad.
 */

/**
 * Qué se le pide a Meta.
 *
 * A nivel de conjunto (`adset`) y no de campaña porque el presupuesto se
 * decide por conjunto: es la unidad que alguien puede subir, bajar o
 * apagar mañana. Un consejo sobre una campaña que adentro tiene tres
 * conjuntos con rendimientos distintos no se puede ejecutar.
 */
const META_NIVEL = 'adset';

/**
 * Y los anuncios, uno por uno, en su propia hoja.
 *
 * Es otra pregunta: el conjunto dice dónde va el presupuesto, el anuncio
 * dice cuál creativo tira del carro. Se piden aparte y se guardan aparte
 * porque el gasto está en los dos —el conjunto de 100 son los mismos 100
 * repartidos entre sus anuncios— y sumarlos juntos contaría todo dos
 * veces.
 *
 * Cuesta llamadas: un conjunto con cuatro creativos son cuatro filas por
 * día en vez de una. Por eso se puede apagar por tienda, y por eso la
 * lectura diaria trae menos días de anuncios que de conjuntos.
 */
const META_NIVEL_ANUNCIO = 'ad';

const META_CAMPOS_ANUNCIO = [
  'date_start', 'account_currency', 'campaign_name', 'adset_name',
  'ad_name', 'ad_id', 'spend', 'impressions', 'reach', 'frequency',
  'clicks', 'ctr', 'cpc', 'cpm', 'actions', 'action_values',
];

/** Cuántos días de anuncios se repiden a diario. Menos que los conjuntos. */
const META_DIAS_ANUNCIOS = 3;

const META_CAMPOS = [
  'date_start', 'date_stop', 'account_currency',
  'campaign_name', 'adset_name', 'adset_id',
  'spend', 'impressions', 'reach', 'frequency',
  'clicks', 'ctr', 'cpc', 'cpm',
  'actions', 'action_values',
];

/**
 * Cuántos días se vuelven a pedir en cada lectura diaria.
 *
 * No basta con pedir "ayer". Meta sigue atribuyendo conversiones hasta
 * días después de que ocurrió el clic, así que las cifras de ayer
 * cambian pasado mañana. Se vuelven a pedir siete días y se reescriben
 * encima: el número siempre es el último que Meta conoce, no el primero
 * que dijo.
 */
const META_DIAS_DIARIO = 7;

/**
 * Cómo se reconoce una compra, y en qué orden.
 *
 * Meta devuelve una lista de "acciones" con nombres técnicos y varias
 * pueden ser la misma venta contada de dos formas. Se toma la primera de
 * esta lista que exista, y se DICE cuál se tomó, porque de esto depende
 * el CPA entero y quien lo mire tiene derecho a saber de dónde salió.
 */
const META_ACCIONES_COMPRA = [
  'offsite_conversion.fb_pixel_purchase',
  'onsite_web_purchase',
  'omni_purchase',
  'purchase',
];

/**
 * Y cómo se reconoce un registro.
 *
 * En contraentrega mucha gente no optimiza a compra sino a formulario:
 * el pedido se confirma después, por teléfono. Si no hay compras, el
 * resultado del conjunto son esos registros — pero nunca se mezclan los
 * dos en la misma columna.
 */
const META_ACCIONES_LEAD = [
  'offsite_conversion.fb_pixel_lead',
  'onsite_conversion.lead_grouped',
  'lead',
];

const META_ACCION_LP = ['landing_page_view'];

/** El primer tipo de la lista que Meta reportó, con su valor. */
function metaAccion_(acciones, tipos) {
  if (!acciones || !acciones.length) return null;
  const porTipo = {};
  acciones.forEach(function (a) { porTipo[a.action_type] = a.value; });
  for (let i = 0; i < tipos.length; i++) {
    if (porTipo[tipos[i]] !== undefined) {
      return { tipo: tipos[i], valor: Number(porTipo[tipos[i]]) || 0 };
    }
  }
  return null;
}

/**
 * Una llamada a Meta, con sus páginas.
 *
 * Meta parte las respuestas largas y deja un enlace al resto. Sin seguir
 * ese enlace, una cuenta con muchos conjuntos devolvería los primeros y
 * el resto simplemente no existiría — un gasto que falta sin que nada
 * avise es peor que un error.
 */
function metaPedir_(url, token) {
  const filas = [];
  let siguiente = url;
  let vueltas = 0;

  while (siguiente && vueltas < 40) {   // tope: una cuenta rota no cuelga el script
    vueltas++;
    let r;
    try {
      r = UrlFetchApp.fetch(siguiente, { muteHttpExceptions: true });
    } catch (err) {
      return { ok: false, error: 'No se pudo hablar con Meta: ' + err.message };
    }

    let j;
    try { j = JSON.parse(r.getContentText()); }
    catch (err) { return { ok: false, error: 'Meta contestó algo que no entiendo.' }; }

    if (j.error) return Object.assign({ ok: false }, explicarErrorMeta(j.error));

    (j.data || []).forEach(function (f) { filas.push(f); });

    // El enlace siguiente ya trae la llave adentro; el primero no.
    siguiente = (j.paging && j.paging.next) ? j.paging.next : '';
    if (siguiente && siguiente.indexOf('access_token') === -1) {
      siguiente += '&access_token=' + encodeURIComponent(token);
    }
  }

  return { ok: true, filas: filas, paginas: vueltas };
}

/** yyyy-MM-dd de hace N días, en la zona del script. */
function metaFecha_(diasAtras) {
  const d = new Date(Date.now() - diasAtras * 86400000);
  return Utilities.formatDate(d, Session.getScriptTimeZone(), 'yyyy-MM-dd');
}

/**
 * Trae el gasto de una tienda y lo escribe en Pauta.
 *
 * Devuelve un informe de lo que hizo, no un "listo". Quien paga por esto
 * necesita poder auditarlo: cuántas filas entraron, cuántas cambiaron,
 * qué acción se contó como compra y si hay filas viejas que se solapan.
 *
 * @param {string} sheetId  la hoja del cliente
 * @param {string} tienda   el id de la tienda
 * @param {string} token    la llave de Meta de ESE cliente
 * @param {number} dias     cuántos días hacia atrás pedir
 */
function metaLeerTienda_(sheetId, tienda, token, dias) {
  const ss = SpreadsheetApp.openById(sheetId);
  const cuenta = metaCuenta(ss, tienda);
  const informe = {
    tienda: tienda, cuenta: cuenta, ok: false, error: '',
    filas: 0, nuevas: 0, actualizadas: 0, iguales: 0,
    desde: '', hasta: '', moneda: '', monedaTienda: monedaDeTienda(ss, tienda) || '',
    gasto: 0, accionCompra: '', accionResultado: '', solapadas: 0, avisos: [],
  };

  if (!cuenta) {
    informe.error = 'Esta tienda no tiene número de cuenta publicitaria. ' +
                    'Se pone en Configuración → Conexión con Meta.';
    return informe;
  }

  const hasta = metaFecha_(0);
  const desde = metaFecha_(Math.max(1, dias || META_DIAS_DIARIO));
  informe.desde = desde; informe.hasta = hasta;

  const url = META_API + 'act_' + cuenta + '/insights' +
    '?level=' + META_NIVEL +
    '&time_increment=1' +
    '&time_range=' + encodeURIComponent(JSON.stringify({ since: desde, until: hasta })) +
    '&fields=' + META_CAMPOS.join(',') +
    '&limit=200' +
    '&access_token=' + encodeURIComponent(token);

  const r = metaPedir_(url, token);
  if (!r.ok) { informe.error = r.error; return informe; }

  informe.filas = r.filas.length;
  if (!r.filas.length) {
    informe.ok = true;
    informe.avisos.push('Meta no reportó gasto en esos días.');
    return informe;
  }

  const tiposCompra = {}, tiposResultado = {};
  const filas = r.filas.map(function (f) {
    const fecha = String(f.date_start || '').slice(0, 10);
    const moneda = String(f.account_currency || '').toUpperCase();
    if (moneda) informe.moneda = moneda;

    const compra = metaAccion_(f.actions, META_ACCIONES_COMPRA);
    const lead   = metaAccion_(f.actions, META_ACCIONES_LEAD);
    const lp     = metaAccion_(f.actions, META_ACCION_LP);
    const valor  = metaAccion_(f.action_values, META_ACCIONES_COMPRA);

    if (compra) tiposCompra[compra.tipo] = 1;
    const resultado = compra || lead;
    if (resultado) tiposResultado[resultado.tipo] = 1;

    const gasto = Number(f.spend) || 0;
    informe.gasto += gasto;

    /**
     * El identificador es el MISMO que arma el importador del Excel.
     *
     * A propósito: si alguien sube el export del mismo día y el mismo
     * conjunto, las dos filas son la misma fila y la segunda pisa a la
     * primera en vez de sumarse. Dos formas de traer el dato, un solo
     * número.
     */
    const conjunto = String(f.adset_name || f.campaign_name || 'Sin nombre').trim();
    const id = ['meta', tienda, fecha, fecha, norm(conjunto)].join('-')
      .replace(/\s+/g, '_').slice(0, 180);

    return {
      id: id,
      fecha: fecha,
      fecha_fin: fecha,          // una fila por día: el periodo es el día
      tienda: tienda,
      plataforma: 'meta',
      cuenta: cuenta,
      campana: String(f.campaign_name || '').trim(),
      conjunto: conjunto,
      // `entrega` y `presupuesto` no vienen en este informe. Vacías, no
      // en cero: un cero diría que no había presupuesto.
      gasto: gasto,
      moneda_gasto: moneda,
      impresiones: Number(f.impressions) || 0,
      alcance: Number(f.reach) || 0,
      frecuencia: Number(f.frequency) || 0,
      clics: Number(f.clicks) || 0,
      ctr: Number(f.ctr) || 0,
      cpc: Number(f.cpc) || 0,
      cpm: Number(f.cpm) || 0,
      resultados: resultado ? resultado.valor : 0,
      compras: compra ? compra.valor : 0,
      // CPA y ROAS se calculan aquí porque los dos números que los forman
      // están aquí y son del mismo día. Recalcularlos después obliga a
      // adivinar contra qué gasto se dividían.
      cpa: (resultado && resultado.valor) ? gasto / resultado.valor : '',
      roas: (valor && gasto) ? valor.valor / gasto : '',
      valor_conv: valor ? valor.valor : '',
      visitas_lp: lp ? lp.valor : '',
    };
  });

  informe.accionCompra = Object.keys(tiposCompra).join(', ');
  informe.accionResultado = Object.keys(tiposResultado).join(', ');

  // Si Meta no optimiza a compra, el CPA que sale no es por venta
  if (!informe.accionCompra && informe.accionResultado) {
    informe.avisos.push('Meta no reportó compras en estos conjuntos: el resultado ' +
      'que se contó es «' + informe.accionResultado + '». El CPA de esta pantalla ' +
      'es por ese resultado, no por venta entregada.');
  }
  if (informe.moneda && informe.monedaTienda && informe.moneda !== informe.monedaTienda) {
    informe.avisos.push('Meta cobra en ' + informe.moneda + ' y la tienda factura en ' +
      informe.monedaTienda + '. Nova convierte con la tasa del día de cada gasto; ' +
      'si faltan tasas, ese gasto se cuenta aparte y la pantalla lo dice.');
  }

  let esc;
  try {
    esc = escribirFilas(ss, 'Pauta', filas, 'meta');
  } catch (err) {
    informe.error = err.message;
    return informe;
  }
  informe.nuevas = esc.nuevas;
  informe.actualizadas = esc.actualizadas;
  informe.iguales = esc.iguales;

  informe.solapadas = metaFilasSolapadas_(ss, tienda, desde, hasta);
  if (informe.solapadas) {
    informe.avisos.push(informe.solapadas + ' fila(s) de un Excel subido a mano cubren ' +
      'un rango de varios días dentro de este periodo. Esas SÍ se suman por separado ' +
      'y el gasto quedaría contado dos veces. Bórralas en la hoja Pauta.');
  }

  informe.ok = true;
  return informe;
}

/**
 * Filas de Pauta que cubren un rango de varios días dentro del periodo.
 *
 * Son las únicas que se pueden contar dos veces. Las de un solo día
 * comparten identificador con las que trae la API, así que se pisan
 * entre ellas y no hay riesgo. Estas no, y no se borran solas: borrar
 * datos que alguien subió es una decisión suya, no de Nova.
 */
function metaFilasSolapadas_(ss, tienda, desde, hasta) {
  const sh = ss.getSheetByName('Pauta');
  if (!sh || sh.getLastRow() < 2) return 0;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  let n = 0;
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')]).trim() !== tienda) continue;
    if (norm(d[i][c('plataforma')]) !== 'meta') continue;
    const f1 = aISO(d[i][c('fecha')], 'UTC');
    const f2 = aISO(d[i][c('fecha_fin')], 'UTC');
    if (!f1 || !f2 || f1 === f2) continue;      // de un solo día: no hay riesgo
    if (f2 < desde || f1 > hasta) continue;     // fuera del periodo leído
    n++;
  }
  return n;
}

/**
 * Los anuncios de una tienda, a su propia hoja.
 *
 * No toca Pauta. Nada de lo que escribe aquí entra en ninguna suma de
 * gasto: esta hoja existe para comparar creativos entre sí, no para
 * cuadrar plata. El total de la cuenta sigue saliendo de Pauta.
 */
function metaLeerAnuncios_(sheetId, tienda, token, dias) {
  const ss = SpreadsheetApp.openById(sheetId);
  const cuenta = metaCuenta(ss, tienda);
  const informe = { tienda: tienda, ok: false, error: '', filas: 0,
                    nuevas: 0, actualizadas: 0, iguales: 0, desde: '', hasta: '' };

  if (!cuenta) { informe.error = 'Esta tienda no tiene cuenta publicitaria.'; return informe; }
  if (!ss.getSheetByName('Anuncios')) {
    informe.error = 'Falta la hoja Anuncios. Corre bootstrapTodo() una vez.';
    return informe;
  }

  const hasta = metaFecha_(0);
  const desde = metaFecha_(Math.max(1, dias || META_DIAS_ANUNCIOS));
  informe.desde = desde; informe.hasta = hasta;

  const url = META_API + 'act_' + cuenta + '/insights' +
    '?level=' + META_NIVEL_ANUNCIO +
    '&time_increment=1' +
    '&time_range=' + encodeURIComponent(JSON.stringify({ since: desde, until: hasta })) +
    '&fields=' + META_CAMPOS_ANUNCIO.join(',') +
    '&limit=200' +
    '&access_token=' + encodeURIComponent(token);

  const r = metaPedir_(url, token);
  if (!r.ok) { informe.error = r.error; return informe; }

  informe.filas = r.filas.length;
  if (!r.filas.length) { informe.ok = true; return informe; }

  const filas = r.filas.map(function (f) {
    const fecha = String(f.date_start || '').slice(0, 10);
    const compra = metaAccion_(f.actions, META_ACCIONES_COMPRA);
    const lead   = metaAccion_(f.actions, META_ACCIONES_LEAD);
    const lp     = metaAccion_(f.actions, META_ACCION_LP);
    const valor  = metaAccion_(f.action_values, META_ACCIONES_COMPRA);
    const resultado = compra || lead;
    const gasto = Number(f.spend) || 0;

    // El id lleva el ad_id de Meta, que sí es único y estable. En Pauta
    // no se puede: el informe de conjuntos no siempre trae adset_id.
    return {
      id: ['meta', tienda, fecha, String(f.ad_id || norm(f.ad_name || ''))]
            .join('-').replace(/\s+/g, '_').slice(0, 180),
      fecha: fecha, tienda: tienda, plataforma: 'meta', cuenta: cuenta,
      campana: String(f.campaign_name || '').trim(),
      conjunto: String(f.adset_name || '').trim(),
      anuncio: String(f.ad_name || '').trim(),
      anuncio_id: String(f.ad_id || ''),
      gasto: gasto,
      moneda_gasto: String(f.account_currency || '').toUpperCase(),
      impresiones: Number(f.impressions) || 0,
      alcance: Number(f.reach) || 0,
      frecuencia: Number(f.frequency) || 0,
      clics: Number(f.clicks) || 0,
      ctr: Number(f.ctr) || 0,
      cpc: Number(f.cpc) || 0,
      cpm: Number(f.cpm) || 0,
      resultados: resultado ? resultado.valor : 0,
      compras: compra ? compra.valor : 0,
      cpa: (resultado && resultado.valor) ? gasto / resultado.valor : '',
      valor_conv: valor ? valor.valor : '',
      visitas_lp: lp ? lp.valor : '',
    };
  });

  try {
    const esc = escribirFilas(ss, 'Anuncios', filas, 'meta');
    informe.nuevas = esc.nuevas;
    informe.actualizadas = esc.actualizadas;
    informe.iguales = esc.iguales;
    informe.ok = true;
  } catch (err) {
    informe.error = err.message;
  }
  return informe;
}

/** Todas las tiendas de un cliente. */
function metaLeerCliente_(sheetId, dias) {
  const token = PropertiesService.getScriptProperties()
    .getProperty('META_TOKEN_' + String(sheetId).slice(0, 44)) || '';
  if (!token) return { sheetId: sheetId, sinLlave: true, tiendas: [] };

  const ss = SpreadsheetApp.openById(sheetId);
  const shT = ss.getSheetByName('Tiendas');
  if (!shT || shT.getLastRow() < 2) return { sheetId: sheetId, tiendas: [] };

  const filas = shT.getDataRange().getValues();
  const enc = filas[0].map(norm);
  const cId = enc.indexOf('id'), cEstado = enc.indexOf('estado');

  const out = [];
  filas.slice(1).forEach(function (f) {
    const id = String(f[cId] || '').trim();
    if (!id) return;
    if (cEstado !== -1 && norm(f[cEstado]) === 'inactiva') return;
    if (!metaCuenta(ss, id)) return;            // esa tienda no tiene cuenta conectada
    const r = metaLeerTienda_(sheetId, id, token, dias);
    // Los anuncios solo si esa tienda los quiere. Y si fallan, no
    // arrastran a la lectura de conjuntos: el gasto ya quedó escrito.
    if (norm(ajustes(ss, id).meta_anuncios) === 'si') {
      try {
        r.anuncios = metaLeerAnuncios_(sheetId, id, token, META_DIAS_ANUNCIOS);
        if (!r.anuncios.ok && r.anuncios.error) {
          r.avisos.push('Los conjuntos entraron bien, pero los anuncios no: ' +
                        r.anuncios.error);
        }
      } catch (e) {
        r.avisos.push('Los conjuntos entraron bien, pero los anuncios no: ' + e.message);
      }
    }
    out.push(r);
  });
  return { sheetId: sheetId, tiendas: out };
}

/**
 * Lo que corre el disparador: todos los clientes, todas las mañanas.
 *
 * Un cliente que falla no puede impedir que se lea el siguiente. Se anota
 * el fallo y se sigue.
 */
function leerMetaDiario() {
  const lineas = [];
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try {
      const r = metaLeerCliente_(c.sheetId, META_DIAS_DIARIO);
      if (r.sinLlave) return;                   // no está conectado: no es un fallo
      r.tiendas.forEach(function (t) {
        lineas.push(c.empresa + ' · ' + t.tienda + ': ' +
          (t.ok ? t.filas + ' filas (' + t.nuevas + ' nuevas, ' +
                  t.actualizadas + ' corregidas)' +
                  (t.avisos.length ? '  OJO · ' + t.avisos.join(' ') : '')
                : 'FALLÓ · ' + t.error));
      });
    } catch (e) {
      lineas.push(c.empresa + ': FALLÓ · ' + e.message);
    }
  });

  const msg = lineas.length ? lineas.join('\n') : 'Ninguna cuenta tiene Meta conectado.';
  Logger.log(msg);
  return msg;
}

/**
 * El botón «Traer ahora» de la pantalla.
 *
 * Existe por dos razones distintas. Una: al conectar, nadie quiere
 * esperar hasta mañana para ver si sirve. Dos: para traer el historial
 * de una cuenta vieja, que es lo que hace falta el primer día.
 */
function apiMetaTraer(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña trae la pauta.' };

  const token = metaToken(s);
  if (!token) return { ok: false, error: 'Todavía no has guardado la llave.' };

  /**
   * El tope de días no es un capricho.
   *
   * Meta solo guarda insights de los últimos 37 meses, y pedir de a un
   * día multiplica las llamadas: 400 días son 400 filas por conjunto y
   * varias páginas. El nivel de acceso de una app nueva permite 300
   * llamadas por hora, así que un historial largo se trae por tramos.
   */
  const dias = Math.min(Math.max(Number(p.dias) || META_DIAS_DIARIO, 1), 400);

  const tienda = String(p.tienda || s.tiendas[0]);
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }

  const informe = metaLeerTienda_(s.sheetId, tienda, token, dias);
  if (!informe.ok) return { ok: false, error: informe.error };

  /**
   * Los anuncios se traen con menos días que los conjuntos, aunque se
   * pidan muchos. Un año a nivel de anuncio son miles de filas y varias
   * páginas por cuenta: se llenaría la cuota de la hora y quedaría a
   * medias, que es peor que no traerlo.
   */
  const ss2 = SpreadsheetApp.openById(s.sheetId);
  if (norm(ajustes(ss2, tienda).meta_anuncios) === 'si') {
    try {
      informe.anuncios = metaLeerAnuncios_(s.sheetId, tienda, token, Math.min(dias, 90));
      if (!informe.anuncios.ok && informe.anuncios.error) {
        informe.avisos.push('Los anuncios no se pudieron traer: ' + informe.anuncios.error);
      }
    } catch (e) {
      informe.avisos.push('Los anuncios no se pudieron traer: ' + e.message);
    }
  }

  registrarMovimiento(s, 'Pauta', 'meta', 'traer', tienda,
    informe.nuevas + ' nuevas / ' + informe.actualizadas + ' corregidas');

  return { ok: true, informe: informe };
}
