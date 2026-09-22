/**
 * Cuántas llamadas le hace Nova a Meta. Contadas, no estimadas.
 *
 * Se corre el código real del bundle con UrlFetchApp remedado: el
 * remedo devuelve tantas filas como tendría una cuenta de verdad y
 * pagina igual que Meta, así que el contador cuenta llamadas reales.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const HOJAS = {};
function hoja(n) {
  const m = HOJAS[n];
  if (!m) return null;
  return {
    getLastRow: () => m.length, getLastColumn: () => (m[0] ? m[0].length : 0),
    getDataRange: () => ({ getValues: () => m }),
    getRange: (f, c, nf, nc) => ({
      getValues: () => { const o = []; for (let i = 0; i < (nf || 1); i++) {
        const fila = m[f - 1 + i] || []; o.push(fila.slice(c - 1, c - 1 + (nc || fila.length))); } return o; },
      setValues: (v) => { v.forEach((r, i) => { const y = f - 1 + i;
        while (m.length <= y) m.push(new Array(m[0].length).fill(''));
        r.forEach((x, j) => { m[y][c - 1 + j] = x; }); }); },
      setValue: () => {},
    }),
  };
}
const SS = { getSheetByName: hoja };

let LLAMADAS = 0;
let FILAS_POR_DIA = 5;      // cuántos conjuntos (o anuncios) activos tiene la cuenta
let LIMITE = 500;           // el `limit` que Nova pide

const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: k => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: (k, v) => { PROPS[k] = v; }, deleteProperty: k => { delete PROPS[k]; } }) };
global.SpreadsheetApp = { openById: () => SS, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' }, WeekDay: { MONDAY: 'M' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.Utilities = { sleep: () => {}, getUuid: () => 'u',
  formatDate: (d, tz, p) => { const i = new Date(d).toISOString();
    return p === 'yyyy-MM-dd' ? i.slice(0, 10) : p === 'yyyy-MM' ? i.slice(0, 7) : i; } };

/**
 * Meta, remedado: devuelve `dias × FILAS_POR_DIA` filas en páginas de
 * LIMITE, exactamente como lo hace la API de verdad.
 */
global.UrlFetchApp = {
  fetch: (url) => {
    LLAMADAS++;
    const m = url.match(/since[^0-9]*(\d{4}-\d{2}-\d{2})[^0-9]*until[^0-9]*(\d{4}-\d{2}-\d{2})/) ||
              url.match(/%22since%22%3A%22(\d{4}-\d{2}-\d{2})%22%2C%22until%22%3A%22(\d{4}-\d{2}-\d{2})%22/);
    let total;
    if (url.indexOf('__pag=') !== -1) {
      total = Number(url.match(/__tot=(\d+)/)[1]);
    } else {
      const desde = m ? new Date(m[1] + 'T00:00:00Z') : new Date();
      const hasta = m ? new Date(m[2] + 'T00:00:00Z') : new Date();
      const dias = Math.round((hasta - desde) / 86400000) + 1;
      total = dias * FILAS_POR_DIA;
    }
    const servidas = Number((url.match(/__srv=(\d+)/) || [0, 0])[1]);
    const quedan = total - servidas;
    const n = Math.min(LIMITE, Math.max(quedan, 0));
    const data = [];
    for (let i = 0; i < n; i++) {
      data.push({ date_start: '2026-09-01', date_stop: '2026-09-01', account_currency: 'GTQ',
        campaign_name: 'C', adset_name: 'A' + i, adset_id: 'i' + i,
        ad_name: 'AD' + i, ad_id: 'ad' + i, spend: '10', impressions: '1', reach: '1',
        frequency: '1', clicks: '1', ctr: '1', cpc: '1', cpm: '1', actions: [], action_values: [] });
    }
    const out = { data: data };
    if (servidas + n < total) {
      out.paging = { next: 'https://graph.facebook.com/n?__pag=1&__tot=' + total +
                           '&__srv=' + (servidas + n) };
    }
    return { getContentText: () => JSON.stringify(out) };
  },
};

(0, eval)(src + '\n;globalThis.__F = { metaLeerTienda_, metaLeerAnuncios_, META_DIAS_DIARIO, META_DIAS_ANUNCIOS };');
const F = globalThis.__F;

const PAU = ['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
  'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado','impresiones','alcance',
  'frecuencia','clics','ctr','cpc','cpm','resultados','compras','cpa','roas','valor_conv','visitas_lp'];
const ANU = ['id','fecha','tienda','plataforma','cuenta','campana','conjunto','anuncio',
  'anuncio_id','gasto','moneda_gasto','impresiones','alcance','frecuencia','clics','ctr',
  'cpc','cpm','resultados','compras','cpa','valor_conv','visitas_lp'];

function montar() {
  Object.keys(HOJAS).forEach(k => delete HOJAS[k]);
  HOJAS.Parametros = [['tienda','clave','valor','actualizado_en','actualizado_por'],
                      ['', 'moneda_reporte', 'GTQ', '', ''], ['gt', 'meta_cuenta', '123', '', '']];
  HOJAS.Tiendas = [['id','nombre','marca','pais','sociedad','nit','moneda','zona_horaria',
                    'corte_despacho','modalidad','estado'],
                   ['gt','Nutrea GT','','GT','','','GTQ','UTC','','contraentrega','activa']];
  HOJAS.Pauta = [PAU];
  HOJAS.Anuncios = [ANU];
  HOJAS.Movimientos = [['fecha','quien','entidad','id','campo','antes','ahora']];
}

function medir(etiqueta, fn, activos) {
  FILAS_POR_DIA = activos;
  montar();
  LLAMADAS = 0;
  const r = fn();
  const filas = r.filas;
  console.log('  ' + etiqueta.padEnd(46) +
    String(filas).padStart(6) + ' filas  ' +
    String(LLAMADAS).padStart(4) + ' llamada' + (LLAMADAS === 1 ? '' : 's'));
  return { llamadas: LLAMADAS, filas: filas };
}

console.log('\nUNA CUENTA CHICA · 5 conjuntos, 12 anuncios activos');
const d1 = medir('lectura diaria · conjuntos (7 días)',
  () => F.metaLeerTienda_('h', 'gt', 'T', F.META_DIAS_DIARIO), 5);
const d2 = medir('lectura diaria · anuncios (3 días)',
  () => F.metaLeerAnuncios_('h', 'gt', 'T', F.META_DIAS_ANUNCIOS), 12);
console.log('  ' + 'TOTAL DE UNA MAÑANA'.padEnd(46) + ' '.repeat(13) +
  String(d1.llamadas + d2.llamadas).padStart(4) + ' llamadas');

console.log('\nUNA CUENTA GRANDE · 40 conjuntos, 150 anuncios activos');
const g1 = medir('lectura diaria · conjuntos (7 días)',
  () => F.metaLeerTienda_('h', 'gt', 'T', F.META_DIAS_DIARIO), 40);
const g2 = medir('lectura diaria · anuncios (3 días)',
  () => F.metaLeerAnuncios_('h', 'gt', 'T', F.META_DIAS_ANUNCIOS), 150);
console.log('  ' + 'TOTAL DE UNA MAÑANA'.padEnd(46) + ' '.repeat(13) +
  String(g1.llamadas + g2.llamadas).padStart(4) + ' llamadas');

console.log('\nEL HISTORIAL, UNA SOLA VEZ · cuenta grande');
medir('conjuntos · 90 días', () => F.metaLeerTienda_('h', 'gt', 'T', 90), 40);
medir('conjuntos · 365 días', () => F.metaLeerTienda_('h', 'gt', 'T', 365), 40);
medir('anuncios · 90 días', () => F.metaLeerAnuncios_('h', 'gt', 'T', 90), 150);
const a365 = medir('anuncios · 365 días', () => F.metaLeerAnuncios_('h', 'gt', 'T', 365), 150);

console.log('\nEL TOPE DE 40 PÁGINAS · ¿avisa o se calla?');
FILAS_POR_DIA = 150; montar(); LLAMADAS = 0;
const r = F.metaLeerAnuncios_('h', 'gt', 'T', 365);
const esperadas = 366 * 150;
console.log('  filas que debería traer:   ' + esperadas);
console.log('  filas que trajo:           ' + r.filas);
console.log('  llamadas:                  ' + LLAMADAS);
console.log('  ¿lo dice?                  ' +
  (r.filas < esperadas ? (r.error || (r.avisos || []).join(' ') || '*** NO DICE NADA ***')
                       : 'no hizo falta'));
