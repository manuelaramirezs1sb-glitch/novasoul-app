/**
 * La lectura de Meta, con Meta remedado.
 *
 * Se carga el bundle real con los servicios de Google y UrlFetchApp
 * falsos: lo que se prueba es qué le pide Nova a Meta, qué escribe en
 * Pauta y qué informa — sin tocar una cuenta de verdad ni gastar
 * llamadas de la cuota.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const HOJAS = {};
const PEDIDOS = [];          // las URL que Nova le pidió a Meta
let RESPUESTAS = [];         // lo que Meta contesta, en orden

function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  return {
    getLastRow: () => m.length,
    getLastColumn: () => (m[0] ? m[0].length : 0),
    getDataRange: () => ({ getValues: () => m }),
    getRange: (fila, col, nFilas, nCols) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < (nFilas || 1); i++) {
          const f = m[fila - 1 + i] || [];
          out.push(f.slice(col - 1, col - 1 + (nCols || f.length)));
        }
        return out;
      },
      setValues: (vals) => {
        vals.forEach((v, i) => {
          const r = fila - 1 + i;
          while (m.length <= r) m.push(new Array(m[0].length).fill(''));
          v.forEach((x, j) => { m[r][col - 1 + j] = x; });
        });
      },
      setValue: () => {},
      setFormula: () => {},
    }),
    clear: () => {},
    hideSheet: () => {},
    getParent: () => SS,
  };
}
const SS = { getSheetByName: hoja, insertSheet: (n) => { HOJAS[n] = [[]]; return hoja(n); } };

const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a',
                'META_TOKEN_hoja-1': 'EAA' + 'x'.repeat(200) };
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: (k, v) => { PROPS[k] = v; }, deleteProperty: (k) => { delete PROPS[k]; },
}) };
global.SpreadsheetApp = { openById: () => SS, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.Utilities = {
  sleep: () => {}, getUuid: () => 'uuid-0000',
  formatDate: (d, tz, patron) => {
    const iso = new Date(d).toISOString();
    if (patron === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (patron === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(patron)) return iso.slice(0, 19);
    return iso;
  },
};
global.UrlFetchApp = {
  fetch: (url) => {
    PEDIDOS.push(url);
    const r = RESPUESTAS.shift() || { data: [] };
    return { getContentText: () => JSON.stringify(r) };
  },
};

(0, eval)(src + '\n;globalThis.__F = { metaLeerTienda_, metaAccion_, metaPedir_, TRABAJOS, estadoAutomatico };');
const F = globalThis.__F;

let fallas = 0;
function ok(nombre, cond, detalle) {
  if (cond) { console.log('  ok     ' + nombre); return; }
  fallas++; console.log('  FALLA  ' + nombre + (detalle ? '\n         ' + detalle : ''));
}

const PAUTA_ENC = ['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
  'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado','impresiones','alcance',
  'frecuencia','clics','ctr','cpc','cpm','resultados','compras','cpa','roas','valor_conv','visitas_lp'];

function montar(pautaExtra) {
  Object.keys(HOJAS).forEach(k => delete HOJAS[k]);
  HOJAS.Parametros = [['tienda','clave','valor','actualizado_en','actualizado_por'],
                      ['', 'moneda_reporte', 'COP', '', ''],
                      ['gt', 'meta_cuenta', '123456789', '', '']];
  HOJAS.Tiendas = [['id','nombre','marca','pais','sociedad','nit','moneda',
                    'zona_horaria','corte_despacho','modalidad','estado'],
                   ['gt', 'Nutrea GT', '', 'GT', '', '', 'GTQ', 'UTC', '', 'contraentrega', 'activa']];
  HOJAS.Pauta = [PAUTA_ENC].concat(pautaExtra || []);
  HOJAS.Movimientos = [['fecha','quien','entidad','id','campo','antes','ahora']];
  PEDIDOS.length = 0;
}

const ayer = new Date(Date.now() - 86400000).toISOString().slice(0, 10);
const anteayer = new Date(Date.now() - 2 * 86400000).toISOString().slice(0, 10);

function filaMeta(fecha, conjunto, spend, acciones) {
  return {
    date_start: fecha, date_stop: fecha, account_currency: 'COP',
    campaign_name: 'Camp 1', adset_name: conjunto, adset_id: 'a1',
    spend: String(spend), impressions: '1000', reach: '800', frequency: '1.25',
    clicks: '50', ctr: '5', cpc: '100', cpm: '5000',
    actions: acciones || [], action_values: [],
  };
}

console.log('\nLO QUE SE LE PIDE A META');
montar();
RESPUESTAS = [{ data: [filaMeta(ayer, 'Frío', 50000, [{ action_type: 'offsite_conversion.fb_pixel_purchase', value: '5' }])] }];
let i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
const url = PEDIDOS[0] || '';
ok('pide una fila por día (time_increment=1)', url.includes('time_increment=1'), url.slice(0, 160));
ok('pide a nivel de conjunto', url.includes('level=adset'));
ok('pide la moneda de la cuenta', url.includes('account_currency'));
ok('la llave viaja en la URL y no en el cuerpo', url.includes('access_token='));

console.log('\nLO QUE SE ESCRIBE EN PAUTA');
const fila = HOJAS.Pauta[1];
const col = (n) => fila[PAUTA_ENC.indexOf(n)];
ok('fecha y fecha_fin son el mismo día', col('fecha') === ayer && col('fecha_fin') === ayer,
   col('fecha') + ' / ' + col('fecha_fin'));
ok('la moneda la declara Meta, no se adivina', col('moneda_gasto') === 'COP');
ok('el gasto entra como número', col('gasto') === 50000, String(col('gasto')));
ok('las compras salen de la acción del píxel', col('compras') === 5, String(col('compras')));
ok('el CPA se calcula con el gasto de ese mismo día', col('cpa') === 10000, String(col('cpa')));
ok('presupuesto queda VACÍO, no en cero', col('presupuesto') === '', JSON.stringify(col('presupuesto')));
ok('entrega queda VACÍA, no en cero', col('entrega') === '', JSON.stringify(col('entrega')));
ok('el informe dice qué contó como compra',
   i.accionCompra === 'offsite_conversion.fb_pixel_purchase', i.accionCompra);

console.log('\nVOLVER A LEER NO DUPLICA');
RESPUESTAS = [{ data: [filaMeta(ayer, 'Frío', 50000, [{ action_type: 'offsite_conversion.fb_pixel_purchase', value: '5' }])] }];
const antes = HOJAS.Pauta.length;
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('la misma lectura no agrega filas', HOJAS.Pauta.length === antes,
   antes + ' → ' + HOJAS.Pauta.length);
ok('y el informe lo dice: 0 nuevas', i.nuevas === 0, 'nuevas=' + i.nuevas);

console.log('\nMETA CORRIGE SUS PROPIOS NÚMEROS');
// Meta sigue atribuyendo conversiones: el mismo día vuelve con más compras.
RESPUESTAS = [{ data: [filaMeta(ayer, 'Frío', 50000, [{ action_type: 'offsite_conversion.fb_pixel_purchase', value: '8' }])] }];
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('la fila se corrige en vez de duplicarse', HOJAS.Pauta.length === antes);
ok('y queda el número nuevo', HOJAS.Pauta[1][PAUTA_ENC.indexOf('compras')] === 8,
   String(HOJAS.Pauta[1][PAUTA_ENC.indexOf('compras')]));
ok('el informe la cuenta como corregida', i.actualizadas === 1, 'actualizadas=' + i.actualizadas);

console.log('\nSIN COMPRAS: NO SE INVENTA UNA');
montar();
RESPUESTAS = [{ data: [filaMeta(ayer, 'Registro', 30000, [{ action_type: 'lead', value: '10' }])] }];
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('compras queda en 0', HOJAS.Pauta[1][PAUTA_ENC.indexOf('compras')] === 0);
ok('pero el resultado sí se cuenta', HOJAS.Pauta[1][PAUTA_ENC.indexOf('resultados')] === 10);
ok('y avisa que ese CPA no es por venta',
   i.avisos.some(a => a.includes('no es por venta') || a.includes('no por venta')),
   JSON.stringify(i.avisos));

console.log('\nLA FUGA: UN EXCEL DE VARIOS DÍAS QUE SE SOLAPA');
// Una fila subida a mano que cubre un rango NO comparte identificador con
// las diarias, así que se sumaría además de ellas.
const rango = new Array(PAUTA_ENC.length).fill('');
rango[0] = 'meta-gt-' + anteayer + '-' + ayer + '-frio';
rango[1] = anteayer; rango[2] = ayer; rango[3] = 'gt'; rango[4] = 'meta';
rango[7] = 'Frío'; rango[10] = 99999; rango[11] = 'COP';
montar([rango]);
RESPUESTAS = [{ data: [filaMeta(ayer, 'Frío', 50000, [])] }];
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('la detecta', i.solapadas === 1, 'solapadas=' + i.solapadas);
ok('lo dice en el informe', i.avisos.some(a => a.includes('dos veces')), JSON.stringify(i.avisos));
ok('y NO la borra sola', HOJAS.Pauta.some(f => f[0] === rango[0]));

console.log('\nMETA CONTESTA UN ERROR');
montar();
RESPUESTAS = [{ error: { code: 190, message: 'Session expired' } }];
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('no escribe nada', HOJAS.Pauta.length === 1);
ok('y explica que hay que cambiar la llave', /llave ya no sirve/i.test(i.error), i.error);

console.log('\nRESPUESTA PARTIDA EN PÁGINAS');
montar();
RESPUESTAS = [
  { data: [filaMeta(ayer, 'A', 100, [])], paging: { next: 'https://graph.facebook.com/siguiente' } },
  { data: [filaMeta(ayer, 'B', 200, [])] },
];
i = F.metaLeerTienda_('hoja-1', 'gt', 'TOKEN', 7);
ok('sigue el enlace y no pierde la segunda página', i.filas === 2, 'filas=' + i.filas);
ok('escribe las dos', HOJAS.Pauta.length === 3, String(HOJAS.Pauta.length - 1));

console.log('\nEL ORDEN DE LAS HORAS');
const horas = F.TRABAJOS.map(t => t.fn + '@' + t.hora);
const hTasas = F.TRABAJOS.filter(t => t.fn === 'actualizarTasasDiario')[0].hora;
const hMeta  = F.TRABAJOS.filter(t => t.fn === 'leerMetaDiario')[0].hora;
const hAlarm = F.TRABAJOS.filter(t => t.fn === 'revisarAlarmasTodos')[0].hora;
ok('las tasas van antes que Meta', hTasas < hMeta, horas.join(' · '));
ok('y las alarmas después de Meta', hMeta < hAlarm, horas.join(' · '));

console.log(fallas ? '\n' + fallas + ' FALLA(S)\n' : '\nTodo pasa.\n');
process.exit(fallas ? 1 : 0);
