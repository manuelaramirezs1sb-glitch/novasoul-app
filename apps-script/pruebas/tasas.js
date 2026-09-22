/**
 * Prueba paresDeGasto, paresNecesarios y revisarTasasDe_ con hojas falsas.
 * Se cargan las funciones reales del bundle, con los servicios de Google
 * remedados: lo que se prueba es la lógica, no Sheets.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../NOVA-COMPLETO.gs', 'utf8');

const HOJAS = {};                       // nombre -> matriz
function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  return {
    getLastRow: () => m.length,
    getDataRange: () => ({ getValues: () => m }),
  };
}
const SS = { getSheetByName: hoja };

global.SpreadsheetApp = { openById: () => SS };
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 'soul', ID_ACADEMY: 'aca' };
global.PropertiesService = {
  getScriptProperties: () => ({ getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS }),
};
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [] };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }) };
global.UrlFetchApp = {};
global.Utilities = {
  sleep: () => {},
  // Solo el formato que usan aISO y ahoraISO: yyyy-MM-dd y el ISO completo.
  formatDate: (d, tz, patron) => {
    const iso = new Date(d).toISOString();
    if (patron === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (/HH:mm/.test(patron)) return iso.slice(0, 19);
    return iso;
  },
};
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
global.SpreadsheetApp.flush = () => {};

// El bundle es un script clásico: se evalúa en el ámbito global.
(0, eval)(src + '\n;globalThis.__F = { paresEnUso, paresDeGasto, paresNecesarios, estadoTasasDe_, estadoAutomatico, verAutomatico };');
const F = globalThis.__F;

let fallas = 0;
function caso(nombre, esperado, real) {
  const ok = JSON.stringify(real) === JSON.stringify(esperado);
  if (!ok) fallas++;
  console.log((ok ? '  ok   ' : '  FALLA') + '  ' + nombre);
  if (!ok) { console.log('        esperaba ' + JSON.stringify(esperado));
             console.log('        obtuve   ' + JSON.stringify(real)); }
}
function contiene(nombre, texto, trozo) {
  const ok = String(texto).indexOf(trozo) !== -1;
  if (!ok) fallas++;
  console.log((ok ? '  ok   ' : '  FALLA') + '  ' + nombre);
  if (!ok) console.log('        no encontré «' + trozo + '» en:\n        ' + texto);
}
function montar(reporte, tiendas, pauta, tasas) {
  Object.keys(HOJAS).forEach(k => delete HOJAS[k]);
  HOJAS.Parametros = [['tienda','clave','valor','actualizado_en','actualizado_por'],
                      ['', 'moneda_reporte', reporte, '', '']];
  HOJAS.Tiendas = [['id','nombre','marca','pais','sociedad','nit','moneda',
                    'zona_horaria','corte_despacho','modalidad','estado']]
                  .concat(tiendas);
  HOJAS.Pauta = [['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
                  'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado',
                  'impresiones','alcance','frecuencia','clics','ctr','cpc','cpm',
                  'resultados','compras','cpa','roas','valor_conv','visitas_lp']]
                 .concat(pauta);
  if (tasas) HOJAS.Tasas = [['fecha','moneda_origen','moneda_destino','tasa']].concat(tasas);
  // listarClientes la lee para saber a quién hay que revisar. Una sola
  // cuenta: el stub devuelve las mismas hojas para cualquier sheetId.
  const cl = new Array(14).fill('');
  cl[0] = 'c1'; cl[1] = 'La cuenta'; cl[6] = 'activo'; cl[12] = 1; cl[13] = 'hoja-1';
  HOJAS.Clientes = [['id','empresa','pais','plan','tarifa','costo','estado','fecha_alta',
                     'fecha_corte','ultimo_pago','nota','x','tiendas','sheet_id'], cl];
}
const T = (id, mon, estado) => [id, id, '', '', '', '', mon, 'America/Bogota', '', 'contraentrega', estado || 'activa'];
const P = (tienda, mon, fecha) => { const f = new Array(26).fill(''); f[1] = fecha || '2026-09-01'; f[3] = tienda; f[10] = 100; f[11] = mon; return f; };

const hoy = new Date().toISOString().slice(0, 10);
const hace = n => new Date(Date.now() - n * 86400000).toISOString().slice(0, 10);

console.log('\nPARES DE GASTO — el caso que se caía');
// Tienda GTQ, reporta en GTQ, Meta le cobra en USD.
// paresEnUso dice que no hace falta nada. paresDeGasto ve el hueco.
montar('GTQ', [T('Tienda GT', 'GTQ')], [P('Tienda GT', 'USD')]);
caso('paresEnUso no ve nada (y es correcto: tienda y reporte coinciden)', [], F.paresEnUso(SS));
caso('paresDeGasto ve USD->GTQ', [{ origen: 'USD', destino: 'GTQ' }], F.paresDeGasto(SS));
caso('paresNecesarios los junta', [{ origen: 'USD', destino: 'GTQ' }], F.paresNecesarios(SS));

console.log('\nEL CASO DE NUTREA — dos tiendas, reporte en COP');
montar('COP', [T('Nutrea EC', 'USD'), T('Nutrea GT', 'GTQ')],
       [P('Nutrea EC', 'COP'), P('Nutrea GT', 'COP')]);
caso('paresEnUso: las dos tiendas contra el reporte',
     [{ origen: 'USD', destino: 'COP' }, { origen: 'GTQ', destino: 'COP' }], F.paresEnUso(SS));
caso('paresDeGasto: el gasto en COP contra cada tienda',
     [{ origen: 'COP', destino: 'USD' }, { origen: 'COP', destino: 'GTQ' }], F.paresDeGasto(SS));
caso('paresNecesarios NO repite el inverso — buscarTasa sabe dividir',
     [{ origen: 'USD', destino: 'COP' }, { origen: 'GTQ', destino: 'COP' }], F.paresNecesarios(SS));

console.log('\nNADA QUE CONVERTIR');
montar('COP', [T('Tienda CO', 'COP')], [P('Tienda CO', 'COP')]);
caso('todo en la misma moneda: ningún par', [], F.paresNecesarios(SS));

console.log('\nTIENDA INACTIVA Y FILAS SUCIAS');
montar('COP', [T('Vieja', 'PEN', 'inactiva'), T('Tienda CO', 'COP')],
       [P('Tienda CO', 'COP'), P('', 'USD'), P('Tienda CO', ''), P('Fantasma', 'USD')]);
caso('la tienda inactiva no arrastra su par, y las filas sin tienda o sin moneda se ignoran',
     [], F.paresNecesarios(SS));

console.log('\nESTADO DE LAS TASAS — lo que ve la consola');
montar('GTQ', [T('Tienda GT', 'GTQ')], [P('Tienda GT', 'USD')]);   // sin hoja Tasas
let e = F.estadoTasasDe_('x', 'Tienda GT');
caso('sin hoja Tasas: las necesita y no está al día',
     { necesita: true, alDia: false, dias: -1 },
     { necesita: e.necesita, alDia: e.alDia, dias: e.pares[0].dias });

montar('COP', [T('Tienda CO', 'COP')], [P('Tienda CO', 'COP')]);
e = F.estadoTasasDe_('x', 'Tienda CO');
caso('quien no las necesita queda al día, no en falta',
     { necesita: false, alDia: true, pares: 0 },
     { necesita: e.necesita, alDia: e.alDia, pares: e.pares.length });

montar('GTQ', [T('Tienda GT', 'GTQ')], [P('Tienda GT', 'USD')],
       [[hoy, 'USD', 'GTQ', 7.8]]);
e = F.estadoTasasDe_('x', 'Tienda GT');
caso('con la tasa de hoy: al día', { alDia: true, bien: true, dias: 0 },
     { alDia: e.alDia, bien: e.pares[0].bien, dias: e.pares[0].dias });

montar('GTQ', [T('Tienda GT', 'GTQ')], [P('Tienda GT', 'USD')],
       [[hace(20), 'USD', 'GTQ', 7.8]]);
e = F.estadoTasasDe_('x', 'Tienda GT');
caso('con la tasa de hace 20 días: atrasada', { alDia: false, dias: 20 },
     { alDia: e.alDia, dias: e.pares[0].dias });

montar('GTQ', [T('Tienda GT', 'GTQ')], [P('Tienda GT', 'USD')],
       [[hoy, 'GTQ', 'USD', 0.128]]);
caso('el par inverso cuenta como cargado', true, F.estadoTasasDe_('x', 'Tienda GT').alDia);

// EL CASO QUE JUSTIFICA LA REVISIÓN POR PAR:
// un par al día tapando otro que no tiene ni una fila.
montar('COP', [T('Nutrea EC', 'USD'), T('Nutrea GT', 'GTQ')],
       [P('Nutrea EC', 'COP'), P('Nutrea GT', 'COP')],
       [[hoy, 'USD', 'COP', 4000]]);
e = F.estadoTasasDe_('x', 'Nutrea');
caso('un par al día no tapa al otro', { alDia: false, uno: true, otro: -1 },
     { alDia: e.alDia, uno: e.pares[0].bien, otro: e.pares[1].dias });

console.log('\nEL TEXTO DEL APPS SCRIPT dice lo mismo que los datos');
const texto = F.verAutomatico();
contiene('nombra el par que falta', texto, 'GTQ→COP SIN NINGUNA TASA');
contiene('manda a prenderAutomatico', texto, 'prenderAutomatico');
contiene('los trabajos apagados salen marcados', texto, '✗');

console.log(fallas ? '\n' + fallas + ' FALLA(S)\n' : '\nTodo pasa.\n');
process.exit(fallas ? 1 : 0);
