/**
 * POR QUÉ SE DEMORA EN CARGAR. Medido, no opinado.
 *
 * Ella preguntó: «arregla eso de que se está demorando en cargar, ¿por
 * qué es? ¿qué le falta? ¿o qué le sobra?».
 *
 * Al entrar, la pantalla dispara TRECE peticiones en paralelo. Cada una
 * es una petición completa a Apps Script, y cada una abre el libro y
 * casi todas leen la hoja Pedidos ENTERA con `getDataRange()`.
 *
 * Esta prueba cuenta dos cosas que nadie había contado:
 *   · cuántas veces se lee cada hoja de punta a punta en un solo
 *     arranque;
 *   · cuántas CELDAS se mueven en total.
 *
 * No propone una cifra ideal. Cuenta lo que hay, para que la decisión
 * de qué juntar se tome mirando el número y no la intuición.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

/** Cuántos pedidos tiene una cuenta que lleva un rato trabajando. */
const PEDIDOS = Number(process.env.PEDIDOS || 4000);
const NOVEDADES = Math.round(PEDIDOS * 0.35);

const LECTURAS = {};     // hoja → cuántas veces se leyó entera
const CELDAS = {};       // hoja → cuántas celdas se movieron
let ABIERTAS = 0;        // cuántas veces se abrió el libro

const HOJAS = {};
function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  const ancho = m[0] ? m[0].length : 0;
  return {
    getName: () => nombre,
    getLastRow: () => m.length,
    getLastColumn: () => ancho,
    getDataRange: () => ({
      getValues: () => {
        LECTURAS[nombre] = (LECTURAS[nombre] || 0) + 1;
        CELDAS[nombre] = (CELDAS[nombre] || 0) + m.length * ancho;
        return m;
      },
    }),
    appendRow: () => {},
    getRange: (f, c, nf, nc) => ({
      getValues: () => {
        const filas = nf || 1, cols = nc || ancho;
        CELDAS[nombre] = (CELDAS[nombre] || 0) + filas * cols;
        const o = [];
        for (let i = 0; i < filas; i++) {
          const fila = m[f - 1 + i] || [];
          o.push(fila.slice(c - 1, c - 1 + cols));
        }
        return o;
      },
      setValues: () => {}, setValue: () => {},
    }),
  };
}
const SS = { getSheetByName: hoja, insertSheet: () => hoja('Pedidos') };

const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T14:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: () => { ABIERTAS++; return SS; }, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{}' }) };
global.Utilities = { sleep: () => {}, getUuid: () => 'u1',
  formatDate: (d, tz, pat) => {
    const iso = (d && d.getTime && d.getTime() === new Date(HOY).getTime())
      ? HOY : new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(pat)) return iso.slice(0, 19).replace('T', ' ');
    return iso;
  } };
const RealDate = Date;
global.Date = class extends RealDate {
  constructor(...a) { return a.length ? new RealDate(...a) : new RealDate(HOY); }
  static now() { return new RealDate(HOY).getTime(); }
  static UTC(...a) { return RealDate.UTC(...a); }
};

(0, eval)(src + '\n;globalThis.__F = { apiListar, apiEquipo, apiAuditoria, apiAlarmas,' +
  ' apiRecuento, apiProductos, apiReporteDia, apiHistorial, apiCas, apiEstados,' +
  ' apiMetaEstado, apiSemaforo };');
const F = globalThis.__F;

// ── Una cuenta como la suya, con historia ──
const C_PED = ['id','fuente','id_externo','fecha','tienda','cliente','cedula','correo',
  'telefono','telefono_norm','telefono_2','telefono_2_norm','ciudad','departamento',
  'direccion','producto','sku','cantidad','valor','costo_producto','costo_envio',
  'metodo_pago','bodega','estado','estado_transportadora','estado_canonico',
  'transportadora','guia','intentos','gestora_asignada','fecha_promesa','fecha_entrega',
  'razon_cancelacion','estado_nova','nota','ultimo_movimiento','adelanto',
  'acuerdo_oficina','confirmado_oficina','actualizado_en','actualizado_por'];

function construir() {
  const peds = [C_PED];
  for (let i = 0; i < PEDIDOS; i++) {
    const f = new RealDate(RealDate.parse('2026-09-24') - (i % 120) * 86400000)
      .toISOString().slice(0, 10);
    const fila = C_PED.map(() => '');
    fila[C_PED.indexOf('id')] = 'dropi-' + i;
    fila[C_PED.indexOf('tienda')] = 'ec';
    fila[C_PED.indexOf('fecha')] = f;
    fila[C_PED.indexOf('cliente')] = 'Clienta ' + i;
    fila[C_PED.indexOf('producto')] = i % 3 ? 'TAG RECEDE' : 'DR MELAXIN';
    fila[C_PED.indexOf('valor')] = 50;
    fila[C_PED.indexOf('estado_canonico')] =
      ['entregado', 'devolucion', 'en_transito', 'pendiente', 'novedad'][i % 5];
    peds.push(fila);
  }
  HOJAS.Pedidos = peds;

  const novs = [['id','fuente','id_externo','pedido_id','fecha','tipo','motivo','grupo',
                 'estado','solucionada','fecha_solucion','desenlace','gestora','solucion',
                 'nota','intentos','resuelta_en','actualizado_en','actualizado_por']];
  for (let i = 0; i < NOVEDADES; i++) {
    novs.push(['nv' + i, 'dropi', String(i), 'dropi-' + i, '2026-09-20', '', 'No contesta',
               'no_contacta', i % 3 ? 'abierta' : 'resuelta', '', '', '', '', '', '', '',
               '', '', '']);
  }
  HOJAS.Novedades = novs;

  HOJAS.Tiendas = [['id','nombre','moneda','estado','modalidad'],
                   ['ec','Nutrea EC','USD','activa','marca_propia']];
  HOJAS.Parametros = [['tienda','clave','valor','nota','tipo'],
                      ['ec','cpa_verde_pct','70','','']];
  HOJAS.Equipo = [['id','nombre','correo','rol','tienda','estado','casos_asignados',
                   'casos_resueltos','nota_auditoria','ultima_conexion','permisos'],
                  ['e1','Manuela','m@nova.com','dueno','*','activo','','','','','']];
  HOJAS.Inventario = [['id','sku','producto','tienda','fuente','origen','categoria',
    'proveedor','landing','stock','costo_unitario','precio','precio_2','precio_3',
    'minimo','resp_1','resp_2','resp_3','resp_4','dias_cobertura','ultimo_conteo',
    'nota','activo','actualizado_en','actualizado_por'],
    ['i1','TR-001','TAG RECEDE','ec','','manual','estrella','','','40','6','26','','',
     '10','','','','','','2026-09-01','','si','','']];
  HOJAS.Pauta = [['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana',
    'conjunto','entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado',
    'impresiones','alcance','frecuencia','clics','ctr','cpc','cpm','resultados',
    'compras','cpa','roas','valor_conv','visitas_lp']];
  HOJAS.Estados = [['fuente','texto','estado_nova','origen','pedidos','primera_vez',
                    'ultima_vez','decidido_por','nota']];
  HOJAS.Movimientos = [['cuando','quien','entidad','entidad_id','campo','antes','ahora']];
  HOJAS.Cierres = [['tienda','mes','cerrado_en']];
  HOJAS.Gastos = [['id','tienda','mes','tipo','nombre','valor','moneda','nota']];
  HOJAS.Tasas = [['fecha','de','a','tasa']];
  HOJAS.Llamadas = [['id','fuente','id_externo','fecha_hora','tienda','sentido','estado',
    'extension','agente','telefono','telefono_norm','pedido_id','seg_conversado',
    'seg_espera','seg_total','campana','grabacion','etiqueta','observacion']];
  HOJAS.Fuentes = [['id','nombre','tipo','estado']];
  HOJAS.Alarmas = [['id','tienda','tipo','nivel','titulo','accion','estado','creada_en']];
}

const S = { sheetId: 'emp', rol: 'dueno', email: 'm@nova.com', nombre: 'Manuela',
            tiendas: ['ec'], permisos: [], modulos: ['empresarial','pauta','dinero'] };

/** Las trece que dispara la pantalla al entrar, en el mismo orden. */
const ARRANQUE = [
  ['listar · Pedidos',   () => F.apiListar(S, { entidad: 'Pedidos', tienda: 'ec', limite: 300 })],
  ['listar · Novedades', () => F.apiListar(S, { entidad: 'Novedades', tienda: 'ec', limite: 300 })],
  ['equipo',             () => F.apiEquipo(S, { tienda: 'ec' })],
  ['auditoria',          () => F.apiAuditoria(S, { tienda: 'ec' })],
  ['alarmas',            () => F.apiAlarmas(S, { tienda: 'ec' })],
  ['recuento',           () => F.apiRecuento(S, { tienda: 'ec' })],
  ['productos',          () => F.apiProductos(S, { tienda: 'ec' })],
  ['reporte_dia',        () => F.apiReporteDia(S, { tienda: 'ec' })],
  ['historial',          () => F.apiHistorial(S, { tienda: 'ec' })],
  ['cas',                () => F.apiCas(S, { tienda: 'ec' })],
  ['estados',            () => F.apiEstados(S, { tienda: 'ec' })],
  ['meta_estado',        () => F.apiMetaEstado(S, { tienda: 'ec' })],
  ['semaforo',           () => F.apiSemaforo(S, { tienda: 'ec' })],
];

console.log('\nUNA CUENTA CON ' + PEDIDOS.toLocaleString('es') + ' PEDIDOS Y ' +
            NOVEDADES.toLocaleString('es') + ' NOVEDADES\n');
const col = (s, n, der) => {
  s = String(s);
  const hueco = Math.max(0, n - s.length);
  return der ? ' '.repeat(hueco) + s : s + ' '.repeat(hueco);
};
console.log(col('LLAMADA', 22) + col('ms', 7, 1) + col('lee Pedidos', 13, 1) +
            col('celdas', 14, 1));
console.log('─'.repeat(56));

let totalMs = 0, totalPed = 0, totalCeldas = 0, fallos = [];
ARRANQUE.forEach(function (par) {
  construir();
  Object.keys(LECTURAS).forEach(k => delete LECTURAS[k]);
  Object.keys(CELDAS).forEach(k => delete CELDAS[k]);
  const t0 = RealDate.now();
  try { par[1](); } catch (e) { fallos.push(par[0] + ': ' + e.message); }
  const ms = RealDate.now() - t0;
  const ped = LECTURAS.Pedidos || 0;
  const celdas = Object.keys(CELDAS).reduce((a, k) => a + CELDAS[k], 0);
  totalMs += ms; totalPed += ped; totalCeldas += celdas;
  console.log(col(par[0], 22) + col(ms, 7, 1) + col(ped || '—', 13, 1) +
              col(celdas.toLocaleString('es'), 14, 1));
});

console.log('─'.repeat(56));
console.log(col('TOTAL', 22) + col(totalMs, 7, 1) + col(totalPed, 13, 1) +
            col(totalCeldas.toLocaleString('es'), 14, 1));
console.log('');
console.log('Peticiones al entrar:        ' + ARRANQUE.length);
console.log('Veces que se lee Pedidos:    ' + totalPed +
            '  (' + (PEDIDOS * C_PED.length * totalPed).toLocaleString('es') + ' celdas)');
console.log('Celdas movidas en total:     ' + totalCeldas.toLocaleString('es'));
if (fallos.length) {
  console.log('');
  console.log('No se pudieron medir: ' + fallos.join(' · '));
}
console.log('');
console.log('Cada petición es un viaje entero a Apps Script: abre el libro,');
console.log('lee, calcula y responde. Trece viajes que releen la misma hoja');
console.log('es trece veces el mismo trabajo — eso es lo que se siente lento.');
