/**
 * El semáforo semanal, sobre hojas falsas.
 *
 * Lo que se comprueba no es que el código corra: es que las reglas de
 * cálculo sean LAS QUE SON. Solo los entregados generan ganancia, la
 * devolución cuesta el flete, el cancelado cuesta cero, y donde no hay
 * dato dice sin dato. Si alguna de esas cambia sin que alguien lo
 * decida, esto falla.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const HOJAS = {};
function hoja(nombre) {
  const m = HOJAS[nombre];
  if (!m) return null;
  return {
    getLastRow: () => m.length,
    getLastColumn: () => (m[0] ? m[0].length : 0),
    getDataRange: () => ({ getValues: () => m }),
    getRange: (f, c, nf, nc) => ({
      getValues: () => {
        const out = [];
        for (let i = 0; i < (nf || 1); i++) {
          const fila = m[f - 1 + i] || [];
          out.push(fila.slice(c - 1, c - 1 + (nc || fila.length)));
        }
        return out;
      },
      setValues: () => {}, setValue: () => {},
    }),
  };
}
const SS = { getSheetByName: hoja };

const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const CORREOS = [];
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: (k, v) => { PROPS[k] = v; }, deleteProperty: (k) => { delete PROPS[k]; } }) };
global.SpreadsheetApp = { openById: () => SS, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: (o) => CORREOS.push(o) };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{"data":[]}' }) };
global.Utilities = { sleep: () => {}, getUuid: () => 'u',
  formatDate: (d, tz, pat) => {
    const iso = new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(pat)) return iso.slice(0, 19);
    return iso;
  } };

(0, eval)(src + '\n;globalThis.__F = { libroOlvidar_, semaforoSemanal, semaforoTexto, lunesDe_, semanaCerrada_,' +
  ' masDias_, puedeVerSemaforo, numerosSemana_, TRABAJOS };');
const F = globalThis.__F;

let fallas = 0;
function ok(n, c, d) { if (c) console.log('  ok     ' + n);
                       else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } }
function igual(n, esperado, real) {
  ok(n, JSON.stringify(esperado) === JSON.stringify(real),
     'esperaba ' + JSON.stringify(esperado) + ', obtuve ' + JSON.stringify(real));
}

const PED = ['id','fuente','id_externo','fecha','tienda','cliente','cedula','correo',
  'telefono','telefono_norm','telefono_2','telefono_2_norm','ciudad','departamento',
  'direccion','producto','sku','cantidad','valor','costo_producto','costo_envio',
  'metodo_pago','bodega','estado','estado_transportadora','estado_canonico',
  'transportadora','guia','intentos','gestora_asignada','fecha_promesa','fecha_entrega',
  'razon_cancelacion','estado_nova','nota','ultimo_movimiento','adelanto',
  'acuerdo_oficina','confirmado_oficina','actualizado_en','actualizado_por'];
const PAU = ['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
  'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado','impresiones','alcance',
  'frecuencia','clics','ctr','cpc','cpm','resultados','compras','cpa','roas','valor_conv','visitas_lp'];

function pedido(fecha, estado, valor, costoProd, flete, producto) {
  const f = new Array(PED.length).fill('');
  f[PED.indexOf('fecha')] = fecha;
  f[PED.indexOf('tienda')] = 'gt';
  f[PED.indexOf('producto')] = producto || 'TAG RECEDE';
  f[PED.indexOf('valor')] = valor;
  f[PED.indexOf('costo_producto')] = costoProd;
  f[PED.indexOf('costo_envio')] = flete;
  f[PED.indexOf('estado_nova')] = estado;
  return f;
}
function gasto(fecha, monto, moneda) {
  const f = new Array(PAU.length).fill('');
  f[1] = fecha; f[2] = fecha; f[3] = 'gt'; f[4] = 'meta';
  f[7] = 'Frío'; f[10] = monto; f[11] = moneda || 'GTQ';
  return f;
}

function montar(pedidos, pauta, tasas, params) {
  Object.keys(HOJAS).forEach(k => delete HOJAS[k]);
  /**
   * El mundo cambió: hay que decírselo al libro.
   *
   * `libro_()` guarda el manejador Y lo leído mientras dura la ejecución
   * —es lo que bajó el arranque de 75 lecturas de pestaña a 8—. Esta
   * función reemplaza las pestañas enteras por debajo, que es algo que
   * en producción no pasa nunca: ahí solo se escribe por la API, y
   * escribir ya tira lo guardado solo.
   *
   * Sin esta línea la prueba seguiría leyendo el mundo anterior.
   */
  F.libroOlvidar_();

  HOJAS.Parametros = [['tienda','clave','valor','actualizado_en','actualizado_por'],
                      ['', 'moneda_reporte', 'GTQ', '', '']]
    .concat((params || []).map(p => ['gt', p[0], p[1], '', '']));
  HOJAS.Tiendas = [['id','nombre','marca','pais','sociedad','nit','moneda',
                    'zona_horaria','corte_despacho','modalidad','estado'],
                   ['gt','Nutrea GT','','GT','','','GTQ','UTC','','contraentrega','activa']];
  HOJAS.Pedidos = [PED].concat(pedidos || []);
  HOJAS.Pauta = [PAU].concat(pauta || []);
  if (tasas) HOJAS.Tasas = [['fecha','moneda_origen','moneda_destino','tasa']].concat(tasas);
  HOJAS.Equipo = [['id','nombre','correo','rol','tienda','estado','permisos'],
                  ['e1','Manuela','due@x.com','dueño','*','activo','']];
  HOJAS.Clientes = [['id','empresa','pais','plan','tarifa','costo','estado','fecha_alta',
                     'fecha_corte','ultimo_pago','nota','x','tiendas','sheet_id'],
                    ['c1','Nutrea','GT','Completo','','','activo','','','','','',1,'hoja-1']];
  HOJAS.Planes = [['id','nombre','modulos','limite_usuarios','limite_tiendas',
                   'costo_calculado','precio_sugerido','tarifa_fijada'],
                  ['p1','Completo','empresarial,pauta,dinero',10,5,'','','']];
}

// Una semana cerrada conocida, lejos de hoy para que nada dependa del día.
const L = '2026-08-17';                       // lunes
const d = (n) => F.masDias_(L, n);

console.log('\nLAS SEMANAS');
igual('el domingo pertenece a la semana que cierra, no a la que abre',
      '2026-08-17', F.lunesDe_('2026-08-23'));
igual('y el lunes abre la suya', '2026-08-24', F.lunesDe_('2026-08-24'));

console.log('\nLAS REGLAS DE CÁLCULO');
// 3 entregados (valor 300, costo 100, flete 30) → ganancia 3 × 170 = 510
// 1 devolución → cuesta el flete: −30
// 2 cancelados → cuestan 0
montar([
  pedido(d(0), 'entregado', 300, 100, 30),
  pedido(d(1), 'entregado', 300, 100, 30),
  pedido(d(2), 'entregado', 300, 100, 30),
  pedido(d(3), 'devolucion', 300, 100, 30),
  pedido(d(4), 'cancelado', 300, 100, 30),
  pedido(d(4), 'cancelado', 300, 100, 30),
  pedido(d(5), 'en_transito', 300, 100, 30),
], [gasto(d(0), 100), gasto(d(1), 100)]);

let s = F.semaforoSemanal('hoja-1', 'gt', L);
igual('pedidos creados', 7, s.hoy.pedidos);
igual('solo los entregados generan ganancia', 510, s.hoy.ganancia);
igual('la devolución cuesta el flete, no el producto', 30, s.hoy.costoDevoluciones);
igual('los cancelados cuestan cero', 2, s.hoy.cancelados);
igual('utilidad antes de pauta = ganancia − devoluciones', 480, s.hoy.utilidadAntesPauta);
igual('la pauta se suma', 200, s.hoy.gasto);
igual('utilidad real = antes de pauta − pauta', 280, s.hoy.utilidadReal);
igual('el techo es utilidad antes de pauta ÷ pedidos creados',
      Math.round(480 / 7 * 100) / 100, Math.round(s.hoy.techo * 100) / 100);
igual('la tasa de entrega es entregados ÷ (entregados + devoluciones)', 75, s.hoy.tasaEntrega);
igual('lo que sigue en tránsito queda fuera de esa cuenta', 1, s.hoy.enTransito);
igual('ticket = ventas entregadas ÷ entregados', 300, s.hoy.ticket);

console.log('\nDONDE NO HAY DATO, DICE SIN DATO');
montar([pedido(d(0), 'entregado', 300, 100, 30)], []);   // sin pauta
s = F.semaforoSemanal('hoja-1', 'gt', L);
ok('sin pauta, la utilidad real es null y no la de antes de pauta', s.hoy.utilidadReal === null);
ok('sin pauta, el CPA es null y no cero', s.hoy.cpaPorEntrega === null);
ok('y la luz de utilidad queda incompleta, no verde',
   s.luces[3].estado === 'incompleto', s.luces[3].estado);
ok('el texto lo dice', F.semaforoTexto(s).includes('sin dato'));

console.log('\nUN ENTREGADO SIN COSTO NO INVENTA GANANCIA');
montar([pedido(d(0), 'entregado', 300, 0, 30), pedido(d(1), 'entregado', 300, 100, 30)], []);
s = F.semaforoSemanal('hoja-1', 'gt', L);
igual('solo cuenta el que tiene costo', 170, s.hoy.ganancia);
igual('y dice cuántos quedaron fuera', 1, s.hoy.sinCosto);
ok('con una alerta que explica dónde se arregla',
   s.alertas.some(a => a.titulo.includes('costo de producto')));

console.log('\nLA SEMANA VACÍA NO ES UNA SEMANA MALA');
montar([], []);
s = F.semaforoSemanal('hoja-1', 'gt', L);
ok('lo dice sin adornos', s.hayDatos === false);
ok('y el correo explica la diferencia',
   F.semaforoTexto(s).includes('no se actualizó'));
ok('la alerta señala las semanas vacías',
   s.alertas[0].titulo.includes('sin un solo pedido'), s.alertas[0].titulo);

console.log('\nMUESTRA CHICA: SE DICE CUÁNTO FALTA, NO SE OPINA');
montar([pedido(d(0), 'entregado', 300, 100, 30),
        pedido(d(1), 'devolucion', 300, 100, 30)], []);
s = F.semaforoSemanal('hoja-1', 'gt', L);
ok('la luz de entrega no juzga con 2 resueltos',
   s.luces[2].estado === 'muestra_chica', s.luces[2].estado);
ok('y el producto dice cuántos faltan',
   /faltan \d+ resueltos/.test(s.productos[0].señal), s.productos[0].señal);
ok('no hay alerta de entrega baja con esa muestra',
   !s.alertas.some(a => a.titulo.includes('bajo tu mínimo')));

console.log('\nCON MUESTRA SUFICIENTE SÍ SE OPINA');
const muchos = [];
for (let i = 0; i < 6; i++) muchos.push(pedido(d(i % 7), 'entregado', 300, 100, 30));
for (let i = 0; i < 6; i++) muchos.push(pedido(d(i % 7), 'devolucion', 300, 100, 30));
montar(muchos, []);
s = F.semaforoSemanal('hoja-1', 'gt', L);
igual('entrega al 50%', 50, s.hoy.tasaEntrega);
ok('ahora sí es rojo', s.luces[2].estado === 'rojo', s.luces[2].estado);
ok('y sale la alerta con su acción',
   s.alertas.some(a => a.titulo.includes('bajo tu mínimo')));

console.log('\nPAUTA EN OTRA MONEDA SIN TASA: NO SE SUMA Y SE DICE');
montar([pedido(d(0), 'entregado', 300, 100, 30)],
       [gasto(d(0), 400000, 'COP')]);        // tienda en GTQ, sin hoja Tasas
s = F.semaforoSemanal('hoja-1', 'gt', L);
igual('no entra al gasto', 0, s.hoy.gasto);
igual('queda contado aparte', 400000, s.hoy.gastoSinConvertir);
ok('y hay una alerta roja que manda a arreglarlo',
   s.alertas.some(a => a.titulo.includes('sin convertir') && a.nivel === 'rojo'));

console.log('\nCON LA TASA, SÍ SE SUMA');
montar([pedido(d(0), 'entregado', 300, 100, 30)],
       [gasto(d(0), 400000, 'COP')],
       [[d(0), 'COP', 'GTQ', 0.002]]);
s = F.semaforoSemanal('hoja-1', 'gt', L);
igual('convertido con la tasa de ese día', 800, s.hoy.gasto);

console.log('\nQUIÉN PUEDE VERLO');
ok('sin el módulo de dinero en el plan, no',
   !F.puedeVerSemaforo({ modulos: ['empresarial','pauta'], permisos: [], rol: 'dueno' }).puede);
ok('y el mensaje manda a Nova Central, que es donde se resuelve',
   F.puedeVerSemaforo({ modulos: ['empresarial'], permisos: [], rol: 'dueno' })
    .error.includes('Nova Central'));
ok('con el plan completo, la dueña sí',
   F.puedeVerSemaforo({ modulos: ['empresarial','pauta','dinero'], permisos: [], rol: 'dueno' }).puede);
ok('una admin sin ver_dinero, no',
   !F.puedeVerSemaforo({ modulos: ['empresarial','pauta','dinero'], permisos: ['subir_pedidos'],
                         rol: 'admin' }).puede);
ok('la misma admin con ver_dinero, sí',
   F.puedeVerSemaforo({ modulos: ['empresarial','pauta','dinero'],
                        permisos: ['subir_pedidos','ver_dinero'], rol: 'admin' }).puede);
ok('el plan manda sobre el permiso: sin módulo no basta el permiso',
   !F.puedeVerSemaforo({ modulos: ['empresarial'], permisos: ['ver_dinero'], rol: 'admin' }).puede);

console.log('\nEL DISPARADOR DEL LUNES');
const sem = F.TRABAJOS.filter(t => t.fn === 'semaforoLunes')[0];
ok('existe y es semanal', !!sem && sem.dia === 'MONDAY');
/**
 * El semáforo tiene que correr después de los TRES que lo alimentan, y
 * se nombran uno por uno a propósito.
 *
 * Antes esto decía «después de todos los diarios», que es una regla más
 * ancha de lo que hace falta: el primer trabajo diario que no tuviera
 * nada que ver con la tienda —el aviso de cobros de ella, por ejemplo—
 * rompía la prueba sin que hubiera nada roto. Lo que importa es que el
 * gasto de pauta esté adentro y convertido cuando el semáforo juzgue.
 */
const ALIMENTAN = ['actualizarTasasDiario', 'leerMetaDiario', 'revisarAlarmasTodos'];
const previos = F.TRABAJOS.filter(t => ALIMENTAN.indexOf(t.fn) !== -1);
ok('los tres que lo alimentan siguen existiendo y son diarios',
   previos.length === 3 && previos.every(t => !t.dia),
   previos.map(t => t.fn).join(','));
ok('corre después de las tasas, Meta y las alarmas',
   sem.hora > Math.max.apply(null, previos.map(t => t.hora)),
   'semáforo ' + (sem && sem.hora) + ' vs ' +
   previos.map(t => t.fn + '@' + t.hora).join(', '));

console.log(fallas ? '\n' + fallas + ' FALLA(S)\n' : '\nTodo pasa.\n');
process.exit(fallas ? 1 : 0);
