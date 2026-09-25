/**
 * EL PUNTO DE EQUILIBRIO · una sola cuenta, y que no se contradiga.
 *
 * ┌─ LO QUE ESTA PRUEBA EXISTE PARA QUE NO VUELVA A PASAR ─────┐
 * │                                                            │
 * │ «el punto de equilibrio lleva días saliendo en cero,        │
 * │  revísalo en la parte de dinero, y en la parte del hoy      │
 * │  está disque en 31. Las páginas no están dando los          │
 * │  números y no coincide lo que dice: una parte dice algo y   │
 * │  en otra dice algo diferente».                             │
 * │                                                            │
 * │ Tenía tres problemas encima del mismo mes:                 │
 * │                                                            │
 * │  · Hoy decía «31 entregas · 47 entregadas · $ 1.842.000     │
 * │    de gastos fijos». Nada de eso salía de su hoja: era la   │
 * │    maqueta, en pesos colombianos, sobre una tienda que      │
 * │    factura en dólares. `aplicarReales` nunca tocaba esa     │
 * │    tarjeta.                                                │
 * │  · Dinero decía «te faltan 0 · 100% del punto de            │
 * │    equilibrio» tres centímetros debajo de «QUEDA LIMPIO     │
 * │    $ −109». Las dos cifras salían de la misma hoja y se     │
 * │    contradecían.                                           │
 * │  · Y las dos pantallas calculaban por su cuenta, con        │
 * │    fórmulas distintas.                                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ EL INVARIANTE ────────────────────────────────────────────┐
 * │                                                            │
 * │        faltan > 0    ⟺    utilidad < 0                     │
 * │                                                            │
 * │ No es un detalle de presentación. Es la única propiedad     │
 * │ que hace IMPOSIBLE que «ya cubriste tus gastos» y «estás    │
 * │ perdiendo plata» aparezcan juntas otra vez. Se afirma       │
 * │ abajo sobre cientos de combinaciones al azar, porque un     │
 * │ caso escogido a mano solo prueba el caso que escogí.        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libro(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (nombre) => {
    const m = hojas[nombre];
    if (!m) return null;
    return {
      getName: () => nombre,
      getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (n) => m.splice(n - 1, 1),
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
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-25T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: (id) => libro(id), flush: () => {} };
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

(0, eval)(src + '\n;globalThis.__F = { agregarMes, libroOlvidar_, soulOlvidar_,' +
  ' ESQUEMA_EMPRESARIAL };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_PED = F.ESQUEMA_EMPRESARIAL.Pedidos;
const C_TIE = F.ESQUEMA_EMPRESARIAL.Tiendas;
const C_GAS = F.ESQUEMA_EMPRESARIAL.Gastos;
const C_PAU = F.ESQUEMA_EMPRESARIAL.Pauta;
const C_INV = F.ESQUEMA_EMPRESARIAL.Inventario;
const C_CAR = F.ESQUEMA_EMPRESARIAL.Cartera;
function f(cols, o) { return cols.map(c => (o[c] !== undefined ? o[c] : '')); }

const S = { email: 'm@nova.com', nombre: 'Manuela', rol: 'dueno',
            sheetId: 'emp', tiendas: ['ec'] };

/**
 * Un mes armado a mano: N entregas de `valor` cada una, con su costo de
 * producto y su flete, más la pauta y los gastos fijos que se le pasen.
 */
function mes(o) {
  F.libroOlvidar_(); F.soulOlvidar_();
  const peds = [C_PED];
  // Las devoluciones son las que activan el cobro de retorno, que es el
  // costo que la contribución NO estaba restando.
  for (let i = 0; i < (o.devueltos || 0); i++) {
    peds.push(f(C_PED, { id: 'v' + i, tienda: 'ec', fecha: '2026-09-10',
      producto: 'TAG', cantidad: 1, valor: o.valor,
      costo_producto: o.costoProd, costo_envio: o.flete,
      estado: 'DEVOLUCION', estado_canonico: 'devolucion', estado_nova: 'devolucion' }));
  }
  for (let i = 0; i < o.entregados; i++) {
    peds.push(f(C_PED, { id: 'p' + i, tienda: 'ec', fecha: '2026-09-10',
      producto: 'TAG', cantidad: 1, valor: o.valor,
      costo_producto: o.costoProd, costo_envio: o.flete,
      estado: 'ENTREGADO', estado_canonico: 'entregado', estado_nova: 'entregado',
      fecha_entrega: '2026-09-12' }));
  }
  LIBROS.emp = {
    Tiendas: [C_TIE, f(C_TIE, { id: 'ec', nombre: 'Nutrea Ecuador', moneda: 'USD',
                                zona_horaria: 'UTC' })],
    Pedidos: peds,
    Inventario: [C_INV],
    Pauta: [C_PAU, f(C_PAU, { id: 'a1', fecha: '2026-09-01', tienda: 'ec',
      plataforma: 'Meta', campana: 'C1', gasto: o.pauta, moneda_gasto: 'USD',
      gasto_normalizado: o.pauta, resultados: 10 })],
    Gastos: [C_GAS].concat((o.fijos || []).map(function (g, i) {
      return f(C_GAS, { id: 'g' + i, tienda: 'ec', nombre: g.nombre, valor: g.valor,
                        tipo: 'fijo', activo: 'si' });
    })),
    Cartera: [C_CAR].concat((o.cobrosRetorno || []).map(function (m, i) {
      return f(C_CAR, { id: 'c' + i, tienda: 'ec', fecha: '2026-09-15',
                        clase: 'devolucion', monto: -m });
    })),
  };
  return F.agregarMes(libro('emp'), 'ec', '2026-09', S);
}

console.log('\n── 1 · el mes que pierde plata NO dice «equilibrio cubierto» ──');
/**
 * Esta es la forma exacta de lo que ella fotografió: una tienda con la
 * utilidad en rojo y el punto de equilibrio al 100%. Con las dos cuentas
 * separadas eso era posible; ahora no.
 */
let d = mes({ entregados: 40, valor: 30, costoProd: 8, flete: 7, pauta: 550,
              fijos: [{ nombre: 'Shopify', valor: 35 },
                      { nombre: 'Claude', valor: 20 }] });
igual('40 entregas', 40, d.entregados);
igual('55 de gastos fijos', 55, d.fijos);
ok('la utilidad es negativa', d.utilidad < 0, String(d.utilidad));
ok('y por tanto FALTAN entregas, no cero', d.equilibrio.faltan > 0,
   JSON.stringify(d.equilibrio));
ok('el avance no puede ser 100%', d.equilibrio.avance < 100, String(d.equilibrio.avance));

console.log('\n── 1b · las comisiones entran en lo que hay que cubrir ──');
/**
 * Este era el segundo de los dos olvidos: la comisión internacional del
 * banco sobre la pauta se restaba en la utilidad pero NO se sumaba a lo
 * que el punto de equilibrio tenía que cubrir. Sola, esa diferencia ya
 * bastaba para que un mes justo apareciera como cubierto.
 */
ok('las comisiones son parte del cálculo', d.comisiones > 0, String(d.comisiones));
ok('y aparecen nombradas en el desglose',
   d.equilibrio.detalle.some(x => /omisi/.test(x.nombre)),
   JSON.stringify(d.equilibrio.detalle.map(x => x.nombre)));

console.log('\n── 2 · lo que hay que cubrir se nombra entero ──');
const nombres = d.equilibrio.detalle.map(x => x.nombre);
ok('está la pauta', nombres.indexOf('Pauta del mes') !== -1, JSON.stringify(nombres));
ok('están los fijos uno por uno',
   nombres.indexOf('Shopify') !== -1 && nombres.indexOf('Claude') !== -1,
   JSON.stringify(nombres));
const suma = d.equilibrio.detalle.reduce((a, x) => a + x.valor, 0);
ok('y el desglose suma exactamente lo que hay que cubrir',
   Math.abs(suma - d.equilibrio.aCubrir) < 0.001,
   suma + ' vs ' + d.equilibrio.aCubrir);

console.log('\n── 3 · con el mes cubierto, faltan cero ──');
d = mes({ entregados: 100, valor: 40, costoProd: 5, flete: 5, pauta: 500,
          fijos: [{ nombre: 'Shopify', valor: 100 }] });
ok('la utilidad es positiva', d.utilidad > 0, String(d.utilidad));
igual('no falta ninguna entrega', 0, d.equilibrio.faltan);
igual('y el avance es 100%', 100, d.equilibrio.avance);

console.log('\n── 4 · sin entregas no se inventa un número ──');
d = mes({ entregados: 0, valor: 0, costoProd: 0, flete: 0, pauta: 300, fijos: [] });
igual('dice que no hay con qué calcular', false, d.equilibrio.hay);
// null y no 0: un cero en «faltan» se lee como «ya está cubierto», que
// es lo contrario de «todavía no sé».
igual('y no finge una meta', null, d.equilibrio.necesarias);
igual('ni un «faltan 0»', null, d.equilibrio.faltan);

console.log('\n── 5 · EL INVARIANTE, sobre 400 meses al azar ──');
/**
 * Un caso escogido a mano solo prueba el caso que escogí. Estos salen de
 * un generador con semilla fija —así el fallo se puede repetir— y cubren
 * meses que ganan, meses que pierden y meses en el filo.
 */
let semilla = 20260925;
const azar = function () {
  semilla = (semilla * 1103515245 + 12345) % 2147483648;
  return semilla / 2147483648;
};
let rotos = [], probados = 0, ganan = 0, pierden = 0;
for (let i = 0; i < 400; i++) {
  const entregados = 1 + Math.floor(azar() * 120);
  const valor = 10 + azar() * 60;
  const costoProd = azar() * valor * 0.4;
  const flete = azar() * valor * 0.5;
  const pauta = azar() * entregados * valor * 0.6;
  const fijos = [{ nombre: 'Fijo', valor: azar() * 400 }];
  const r = mes({ entregados, valor, costoProd, flete, pauta, fijos });
  probados++;
  if (r.utilidad > 0) ganan++; else if (r.utilidad < 0) pierden++;
  const faltan = r.equilibrio.faltan;
  // faltan > 0  ⟺  utilidad < 0
  const coherente = (faltan > 0) === (r.utilidad < -0.000001) ||
                    Math.abs(r.utilidad) < 0.000001;
  if (!coherente) {
    rotos.push({ entregados, utilidad: r.utilidad, faltan: faltan,
                 necesarias: r.equilibrio.necesarias });
  }
}
igual('400 meses probados', 400, probados);
ok('hay meses que ganan y meses que pierden en la muestra',
   ganan > 40 && pierden > 40, ganan + ' ganan, ' + pierden + ' pierden');
ok('en TODOS, faltan>0 ⟺ utilidad<0', rotos.length === 0,
   rotos.length + ' rotos, p.ej. ' + JSON.stringify(rotos[0]));

console.log('\n── 6 · la contribución es la que dice ser ──');
d = mes({ entregados: 50, valor: 30, costoProd: 8, flete: 7, pauta: 400,
          fijos: [{ nombre: 'Fijo', valor: 100 }] });
// 30 − 8 − 7 = 15 por entrega. 500 a cubrir ÷ 15 = 34 entregas.
ok('cada entrega deja valor − producto − flete',
   Math.abs(d.equilibrio.contribucion - 15) < 0.001, String(d.equilibrio.contribucion));
igual('necesarias = techo(a cubrir / contribución)', 34, d.equilibrio.necesarias);
igual('y con 50 hechas no falta ninguna', 0, d.equilibrio.faltan);
ok('la contribución × entregas − a cubrir es la utilidad',
   Math.abs(d.equilibrio.contribucion * d.entregados - d.equilibrio.aCubrir - d.utilidad) < 0.001,
   String(d.utilidad));

console.log('\n── 7 · el cobro de retorno, que era el primer olvido ──');
/**
 * La contribución calculaba «valor − producto − flete» y se saltaba lo
 * que la plataforma cobra por devolver el paquete. En agosto de Nutrea
 * eso eran 95 dólares que ningún cierre contaba. Restarlo en la utilidad
 * y no en la contribución es justo lo que producía «100% del punto de
 * equilibrio» encima de un «queda limpio» negativo.
 */
const sinRetorno = mes({ entregados: 40, valor: 30, costoProd: 8, flete: 7,
  pauta: 300, fijos: [{ nombre: 'Fijo', valor: 50 }] });
const conRetorno = mes({ entregados: 40, valor: 30, costoProd: 8, flete: 7,
  pauta: 300, fijos: [{ nombre: 'Fijo', valor: 50 }],
  devueltos: 20, cobrosRetorno: [100, 100] });

igual('la cartera manda sobre el flete del export', 'cartera',
      conRetorno.costoDevolucionFuente);
ok('con devoluciones cobradas, cada entrega deja MENOS',
   conRetorno.equilibrio.contribucion < sinRetorno.equilibrio.contribucion,
   conRetorno.equilibrio.contribucion + ' vs ' + sinRetorno.equilibrio.contribucion);
/**
 * Y aquí la prueba encontró un error de verdad: con un cobro de retorno
 * grande la contribución se vuelve NEGATIVA, y la primera versión
 * devolvía `necesarias = 0` y por tanto `faltan = 0`. La pantalla decía
 * «equilibrio cubierto» sobre una tienda que perdía 1.896 dólares.
 *
 * Cuando cada entrega pierde plata no existe un número de entregas que
 * llegue al equilibrio: vender más lo aleja. Eso hay que decirlo, no
 * devolver un cero que se lee al revés.
 */
ok('la contribución se vuelve negativa', conRetorno.equilibrio.contribucion < 0,
   String(conRetorno.equilibrio.contribucion));
igual('y se dice que el equilibrio NO es alcanzable', false,
      conRetorno.equilibrio.alcanzable);
igual('sin fingir un «faltan 0»', null, conRetorno.equilibrio.faltan);
igual('ni una meta que no existe', null, conRetorno.equilibrio.necesarias);
igual('el avance no es 100%', 0, conRetorno.equilibrio.avance);
ok('y sigue habiendo entregas con las que hablar', conRetorno.equilibrio.hay === true);
ok('la utilidad, por supuesto, es negativa', conRetorno.utilidad < 0,
   String(conRetorno.utilidad));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
