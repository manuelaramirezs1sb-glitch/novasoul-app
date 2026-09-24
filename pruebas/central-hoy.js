/**
 * NOVA CENTRAL · HOY, LEYENDO LA RED DE VERDAD.
 *
 * ┌─ LO QUE ELLA DIJO ─────────────────────────────────────────┐
 * │                                                            │
 * │ «¿Cómo que aún no lee la hoja? Ni la hoja ni Nova           │
 * │  Empresarial. Recuerda que Central es la RED que conecta   │
 * │  todo con todos. Así que ponte las pilas y no dejes que    │
 * │  eso se pierda, porque más adelante es duro arreglarlo.»   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Se prueban las dos mitades, y la segunda es la que se salta casi
 * siempre:
 *
 *   SERVIDOR · que los números salgan de las hojas del cliente y
 *              coincidan con los que da Nova Empresarial.
 *
 *   PANTALLA · que se pinten. Escribir el pintor no es terminarlo:
 *              la primera versión de esta pantalla llamaba a `esc()`
 *              y a `traducirError()`, que existen en NovaSoul y NO en
 *              Nova Central, donde la función se llama `escC`. Eso no
 *              lo ve la prueba de sintaxis —es un nombre que falta en
 *              tiempo de ejecución, no un error de escritura— y habría
 *              llegado a su pantalla como una tarjeta en blanco.
 */
const fs = require('fs');
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

// ─── EL SERVIDOR ─────────────────────────────────────────────
const LIBROS = { cen: {}, s: {}, a: {}, emp: {}, emp2: {} };
function libroStub(id) {
  const hojas = LIBROS[id];
  if (!hojas) throw new Error('No existe el libro ' + id);
  return { getSheetByName: (n) => {
    const m = hojas[n];
    if (!m) return null;
    return {
      getName: () => n, getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (k) => m.splice(k - 1, 1),
      getRange: (f, c, nf, nc) => ({
        getValues: () => {
          const o = [];
          for (let i = 0; i < (nf || 1); i++) {
            const r = m[f - 1 + i] || [];
            o.push(r.slice(c - 1, c - 1 + (nc || r.length)));
          }
          return o;
        },
        setValues: (v) => { v.forEach((r, i) => {
          const d = m[f - 1 + i] || (m[f - 1 + i] = []);
          r.forEach((x, j) => { d[c - 1 + j] = x; }); }); },
        setValue: (x) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = x; },
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: libroStub, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{}' }) };
let UUID = 0;
global.Utilities = { sleep: () => {}, getUuid: () => 'u' + (++UUID) + '0000000',
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
(0, eval)(fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8') +
  '\n;globalThis.__S = { centralHoy, centralRed_, apiRecuento, agregarMes,' +
  ' libroOlvidar_, soulOlvidar_, ESQUEMA_EMPRESARIAL };');
const SRV = globalThis.__S;
const E = SRV.ESQUEMA_EMPRESARIAL;
const fila = (cols, o) => cols.map(c => (o[c] !== undefined ? o[c] : ''));

const C_CLI = ['id','empresa','correo','plan','x5','x6','estado','x8','x9','x10',
               'x11','x12','x13','sheet_id'];

/**
 * Una tienda con historia: 40 pedidos de este mes, de los cuales 18
 * entregados, 6 devueltos y el resto todavía en camino.
 *
 * Los números se eligen para que la efectividad NO sea redonda ni
 * coincida con ninguna otra: 18 de 24 cerrados = 75%. Si la pantalla
 * la contara sobre los 40, daría 45% — y eso se vería en la prueba.
 */
function sembrar(opts) {
  opts = opts || {};
  UUID = 0; SRV.libroOlvidar_(); SRV.soulOlvidar_();

  LIBROS.cen = { Clientes: [C_CLI.slice()] };
  if (!opts.sinClientes) {
    LIBROS.cen.Clientes.push(fila(C_CLI, { id:'c1', empresa:'Nutrea', correo:'m@nova.com',
      plan:'pro', estado:'activo', sheet_id:'emp' }));
    if (opts.clienteSinLibro) {
      LIBROS.cen.Clientes.push(fila(C_CLI, { id:'c2', empresa:'Cliente a medias',
        plan:'basico', estado:'activo', sheet_id:'' }));
    }
    if (opts.inactivo) {
      LIBROS.cen.Clientes.push(fila(C_CLI, { id:'c3', empresa:'Ya no está',
        plan:'pro', estado:'inactivo', sheet_id:'emp' }));
    }
  }

  const base = {
    Pedidos: [E.Pedidos.slice()], Novedades: [E.Novedades.slice()],
    Equipo: [E.Equipo.slice()], Tiendas: [['id','nombre','moneda','estado','modalidad']],
    Parametros: [['tienda','clave','valor','nota','tipo']],
    Estados: [['fuente','texto','estado_nova','origen','pedidos','primera_vez',
               'ultima_vez','decidido_por','nota']],
    Movimientos: [['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo']],
    Alarmas: [['id','tienda','tipo','nivel','titulo','accion','estado','creada_en']],
    Inventario: [['id','sku','producto','tienda','stock','minimo','activo']],
    Gastos: [['id','tienda','mes','tipo','nombre','valor','moneda','nota']],
    Cierres: [['tienda','mes','cerrado_en']],
    Llamadas: [E.Llamadas.slice()],
  };
  LIBROS.emp = JSON.parse(JSON.stringify(base));
  LIBROS.emp.Tiendas.push(['gt', 'Nutrea GT', 'GTQ', 'activa', 'marca_propia']);
  if (!opts.unaSolaTienda) {
    LIBROS.emp.Tiendas.push(['ec', 'Nutrea EC', 'USD', 'activa', 'marca_propia']);
  }

  if (!opts.sinPedidos) {
    for (let i = 0; i < 40; i++) {
      const est = i < 18 ? 'entregado' : (i < 24 ? 'devolucion' : 'en_transito');
      LIBROS.emp.Pedidos.push(fila(E.Pedidos, { id: 'p' + i, tienda: 'gt',
        fecha: '2026-09-' + (i % 20 + 1 < 10 ? '0' : '') + (i % 20 + 1),
        cliente: 'Clienta ' + i, valor: 250, estado_canonico: est,
        telefono_norm: '5025000' + i }));
    }
    // Nueve novedades del mes.
    for (let i = 0; i < 9; i++) {
      LIBROS.emp.Novedades.push(fila(E.Novedades, { id: 'n' + i, pedido_id: 'p' + i,
        fecha: '2026-09-1' + (i % 9), motivo: 'No contesta',
        estado: i < 4 ? 'abierta' : 'resuelta' }));
    }
  }

  if (!opts.sinEquipo) {
    LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e1', nombre:'Manuela Ramírez',
      correo:'m@nova.com', rol:'dueno', tienda:'*', estado:'activo',
      ultima_conexion:'2026-09-24T09:00:00Z' }));
    LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e2', nombre:'Paola Gómez',
      correo:'p@nutrea.com', rol:'gestora', tienda:'gt', estado:'activo',
      ultima_conexion:'' }));
    // Alguien de OTRA tienda: no puede salir en la tarjeta de GT.
    LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e3', nombre:'Lucía Torres',
      correo:'l@nutrea.com', rol:'gestora', tienda:'ec', estado:'activo' }));
    // Y alguien inactivo: tampoco.
    LIBROS.emp.Equipo.push(fila(E.Equipo, { id:'e4', nombre:'Quien se fue',
      correo:'x@nutrea.com', rol:'gestora', tienda:'gt', estado:'inactivo' }));
  }
}

const S = { correo: 'm@nova.com', nombre: 'Manuela', rol: 'socia', id: 'op1' };

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

// ══════════════════════════════════════════════════════════════
console.log('\n── Central abre el libro del cliente y cuenta ──');
sembrar();
let r = SRV.centralHoy(S, {});
ok('responde', r.ok);
igual('encuentra a su cliente', 'Nutrea', r.cliente.empresa);
igual('y sus dos tiendas', ['Nutrea GT', 'Nutrea EC'], r.tiendas.map(t => t.nombre));
igual('mira una sola a la vez', 'gt', r.tienda);
igual('40 pedidos del mes', 40, r.kpis.pedidos);
igual('en la moneda de esa tienda', 'GTQ', r.kpis.moneda);
/** 18 entregados × 250 = 4.500. Solo lo entregado cuenta como venta. */
igual('las ventas son solo lo entregado', 4500, r.kpis.ventas);

console.log('\n── La efectividad se cuenta sobre lo que YA terminó ──');
/**
 * 18 entregados de 24 cerrados = 75%. Sobre los 40 daría 45%, que
 * castigaría a la tienda por tener pedidos en camino — que a mitad de
 * mes es exactamente lo normal.
 */
igual('75%, no 45%', 75, r.kpis.efectividad);
igual('y dice sobre cuántos', 24, r.kpis.efectividadSobre);
igual('los que siguen en camino se cuentan aparte', 16, r.kpis.enCamino);

console.log('\n── Sin pedidos terminados NO dice 0% ──');
sembrar({ sinPedidos: true });
r = SRV.centralHoy(S, {});
igual('la efectividad es «no se sabe», no cero', null, r.kpis.efectividad);
ok('y lo explica', r.avisos.join(' ').indexOf('No es 0%') !== -1, r.avisos.join(' | '));
igual('la tienda se marca como vacía', true, r.kpis.vacio);

console.log('\n── Los mismos números que da Nova Empresarial ──');
/**
 * Central NO recalcula: llama a `agregarMes`, la misma función. Si
 * algún día alguien escribe una segunda cuenta, las dos pantallas
 * podrían decirle cifras distintas de la misma tienda el mismo día.
 */
sembrar();
const sesionEmp = { sheetId: 'emp', rol: 'dueno', email: 'm@nova.com',
                    tiendas: ['gt'], permisos: [], modulos: ['empresarial'] };
SRV.libroOlvidar_();
const desdeEmp = SRV.agregarMes(libroStub('emp'), 'gt', '2026-09', sesionEmp, null);
SRV.libroOlvidar_();
const desdeCen = SRV.centralHoy(S, {});
igual('pedidos coinciden', desdeEmp.pedidos, desdeCen.kpis.pedidos);
igual('ventas coinciden', desdeEmp.ventas, desdeCen.kpis.ventas);
igual('novedades coinciden', desdeEmp.novedades, desdeCen.kpis.novedades);

console.log('\n── El equipo es el de verdad, y solo el de esa tienda ──');
sembrar();
r = SRV.centralHoy(S, {});
igual('dos personas en GT', ['Manuela Ramírez', 'Paola Gómez'],
      r.equipo.map(x => x.nombre));
ok('la de la otra tienda NO aparece',
   r.equipo.filter(x => x.nombre === 'Lucía Torres').length === 0);
ok('ni la inactiva', r.equipo.filter(x => x.nombre === 'Quien se fue').length === 0);
igual('quien entró trae su fecha', '2026-09-24',
      r.equipo.filter(x => x.nombre === 'Manuela Ramírez')[0].ultimaConexion);
igual('y quien no ha entrado, vacío — no una fecha inventada', '',
      r.equipo.filter(x => x.nombre === 'Paola Gómez')[0].ultimaConexion);

console.log('\n── Un equipo vacío se dice, no se rellena ──');
sembrar({ sinEquipo: true });
r = SRV.centralHoy(S, {});
igual('no hay nadie', 0, r.equipo.length);
ok('y lo explica', /no hay nadie asignado/i.test(r.avisos.join(' ')), r.avisos.join(' | '));

console.log('\n── La red: cuántos clientes y cuántos conectados ──');
sembrar({ clienteSinLibro: true, inactivo: true });
r = SRV.centralHoy(S, {});
igual('dos clientes activos', 2, r.red.clientes);
igual('uno con libro', 1, r.red.conLibro);
igual('y uno a medias', 1, r.red.sinLibro);
ok('el inactivo no cuenta', r.clientes.filter(c => c.empresa === 'Ya no está').length === 0);
igual('elige al que sí tiene libro', 'Nutrea', r.cliente.empresa);

console.log('\n── Sin clientes, no finge ──');
sembrar({ sinClientes: true });
r = SRV.centralHoy(S, {});
igual('responde igual', true, r.ok);
igual('sin cliente', null, r.cliente);
igual('sin números', null, r.kpis);
ok('y dice qué falta', /primer cliente/i.test(r.avisos.join(' ')), r.avisos.join(' | '));

// ══════════════════════════════════════════════════════════════
(async () => {
  console.log('\n── Y AHORA LA PANTALLA, que es la mitad que se olvida ──');
  sembrar();
  const RESP = SRV.centralHoy(S, {});
  const VACIO = (function () { sembrar({ sinClientes: true }); return SRV.centralHoy(S, {}); })();

  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  for (const ancho of [1280, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1100 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    p.on('console', m => {
      if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) {
        errores.push('console: ' + m.text());
      }
    });

    await p.goto('file:///home/claude/repo/novacentral.html');
    await p.waitForLoadState('load');
    await p.evaluate((R) => {
      const puerta = document.getElementById('login-screen') || document.getElementById('login');
      if (puerta) puerta.style.display = 'none';
      const app = document.getElementById('app');
      if (app) app.style.display = 'flex';
      window.nc = async (accion) => (accion === 'nc_hoy'
        ? JSON.parse(JSON.stringify(R)) : { ok: true });
      return cargarRed('', '');
    }, RESP);
    await p.waitForTimeout(300);

    const kpis = (await p.textContent('#dash-kpis')).replace(/\s+/g, ' ');
    const abajo = (await p.textContent('#dash-bottom')).replace(/\s+/g, ' ');

    /**
     * «GTQ 4500» y no «GTQ 4.500»: en español los números de cuatro
     * cifras NO llevan separador de miles (a partir de cinco, sí). Mi
     * primera versión de esta prueba esperaba el punto y falló — el
     * error era mío, no de la pantalla.
     */
    ok('pinta las ventas en su moneda' + A, /GTQ 4500/.test(kpis), kpis);
    ok('los 40 pedidos' + A, /40/.test(kpis));
    ok('la efectividad sobre los cerrados' + A,
       /75%/.test(kpis) && /24 ya cerrados/.test(kpis), kpis);
    ok('y las novedades' + A, /Novedades del mes/.test(kpis));

    ok('el equipo real' + A, /Manuela Ram/.test(abajo) && /Paola G/.test(abajo), abajo);
    ok('dice quién nunca ha entrado' + A, /Nunca ha entrado/.test(abajo));

    console.log('\n── Y nada inventado' + A + ' ──');
    const INVENTADOS = ['Daniela R.', 'Camila R.', 'Katherin P.', '284.500', '342', '88%'];
    const quedan = INVENTADOS.filter(x => (kpis + abajo).indexOf(x) !== -1);
    ok('ninguna cifra ni persona de la maqueta' + A, quedan.length === 0,
       'siguen ahí: ' + quedan.join(', '));

    console.log('\n── Sin clientes, la pantalla lo dice' + A + ' ──');
    await p.evaluate((V) => {
      window.nc = async () => JSON.parse(JSON.stringify(V));
      return cargarRed('', '');
    }, VACIO);
    await p.waitForTimeout(250);
    const vacio = (await p.textContent('#dash-kpis')).replace(/\s+/g, ' ');
    ok('explica qué falta en vez de mostrar ceros' + A,
       /primer cliente/i.test(vacio) && /no hay clientes/i.test(vacio), vacio);

    ok('sin errores de JavaScript' + A, errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
  process.exit(fallas ? 1 : 0);
})();
