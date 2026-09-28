/**
 * REVISAR ANTES DE CERRAR.
 *
 * ┌─ LO QUE PIDIÓ, TRES VECES ─────────────────────────────────┐
 * │                                                            │
 * │ «el cierre de mes sigue igual, sin la opción de editar o    │
 * │  revisar o mirar y analizar. Saber qué quedó pendiente, de  │
 * │  una da cerrar mes y te avisa los pedidos que faltan por    │
 * │  cambiar de estado para tener un cierre limpio».            │
 * │                                                            │
 * │ El flujo era: `prompt` pidiendo AAAA-MM → el servidor dice  │
 * │ «quedan 14 sin resolver» → `confirm` → cerrado. Le decía    │
 * │ CUÁNTOS y nunca CUÁLES, y con catorce pedidos anónimos lo   │
 * │ único que se puede hacer es aceptar. Nova le hacía CERRAR   │
 * │ PARA ENTERARSE, y cerrar congela las cifras para siempre.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ NO ES UN ARCHIVO ─────────────────────────────────┐
 * │                                                            │
 * │ Ella propuso descargar un archivo. El diagnóstico estaba    │
 * │ bien —el cuadrito no sirve— pero un archivo es un callejón  │
 * │ sin salida para una tarea que es ARREGLAR COSAS: se abre    │
 * │ en otra parte, no tiene botones, y los catorce pedidos      │
 * │ siguen sin corregirse.                                      │
 * │                                                            │
 * │ La aserción que sostiene esa decisión es la 3: cambiar el   │
 * │ estado DESDE la revisión, y que las cifras de arriba se     │
 * │ rehagan con ese pedido dentro. Eso es lo que un archivo no  │
 * │ puede hacer.                                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/**
 * El mes que se revisa es el ANTERIOR al de hoy, que es el que la
 * pantalla propone. Se calcula, no se escribe: escrito, esta prueba
 * pasaría este mes y fallaría el que viene.
 */
const d = new Date(); d.setMonth(d.getMonth() - 1);
const MES = d.toISOString().slice(0, 7);

const ped = (id, cliente, estado, valor, dias) => ({
  id: id, idExterno: 'X' + id, cliente: cliente, producto: 'TÓNICO',
  ciudad: 'Quito', valor: valor, estado: estado, gestora: 'Sara',
  fecha: MES + '-10', diasQuieto: dias,
});

/** Un mes con tres pedidos colgando, uno de ellos muy quieto. */
const CONPENDIENTES = {
  ok: true, previo: true, tienda: 'ec', mes: MES, moneda: 'USD',
  cerrado: false, cerrado_en: '', mesEnCurso: false, limpio: false,
  actual: { ventas: 4000, pedidos: 100, entregados: 60, devueltos: 10,
            despachados: 80, efectividad: 85.7, tasaDevolucion: 12.5,
            gasto: 1000, margen: 900, pendientes: 3 },
  anterior: { ventas: 5000, pedidos: 110, entregados: 70, devueltos: 8,
              gasto: 1200, margen: 1100 },
  abiertos: [ped('p1', 'ANA QUIETA', 'en_transito', 50, 19),
             ped('p2', 'BEA MEDIANA', 'pendiente', 70, 4),
             ped('p3', 'CARO NUEVA', 'novedad', 30, 1)],
  porEstado: [{ estado: 'en_transito', n: 1, valor: 50 },
              { estado: 'pendiente', n: 1, valor: 70 },
              { estado: 'novedad', n: 1, valor: 30 }],
  valorAbierto: 150,
};
/** El mismo mes, ya sin nada colgando: es lo que devuelve tras arreglarlo. */
const LIMPIO = JSON.parse(JSON.stringify(CONPENDIENTES));
LIMPIO.limpio = true; LIMPIO.abiertos = []; LIMPIO.porEstado = [];
LIMPIO.valorAbierto = 0; LIMPIO.actual.pendientes = 0;
LIMPIO.actual.ventas = 4050;          // entró la venta del pedido arreglado
LIMPIO.actual.entregados = 61;

const YACERRADO = JSON.parse(JSON.stringify(LIMPIO));
YACERRADO.cerrado = true; YACERRADO.cerrado_en = MES + '-28T10:00:00Z';

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage({ viewport: { width: 1280, height: 1000 } });
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message));
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');

  await p.evaluate(({ CONPENDIENTES }) => {
    window.__llamadas = [];
    window.__resp = { cierre_previo: CONPENDIENTES };
    window.fetch = async (u, o) => {
      const c = JSON.parse(o.body);
      window.__llamadas.push(c);
      const r = window.__resp[c.accion];
      if (r) return { text: async () => JSON.stringify(r) };
      if (c.accion === 'escribir') return { text: async () => JSON.stringify({ ok: true }) };
      if (c.accion === 'cerrarmes') {
        return { text: async () => JSON.stringify({ ok: true, datos: { ventas: 1 } }) };
      }
      return { text: async () => JSON.stringify({ ok: true }) };
    };
    document.getElementById('login').style.display = 'none';
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO=true; TOKEN="t"; ST="ec"; ROL="dueno";');
    (0, eval)('STORES.ec = { name: "Nutrea EC", sym: "$", moneda: "USD" };');
  }, { CONPENDIENTES });

  console.log('\n══ 1 · EL BOTÓN ABRE UNA REVISIÓN, NO CIERRA NADA ══');
  /**
   * La aserción que evita el peor error posible: tocar «Cerrar el mes»
   * ya no puede cerrar nada. Cerrar congela las cifras para siempre.
   */
  await p.evaluate(() => cerrarMesActual());
  await p.waitForTimeout(400);
  let st = await p.evaluate(() => ({
    abierta: document.getElementById('cierre-prev').classList.contains('on'),
    acciones: window.__llamadas.map((c) => c.accion),
  }));
  ok('la revisión se abre', st.abierta);
  igual('y lo único que pidió es la revisión', ['cierre_previo'], st.acciones);
  ok('no cerró nada', st.acciones.indexOf('cerrarmes') === -1,
     JSON.stringify(st.acciones));
  /** Y el mes ya no se escribe a mano: un dedo torcido cerraba otro mes. */
  igual('el mes que propone es el anterior', MES,
        await p.evaluate(() => window.__llamadas[0].mes));

  console.log('\n══ 2 · DICE CUÁLES, NO SOLO CUÁNTOS ══');
  const texto = await p.evaluate(() =>
    (document.getElementById('cp-body').innerText || '').replace(/\s+/g, ' '));
  ok('sale el análisis del mes, no solo el aviso',
     /VENTAS/.test(texto) && /ENTREGADOS/.test(texto) && /MARGEN/.test(texto),
     texto.slice(0, 150));
  ok('dice cuántos faltan y cuánta plata es',
     /3 pedidos no tienen desenlace/.test(texto) && /150/.test(texto),
     texto.slice(0, 260));
  /** Lo que el cuadrito viejo nunca hizo. */
  ['ANA QUIETA', 'BEA MEDIANA', 'CARO NUEVA'].forEach((n) => {
    ok('y nombra a ' + n, texto.indexOf(n) !== -1);
  });
  ok('lo más quieto va primero', texto.indexOf('ANA QUIETA') < texto.indexOf('CARO NUEVA'));
  igual('cada uno con su selector para cambiarlo', 3,
        await p.evaluate(() => document.querySelectorAll('.cp-fila select').length));
  ok('el botón avisa que cerraría sucio',
     /con 3 sin resolver/.test(await p.evaluate(() =>
       document.getElementById('cp-cerrar').textContent)));

  console.log('\n══ 3 · Y SE ARREGLAN AHÍ MISMO ══');
  /**
   * Esta es la que sostiene la decisión de no hacer un archivo: se
   * cambia el estado sin salir, y las CIFRAS DE ARRIBA se rehacen con
   * ese pedido dentro. Un archivo descargado no puede hacer ninguna de
   * las dos cosas.
   */
  await p.evaluate(() => { window.__llamadas = []; window.__resp.cierre_previo = null; });
  await p.evaluate((LIMPIO) => { window.__resp.cierre_previo = LIMPIO; }, LIMPIO);
  await p.evaluate(() => cambiarEstadoDesdeCierre(0, 'entregado'));
  await p.waitForTimeout(500);
  st = await p.evaluate(() => ({
    acciones: window.__llamadas.map((c) => c.accion),
    escritura: window.__llamadas.filter((c) => c.accion === 'escribir')[0] || null,
    texto: (document.getElementById('cp-body').innerText || '').replace(/\s+/g, ' '),
    boton: document.getElementById('cp-cerrar').textContent,
  }));
  igual('guarda el estado del pedido que se tocó',
        { entidad: 'Pedidos', id: 'p1', campos: { estado_nova: 'entregado' } },
        st.escritura && { entidad: st.escritura.entidad, id: st.escritura.id,
                          campos: st.escritura.campos });
  ok('y vuelve a pedir la revisión, para rehacer las cifras',
     st.acciones.indexOf('cierre_previo') !== -1, JSON.stringify(st.acciones));
  ok('ahora dice que el cierre está limpio', /Cierre limpio/.test(st.texto),
     st.texto.slice(0, 160));
  ok('ANA QUIETA ya no está en la lista', st.texto.indexOf('ANA QUIETA') === -1);
  ok('y las ventas se rehicieron con su venta dentro', /4\.050|4050/.test(st.texto),
     st.texto.slice(0, 200));
  ok('el botón ya no avisa de pendientes', !/sin resolver/.test(st.boton), st.boton);

  console.log('\n══ 4 · CERRAR ES EL ÚLTIMO PASO, Y SOLO SI SE TOCA ══');
  await p.evaluate(() => { window.__llamadas = []; });
  await p.evaluate(() => confirmarCierre());
  await p.waitForTimeout(400);
  st = await p.evaluate(() => ({
    acciones: window.__llamadas.map((c) => c.accion),
    cerrarmes: window.__llamadas.filter((c) => c.accion === 'cerrarmes')[0] || null,
    abierta: document.getElementById('cierre-prev').classList.contains('on'),
  }));
  ok('ahora sí cierra', st.acciones.indexOf('cerrarmes') !== -1,
     JSON.stringify(st.acciones));
  igual('el mes que cierra es el que se estaba revisando', MES,
        st.cerrarmes && st.cerrarmes.mes);
  /**
   * Con el cierre limpio no hace falta forzar. Forzar siempre haría que
   * el aviso de «quedan pendientes» del servidor no sirviera de nada.
   */
  igual('y no fuerza, porque no hay nada que forzar', false,
        !!(st.cerrarmes && st.cerrarmes.forzar));
  ok('y se cierra el panel', !st.abierta);

  console.log('\n══ 5 · UN MES YA CERRADO NO SE VUELVE A CERRAR ══');
  await p.evaluate((YACERRADO) => { window.__resp.cierre_previo = YACERRADO; }, YACERRADO);
  await p.evaluate((MES) => abrirRevisionCierre(MES), MES);
  await p.waitForTimeout(400);
  st = await p.evaluate(() => ({
    texto: (document.getElementById('cp-body').innerText || '').replace(/\s+/g, ' '),
    apagado: document.getElementById('cp-cerrar').disabled,
  }));
  ok('lo dice', /Ya lo cerraste/.test(st.texto), st.texto.slice(0, 140));
  ok('y el botón de cerrar está apagado', st.apagado);
  await p.evaluate(() => { window.__llamadas = []; });
  await p.evaluate(() => confirmarCierre());
  await p.waitForTimeout(250);
  igual('tocarlo igual no manda nada', [],
        await p.evaluate(() => window.__llamadas.map((c) => c.accion)));

  igual('sin errores de consola', [], errs);
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
