/**
 * LAS TIENDAS DE NOVASOUL, TODAS, EN CHICO.
 *
 * ┌─ LO QUE PIDIÓ ─────────────────────────────────────────────┐
 * │                                                            │
 * │ «en la pantalla de soul en resumen de tiendas pones unos    │
 * │  cuadros para cada tienda pero más pequeños».               │
 * │                                                            │
 * │ Antes había que ir tienda por tienda con los botones de     │
 * │ arriba para saber si alguna estaba en rojo.                 │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y LA REGLA QUE NO SE PODÍA ROMPER ────────────────────────┐
 * │                                                            │
 * │ «dentro de Nova nunca se comparan dos tiendas al mismo      │
 * │  tiempo en una misma pantalla» — regla suya, y con razón:   │
 * │ 4.000 quetzales al lado de 1.200 dólares no significan      │
 * │ nada.                                                       │
 * │                                                            │
 * │ Por eso las tarjetas llevan ESTADO y no PLATA. Esta prueba  │
 * │ lo afirma: si algún día alguien le mete una cifra de        │
 * │ dinero a una tarjeta, falla.                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Y mide el ANCHO y el COLOR en pantalla, no en el CSS. Las dos
 * primeras versiones estaban mal y el CSS se veía razonable en las dos:
 *
 *   · las tarjetas salían de 570px —«más pequeños», decía ella— porque
 *     `flex:1 1 150px` con dos tiendas las estira a media pantalla;
 *   · y el verde salía DORADO, porque lo pinté con `var(--acc)` y
 *     `--acc` cambia con el tema que ella elija. Un semáforo donde
 *     «bien» y «regular» son del mismo color no es un semáforo.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const luz = (etiqueta, estado) => ({ etiqueta, estado, valor: 1, porque: 'x', delta: null });
/** Tres tiendas: una en rojo, una tranquila, y una que no se pudo leer. */
const FAM = {
  ok: true,
  clientes: [{ id: 'c1', empresa: 'Nutrea', sheetId: 'h1' }],
  cliente: { id: 'c1', empresa: 'Nutrea', sheetId: 'h1' },
  tiendas: [{ id: 'ec', nombre: 'Nutrea EC' }, { id: 'gt', nombre: 'Nutrea Guatemala' },
            { id: 'mx', nombre: 'Nutrea MX' }],
  tienda: 'ec',
  alarmas: [{ id: 'a1', nivel: 'mal', nombre: 'n', titulo: 'Tres pedidos sin mover',
              detalle: 'llevan 5 días', casos: 3 }],
  semaforo: { tienda: 'ec', moneda: 'USD',
              semana: { lunes: '2026-09-28', domingo: '2026-10-04' },
              enCurso: true, hastaDia: '2026-09-30', diasCorridos: 3,
              luces: [luz('CPA por entrega', 'verde'), luz('Ticket promedio', 'amarillo'),
                      luz('Tasa de entrega', 'rojo'), luz('Utilidad real', 'sin_medir')],
              alertas: [{ nivel: 'rojo', titulo: 'La entrega cayó', accion: 'Revisa el courier' }],
              hayDatos: true },
  resumen: [
    { id: 'ec', nombre: 'Nutrea EC', error: '', alarmas: 1, peor: 'rojo', hayDatos: true,
      luces: [luz('CPA', 'verde'), luz('Ticket', 'amarillo'), luz('Entrega', 'rojo'),
              luz('Utilidad', 'sin_medir')] },
    { id: 'gt', nombre: 'Nutrea Guatemala', error: '', alarmas: 0, peor: 'verde',
      hayDatos: false,
      luces: [luz('CPA', 'sin_medir'), luz('Ticket', 'sin_medir'),
              luz('Entrega', 'sin_medir'), luz('Utilidad', 'sin_medir')] },
    { id: 'mx', nombre: 'Nutrea MX', error: 'sin permiso', alarmas: 0, peor: '',
      hayDatos: false, luces: [] },
  ],
  central: null, academy: null, errorTienda: '',
};

/** Cifras de dinero: lo que NO puede aparecer en una tarjeta. */
const PLATA = /\$|USD|GTQ|COP|MXN|\d[\d.,]{3,}/;

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const errs = [];

  const abrir = async (w) => {
    const p = await b.newPage({ viewport: { width: w, height: 1100 } });
    p.on('pageerror', (e) => errs.push(e.message));
    await p.goto('file:///home/claude/repo/novasoul.html');
    await p.waitForLoadState('load');
    await p.evaluate((FAM) => {
      (0, eval)('FAM = ' + JSON.stringify(FAM));
      const lg = document.getElementById('login-screen');
      if (lg) lg.style.display = 'none';
      const app = document.getElementById('app');
      if (app) app.style.display = 'block';
      /**
       * La vista tiene que estar VISIBLE o los anchos salen todos en 0,
       * y entonces «ninguna pasa de 240px» pasa por la razón
       * equivocada. Me pasó: la aserción estaba verde con las tarjetas
       * sin dibujar. Por eso hay también un mínimo.
       */
      const c = document.getElementById('fam-cuerpo');
      if (c && c.closest('.view')) c.closest('.view').style.display = 'block';
      pintarFamily();
    }, FAM);
    await p.waitForTimeout(220);
    return p;
  };

  console.log('\n══ 1 · UNA TARJETA POR TIENDA, Y CHICA ══');
  let p = await abrir(1200);
  let t = await p.evaluate(() => {
    const cs = [...document.querySelectorAll('.fam-t')];
    return {
      n: cs.length,
      textos: cs.map((c) => (c.innerText || '').replace(/\s+/g, ' ').trim()),
      anchos: cs.map((c) => Math.round(c.getBoundingClientRect().width)),
      puntos: cs.map((c) => c.querySelectorAll('.fam-pt').length),
      haySelectorDeTiendas: !!document.querySelector('#fam-sel .kbtn'),
    };
  });
  igual('sale una por tienda, las tres', 3, t.n);
  /**
   * El número concreto no importa; que sean CHICAS sí. La primera
   * versión daba 570px cada una con dos tiendas, que es justo lo
   * contrario de lo que pidió.
   */
  ok('ninguna pasa de 240px de ancho', t.anchos.every((a) => a <= 240),
     JSON.stringify(t.anchos));
  ok('ni baja de 120, que ya no se leería', t.anchos.every((a) => a >= 120),
     JSON.stringify(t.anchos));
  ok('cada nombre de tienda está en la suya',
     t.textos[0].indexOf('Nutrea EC') === 0 &&
     t.textos[1].indexOf('Nutrea Guatemala') === 0 &&
     t.textos[2].indexOf('Nutrea MX') === 0, JSON.stringify(t.textos));
  ok('la que no se pudo leer lo dice en vez de fingir estar bien',
     /no pude leerla/.test(t.textos[2]), t.textos[2]);
  ok('y no le inventa luces', t.puntos[2] === 0, String(t.puntos[2]));
  ok('con las tiendas en tarjetas, los botones de tienda sobran',
     !t.haySelectorDeTiendas);

  console.log('\n══ 2 · ESTADO, NUNCA PLATA ══');
  /**
   * La aserción que sostiene la decisión: la regla de «nunca dos tiendas
   * comparándose» se respeta porque en las tarjetas no hay cifras que
   * comparar. Si alguien le mete una, esto falla.
   */
  const conPlata = t.textos.filter((x) => PLATA.test(x));
  igual('ninguna tarjeta enseña dinero', [], conPlata);

  console.log('\n══ 3 · EL SEMÁFORO SE DISTINGUE POR COLOR ══');
  const col = await p.evaluate(() => {
    const c = document.querySelectorAll('.fam-t')[0];
    return [...c.querySelectorAll('.fam-pt')].map((x) => getComputedStyle(x).backgroundColor);
  });
  // [alarmas, CPA verde, Ticket amarillo, Entrega rojo, Utilidad sin medir]
  const distintos = new Set([col[1], col[2], col[3]]).size;
  ok('verde, amarillo y rojo son TRES colores distintos', distintos === 3,
     JSON.stringify([col[1], col[2], col[3]]));
  /**
   * Y el verde tiene que ser verde. Pintado con `var(--acc)` salía
   * dorado —el acento del tema— e igualaba al amarillo.
   */
  const rgb = (s) => (s.match(/\d+/g) || []).map(Number);
  const v = rgb(col[1]);
  ok('el verde es verde de verdad, no el dorado del tema',
     v[1] > v[0] && v[1] > v[2], col[1]);

  console.log('\n══ 4 · EN TELÉFONO TAMBIÉN ══');
  await p.close();
  p = await abrir(390);
  t = await p.evaluate(() => ({
    n: document.querySelectorAll('.fam-t').length,
    anchos: [...document.querySelectorAll('.fam-t')]
      .map((c) => Math.round(c.getBoundingClientRect().width)),
    desborda: document.documentElement.scrollWidth > window.innerWidth + 1,
  }));
  igual('siguen estando las tres', 3, t.n);
  ok('no empuja la página de lado', !t.desborda);
  ok('y siguen siendo legibles', t.anchos.every((a) => a >= 120), JSON.stringify(t.anchos));

  igual('sin errores de consola', [], errs);
  await p.close();
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
