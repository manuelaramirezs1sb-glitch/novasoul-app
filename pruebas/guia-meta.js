const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  for (const [nom, w] of [['ancho', 900], ['movil', 390]]) {
    const p = await b.newPage({ viewport: { width: w, height: 900 } });
    const errores = [];
    p.on('pageerror', e => errores.push('JS: ' + e.message));
    await p.goto('file:///home/claude/repo/guia-meta.html');
    await p.waitForLoadState('load');

    // imágenes que no cargaron
    const imgsMalas = await p.$$eval('img', els =>
      els.filter(i => !i.complete || i.naturalWidth === 0).map(i => i.getAttribute('src')));

    // scroll lateral
    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);

    // anclas del mapa que no existen
    const anclasRotas = await p.$$eval('a[href^="#"]', els =>
      els.map(a => a.getAttribute('href').slice(1))
         .filter(id => id && !document.getElementById(id)));

    // texto que se sale de su caja
    const desborde = await p.evaluate(() => {
      const malos = [];
      document.querySelectorAll('section, .nota, .ojo, table, pre, code').forEach(el => {
        if (el.scrollWidth > el.clientWidth + 2 && getComputedStyle(el).overflowX === 'visible')
          malos.push((el.tagName + '.' + el.className).slice(0, 40));
      });
      return malos;
    });

    const secciones = await p.$$eval('section', e => e.length);
    console.log(`[${nom} ${w}px] secciones=${secciones} scrollLateral=${lateral}px`);
    if (errores.length)     console.log('  ERRORES JS:', errores);
    if (imgsMalas.length)   console.log('  IMGS SIN CARGAR:', imgsMalas);
    if (anclasRotas.length) console.log('  ANCLAS ROTAS:', anclasRotas);
    if (desborde.length)    console.log('  DESBORDE:', desborde);
    if (!errores.length && !imgsMalas.length && !anclasRotas.length && !desborde.length && lateral === 0)
      console.log('  OK');

    await p.screenshot({ path: OUT + 'guia-' + nom + '.png', fullPage: false });
    // captura del paso 7
    await p.evaluate(() => document.getElementById('p7').scrollIntoView());
    await p.screenshot({ path: OUT + 'guia-' + nom + '-p7.png' });
    await p.close();
  }
  await b.close();
})();
