/**
 * La tarjeta de Meta en la pantalla de la dueña.
 *
 * Comprueba lo que no se ve leyendo el código: que «Traer ahora» solo
 * aparezca cuando la conexión ya está comprobada, y que el informe
 * muestre los avisos en vez de tragárselos.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
                            else { fallas++; console.log('  FALLA  ' + n + (d ? ' — ' + d : '')); } };

  for (const ancho of [1100, 390]) {
    const p = await b.newPage({ viewport: { width: ancho, height: 1000 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    await p.goto('file:///home/claude/repo/empresarial.html');
    await p.waitForLoadState('load');

    // El estado mínimo que pintarMeta necesita, sin sesión ni servidor.
    await p.evaluate(() => {
      (0, eval)('ST = "gt"; STORES = { gt: { name: "Nutrea GT" } };');
      // La pantalla ya trae su propia caja: se usa esa, no una copia.
      if (!document.getElementById('meta-caja')) {
        const d = document.createElement('div'); d.id = 'meta-caja';
        document.body.appendChild(d);
      }
      if (!document.getElementById('meta-estado-lbl')) {
        const s = document.createElement('span'); s.id = 'meta-estado-lbl';
        document.body.appendChild(s);
      }
    });

    // 1 · Guardada pero sin probar: no debe haber botón de traer.
    await p.evaluate(() => {
      window.META_DATA = { hayLlave: true, ultimaPrueba: null, guardadaEn: '2026-09-22',
                           cuentas: [{ tienda: 'gt', cuenta: '123', moneda: 'GTQ' }] };
      (0, eval)('META_DATA = window.META_DATA;');
      pintarMeta();
    });
    ok('sin probar: no aparece «Traer ahora» (@' + ancho + ')',
       await p.$('#meta-dias') === null);

    // 2 · Conectada: sí aparece.
    await p.evaluate(() => {
      window.META_DATA = { hayLlave: true, guardadaEn: '2026-09-22',
        ultimaPrueba: { cuando: '2026-09-22', cuenta: '123', nombre: 'CP Nutrea' },
        cuentas: [{ tienda: 'gt', cuenta: '123', moneda: 'GTQ' }] };
      (0, eval)('META_DATA = window.META_DATA;');
      pintarMeta();
    });
    ok('conectada: aparece el selector de días (@' + ancho + ')',
       await p.$('#meta-dias') !== null);
    ok('y el botón (@' + ancho + ')',
       (await p.textContent('#meta-caja')).includes('Traer ahora'));

    // 3 · El informe, con un aviso que NO se puede perder.
    await p.evaluate(() => {
      pintarInformeMeta({
        desde: '2026-09-15', hasta: '2026-09-22', filas: 14, nuevas: 14,
        actualizadas: 0, iguales: 0, gasto: 1386315, moneda: 'COP',
        accionCompra: '', accionResultado: 'lead', solapadas: 2,
        avisos: ['Meta no reportó compras en estos conjuntos.',
                 '2 fila(s) de un Excel subido a mano se contarían dos veces.'],
      });
    });
    const texto = await p.textContent('#meta-informe');
    ok('el informe muestra los dos avisos (@' + ancho + ')',
       texto.includes('dos veces') && texto.includes('no reportó compras'), texto.slice(0, 120));
    ok('dice qué contó como compra, aunque sea ninguna (@' + ancho + ')',
       texto.includes('ninguna'));
    ok('y muestra el gasto con separadores (@' + ancho + ')',
       /1[.,]386[.,]315/.test(texto), texto.match(/COP[^A-Za-z]*/)?.[0]);

    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral (@' + ancho + ')', lateral === 0, lateral + 'px');
    ok('sin errores de JS (@' + ancho + ')', errores.length === 0, errores.join(' | '));

    if (ancho === 1100) {
      // La caja vive dentro de Configuración, que está oculta: se
      // destapan sus padres solo para la foto.
      await p.evaluate(() => {
        // Y se tapa el login, que está por encima de todo.
        document.querySelectorAll('.login, #login, #login-screen, .gate, #gate')
          .forEach(e => { e.style.display = 'none'; });
        let el = document.getElementById('meta-caja');
        while (el && el !== document.body) {
          el.style.display = el.tagName === 'SPAN' ? 'inline' : 'block';
          el.style.visibility = 'visible';
          el = el.parentElement;
        }
      });
      await p.locator('#meta-caja').first().screenshot({ path: OUT + 'meta-tarjeta.png' });
    }
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
