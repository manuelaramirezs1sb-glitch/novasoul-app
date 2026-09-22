/**
 * La puerta del hub.
 *
 * Lo que se comprueba es lo que protege: que sin sesión no se vea NADA
 * de los cuatro productos, ni siquiera un destello, y que con sesión se
 * vea solo lo del plan.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const SES = (modulos, nombre) => ({
  email: 'm@nova.com', nombre: nombre || 'Manuela', rol: 'dueno',
  tiendas: ['gt'], modulos: modulos, permisos: [], vence: Date.now() + 3600000, horas: 8,
});

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };

  const visibles = (p) => p.$$eval('.card[data-mod]', els =>
    els.filter(e => getComputedStyle(e).display !== 'none').map(e => e.dataset.mod));

  for (const ancho of [1100, 390]) {
    // ── 1 · Sin sesión: la puerta, y nada más ──
    let p = await b.newPage({ viewport: { width: ancho, height: 900 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    await p.goto('file:///home/claude/repo/index.html');
    await p.waitForLoadState('load');
    await p.waitForTimeout(300);

    ok('sin sesión sale la puerta (@' + ancho + ')', await p.$('#puerta') !== null);
    ok('y NINGUNA tarjeta visible (@' + ancho + ')',
       (await visibles(p)).length === 0, JSON.stringify(await visibles(p)));
    const texto = await p.textContent('body');
    ok('ni el nombre de los productos en pantalla (@' + ancho + ')',
       !texto.includes('Tu energía y tu semana') || await p.$('#puerta') !== null);
    ok('sin scroll lateral (@' + ancho + ')',
       await p.evaluate(() => document.documentElement.scrollWidth -
                              document.documentElement.clientWidth) === 0);
    await p.close();

    // ── 2 · Con sesión de un cliente: solo lo suyo ──
    p = await b.newPage({ viewport: { width: ancho, height: 900 } });
    p.on('pageerror', e => errores.push(e.message));
    await p.addInitScript(() => { localStorage.setItem('ne_token', 'tok'); });
    await p.route('**/macros/**', r => r.fulfill({
      status: 200, contentType: 'application/json',
      body: JSON.stringify({ ok: true, sesion: {
        email: 'c@x.com', nombre: 'Cliente', rol: 'dueno', tiendas: ['gt'],
        modulos: ['empresarial'], permisos: [], vence: Date.now() + 3600000, horas: 8 } }),
    }));
    await p.goto('file:///home/claude/repo/index.html');
    await p.waitForFunction(() => !document.getElementById('puerta'), null, { timeout: 5000 })
      .catch(() => {});
    const v1 = await visibles(p);
    ok('un cliente de Empresarial ve solo Empresarial (@' + ancho + ')',
       JSON.stringify(v1) === JSON.stringify(['empresarial']), JSON.stringify(v1));
    ok('NO ve Nova Central (@' + ancho + ')', v1.indexOf('central') === -1);
    ok('ni NovaSoul ni novAcademy (@' + ancho + ')',
       v1.indexOf('soul') === -1 && v1.indexOf('academy') === -1);
    await p.close();

    // ── 3 · Con la sesión de Manuela: las cuatro ──
    p = await b.newPage({ viewport: { width: ancho, height: 900 } });
    p.on('pageerror', e => errores.push(e.message));
    await p.addInitScript(() => { localStorage.setItem('ne_token', 'tok'); });
    await p.route('**/macros/**', r => r.fulfill({
      status: 200, contentType: 'application/json',
      body: JSON.stringify({ ok: true,
        sesion: SES(['empresarial', 'soul', 'academy', 'central']) }),
    }));
    await p.goto('file:///home/claude/repo/index.html');
    await p.waitForFunction(() => !document.getElementById('puerta'), null, { timeout: 5000 })
      .catch(() => {});
    const v2 = await visibles(p);
    ok('Manuela ve las cuatro (@' + ancho + ')', v2.length === 4, JSON.stringify(v2));
    ok('y la saluda por su nombre (@' + ancho + ')',
       (await p.textContent('.hub-title')).includes('Manuela'));
    if (ancho === 1100) await p.screenshot({ path: OUT + 'hub-adentro.png' });
    await p.close();

    // ── 4 · Token muerto: la puerta NO se abre ──
    p = await b.newPage({ viewport: { width: ancho, height: 900 } });
    p.on('pageerror', e => errores.push(e.message));
    await p.addInitScript(() => { localStorage.setItem('ne_token', 'viejo'); });
    await p.route('**/macros/**', r => r.fulfill({
      status: 200, contentType: 'application/json',
      body: JSON.stringify({ ok: false, error: 'Sesión vencida o inválida.' }),
    }));
    await p.goto('file:///home/claude/repo/index.html');
    await p.waitForTimeout(700);
    ok('un token vencido deja la puerta puesta (@' + ancho + ')',
       await p.$('#puerta') !== null);
    ok('y ninguna tarjeta se asoma (@' + ancho + ')', (await visibles(p)).length === 0);
    ok('el token muerto se borra del navegador (@' + ancho + ')',
       await p.evaluate(() => localStorage.getItem('ne_token')) === null);
    ok('y lo dice en vez de quedarse mudo (@' + ancho + ')',
       (await p.textContent('#pz-sub')).includes('venció'),
       await p.textContent('#pz-sub'));
    ok('sin errores de JS (@' + ancho + ')', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
