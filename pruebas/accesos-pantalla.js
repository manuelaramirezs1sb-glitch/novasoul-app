/**
 * LOS ACCESOS, EN LA PANTALLA.
 *
 * ┌─ QUÉ SE ESTÁ COMPROBANDO ──────────────────────────────────┐
 * │                                                            │
 * │ «okay, si ya te quedó claro lo de los accesos entonces es   │
 * │  posible que se acomoden en el lugar donde van dentro de     │
 * │  Nova?» Y de la gestora: «tener el link directo para        │
 * │  entrar a la parte de accesos».                            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y LAS DOS COSAS QUE NO SON DECORACIÓN ────────────────────┐
 * │                                                            │
 * │ 1. La contraseña NO está en el HTML mientras está tapada.   │
 * │    Taparla con CSS —color del fondo, `-webkit-text-security`│
 * │    o un `opacity`— se ve igual y no protege de nada: sigue   │
 * │    ahí para cualquiera que mire el inspector, y para una    │
 * │    captura de pantalla compartida. Aquí se mide el TEXTO.   │
 * │                                                            │
 * │ 2. Los enlaces salen con `target=_blank` y `rel=noopener`.  │
 * │    Sin `noopener`, la página que se abre puede redirigir la │
 * │    pestaña de Nova a una copia que pida la contraseña otra  │
 * │    vez. Es un atributo, no una opinión: o está o no está.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const CLAVE = 'LaQueSeaAqui2026';
const ACCESOS = {
  tienda: 'ec', plataforma: 'Dropi',
  url: 'https://app.dropi.co/dashboard/home',
  usuario: 'gestion.vip@correodealejandro.com',
  clave: CLAVE,
  correo_codigo: 'alejandro.escobar@correodealejandro.com',
  canal_nombre: 'Chat Center', canal_url: 'https://chateapro.app/login',
  nota: 'La subcuenta solo confirma y gestiona novedades.',
  actualizado_en: '2026-09-25 09:00:00', actualizado_por: 'sara@x.com',
};

async function abrir(b, rol, accesos) {
  const p = await b.newPage({ viewport: { width: 1280, height: 1100 } });
  const errores = [];
  p.on('pageerror', e => errores.push(e.message));
  p.on('console', m => {
    if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) {
      errores.push('console: ' + m.text());
    }
  });
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');

  const enviados = [];
  await p.exposeFunction('__anotar', (x) => { enviados.push(JSON.parse(x)); });

  await p.evaluate(({ rol, accesos }) => {
    document.getElementById('login').style.display = 'none';
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO = true; ST = "ec"; ROL = "' + rol + '";');
    document.documentElement.setAttribute('data-rol', rol);
    (0, eval)('SESION = ' + JSON.stringify({
      nombre: rol === 'gestora' ? 'Andrea Ramírez' : 'Sara Escobar', rol: rol,
      correo: rol === 'gestora' ? 'andrea@x.com' : 'sara@x.com',
      tiendas: ['ec'], permisos: [], modulos: ['empresarial'] }));
    window.api = async (accion, params) => {
      window.__anotar(JSON.stringify({ accion: accion, params: params || {} }));
      if (accion === 'accesos') {
        return { ok: true, tienda: 'ec', nombreTienda: 'Nutrea Ecuador',
                 accesos: JSON.parse(JSON.stringify(accesos)),
                 puedeEditar: rol === 'dueno',
                 hay: !!(accesos.url || accesos.canal_url) };
      }
      if (accion === 'accesos_guardar') {
        return { ok: true, accesos: Object.assign({}, accesos, (params || {}).cambios || {}) };
      }
      return { ok: true };
    };
  }, { rol, accesos });

  return { p, enviados, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 1 · LA GESTORA ENTRA Y VE SUS DOS PUERTAS ══');
  {
    const { p, errores } = await abrir(b, 'gestora', ACCESOS);
    await p.evaluate(() => cargarAccesos());
    await p.waitForTimeout(250);

    console.log('\n── El atajo en su panel, donde ya está parada ──');
    await p.evaluate(() => go('g-hoy', null));
    await p.waitForTimeout(150);
    const atajo = await p.$$eval('#g-accesos a', els => els.map(e => ({
      txt: e.textContent.trim(), href: e.getAttribute('href'),
      target: e.getAttribute('target'), rel: e.getAttribute('rel'),
    })));
    igual('dos botones: la plataforma y el chat', 2, atajo.length);
    igual('el primero abre Dropi', 'https://app.dropi.co/dashboard/home',
          atajo[0] && atajo[0].href);
    igual('con su nombre', 'Abrir Dropi', atajo[0] && atajo[0].txt);
    igual('el segundo abre el chat de las compradoras', 'https://chateapro.app/login',
          atajo[1] && atajo[1].href);
    ok('los dos en pestaña nueva y con noopener',
       atajo.every(a => a.target === '_blank' && /noopener/.test(a.rel || '')),
       JSON.stringify(atajo));
    ok('y el link directo a los accesos completos',
       /Usuario y contraseña/.test(await p.textContent('#g-accesos')),
       await p.textContent('#g-accesos'));
    ok('el atajo se ve de verdad', await p.evaluate(
       () => document.getElementById('g-accesos').offsetParent !== null));

    console.log('\n── Y tiene la entrada de menú ──');
    ok('«Accesos» está en su menú', await p.evaluate(() => {
      const el = document.querySelector('.ni[data-v="accesos"]');
      return !!el && el.offsetParent !== null;
    }));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 2 · LA CONTRASEÑA NO ESTÁ EN LA PANTALLA HASTA QUE SE PIDE ══');
  {
    const { p, errores } = await abrir(b, 'gestora', ACCESOS);
    await p.evaluate(() => { go('accesos', null); });
    await p.waitForTimeout(300);

    let html = await p.content();
    ok('la contraseña NO está en el HTML', html.indexOf(CLAVE) === -1);
    ok('se ven puntos en su lugar',
       /•/.test(await p.textContent('#acc-cards')),
       (await p.textContent('#acc-cards')).slice(0, 300));
    igual('con la etiqueta CONTRASEÑA', true,
          /CONTRASEÑA/.test(await p.textContent('#acc-cards')));

    console.log('\n── Y al pedirla, aparece ──');
    await p.evaluate(() => accVerClave());
    await p.waitForTimeout(150);
    html = await p.content();
    ok('ahora sí está', html.indexOf(CLAVE) !== -1);
    ok('y el botón invita a ocultarla',
       /ocultar/.test(await p.textContent('#acc-cards')));

    await p.evaluate(() => accVerClave());
    await p.waitForTimeout(150);
    ok('se puede volver a tapar', (await p.content()).indexOf(CLAVE) === -1);

    console.log('\n── El aviso del código, que es el error más caro ──');
    const t = (await p.textContent('#acc-cards')).replace(/\s+/g, ' ');
    ok('avisa que va a pedir un código', /código/i.test(t), t.slice(0, 400));
    ok('y dice a qué correo llega',
       t.indexOf('alejandro.escobar@correodealejandro.com') !== -1, t.slice(0, 500));
    ok('y que no cambie la contraseña por eso',
       /no la cambies/i.test(t), t.slice(0, 600));

    console.log('\n── Lo que la gestora NO puede hacer ──');
    ok('no ve el formulario de cambio', !(await p.isVisible('#acc-editar-card')));
    igual('y no se pinta nada dentro', '', await p.textContent('#acc-form'));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 3 · LA DUEÑA LOS CAMBIA ══');
  {
    const { p, enviados, errores } = await abrir(b, 'dueno', ACCESOS);
    await p.evaluate(() => { go('accesos', null); });
    await p.waitForTimeout(300);

    ok('ve el formulario', await p.isVisible('#acc-editar-card'));
    igual('con la plataforma ya puesta', 'Dropi', await p.inputValue('#acc-plataforma'));
    igual('el enlace', 'https://app.dropi.co/dashboard/home', await p.inputValue('#acc-url'));
    igual('el correo del código', 'alejandro.escobar@correodealejandro.com',
          await p.inputValue('#acc-correo_codigo'));
    igual('y el enlace del chat', 'https://chateapro.app/login',
          await p.inputValue('#acc-canal_url'));
    ok('dice quién los cambió la última vez',
       /sara@x.com/.test(await p.textContent('#acc-form')),
       await p.textContent('#acc-form'));
    ok('y avisa que la contraseña no va a la bitácora',
       /bitácora/i.test(await p.textContent('#acc-form')),
       await p.textContent('#acc-form'));

    await p.fill('#acc-canal_nombre', 'WhatsApp de la tienda');
    await p.click('#acc-form .nsavebtn');
    await p.waitForTimeout(300);

    const g = enviados.filter(x => x.accion === 'accesos_guardar');
    igual('guardar manda los accesos', 1, g.length);
    igual('con el nombre nuevo del canal', 'WhatsApp de la tienda',
          g[0] && g[0].params.cambios.canal_nombre);
    igual('y la tienda en la que está', 'ec', g[0] && g[0].params.tienda);
    ok('no manda la contraseña vacía por accidente',
       g[0] && g[0].params.cambios.clave === CLAVE,
       JSON.stringify(g[0] && g[0].params.cambios.clave));
    ok('avisa que quedó guardado',
       /guardados/i.test(await p.textContent('#toast')), await p.textContent('#toast'));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 4 · SIN ACCESOS TODAVÍA, CADA UNA SABE QUÉ HACER ══');
  const VACIO = { tienda: 'ec', plataforma: '', url: '', usuario: '', clave: '',
                  correo_codigo: '', canal_nombre: '', canal_url: '', nota: '',
                  actualizado_en: '', actualizado_por: '' };
  {
    const { p, errores } = await abrir(b, 'dueno', VACIO);
    await p.evaluate(() => { go('accesos', null); });
    await p.waitForTimeout(300);
    const t = (await p.textContent('#acc-cards')).replace(/\s+/g, ' ');
    ok('a la dueña se le dice que los llene', /Llénalos abajo/.test(t), t);
    ok('nombrando su tienda', /Nutrea Ecuador/.test(t), t);
    ok('y el formulario está listo', await p.isVisible('#acc-editar-card'));
    igual('sin errores de consola', [], errores);
    await p.close();
  }
  {
    const { p, errores } = await abrir(b, 'gestora', VACIO);
    await p.evaluate(() => { go('accesos', null); });
    await p.waitForTimeout(300);
    const t = (await p.textContent('#acc-cards')).replace(/\s+/g, ' ');
    ok('a la gestora se le dice a quién pedírselos', /dueña/.test(t), t);
    ok('sin decirle que los llene, que no puede', !/Llénalos/.test(t), t);

    await p.evaluate(() => go('g-hoy', null));
    await p.waitForTimeout(150);
    const atajo = (await p.textContent('#g-accesos')).replace(/\s+/g, ' ');
    igual('y el atajo de su panel no queda como un botón roto', 0,
          (await p.$$('#g-accesos a')).length);
    ok('sino diciendo qué falta', /no ha puesto los accesos/.test(atajo), atajo);

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 5 · SOLO LA PLATAFORMA, SIN CANAL ══');
  {
    const solo = Object.assign({}, ACCESOS, { canal_nombre: '', canal_url: '' });
    const { p, errores } = await abrir(b, 'gestora', solo);
    await p.evaluate(() => { go('accesos', null); });
    await p.waitForTimeout(300);
    const t = (await p.textContent('#acc-cards')).replace(/\s+/g, ' ');
    ok('la plataforma se ve', /Abrir Dropi/.test(t), t.slice(0, 200));
    ok('y no hay tarjeta de canal vacía', !/compradoras/i.test(t), t.slice(0, 400));
    igual('un solo botón de abrir', 1, (await p.$$('#acc-cards a.acc-abrir')).length);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
