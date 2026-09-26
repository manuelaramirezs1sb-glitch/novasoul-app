/**
 * LA PUERTA DEL HUB, CUANDO GOOGLE CONTESTA RARO.
 *
 * ┌─ LO QUE PASÓ ──────────────────────────────────────────────┐
 * │                                                            │
 * │ Ella escribió su código de acceso y el hub contestó         │
 * │ «Falta la acción.» en rojo. No se podía entrar a Nova.      │
 * │                                                            │
 * │ No es un error del código: es cómo funciona Apps Script.    │
 * │ No responde el POST — contesta un redirect, el navegador lo │
 * │ sigue convirtiéndolo en GET, y en los minutos siguientes a  │
 * │ publicar una versión nueva ese GET llega pelado. `doGet`    │
 * │ contesta «Falta la acción» sobre una petición que sí era    │
 * │ un POST.                                                    │
 * │                                                            │
 * │ Empresarial ya sabía esperar y reintentar desde hacía       │
 * │ tiempo. El hub no: tiene su propia función de red, más      │
 * │ corta, y el arreglo nunca llegó hasta aquí.                 │
 * │                                                            │
 * │ Y el hub es LA PUERTA. Si falla, no se entra a nada —ni a   │
 * │ Empresarial, ni a NovaSoul, ni a la Universidad—. Que la    │
 * │ parte menos vigilada sea la que abre todo es exactamente    │
 * │ al revés de como debería.                                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** Lo que contesta Apps Script cuando el POST le llega como GET pelado. */
const SIN_ACCION = { ok: false, error: 'Falta la acción.',
  diagnostico: { metodo: 'GET', hubo_cuerpo: false, largo: 0, empieza: '',
                 tipo: '', parametros: [] } };

const SESION = { email: 'novasoul959@gmail.com', nombre: 'Manuela', rol: 'dueno',
                 tiendas: ['ec'], modulos: ['empresarial', 'soul'], permisos: [],
                 vence: Date.now() + 3600000, horas: 8 };

async function abrir(b, guion) {
  const p = await b.newPage({ viewport: { width: 1100, height: 900 } });
  const errores = [];
  p.on('pageerror', e => errores.push(e.message));
  await p.goto('file:///home/claude/repo/index.html');
  await p.waitForLoadState('load');
  await p.evaluate(({ guion }) => {
    window.__n = 0;
    window.fetch = async () => {
      const paso = guion[Math.min(window.__n, guion.length - 1)];
      window.__n++;
      if (paso === 'basura') return { text: async () => '<html>404</html>' };
      return { text: async () => JSON.stringify(paso) };
    };
  }, { guion });
  return { p, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  console.log('\n══ 1 · EL CASO DE LA FOTO: FALLA UNA VEZ Y DESPUÉS ENTRA ══');
  {
    const { p, errores } = await abrir(b, [SIN_ACCION,
      { ok: true, token: 't1', sesion: SESION }]);
    await p.evaluate(() => { PZ_EMAIL = 'novasoul959@gmail.com';
      document.getElementById('pz-cod').value = '355498'; pzVerificar(); });
    await p.waitForTimeout(3000);

    igual('reintentó: dos llamadas, no una', 2, await p.evaluate(() => window.__n));
    ok('y entró', await p.evaluate(() => !!localStorage.getItem('ne_token')),
       await p.evaluate(() => localStorage.getItem('ne_token')));
    // Al entrar, la puerta desaparece del DOM: el aviso no existe o
    // está vacío. Buscarlo con textContent se quedaba colgado 30s.
    const aviso = await p.evaluate(() => {
      const e = document.getElementById('pz-e2');
      return e ? (e.style.display === 'none' ? '' : e.textContent) : '';
    });
    ok('sin enseñar ningún error', !aviso.trim(), aviso);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 2 · SI NO SE CURA, LO DICE CON ALGO QUE SE PUEDE HACER ══');
  {
    const { p, errores } = await abrir(b, [SIN_ACCION]);
    await p.evaluate(() => { PZ_EMAIL = 'novasoul959@gmail.com';
      document.getElementById('pz-cod').value = '355498'; pzVerificar(); });
    await p.waitForTimeout(13000);

    const txt = await p.evaluate(() => {
      const e = document.getElementById('pz-e2');
      return e ? e.textContent : '';
    });
    ok('lo intentó varias veces', await p.evaluate(() => window.__n) >= 4,
       String(await p.evaluate(() => window.__n)));
    ok('no deja «Falta la acción» a secas', !/^Falta la acción\.$/.test(txt.trim()), txt);
    ok('explica que se está propagando', /propagando/.test(txt), txt);
    ok('y dice qué revisar en Apps Script', /Cualquier usuario/.test(txt), txt);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 3 · UNA PÁGINA DE ERROR DE GOOGLE TAMPOCO ROMPE NADA ══');
  /**
   * Mientras Google publica, /exec contesta una página HTML de 404 en
   * vez de JSON. `r.json()` reventaba con un error de sintaxis que no
   * decía nada útil.
   */
  {
    const { p, errores } = await abrir(b, ['basura', 'basura',
      { ok: true, token: 't2', sesion: SESION }]);
    await p.evaluate(() => { PZ_EMAIL = 'novasoul959@gmail.com';
      document.getElementById('pz-cod').value = '355498'; pzVerificar(); });
    await p.waitForTimeout(6000);
    ok('aguanta y acaba entrando',
       await p.evaluate(() => !!localStorage.getItem('ne_token')),
       String(await p.evaluate(() => window.__n)));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 4 · UN CÓDIGO MALO SIGUE SIENDO UN CÓDIGO MALO ══');
  /**
   * El reintento no puede tragarse los errores de verdad: si el código
   * está mal, hay que decirlo a la primera y no hacerla esperar doce
   * segundos para lo mismo.
   */
  {
    const { p, errores } = await abrir(b, [{ ok: false, error: 'Código incorrecto.' }]);
    await p.evaluate(() => { PZ_EMAIL = 'novasoul959@gmail.com';
      document.getElementById('pz-cod').value = '000000'; pzVerificar(); });
    await p.waitForTimeout(1200);
    igual('una sola llamada', 1, await p.evaluate(() => window.__n));
    const err4 = await p.evaluate(() => {
      const e = document.getElementById('pz-e2');
      return e ? e.textContent : '';
    });
    ok('y lo dice', /Código incorrecto/.test(err4), err4);
    ok('sin entrar', !(await p.evaluate(() => !!localStorage.getItem('ne_token'))));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
