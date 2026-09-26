/**
 * LA PUERTA · que cualquier cliente pueda entrar, no solo Nutrea.
 *
 * ┌─ EL BLOQUEO QUE ESTO EXISTE PARA QUE NO VUELVA ────────────┐
 * │                                                            │
 * │ Sara entró por primera vez a su cuenta y se quedó parada    │
 * │ en «Paso 3 de 3 · ¿Con cuál abrimos?» con la lista VACÍA.   │
 * │ No podía pasar de ahí.                                      │
 * │                                                            │
 * │ El paso 3 tenía DOS tarjetas escritas a mano en el HTML —   │
 * │ «Nutrea Ecuador» y «Nutrea GT»— y el código solo sabía      │
 * │ ESCONDER la que no correspondiera. Nunca creaba una.       │
 * │                                                            │
 * │ O sea: la puerta de Nova solo se abría para una cuenta      │
 * │ cuyas tiendas se llamaran exactamente `ec` y `gt`. La       │
 * │ nuestra. Todo cliente con otro id se quedaba afuera, y sin  │
 * │ ningún mensaje: desde el punto de vista del código no había │
 * │ fallado nada.                                               │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

async function puerta(b, tiendas, fichas) {
  const p = await b.newPage({ viewport: { width: 1280, height: 900 } });
  const errores = [];
  p.on('pageerror', e => errores.push(e.message));
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');
  await p.evaluate(({ tiendas, fichas }) => {
    // Sin red: lo que se mide es la pantalla de entrada.
    window.fetch = async () => ({ text: async () => JSON.stringify({ ok: true }) });
    (0, eval)('CONECTADO = true; TOKEN = "t";');
    document.getElementById('em').value = 'sara@quickfix.com';
    step(3);
    pintarTiendasPermitidas(tiendas, fichas);
  }, { tiendas, fichas });
  await p.waitForTimeout(250);
  return { p, errores };
}

const tarjetas = (p) => p.$$eval('#s3 .spick', els => els.map(e => ({
  txt: e.textContent.replace(/\s+/g, ' ').trim(),
  id: (e.getAttribute('onclick') || '').match(/setStore\('([^']*)'\)/)?.[1] || '',
})));

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  console.log('\n══ 1 · EL CASO DE SARA: UNA TIENDA QUE NO SE LLAMA «ec» ══');
  {
    const { p, errores } = await puerta(b, ['qfx'],
      [{ id: 'qfx', nombre: 'QuickFix', moneda: 'COP', pais: 'Colombia' }]);

    /**
     * Con una sola tienda no hay nada que elegir: se entra directo. Lo
     * que se comprueba es que ENTRÓ, que es lo que Sara no podía hacer.
     */
    ok('la pantalla de entrada se cerró', !(await p.isVisible('#login')));
    ok('y está dentro de la app',
       await p.evaluate(() => document.getElementById('app').classList.contains('on')));
    igual('con su tienda puesta', 'qfx', await p.evaluate(() => ST));
    igual('y con su nombre de verdad, no el nuestro', 'QuickFix',
          await p.evaluate(() => (STORES[ST] || {}).name));
    igual('y su moneda', 'COP', await p.evaluate(() => (STORES[ST] || {}).moneda));
    ok('nunca aparece Nutrea', !/Nutrea/.test(await p.textContent('#s3')),
       await p.textContent('#s3'));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 2 · VARIAS TIENDAS: UNA TARJETA POR CADA UNA ══');
  {
    const { p, errores } = await puerta(b, ['qfx', 'mx', 'pe'], [
      { id: 'qfx', nombre: 'QuickFix', moneda: 'COP', pais: 'Colombia' },
      { id: 'mx',  nombre: 'QuickFix México', moneda: 'MXN', pais: 'México' },
      { id: 'pe',  nombre: 'QuickFix Perú', moneda: 'PEN', pais: 'Perú' },
    ]);
    const t = await tarjetas(p);
    igual('tres tarjetas', 3, t.length);
    igual('con sus ids', ['qfx', 'mx', 'pe'], t.map(x => x.id));
    ok('con sus nombres', /QuickFix México/.test(t[1].txt), JSON.stringify(t[1]));
    ok('y su moneda y país', /MXN · México/.test(t[1].txt), JSON.stringify(t[1]));
    ok('sigue sin aparecer Nutrea', !/Nutrea/.test(await p.textContent('#s3')));
    ok('todavía no ha entrado: hay que elegir', await p.isVisible('#login'));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 3 · DIEZ TIENDAS, COMO LAS DE SARA ══');
  {
    const muchas = [];
    for (let i = 0; i < 10; i++) muchas.push({ id: 't' + i, nombre: 'Tienda ' + i,
                                               moneda: 'COP', pais: 'Colombia' });
    const { p, errores } = await puerta(b, muchas.map(x => x.id), muchas);
    igual('las diez salen', 10, (await tarjetas(p)).length);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 4 · UNA TIENDA SIN MONEDA LO DICE EN LA PUERTA ══');
  /**
   * Una tienda sin moneda no suma bien nada. Enterarse aquí es mucho
   * mejor que tres pantallas adentro, cuando ya hay pedidos cargados.
   */
  {
    const { p, errores } = await puerta(b, ['a', 'b'], [
      { id: 'a', nombre: 'Con moneda', moneda: 'COP', pais: 'Colombia' },
      { id: 'b', nombre: 'Sin moneda', moneda: '', pais: 'Colombia' },
    ]);
    const t = await tarjetas(p);
    ok('lo dice', /sin moneda/i.test(t[1].txt), JSON.stringify(t[1]));
    ok('y en rojo', await p.evaluate(() => {
      const s = document.querySelectorAll('#s3 .spick')[1].querySelector('small');
      return getComputedStyle(s).color === 'rgb(184, 50, 47)';
    }));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 5 · SIN NINGUNA TIENDA, NO SE OFRECEN LAS AJENAS ══');
  /**
   * Antes esto salía sin tocar nada y dejaba en pantalla las dos
   * tarjetas de ejemplo: dos tiendas de OTRA cuenta, clicables. Se
   * entraba a una tienda inexistente y todo daba error sin decir por qué.
   */
  {
    const { p, errores } = await puerta(b, [], []);
    igual('no hay ninguna tarjeta', 0, (await tarjetas(p)).length);
    const txt = (await p.textContent('#s3')).replace(/\s+/g, ' ');
    ok('NO se ofrece Nutrea', !/Nutrea/.test(txt), txt);
    ok('se dice qué pasa', /no tiene tiendas/i.test(txt), txt);
    ok('y qué hacer', /quien te creó la cuenta/i.test(txt), txt);
    ok('y no se entró a ninguna parte', await p.isVisible('#login'));
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
