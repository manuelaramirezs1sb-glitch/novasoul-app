/**
 * EL CHAT DEL EQUIPO, EN LA PANTALLA.
 *
 * ┌─ QUÉ SE ESTÁ COMPROBANDO ──────────────────────────────────┐
 * │                                                            │
 * │ «el dueño, la admin y las gestoras deben tener un mismo     │
 * │  chat en Nova (…) que pueda hablar con su equipo completo   │
 * │  o solo dejarle un mensaje al admin o gestor en privado,    │
 * │  dentro de Nova, no por fuera».                            │
 * │                                                            │
 * │ La pestaña «Nova Chat» era un BOTÓN que abría un grupo de   │
 * │ WhatsApp. No fallaba: no era un chat. Así que la primera    │
 * │ aserción de este archivo es que la caja de escribir esté    │
 * │ ahí y que lo escrito SALGA hacia el servidor — porque eso   │
 * │ es exactamente lo que no pasaba.                           │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y LA TRAMPA QUE YA CAÍ UNA VEZ ───────────────────────────┐
 * │                                                            │
 * │ Las clases del chat se llamaban `.eq-av`, `.eq-sec`,        │
 * │ `.eq-vacio`… y esos nombres YA existían para las tarjetas   │
 * │ de Equipo. El círculo del chat heredaba 34px y un           │
 * │ line-height de otra pantalla, y nada avisaba. Por eso aquí  │
 * │ se mide el tamaño real del círculo, no solo que exista.     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** Lo que devolvería el servidor de verdad, con la forma exacta de apiChat. */
function respuesta(quien, con, marcar) {
  const YO = {
    duena: { correo: 'm@nova.com', nombre: 'Manuela Ramírez', tiendas: ['ec', 'gt'] },
    gestora: { correo: 'andrea@nutrea.com', nombre: 'Andrea Ramírez', tiendas: ['ec'] },
  }[quien];

  const GENTE = [
    { correo: 'm@nova.com', nombre: 'Manuela Ramírez', rol: 'dueno', tiendas: '*' },
    { correo: 'kat@nutrea.com', nombre: 'Katherin Ortiz', rol: 'admin', tiendas: 'ec,gt' },
    { correo: 'andrea@nutrea.com', nombre: 'Andrea Ramírez', rol: 'gestora', tiendas: 'ec' },
  ].filter(g => g.correo !== YO.correo);

  const TIE = { ec: 'Nutrea Ecuador', gt: 'Nutrea Guatemala' };
  const MSG = {
    'g:ec': [
      { id: 'm1', de: 'm@nova.com', deNombre: 'Manuela Ramírez',
        texto: 'Hoy priorizamos las 12 sin confirmar de Ecuador.',
        cuando: '2026-09-25 08:12:00', borrado: false, mio: YO.correo === 'm@nova.com' },
      { id: 'm2', de: 'andrea@nutrea.com', deNombre: 'Andrea Ramírez',
        texto: 'Voy con esas. Dos no contestan desde ayer.',
        cuando: '2026-09-25 08:40:00', borrado: false, mio: YO.correo === 'andrea@nutrea.com' },
      { id: 'm3', de: 'kat@nutrea.com', deNombre: 'Katherin Ortiz',
        texto: '', cuando: '2026-09-25 08:41:00', borrado: true, mio: false },
    ],
    'kat@nutrea.com': [
      { id: 'm9', de: 'kat@nutrea.com', deNombre: 'Katherin Ortiz',
        texto: 'Andrea, te faltó la nota del pedido NE-4412.',
        cuando: '2026-09-25 09:05:00', borrado: false, mio: false },
    ],
  };

  const hilos = {};
  if (quien === 'gestora') {
    hilos['g:ec'] = { id: 'g:ec', sinLeer: 2,
      ultimo: { de: 'Katherin Ortiz', texto: '(mensaje borrado)',
                cuando: '2026-09-25 08:41:00', mio: false } };
    hilos['kat@nutrea.com'] = { id: 'kat@nutrea.com', sinLeer: 1,
      ultimo: { de: 'Katherin Ortiz', texto: 'Andrea, te faltó la nota del pedido NE-4412.',
                cuando: '2026-09-25 09:05:00', mio: false } };
  } else {
    hilos['g:ec'] = { id: 'g:ec', sinLeer: 1,
      ultimo: { de: 'Katherin Ortiz', texto: '(mensaje borrado)',
                cuando: '2026-09-25 08:41:00', mio: false } };
  }
  let sinLeerTotal = Object.keys(hilos).reduce((a, k) => a + hilos[k].sinLeer, 0);

  const abierta = con && (con.indexOf('g:') !== 0 || YO.tiendas.indexOf(con.slice(2)) !== -1)
    ? con : '';
  // Como el servidor de verdad: si se pide `marcar`, ese hilo queda en
  // cero y el total baja con él.
  if (abierta && marcar && hilos[abierta]) {
    sinLeerTotal -= hilos[abierta].sinLeer;
    hilos[abierta] = Object.assign({}, hilos[abierta], { sinLeer: 0 });
  }
  return {
    ok: true, yo: YO.correo, yoNombre: YO.nombre, gente: GENTE,
    grupos: YO.tiendas.map(t => ({ id: 'g:' + t, tienda: t, nombre: TIE[t] })),
    hilos: hilos, sinLeer: sinLeerTotal,
    con: abierta || (con ? '' : 'g:' + YO.tiendas[0]),
    mensajes: abierta ? (MSG[abierta] || []) : [],
  };
}

async function abrir(b, rol) {
  const p = await b.newPage({ viewport: { width: 1280, height: 1000 } });
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
  await p.exposeFunction('__resp', (quien, con, marcar) => respuesta(quien, con || '', marcar));

  await p.evaluate(async ({ rol }) => {
    document.getElementById('login').style.display = 'none';
    // #app es display:none hasta entrar; sin esto nada tiene tamaño y
    // medir estilos devuelve valores que no son los de la pantalla real.
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO = true; ST = "ec"; ROL = "' + rol + '";');
    document.documentElement.setAttribute('data-rol', rol);
    (0, eval)('SESION = ' + JSON.stringify({
      nombre: rol === 'gestora' ? 'Andrea Ramírez' : 'Manuela Ramírez', rol: rol,
      correo: rol === 'gestora' ? 'andrea@nutrea.com' : 'm@nova.com',
      tiendas: rol === 'gestora' ? ['ec'] : ['ec', 'gt'],
      permisos: [], modulos: ['empresarial'] }));

    const quien = rol === 'gestora' ? 'gestora' : 'duena';
    window.api = async (accion, params) => {
      window.__anotar(JSON.stringify({ accion: accion, params: params || {} }));
      if (accion === 'chat') {
        return await window.__resp(quien, (params || {}).con || '', !!(params || {}).marcar);
      }
      return { ok: true };
    };
  }, { rol });

  /**
   * Abrir el panel y pasarse a la pestaña del equipo, como lo haría una
   * persona. Hace falta de verdad: `.chat-pane` está en display:none
   * hasta que la pestaña se marca, y medir «¿se ve la conversación?» con
   * el panel cerrado da «no» siempre, tenga o no razón.
   */
  await p.evaluate(() => {
    toggleChat();
    document.querySelectorAll('.chat-tab')[1].click();
  });
  await p.waitForTimeout(200);

  return { p, enviados, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 1 · YA NO ES UN BOTÓN QUE MANDA AFUERA ══');
  {
    const { p, errores } = await abrir(b, 'gestora');
    const html = await p.content();
    ok('no queda el botón «Abrir el canal» en el panel de chat',
       !/nc-canal/.test(html), 'sigue el div del canal externo');
    ok('la pestaña se llama «Mi equipo»',
       /Mi equipo/.test(await p.textContent('.chat-tabs')),
       await p.textContent('.chat-tabs'));
    ok('hay una caja para escribir', await p.$('#nch-inp') !== null);
    ok('y un botón de enviar', await p.$('.nch-send') !== null);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 2 · LA LISTA DE CONVERSACIONES ══');
  {
    const { p, errores } = await abrir(b, 'gestora');
    await p.evaluate(() => cargarChat(''));
    await p.waitForTimeout(200);

    const lista = (await p.textContent('#nch-lista')).replace(/\s+/g, ' ');
    ok('aparece el grupo de su tienda, con nombre de tienda',
       /Equipo · Nutrea Ecuador/.test(lista), lista.slice(0, 260));
    ok('NO aparece el grupo de la tienda que no tiene',
       !/Guatemala/.test(lista), lista.slice(0, 300));
    ok('aparece la admin para escribirle en privado',
       /Katherin Ortiz/.test(lista), lista.slice(0, 300));
    ok('con el rol en español, no «dueno»',
       /Dueña/.test(lista) && !/dueno/.test(lista), lista.slice(0, 300));
    // Ojo: «Andrea» sí aparece — dentro del mensaje que le escribió la
    // admin. Lo que no debe haber es una FILA suya. Buscar la palabra en
    // todo el bloque medía otra cosa y fallaba teniendo razón la pantalla.
    const nombres = await p.$$eval('#nch-lista .nch-nom', e => e.map(x => x.textContent));
    ok('no se ofrece hablar consigo misma',
       nombres.indexOf('Andrea Ramírez') === -1, JSON.stringify(nombres));

    igual('tres sin leer en el punto rojo', '3',
          await p.textContent('#chat-fab-n'));
    ok('el punto rojo se ve', await p.evaluate(
       () => document.getElementById('chat-fab-n').classList.contains('on')));

    const ns = await p.$$eval('.nch-n', els => els.map(e => e.textContent));
    igual('cada conversación con su propio conteo', ['2', '1'], ns);

    console.log('\n── La colisión de clases que ya me pasó ──');
    const av = await p.$$eval('#nch-lista .nch-av',
      els => els.slice(0, 1).map(e => {
        const c = getComputedStyle(e);
        return { w: c.width, h: c.height, lh: c.lineHeight };
      })[0]);
    igual('el círculo mide lo del chat, no lo de las tarjetas de Equipo',
          { w: '32px', h: '32px' }, { w: av.w, h: av.h });

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 3 · ABRIR UNA CONVERSACIÓN Y LEERLA ══');
  {
    const { p, enviados, errores } = await abrir(b, 'gestora');
    await p.evaluate(() => cargarChat(''));
    await p.waitForTimeout(150);
    await p.evaluate(() => nchAbrir('g:ec'));
    await p.waitForTimeout(250);

    ok('la lista se esconde', !(await p.isVisible('#nch-lista')));
    ok('y se ve la conversación', await p.isVisible('#nch-hilo'));
    igual('con el nombre del grupo arriba', 'Equipo · Nutrea Ecuador',
          (await p.textContent('#nch-htitulo')).trim());
    ok('y a quién lo lee', /todo el equipo/i.test(await p.textContent('#nch-hsub')),
       await p.textContent('#nch-hsub'));

    const ms = await p.$$eval('.nch-m', els => els.map(e => ({
      txt: e.querySelector('.nch-m-txt').textContent,
      de: (e.querySelector('.nch-m-de') || {}).textContent || '',
      mio: e.classList.contains('mio'), ido: e.classList.contains('ido'),
      borrar: !!e.querySelector('.nch-m-x'),
    })));
    igual('llegan los tres mensajes', 3, ms.length);
    igual('con el nombre de quien escribió', 'Manuela Ramírez', ms[0].de);
    ok('el suyo va a la derecha y sin nombre', ms[1].mio && ms[1].de === '',
       JSON.stringify(ms[1]));
    ok('solo el suyo se puede borrar',
       ms[1].borrar && !ms[0].borrar && !ms[2].borrar, JSON.stringify(ms.map(m => m.borrar)));
    ok('el borrado dice que se borró, no desaparece',
       ms[2].ido && /borrado/.test(ms[2].txt), JSON.stringify(ms[2]));

    console.log('\n── Abrir es leer, y en un solo viaje ──');
    /**
     * Eran DOS peticiones seguidas para un solo gesto: `chat` para traer
     * y `chat_visto` para marcar, y la segunda releía la hoja entera que
     * la primera acababa de leer. Ahora abrir pide una vez, con
     * `marcar`, y el servidor hace las dos cosas.
     */
    const abre = enviados.filter(x => x.accion === 'chat' && x.params.con === 'g:ec');
    igual('abrir la conversación es UNA petición', 1, abre.length);
    igual('y pide que se marque en el mismo viaje', true, abre[0] && abre[0].params.marcar);
    igual('ya no hay una segunda petición para marcar', 0,
          enviados.filter(x => x.accion === 'chat_visto').length);

    console.log('\n── Y se puede volver ──');
    await p.click('.nch-atras');
    await p.waitForTimeout(120);
    ok('vuelve a la lista', await p.isVisible('#nch-lista'));
    ok('y la conversación se cierra', !(await p.isVisible('#nch-hilo')));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 4 · ESCRIBIR, QUE ERA LO QUE NO PASABA ══');
  {
    const { p, enviados, errores } = await abrir(b, 'gestora');
    await p.evaluate(() => cargarChat(''));
    await p.waitForTimeout(150);
    await p.evaluate(() => nchAbrir('kat@nutrea.com'));
    await p.waitForTimeout(250);

    igual('el privado se titula con el nombre de la persona', 'Katherin Ortiz',
          (await p.textContent('#nch-htitulo')).trim());
    ok('y dice que es solo entre las dos',
       /solo ustedes dos/i.test(await p.textContent('#nch-hsub')),
       await p.textContent('#nch-hsub'));

    await p.fill('#nch-inp', 'Ya quedó la nota, Katherin.');
    await p.keyboard.press('Enter');
    await p.waitForTimeout(300);

    const env = enviados.filter(x => x.accion === 'chat_enviar');
    igual('Enter manda el mensaje', 1, env.length);
    igual('con el texto', 'Ya quedó la nota, Katherin.', env[0] && env[0].params.texto);
    igual('a la conversación abierta', 'kat@nutrea.com', env[0] && env[0].params.con);
    igual('la caja queda vacía', '', await p.inputValue('#nch-inp'));

    console.log('\n── Shift+Enter no manda: hace párrafo ──');
    await p.fill('#nch-inp', 'primera línea');
    await p.keyboard.down('Shift'); await p.keyboard.press('Enter'); await p.keyboard.up('Shift');
    await p.waitForTimeout(150);
    igual('sigue habiendo un solo envío', 1,
          enviados.filter(x => x.accion === 'chat_enviar').length);
    ok('y lo escrito no se perdió',
       /primera línea/.test(await p.inputValue('#nch-inp')),
       await p.inputValue('#nch-inp'));

    console.log('\n── Un mensaje en blanco no viaja ──');
    await p.fill('#nch-inp', '   ');
    await p.keyboard.press('Enter');
    await p.waitForTimeout(150);
    igual('no se manda nada', 1, enviados.filter(x => x.accion === 'chat_enviar').length);

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 5 · SI FALLA EL ENVÍO, NO SE PIERDE LO ESCRITO ══');
  {
    const { p, errores } = await abrir(b, 'gestora');
    await p.evaluate(() => {
      const real = window.api;
      window.api = async (a, prm) => {
        if (a === 'chat_enviar') return { ok: false, error: 'Sin conexión: no se envió' };
        return real(a, prm);
      };
    });
    await p.evaluate(() => cargarChat(''));
    await p.waitForTimeout(150);
    await p.evaluate(() => nchAbrir('g:ec'));
    await p.waitForTimeout(250);

    const antes = (await p.$$('.nch-m')).length;
    await p.fill('#nch-inp', 'esto no va a salir');
    await p.keyboard.press('Enter');
    await p.waitForTimeout(350);

    igual('lo escrito vuelve a la caja', 'esto no va a salir',
          await p.inputValue('#nch-inp'));
    igual('y no queda un mensaje fantasma en pantalla', antes,
          (await p.$$('.nch-m')).length);
    ok('avisa que no se envió',
       /no se envió/i.test(await p.textContent('#toast')),
       await p.textContent('#toast'));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 6 · LA DUEÑA CON DOS TIENDAS ══');
  {
    const { p, errores } = await abrir(b, 'dueno');
    await p.evaluate(() => cargarChat(''));
    await p.waitForTimeout(200);

    const lista = (await p.textContent('#nch-lista')).replace(/\s+/g, ' ');
    ok('tiene los dos grupos',
       /Nutrea Ecuador/.test(lista) && /Nutrea Guatemala/.test(lista), lista.slice(0, 320));
    ok('y dice que son por tienda', /GRUPOS POR TIENDA/.test(lista), lista.slice(0, 120));
    ok('con su equipo para los privados',
       /Katherin/.test(lista) && /Andrea/.test(lista), lista.slice(0, 400));

    console.log('\n── Pedir un grupo que no es suyo no lo abre ──');
    await p.evaluate(() => nchAbrir('g:xx'));
    await p.waitForTimeout(250);
    ok('vuelve a la lista', await p.isVisible('#nch-lista'));
    ok('y no queda una conversación abierta a medias',
       !(await p.isVisible('#nch-hilo')));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 7 · SOLA EN NOVA, SIN EQUIPO TODAVÍA ══');
  {
    const { p, errores } = await abrir(b, 'dueno');
    await p.evaluate(() => {
      window.api = async (a) => (a === 'chat'
        ? { ok: true, yo: 'm@nova.com', yoNombre: 'Manuela', gente: [], grupos: [],
            hilos: {}, sinLeer: 0, con: '', mensajes: [] }
        : { ok: true });
      return cargarChat('');
    });
    await p.waitForTimeout(200);

    const lista = (await p.textContent('#nch-lista')).replace(/\s+/g, ' ');
    ok('no finge conversaciones', !/Equipo ·/.test(lista), lista.slice(0, 200));
    ok('dice qué falta para usarlo', /equipo/i.test(lista), lista.slice(0, 220));
    ok('y el punto rojo no se ve', !(await p.evaluate(
       () => document.getElementById('chat-fab-n').classList.contains('on'))));

    igual('sin errores de consola', [], errores);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
