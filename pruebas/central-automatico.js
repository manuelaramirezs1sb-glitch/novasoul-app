const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const ESCENAS = {
  apagado: {
    trabajos: [
      { nombre: 'Tasas de cambio', hora: 6, prendido: false,
        porque: 'Sin esto, la pauta que Meta cobra en otra moneda no se puede sumar.' },
      { nombre: 'Revisión de alarmas', hora: 7, prendido: false,
        porque: 'Sin esto, las alarmas solo se calculan cuando alguien abre Nova.' },
    ],
    todoPrendido: false, enFalta: 1, error: '',
    clientes: [
      { empresa: 'Nutrea', necesita: true, alDia: false, error: '',
        pares: [{ par: 'USD→COP', ultima: '', dias: -1, bien: false },
                { par: 'GTQ→COP', ultima: '', dias: -1, bien: false }] },
      { empresa: 'Tienda CO', necesita: false, alDia: true, pares: [], error: '' },
    ],
  },
  aMedias: {
    trabajos: [
      { nombre: 'Tasas de cambio', hora: 6, prendido: true, porque: '' },
      { nombre: 'Revisión de alarmas', hora: 7, prendido: true, porque: '' },
    ],
    todoPrendido: true, enFalta: 1, error: '',
    clientes: [
      { empresa: 'Nutrea', necesita: true, alDia: false, error: '',
        pares: [{ par: 'USD→COP', ultima: '2026-09-22', dias: 0, bien: true },
                { par: 'GTQ→COP', ultima: '', dias: -1, bien: false }] },
    ],
  },
  bien: {
    trabajos: [
      { nombre: 'Tasas de cambio', hora: 6, prendido: true, porque: '' },
      { nombre: 'Revisión de alarmas', hora: 7, prendido: true, porque: '' },
    ],
    todoPrendido: true, enFalta: 0, error: '',
    clientes: [
      { empresa: 'Nutrea', necesita: true, alDia: true, error: '',
        pares: [{ par: 'USD→COP', ultima: '2026-09-22', dias: 0, bien: true },
                { par: 'GTQ→COP', ultima: '2026-09-22', dias: 0, bien: true }] },
      { empresa: 'Tienda CO', necesita: false, alDia: true, pares: [], error: '' },
    ],
  },
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;

  for (const [nombre, datos] of Object.entries(ESCENAS)) {
    for (const ancho of [1100, 390]) {
      const p = await b.newPage({ viewport: { width: ancho, height: 1000 } });
      const errores = [];
      p.on('pageerror', e => errores.push(e.message));
      p.on('console', m => { if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) errores.push('console: ' + m.text()); });

      await p.goto('file:///home/claude/repo/novacentral.html');
      await p.waitForLoadState('load');

      // El servidor, simulado: la tarjeta se prueba sin sesión real.
      await p.evaluate((d) => {
        window.__prendido = false;
        window.nc = async (accion) => {
          if (accion === 'nc_automatico') return { ok: true, automatico: d };
          if (accion === 'nc_automatico_prender') {
            window.__prendido = true;
            const todo = JSON.parse(JSON.stringify(d));
            todo.trabajos.forEach(t => { t.prendido = true; });
            todo.todoPrendido = true;
            return { ok: true, automatico: todo, cargando: true,
                     mensaje: 'Listo. Las tasas empiezan a bajar en un minuto.' };
          }
          return { ok: true };
        };
        document.getElementById('login-screen').style.display = 'none';
        document.getElementById('app').style.display = 'flex';
        document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
        document.getElementById('v-clientes').classList.add('on');
      }, datos);

      await p.evaluate(() => cargarAutomatico());
      await p.waitForFunction(() =>
        !document.getElementById('auto-cuerpo').textContent.includes('Revisando'),
        null, { timeout: 5000 });

      const txt = await p.textContent('#auto-cuerpo');
      const hayBoton = await p.$('#auto-btn') !== null;
      const lateral = await p.evaluate(() =>
        document.documentElement.scrollWidth - document.documentElement.clientWidth);

      const esperaBoton = nombre !== 'bien';
      let mal = [];
      if (hayBoton !== esperaBoton) mal.push('botón=' + hayBoton + ' esperaba ' + esperaBoton);
      if (lateral > 0) mal.push('scroll lateral ' + lateral + 'px');
      if (errores.length) mal.push('JS: ' + errores.join(' | '));
      // Una cuenta que no necesita tasas no debe aparecer en la lista de tasas
      if (txt.includes('Tienda CO')) mal.push('nombra a quien no necesita tasas');

      if (mal.length) { fallas++; console.log('  FALLA ' + nombre + ' @' + ancho + 'px: ' + mal.join('; ')); }
      else console.log('  ok    ' + nombre + ' @' + ancho + 'px');

      if (ancho === 1100 && esperaBoton) {
        await p.click('#auto-btn');
        await p.waitForFunction(() => window.__prendido === true, null, { timeout: 5000 });
        await p.waitForTimeout(200);
        const despues = await p.textContent('#auto-cuerpo');
        if (despues.includes('Apagado')) { fallas++; console.log('  FALLA ' + nombre + ': sigue diciendo Apagado tras prender'); }
        else console.log('  ok    ' + nombre + ' · el botón repinta el estado');
      }

      if (ancho === 1100) {
        await p.locator('#auto-card').screenshot({ path: OUT + 'auto-' + nombre + '.png' });
      }
      await p.close();
    }
  }
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
