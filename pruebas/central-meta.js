/**
 * Meta, visto desde la consola.
 *
 * Lo que había antes en esta pantalla era una maqueta: un cuadro para
 * subir un CSV y cuatro números escritos a mano en el HTML —«Alcance
 * 184K», «CPA $9.80», «ROAS 3.2×»— que no salían de ningún lado y no
 * cambiaban nunca. Una pantalla vacía se nota; un número inventado se
 * cree. La primera prueba de este archivo es que esos números ya no
 * estén, porque si vuelven, vuelven en silencio.
 *
 * Lo demás que se comprueba:
 *
 *   · que haya TRES estados y no dos — «con llave pero sin número de
 *     cuenta» es el que se pierde si solo hay sí y no, y es justo el
 *     que deja a alguien creyendo que el gasto está entrando;
 *   · que la consola NO tenga dónde escribir una llave, que es una
 *     pantalla que no se debe construir;
 *   · que de Nova y novAcademy diga «todavía no pauta» en vez de
 *     dibujar campañas en cero.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const fs = require('fs');

const DATOS = {
  ok: true, hoy: '2026-09-23', error: '',
  clientes: [
    { id: 'c1', empresa: 'Nutrea', estado: 'activo', tienePauta: true,
      hayLlave: true, guardadaEn: '2026-08-01', sinCuenta: 0, listo: true,
      ultimaPrueba: { cuando: '2026-09-22', nombre: 'Nutrea Ads' },
      tiendas: [{ id: 'gt', nombre: 'Nutrea GT', cuenta: '11122233', moneda: 'GTQ' },
                { id: 'ec', nombre: 'Nutrea EC', cuenta: '44455566', moneda: 'USD' }] },
    { id: 'c2', empresa: 'Tienda Bogotá', estado: 'activo', tienePauta: true,
      hayLlave: true, guardadaEn: '2026-09-10', sinCuenta: 1, listo: false,
      ultimaPrueba: null,
      tiendas: [{ id: 'bog', nombre: 'Tienda Bogotá', cuenta: '', moneda: 'COP' }] },
    { id: 'c3', empresa: 'Tienda Lima', estado: 'activo', tienePauta: true,
      hayLlave: false, guardadaEn: '', sinCuenta: 1, listo: false, ultimaPrueba: null,
      tiendas: [{ id: 'lim', nombre: 'Tienda Lima', cuenta: '', moneda: 'PEN' }] },
    { id: 'c4', empresa: 'Sin plan de pauta', estado: 'activo', tienePauta: false,
      hayLlave: false, guardadaEn: '', sinCuenta: 0, listo: false, ultimaPrueba: null,
      tiendas: [] },
  ],
  resumen: { total: 4, conPlanDePauta: 3, listos: 1, aMedias: 1, sinLlave: 1 },
  propias: [
    { id: 'nova', nombre: 'Nova', que: 'La plataforma que le vende a las tiendas.', pauta: false },
    { id: 'academy', nombre: 'novAcademy', que: 'Los cursos y el acompañamiento.', pauta: false },
  ],
};

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

(async () => {
  /**
   * ── LO PRIMERO: QUE LOS NÚMEROS INVENTADOS NO ESTÉN ──
   * Se mira el archivo, no la pantalla. Un número puede estar escondido
   * en un div oculto y seguir estando listo para aparecer.
   */
  console.log('\n── Los números inventados ──');
  const html = fs.readFileSync(__dirname + '/../novacentral.html', 'utf8');
  [['184K', 'el alcance'], ['3.2×', 'el ROAS'], ['9.80', 'el CPA'],
   ['Q 21,180', 'la inversión de GT'], ['$ 4,280', 'la inversión de EC'],
  ].forEach(function (par) {
    ok(par[1] + ' ya no está en el archivo', html.indexOf(par[0]) === -1);
  });
  ok('y tampoco el cuadro de subir el CSV', !/onMetaFileC/.test(html));

  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  for (const ancho of [1100, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1100 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    p.on('console', m => {
      if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) {
        errores.push('console: ' + m.text());
      }
    });

    await p.goto('file:///home/claude/repo/novacentral.html');
    await p.waitForLoadState('load');
    await p.evaluate((d) => {
      window.nc = async (accion) => {
        if (accion === 'nc_meta') return d;
        return { ok: true };
      };
      cargarMetaC();
    }, DATOS);
    await p.waitForTimeout(150);
    await p.evaluate(() => go('meta', document.querySelector('[data-v=meta]')));
    await p.waitForTimeout(80);

    const t = await p.textContent('#v-meta');

    if (ancho === 1100) {
      console.log('\n── Los tres estados ──');
      ok('la conectada de verdad' + A, /Nutrea[\s\S]*?Conectada/.test(t), t);
      ok('la que tiene llave pero le falta el número de cuenta' + A,
         /Falta la cuenta/.test(t), t);
      ok('y la que no ha conectado nada' + A, /Sin conectar/.test(t), t);
      ok('se dice cuántas de cuántas, arriba' + A,
         /1 de 3 conectadas/.test(await p.textContent('#meta-lbl')));
      ok('la que no tiene cuenta lo dice en su renglón' + A, /sin cuenta/.test(t), t);

      console.log('\n── Los nombres de tienda ──');
      /** GLISEN y DAFERASE son hojas de gastos, nunca nombres de tienda. */
      ok('se usan los de pantalla' + A, /Nutrea GT/.test(t) && /Nutrea EC/.test(t), t);
      ok('y nunca los de las hojas de gastos' + A,
         !/GLISEN/i.test(t) && !/DAFERASE/i.test(t), t);

      console.log('\n── Lo que la consola NO hace ──');
      /**
       * Ni un campo de texto, ni un botón de guardar. Si algún día
       * aparecen aquí, esta prueba es la que tiene que doler.
       */
      const campos = await p.$$eval('#v-meta input, #v-meta textarea', e => e.length);
      ok('no hay ni un campo donde escribir una llave' + A, campos === 0, campos + ' campos');
      ok('y se dice dónde se conecta de verdad' + A,
         /se escribe en la Nova Empresarial de cada cliente/.test(t), t);
      ok('y que desde aquí nunca se ve cuál es' + A,
         /nunca cuál es/.test(t), t);

      console.log('\n── Nova y novAcademy ──');
      ok('dicen que todavía no pautan' + A,
         (t.match(/Todavía no pauta/g) || []).length === 2, t);
      ok('y se explica por qué no hay un panel en cero' + A,
         /se lee como una mala semana/.test(t), t);

      console.log('\n── Los que no tienen pauta en el plan ──');
      ok('van aparte, no mezclados con los conectables' + A,
         /Sin pauta en el plan[\s\S]*Sin plan de pauta/.test(t), t);

      console.log('\n── Cuando algo falla ──');
      await p.evaluate(() => {
        window.nc = async () => ({ ok: false, error: 'Acción desconocida: nc_meta' });
        return cargarMetaC();
      });
      await p.waitForTimeout(120);
      const err = await p.textContent('#v-meta');
      ok('«acción desconocida» dice que falta publicar el script' + A,
         /NOVA-COMPLETO\.gs/.test(err), err);
      await p.evaluate(() => {
        window.nc = async () => { throw new Error('sin internet'); };
        return cargarMetaC();
      });
      await p.waitForTimeout(120);
      ok('y un fallo de red NO se disfraza de otra cosa' + A,
         /sin internet/.test(await p.textContent('#v-meta')));
    }

    if (ancho === 390) {
      const sc = await p.evaluate(() =>
        document.documentElement.scrollWidth - document.documentElement.clientWidth);
      ok('sin scroll lateral' + A, sc <= 0, sc + 'px');
    }

    ok('sin errores de JS' + A, errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
