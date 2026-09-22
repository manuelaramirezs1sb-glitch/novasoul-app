/**
 * El semáforo en la pantalla de la dueña.
 *
 * Tres situaciones, y la tercera es la que más importa: cuando el
 * servidor dice que no —plan sin dinero, o persona sin permiso— la
 * pantalla no puede quedar con un hueco ni con un cartel.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const SEM = {
  tienda: 'gt', moneda: 'GTQ', hayDatos: true,
  semana: { lunes: '2026-09-14', domingo: '2026-09-20' },
  umbrales: { ticket_minimo: 330, entrega_minima: 72, muestra_minima: 10, cpa_verde_pct: 70 },
  luces: [
    { estado: 'verde', valor: 66.67, etiqueta: 'CPA por entrega', delta: -4.2,
      porque: 'El techo es 116.92 GTQ. Queda 50.26 de margen.' },
    { estado: 'amarillo', valor: 305.56, etiqueta: 'Ticket promedio', delta: -14.6,
      porque: 'El mínimo que definiste es 330 GTQ.' },
    { estado: 'verde', valor: 81.82, etiqueta: 'Entrega sobre resuelto', delta: 9.3,
      porque: '9 de 11 resueltos · 1 sigue en tránsito, fuera de esta cuenta' },
    { estado: 'incompleto', valor: 920, etiqueta: 'Utilidad de la semana', delta: -322.7,
      porque: '1386315 de pauta quedaron sin convertir por falta de tasa.' },
  ],
  hoy: { pedidos: 13, entregados: 9, devoluciones: 2, cancelados: 1, enTransito: 1,
         ganancia: 1580, costoDevoluciones: 60, utilidadAntesPauta: 1520,
         hayPauta: true, gasto: 600, utilidadReal: 920, techo: 116.92 },
  alertas: [
    { nivel: 'rojo', titulo: '1386315 de pauta sin convertir',
      accion: 'Meta cobra en otra moneda y faltan las tasas de esos días.' },
    { nivel: 'amarillo', titulo: 'La semana tiene 6 de 7 días con registros',
      accion: 'Las cifras son parciales, no bajas. No se extrapoló nada.' },
  ],
  productos: [
    { producto: 'TAG RECEDE', pedidos: 12, entregados: 8, devoluciones: 2,
      tasaEntrega: 80, gananciaPorPedido: 117.5, señal: 'medible' },
    { producto: 'TRULY ROSA', pedidos: 1, entregados: 1, devoluciones: 0,
      tasaEntrega: 100, gananciaPorPedido: 170,
      señal: 'faltan 9 resueltos para que signifique algo' },
  ],
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };

  for (const ancho of [1100, 390]) {
    const p = await b.newPage({ viewport: { width: ancho, height: 1100 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    await p.goto('file:///home/claude/repo/empresarial.html');
    await p.waitForLoadState('load');
    await p.evaluate(() => {
      document.getElementById('login').style.display = 'none';
      (0, eval)('CONECTADO = true; ST = "gt";');
    });

    // 1 · Con datos
    await p.evaluate((s) => {
      window.api = async () => ({ ok: true, semaforo: s });
      return cargarSemaforo();
    }, SEM);
    const hoy = await p.textContent('#sem-luces');
    const pauta = await p.textContent('#sem-pauta');

    ok('las cuatro luces salen en Hoy (@' + ancho + ')',
       (await p.$$('#sem-luces .sem-luz')).length === 4);
    ok('Hoy NO repite la tabla ni las alertas (@' + ancho + ')',
       !hoy.includes('Techo de CPA') && !hoy.includes('Qué mirar'));
    ok('Pauta trae luces, números, alertas y producto (@' + ancho + ')',
       pauta.includes('Techo de CPA') && pauta.includes('Qué mirar') &&
       pauta.includes('TAG RECEDE'));
    ok('el producto sin muestra no muestra una ganancia que no significa nada (@' + ancho + ')',
       /TRULY ROSA[\s\S]{0,120}faltan 9 resueltos/.test(pauta) &&
       !pauta.includes('170'), 'no debería imprimir 170');
    ok('la utilidad incompleta dice por qué (@' + ancho + ')',
       hoy.includes('sin convertir'));
    ok('el delta se ve con su signo (@' + ancho + ')', hoy.includes('▼'));
    ok('el pie explica cómo se calculó (@' + ancho + ')',
       pauta.includes('Solo los entregados generan ganancia'));

    // 2 · Semana vacía
    await p.evaluate((s) => {
      const v = JSON.parse(JSON.stringify(s));
      v.hayDatos = false;
      window.api = async () => ({ ok: true, semaforo: v });
      return cargarSemaforo();
    }, SEM);
    const vacio = await p.textContent('#sem-luces');
    // La fecha lleva ceros, así que buscar «0» no dice nada. Lo que no
    // puede pasar es que dibuje las luces: cuatro ceros con pinta de
    // medición son exactamente la mentira que este caso evita.
    ok('la semana vacía no dibuja ninguna luz (@' + ancho + ')',
       (await p.$$('#sem-luces .sem-luz')).length === 0);
    ok('y explica la diferencia entre no vender y no importar (@' + ancho + ')',
       vacio.includes('no se actualizó') && vacio.includes('no significa que no se vendió'),
       vacio.slice(0, 100));

    // 3 · Sin permiso: nada, ni un cartel
    await p.evaluate(() => {
      window.api = async () => ({ ok: false, error: 'Tu plan no incluye…', porque: 'plan' });
      return cargarSemaforo();
    });
    ok('sin permiso no deja hueco ni cartel en Hoy (@' + ancho + ')',
       (await p.textContent('#sem-luces')).trim() === '');
    ok('ni en Pauta (@' + ancho + ')',
       (await p.textContent('#sem-pauta')).trim() === '');

    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral (@' + ancho + ')', lateral === 0, lateral + 'px');
    ok('sin errores de JS (@' + ancho + ')', errores.length === 0, errores.join(' | '));

    if (ancho === 1100) {
      await p.evaluate((s) => {
        window.api = async () => ({ ok: true, semaforo: s });
        return cargarSemaforo();
      }, SEM);
      await p.evaluate(() => {
        let el = document.getElementById('sem-pauta');
        while (el && el !== document.body) { el.style.display = 'block'; el = el.parentElement; }
      });
      await p.locator('#sem-pauta').screenshot({ path: OUT + 'semaforo-pantalla.png' });
    }
    await p.close();
  }
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
