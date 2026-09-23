/**
 * Pedidos y Novedades: tabla a la izquierda, detalle a la derecha.
 *
 * Es la estructura que ella pidió con una referencia (tabla de órdenes
 * con panel lateral). Se toma la ESTRUCTURA; los colores siguen siendo
 * los de Nova.
 *
 * ── LOS DOS FALLOS QUE ESTA PRUEBA GUARDA ──
 *
 * 1· La etiqueta de estado miraba solo `estado_canonico`. Ella elegía un
 *    estado en el detalle del pedido, le daba Guardar, salía «Guardado
 *    en la hoja»… y la etiqueta seguía igual, el pedido no cambiaba de
 *    banda y la barra de arriba no se movía. Se guardaba de verdad: la
 *    que no se enteraba era la pantalla.
 *
 * 2· Guardar no repintaba nada. Aunque la etiqueta hubiera mirado bien,
 *    seguiría viéndose el estado viejo hasta recargar.
 *
 * Y una tercera cosa, de método: al armar esta pantalla dejé un `let`
 * repetido que tumbaba TODO el JavaScript del archivo. Ninguna prueba
 * dibujaba Pedidos, así que nadie se enteró hasta que reventó otra
 * pantalla. Esta prueba existe también para eso.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

const PEDIDOS = [
  { id: 'dropi-1', id_externo: '90001', cliente: 'Lucía Marín', producto: 'Tag Recede',
    ciudad: 'Quito', direccion: 'Av. Amazonas 120', valor: 32.9, fecha: '2026-09-20',
    estado: 'RECHAZADO', estado_canonico: 'sin_clasificar', estado_nova: '',
    telefono: '0991112233', transportadora: 'Servientrega', guia: 'SE-778',
    gestora_asignada: 'Katherin', nota: '' },
  { id: 'dropi-2', id_externo: '90002', cliente: 'Andrés Poveda', producto: 'Truly Rosa',
    ciudad: 'Guayaquil', direccion: 'Cdla. Kennedy m5', valor: 41, fecha: '2026-09-19',
    estado: 'NOVEDAD', estado_canonico: 'novedad', estado_nova: '',
    telefono: '0987654321', transportadora: 'Laar', guia: 'LA-221',
    gestora_asignada: '', nota: 'No contesta' },
  /**
   * El que tiene corrección manual. La plataforma dice que está en
   * tránsito; ella dijo que se entregó. Manda ella.
   */
  { id: 'dropi-3', id_externo: '90003', cliente: 'Sofía Benítez', producto: 'Tag Recede',
    ciudad: 'Cuenca', direccion: 'Calle Larga 4-20', valor: 32.9, fecha: '2026-09-18',
    estado: 'EN TRANSITO', estado_canonico: 'en_transito', estado_nova: 'entregado',
    telefono: '0966554433', transportadora: 'Servientrega', guia: 'SE-901',
    gestora_asignada: 'Katherin', nota: '' },
];

const NOVEDADES = [
  { id: 'nv-1', pedido_id: '90002', fuente: 'dropi', fecha: '2026-09-19',
    motivo: 'Cliente no contesta el teléfono', grupo: 'contacto', estado: 'abierta',
    desenlace: '', solucion: '', nota: 'Llamé dos veces', gestora_asignada: 'Katherin' },
  { id: 'nv-2', pedido_id: '90007', fuente: 'dropi', fecha: '2026-09-22',
    motivo: 'Dirección incompleta', grupo: 'direccion', estado: 'resuelta',
    desenlace: 'devolucion', solucion: 'Reprogramar con dirección nueva',
    nota: '', gestora_asignada: '' },
];

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  for (const ancho of [1200, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1200 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    p.on('console', m => {
      if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) {
        errores.push('console: ' + m.text());
      }
    });

    await p.goto('file:///home/claude/repo/empresarial.html');
    await p.waitForLoadState('load');
    await p.evaluate(() => {
      document.getElementById('login').style.display = 'none';
      (0, eval)('CONECTADO = true; ST = "ec";');
    });

    // El servidor, simulado.
    const guardados = [];
    await p.exposeFunction('__anotar', (x) => { guardados.push(x); });
    await p.evaluate(({ peds, novs }) => {
      window.__peds = peds;
      window.api = async (accion, params) => {
        if (accion === 'listar' && params.entidad === 'Pedidos') {
          return { ok: true, filas: JSON.parse(JSON.stringify(window.__peds)), total: peds.length };
        }
        if (accion === 'listar' && params.entidad === 'Novedades') {
          return { ok: true, filas: JSON.parse(JSON.stringify(novs)), total: novs.length };
        }
        if (accion === 'escribir') {
          window.__anotar(JSON.stringify(params));
          // Como haría la hoja: el cambio queda guardado de verdad.
          const f = window.__peds.filter(x => x.id === params.id)[0];
          if (f) Object.assign(f, params.campos);
          return { ok: true };
        }
        return { ok: true, filas: [], total: 0 };
      };
      return cargarPedidosReales();
    }, { peds: PEDIDOS, novs: NOVEDADES });
    await p.waitForTimeout(120);
    await p.evaluate(() => go('pedidos', document.querySelector('[data-v=pedidos]')));
    /**
     * Al entrar, el filtro es «Necesitan acción» — así debe ser, es lo
     * que hay que resolver hoy. Para mirar la tabla completa se pasa a
     * Todos, que es lo que haría ella.
     */
    await p.evaluate(() => filtrarPedidos('todos'));
    await p.waitForTimeout(80);

    console.log('\n── La tabla' + A + ' ──');
    ok('hay una fila por pedido' + A,
       (await p.$$('#pedlist .trow')).length === 3);
    const ped = await p.textContent('#v-pedidos');
    ok('con sus columnas' + A,
       /Cliente/.test(ped) && /Producto/.test(ped) && /Destino/.test(ped) &&
       /Valor/.test(ped) && /Estado/.test(ped), ped.slice(0, 300));
    ok('y el cliente y su número de pedido' + A,
       /Lucía Marín/.test(ped) && /#90001/.test(ped));

    /**
     * ── LA QUE IMPORTA ──
     * Sofía viene «EN TRANSITO» de la plataforma, pero ella la marcó
     * entregada. La tabla tiene que decir Entregado.
     */
    console.log('\n── Lo que ella dijo manda' + A + ' ──');
    const filaSofia = await p.evaluate(() => {
      const r = [...document.querySelectorAll('#pedlist .trow')]
        .filter(x => x.textContent.includes('Sofía'))[0];
      return r ? r.textContent : '';
    });
    ok('el pedido corregido a mano se ve Entregado, no En tránsito' + A,
       /Entregado/.test(filaSofia) && !/En tránsito/.test(filaSofia), filaSofia);
    /**
     * Pero el texto de la plataforma NO se pierde: se guarda para el
     * panel. Que ella lo corrija no borra lo que dijo la transportadora,
     * que es lo que hay que mirar cuando las dos versiones no cuadran.
     */
    await p.evaluate(() => {
      [...document.querySelectorAll('#pedlist .trow')]
        .filter(x => x.textContent.includes('Sofía'))[0].click();
    });
    await p.waitForTimeout(80);
    const panelSofia = await p.textContent('#pedpanel');
    ok('y el estado de la plataforma sigue a la vista en el detalle' + A,
       /EN TRANSITO/.test(panelSofia), panelSofia.slice(0, 400));
    // Y se vuelve a cerrar, para seguir probando desde cero.
    await p.evaluate(() => {
      [...document.querySelectorAll('#pedlist .trow')]
        .filter(x => x.textContent.includes('Sofía'))[0].click();
    });
    await p.waitForTimeout(60);

    console.log('\n── El panel' + A + ' ──');
    ok('antes de tocar nada, invita a tocar' + A,
       /Toca un pedido/.test(await p.textContent('#pedpanel')));
    await p.evaluate(() => {
      const r = [...document.querySelectorAll('#pedlist .trow')]
        .filter(x => x.textContent.includes('Lucía'))[0];
      r.click();
    });
    await p.waitForTimeout(80);
    let panel = await p.textContent('#pedpanel');
    ok('al tocar una fila sale su detalle' + A,
       /Lucía Marín/.test(panel) && /0991112233/.test(panel) &&
       /Av\. Amazonas 120/.test(panel), panel.slice(0, 400));
    ok('con la guía y la transportadora' + A,
       /Servientrega/.test(panel) && /SE-778/.test(panel));
    ok('y el texto crudo que mandó la plataforma' + A, /RECHAZADO/.test(panel));
    ok('la fila queda marcada' + A,
       (await p.$$('#pedlist .trow.on')).length === 1);
    /**
     * Cambiar el estado de UN pedido no es lo mismo que cambiar qué
     * significa esa palabra para todos. Si no se dice, se confunden.
     */
    ok('dice que vale solo para este pedido, y dónde se cambia para todos' + A,
       /Vale solo para este pedido/.test(panel) && /Ajustes → Estados/.test(panel));

    console.log('\n── Guardar un estado repinta' + A + ' ──');
    await p.evaluate(() => {
      [...document.querySelectorAll('#pedpanel .nstates .nst')]
        .filter(x => x.textContent.trim() === 'Devolución')[0].click();
      return guardarPedido();
    });
    await p.waitForTimeout(150);
    ok('se mandó a la hoja' + A,
       guardados.some(g => /"estado_nova":"devolucion"/.test(g)),
       JSON.stringify(guardados));
    const filaLucia = await p.evaluate(() => {
      const r = [...document.querySelectorAll('#pedlist .trow')]
        .filter(x => x.textContent.includes('Lucía'))[0];
      return r ? r.textContent : '';
    });
    ok('y la TABLA ya lo dice, sin recargar' + A,
       /Devolución/.test(filaLucia), filaLucia);
    ok('el pedido se movió a la banda de cerrados' + A,
       /CERRADOS/.test(await p.textContent('#pedlist')));

    if (ancho === 1200) {
      console.log('\n── Novedades ──');
      await p.evaluate(() => cargarNovedadesReales());
      await p.waitForTimeout(120);
      await p.evaluate(() => go('novedades', document.querySelector('[data-v=novedades]')));
      await p.waitForTimeout(80);
      ok('también es tabla' + A, (await p.$$('#novlist .trow')).length === 2);
      const nov = await p.textContent('#v-novedades');
      ok('con el motivo y el grupo' + A,
         /Cliente no contesta/.test(nov) && /contacto/.test(nov));
      ok('y los días que lleva abierta' + A, /\d+ d/.test(nov));
      ok('el panel invita a tocar' + A,
         /Toca una novedad/.test(await p.textContent('#novpanel')));
      await p.evaluate(() => {
        document.querySelectorAll('#novlist .trow')[1].click();
      });
      await p.waitForTimeout(80);
      const np = await p.textContent('#novpanel');
      ok('al tocarla, sale qué se le pidió al courier' + A,
         /Reprogramar con dirección nueva/.test(np), np.slice(0, 300));
      ok('y que terminó en devolución pese a estar resuelta' + A,
         /Terminó devolución/.test(np), np);
    }

    if (ancho === 390) {
      console.log('\n── En el teléfono ──');
      const sc = await p.evaluate(() =>
        document.documentElement.scrollWidth - document.documentElement.clientWidth);
      ok('sin scroll lateral' + A, sc <= 0, sc + 'px');
      // Una sola columna: el panel va debajo, no al lado.
      const dos = await p.evaluate(() => {
        const t = document.querySelector('#v-pedidos .tsplit');
        return getComputedStyle(t).gridTemplateColumns.split(' ').length;
      });
      ok('la tabla y el panel se apilan, no se aprietan' + A, dos === 1, dos + ' columnas');
    }

    ok('sin errores de JS' + A, errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
