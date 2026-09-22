/**
 * «Mi trabajo» y «Mi plata» en Nova Central.
 *
 * Lo que se comprueba no es que pinte: es que no mienta. Que no sume
 * monedas distintas, que separe lo atrasado de lo que aún no vence, que
 * no dibuje una barra de progreso sin meta, y que una operadora no vea
 * nada de esto.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const HOY = '2026-09-22';
const MIO = {
  ok: true, hoy: HOY, mes: '2026-09',
  trabajos: [
    { id: 't1', nombre: 'Son de Sky', contraparte: 'Dueño de Salsabor', tipo: 'cliente',
      estado: 'activo', moneda: 'COP', valor: 2400000, cobrado: 800000, falta: 1600000,
      entrega: '2026-09-30', horasSemana: 8, especificacion: 'Carta, web y precios',
      documento: 'https://ejemplo.com/contrato', nota: '' },
    { id: 't2', nombre: 'PHH', contraparte: 'Upwork', tipo: 'empleo', estado: 'activo',
      moneda: 'USD', valor: 0, cobrado: 0, falta: null, entrega: '', horasSemana: 12,
      especificacion: '', documento: '', nota: '' },
    { id: 't3', nombre: 'Parcial de Estadística', contraparte: 'Universidad',
      tipo: 'estudio', estado: 'activo', moneda: '', valor: 0, cobrado: 0, falta: null,
      entrega: '2026-09-20', horasSemana: 6, especificacion: '', documento: '', nota: '' },
  ],
  atrasados: [
    { id: 'c1', trabajo_id: 't1', trabajo: 'Son de Sky', concepto: 'Segundo pago',
      monto: 800000, moneda: 'COP', esperada: '2026-09-05', dias: 17 },
  ],
  porCobrar: [
    { id: 'c2', trabajo_id: 't1', trabajo: 'Son de Sky', concepto: 'Entrega final',
      monto: 800000, moneda: 'COP', esperada: '2026-10-05', dias: -13 },
    { id: 'c3', trabajo_id: 't2', trabajo: 'PHH', concepto: 'Semana 38',
      monto: 240, moneda: 'USD', esperada: '2026-09-26', dias: -4 },
  ],
  totales: {
    atrasado: { COP: 800000 },
    porCobrar: { COP: 800000, USD: 240 },
    ingresosMes: { COP: 800000 },
    gastosMes: { COP: 1250000 },
  },
  finanzas: [
    { id: 'f1', fecha: '2026-09-18', flujo: 'ingreso', categoria: 'trabajo',
      concepto: 'Primer pago Son de Sky', monto: 800000, moneda: 'COP', cuenta: 'Bancolombia' },
    { id: 'f2', fecha: '2026-09-02', flujo: 'gasto', categoria: 'vivienda',
      concepto: 'Arriendo', monto: 1250000, moneda: 'COP', cuenta: 'Bancolombia' },
  ],
  metas: [
    { id: 'm1', tipo: 'deuda', nombre: 'Semestre', conQuien: 'Universidad',
      meta: 4000000, saldo: 1500000, moneda: 'COP', cuota: 500000, dia: 5,
      fechaMeta: '2026-12-05', pct: 38, estado: 'activa' },
    // Sin monto meta: NO debe dibujar barra.
    { id: 'm2', tipo: 'ahorro', nombre: 'Colchón', conQuien: '', meta: 0,
      saldo: 300000, moneda: 'COP', cuota: 0, dia: 0, fechaMeta: '', pct: null,
      estado: 'activa' },
  ],
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };

  const abrir = async (ancho, permisos, datos) => {
    const p = await b.newPage({ viewport: { width: ancho, height: 1100 } });
    p.on('pageerror', e => console.log('  JS ERROR: ' + e.message));
    await p.goto('file:///home/claude/repo/novacentral.html');
    await p.waitForLoadState('load');
    await p.evaluate((x) => {
      window.nc = async (accion) => {
        if (accion === 'nc_mio') return x.datos;
        if (accion === 'nc_automatico') return { ok: true, automatico: {
          trabajos: [], todoPrendido: true, enFalta: 0, clientes: [], error: '' } };
        return { ok: true, clientes: [] };
      };
      window.NC_SESION = { nombre: 'Manuela', rol: 'socia', permisos: x.permisos };
      (0, eval)('NC_SESION = window.NC_SESION;');
      document.getElementById('login-screen').style.display = 'none';
      document.getElementById('app').style.display = 'flex';
    }, { permisos, datos });
    return p;
  };

  for (const ancho of [1100, 390]) {
    // ── La socia ──
    let p = await abrir(ancho, ['ver', 'facturacion'], MIO);
    await p.evaluate(() => cargarMio());
    await p.waitForTimeout(250);

    const trab = await p.textContent('#trab-lista');
    ok('lista los tres trabajos (@' + ancho + ')',
       trab.includes('Son de Sky') && trab.includes('PHH') && trab.includes('Estadística'));
    ok('la universidad aparece aunque no facture (@' + ancho + ')',
       trab.includes('no factura'));
    ok('dice los días que faltan, no solo la fecha (@' + ancho + ')',
       /faltan \d+ días/.test(trab), trab.match(/faltan[^·]*/)?.[0]);
    ok('y marca lo que ya venció (@' + ancho + ')',
       /venció hace \d+ días/.test(trab));

    const kpis = await p.textContent('#plata-kpis');
    ok('no suma monedas distintas: las muestra aparte (@' + ancho + ')',
       kpis.includes('COP 800.000') && kpis.includes('USD 240'), kpis.replace(/\s+/g, ' ').slice(0, 130));

    const cob = await p.textContent('#cobros-lista');
    ok('lo atrasado va primero y dice cuántos días (@' + ancho + ')',
       cob.indexOf('17 días tarde') !== -1 &&
       cob.indexOf('Segundo pago') < cob.indexOf('Entrega final'));

    const barras = await p.$$eval('#metas-lista .mio-barra', e => e.length);
    ok('la deuda con meta dibuja barra (@' + ancho + ')', barras === 1, barras + ' barras');
    ok('y el ahorro SIN meta no dibuja ninguna (@' + ancho + ')', barras === 1);

    const fam = await p.$$eval('#fam-links .fam', e => e.map(x => x.getAttribute('href')));
    ok('los cuatro accesos de la familia (@' + ancho + ')', fam.length === 4, JSON.stringify(fam));
    ok('y llevan a las páginas de verdad (@' + ancho + ')',
       fam.join(',').includes('novasoul.html') && fam.join(',').includes('empresarial.html'));


    // ── HOY ──
    await p.evaluate(() => {
      document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
      document.getElementById('v-dashboard').classList.add('on');
    });
    const hoyTxt = await p.textContent('#hoy-cifras');
    ok('Hoy cuenta los proyectos activos (@' + ancho + ')',
       hoyTxt.includes('Proyectos activos'));
    ok('y marca en rojo que hay una entrega vencida (@' + ancho + ')',
       await p.$$eval('#hoy-cifras .kpi-val.mal', e => e.length) === 2,
       await p.$$eval('#hoy-cifras .kpi-val.mal', e => e.length) + ' en rojo');

    const cuerpo = await p.textContent('#hoy-cuerpo');
    ok('reparte la semana por HORAS, no por plata (@' + ancho + ')',
       cuerpo.includes('26 h comprometidas'), cuerpo.match(/\d+ h comprometidas/)?.[0]);
    ok('y dice que los horarios fijos todavía no están (@' + ancho + ')',
       cuerpo.includes('sería inventarte la semana'));
    // Se mira SOLO la tercera tarjeta: «Son de Sky» también aparece en
    // la repartición de arriba, y buscarlo en todo el cuerpo comparaba
    // dos cosas distintas.
    const viene = await p.$eval('#hoy-cuerpo .card:nth-child(3)', e => e.textContent);
    ok('«Lo que viene» ordena por fecha y traduce a días (@' + ancho + ')',
       viene.indexOf('Parcial de Estadística') < viene.indexOf('Son de Sky') &&
       /venció hace \d+ días/.test(viene), viene.replace(/\s+/g, ' ').slice(0, 110));

    const dias = await p.$$eval('.hoy-dia', e => e.length);
    ok('la semana tiene siete días (@' + ancho + ')', dias === 7, dias + ' días');

    const proy = await p.$$eval('#sb-proyectos .ni-proy', e => e.map(x => x.textContent));
    ok('los proyectos salen también en el menú (@' + ancho + ')',
       proy.length === 3, JSON.stringify(proy));
    ok('y el vencido lleva su punto rojo (@' + ancho + ')',
       await p.$$eval('#sb-proyectos .ni-proy.tarde', e => e.length) === 1);

    const soul = await p.$$eval('.ni-link', e => e.map(x => x.getAttribute('href')));
    ok('NovaSoul tiene acceso directo desde el menú (@' + ancho + ')',
       soul.indexOf('./novasoul.html') !== -1, JSON.stringify(soul));

    // Un error al dibujar no puede decir «sin conexión»
    const msg = await p.evaluate(() => {
      const viejo = window.pintarHoy;
      window.pintarHoy = () => { throw new Error('a propósito'); };
      (0, eval)('pintarHoy = window.pintarHoy;');
      return cargarMio().then(() => {
        const t = document.getElementById('trab-lista').textContent;
        (0, eval)('pintarHoy = ' + viejo.toString() + ';');
        return t;
      });
    });
    ok('un fallo al dibujar NO se disfraza de falta de conexión (@' + ancho + ')',
       msg.includes('no los pude dibujar') && !msg.includes('Sin conexión'), msg.slice(0, 90));
    await p.evaluate(() => cargarMio());

    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral (@' + ancho + ')', lateral === 0, lateral + 'px');

    if (ancho === 1100) {
      await p.evaluate(() => {
        document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
        document.getElementById('v-trabajos').classList.add('on');
      });
      await p.evaluate(() => {
        document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
        document.getElementById('v-dashboard').classList.add('on');
      });
      await p.locator('#v-dashboard').screenshot({ path: OUT + 'central-hoy.png' });
      await p.evaluate(() => {
        document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
        document.getElementById('v-trabajos').classList.add('on');
      });
      await p.locator('#v-trabajos').screenshot({ path: OUT + 'central-trabajos.png' });
      await p.evaluate(() => {
        document.querySelectorAll('.view').forEach(v => v.classList.remove('on'));
        document.getElementById('v-plata').classList.add('on');
      });
      await p.locator('#v-plata').screenshot({ path: OUT + 'central-plata.png' });
    }
    await p.close();

    // ── Una operadora: nada de esto es suyo ──
    p = await abrir(ancho, ['ver', 'crear_cliente'],
                    { ok: false, error: 'Esta parte es solo de la socia.' });
    await p.evaluate(() => {
      if ((NC_SESION.permisos || []).indexOf('facturacion') !== -1) cargarMio();
      else document.querySelectorAll('.socia-solo').forEach(e => { e.style.display = 'none'; });
    });
    const visibles = await p.$$eval('.socia-solo', els =>
      els.filter(e => getComputedStyle(e).display !== 'none').length);
    ok('una operadora no ve «Mi trabajo» ni «Mi plata» (@' + ancho + ')', visibles === 0);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
