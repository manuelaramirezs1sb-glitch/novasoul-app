/**
 * NovaSoul en pantalla.
 *
 * Lo que se comprueba no es que pinte bonito: es que diga la verdad.
 * Que sin horas libres NO dibuje una semana holgada. Que lo vencido se
 * vea vencido. Que el tablero mueva de verdad y no solo de mentira.
 * Que un error al dibujar no se disfrace de falta de internet. Y que
 * nunca se vean dos tiendas a la vez.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const HOY = '2026-09-23';
const RIESGOS = {
  inamovible: { nombre: 'No se puede mover', orden: 3, ayuda: 'Un parcial, una entrega con fecha de ellos.' },
  acordado:   { nombre: 'Acordado con alguien', orden: 2, ayuda: 'Se puede correr, pero hay que avisar.' },
  corrible:   { nombre: 'Se puede correr', orden: 1, ayuda: 'Nadie está esperando una fecha exacta.' },
};
const TRABAJOS = [
  { id: 't1', nombre: 'PHH', tipo: 'empleo', estado: 'activo', horasSemana: 12, entrega: '' },
  { id: 't2', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo', horasSemana: 0, entrega: '2026-09-30' },
  { id: 't3', nombre: 'Universidad', tipo: 'estudio', estado: 'activo', horasSemana: 6, entrega: '' },
];
const T = (id, texto, o) => Object.assign({
  id, texto, trabajoId: '', proyecto: '', tipo: '', origen: '', fecha: '', dias: null,
  estado: 'pendiente', prioridad: 'media', riesgo: 'corrible', horas: 0, horasReales: 0,
  hechoEn: '', nota: '',
}, o);

const PEND = [
  T('p1', 'Entregar la carta de cocteles', { trabajoId: 't2', proyecto: 'Son de Sky', fecha: '2026-09-25', dias: -2, horas: 6, riesgo: 'acordado' }),
  T('p2', 'Parcial de Estadística', { trabajoId: 't3', proyecto: 'Universidad', fecha: '2026-09-20', dias: 3, horas: 4, riesgo: 'inamovible', prioridad: 'alta' }),
  T('p3', 'Encargo de PHH', { trabajoId: 't1', proyecto: 'PHH', fecha: '2026-09-24', dias: -1, horas: 5, estado: 'haciendo' }),
  T('p4', 'Comprar Omega 3', { fecha: '2026-09-26', dias: -3, horas: 1 }),
  T('p5', 'Llamar a mamá', { fecha: '2026-09-21', dias: 2, estado: 'hecho', hechoEn: '2026-09-21' }),
];

const dia = (f, dow, esHoy, libres, textos) => ({
  fecha: f, dow, esHoy, libres,
  entregas: textos.length, horas: textos.reduce((a, t) => a + t.horas, 0), textos,
});

function foto(conHoras) {
  return {
    ok: true, hoy: HOY, mes: '2026-09',
    semana: { lunes: '2026-09-21', domingo: '2026-09-27', dias: [
      dia('2026-09-21', 1, false, conHoras ? 4 : null, []),
      dia('2026-09-22', 2, false, conHoras ? 4 : null, []),
      dia('2026-09-23', 3, true, conHoras ? 4 : null, []),
      dia('2026-09-24', 4, false, conHoras ? 4 : null, [{ id: 'p3', texto: 'Encargo de PHH', proyecto: 'PHH', horas: 5 }]),
      dia('2026-09-25', 5, false, conHoras ? 4 : null, [{ id: 'p1', texto: 'Entregar la carta de cocteles', proyecto: 'Son de Sky', horas: 6 }]),
      dia('2026-09-26', 6, false, conHoras ? 0 : null, [{ id: 'p4', texto: 'Comprar Omega 3', proyecto: '', horas: 1 }]),
      dia('2026-09-27', 7, false, conHoras ? 0 : null, []),
    ] },
    trabajos: TRABAJOS,
    tipos: { cliente: { nombre: 'Cliente' }, empleo: { nombre: 'Empleo' },
             propio: { nombre: 'Propio (Nova)' }, estudio: { nombre: 'Universidad' } },
    riesgos: RIESGOS,
    pendientes: PEND,
    urgente: PEND[1],
    resumen: {
      dia: { entregas: 0, horas: 0, vencidas: 1, hechasHoy: 0, haciendo: 1 },
      semana: { entregas: 3, horasEntregas: 12, hechas: 1 },
      mes: { entregas: 5, hechas: 1, abiertas: 4, porProyecto: [
        { id: 't2', nombre: 'Son de Sky', abiertas: 1, horas: 6, vencidas: 0 },
        { id: 't1', nombre: 'PHH', abiertas: 1, horas: 5, vencidas: 0 },
        { id: 't3', nombre: 'Universidad', abiertas: 1, horas: 4, vencidas: 1 },
        { id: '·sueltas', nombre: 'Sin proyecto', abiertas: 1, horas: 1, vencidas: 0 },
      ] },
    },
    riesgo: conHoras
      ? { libres: 20, fijas: 18, extra: 7, comprometidas: 25, sobra: -5, dentroDeFijas: 2,
          candidatas: [
            { id: 'p4', texto: 'Comprar Omega 3', horas: 1, proyecto: '', riesgo: 'corrible', fecha: '2026-09-26' },
            { id: 'p1', texto: 'Entregar la carta de cocteles', horas: 6, proyecto: 'Son de Sky', riesgo: 'acordado', fecha: '2026-09-25' },
          ], noAlcanza: 0, sinHoras: false, porque: '' }
      : { libres: null, fijas: 18, extra: 7, comprometidas: 25, sobra: null, dentroDeFijas: 2,
          candidatas: [], noAlcanza: 0, sinHoras: true,
          porque: 'Todavía no me has dicho cuántas horas libres tienes cada día. ' +
                  'Sin ese número no puedo decirte si la semana cabe: lo demás sería un adorno.' },
    horas: conHoras ? { 1: 4, 2: 4, 3: 4, 4: 4, 5: 4, 6: 0, 7: 0 } : null,
    mindlab: {
      inicio: '2026-09-28', fin: '2026-12-20', hechas: 1, total: 12, horasTotales: 37,
      semanaActual: null,
      plan: [1,2,3,4,5,6,7,8,9,10,11,12].map(function (n) {
        return { id: 'ml' + n, semana: n, mes: n <= 4 ? 1 : n <= 8 ? 2 : 3,
                 tema: 'Tema ' + n, tarea: 'Tarea de la semana ' + n,
                 horas: n === 11 || n === 12 ? 2 : 3,
                 desde: '2026-09-28', hasta: '2026-10-04',
                 estado: n === 1 ? 'hecho' : 'pendiente', nota: '' };
      }),
    },
  };
}

const FINANZAS = {
  ok: true, mes: '2026-09', hoy: HOY,
  categorias: [
    { id: 'arriendo', nombre: 'Arriendo', tipo: 'gasto',
      lineas: [{ id: 'x1', categoria: 'arriendo', concepto: 'Apartamento', monto: 1250000, moneda: 'COP', dia: 5, activo: true, nota: '' }],
      planeado: { COP: 1250000 }, gastado: { COP: 1400000 } },
    { id: 'mercado', nombre: 'Mercado', tipo: 'gasto', lineas: [], planeado: null, gastado: { COP: 420000 } },
    { id: 'servicios', nombre: 'Servicios', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'internet', nombre: 'Internet', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'credito', nombre: 'Crédito', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'deudas', nombre: 'Deudas', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'movil', nombre: 'Móvil', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'varios', nombre: 'Gastos varios', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'ahorro', nombre: 'Ahorro', tipo: 'ahorro',
      lineas: [{ id: 'x2', categoria: 'ahorro', concepto: 'Colchón', monto: 300000, moneda: 'COP', dia: 30, activo: true, nota: '' }],
      planeado: { COP: 300000 }, gastado: null },
  ],
  totales: { plan: { COP: 1250000 }, ahorro: { COP: 300000 } },
  error: '',
};

const FAMILY = {
  ok: true,
  clientes: [{ id: 'c1', empresa: 'Nutrea', sheetId: 'sid' }],
  cliente: { id: 'c1', empresa: 'Nutrea', sheetId: 'sid' },
  tiendas: [{ id: 'gt', nombre: 'Nutrea GT' }, { id: 'ec', nombre: 'Nutrea EC' }],
  tienda: 'gt',
  alarmas: [{ id: 'a1', nivel: 'rojo', nombre: 'Novedades', titulo: '4 novedades sin resolver',
              detalle: 'Llevan más de 24 horas.', casos: 4 }],
  semaforo: { tienda: 'gt', moneda: 'GTQ', semana: { lunes: '2026-09-14', domingo: '2026-09-20' },
    luces: { entrega: { estado: 'verde', valor: '78%', etiqueta: 'Tasa de entrega', porque: 'sobre 65% mínimo' },
             cpa: { estado: 'rojo', valor: 'GTQ 92', etiqueta: 'CPA contra el techo', porque: 'el techo es GTQ 70' } },
    alertas: [{ titulo: 'El CPA pasó el techo tres semanas seguidas.' }],
    hayDatos: true },
  errorTienda: '',
  central: { proyectos: 3, entregasVencidas: 1, cobrosAtrasados: 1,
             atrasadoPorMoneda: { COP: 800000 }, carga: {} },
  academy: { estudiantes: 0, porque: 'Todavía no has dado de alta a nadie en novAcademy.' },
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };

  const abrir = async (ancho, conHoras, rol) => {
    const p = await b.newPage({ viewport: { width: ancho, height: 1200 } });
    p.on('pageerror', e => { fallas++; console.log('  JS ERROR: ' + e.message); });
    await p.goto('file:///home/claude/repo/novasoul.html');
    await p.waitForLoadState('load');
    await p.evaluate((x) => {
      window.LLAMADAS = [];
      window.nc = async (accion, datos) => {
        window.LLAMADAS.push({ accion, datos });
        if (accion === 'nc_soul') return x.foto;
        if (accion === 'nc_soul_finanzas') return x.fin;
        if (accion === 'nc_soul_family') return x.fam;
        return { ok: true };
      };
      (0, eval)('nc = window.nc;');
      (0, eval)('SES = ' + JSON.stringify({ correo: 'm@nova.com', nombre: 'Manuela', rol: x.rol }) + ';');
      (0, eval)('entrar();');
    }, { foto: foto(conHoras), fin: FINANZAS, fam: FAMILY, rol: rol || 'socia' });
    await p.waitForTimeout(200);
    return p;
  };

  for (const ancho of [1200, 390]) {
    const A = ' (@' + ancho + ')';

    // ══ Sin horas libres: el caso que más importa ══
    let p = await abrir(ancho, false);
    ok('la puerta se cierra al entrar' + A,
       await p.$eval('#login-screen', e => getComputedStyle(e).display) === 'none');

    let riesgo = await p.textContent('#hoy-riesgo');
    ok('sin horas NO dibuja una semana holgada: pide el dato' + A,
       /cuántas horas libres/.test(riesgo), riesgo.replace(/\s+/g, ' ').slice(0, 90));
    ok('y no inventa una barra' + A,
       await p.$$eval('#hoy-riesgo .hbar', e => e.length) === 0);
    ok('ofrece el botón para darlo' + A,
       /Decirle cuántas horas tengo/.test(riesgo));

    let semDias = await p.textContent('#sem-dias');
    ok('y cada día dice que faltan las horas, en vez de poner cero' + A,
       /sin horas dichas/.test(semDias));
    await p.close();

    // ══ Con horas: la cuenta completa ══
    p = await abrir(ancho, true);

    const kpis = await p.textContent('#hoy-kpis');
    ok('las cifras de arriba: hoy, vencidas, semana y mes' + A,
       /ENTREGAS DE HOY/.test(kpis) && /VENCIDAS/.test(kpis) &&
       /ESTA SEMANA/.test(kpis) && /ESTE MES/.test(kpis));
    ok('lo vencido va en rojo' + A,
       await p.$$eval('#hoy-kpis .kpi-v.mal', e => e.length) === 1,
       await p.$$eval('#hoy-kpis .kpi-v.mal', e => e.length) + ' en rojo');

    const urg = await p.textContent('#hoy-urgente');
    ok('«lo primero» es UNA sola cosa, la vencida' + A,
       /Parcial de Estadística/.test(urg) && /venció hace 3 días/.test(urg));
    ok('y dice que no se puede mover' + A, /No se puede mover/.test(urg));

    riesgo = await p.textContent('#hoy-riesgo');
    ok('dice cuántas horas hay y cuántas se pidieron' + A,
       /20 h libres/.test(riesgo) && /25 h comprometidas/.test(riesgo),
       riesgo.replace(/\s+/g, ' ').slice(0, 120));
    ok('y cuántas faltan' + A, /FALTAN 5 h/.test(riesgo));
    ok('explica que lo de un proyecto con horas fijas no se cuenta dos veces' + A,
       /no se suman otra vez/.test(riesgo));
    ok('propone candidatas a caerse' + A, /ALGO SE VA A CAER/.test(riesgo));
    ok('y el parcial NO está entre ellas' + A,
       riesgo.split('ALGO SE VA A CAER')[1].indexOf('Parcial') === -1);
    ok('la barra reparte en tres tramos' + A,
       await p.$$eval('#hoy-riesgo .hbar span', e => e.length) === 3);

    // ══ El tablero ══
    await p.evaluate(() => go('pendientes', null));
    const cols = await p.$$eval('#pend-kan .kcol', e => e.length);
    ok('el tablero tiene tres columnas' + A, cols === 3, cols + ' columnas');
    const tarjetas = await p.$$eval('#pend-kan .kcard', e => e.length);
    ok('y las cinco tarjetas' + A, tarjetas === 5, tarjetas + ' tarjetas');
    const enHaciendo = await p.$eval('#kcol-haciendo', e => e.textContent);
    ok('cada una en su columna' + A, /Encargo de PHH/.test(enHaciendo));
    const kanTxt = await p.textContent('#pend-kan');
    ok('lo vencido dice cuántos días lleva' + A, /venció hace 3 días/.test(kanTxt));
    ok('y lo inamovible se marca' + A, /no se mueve/.test(kanTxt));
    ok('lo hecho lleva su fecha, no «hoy»' + A, /21 sep/.test(kanTxt));

    // Mover de verdad llama al servidor con el estado nuevo
    await p.evaluate(() => { window.LLAMADAS.length = 0; });
    await p.click('#kcol-pendiente .kcard .kbtn');
    await p.waitForTimeout(120);
    const llamadas = await p.evaluate(() => window.LLAMADAS);
    ok('mover una tarjeta lo guarda en el servidor' + A,
       llamadas.length >= 1 && llamadas[0].accion === 'nc_soul_guardar' &&
       llamadas[0].datos.datos.estado === 'haciendo',
       JSON.stringify(llamadas[0] || null));
    ok('y vuelve a traer la foto, para no pintar lo que cree que pasó' + A,
       llamadas.filter(l => l.accion === 'nc_soul').length === 1);

    // El filtro por proyecto
    const filtros = await p.$$eval('#pend-filtros .kbtn', e => e.map(x => x.textContent));
    ok('se puede filtrar por proyecto' + A,
       filtros.indexOf('Son de Sky') !== -1, JSON.stringify(filtros));

    // ══ El formulario ══
    await p.evaluate(() => nuevaTarea());
    ok('el formulario abre' + A,
       await p.$eval('#m-tarea', e => getComputedStyle(e).display) !== 'none');
    const proys = await p.$$eval('#mt-trabajo option', e => e.map(x => x.textContent));
    ok('y ofrece los proyectos de Central' + A,
       proys.length === 4 && proys.indexOf('PHH') !== -1, JSON.stringify(proys));
    ok('tiene horas estimadas y fecha de entrega' + A,
       (await p.$$('#mt-horas')).length === 1 && (await p.$$('#mt-fecha')).length === 1);
    ok('y pregunta qué pasa si no se entrega' + A,
       (await p.$$eval('#mt-riesgo option', e => e.length)) === 3);
    await p.fill('#mt-texto', '');
    await p.click('#mt-guardar');
    await p.waitForTimeout(80);
    ok('no deja guardar una tarea sin texto' + A,
       /Escribe qué hay que hacer/.test(await p.textContent('#mt-err')));
    await p.evaluate(() => cerrarModal('m-tarea'));

    // ══ Mindlab ══
    await p.evaluate(() => go('mindlab', null));
    const ml = await p.textContent('#v-mindlab');
    ok('Mindlab son doce semanas' + A,
       await p.$$eval('#ml-lista .fila', e => e.length) === 12);
    ok('con sus tres meses' + A,
       /MES 1 · TRÁFICO/.test(ml) && /MES 3 · VENTA/.test(ml));
    ok('dice cuánto pesa en total y que lo recorté' + A,
       /37 h en total/.test(ml) && /último trimestre/.test(ml), ml.replace(/\s+/g, ' ').slice(0, 160));
    ok('y por qué las dos últimas semanas del año quedan libres' + A,
       /finales y fiestas/.test(ml));

    // ══ Gastos fijos ══
    await p.evaluate(() => go('plata', null));
    await p.waitForTimeout(150);
    const fin = await p.textContent('#fin-lista');
    ok('están las nueve categorías que pidió' + A,
       await p.$$eval('#fin-lista .fin-row', e => e.length) === 10, // 9 + la cabecera
       await p.$$eval('#fin-lista .fin-row', e => e.length) + ' filas');
    ok('con sus nombres' + A,
       ['Arriendo','Mercado','Servicios','Internet','Crédito','Deudas','Móvil',
        'Gastos varios','Ahorro'].every(n => fin.indexOf(n) !== -1));
    ok('lo que no tiene monto dice «sin definir», no cero' + A,
       /sin definir/.test(fin) && !/COP 0/.test(fin));
    ok('y donde se pasó, lo dice' + A, /se pasó COP 150.000/.test(fin));
    ok('el ahorro no se suma con los gastos' + A,
       /va aparte, no es un gasto/.test(await p.textContent('#fin-kpis')));

    // ══ Nova Family ══
    await p.evaluate(() => go('family', null));
    await p.waitForTimeout(200);
    const fam = await p.textContent('#fam-cuerpo');
    ok('trae las alarmas de la tienda' + A, /4 novedades sin resolver/.test(fam));
    ok('y el semáforo de la pauta' + A,
       /EL SEMÁFORO DE LA PAUTA/.test(fam) && /CPA contra el techo/.test(fam));
    ok('dice que es el mismo del correo del lunes' + A, /correo los lunes/.test(fam));
    ok('resume Nova Central' + A, /Cobros atrasados/.test(fam) && /COP 800.000/.test(fam));
    ok('y dice la verdad de novAcademy en vez de un cero mudo' + A,
       /Todavía no has dado de alta/.test(fam));
    /**
     * La regla de toda Nova: nunca dos tiendas en la misma pantalla.
     * El selector las ofrece, el cuerpo muestra una.
     */
    ok('NUNCA muestra dos tiendas a la vez' + A,
       fam.indexOf('Nutrea EC') === -1 && /Nutrea GT/.test(fam));
    ok('pero deja cambiar de tienda' + A,
       (await p.$$eval('#fam-sel .kbtn', e => e.map(x => x.textContent))).join(',') === 'Nutrea GT,Nutrea EC');

    // ══ Un fallo al dibujar no se disfraza de falta de internet ══
    const msg = await p.evaluate(async () => {
      const viejo = window.pintarHoy;
      window.pintarHoy = () => { throw new Error('a propósito'); };
      (0, eval)('pintarHoy = window.pintarHoy;');
      await cargar();
      const t = document.getElementById('av-global').textContent;
      (0, eval)('pintarHoy = ' + viejo.toString() + ';');
      return t;
    });
    ok('un fallo al dibujar NO dice «sin conexión»' + A,
       /no los pude dibujar/.test(msg) && !/conexión/.test(msg) && !/hablar con la hoja/.test(msg),
       msg.slice(0, 90));
    await p.evaluate(() => cargar());
    await p.waitForTimeout(120);

    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral' + A, lateral === 0, lateral + 'px');

    if (ancho === 1200) {
      for (const [v, f] of [['hoy', 'soul-hoy'], ['pendientes', 'soul-tablero'],
                            ['semana', 'soul-semana'], ['plata', 'soul-plata'],
                            ['family', 'soul-family'], ['mindlab', 'soul-mindlab']]) {
        await p.evaluate((x) => go(x, null), v);
        await p.waitForTimeout(120);
        await p.locator('#v-' + v).screenshot({ path: OUT + f + '.png' });
      }
    }
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
