/**
 * EL PRIMER DÍA DE SU HERMANA, EN LA PANTALLA.
 *
 * ┌─ QUÉ SE ESTÁ COMPROBANDO ──────────────────────────────────┐
 * │                                                            │
 * │ «le voy a crear el usuario a mi hermana y necesito que      │
 * │  esté lista para que ella lo use sin daños, ni              │
 * │  confusiones o bugs, o cosas que no sirvan».                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y POR QUÉ CAMBIÓ ─────────────────────────────────────────┐
 * │                                                            │
 * │ La primera versión probaba un reparto pedido por pedido.   │
 * │ Ella lo corrigió: «no se asignan pedidos, se asignan       │
 * │ tiendas. La gestora de la tienda gestiona la tienda que    │
 * │ tiene asignada».                                           │
 * │                                                            │
 * │ Así que lo que hay que comprobar es otro: que una gestora  │
 * │ con tienda vea TODO lo de esa tienda desde el primer       │
 * │ minuto, y que una sin tienda sepa exactamente qué le       │
 * │ falta — que no es lo mismo.                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

const GENTE = [
  { id: 'e1', nombre: 'Manuela Ramírez', correo: 'm@nova.com', rol: 'dueno',
    total: 0, abiertos: 0 },
  { id: 'e2', nombre: 'Andrea Ramírez', correo: 'andrea@nutrea.com', rol: 'gestora',
    total: 0, abiertos: 0 },
];

const pedido = (i, gestora) => ({
  id: 'd' + i, id_externo: 'NE-44' + i, tienda: 'ec', fecha: '2026-09-22',
  cliente: 'Clienta ' + i, telefono: '59399900' + i, telefono_norm: '59399900' + i,
  ciudad: 'Quito', direccion: 'Av. Siempre Viva ' + i, producto: 'TAG RECEDE',
  valor: 50, estado: 'EN TRANSITO', estado_canonico: 'en_transito',
  estado_nova: '', gestora_asignada: gestora || '', nota: '',
  transportadora: 'Servientrega', guia: 'G' + i, ultimo_movimiento: '2026-09-22',
});

const PEDIDOS = [0, 1, 2, 3, 4, 5].map(i => pedido(i));

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

async function abrir(b, rol, pedidos, reparto, tiendas) {
  const p = await b.newPage({ viewport: { width: 1280, height: 1200 } });
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

  await p.evaluate(({ rol, pedidos, reparto, tiendas }) => {
    document.getElementById('login').style.display = 'none';
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO = true; ST = "ec"; ROL = "' + rol + '";');
    document.documentElement.setAttribute('data-rol', rol);
    /**
     * `SESION` está declarada con `let`, así que `window.SESION = …` NO
     * la toca: crea una propiedad distinta con el mismo nombre y la
     * variable real sigue en null. Hay que asignarla por eval, igual
     * que las otras. (Mi primera versión de esta prueba lo hizo mal y
     * «falló» la frase del saludo — el error era de la prueba.)
     */
    (0, eval)('SESION = ' + JSON.stringify({
      nombre: rol === 'gestora' ? 'Andrea Ramírez' : 'Manuela Ramírez',
      rol: rol, correo: rol === 'gestora' ? 'andrea@nutrea.com' : 'm@nova.com',
      tiendas: tiendas || ['ec'], permisos: [], modulos: ['empresarial'] }));
    window.api = async (accion, params) => {
      if (accion === 'listar' && params.entidad === 'Pedidos') {
        return { ok: true, filas: JSON.parse(JSON.stringify(pedidos)),
                 total: pedidos.length, entidad: 'Pedidos' };
      }
      if (accion === 'listar') return { ok: true, filas: [], total: 0, entidad: params.entidad };
      if (accion === 'reparto') return JSON.parse(JSON.stringify(reparto));
      if (accion === 'asignar') {
        window.__anotar(JSON.stringify({ accion, params }));
        return { ok: true, asignados: (params.ids || []).length, a: 'Andrea Ramírez',
                 novedades: 2, iguales: 0, desasignado: !params.a };
      }
      return { ok: true };
    };
  }, { rol, pedidos, reparto, tiendas });

  return { p, enviados, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 1 · SU HERMANA ENTRA SIN TIENDA ASIGNADA ══');
  {
    const { p, errores } = await abrir(b, 'gestora', [], { ok: false }, []);
    await p.evaluate(() => { PEDS_REALES = []; NOVS_REALES = []; pintarPantallasGestora(); });
    await p.waitForTimeout(150);

    const hoy = (await p.textContent('#v-g-hoy')).replace(/\s+/g, ' ');
    ok('NO le dice «todo al día» con cuatro ceros', !/0%|MI EFECTIVIDAD/.test(hoy), hoy.slice(0, 200));
    ok('le dice que no tiene ninguna tienda',
       /no tienes ninguna tienda asignada/i.test(hoy), hoy.slice(0, 250));
    ok('la llama por su nombre', /Andrea/.test(hoy));
    ok('y dice a quién pedírselo', /Permisos/.test(hoy), hoy.slice(0, 350));

    console.log('\n── Y no ve nada que no le toca ──');
    for (const v of ['pauta', 'dinero', 'permisos', 'config', 'cierre', 'productos']) {
      const oculto = await p.evaluate((id) => {
        const el = document.getElementById('v-' + id);
        if (!el) return true;
        return getComputedStyle(el).display === 'none';
      }, v);
      ok('no ve ' + v, oculto);
    }
    ok('sin errores de JavaScript', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  console.log('\n══ 1b · CON TIENDA, PERO LA TIENDA ESTÁ VACÍA ══');
  {
    /**
     * Es un problema DISTINTO del anterior y por eso el mensaje también:
     * aquí no falta un permiso, falta importar. Confundirlos manda a la
     * persona a pedir algo que ya tiene.
     */
    const { p, errores } = await abrir(b, 'gestora', [], { ok: false }, ['ec']);
    await p.evaluate(() => { PEDS_REALES = []; NOVS_REALES = []; pintarPantallasGestora(); });
    await p.waitForTimeout(150);
    const hoy = (await p.textContent('#v-g-hoy')).replace(/\s+/g, ' ');
    ok('no le habla de permisos', !/Permisos/.test(hoy), hoy.slice(0, 250));
    ok('le dice que la tienda está vacía',
       /No hay pedidos en/i.test(hoy), hoy.slice(0, 250));
    ok('y le recuerda que ve la tienda entera',
       /toda<\/strong> esta tienda|toda esta tienda/i.test(hoy), hoy.slice(0, 300));
    ok('sin errores de JavaScript', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 2 · LO QUE VE MANUELA DE SU TIENDA ══');
  {
    /**
     * Aquí había una barra para repartir pedidos, con un selector de
     * persona y un botón de «asignarle los N de la lista». Se borró:
     * «de nada sirve asignar pedidos si ya asignaste la tienda».
     *
     * Lo que quedó no pide nada — informa. Y es lo que sí sirve todos
     * los días: quién tiene esta tienda y cuánto lleva hecho cada quien.
     */
    const reparto = {
      ok: true, tienda: 'ec',
      gente: GENTE,
      total: 6, sinTocar: 4, aNadieConocido: [], huerfanos: {},
      etiquetas: [
        { nombre: 'JAIME', total: 1, abiertos: 1, esUsuario: false },
        { nombre: 'Andrea Ramírez', total: 1, abiertos: 0, esUsuario: true },
      ],
    };
    const { p, errores } = await abrir(b, 'dueno', PEDIDOS, reparto, ['ec']);
    await p.evaluate(() => cargarPedidosReales());
    await p.waitForTimeout(200);
    await p.evaluate(() => cargarReparto());
    await p.waitForTimeout(200);
    await p.evaluate(() => go('pedidos', document.querySelector('[data-v=pedidos]')));
    await p.waitForTimeout(150);

    const rep = (await p.textContent('#ped-reparto')).replace(/\s+/g, ' ');
    ok('dice cuántos no ha tocado nadie', /4 sin tocar de 6/.test(rep), rep.slice(0, 220));
    ok('y deja claro que no hay que repartir',
       /No hay que repartir nada/.test(rep), rep.slice(0, 260));
    ok('nombra a quien tiene la tienda', /Andrea Ram/.test(rep));

    console.log('\n── Y no queda ni rastro del reparto viejo ──');
    ok('no hay selector de persona',
       (await p.$$('#rep-quien')).length === 0);
    ok('ni botón de asignar',
       !/Asignarle los/.test(rep), rep.slice(0, 260));
    ok('ni selector en el detalle del pedido',
       await p.evaluate(() => { PED_SEL_ID = 'd0'; pintarPanelPedido();
                                return document.getElementById('pd-gestora') === null; }));

    console.log('\n── Pero sí dice quién lo trabajó ──');
    const panel = (await p.textContent('#pedpanel')).replace(/\s+/g, ' ');
    ok('lo pregunta', /LO TRABAJÓ/.test(panel), panel.slice(0, 250));
    ok('y si nadie, lo dice', /Nadie todav/.test(panel));

    console.log('\n── Quién hizo el trabajo, aunque no tenga cuenta ──');
    ok('JAIME aparece', /JAIME/.test(rep), rep.slice(0, 300));
    ok('marcado como sin cuenta', /sin cuenta/.test(rep));
    ok('y Andrea no lleva esa marca',
       rep.indexOf('Andrea Ramírez sin cuenta') === -1);

    console.log('\n── Sin equipo todavía, lo explica ──');
    await p.evaluate(() => {
      REPARTO.gente = [{ id: 'e1', nombre: 'Manuela Ramírez', rol: 'dueno' }];
      pintarReparto();
    });
    await p.waitForTimeout(120);
    const rep3 = (await p.textContent('#ped-reparto')).replace(/\s+/g, ' ');
    ok('dice que agregue a alguien y en qué tiendas',
       /Permisos/.test(rep3) && /en qué tiendas trabaja/.test(rep3), rep3.slice(0, 300));

    ok('sin errores de JavaScript', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 2b · «VER LO QUE VE» SU HERMANA ══');
  {
    /**
     * Esta pantalla tenía tres personas inventadas —Camila Restrepo,
     * Valentina Ospina, Daniela Ríos— con pedidos falsos, y encima no
     * la llamaba nadie. Ahora cuenta de verdad, con los pedidos que la
     * dueña ya tiene cargados.
     */
    const mezcla = [
      pedido(0, 'Andrea Ramírez'), pedido(1, 'Andrea Ramírez'),
      pedido(2, 'Andrea Ramírez'), pedido(3), pedido(4),
    ];
    // Ahora lo que cuenta es lo que TRABAJÓ, no lo que le asignaron.
    mezcla.forEach(function (x) { x.gestionado_por = x.gestora_asignada; });
    mezcla[0].estado_canonico = 'entregado';
    mezcla[1].estado_canonico = 'devolucion';
    mezcla[2].estado_canonico = 'novedad';

    const reparto = { ok: true, tienda: 'ec', gente: GENTE, total: 5, sinAsignar: 2,
                      aNadieConocido: [], huerfanos: {} };
    const { p, errores } = await abrir(b, 'dueno', mezcla, reparto);
    await p.evaluate((m) => {
      PEDS_REALES = m; PEDS_VISTA = m; NOVS_REALES = [];
      MI_EQUIPO = [{ id: 'e2', nombre: 'Andrea Ramírez', correo: 'andrea@nutrea.com',
                     rol: 'gestora', tienda: 'ec', estado: 'activo',
                     ultima_conexion: '2026-09-25T08:00:00Z',
                     pedidos: 3, efectividad: 50, sinMover: 0 }];
      showSombra('Andrea Ramírez');
    }, mezcla);
    await p.waitForTimeout(200);

    const som = (await p.textContent('#sombra')).replace(/\s+/g, ' ');
    const INVENTADOS = ['Camila Restrepo', 'Valentina Ospina', 'Daniela Ríos',
                        'Carlos Mendoza', 'Ana Ruiz', 'María Salazar',
                        'Proteína Vainilla', 'Colágeno'];
    const quedan = INVENTADOS.filter(x => som.indexOf(x) !== -1);
    ok('ni una persona ni un producto inventado', quedan.length === 0,
       'siguen ahí: ' + quedan.join(', '));

    ok('cuenta los 3 que trabajó', /Trabajados\s*3/.test(som), som.slice(0, 250));
    /** 1 entregado de 2 resueltos (entregado + devolución) = 50%. */
    ok('y la efectividad sobre los RESUELTOS: 50%', /50%/.test(som), som.slice(0, 250));
    ok('dice cuándo entró', /entr[óo] el 2026-09-25/.test(som), som.slice(0, 250));
    ok('y muestra sus clientas de verdad', /Clienta 2/.test(som), som.slice(0, 400));

    console.log('\n── Y si no le han asignado nada, lo dice ──');
    await p.evaluate(() => {
      MI_EQUIPO = [{ id: 'e3', nombre: 'Recién Llegada', correo: 'r@nutrea.com',
                     rol: 'gestora', tienda: 'ec', estado: 'activo',
                     ultima_conexion: '', pedidos: 0, efectividad: 0, sinMover: 0 }];
      showSombra('Recién Llegada');
    });
    await p.waitForTimeout(150);
    const som2 = (await p.textContent('#sombra')).replace(/\s+/g, ' ');
    ok('dice que no ha tocado nada',
       /no ha tocado ning[úu]n caso/i.test(som2), som2.slice(0, 300));
    /** Y aclara que NO es falta de acceso: ve la tienda entera. */
    ok('y aclara que sí tiene acceso',
       /ve la tienda entera/.test(som2), som2.slice(0, 300));
    ok('y que nunca ha entrado', /nunca ha entrado/.test(som2));

    ok('sin errores de JavaScript', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 3 · SU HERMANA VUELVE, YA CON TRABAJO ══');
  {
    const mios = [0, 1, 2].map(i => pedido(i, 'Andrea Ramírez'));
    mios[0].estado_canonico = 'novedad';
    const { p, errores } = await abrir(b, 'gestora', mios, { ok: false });
    await p.evaluate((m) => {
      PEDS_REALES = m; NOVS_REALES = [];
      pintarPantallasGestora();
    }, mios);
    await p.waitForTimeout(150);

    const hoy = (await p.textContent('#v-g-hoy')).replace(/\s+/g, ' ');
    ok('ya NO sale el aviso de vacío', !/no tienes pedidos asignados/i.test(hoy));
    ok('ve cuántos le faltan', /ME FALTAN/.test(hoy), hoy.slice(0, 200));
    ok('y cuál está en novedad', /EN NOVEDAD/.test(hoy));
    ok('con sus clientas', /Clienta/.test(hoy), hoy.slice(0, 300));
    ok('sin errores de JavaScript', errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
  process.exit(fallas ? 1 : 0);
})();
