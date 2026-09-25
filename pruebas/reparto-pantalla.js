/**
 * EL PRIMER DÍA DE SU HERMANA, EN LA PANTALLA.
 *
 * ┌─ QUÉ SE ESTÁ COMPROBANDO ──────────────────────────────────┐
 * │                                                            │
 * │ Ella: «le voy a crear el usuario a mi hermana y necesito   │
 * │  que esté lista para que ella lo use sin daños, ni         │
 * │  confusiones o bugs, o cosas que no sirvan».               │
 * │                                                            │
 * │ Lo que se prueba aquí es exactamente eso, en el orden en   │
 * │ que va a pasar:                                            │
 * │                                                            │
 * │   1· La hermana entra sin nada asignado. ¿Qué ve?          │
 * │      Tiene que ver POR QUÉ está vacío — no cuatro ceros    │
 * │      verdes que parecen «todo al día».                     │
 * │                                                            │
 * │   2· Manuela reparte. ¿Puede, desde dónde, y se entera de  │
 * │      cuántos quedan sin repartir?                          │
 * │                                                            │
 * │   3· La hermana vuelve. ¿Ahora sí ve su trabajo?           │
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

async function abrir(b, rol, pedidos, reparto) {
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

  await p.evaluate(({ rol, pedidos, reparto }) => {
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
      rol: rol, correo: rol === 'gestora' ? 'andrea@nutrea.com' : 'm@nova.com' }));
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
  }, { rol, pedidos, reparto });

  return { p, enviados, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 1 · SU HERMANA ENTRA Y NO TIENE NADA ASIGNADO ══');
  {
    /** El servidor filtra por rol: a ella le llegan CERO filas. */
    const { p, errores } = await abrir(b, 'gestora', [], { ok: false });
    await p.evaluate(() => { PEDS_REALES = []; NOVS_REALES = []; pintarPantallasGestora(); });
    await p.waitForTimeout(150);

    const hoy = (await p.textContent('#v-g-hoy')).replace(/\s+/g, ' ');
    ok('NO le dice «todo al día» con cuatro ceros', !/0%|MI EFECTIVIDAD/.test(hoy), hoy.slice(0, 200));
    ok('le dice que no tiene pedidos asignados',
       /Todav[ií]a no tienes pedidos asignados/.test(hoy), hoy.slice(0, 250));
    ok('la llama por su nombre', /Andrea/.test(hoy));
    ok('aclara que NO es que algo falló',
       /no haya fallado|no es que est[eé] vac[ií]o/i.test(hoy), hoy.slice(0, 300));
    ok('y dice qué tiene que pasar para que cambie',
       /Repartir el trabajo/.test(hoy), hoy.slice(0, 350));

    const peds = (await p.textContent('#v-g-pedidos')).replace(/\s+/g, ' ');
    ok('lo mismo en su pantalla de pedidos', /no tienes pedidos asignados/i.test(peds));

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

  // ══════════════════════════════════════════════════════════
  console.log('\n══ 2 · MANUELA REPARTE ══');
  {
    const reparto = { ok: true, tienda: 'ec', gente: GENTE, total: 6, sinAsignar: 6,
                      aNadieConocido: [], huerfanos: {} };
    const { p, enviados, errores } = await abrir(b, 'dueno', PEDIDOS, reparto);
    await p.evaluate(() => cargarPedidosReales());
    await p.waitForTimeout(200);
    await p.evaluate(() => cargarReparto());
    await p.waitForTimeout(200);
    await p.evaluate(() => go('pedidos', document.querySelector('[data-v=pedidos]')));
    await p.waitForTimeout(150);

    const rep = (await p.textContent('#ped-reparto')).replace(/\s+/g, ' ');
    ok('ve cuántos están sin asignar', /6 sin asignar de 6/.test(rep), rep.slice(0, 200));
    ok('y le explican por qué importa',
       /nadie m[áa]s|pantalla vac[ií]a/i.test(rep), rep.slice(0, 300));
    ok('puede elegir a su hermana', /Andrea Ram/.test(rep));

    console.log('\n── Repartir en bloque ──');
    await p.evaluate(() => {
      document.getElementById('rep-quien').value = 'e2';
      window.confirm = () => true;
      return repartirVisibles();
    });
    await p.waitForTimeout(300);
    ok('la petición salió', enviados.length >= 1, 'no salió ninguna');
    if (enviados.length) {
      const e = enviados[enviados.length - 1];
      ok('con los 6 pedidos de la lista', (e.params.ids || []).length === 6,
         JSON.stringify((e.params.ids || []).length));
      ok('y a quién', e.params.a === 'e2', e.params.a);
      ok('y de qué tienda', e.params.tienda === 'ec', e.params.tienda);
    }

    console.log('\n── Asignar uno solo, desde su detalle ──');
    await p.evaluate(() => { PED_SEL_ID = 'd0'; pintarPanelPedido(); });
    await p.waitForTimeout(120);
    ok('el detalle trae un selector, no solo texto',
       (await p.$$('#pd-gestora')).length === 1);
    const antes = enviados.length;
    await p.selectOption('#pd-gestora', 'e2');
    await p.waitForTimeout(300);
    ok('elegir dispara la asignación', enviados.length > antes);
    if (enviados.length > antes) {
      const e = enviados[enviados.length - 1];
      ok('de ese pedido solo', JSON.stringify(e.params.ids) === '["d0"]',
         JSON.stringify(e.params.ids));
    }

    console.log('\n── Un nombre que no es de nadie se DENUNCIA ──');
    await p.evaluate(() => {
      REPARTO.aNadieConocido = [{ nombre: 'Karen (la de antes)', pedidos: 12 }];
      pintarReparto();
    });
    await p.waitForTimeout(120);
    const rep2 = (await p.textContent('#ped-reparto')).replace(/\s+/g, ' ');
    ok('lo nombra', /Karen \(la de antes\)/.test(rep2), rep2.slice(0, 300));
    ok('y dice lo que significa', /Nadie los est[áa] viendo/.test(rep2));

    console.log('\n── Sin equipo todavía, lo explica ──');
    await p.evaluate(() => {
      ASIGNABLES = [{ id: 'e1', nombre: 'Manuela Ramírez', rol: 'dueno',
                      total: 0, abiertos: 0 }];
      pintarReparto();
    });
    await p.waitForTimeout(120);
    const rep3 = (await p.textContent('#ped-reparto')).replace(/\s+/g, ' ');
    ok('dice que agregue a alguien primero',
       /Permisos/.test(rep3) && /entra y no ve nada/.test(rep3), rep3.slice(0, 300));

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

    ok('cuenta sus 3 asignados', /Asignados\s*3/.test(som), som.slice(0, 250));
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
    ok('dice que no tiene nada asignado',
       /No tiene ning[úu]n pedido asignado/.test(som2), som2.slice(0, 300));
    ok('y lo que eso significa para ella',
       /ve la pantalla vac[ií]a/.test(som2));
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
