/**
 * CAMBIAR DE TIENDA SIN QUE SE MEZCLE NADA.
 *
 * ┌─ LO QUE PASÓ ──────────────────────────────────────────────┐
 * │                                                            │
 * │ «creo que está combinando datos porque me aparecen          │
 * │  productos de Guatemala en Ecuador (…) sigue apareciéndome  │
 * │  el cuadro de los trabajadores que no tengo».               │
 * │                                                            │
 * │ El servidor no combinaba nada: cada respuesta venía de su   │
 * │ tienda. Lo que se quedaba era la PANTALLA — y no como yo     │
 * │ creí al principio.                                          │
 * │                                                            │
 * │ Mi primera explicación fue que diecisiete variables         │
 * │ guardaban cosas de una tienda y `cambiarTienda` limpiaba    │
 * │ cinco. Lo arreglé, escribí esta prueba, y al desactivar el  │
 * │ arreglo a propósito la prueba SEGUÍA EN VERDE. O sea que    │
 * │ mi explicación no era la causa.                             │
 * │                                                            │
 * │ La causa, al mirar la pantalla de Ecuador en vez de razonar │
 * │ sobre ella, era otra: DOS PINTORES SE CALLABAN CUANDO LA    │
 * │ RESPUESTA VENÍA VACÍA.                                      │
 * │                                                            │
 * │     if (!r.ok || !r.filas.length) return false;            │
 * │                                                            │
 * │ Cero pedidos y cero novedades se trataban como «no hay      │
 * │ respuesta», así que no se pintaba nada… y quedaba el HTML   │
 * │ DE LA MAQUETA: seis pedidos con nombres ecuatorianos        │
 * │ inventados, dos novedades con su courier, «3 gestoras ·     │
 * │ agosto 2026», «VENTAS $ 68.420.000», «HOY $ 1.240».         │
 * │                                                            │
 * │ Con el cartel verde «Datos reales de tu hoja» encima.       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ QUÉ AFIRMA ───────────────────────────────────────────────┐
 * │                                                            │
 * │ Dos cosas, y ninguna enumera variables —la lista va a       │
 * │ volver a crecer y una prueba que la copia se queda corta    │
 * │ igual que se quedó el código:                               │
 * │                                                            │
 * │ 1. Al pasar a una tienda vacía no queda nada de la          │
 * │    anterior, se llame como se llame la variable.            │
 * │ 2. Y no aparece NADA DE LA MAQUETA. Los nombres de esos     │
 * │    datos de muestra solo existen escritos a mano en el      │
 * │    HTML: si uno sale en pantalla con la sesión abierta, es  │
 * │    que un pintor se calló.                                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** Guatemala tiene de todo. Ecuador está recién abierta y no tiene nada. */
const GT = {
  ok: true, tienda: 'gt',
  datos: { ventas: 9999, entregados: 7, equilibrio: { hay: false }, detalleFijos: [],
           productos: {}, productosDetalle: {} },
  filas: [{ id: 'pgt1', tienda: 'gt', cliente: 'CLIENTA DE GUATEMALA', valor: 30,
            producto: 'PRODUCTO GUATEMALTECO', estado_nova: 'novedad', nota: '' }],
  total: 1,
  /**
   * Los campos son los que arma `apiProductos`, no los de la hoja: la
   * pantalla lee `nombre`, y mi primera versión puso `producto` —el nombre
   * de la COLUMNA— así que el catálogo decía «1 en total» y la fila salía
   * en blanco. La prueba pasaba a medias por culpa del arnés, no del
   * producto.
   */
  productos: [{ nombre: 'PRODUCTO GUATEMALTECO', sku: 'GT-1',
                pedidos: 6, entregados: 5, devueltos: 1, cancelados: 0,
                ventas: 150, unidades: 6, unidadesEntregadas: 5,
                ticket: 30, margen: 10, costoUnitario: 20, falta: [],
                coberturaDias: 12,
                ficha: { id: 'i1', nombre: 'PRODUCTO GUATEMALTECO', sku: 'GT-1',
                         stock: 40, minimo: 5, costo: 20, precio: 30,
                         categoria: 'ganador' } }],
  personas: [{ id: 'e9', nombre: 'TRABAJADORA DE GUATEMALA', correo: 'g@x.com',
               rol: 'gestora', tienda: 'gt', estado: 'activo' }],
  casos: [], cas: [], candidatos: [], alarmas: [], notas: {}, accesos: {},
  configuradas: [], umbrales: {}, ajustes: {}, gente: [], grupos: [], hilos: {},
  mensajes: [], vacio: true, historial: [], sesion: {},
};
const EC = {
  ok: true, tienda: 'ec',
  datos: { ventas: 0, entregados: 0, equilibrio: { hay: false }, detalleFijos: [],
           productos: {}, productosDetalle: {} },
  filas: [], total: 0, productos: [], personas: [],
  casos: [], cas: [], candidatos: [], alarmas: [], notas: {}, accesos: {},
  configuradas: [], umbrales: {}, ajustes: {}, gente: [], grupos: [], hilos: {},
  mensajes: [], vacio: true, historial: [], sesion: {},
};
const fs = require('fs');
const GS = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');
const bl = GS.slice(GS.indexOf('const ARRANQUE_EMP = ['));
const PARTES = (bl.slice(0, bl.indexOf('\n];')).match(/^\s*\['([a-z_]+)'/gm) || [])
  .map(l => l.replace(/^\s*\['/, '').replace("'", ''));

/** Lo que solo puede venir de Guatemala. */
const DE_GUATEMALA = /GUATEMALTECO|CLIENTA DE GUATEMALA|TRABAJADORA DE GUATEMALA/;

/**
 * Lo que solo puede venir de la maqueta.
 *
 * Ninguno de estos textos puede salir del servidor: están escritos a
 * mano en el HTML para que la pantalla se pueda ver antes de entrar. Con
 * la sesión abierta, cualquiera de ellos en pantalla es un pintor que se
 * calló.
 */
const DE_LA_MAQUETA = [
  ['un pedido inventado',      /Fausto Riofr|Monserrath Cando|Melany Herrera|Jonathan Piedra/],
  ['un producto inventado',    /TAG RECEDE|EVILGOODS|DR MELAXIN|TRULY ROSA/],
  ['una novedad inventada',    /Servientrega|Laar Courier|Rechazo en puerta/],
  ['ventas inventadas',        /68\.420\.000/],
  ['recaudo del día inventado',/Mejor día: miércoles|11% por encima del promedio|1\.240/],
  ['gestoras que no existen',  /3 gestoras · agosto 2026/],
];

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage({ viewport: { width: 1400, height: 1100 } });
  const errs = [];
  p.on('pageerror', e => errs.push(e.message));
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');

  await p.evaluate(({ GT, EC, PARTES }) => {
    window.__respuesta = (tienda) => {
      const base = tienda === 'gt' ? GT : EC;
      const partes = {};
      PARTES.forEach(k => partes[k] = base);
      return { ok: true, partes, fallaron: [], ms: 100 };
    };
    window.fetch = async (u, o) => {
      const c = JSON.parse(o.body);
      const t = c.tienda || (c.params || {}).tienda || ST;
      if (c.accion === 'arranque') {
        return { text: async () => JSON.stringify(window.__respuesta(t)) };
      }
      return { text: async () => JSON.stringify(t === 'gt' ? GT : EC) };
    };
    document.getElementById('login').style.display = 'none';
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO=true; TOKEN="t"; ST="gt"; ROL="dueno";');
    document.documentElement.setAttribute('data-rol', 'dueno');
    (0, eval)('SESION=' + JSON.stringify({ nombre: 'Manuela', rol: 'dueno',
      correo: 'm@n.com', tiendas: ['gt', 'ec'], permisos: [], modulos: ['empresarial'],
      fichas: [{ id: 'gt', nombre: 'Nutrea GT', moneda: 'GTQ', pais: 'Guatemala' },
               { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD', pais: 'Ecuador' }] }));
    pintarTiendasPermitidas(['gt', 'ec'],
      [{ id: 'gt', nombre: 'Nutrea GT', moneda: 'GTQ', pais: 'Guatemala' },
       { id: 'ec', nombre: 'Nutrea EC', moneda: 'USD', pais: 'Ecuador' }]);
  }, { GT, EC, PARTES });
  await p.waitForTimeout(400);

  console.log('\n══ 1 · EN GUATEMALA SE VE GUATEMALA ══');
  await p.evaluate(() => { setStore('gt'); return cargarReales(); });
  await p.waitForTimeout(1200);
  const vistas = ['pedidos', 'productos', 'inventario', 'gestoras', 'hoy',
                  'novedades', 'dinero'];
  const leer = async () => {
    const out = {};
    for (const v of vistas) {
      await p.evaluate((v) => go(v, null), v);
      await p.waitForTimeout(140);
      out[v] = await p.evaluate((v) => {
        const el = document.getElementById('v-' + v);
        return el ? (el.innerText || '').replace(/\s+/g, ' ') : '';
      }, v);
    }
    return out;
  };
  let txt = await leer();
  ok('los productos de Guatemala se ven', DE_GUATEMALA.test(txt.productos),
     txt.productos.slice(0, 120));
  ok('y su gente también', DE_GUATEMALA.test(txt.gestoras), txt.gestoras.slice(0, 120));

  console.log('\n══ 2 · AL PASAR A ECUADOR, NO QUEDA NADA DE GUATEMALA ══');
  /**
   * Ecuador está recién abierta: el servidor contesta vacío en todo. Si
   * la pantalla no olvidó lo anterior, es justo cuando se nota — no hay
   * datos nuevos que pisen los viejos.
   */
  await p.evaluate(() => cambiarTienda('ec'));
  await p.waitForTimeout(1600);
  txt = await leer();
  const sucias = Object.keys(txt).filter((v) => DE_GUATEMALA.test(txt[v]));
  igual('ninguna pantalla conserva datos de Guatemala', [], sucias);
  if (sucias.length) {
    sucias.forEach((v) => console.log('         ' + v + ' → ' +
      txt[v].slice(0, 140)));
  }
  igual('y la tienda abierta es Ecuador', 'ec', await p.evaluate(() => ST));

  console.log('\n══ 3 · Y EN UNA TIENDA VACÍA NO SE INVENTA NADA ══');
  /**
   * Esta es la que de verdad encontró el error. Las de arriba pasaban
   * con el arreglo desactivado; esta no.
   */
  DE_LA_MAQUETA.forEach(([qué, re]) => {
    const dónde = Object.keys(txt).filter((v) => re.test(txt[v]));
    ok('no se ve ' + qué, !dónde.length,
       dónde.length ? 'en ' + dónde.join(', ') + ' → ' +
         (txt[dónde[0]].match(re) || [''])[0] : undefined);
  });
  /**
   * Y el cartel no puede prometer más de lo que hay: si algo no cargó,
   * lo dice con nombre propio en vez de desmentir la pantalla entera.
   */
  const cartel = await p.evaluate(() => {
    const b = document.getElementById('origen-datos');
    return b ? b.textContent : '(no hay cartel)';
  });
  ok('el cartel de origen no dice EJEMPLO con datos reales',
     !/EJEMPLO/.test(cartel), cartel);

  console.log("\n══ 4 · Y AL VOLVER, GUATEMALA ESTÁ ENTERA ══");
  /**
   * Olvidar de más también es un error: si al volver no se recarga,
   * queda una tienda en blanco que sí tiene datos.
   */
  await p.evaluate(() => cambiarTienda('gt'));
  await p.waitForTimeout(1600);
  txt = await leer();
  ok('los productos volvieron', DE_GUATEMALA.test(txt.productos),
     txt.productos.slice(0, 120));
  ok('y la gente también', DE_GUATEMALA.test(txt.gestoras), txt.gestoras.slice(0, 120));

  console.log('\n══ 5 · UN PINTOR QUE REVIENTA NO TUMBA A LOS DEMÁS ══');
  /**
   * Los dieciocho cargadores iban en un `Promise.all` desnudo. Uno que
   * se cayera —me pasó con un `nombreMes(undefined)`— hacía que el
   * `catch` de `cargarReales` sellara la pantalla COMPLETA como «Datos
   * de EJEMPLO», y con ella volvían las tablas de la maqueta.
   *
   * Aquí se rompe uno a propósito y se comprueba que el resto aguanta y
   * que el aviso lo dice por su nombre.
   */
  await p.evaluate(() => {
    (0, eval)('cargarRecuento = async function () { ' +
      'throw new Error("reventé a propósito"); };');
  });
  await p.evaluate(() => cargarReales());
  await p.waitForTimeout(1400);
  txt = await leer();
  ok('las demás pantallas siguen con datos de verdad',
     DE_GUATEMALA.test(txt.productos) && DE_GUATEMALA.test(txt.gestoras),
     'productos → ' + txt.productos.slice(0, 100));
  DE_LA_MAQUETA.forEach(([qué, re]) => {
    const dónde = Object.keys(txt).filter((v) => re.test(txt[v]));
    ok('tampoco vuelve ' + qué, !dónde.length,
       dónde.length ? 'en ' + dónde.join(', ') : undefined);
  });
  const aviso = await p.evaluate(() => {
    const b = document.getElementById('origen-datos');
    return b ? b.textContent : '(no hay cartel)';
  });
  ok('el aviso nombra lo que no cargó', /recuento/.test(aviso), aviso);
  ok('y no desmiente el resto de la pantalla', !/EJEMPLO/.test(aviso), aviso);

  igual('sin errores de consola', [], errs);
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
