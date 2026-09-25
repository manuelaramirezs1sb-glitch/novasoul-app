/**
 * CUÁNTAS PETICIONES HACE LA PANTALLA AL ABRIRSE.
 *
 * ┌─ QUÉ SE ESTÁ MIDIENDO ─────────────────────────────────────┐
 * │                                                            │
 * │ «se demora mucho en cargar».                                │
 * │                                                            │
 * │ No son los datos: son los VIAJES. Cada petición a Apps       │
 * │ Script arranca un motor de Google, y Google SERIALIZA las    │
 * │ del mismo usuario — las que el navegador manda «a la vez»    │
 * │ se ejecutan una detrás de otra, cada una abriendo la misma   │
 * │ hoja otra vez.                                              │
 * │                                                            │
 * │ Medido en un navegador de verdad: eran DIECINUEVE.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y POR QUÉ ESTA PRUEBA MIDE `fetch` Y NO `api()` ──────────┐
 * │                                                            │
 * │ Mi primera versión remedaba `api()`, que es justo la        │
 * │ función que sirve de la despensa del arranque. Reemplazarla │
 * │ borraba lo que había que medir: contaba 20 y parecía que no │
 * │ había mejorado nada. (Y `ARRANQUE` es un `let`, así que     │
 * │ `window.ARRANQUE` tampoco lo tocaba — la misma trampa de    │
 * │ `SESION` que ya me había mordido antes.)                    │
 * │                                                            │
 * │ Lo que cuesta segundos es lo que sale a la red. Se mide     │
 * │ eso.                                                       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

// `tienda` va dentro a propósito: `cargarAccesos` la usa para saber si
// lo que tiene en memoria es de la tienda abierta. Sin ella volvía a
// pedirlos y la prueba contaba un viaje de más que no existe.
const VACIO = { ok:true, tienda:'ec', datos:{}, filas:[], total:0, notas:{},
                accesos:{}, sesion:{}, personas:[], alarmas:[], productos:[],
                casos:[], configuradas:[], umbrales:{}, ajustes:{} };
/**
 * Las secciones NO se escriben a mano aquí: se sacan del servidor.
 *
 * Las tenía copiadas y se me quedó el chat fuera al agregarlo, así que
 * la prueba «medía» un arranque que no traía el chat y contaba un viaje
 * de más que no existía. Una lista duplicada es una lista que se
 * desincroniza; ésta se lee de ARRANQUE_EMP y no puede.
 */
const fs = require('fs');
const GS = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');
const bloque = GS.slice(GS.indexOf('const ARRANQUE_EMP = ['));
const PARTES = (bloque.slice(0, bloque.indexOf('\n];'))
  .match(/^\s*\['([a-z_]+)'/gm) || []).map(l => l.replace(/^\s*\['/, '').replace("'", ''));

async function abrir(b, conArranque) {
  const p = await b.newPage({ viewport: { width: 1280, height: 1000 } });
  const errores = [];
  p.on('pageerror', e => errores.push(e.message));
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');

  const acciones = [];
  await p.exposeFunction('__log', (a) => acciones.push(a));
  await p.evaluate(({ VACIO, PARTES, conArranque }) => {
    document.getElementById('login').style.display = 'none';
    document.getElementById('app').style.display = 'flex';
    (0, eval)('CONECTADO=true; TOKEN="t"; ST="ec"; ROL="dueno";');
    document.documentElement.setAttribute('data-rol', 'dueno');
    (0, eval)('SESION=' + JSON.stringify({ nombre:'M', rol:'dueno', correo:'m@n.com',
      tiendas:['ec'], permisos:[], modulos:['empresarial'] }));
    // Se remeda la RED, no `api()`: `api` es lo que hay que medir.
    window.fetch = async (url, opt) => {
      const cuerpo = JSON.parse(opt.body);
      window.__log(cuerpo.accion);
      let resp = VACIO;
      if (cuerpo.accion === 'arranque') {
        if (!conArranque) resp = { ok:false, error:'Acción desconocida: arranque' };
        else {
          const partes = {};
          PARTES.forEach(k => partes[k] = VACIO);
          resp = { ok:true, partes, fallaron: [], ms: 900 };
        }
      }
      return { text: async () => JSON.stringify(resp) };
    };
  }, { VACIO, PARTES, conArranque });

  // `cargarAccesos` ya va dentro de `cargarReales`, en el mismo lote.
  await p.evaluate(() => { cargarReales(); initChat(); });
  await p.waitForTimeout(1200);
  return { p, acciones, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  console.log('\n══ 0 · LAS DOS LISTAS DICEN LO MISMO ══');
  /**
   * El servidor decide QUÉ trae el arranque (ARRANQUE_EMP) y la pantalla
   * decide QUÉ sabe servir de la despensa (claveArranque). Si una tiene
   * algo que la otra no, esa sección se trae y se vuelve a pedir: un
   * viaje de más, invisible, por cada vez que alguien entra.
   */
  ok('el servidor declara sus secciones', PARTES.length > 15,
     PARTES.length + ': ' + PARTES.join(' '));
  {
    const p0 = await b.newPage();
    await p0.goto('file:///home/claude/repo/empresarial.html');
    await p0.waitForLoadState('load');
    const huerfanas = await p0.evaluate((PARTES) => {
      return PARTES.filter(function (k) {
        // `pedidos` y `novedades` llegan por `listar`, no por su nombre.
        if (k === 'pedidos') return claveArranque('listar', { entidad: 'Pedidos' }) !== k;
        if (k === 'novedades') return claveArranque('listar', { entidad: 'Novedades' }) !== k;
        return claveArranque(k, {}) !== k;
      });
    }, PARTES);
    igual('la pantalla sabe servir todas', [], huerfanas);
    await p0.close();
  }

  console.log('\n══ 1 · CON EL ARRANQUE: UN SOLO VIAJE ══');
  {
    const { p, acciones, errores } = await abrir(b, true);
    igual('una sola petición al abrir', 1, acciones.length);
    igual('y es el arranque', ['arranque'], acciones);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 2 · SIN EL ARRANQUE, LA PANTALLA SIGUE FUNCIONANDO ══');
  /**
   * Es lo que pasa mientras el código nuevo no esté pegado en Apps
   * Script. El servidor contesta «Acción desconocida» y cada pantalla
   * tiene que salir a pedir lo suyo, como antes. Lento, pero entero.
   *
   * Un arranque que al fallar deja la pantalla en blanco sería peor que
   * no tenerlo.
   */
  {
    const { p, acciones, errores } = await abrir(b, false);
    ok('vuelve a pedir de a una', acciones.length > 15, String(acciones.length));
    ['resumen', 'listar', 'notas', 'accesos', 'alarmas', 'equipo'].forEach(function (a) {
      ok('sigue pidiendo ' + a, acciones.indexOf(a) !== -1, JSON.stringify(acciones));
    });
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  console.log('\n══ 3 · LA DESPENSA SE SIRVE UNA VEZ ══');
  /**
   * La segunda vez que algo pide `listar` es porque acaba de guardar y
   * necesita lo nuevo. Si la despensa lo siguiera sirviendo, la pantalla
   * se quedaría enseñando lo de hace diez minutos — que es peor que ser
   * lenta, porque no se nota.
   */
  {
    const { p, acciones, errores } = await abrir(b, true);
    igual('el arranque, y nada más', 1, acciones.length);
    await p.evaluate(() => cargarPedidosReales());
    await p.waitForTimeout(400);
    igual('volver a pedir pedidos SÍ viaja', 2, acciones.length);
    igual('y va a la red', 'listar', acciones[1]);
    igual('sin errores de consola', [], errores);
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
