/**
 * El inventario en la pantalla: abrir la ficha del TAG RECEDE y editarla.
 *
 * Ella lleva semanas diciendo que no puede. El servidor guarda bien —eso
 * ya está probado en pruebas/inventario.js— así que si algo se rompe,
 * se rompe aquí: al abrir la ficha, al leer los campos o al mandarlos.
 *
 * Esta prueba hace lo que haría ella: entra a Inventario, toca el
 * producto, escribe el stock, los precios y LA LANDING, y le da Guardar.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

/** La hoja al día, con todas las columnas. */
const COLUMNAS = ['id', 'sku', 'producto', 'tienda', 'fuente', 'origen', 'categoria',
  'proveedor', 'landing', 'stock', 'costo_unitario', 'precio', 'precio_2', 'precio_3',
  'minimo', 'resp_1', 'resp_2', 'resp_3', 'resp_4', 'dias_cobertura', 'ultimo_conteo',
  'nota', 'activo', 'actualizado_en', 'actualizado_por'];

const PRODUCTOS = [
  { clave: 'tag recede', nombre: 'TAG RECEDE', sku: 'TR-001',
    pedidos: 48, entregados: 31, devueltos: 6, cancelados: 2, unidades: 62,
    unidadesEntregadas: 40, ventas: 1612, ticket: 52, entrega: 83.8,
    primera: '2026-07-01', ultima: '2026-09-22',
    ritmo: 1.2, coberturaDias: 33, ventanaDias: 83, margen: 46, falta: [],
    ficha: { id: 'inv-tag', nombre: 'TAG RECEDE', sku: 'TR-001', stock: 40,
             minimo: 10, costo: 6, precio: 26, precio2: 0, precio3: 0,
             categoria: 'estrella', proveedor: 'Laboratorio', landing: '',
             respuestas: ['', '', '', ''], nota: '', ultimoConteo: '2026-09-01' } },
  { clave: 'dr melaxin', nombre: 'DR MELAXIN', sku: '',
    pedidos: 12, entregados: 7, devueltos: 1, cancelados: 0, unidades: 12,
    unidadesEntregadas: 7, ventas: 413, ticket: 59, entrega: 87.5,
    primera: '2026-08-10', ultima: '2026-09-20',
    ritmo: null, coberturaDias: null, ventanaDias: 41, margen: null,
    falta: ['sku'], ficha: null },
];

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  for (const ancho of [1200, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1300 } });
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
      /**
       * `#app` nace oculto y solo se muestra al entrar de verdad. Sin
       * esto la prueba mide una pantalla invisible: los campos existen,
       * tienen altura cero, y Playwright dice «no se puede escribir» —
       * que parece el fallo de ella y no lo es.
       */
      document.getElementById('app').style.display = 'flex';
      (0, eval)('CONECTADO = true; ST = "ec"; ROL = "dueno";');
    });

    const enviados = [];
    await p.exposeFunction('__anotar', (x) => { enviados.push(JSON.parse(x)); });
    await p.evaluate(({ prods, cols }) => {
      window.api = async (accion, params) => {
        if (accion === 'productos') {
          return { ok: true, tienda: 'ec', productos: JSON.parse(JSON.stringify(prods)),
                   preguntas: ['¿Para qué sirve?', '¿Cómo se usa?',
                               '¿En cuánto tiempo se ven resultados?',
                               '¿Tiene contraindicaciones?'],
                   columnas: cols, modalidad: 'marca_propia', moneda: 'USD' };
        }
        if (accion === 'escribir' || accion === 'crear') {
          window.__anotar(JSON.stringify({ accion: accion, params: params }));
          return { ok: true, escritos: Object.keys(params.campos || params.datos || {}),
                   rechazados: [] };
        }
        return { ok: true };
      };
      return cargarProductos();
    }, { prods: PRODUCTOS, cols: COLUMNAS });
    await p.waitForTimeout(150);
    await p.evaluate(() => go('inventario', document.querySelector('[data-v=inventario]')));
    await p.waitForTimeout(80);

    console.log('\n── La lista' + A + ' ──');
    const inv = await p.textContent('#v-inventario');
    ok('aparece el TAG RECEDE' + A, /TAG RECEDE/.test(inv), inv.slice(0, 300));
    ok('y el que no tiene ficha, marcado' + A,
       /DR MELAXIN/.test(inv) && /Falta ficha/.test(inv));
    ok('con el botón de agregar uno nuevo' + A,
       /Producto que todavía no has vendido/.test(inv));

    console.log('\n── Abrir la ficha' + A + ' ──');
    await p.evaluate(() => {
      const i = INV_VISTA.findIndex(x => x.nombre === 'TAG RECEDE');
      abrirFicha(i);
    });
    await p.waitForTimeout(100);
    ok('se abre el editor' + A, (await p.$$('#v-inventario .inv-dtl')).length === 1);
    /**
     * Los campos que ella dice que no puede montar. Si alguno no existe,
     * no es que «no deje editar»: es que no hay dónde escribir.
     */
    for (const [id, nombre] of [['fi-producto', 'nombre'], ['fi-stock', 'stock'],
                                ['fi-costo_unitario', 'costo'], ['fi-precio', 'precio'],
                                ['fi-precio_2', 'precio de 2'], ['fi-precio_3', 'precio de 3'],
                                ['fi-minimo', 'mínimo'], ['fi-proveedor', 'proveedor'],
                                ['fi-landing', 'LANDING'], ['fi-nota', 'nota'],
                                ['fi-resp1', 'respuesta 1'], ['fi-resp4', 'respuesta 4']]) {
      ok('hay campo para ' + nombre + A, (await p.$$('#' + id)).length === 1);
    }
    ok('y el nombre viene cargado' + A,
       (await p.inputValue('#fi-producto')) === 'TAG RECEDE');
    ok('y el stock también' + A, (await p.inputValue('#fi-stock')) === '40');

    console.log('\n── Escribir y guardar' + A + ' ──');
    await p.fill('#fi-stock', '38');
    await p.fill('#fi-precio_2', '46');
    await p.fill('#fi-precio_3', '63');
    await p.fill('#fi-landing', 'https://nutrea.co/tag-recede');
    await p.fill('#fi-resp1', 'Para las manchas');
    await p.evaluate(() => {
      const c = [...document.querySelectorAll('#fi-cats .nst')]
        .filter(x => x.dataset.cat === 'estrella')[0];
      if (c && !c.classList.contains('on')) c.click();
    });
    await p.evaluate(() => guardarFicha());
    await p.waitForTimeout(200);

    ok('se manda algo al servidor' + A, enviados.length > 0,
       JSON.stringify(enviados));
    const env = enviados[enviados.length - 1] || {};
    const campos = (env.params || {}).campos || (env.params || {}).datos || {};
    ok('y es una edición, no una ficha nueva' + A, env.accion === 'escribir',
       env.accion);
    ok('con el id de la ficha' + A, (env.params || {}).id === 'inv-tag');
    ok('el stock nuevo va' + A, String(campos.stock) === '38', JSON.stringify(campos));
    ok('los precios del combo van' + A,
       String(campos.precio_2) === '46' && String(campos.precio_3) === '63',
       JSON.stringify(campos));
    /** El que ella nombró por su nombre. */
    ok('LA LANDING VA' + A,
       campos.landing === 'https://nutrea.co/tag-recede', JSON.stringify(campos));
    ok('la respuesta va' + A, campos.resp_1 === 'Para las manchas');
    ok('y la categoría' + A, campos.categoria === 'estrella', JSON.stringify(campos));

    console.log('\n── Si la hoja está sin migrar' + A + ' ──');
    /**
     * El caso de una cuenta vieja: la hoja no tiene las columnas nuevas.
     * La pantalla no puede quedarse callada ni fingir que guardó.
     */
    await p.evaluate(() => {
      window.api = async (accion) => {
        if (accion === 'productos') {
          return { ok: true, productos: [], preguntas: [],
                   columnas: ['id', 'sku', 'producto', 'tienda', 'stock', 'precio'],
                   modalidad: 'marca_propia', moneda: 'USD' };
        }
        return { ok: true };
      };
      return cargarProductos();
    });
    await p.waitForTimeout(100);
    await p.evaluate(() => abrirFicha('nuevo'));
    await p.waitForTimeout(100);
    const aviso = await p.textContent('#v-inventario');
    ok('avisa qué columnas faltan' + A,
       /todavía no tiene/.test(aviso) && /landing/.test(aviso), aviso.slice(0, 400));
    ok('y dice cómo arreglarlo' + A, /bootstrapTodo/.test(aviso));
    /**
     * ── EL CAMPO QUE MENTÍA ──
     *
     * Antes la landing se pintaba igual aunque la hoja no tuviera la
     * columna: ella la escribía, le daba guardar, leía «Ficha
     * actualizada» y no se guardaba nada. Un campo que se deja escribir
     * y descarta lo escrito es peor que uno que no está.
     */
    ok('la landing queda DESACTIVADA, no escribible' + A,
       await p.isDisabled('#fi-landing'));
    ok('y lo dice en el propio campo, no solo arriba' + A,
       /LANDING[\s\S]{0,60}no se puede guardar/.test(aviso), aviso.slice(0, 600));
    ok('el stock, que sí cabe, sigue escribible' + A,
       !(await p.isDisabled('#fi-stock')));
    ok('las respuestas también se desactivan' + A,
       await p.isDisabled('#fi-resp1'));

    // Y guardar NO puede decir que quedó todo bien.
    await p.evaluate(() => {
      window.api = async (accion, params) => {
        if (accion === 'crear') return { ok: true, id: 'nuevo-1' };
        if (accion === 'productos') {
          return { ok: true, productos: [], preguntas: [],
                   columnas: ['id', 'sku', 'producto', 'tienda', 'stock', 'precio'],
                   modalidad: 'marca_propia', moneda: 'USD' };
        }
        return { ok: true };
      };
      document.getElementById('fi-producto').value = 'PRODUCTO NUEVO';
      return guardarFicha();
    });
    await p.waitForTimeout(150);
    ok('al guardar NO dice que quedó todo, dice qué falta' + A,
       /Se guardó lo que cabe/.test(await p.textContent('#v-inventario')),
       (await p.textContent('#v-inventario')).slice(0, 500));

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
