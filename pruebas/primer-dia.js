/**
 * EL PRIMER DÍA DE UNA CLIENTA NUEVA.
 *
 * ┌─ POR QUÉ EXISTE ───────────────────────────────────────────┐
 * │                                                            │
 * │ Sara entró por primera vez y no pudo pasar de la puerta.    │
 * │ Al arreglar la puerta pregunté lo obvio: ¿y una vez         │
 * │ adentro? Nadie había mirado nunca una cuenta recién nacida  │
 * │ —una tienda, cero pedidos, cero equipo— porque la única     │
 * │ cuenta que existía era la nuestra, llena de datos.          │
 * │                                                            │
 * │ Había dos cosas rotas, y ninguna se veía como un error:     │
 * │                                                            │
 * │  · La Auditoría reventaba. El servidor contesta             │
 * │    `{vacio:true}` cuando todavía no hay novedades, y la     │
 * │    pantalla leía `AUD.recuento.revisados` a pelo. Y como    │
 * │    corre dentro del `Promise.all` del arranque, se llevaba  │
 * │    por delante todo lo que faltara.                         │
 * │                                                            │
 * │  · Y entonces el Hoy quedaba con «Datos de EJEMPLO · sin    │
 * │    conexión» ENCIMA de números que sí eran reales. Lo peor  │
 * │    de los dos mundos: enseña a desconfiar de lo que sí se   │
 * │    puede creer, y manda a buscar el problema al router.     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE HACE DISTINTA A ESTA PRUEBA ───────────────────────┐
 * │                                                            │
 * │ La respuesta NO está escrita a mano: la calcula el servidor │
 * │ de verdad sobre una hoja recién creada, y la pantalla la    │
 * │ recibe tal cual.                                            │
 * │                                                            │
 * │ Mi primer intento sí inventaba la respuesta, y me dio tres  │
 * │ fallos que no existían (campos que mi simulador no mandaba  │
 * │ y el servidor sí) mientras tapaba el que sí era real. Una   │
 * │ prueba que se inventa lo que el servidor contesta prueba mi │
 * │ imaginación, no el producto.                                │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs','utf8');
const LIBROS = { cen:{}, s:{}, a:{}, emp:{} };
function libro(id){ const h = LIBROS[id]||{}; return { getSheetByName:(n)=>{
  const m=h[n]; if(!m) return null;
  return { getName:()=>n, getLastRow:()=>m.length, getLastColumn:()=>(m[0]?m[0].length:0),
    getDataRange:()=>({getValues:()=>m.map(f=>f.slice())}),
    appendRow:f=>m.push(f.slice()), deleteRow:i=>m.splice(i-1,1),
    getRange:(f,c,nf,nc)=>({ getValues:()=>{const o=[];for(let i=0;i<(nf||1);i++){
      const fl=m[f-1+i]||[];o.push(fl.slice(c-1,c-1+(nc||fl.length)));}return o;},
      setValues:()=>{}, setValue:()=>{} }) };
}};}
const PROPS={ID_EMPRESARIAL:'emp',ID_CENTRAL:'cen',ID_SOUL:'s',ID_ACADEMY:'a'};
global.PropertiesService={getScriptProperties:()=>({getProperty:k=>PROPS[k]||'',
  getProperties:()=>PROPS,setProperty:()=>{},deleteProperty:()=>{}})};
global.SpreadsheetApp={openById:id=>libro(id),flush:()=>{}};
global.Logger={log:()=>{}};
global.ScriptApp={getProjectTriggers:()=>[],EventType:{CLOCK:'CLOCK'},WeekDay:{MONDAY:'MONDAY'}};
global.Session={getScriptTimeZone:()=>'UTC',getActiveUser:()=>({getEmail:()=>''}),
  getEffectiveUser:()=>({getEmail:()=>''})};
global.LockService={getScriptLock:()=>({tryLock:()=>true,releaseLock:()=>{}})};
global.CacheService={getScriptCache:()=>({get:()=>null,put:()=>{},remove:()=>{}})};
global.MailApp={sendEmail:()=>{}};
global.UrlFetchApp={fetch:()=>({getContentText:()=>'{}'})};
global.Utilities={sleep:()=>{},getUuid:()=>'u1',formatDate:(d,tz,pat)=>{
  const iso=new Date(d).toISOString();
  if(pat==='yyyy-MM-dd')return iso.slice(0,10); if(pat==='yyyy-MM')return iso.slice(0,7);
  if(/HH:mm/.test(pat))return iso.slice(0,19).replace('T',' '); return iso;}};
(0,eval)(src + '\n;globalThis.__F={apiArranque,libroOlvidar_,ESQUEMA_EMPRESARIAL,ARRANQUE_EMP,publico};');
const F=globalThis.__F, E=F.ESQUEMA_EMPRESARIAL;
const f=(c,o)=>c.map(k=>o[k]!==undefined?o[k]:'');

// La cuenta de Sara el primer día: la plantilla completa, todo vacío.
const emp={};
Object.keys(E).forEach(k=>{ emp[k]=[E[k]]; });
emp.Movimientos=[['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo']];
emp.Tiendas.push(f(E.Tiendas,{id:'qfx',nombre:'QuickFix',moneda:'COP',pais:'Colombia',
                              zona_horaria:'America/Bogota',estado:'activa'}));
emp.Equipo.push(f(E.Equipo,{id:'e1',nombre:'Sara Ramirez',correo:'sara@quickfix.com',
                            rol:'dueno',tienda:'*',estado:'activo'}));
LIBROS.emp=emp;
F.libroOlvidar_();
const S={email:'sara@quickfix.com',nombre:'Sara Ramirez',rol:'dueno',
         sheetId:'emp',tiendas:['qfx'],permisos:[],modulos:['empresarial']};

const ARR = F.apiArranque(S, { tienda: 'qfx',
  mes: new Date().toISOString().slice(0, 7), limite: 300 });

const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

/** Nombres y cifras que solo pueden venir de NUESTRA cuenta. */
const NUESTRO = /Nutrea|Camila|Daniela|Colágeno|Melaxin|Tag Recede|Truly|Evil Goods/;
const CIFRAS_EJEMPLO = /\$ ?\d{1,3}\.\d{3}\.\d{3}|Q ?\d{2}\.\d{3}/;

(async () => {
  console.log('\n══ 0 · EL SERVIDOR CONTESTA SIN CAERSE ══');
  igual('ninguna sección falla', [], ARR.fallaron);
  ok('trae todas las secciones', Object.keys(ARR.partes).length > 15,
     String(Object.keys(ARR.partes).length));
  /**
   * Esta es la forma exacta que reventaba la pantalla, y hay que
   * conservarla: la auditoría responde BIEN —`ok:true`— pero SIN
   * `recuento`, porque todavía no hay novedades que recontar. La
   * pantalla leía `AUD.recuento.revisados` a pelo.
   *
   * (Mi primera versión de esta aserción decía `vacio:true`, que me lo
   * inventé. El servidor manda `porque`. Comprobar lo que imaginé en
   * vez de lo que pasa es cómo no se encuentra un error.)
   */
  const A = ARR.partes.auditoria_casos;
  ok('la auditoría contesta bien', A.ok === true, JSON.stringify(A).slice(0, 120));
  ok('pero sin recuento, porque no hay nada que recontar',
     A.recuento === undefined, JSON.stringify(A.recuento));
  ok('y explica por qué', /novedades/i.test(A.porque || ''), A.porque);

  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage({ viewport: { width: 1400, height: 1100 } });
  const errs = [];
  p.on('pageerror', e => errs.push('PAGE: ' + e.message));
  p.on('console', m => { if (m.type() === 'error' &&
    !/ERR_CERT|Failed to load resource/.test(m.text())) errs.push('CONSOLE: ' + m.text()); });
  await p.goto('file:///home/claude/repo/empresarial.html');
  await p.waitForLoadState('load');

  await p.evaluate(({ ARR }) => {
    window.__pedidas = [];
    window.fetch = async (u, o) => {
      const c = JSON.parse(o.body);
      window.__pedidas.push(c.accion);
      if (c.accion === 'arranque') return { text: async () => JSON.stringify(ARR) };
      return { text: async () => JSON.stringify({ ok: false, error: 'no simulado' }) };
    };
    (0, eval)('CONECTADO=true; TOKEN="t"; ROL="dueno";');
    document.documentElement.setAttribute('data-rol', 'dueno');
    (0, eval)('SESION=' + JSON.stringify({ nombre: 'Sara Ramirez', rol: 'dueno',
      correo: 'sara@quickfix.com', tiendas: ['qfx'], permisos: [], modulos: ['empresarial'],
      fichas: [{ id: 'qfx', nombre: 'QuickFix', moneda: 'COP', pais: 'Colombia' }] }));
    document.getElementById('em').value = 'sara@quickfix.com';
    window.__origen = [];
    (0, eval)('window.__real = marcarOrigen; marcarOrigen = function (o, m) {'
      + ' window.__origen.push(o); return window.__real(o, m); };');
    pintarTiendasPermitidas(['qfx'],
      [{ id: 'qfx', nombre: 'QuickFix', moneda: 'COP', pais: 'Colombia' }]);
  }, { ARR });
  await p.waitForTimeout(2200);

  console.log('\n══ 1 · ENTRA, Y NADA SE CAE ══');
  igual('sin un solo error', [], errs);
  ok('está dentro', await p.evaluate(() =>
     document.getElementById('app').classList.contains('on')));
  igual('un solo viaje', ['arranque'], await p.evaluate(() => window.__pedidas));

  console.log('\n══ 2 · Y SUS DATOS SE MARCAN COMO REALES ══');
  /**
   * El sello no puede acabar en «EJEMPLO». Lo hacía: un error de
   * JavaScript en cualquier pintor caía en el catch de `cargarReales`,
   * que decía «sin conexión» sobre datos que sí habían llegado.
   */
  const origen = await p.evaluate(() => window.__origen);
  ok('el sello acaba en REAL, no en ejemplo', origen[origen.length - 1] === true,
     JSON.stringify(origen));
  const banner = await p.evaluate(() =>
    (document.getElementById('origen-datos') || {}).textContent || '');
  ok('y el aviso lo dice', /reales/i.test(banner), banner);
  ok('sin culpar a la conexión', !/sin conexión/i.test(banner), banner);

  console.log('\n══ 3 · NINGUNA PANTALLA LE ENSEÑA LO NUESTRO ══');
  const vistas = ['hoy', 'pedidos', 'novedades', 'oficina', 'productos', 'inventario',
                  'accesos', 'dinero', 'pauta', 'calc', 'cierre', 'gestoras',
                  'auditoria', 'permisos', 'config'];
  for (const v of vistas) {
    await p.evaluate((v) => go(v, null), v);
    await p.waitForTimeout(160);
    const r = await p.evaluate((v) => {
      const el = document.getElementById('v-' + v);
      if (!el) return { falta: true };
      return { txt: (el.innerText || '').replace(/\s+/g, ' ').trim() };
    }, v);
    ok(v + ': existe', !r.falta);
    if (r.falta) continue;
    ok(v + ': sin nombres nuestros', !NUESTRO.test(r.txt), r.txt.slice(0, 120));
    ok(v + ': sin cifras de ejemplo', !CIFRAS_EJEMPLO.test(r.txt), r.txt.slice(0, 120));
  }

  console.log('\n══ 4 · LOS SITIOS DONDE SE ELIGE UNA TIENDA ══');
  /**
   * `STORES` arrancaba con `ec` y `gt` y esta pantalla solo los AGREGABA
   * a los del cliente. Resultado: en Cierre y —peor— en las casillas de
   * PERMISOS, Sara podía asignarle a su gestora una tienda nuestra.
   */
  igual('el catálogo de tiendas es solo el suyo', ['qfx'],
        await p.evaluate(() => Object.keys(STORES)));
  await p.evaluate(() => go('cierre', null));
  await p.waitForTimeout(400);
  igual('el cierre solo ofrece la suya', ['QuickFix'],
        await p.evaluate(() => Array.from(
          document.querySelectorAll('#cm-tienda option')).map(o => o.textContent)));
  await p.evaluate(() => { go('permisos', null); nuevaPersona(); });
  await p.waitForTimeout(400);
  const casillas = await p.evaluate(() => Array.from(
    document.querySelectorAll('#ep-tiendas-lista label')).map(l => l.textContent.trim()));
  igual('y las casillas de permisos también', ['QuickFix'], casillas);

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
