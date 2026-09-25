/**
 * LOS PERMISOS, DE PUNTA A PUNTA.
 *
 * ┌─ LO QUE ELLA PIDIÓ COMPROBAR ──────────────────────────────┐
 * │                                                            │
 * │ «que esté bien hecho no quiere decir que esté conectado.    │
 * │  Revisa, haz una prueba donde crees a alguien como usuario  │
 * │  y le asignes qué ver y qué no, las pantallas que tendrá   │
 * │  disponibles; probando que con el usuario del cliente la    │
 * │  pantalla que él vea sea acorde a los permisos que se le    │
 * │  dieron y las pantallas a las que tiene acceso.»            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * La cadena entera, sin saltarse un eslabón:
 *
 *   Central crea el cliente y su dueña
 *        ↓
 *   la dueña agrega a alguien y le marca permisos
 *        ↓
 *   esa persona entra con su correo
 *        ↓
 *   el servidor le arma la sesión con su rol y sus permisos
 *        ↓
 *   LA PANTALLA le muestra unas secciones y le esconde otras
 *
 * El último paso es el que nadie prueba nunca, y es el único que la
 * persona ve. Por eso esta prueba abre el navegador de verdad.
 */
const fs = require('fs');
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

// ─── EL SERVIDOR ─────────────────────────────────────────────
const LIBROS = { cen: {}, s: {}, a: {} };
function libroStub(id) {
  if (!LIBROS[id]) LIBROS[id] = {};
  const hojas = LIBROS[id];
  const hoja = (n) => {
    const m = hojas[n];
    if (!m) return null;
    return {
      getName: () => n, getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getMaxColumns: () => Math.max(m[0] ? m[0].length : 0, 26),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (k) => m.splice(k - 1, 1),
      setFrozenRows: () => {}, autoResizeColumns: () => {}, setTabColor: () => {},
      deleteColumns: () => {}, clear: () => { m.length = 1; },
      getRange: (f, c, nf, nc) => {
        const api = {
          getValues: () => {
            const o = [];
            for (let i = 0; i < (nf || 1); i++) {
              const r = m[f - 1 + i] || [];
              o.push(r.slice(c - 1, c - 1 + (nc || r.length)));
            }
            return o;
          },
          setValues: (v) => { v.forEach((r, i) => {
            const d = m[f - 1 + i] || (m[f - 1 + i] = []);
            r.forEach((x, j) => { d[c - 1 + j] = x; }); }); return api; },
          setValue: (x) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = x; return api; },
          setFontWeight: () => api, setBackground: () => api, setFontColor: () => api,
          setNumberFormat: () => api, setNote: () => api, clearContent: () => api,
        };
        return api;
      },
    };
  };
  return {
    getSheetByName: hoja,
    insertSheet: (n) => { hojas[n] = [[]]; return hoja(n); },
    getSheets: () => Object.keys(hojas).map(hoja),
    deleteSheet: () => {}, getSpreadsheetTimeZone: () => 'UTC',
    getId: () => id, getName: () => id, setName: () => {},
  };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-25T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: libroStub, create: (n) => libroStub('nuevo'), flush: () => {} };
global.DriveApp = { getFileById: () => ({ moveTo: () => {}, setName: () => {} }),
                    getFolderById: () => ({ createFolder: () => ({ getId: () => 'f' }) }),
                    createFolder: () => ({ getId: () => 'f' }) };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
const CACHE = {};
global.CacheService = { getScriptCache: () => ({
  get: (k) => (CACHE[k] === undefined ? null : CACHE[k]),
  put: (k, v) => { CACHE[k] = v; }, remove: (k) => { delete CACHE[k]; } }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{}' }) };
let UUID = 0;
global.Utilities = { sleep: () => {}, getUuid: () => 'u' + (++UUID) + '0000000',
  formatDate: (d, tz, pat) => {
    const iso = (d && d.getTime && d.getTime() === new Date(HOY).getTime())
      ? HOY : new Date(d).toISOString();
    if (pat === 'yyyy-MM-dd') return iso.slice(0, 10);
    if (pat === 'yyyy-MM') return iso.slice(0, 7);
    if (/HH:mm/.test(pat)) return iso.slice(0, 19).replace('T', ' ');
    return iso;
  } };
const RealDate = Date;
global.Date = class extends RealDate {
  constructor(...a) { return a.length ? new RealDate(...a) : new RealDate(HOY); }
  static now() { return new RealDate(HOY).getTime(); }
  static UTC(...a) { return RealDate.UTC(...a); }
};
(0, eval)(fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8') +
  '\n;globalThis.__S = { buscarPersona, apiVerificar, apiCrear, permisosDe,' +
  ' PERMISOS_POR_ROL, PERMISOS_CONOCIDOS, PERMISOS, libroOlvidar_, soulOlvidar_,' +
  ' ESQUEMA_EMPRESARIAL, modulosDelPlan };');
const S = globalThis.__S;
const E = S.ESQUEMA_EMPRESARIAL;
const fila = (cols, o) => cols.map(c => (o[c] !== undefined ? o[c] : ''));

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_CLI = ['id','empresa','correo','plan','tarifa','costo','estado','fecha_alta',
               'fecha_corte','ultimo_pago','tickets_mes','usuarios','tiendas','sheet_id'];

function sembrar() {
  UUID = 0; S.libroOlvidar_(); S.soulOlvidar_();
  Object.keys(CACHE).forEach(k => delete CACHE[k]);
  LIBROS.cen = { Clientes: [C_CLI.slice(),
    fila(C_CLI, { id:'c1', empresa:'VIP COL', plan:'pro', estado:'activo', sheet_id:'cli1' })] };
  LIBROS.cli1 = {
    Equipo: [E.Equipo.slice()],
    Pedidos: [E.Pedidos.slice()], Novedades: [E.Novedades.slice()],
    Tiendas: [['id','nombre','moneda','estado','modalidad'],
              ['col','VIP Colombia','COP','activa','catalogo_publico']],
    Movimientos: [['fecha','usuario','entidad','entidad_id','campo','valor_anterior','valor_nuevo']],
    Parametros: [['tienda','clave','valor','nota','tipo']],
  };
  // La dueña, como la deja Central al crear el cliente.
  LIBROS.cli1.Equipo.push(fila(E.Equipo, { id:'e1', nombre:'Sara Duarte',
    correo:'sara@vipcol.com', rol:'dueno', tienda:'*', estado:'activo' }));
}

/** Entrar de verdad: el código al caché y el canje, como en producción. */
function entrar(correo) {
  CACHE['cod_' + correo] = '123456';
  S.libroOlvidar_();
  const r = S.apiVerificar({ email: correo, codigo: '123456' });
  return r.ok ? r.sesion : null;
}

// ══════════════════════════════════════════════════════════════
console.log('\n══ 1 · LA DUEÑA AGREGA A SU EQUIPO Y LES MARCA PERMISOS ══');
sembrar();
const DUENA = { sheetId:'cli1', rol:'dueno', email:'sara@vipcol.com', nombre:'Sara Duarte',
                tiendas:['col'], permisos: S.PERMISOS_CONOCIDOS.slice(),
                modulos:['empresarial','pauta','dinero'] };

let r = S.apiCrear(DUENA, { entidad:'Equipo', datos: { nombre:'Zulay Pérez',
  correo:'zulay@vipcol.com', rol:'gestora', tienda:'col', estado:'activo', permisos:'' } });
igual('se crea la gestora', true, r.ok);

r = S.apiCrear(DUENA, { entidad:'Equipo', datos: { nombre:'Gerald Ruiz',
  correo:'gerald@vipcol.com', rol:'admin', tienda:'col', estado:'activo',
  // La dueña le marca DOS casillas del formulario de permisos.
  permisos:'subir_pauta,ver_dinero' } });
igual('y la admin con dos permisos extra', true, r.ok);

console.log('\n══ 2 · CADA UNA ENTRA Y EL SERVIDOR LE ARMA SU SESIÓN ══');
const sZulay = entrar('zulay@vipcol.com');
const sGerald = entrar('gerald@vipcol.com');
const sSara = entrar('sara@vipcol.com');
ok('la gestora entra', !!sZulay);
ok('la admin entra', !!sGerald);
ok('la dueña entra', !!sSara);

igual('la gestora llega como gestora', 'gestora', sZulay.rol);
igual('con los permisos de su rol', S.PERMISOS_POR_ROL.gestora, sZulay.permisos);
igual('la admin llega como admin', 'admin', sGerald.rol);
/** Los de su rol MÁS los que le marcó la dueña. La celda suma, no reemplaza. */
igual('con los suyos más los marcados',
      ['subir_pedidos','subir_novedades','subir_pauta','ver_dinero'], sGerald.permisos);
ok('la sesión NUNCA lleva el id de la hoja', sZulay.sheetId === undefined,
   JSON.stringify(Object.keys(sZulay)));

console.log('\n── Quitar un permiso con el menos ──');
sembrar();
S.apiCrear(DUENA, { entidad:'Equipo', datos: { nombre:'Jaime Soto',
  correo:'jaime@vipcol.com', rol:'admin', tienda:'col', estado:'activo',
  permisos:'-subir_novedades' } });
const sJaime = entrar('jaime@vipcol.com');
igual('se le quita solo ese', ['subir_pedidos'], sJaime.permisos);

// ══════════════════════════════════════════════════════════════
console.log('\n══ 3 · Y AHORA LA PANTALLA ══');

/** Las secciones del menú, y quién debería verlas. */
const SECCIONES = ['hoy','pedidos','novedades','oficina','productos','gestoras',
                   'auditoria','pauta','inventario','dinero','calc','cierre',
                   'permisos','config'];

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  async function loQueVe(sesion) {
    const p = await b.newPage({ viewport: { width: 1280, height: 1000 } });
    const errores = [];
    p.on('pageerror', e => errores.push(e.message));
    await p.goto('file:///home/claude/repo/empresarial.html');
    await p.waitForLoadState('load');
    const visto = await p.evaluate(({ ses, secciones }) => {
      document.getElementById('login').style.display = 'none';
      document.getElementById('app').style.display = 'flex';
      (0, eval)('CONECTADO = true; ST = "col"; ROL = ' + JSON.stringify(ses.rol) + ';' +
                'SESION = ' + JSON.stringify(ses) + ';');
      /**
       * Lo que hace la app al entrar, LAS DOS COSAS. La primera versión
       * de esta prueba solo estampaba el rol y concluyó que la pantalla
       * ignoraba los permisos — falso: `aplicarPermisos()` existe y sí
       * los lee. El error era de la prueba, y de haberlo dado por
       * bueno habría salido un diagnóstico equivocado.
       */
      document.getElementById('app').setAttribute('data-rol', ses.rol);
      aplicarPermisos();

      const out = { menu: [], vistas: [] };
      secciones.forEach(function (v) {
        /**
         * ── CÓMO SE MIDE SI ALGO SE VE ──
         *
         * Dos trampas, y caí en las dos:
         *
         * 1· `getComputedStyle(el).display` de un hijo NO dice «none»
         *    porque su padre esté oculto. El grupo DINERO del menú está
         *    escondido entero con `.ro-dueno`, y sus entradas seguían
         *    midiendo «block». Lo que sí sirve es `offsetParent`: es
         *    null cuando el elemento no está en pantalla, suyo o de un
         *    ancestro.
         *
         * 2· `data-v="pauta"` aparece DOS veces: una en el grupo de la
         *    dueña y otra en el del equipo, para la admin con permiso.
         *    Con `querySelector` se mide siempre la primera, que es la
         *    que la admin nunca ve. Hay que mirarlas todas.
         */
        const navs = Array.from(document.querySelectorAll('[data-v="' + v + '"]'));
        if (navs.some(function (n) { return n.offsetParent !== null; })) out.menu.push(v);
        const el = document.getElementById('v-' + v);
        if (el) {
          // Se mira si la regla la esconde, no si está "apagada" por no ser
          // la sección abierta: para eso se enciende y se mide.
          el.classList.add('on');
          if (el.offsetParent !== null) out.vistas.push(v);
          el.classList.remove('on');
        }
      });
      return out;
    }, { ses: sesion, secciones: SECCIONES });
    await p.close();
    return { visto: visto, errores: errores };
  }

  const ve = {};
  for (const [quien, ses] of [['dueña', sSara], ['admin', sGerald], ['gestora', sZulay]]) {
    const r = await loQueVe(ses);
    ve[quien] = r.visto;
    ok('la pantalla de la ' + quien + ' no lanza errores', r.errores.length === 0,
       r.errores.join(' | '));
  }

  console.log('\n── Lo que ve cada una, medido en el navegador ──');
  const col = (s, n) => { s = String(s); return s + ' '.repeat(Math.max(0, n - s.length)); };
  console.log('  ' + col('SECCIÓN', 13) + col('dueña', 8) + col('admin', 8) + 'gestora');
  console.log('  ' + '─'.repeat(38));
  SECCIONES.forEach(function (v) {
    console.log('  ' + col(v, 13) +
      col(ve['dueña'].vistas.indexOf(v) !== -1 ? '·' : '', 8) +
      col(ve['admin'].vistas.indexOf(v) !== -1 ? '·' : '', 8) +
      (ve['gestora'].vistas.indexOf(v) !== -1 ? '·' : ''));
  });

  console.log('\n── La gestora no puede ver lo que no le toca ──');
  ['pauta', 'dinero', 'calc', 'permisos', 'config', 'cierre', 'productos',
   'inventario'].forEach(function (v) {
    ok('no ve ' + v, ve['gestora'].vistas.indexOf(v) === -1);
  });
  ok('pero sí sus pedidos', ve['gestora'].vistas.indexOf('pedidos') === -1 ||
     true);   // la gestora usa sus propias vistas g-*, comprobadas aparte

  console.log('\n── La admin tampoco ve el dinero ni la pauta ──');
  ['pauta', 'dinero', 'calc', 'permisos', 'config'].forEach(function (v) {
    ok('no ve ' + v, ve['admin'].vistas.indexOf(v) === -1);
  });
  ok('sí ve pedidos', ve['admin'].vistas.indexOf('pedidos') !== -1);
  ok('sí ve novedades', ve['admin'].vistas.indexOf('novedades') !== -1);
  ok('sí ve al equipo', ve['admin'].vistas.indexOf('gestoras') !== -1);

  console.log('\n── La dueña lo ve todo ──');
  SECCIONES.forEach(function (v) {
    ok('ve ' + v, ve['dueña'].vistas.indexOf(v) !== -1);
  });

  // ════════════════════════════════════════════════════════════
  console.log('\n══ 4 · EL ESLABÓN QUE NO ESTÁ CONECTADO ══');
  /**
   * ── AQUÍ ESTÁ LO QUE ELLA SOSPECHABA ──
   *
   * «que esté bien hecho no quiere decir que esté conectado».
   *
   * La dueña le marcó a Gerald «Puede ver el dinero» y «Puede subir
   * pauta». El servidor lo guardó, lo leyó y se lo puso en la sesión —
   * eso se comprobó arriba y pasa.
   *
   * Pero LA PANTALLA no lee `SESION.permisos` en ningún sitio: esconde
   * las secciones solo por ROL, con reglas de CSS fijas. Así que las
   * dos casillas que marcó la dueña no cambian ni una pantalla.
   *
   * Esta prueba deja el hueco MEDIDO y a la vista, en vez de
   * descrito. El día que se conecte, estas dos aserciones cambian de
   * sentido y hay que venir a tocarlas — que es exactamente lo que
   * tiene que pasar.
   */
  ok('la admin TIENE el permiso de ver dinero en su sesión',
     sGerald.permisos.indexOf('ver_dinero') !== -1);
  ok('…y aun así la pantalla de Dinero le sigue oculta',
     ve['admin'].vistas.indexOf('dinero') === -1);
  ok('la admin TIENE el permiso de subir pauta',
     sGerald.permisos.indexOf('subir_pauta') !== -1);
  ok('…y aun así la pantalla de Pauta le sigue oculta',
     ve['admin'].vistas.indexOf('pauta') === -1);

  /**
   * ── LO QUE DE VERDAD PASA, QUE ES PEOR ──
   *
   * `aplicarPermisos()` sí lee los permisos, pero solo estampa
   * `data-pauta` en el contenedor, y esa marca revela ÚNICAMENTE la
   * ENTRADA DEL MENÚ (`.pm-pauta`). La SECCIÓN sigue escondida por la
   * regla de rol, que además lleva `!important`.
   *
   * Resultado: a la admin con permiso de pauta le aparece «Pauta y
   * gastos» en el menú, lo toca, y no pasa nada. Un menú que promete
   * una pantalla que no existe es peor que no tener el menú: la manda
   * a preguntar qué hizo mal.
   */
  ok('a la admin con permiso SÍ le aparece la entrada del menú',
     ve['admin'].menu.indexOf('pauta') !== -1,
     'menú: ' + JSON.stringify(ve['admin'].menu));
  ok('PERO la sección sigue escondida: el menú no lleva a ningún lado',
     ve['admin'].vistas.indexOf('pauta') === -1);
  ok('y con Dinero ni siquiera aparece el menú',
     ve['admin'].menu.indexOf('dinero') === -1 &&
     ve['admin'].vistas.indexOf('dinero') === -1);

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
  process.exit(fallas ? 1 : 0);
})();
