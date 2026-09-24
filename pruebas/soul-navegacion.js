/**
 * ABRIR CADA SECCIÓN DE NOVASOUL Y VER QUE PINTE ALGO.
 *
 * ┌─ EL ERROR QUE NADIE VIO ───────────────────────────────────┐
 * │                                                            │
 * │ «ya no me carga NovaSoul, ni Universidad ni Cielo me        │
 * │  carga».                                                   │
 * │                                                            │
 * │ El servidor estaba bien —sus pruebas pasaban todas— y la   │
 * │ pantalla compilaba sin un error. Lo que estaba roto era el │
 * │ menú: `go()` decía «si no hay datos, tráelos», y con la    │
 * │ precarga nueva los datos YA estaban, así que no traía      │
 * │ nada… y tampoco pintaba. La sección quedaba en blanco con  │
 * │ todo cargado en memoria.                                   │
 * │                                                            │
 * │ Ninguna prueba de servidor podía verlo. Ninguna prueba de  │
 * │ sintaxis podía verlo. Hacía falta ESTA: entrar, tocar cada │
 * │ entrada del menú, y mirar que aparezca contenido.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Se prueban los DOS caminos, porque son distintos y cada uno se
 * rompe por su lado:
 *
 *   · SIN precarga → al abrir, la sección se trae sola.
 *   · CON precarga → al abrir, la sección pinta lo que ya llegó.
 *
 * El segundo es el que falló. El primero es el que seguía funcionando
 * y por eso el fallo parecía imposible.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

/* ═══════════════════════════════════════════════════════════════
     LOS DATOS SALEN DEL SERVIDOR DE VERDAD, NO DE AQUÍ
   ═══════════════════════════════════════════════════════════════

   El primer intento de esta prueba traía las respuestas escritas a
   mano. Dos de las cinco secciones «fallaron» — y no era la pantalla:
   era que a mis respuestas inventadas les faltaban campos que el
   servidor sí manda (`resumen.dia`, `ocupadasPorDia`).

   Una prueba que puede fallar porque el que la escribió se equivocó
   copiando la forma de los datos no sirve para nada: gasta el tiempo
   en el error de la prueba y deja pasar el del producto.

   Así que aquí se siembran las hojas con lo suyo, se corren LAS
   FUNCIONES DEL BUNDLE, y lo que devuelvan es lo que se le mete a la
   pantalla. Si el servidor cambia de forma, esto cambia con él. */

const fs = require('fs');
const LIBROS = { cen: {}, s: {}, a: {}, emp: {} };
function libroStub(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (n) => {
    const m = hojas[n];
    if (!m) return null;
    return {
      getName: () => n, getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (k) => m.splice(k - 1, 1),
      getRange: (f, c, nf, nc) => ({
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
          r.forEach((x, j) => { d[c - 1 + j] = x; }); }); },
        setValue: (x) => { const d = m[f - 1] || (m[f - 1] = []); d[c - 1] = x; },
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-24T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: libroStub, flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' },
                     WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
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
  '\n;globalThis.__S = { soulArranque, soulCielo, soulCieloLectura, soulPlataOrdenada,' +
  ' soulHoy, soulRutina, soulMaterias, libroOlvidar_, soulOlvidar_, SOUL_HOJAS };');
const SRV = globalThis.__S;
const H = SRV.SOUL_HOJAS;
const YO = 'manuela@nova.com';
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const filaDe = (cols, o) => cols.map(c => (o[c] !== undefined ? o[c] : ''));

/** Sus hojas, con lo suyo: su carta de Horus, sus gastos, sus turnos. */
function sembrarHojas() {
  UUID = 0; SRV.libroOlvidar_(); SRV.soulOlvidar_();
  LIBROS.cen = { Finanzas: [['id','fecha','flujo','categoria','monto','moneda','concepto']],
                 Trabajos: [['id','nombre','tipo','estado','horas_semana','fecha_entrega']],
                 Cobros: [['id']], Metas: [['id']] };
  LIBROS.s = {
    Carta: [['usuario_id','cuerpo','signo','grado','casa','retrogrado','nota']],
    Transitos: [['usuario_id','fecha','casa','casa_placidus','tema','intensidad_pct',
      'texto_transito','por_que','como_trabajarlo','el_otro_lado','cuerpo','aspecto',
      'a_natal','desde','hasta','fuente']],
    Pensum: [H.Pensum.slice()], Fijos: [H.Fijos.slice()], Rutina: [H.Rutina.slice()],
    Turnos: [H.Turnos.slice()], Pendientes: [H.Pendientes.slice()],
    Materias: [H.Materias.slice()], Mindlab: [H.Mindlab.slice()],
    Horas: [['usuario_id','dia_semana','horas_libres','nota']],
    Revolucion: [['usuario_id','anio','desde','hasta','ascendente','casa_sol','tema','texto','nota']],
    Usuarios: [['id','nombre','correo','fecha_nacimiento','hora_nacimiento',
                'lugar_nacimiento','zona_horaria','acento','modo','idioma'],
               ['u1','Manuela',YO,'1995-09-20','13:20','Palmira','America/Bogota','','','']],
  };
  [['ascendente','capricornio',15.1,1],['sol','virgo',27.3,9],['luna','leo',10.2,8],
   ['mercurio','libra',20,10],['venus','libra',5.7,10],['marte','escorpio',9,11],
   ['jupiter','sagitario',9,12],['saturno','piscis',20.9,3],['urano','capricornio',26.6,1],
   ['neptuno','capricornio',22.8,1],['pluton','escorpio',28.3,11],
   ['medio_cielo','libra',19.4,10],['nodo_norte','libra',26.9,10],
  ].forEach(c => LIBROS.s.Carta.push([YO, c[0], c[1], c[2], c[3], '', '']));

  LIBROS.s.Pendientes.push(filaDe(H.Pendientes, { id:'p1', usuario_id:YO,
    texto:'Terminar la propuesta de la carta', tipo:'trabajo', fecha:'2026-09-26',
    estado:'pendiente', prioridad:'alta', horas_estimadas:3, riesgo:'acordado' }));

  const fj = (o) => LIBROS.s.Fijos.push(filaDe(H.Fijos,
    Object.assign({ usuario_id: YO, activo:'si', moneda:'COP', flujo:'gasto' }, o)));
  fj({ id:'f1', categoria:'arriendo', concepto:'apto', monto:788000, dia_del_mes:15, tipo_pago:'mensual' });
  fj({ id:'f2', categoria:'credito', concepto:'nu', monto:223000, dia_del_mes:18,
       tipo_pago:'cuotas', acreedor:'Nu', cuotas_total:12, cuota_desde:'2026-03' });

  LIBROS.s.Rutina.push(filaDe(H.Rutina, { id:'r1', usuario_id:YO, tipo:'turno',
    nombre:'Turno bar', dia_semana:'V', hora_inicio:'18:00', hora_fin:'02:00',
    paga_fija:90000, moneda:'COP', activo:'si' }));
  LIBROS.s.Turnos.push(filaDe(H.Turnos, { id:'t1', usuario_id:YO, rutina_id:'r1',
    fecha:'2026-09-18', paga:90000, propinas:35000, moneda:'COP', estado:'hecho' }));

  LIBROS.s.Materias.push(filaDe(H.Materias, { id:'m1', usuario_id:YO,
    nombre:'Psicología del consumidor', codigo:'PSI-401', profesor:'Ramírez',
    semestre:'2026-2', estado:'activa' }));
}

console.log('Corriendo el servidor de verdad para sacar las respuestas…');
sembrarHojas();
const HOY_ = SRV.soulHoy(SOCIA, {});
const PLATA_ = SRV.soulPlataOrdenada(SOCIA, {});
const CIELO_ = SRV.soulCielo(SOCIA, {});
const LECTURA_ = SRV.soulCieloLectura(SOCIA, { rango: 'semana' });
const RUTINA_ = SRV.soulRutina(SOCIA, {});
const UNI_ = SRV.soulMaterias(SOCIA, {});
sembrarHojas();
const ARRANQUE_ = SRV.soulArranque(SOCIA, {});

[['Hoy', HOY_], ['plata', PLATA_], ['cielo', CIELO_], ['lectura', LECTURA_],
 ['rutina', RUTINA_], ['universidad', UNI_], ['arranque', ARRANQUE_]
].forEach(function (par) {
  if (!par[1] || par[1].ok !== true) {
    console.log('  ⚠ el servidor falló en ' + par[0] + ': ' +
                (par[1] && par[1].error || 'sin respuesta'));
  }
});

/** Qué tiene que aparecer en cada sección para llamarla «cargada». */
const SECCIONES = [
  { v: 'hoy',    nombre: 'Hoy',         busca: /propuesta de la carta/i },
  { v: 'plata',  nombre: 'Mi plata',    busca: /Cuotas y deudas/i },
  { v: 'rutina', nombre: 'Mi rutina',   busca: /Turno bar/i },
  { v: 'uni',    nombre: 'Universidad', busca: /Psicolog/i },
  { v: 'cielo',  nombre: 'El cielo',    busca: /Gibosa creciente/i },
];

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

async function abrirNovaSoul(b, conPrecarga) {
  const p = await b.newPage({ viewport: { width: 1280, height: 1100 } });
  const errores = [];
  p.on('pageerror', e => errores.push(e.message));
  p.on('console', m => {
    if (m.type() === 'error' && !/ERR_CERT|Failed to load resource/.test(m.text())) {
      errores.push('console: ' + m.text());
    }
  });

  await p.goto('file:///home/claude/repo/novasoul.html');
  await p.waitForLoadState('load');

  await p.evaluate((d) => {
    // La puerta se llama `login-screen`, y `#app` nace oculto.
    const puerta = document.getElementById('login-screen');
    if (puerta) puerta.style.display = 'none';
    document.getElementById('app').style.display = 'flex';

    window.__pedidas = [];
    window.nc = async (accion, datos) => {
      window.__pedidas.push(accion);
      const R = {
        nc_soul: d.hoy, nc_soul_plata: d.plata, nc_soul_cielo: d.cielo,
        nc_soul_rutina: d.rutina, nc_soul_materias: d.uni,
        nc_soul_cielo_lectura: d.lectura, nc_soul_arranque: d.arranque,
      }[accion];
      if (!R) return { ok: true };
      // Sin precarga, el arranque «no existe» en esa hoja.
      if (accion === 'nc_soul_arranque' && !d.conPrecarga) {
        return { ok: false, error: 'Acción desconocida.' };
      }
      return JSON.parse(JSON.stringify(R));
    };
  }, { hoy: HOY_, plata: PLATA_, cielo: CIELO_, rutina: RUTINA_, uni: UNI_,
       lectura: LECTURA_, arranque: ARRANQUE_, conPrecarga: conPrecarga });

  await p.evaluate(() => cargar());
  await p.waitForTimeout(500);
  return { p, errores };
}

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  for (const conPrecarga of [false, true]) {
    const C = conPrecarga ? ' (con precarga)' : ' (sin precarga)';
    console.log('\n══ Abriendo cada sección' + C + ' ══');

    const { p, errores } = await abrirNovaSoul(b, conPrecarga);

    if (conPrecarga) {
      const pedidas = await p.evaluate(() => window.__pedidas);
      ok('la precarga se pidió de una sola vez',
         pedidas.filter(x => x === 'nc_soul_arranque').length === 1,
         JSON.stringify(pedidas));
      /**
       * Y —lo que se rompió— que tener los datos en memoria NO impida
       * pintarlos. Si esta parte falla, la sección se ve vacía aunque
       * todo esté cargado.
       */
    }

    for (const S of SECCIONES) {
      await p.evaluate((v) => go(v), S.v);
      await p.waitForTimeout(350);
      const t = await p.textContent('#v-' + S.v);
      const limpio = (t || '').replace(/\s+/g, ' ').trim();
      ok(S.nombre + ' pinta su contenido' + C, S.busca.test(limpio),
         limpio ? 'lo que hay: «' + limpio.slice(0, 160) + '…»' : 'la sección está VACÍA');
    }

    console.log('\n── Y nada de esto puede estar roto' + C + ' ──');
    /** Los textos que se rompieron cuando porMoneda tenía dos versiones. */
    await p.evaluate(() => go('plata'));
    await p.waitForTimeout(250);
    const plata = (await p.textContent('#v-plata')).replace(/\s+/g, ' ');
    ok('el texto de «vacío» no se imprime pegado al monto' + C,
       !/nada aúnCOP|sin definirCOP|nadaCOP/.test(plata),
       plata.slice(0, 200));
    ok('la deuda dice cuándo se acaba' + C, /febrero de 2027/.test(plata), plata.slice(0, 200));
    ok('y cuántas cuotas van' + C, /van 6 de 12/.test(plata));

    await p.evaluate(() => go('cielo'));
    await p.waitForTimeout(400);
    const cielo = (await p.textContent('#v-cielo')).replace(/\s+/g, ' ');
    ok('la luna explica qué es una gibosa' + C, /giba/.test(cielo), cielo.slice(0, 200));
    ok('la lectura trae las cuatro preguntas' + C,
       /TIEMPO PARA QUÉ/.test(cielo) || /tiempo para/i.test(cielo));
    ok('y los ejes del karma' + C, /eje del karma/i.test(cielo));

    ok('sin errores de JavaScript' + C, errores.length === 0, errores.join(' | '));
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
  process.exit(fallas ? 1 : 0);
})();
