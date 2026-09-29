/**
 * EL CIELO EN PANTALLA: de a poco, y activas hoy primero.
 *
 * ┌─ LO QUE PIDIÓ ─────────────────────────────────────────────┐
 * │                                                            │
 * │ «me lo tiraste todo de un solo golpe, mucha información me  │
 * │  sobre-estimula (…) me gusta que ya tenga la info, pero     │
 * │  debe dármela poco a poco cuando la necesite».              │
 * │                                                            │
 * │ «para la parte de temporadas abiertas, que estén como       │
 * │  principal, activas hoy. Eso me interesa leerlo».           │
 * │                                                            │
 * │ «me importa lo de Saturno porque es una energía constante   │
 * │  y duradera y debo aprovecharla».                           │
 * │                                                            │
 * │ «yo usaría Placidus como sistema principal y casas enteras  │
 * │  como segunda capa».                                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Las cuatro son de PANTALLA, así que se comprueban en pantalla. El
 * servidor ya está probado en `cielo-lectura.js`; aquí lo que importa
 * es qué se ve y en qué orden.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const temporada = (titulo, desde, hasta, dias, activa, borde) => ({
  hay: true, titulo: titulo, cuerpo: 'Saturno', aspecto: 'Cuadratura', simbolo: '□',
  aNatal: 'Ascendente', casaNombre: 'Casa 4', casaArea: 'la casa y la raíz',
  casaNumero: 4, grupo: 'social', ritmo: 'dos años y medio por signo',
  duraTipico: 'meses', tension: 'dura', intensidad: 80,
  tiempoPara: 'Tiempo para arreglar la base.',
  queEs: 'Saturno aprieta, ordena y cobra.',
  aFavor: ['Ordena lo que sostiene.'], desequilibrio: ['Se nota en la casa.'],
  nota: 'Lectura común.', desde: desde, hasta: hasta, dias: dias,
  activaHoy: activa, pico: activa, suyo: null,
  borde: borde ? { casa: 3, nombre: 'Casa 3', area: 'el entorno cercano',
    texto: 'Va por el borde: en Placidus cae en tu casa 4 y en casas enteras en tu ' +
           'casa 3. Se lee con Placidus; lo de casas enteras es la segunda capa — ' +
           'también se está moviendo el entorno cercano.' } : null,
});

const LEC = {
  ok: true, hoy: '2026-09-29',
  rangos: { semana: { nombre: 'Esta semana' }, mes: { nombre: 'Este mes' },
            anio: { nombre: 'Este año' } },
  lectura: {
    rango: 'semana', nombre: 'Esta semana',
    que: 'Lo que se decide en días. Aquí manda la Luna.',
    desde: '2026-09-28', hasta: '2026-10-04',
    temporadas: [temporada('Luna llena sobre tu Venus', '2026-09-27', '2026-10-01', 5, true, true),
                 temporada('Marte trígono a tu Sol', '2026-10-02', '2026-10-09', 8, false, false)],
    activas: [temporada('Luna llena sobre tu Venus', '2026-09-27', '2026-10-01', 5, true, true)],
    porVenir: [temporada('Marte trígono a tu Sol', '2026-10-02', '2026-10-09', 8, false, false)],
    deFondo: [
      { titulo: 'Saturno cuadratura a tu Ascendente · por tu casa 4',
        hasta: '2027-03-18', intensidad: 92, dias: 175 },
      { titulo: 'Neptuno oposición a tu Sol · por tu casa 3',
        hasta: '2027-03-03', intensidad: 78, dias: 161 },
    ],
    lunas: [], lunaHoy: { nombre: 'Cuarto menguante', forma: 'media luna', que: 'soltar' },
    fueraDeRango: 'Lo que dura meses no cabe en una semana. Está en Este mes y en Este año.',
    vacio: false, porqueVacio: '',
  },
  sembro: 0, efemerides: { vencen: false },
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  const p = await b.newPage({ viewport: { width: 1100, height: 1400 } });
  const errs = [];
  p.on('pageerror', (e) => errs.push(e.message));
  await p.goto('file:///home/claude/repo/novasoul.html');
  await p.waitForLoadState('load');
  await p.evaluate((LEC) => {
    const lg = document.getElementById('login-screen');
    if (lg) lg.style.display = 'none';
    const app = document.getElementById('app');
    if (app) app.style.display = 'block';
    (0, eval)('LEC = ' + JSON.stringify(LEC) + '; LEC_RANGO = "semana"; LEC_ABIERTA = -1;');
    const c = document.getElementById('ci-lectura');
    if (c && c.closest('.view')) c.closest('.view').style.display = 'block';
    pintarLectura();
  }, LEC);
  await p.waitForTimeout(250);

  const leer = () => p.evaluate(() =>
    (document.getElementById('ci-lectura').innerText || '').replace(/\s+/g, ' '));

  console.log('\n══ 1 · LO ACTIVO HOY ES LO PRINCIPAL ══');
  let txt = await leer();
  ok('el encabezado habla de lo ACTIVO HOY, no de «abiertas»',
     /1 TEMPORADA ACTIVA HOY/.test(txt), txt.slice(0, 200));
  ok('y la activa se ve', /Luna llena sobre tu Venus/.test(txt), txt.slice(0, 260));
  /**
   * La que importa: lo que todavía no empieza NO compite. Antes salían
   * las ocho de golpe, sin jerarquía, y por eso la sobre-estimulaba.
   */
  const visible = await p.evaluate(() => {
    const d = [...document.querySelectorAll('#ci-lectura details')][0];
    return { hayDetalle: !!d, abierto: d ? d.open : null,
             resumen: d ? d.querySelector('summary').innerText.trim() : '' };
  });
  ok('lo que no ha empezado está plegado', visible.hayDetalle && visible.abierto === false,
     JSON.stringify(visible));
  ok('y se dice cuántas son, para poder decidir abrirlo',
     /1 que todavía no empieza/.test(visible.resumen), visible.resumen);

  console.log('\n══ 2 · SATURNO NO DESAPARECE, PERO NO COMPITE ══');
  ok('el fondo largo se nombra', /DE FONDO, CORRIENDO IGUAL/.test(txt) &&
     /Saturno cuadratura a tu Ascendente/.test(txt), txt.slice(-400));
  ok('con hasta cuándo', /hasta/.test(txt));
  ok('y dice dónde se lee entero', /Este año/.test(txt));
  /**
   * Sin su lectura completa: si trajera los cuatro bloques sería otra
   * vez todo de golpe, que es justo lo que pidió que no pasara.
   */
  ok('pero SIN su lectura completa encima',
     txt.indexOf('Saturno aprieta, ordena y cobra') === -1, txt.slice(-300));

  console.log('\n══ 3 · LAS CASAS: PLACIDUS MANDA, EL BORDE INFORMA ══');
  await p.evaluate(() => abrirTemporada(0));
  await p.waitForTimeout(150);
  txt = await leer();
  ok('al abrir una temporada sale su lectura', /TIEMPO PARA QUÉ/.test(txt));
  ok('el borde se dice como información', /Va por el borde/.test(txt) &&
     /segunda capa/.test(txt), txt.slice(0, 400));
  /**
   * Y NO como alarma. El recuadro rojo saltaba en 80 de sus 178
   * tránsitos —el 44%—: avisaba de lo normal, y una alarma que suena la
   * mitad de las veces deja de leerse.
   */
  const alarmas = await p.evaluate(() =>
    document.querySelectorAll('#ci-lectura .aviso.mal.on').length);
  igual('y no hay ningún recuadro de alarma por las casas', 0, alarmas);
  ok('tampoco el texto viejo que lo anunciaba',
     txt.indexOf('Los dos sistemas de casas no coinciden') === -1);

  console.log('\n══ 4 · UNA SEMANA TRANQUILA SE DICE, NO SE DEJA EN BLANCO ══');
  await p.evaluate(() => {
    (0, eval)('LEC.lectura.activas = []; LEC.lectura.porVenir = []; ' +
      'LEC.lectura.temporadas = []; LEC.lectura.vacio = true; ' +
      'LEC.lectura.porqueVacio = "Ninguna temporada corta abierta esta semana: nada ' +
      'que se decida en días."; LEC_ABIERTA = -1;');
    pintarLectura();
  });
  await p.waitForTimeout(150);
  txt = await leer();
  ok('lo explica en vez de dejar el hueco', /nada que se decida en días/.test(txt),
     txt.slice(0, 220));
  ok('y el fondo sigue ahí, que es lo que sí está corriendo',
     /Saturno cuadratura a tu Ascendente/.test(txt), txt.slice(-300));

  igual('sin errores de consola', [], errs);
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
  process.exit(fallas ? 1 : 0);
})();
