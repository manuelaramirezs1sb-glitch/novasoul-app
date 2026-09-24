/**
 * AUDITORÍA EN LA PANTALLA · que el botón haga algo.
 *
 * ┌─ LO QUE FALLÓ, Y POR QUÉ NINGUNA PRUEBA LO VIO ────────────┐
 * │                                                            │
 * │ «auditorías tampoco está guardando los hallazgos».          │
 * │                                                            │
 * │ El botón «Guardar veredicto» no tenía `onclick`. Ninguna    │
 * │ prueba de servidor podía verlo —el servidor estaba bien—   │
 * │ y ninguna prueba de pantalla existía. Se podía marcar el   │
 * │ hallazgo, escribir la nota, darle a guardar, y no pasaba   │
 * │ nada: ni un error, porque no había nada que fallara.       │
 * │                                                            │
 * │ Así que esto hace exactamente lo que hizo ella: entra,     │
 * │ marca, escribe y le da a guardar. Y comprueba que la       │
 * │ petición SALIÓ, con el veredicto y la nota dentro.         │
 * │                                                            │
 * │ Y comprueba lo otro: que no quede ni una clienta           │
 * │ inventada en la pantalla.                                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');

/** Los nombres de la maqueta vieja. Si alguno aparece, alguien volvió. */
const INVENTADOS = ['Andrea Morales', 'Luisa Fernández', 'Carolina Jiménez',
                    'NE-4412', 'NE-4410', 'NE-4408', 'Camila', 'Valentina', 'Daniela'];

const CASOS = [
  { novedadId: 'n1', pedidoId: 'dropi-7781', fecha: '2026-09-21',
    cliente: 'Marcela Andrade', ciudad: 'Quito', motivo: 'Dirección incorrecta',
    desenlace: '', gestora: 'Paola',
    dijoGestora: 'Dirección incorrecta. Llamé 3 veces, no contesta.',
    llamadas: [
      { cuando: '2026-09-21 10:30', agente: 'Paola', sentido: 'saliente',
        estado: 'sin_respuesta', segundos: 0, observacion: '', grabacion: '' },
      { cuando: '2026-09-21 12:15', agente: 'Paola', sentido: 'saliente',
        estado: 'sin_respuesta', segundos: 0, observacion: '', grabacion: '' },
    ],
    senales: [{ tipo: 'llamadas_de_menos', nivel: 'mal',
                que: 'Escribió 3 intentos y en la central hay 2.', dato: { dijo: 3, hay: 2 } }],
    veredicto: 'pendiente', notaAuditoria: '', auditadoPor: '', auditadoEn: '' },
  { novedadId: 'n2', pedidoId: 'dropi-7790', fecha: '2026-09-22',
    cliente: 'Rosa Benítez', ciudad: 'Guayaquil', motivo: 'No contesta',
    desenlace: '', gestora: 'Paola',
    dijoGestora: 'Llamé 1 vez y quedó entregado.',
    llamadas: [{ cuando: '2026-09-22 09:10', agente: 'Paola', sentido: 'saliente',
                 estado: 'contestada', segundos: 95, observacion: 'Confirmó dirección',
                 grabacion: '' }],
    senales: [{ tipo: 'calza', nivel: 'bien',
                que: 'Lo escrito y lo registrado calzan: 1 llamada(s), 1 con conversación.',
                dato: {} }],
    veredicto: 'ok', notaAuditoria: 'Revisado.', auditadoPor: 'm@nova.com',
    auditadoEn: '2026-09-23T10:00:00Z' },
];

const RESPUESTA = {
  ok: true, hoy: '2026-09-24', desde: '2026-09-10', tienda: 'ec',
  casos: CASOS,
  veredictos: { ok: { nombre: 'OK' }, hallazgo: { nombre: 'Hallazgo' },
                pendiente: { nombre: 'Pendiente' } },
  recuento: { cerradasEnVentana: 9, mostrados: 2, revisados: 1, sinRevisar: 1,
              hallazgos: 0, conSenal: 1 },
  sinLlamadas: false, porqueSinLlamadas: '',
};

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });

  for (const ancho of [1200, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1400 } });
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
      // `#app` nace oculto: sin esto todo mide cero y parece roto.
      document.getElementById('app').style.display = 'flex';
      (0, eval)('CONECTADO = true; ST = "ec"; ROL = "dueno";');
    });

    const enviados = [];
    await p.exposeFunction('__anotar', (x) => { enviados.push(JSON.parse(x)); });
    await p.evaluate((R) => {
      window.api = async (accion, params) => {
        if (accion === 'auditoria_casos') return JSON.parse(JSON.stringify(R));
        if (accion === 'auditoria_guardar') {
          window.__anotar(JSON.stringify({ accion, params }));
          return { ok: true, veredicto: params.veredicto, cuando: '2026-09-24T12:00:00Z' };
        }
        return { ok: true };
      };
      return cargarAuditoriaCasos();
    }, RESPUESTA);
    await p.waitForTimeout(150);
    await p.evaluate(() => go('auditoria', document.querySelector('[data-v=auditoria]')));
    await p.waitForTimeout(100);

    const texto = await p.textContent('#v-auditoria');

    console.log('\n── No queda nada inventado' + A + ' ──');
    const quedan = INVENTADOS.filter(n => texto.indexOf(n) !== -1);
    ok('ninguna clienta ni gestora de la maqueta' + A, quedan.length === 0,
       'siguen ahí: ' + quedan.join(', '));

    console.log('\n── Los dos relatos, uno al lado del otro' + A + ' ──');
    ok('la clienta real' + A, /Marcela Andrade/.test(texto));
    ok('lo que escribió la gestora' + A, /Llamé 3 veces/.test(texto));
    ok('y lo que registró la central' + A, /10:30/.test(texto) && /12:15/.test(texto));
    ok('la discrepancia, con los dos números' + A,
       /Escribió 3 intentos y en la central hay 2/.test(texto));
    ok('y cuando calza, también lo dice' + A, /calzan/.test(texto));

    console.log('\n── El botón existe Y tiene quien lo atienda' + A + ' ──');
    const botones = await p.$$('#audrows .nsavebtn');
    ok('hay un botón por caso' + A, botones.length === 2, 'hay ' + botones.length);
    const conHandler = await p.evaluate(() =>
      Array.from(document.querySelectorAll('#audrows .nsavebtn'))
        .filter(b => b.getAttribute('onclick')).length);
    ok('y CADA uno lleva su onclick' + A, conHandler === 2,
       conHandler + ' de 2 tienen handler');

    console.log('\n── Marcar, escribir y guardar' + A + ' ──');
    await p.evaluate(() => marcarVeredicto(0, 'hallazgo'));
    await p.waitForTimeout(40);
    const marcado = await p.evaluate(() =>
      document.querySelector('#aud-0 .nst.on').textContent.trim());
    ok('marcar cambia lo resaltado' + A, marcado === 'Hallazgo', marcado);

    await p.fill('#aud-nota-0', 'Falta el tercer intento. Hablado con ella el 24.');
    await p.click('#audrows .nsavebtn');
    await p.waitForTimeout(250);

    ok('la petición SALIÓ' + A, enviados.length >= 1,
       'no salió ninguna — el botón vuelve a no hacer nada');
    if (enviados.length) {
      const e = enviados[enviados.length - 1];
      ok('con la acción de guardar' + A, e.accion === 'auditoria_guardar', e.accion);
      ok('con el veredicto marcado' + A, e.params.veredicto === 'hallazgo',
         JSON.stringify(e.params.veredicto));
      ok('con la nota que escribió' + A,
         /Falta el tercer intento/.test(e.params.nota || ''), e.params.nota);
      ok('y con el caso al que pertenece' + A, e.params.novedadId === 'n1',
         e.params.novedadId);
      /** Las señales viajan para saber con qué datos se puso el veredicto. */
      ok('y con las señales que Nova había levantado' + A,
         Array.isArray(e.params.senales) &&
         e.params.senales.indexOf('llamadas_de_menos') !== -1,
         JSON.stringify(e.params.senales));
    }

    console.log('\n── Un veredicto ya puesto se ve puesto' + A + ' ──');
    ok('el segundo caso sale marcado OK' + A,
       (await p.evaluate(() =>
         (document.querySelector('#aud-1 .nst.on') || {}).textContent || '')).trim() === 'OK');
    ok('con su nota cargada' + A,
       (await p.inputValue('#aud-nota-1')) === 'Revisado.');
    ok('y diciendo quién lo puso' + A, /m@nova\.com/.test(texto));

    console.log('\n── Los recuentos no se inventan' + A + ' ──');
    const kpi = await p.textContent('#audkpi');
    ok('revisados: 1 de 2' + A, /1 de 2/.test(kpi), kpi);
    ok('y 1 con señal' + A, /CON SEÑAL/.test(kpi));

    console.log('\n── Sin errores de JavaScript' + A + ' ──');
    ok('la pantalla no lanzó ninguno' + A, errores.length === 0, errores.join(' | '));

    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
  process.exit(fallas ? 1 : 0);
})();
