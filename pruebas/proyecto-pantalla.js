/** La pantalla del proyecto a fondo, con datos falsos. */
const { chromium } = require('/opt/node22/lib/node_modules/playwright');
const PR = {
  ok: true, hoy: '2026-09-23', mes: '2026-09',
  semana: { lunes: '2026-09-21', domingo: '2026-09-27' },
  roles: { socia: { nombre: 'Socia' }, trabajadora: { nombre: 'Trabajadora' },
           propio: { nombre: 'Propio' }, estudio: { nombre: 'Estudio' } },
  modalidades: { fijo: { nombre: 'Precio fijo' }, porcentaje: { nombre: 'Porcentaje' },
                 por_hora: { nombre: 'Por hora' }, sin_cobro: { nombre: 'No se cobra' } },
  bases: { utilidad_neta: { nombre: 'Utilidad neta del mes' } },
  tiposFuente: [{ id: 'excel', nombre: 'Excel o CSV', lee: true },
                { id: 'pdf', nombre: 'PDF', lee: false }],
  tiendas: [{ id: 'EC', nombre: 'Nutrea EC', empresa: 'Nutrea' }],
  proyecto: { id: 't1', nombre: 'Nutrea EC', contraparte: 'Sociedad', tipo: 'propio',
    estado: 'activo', moneda: 'USD', valor: 0, inicio: '', entrega: '', horasSemana: 0,
    especificacion: 'Mi 50% de la tienda de Ecuador', documento: '', nota: '',
    rol: 'socia', rolDeducido: false, modalidad: 'porcentaje', modalidadDeducida: false,
    porcentaje: 50, base: 'utilidad_neta', tiendaId: 'EC', clienteId: 'c1',
    confidencial: false },
  fuentes: [{ id: 'f1', nombre: 'Cierre de agosto', tipo: 'excel', tipoNombre: 'Excel o CSV',
              lee: true, enlace: 'https://docs.google.com/x', nota: '', agregado: '2026-09-01' }],
  cobros: [], plata: { cobrado: {}, porCobrar: {}, atrasado: {}, falta: null },
  parte: { hay: true, tienda: 'EC', tiendaNombre: 'Nutrea EC', empresa: 'Nutrea',
    mes: '2026-09', moneda: 'USD', base: 'utilidad_neta',
    baseNombre: 'Utilidad neta del mes', valorBase: 62, porcentaje: 50, parte: 31,
    enPerdida: false, hayDatos: true, faltan: [],
    detalle: { ganancia: 186, costoDevoluciones: 8, pauta: 66, costosFijos: 50 } },
  tareas: [{ id: 's1', texto: 'Revisar el cierre del mes', fecha: '2026-09-25', horas: 2,
             estado: 'pendiente', riesgo: 'corrible', estaSemana: true, vencida: false }],
  carga: { abiertas: 1, horas: 2, vencidas: 0 },
  movimientos: [],
};
(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };
  for (const ancho of [1100, 390]) {
    const A = ' (@' + ancho + ')';
    const p = await b.newPage({ viewport: { width: ancho, height: 1400 } });
    p.on('pageerror', e => { fallas++; console.log('  JS ERROR: ' + e.message); });
    await p.goto('file:///home/claude/repo/novacentral.html');
    await p.waitForLoadState('load');
    await p.evaluate((x) => {
      window.nc = async (a) => (a === 'nc_proyecto' ? x : { ok: true });
      (0, eval)('nc = window.nc;');
      (0, eval)("NC_SESION = { nombre:'M', rol:'socia', permisos:['ver','facturacion'] };");
      document.getElementById('login-screen').style.display = 'none';
      document.getElementById('app').style.display = 'flex';
    }, PR);
    await p.evaluate(() => abrirProyecto('t1'));
    await p.waitForTimeout(200);
    const c = await p.textContent('#pr-cuerpo');
    ok('muestra la especificación' + A, /Mi 50% de la tienda/.test(c));
    ok('dice qué soy y cómo cobro' + A, /Socia/.test(c) && /Porcentaje/.test(c));
    ok('la cuenta del 50% va resta por resta' + A,
       /Ganancia de lo entregado/.test(c) && /Devoluciones/.test(c) &&
       /Pauta/.test(c) && /Costos fijos del mes/.test(c), c.replace(/\s+/g,' ').slice(0,200));
    ok('con la base y el resultado' + A, /Utilidad neta del mes/.test(c) && /Mi 50%/.test(c));
    ok('las fuentes dicen que son enlaces, no archivos' + A,
       /Se guarda el enlace, no el archivo/.test(c));
    ok('y cuál se puede leer de verdad' + A, /Nova puede leerlo/.test(c));
    ok('las entregas de la semana salen' + A,
       /Revisar el cierre del mes/.test(c) && /ESTA SEMANA/.test(c));
    ok('hay dónde pegar el alcance' + A, /Sacar las entregas de un texto/.test(c));
    ok('y dice que el PDF todavía necesita el modelo' + A, /modelo de lenguaje/.test(c));
    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral' + A, lateral === 0, lateral + 'px');
    if (ancho === 1100) await p.locator('#v-proyecto').screenshot({ path: '/tmp/nova-pruebas/central-proyecto.png' });
    await p.close();
  }
  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
