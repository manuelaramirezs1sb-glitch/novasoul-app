/**
 * NovaSoul en pantalla.
 *
 * Lo que se comprueba no es que pinte bonito: es que diga la verdad.
 * Que sin horas libres NO dibuje una semana holgada. Que lo vencido se
 * vea vencido. Que el tablero mueva de verdad y no solo de mentira.
 * Que un error al dibujar no se disfrace de falta de internet. Y que
 * nunca se vean dos tiendas a la vez.
 */
const { chromium } = require(process.env.PLAYWRIGHT || '/opt/node22/lib/node_modules/playwright');
const OUT = '/tmp/nova-pruebas/';

const HOY = '2026-09-23';
const RIESGOS = {
  inamovible: { nombre: 'No se puede mover', orden: 3, ayuda: 'Un parcial, una entrega con fecha de ellos.' },
  acordado:   { nombre: 'Acordado con alguien', orden: 2, ayuda: 'Se puede correr, pero hay que avisar.' },
  corrible:   { nombre: 'Se puede correr', orden: 1, ayuda: 'Nadie está esperando una fecha exacta.' },
};
const TRABAJOS = [
  { id: 't1', nombre: 'PHH', tipo: 'empleo', estado: 'activo', horasSemana: 12, entrega: '' },
  { id: 't2', nombre: 'Son de Sky', tipo: 'cliente', estado: 'activo', horasSemana: 0, entrega: '2026-09-30' },
  { id: 't3', nombre: 'Universidad', tipo: 'estudio', estado: 'activo', horasSemana: 6, entrega: '' },
];
const T = (id, texto, o) => Object.assign({
  id, texto, trabajoId: '', proyecto: '', tipo: '', origen: '', fecha: '', dias: null,
  estado: 'pendiente', prioridad: 'media', riesgo: 'corrible', horas: 0, horasReales: 0,
  hechoEn: '', nota: '',
}, o);

const PEND = [
  T('p1', 'Entregar la carta de cocteles', { trabajoId: 't2', proyecto: 'Son de Sky', fecha: '2026-09-25', dias: -2, horas: 6, riesgo: 'acordado' }),
  T('p2', 'Parcial de Estadística', { trabajoId: 't3', proyecto: 'Universidad', fecha: '2026-09-20', dias: 3, horas: 4, riesgo: 'inamovible', prioridad: 'alta' }),
  T('p3', 'Encargo de PHH', { trabajoId: 't1', proyecto: 'PHH', fecha: '2026-09-24', dias: -1, horas: 5, estado: 'haciendo' }),
  T('p4', 'Comprar Omega 3', { fecha: '2026-09-26', dias: -3, horas: 1 }),
  T('p5', 'Llamar a mamá', { fecha: '2026-09-21', dias: 2, estado: 'hecho', hechoEn: '2026-09-21' }),
];

const BLOQUE_TURNO = { id: 'r1', tipo: 'turno', nombre: 'Salsabor', lugar: 'Salsabor',
  inicio: '18:00', fin: '02:00', horas: 8 };
const BLOQUE_CLASE = { id: 'r3', tipo: 'clase', nombre: 'Estadística', lugar: 'Bloque 3',
  inicio: '07:00', fin: '09:00', horas: 2 };
const dia = (f, dow, esHoy, libres, textos, bloques) => ({
  fecha: f, dow, esHoy, libres,
  utiles: libres === null ? null : libres + (bloques || []).reduce((a, b) => a + b.horas, 0),
  ocupadas: (bloques || []).reduce((a, b) => a + b.horas, 0),
  bloques: bloques || [],
  entregas: textos.length, horas: textos.reduce((a, t) => a + t.horas, 0), textos,
});

function foto(conHoras) {
  return {
    ok: true, hoy: HOY, mes: '2026-09',
    semana: { lunes: '2026-09-21', domingo: '2026-09-27', dias: [
      dia('2026-09-21', 1, false, conHoras ? 4 : null, [], [BLOQUE_CLASE]),
      dia('2026-09-22', 2, false, conHoras ? 4 : null, []),
      dia('2026-09-23', 3, true, conHoras ? 4 : null, []),
      dia('2026-09-24', 4, false, conHoras ? 4 : null, [{ id: 'p3', texto: 'Encargo de PHH', proyecto: 'PHH', horas: 5 }]),
      dia('2026-09-25', 5, false, conHoras ? 4 : null, [{ id: 'p1', texto: 'Entregar la carta de cocteles', proyecto: 'Son de Sky', horas: 6 }], [BLOQUE_TURNO]),
      dia('2026-09-26', 6, false, conHoras ? 0 : null, [{ id: 'p4', texto: 'Comprar Omega 3', proyecto: '', horas: 1 }]),
      dia('2026-09-27', 7, false, conHoras ? 0 : null, []),
    ] },
    trabajos: TRABAJOS,
    tipos: { cliente: { nombre: 'Cliente' }, empleo: { nombre: 'Empleo' },
             propio: { nombre: 'Propio (Nova)' }, estudio: { nombre: 'Universidad' } },
    riesgos: RIESGOS,
    pendientes: PEND,
    urgente: PEND[1],
    resumen: {
      dia: { entregas: 0, horas: 0, vencidas: 1, hechasHoy: 0, haciendo: 1 },
      semana: { entregas: 3, horasEntregas: 12, hechas: 1 },
      mes: { entregas: 5, hechas: 1, abiertas: 4, porProyecto: [
        { id: 't2', nombre: 'Son de Sky', abiertas: 1, horas: 6, vencidas: 0 },
        { id: 't1', nombre: 'PHH', abiertas: 1, horas: 5, vencidas: 0 },
        { id: 't3', nombre: 'Universidad', abiertas: 1, horas: 4, vencidas: 1 },
        { id: '·sueltas', nombre: 'Sin proyecto', abiertas: 1, horas: 1, vencidas: 0 },
      ] },
    },
    riesgo: conHoras
      ? { libres: 20, fijas: 18, extra: 7, comprometidas: 25, sobra: -5, dentroDeFijas: 2, ocupadas: 10,
          candidatas: [
            { id: 'p4', texto: 'Comprar Omega 3', horas: 1, proyecto: '', riesgo: 'corrible', fecha: '2026-09-26' },
            { id: 'p1', texto: 'Entregar la carta de cocteles', horas: 6, proyecto: 'Son de Sky', riesgo: 'acordado', fecha: '2026-09-25' },
          ], noAlcanza: 0, sinHoras: false, porque: '' }
      : { libres: null, fijas: 18, extra: 7, comprometidas: 25, sobra: null, dentroDeFijas: 2, ocupadas: 10,
          candidatas: [], noAlcanza: 0, sinHoras: true,
          porque: 'Todavía no me has dicho cuántas horas libres tienes cada día. ' +
                  'Sin ese número no puedo decirte si la semana cabe: lo demás sería un adorno.' },
    horas: conHoras ? { 1: 4, 2: 4, 3: 4, 4: 4, 5: 4, 6: 0, 7: 0 } : null,
    horasUtiles: conHoras ? { 1: 6, 2: 4, 3: 4, 4: 4, 5: 12, 6: 0, 7: 0 } : null,
    rutina: { bloques: 2, ocupadasSemana: 10 },
    turnosPendientes: 1,
    faltanHojas: [],
    mindlab: {
      inicio: '2026-09-28', fin: '2026-12-20', hechas: 1, total: 12, horasTotales: 37,
      semanaActual: null,
      plan: [1,2,3,4,5,6,7,8,9,10,11,12].map(function (n) {
        return { id: 'ml' + n, semana: n, mes: n <= 4 ? 1 : n <= 8 ? 2 : 3,
                 tema: 'Tema ' + n, tarea: 'Tarea de la semana ' + n,
                 horas: n === 11 || n === 12 ? 2 : 3,
                 desde: '2026-09-28', hasta: '2026-10-04',
                 estado: n === 1 ? 'hecho' : 'pendiente', nota: '' };
      }),
    },
  };
}

const FINANZAS = {
  ok: true, mes: '2026-09', hoy: HOY,
  categorias: [
    { id: 'arriendo', nombre: 'Arriendo', tipo: 'gasto',
      lineas: [{ id: 'x1', categoria: 'arriendo', concepto: 'Apartamento', monto: 1250000, moneda: 'COP', dia: 5, activo: true, nota: '' }],
      planeado: { COP: 1250000 }, gastado: { COP: 1400000 } },
    { id: 'mercado', nombre: 'Mercado', tipo: 'gasto', lineas: [], planeado: null, gastado: { COP: 420000 } },
    { id: 'servicios', nombre: 'Servicios', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'internet', nombre: 'Internet', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'credito', nombre: 'Crédito', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'deudas', nombre: 'Deudas', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'movil', nombre: 'Móvil', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'varios', nombre: 'Gastos varios', tipo: 'gasto', lineas: [], planeado: null, gastado: null },
    { id: 'ahorro', nombre: 'Ahorro', tipo: 'ahorro',
      lineas: [{ id: 'x2', categoria: 'ahorro', concepto: 'Colchón', monto: 300000, moneda: 'COP', dia: 30, activo: true, nota: '' }],
      planeado: { COP: 300000 }, gastado: null },
  ],
  totales: { plan: { COP: 1250000 }, ahorro: { COP: 300000 } },
  error: '',
};

const RUTINA = {
  ok: true, hoy: HOY,
  tipos: { turno: { nombre: 'Turno', paga: true, ayuda: 'Salsabor.' },
           clase: { nombre: 'Clase', paga: false, ayuda: 'Una materia.' },
           otro:  { nombre: 'Otro',  paga: false, ayuda: 'Gimnasio.' } },
  rutinas: [
    { id: 'r1', tipo: 'turno', nombre: 'Salsabor', dia: 5, inicio: '18:00', fin: '02:00',
      horas: 8, lugar: 'Salsabor', trabajoId: 't1', materiaId: '', paga: 80000,
      moneda: 'COP', desde: '', hasta: '', activo: true },
    { id: 'r3', tipo: 'clase', nombre: 'Estadística', dia: 1, inicio: '07:00', fin: '09:00',
      horas: 2, lugar: 'Bloque 3', trabajoId: '', materiaId: 'm1', paga: 0,
      moneda: 'COP', desde: '2026-08-03', hasta: '2026-12-05', activo: true },
  ],
  ocupadasPorDia: { 1: 2, 2: 0, 3: 0, 4: 0, 5: 8, 6: 0, 7: 0 },
  trabajos: TRABAJOS,
  materias: [{ id: 'm1', nombre: 'Estadística' }],
};

const PLATA = {
  ok: true, hoy: HOY, mes: '2026-09',
  categorias: FINANZAS.categorias, totales: FINANZAS.totales, error: '',
  hormigaCategorias: [
    { id: 'bus', nombre: 'Buses' }, { id: 'transporte', nombre: 'Transporte' },
    { id: 'uber', nombre: 'Uber' }, { id: 'antojos', nombre: 'Antojos' },
    { id: 'salidas', nombre: 'Salidas' }, { id: 'hormiga', nombre: 'Otros sueltos' }],
  turnosPendientes: [
    { rutinaId: 'r1', nombre: 'Salsabor', lugar: 'Salsabor', fecha: '2026-09-18',
      inicio: '18:00', fin: '02:00', horas: 8, paga: 80000, pagaFija: 80000,
      moneda: 'COP', propinas: 0, escrito: false, turnoId: '' },
  ],
  resumen: {
    mes: '2026-09', ultimoDia: '2026-09-30',
    entro: { COP: 114000 }, salio: { COP: 97900 }, hormiga: { COP: 97900 },
    hormigaPorCategoria: [
      { id: 'bus', nombre: 'Buses', n: 1, monto: { COP: 2900 } },
      { id: 'transporte', nombre: 'Transporte', n: 0, monto: null },
      { id: 'uber', nombre: 'Uber', n: 1, monto: { COP: 18000 } },
      { id: 'antojos', nombre: 'Antojos', n: 1, monto: { COP: 12000 } },
      { id: 'salidas', nombre: 'Salidas', n: 1, monto: { COP: 65000 } },
      { id: 'hormiga', nombre: 'Otros sueltos', n: 0, monto: null },
    ],
    resultado: [{ moneda: 'COP', entro: 114000, salio: 97900, queda: 16100,
                  fijosPendientes: 1200000, porVenir: 160000, proyectado: -1023900 }],
    turnosPorVenir: 2,
    movimientosHormiga: [
      { id: 'h1', fecha: '2026-09-23', categoria: 'bus', concepto: 'Buses', monto: 2900, moneda: 'COP' },
      { id: 'h2', fecha: '2026-09-20', categoria: 'uber', concepto: 'Uber', monto: 18000, moneda: 'COP' },
    ],
    error: '',
  },
};

const CUERPOS = {
  sol: { nombre: 'Sol', grupo: 'personal' }, luna: { nombre: 'Luna', grupo: 'personal' },
  mercurio: { nombre: 'Mercurio', grupo: 'personal' }, venus: { nombre: 'Venus', grupo: 'personal' },
  marte: { nombre: 'Marte', grupo: 'personal' }, jupiter: { nombre: 'Júpiter', grupo: 'social' },
  saturno: { nombre: 'Saturno', grupo: 'social' }, urano: { nombre: 'Urano', grupo: 'generacional' },
  neptuno: { nombre: 'Neptuno', grupo: 'generacional' }, pluton: { nombre: 'Plutón', grupo: 'generacional' },
  ascendente: { nombre: 'Ascendente', grupo: 'angulo' },
};
const GRUPOS = {
  personal: { nombre: 'Personales', que: 'Lo inmediato, lo mío, lo propio.' },
  social: { nombre: 'Sociales', que: 'Expansión y estructura.' },
  generacional: { nombre: 'Generacionales', que: 'Marcan época. Van más allá del ego.' },
  angulo: { nombre: 'Ángulos', que: 'Por dónde entras y hacia dónde apuntas.' },
};
const MOMENTOS = {
  aprender: { nombre: 'Aprender', que: 'Entra información.' },
  descansar: { nombre: 'Descansar', que: 'Se sostiene lo que hay.' },
  cambiar: { nombre: 'Cambiar', que: 'Se cierra, se suelta, se mueve.' },
};
const lunaDia = (n, ilum, crece) => ({ id: 'llena', nombre: n, momento: 'cambiar',
  que: '', edadDias: 14, iluminacion: ilum, creciendo: crece });

const CIELO = {
  ok: true, hoy: HOY,
  nacimiento: { nombre: 'Manuela', fecha: '1998-09-14', hora: '04:20',
                lugar: 'Medellín, Colombia', zona: 'America/Bogota' },
  cuerpos: CUERPOS, grupos: GRUPOS, momentos: MOMENTOS,
  signos: [{ id: 'leo', nombre: 'Leo' }, { id: 'virgo', nombre: 'Virgo' }],
  fases: [{ id: 'llena', nombre: 'Luna llena', momento: 'cambiar' }],
  tieneCarta: true,
  carta: [], cartaPorGrupo: {
    personal: [{ cuerpo: 'sol', nombre: 'Sol', grupo: 'personal', signo: 'virgo',
                 signoNombre: 'Virgo', grado: 21.5, casa: 2, retrogrado: false },
               { cuerpo: 'marte', nombre: 'Marte', grupo: 'personal', signo: 'virgo',
                 signoNombre: 'Escorpio', grado: 14, casa: 4, retrogrado: true }],
    social: [{ cuerpo: 'saturno', nombre: 'Saturno', grupo: 'social', signo: 'leo',
               signoNombre: 'Aries', grado: 29, casa: 9, retrogrado: false }],
    generacional: [{ cuerpo: 'pluton', nombre: 'Plutón', grupo: 'generacional', signo: 'leo',
                     signoNombre: 'Sagitario', grado: 7, casa: null, retrogrado: true }],
    angulo: [{ cuerpo: 'ascendente', nombre: 'Ascendente', grupo: 'angulo', signo: 'leo',
               signoNombre: 'Leo', grado: 15, casa: null, retrogrado: false }],
  },
  lunaHoy: lunaDia('Luna llena', 96, false),
  semana: { lunes: '2026-09-21', dias: [1,2,3,4,5,6,7].map(function (i) {
    return { fecha: '2026-09-' + (20 + i), esHoy: i === 3,
             luna: lunaDia('Luna llena', 90 + i, false),
             momento: 'cambiar', deDonde: i === 3 ? 'pensum' : 'luna',
             porque: i === 3 ? 'Saturno por casa 10' : 'Luna llena',
             pensum: [], transitos: 1 };
  }) },
  pensum: [{ id: 'pn1', desde: '2026-09-01', hasta: '2026-12-31', titulo: 'Saturno por casa 10',
             cuerpo: 'saturno', cuerpoNombre: 'Saturno', grupo: 'social', casa: 10,
             momento: 'cambiar', quePide: 'Cerrar lo que ya no sostiene.', queEvitar: '', nota: '' }],
  pensumAbierto: [{ id: 'pn1', titulo: 'Saturno por casa 10', momento: 'cambiar' }],
  transitos: [
    { cuerpo: 'saturno', nombre: 'Saturno', grupo: 'social', aspecto: 'cuadratura',
      aNatal: 'mi Luna', casa: 10, tema: 'Estructura', intensidad: null,
      desde: '2026-09-01', hasta: '2026-11-30', texto: '', porQue: '', como: '',
      elOtroLado: '', fuente: 'Horus' },
    { cuerpo: 'luna', nombre: 'Luna', grupo: 'personal', aspecto: '', aNatal: '', casa: 4,
      tema: 'Luna por casa 4', intensidad: null, desde: '2026-09-22', hasta: '2026-09-24',
      texto: '', porQue: '', como: '', elOtroLado: '', fuente: 'Horus' },
  ],
  transitosHoy: [{ cuerpo: 'saturno' }, { cuerpo: 'luna' }],
  revolucion: { anio: 2026, desde: '2026-09-14', hasta: '2027-09-13', edad: 28,
                diasRestantes: 355, transcurrido: 3,
                carta: { anio: 2026, ascendente: 'leo', casaSol: 11,
                         tema: 'El año de mostrar', texto: '', nota: '',
                         planetas: [{ cuerpo: 'marte', casa: 1 }] },
                lectura: { hay: true, momento: 'descansar', momentoNombre: 'Descansar',
                  partes: [
                    { clave: 'ascendente', titulo: 'Cómo entras al año', valor: 'Leo',
                      texto: 'Se entra mostrándose. Lo que hagas este año se va a ver.' },
                    { clave: 'sol', titulo: 'Dónde va tu atención', valor: 'Casa 11 · La gente',
                      texto: 'Red, comunidad, proyectos con otros. Es una casa sucedente, así que el año pide descansar.' }],
                  grupos: [{ grupo: 'personal', nombre: 'Personales',
                    que: 'Lo inmediato, lo mío, lo propio.',
                    cuerpos: [{ cuerpo: 'marte', nombre: 'Marte', casa: 1, area: 'Tú',
                                que: 'Cómo te presentas y qué cuerpo le pones al año.',
                                momento: 'cambiar' }] }],
                  faltan: [], porque: '' } },
  casas: [], pensumPropuesto: [],
  medicion: { minimo: 8, conFecha: 30, promedio: 62,
    fases: [{ id: 'llena', nombre: 'Luna llena', momento: 'cambiar', n: 12, hechas: 5, pct: 42, faltan: 0 },
            { id: 'cuarto_menguante', nombre: 'Cuarto menguante', momento: 'cambiar', n: 10, hechas: 8, pct: 80, faltan: 0 },
            { id: 'nueva', nombre: 'Luna nueva', momento: 'aprender', n: 3, hechas: 2, pct: null, faltan: 5 }],
    conMuestra: 2,
    patron: 'Las entregas que pusiste en cuarto menguante las terminaste el 80% de las veces, contra un 62% en general. En luna llena bajas al 42%.',
    porque: '' },
};

const CARTA_LEIDA = {
  ok: true, cuerpos: CUERPOS, grupos: GRUPOS,
  signos: [{ id: 'virgo', nombre: 'Virgo' }],
  encontradas: [
    { cuerpo: 'sol', nombre: 'Sol', grupo: 'personal', signo: 'virgo', signoNombre: 'Virgo',
      grado: 21.5, casa: 2, retrogrado: false, linea: "Sol en Virgo 21°34' Casa 2" },
    { cuerpo: 'saturno', nombre: 'Saturno', grupo: 'social', signo: 'virgo', signoNombre: 'Aries',
      grado: 29, casa: 9, retrogrado: false, linea: 'Saturno en Aries 29°48 Casa 9' },
  ],
  ignoradas: [{ linea: 'Generado por Horus', porque: 'No encontré ni planeta ni signo.' }],
  faltan: ['Luna', 'Mercurio'],
};

const MATERIAS = {
  ok: true, hoy: HOY,
  trabajos: [{ id: 't3', nombre: 'Universidad', tipo: 'estudio', estado: 'activo', horasSemana: 6, entrega: '' }],
  materias: [
    { id: 'm1', nombre: 'Estadística', codigo: 'EST-301', profesor: 'Prof. Gómez',
      carpeta: 'https://drive.google.com/drive/folders/abc', semestre: '2026-2',
      trabajoId: 't3', estado: 'activa', nota: '', abiertas: 2, vencidas: 1, proxima: '2026-10-15' },
    { id: 'm2', nombre: 'Química Ambiental', codigo: '', profesor: '', carpeta: '',
      semestre: '', trabajoId: 't3', estado: 'activa', nota: '',
      abiertas: 0, vencidas: 0, proxima: '' },
  ],
};

const SILABO_LEIDO = {
  ok: true, maximo: 300, truncado: false,
  encontradas: [
    { titulo: 'Parcial 1', sinTitulo: false, fecha: '2026-10-15', anioInferido: true,
      rango: false, peso: 25, linea: 'Parcial 1 — 15 de octubre — 25%' },
    { titulo: 'Entrega del proyecto', sinTitulo: false, fecha: '2026-10-16', anioInferido: true,
      rango: true, peso: 20, linea: 'Entrega del proyecto: semana del 12 al 16 de octubre 20%' },
  ],
  ignoradas: [
    { linea: 'Bibliografía: Walpole, 9a edición', porque: 'No encontré ninguna fecha.' },
    { linea: 'Participación en clase 10%', porque: 'Tiene un porcentaje pero ninguna fecha.' },
  ],
};

const FAMILY = {
  ok: true,
  clientes: [{ id: 'c1', empresa: 'Nutrea', sheetId: 'sid' }],
  cliente: { id: 'c1', empresa: 'Nutrea', sheetId: 'sid' },
  tiendas: [{ id: 'gt', nombre: 'Nutrea GT' }, { id: 'ec', nombre: 'Nutrea EC' }],
  tienda: 'gt',
  alarmas: [{ id: 'a1', nivel: 'rojo', nombre: 'Novedades', titulo: '4 novedades sin resolver',
              detalle: 'Llevan más de 24 horas.', casos: 4 }],
  semaforo: { tienda: 'gt', moneda: 'GTQ', semana: { lunes: '2026-09-14', domingo: '2026-09-20' },
    luces: { entrega: { estado: 'verde', valor: '78%', etiqueta: 'Tasa de entrega', porque: 'sobre 65% mínimo' },
             cpa: { estado: 'rojo', valor: 'GTQ 92', etiqueta: 'CPA contra el techo', porque: 'el techo es GTQ 70' } },
    alertas: [{ titulo: 'El CPA pasó el techo tres semanas seguidas.' }],
    hayDatos: true },
  errorTienda: '',
  central: { proyectos: 3, entregasVencidas: 1, cobrosAtrasados: 1,
             atrasadoPorMoneda: { COP: 800000 }, carga: {} },
  academy: { estudiantes: 0, porque: 'Todavía no has dado de alta a nadie en novAcademy.' },
};

(async () => {
  const b = await chromium.launch({ executablePath: '/opt/pw-browsers/chromium' });
  let fallas = 0;
  const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
    else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? ' — ' + d : '')); } };

  const abrir = async (ancho, conHoras, rol) => {
    const p = await b.newPage({ viewport: { width: ancho, height: 1200 } });
    p.on('pageerror', e => { fallas++; console.log('  JS ERROR: ' + e.message); });
    await p.goto('file:///home/claude/repo/novasoul.html');
    await p.waitForLoadState('load');
    await p.evaluate((x) => {
      window.LLAMADAS = [];
      window.nc = async (accion, datos) => {
        window.LLAMADAS.push({ accion, datos });
        if (accion === 'nc_soul') return x.foto;
        if (accion === 'nc_soul_plata') return x.plata;
        if (accion === 'nc_soul_rutina') return x.rut;
        if (accion === 'nc_soul_cielo') return x.cielo;
        if (accion === 'nc_soul_carta_leer') return x.cartaLeida;
        if (accion === 'nc_soul_turno') return { ok: true, total: 114000, moneda: 'COP' };
        if (accion === 'nc_soul_family') return x.fam;
        if (accion === 'nc_soul_materias') return x.mat;
        if (accion === 'nc_soul_silabo') return x.sil;
        if (accion === 'nc_soul_silabo_guardar') return { ok: true, creadas: 2, repetidas: 0, errores: [] };
        return { ok: true };
      };
      (0, eval)('nc = window.nc;');
      (0, eval)('SES = ' + JSON.stringify({ correo: 'm@nova.com', nombre: 'Manuela', rol: x.rol }) + ';');
      (0, eval)('entrar();');
    }, { foto: foto(conHoras), plata: PLATA, fam: FAMILY, mat: MATERIAS,
         rut: RUTINA, sil: SILABO_LEIDO, cielo: CIELO, cartaLeida: CARTA_LEIDA,
         rol: rol || 'socia' });
    await p.waitForTimeout(200);
    return p;
  };

  for (const ancho of [1200, 390]) {
    const A = ' (@' + ancho + ')';

    // ══ Sin horas libres: el caso que más importa ══
    let p = await abrir(ancho, false);
    ok('la puerta se cierra al entrar' + A,
       await p.$eval('#login-screen', e => getComputedStyle(e).display) === 'none');

    let riesgo = await p.textContent('#hoy-riesgo');
    ok('sin horas NO dibuja una semana holgada: pide el dato' + A,
       /cuántas horas libres/.test(riesgo), riesgo.replace(/\s+/g, ' ').slice(0, 90));
    ok('y no inventa una barra' + A,
       await p.$$eval('#hoy-riesgo .hbar', e => e.length) === 0);
    ok('ofrece el botón para darlo' + A,
       /Decirle cuántas horas tengo/.test(riesgo));

    let semDias = await p.textContent('#sem-dias');
    ok('y cada día dice que faltan las horas, en vez de poner cero' + A,
       /sin horas dichas/.test(semDias));
    await p.close();

    // ══ Con horas: la cuenta completa ══
    p = await abrir(ancho, true);

    const kpis = await p.textContent('#hoy-kpis');
    ok('las cifras de arriba: hoy, vencidas, semana y mes' + A,
       /ENTREGAS DE HOY/.test(kpis) && /VENCIDAS/.test(kpis) &&
       /ESTA SEMANA/.test(kpis) && /ESTE MES/.test(kpis));
    ok('lo vencido va en rojo' + A,
       await p.$$eval('#hoy-kpis .kpi-v.mal', e => e.length) === 1,
       await p.$$eval('#hoy-kpis .kpi-v.mal', e => e.length) + ' en rojo');

    const urg = await p.textContent('#hoy-urgente');
    ok('«lo primero» es UNA sola cosa, la vencida' + A,
       /Parcial de Estadística/.test(urg) && /venció hace 3 días/.test(urg));
    ok('y dice que no se puede mover' + A, /No se puede mover/.test(urg));

    riesgo = await p.textContent('#hoy-riesgo');
    ok('dice cuántas horas hay y cuántas se pidieron' + A,
       /20 h libres/.test(riesgo) && /25 h comprometidas/.test(riesgo),
       riesgo.replace(/\s+/g, ' ').slice(0, 120));
    ok('y cuántas faltan' + A, /FALTAN 5 h/.test(riesgo));
    ok('explica que lo de un proyecto con horas fijas no se cuenta dos veces' + A,
       /no se suman otra vez/.test(riesgo));
    ok('propone candidatas a caerse' + A, /ALGO SE VA A CAER/.test(riesgo));
    ok('y el parcial NO está entre ellas' + A,
       riesgo.split('ALGO SE VA A CAER')[1].indexOf('Parcial') === -1);
    ok('la barra reparte en tres tramos' + A,
       await p.$$eval('#hoy-riesgo .hbar span', e => e.length) === 3);

    // ══ El tablero ══
    await p.evaluate(() => go('pendientes', null));
    const cols = await p.$$eval('#pend-kan .kcol', e => e.length);
    ok('el tablero tiene tres columnas' + A, cols === 3, cols + ' columnas');
    const tarjetas = await p.$$eval('#pend-kan .kcard', e => e.length);
    ok('y las cinco tarjetas' + A, tarjetas === 5, tarjetas + ' tarjetas');
    const enHaciendo = await p.$eval('#kcol-haciendo', e => e.textContent);
    ok('cada una en su columna' + A, /Encargo de PHH/.test(enHaciendo));
    const kanTxt = await p.textContent('#pend-kan');
    ok('lo vencido dice cuántos días lleva' + A, /venció hace 3 días/.test(kanTxt));
    ok('y lo inamovible se marca' + A, /no se mueve/.test(kanTxt));
    ok('lo hecho lleva su fecha, no «hoy»' + A, /21 sep/.test(kanTxt));

    // Mover de verdad llama al servidor con el estado nuevo
    await p.evaluate(() => { window.LLAMADAS.length = 0; });
    await p.click('#kcol-pendiente .kcard .kbtn');
    await p.waitForTimeout(120);
    const llamadas = await p.evaluate(() => window.LLAMADAS);
    ok('mover una tarjeta lo guarda en el servidor' + A,
       llamadas.length >= 1 && llamadas[0].accion === 'nc_soul_guardar' &&
       llamadas[0].datos.datos.estado === 'haciendo',
       JSON.stringify(llamadas[0] || null));
    ok('y vuelve a traer la foto, para no pintar lo que cree que pasó' + A,
       llamadas.filter(l => l.accion === 'nc_soul').length === 1);

    // El filtro por proyecto
    const filtros = await p.$$eval('#pend-filtros .kbtn', e => e.map(x => x.textContent));
    ok('se puede filtrar por proyecto' + A,
       filtros.indexOf('Son de Sky') !== -1, JSON.stringify(filtros));

    // ══ El formulario ══
    await p.evaluate(() => nuevaTarea());
    ok('el formulario abre' + A,
       await p.$eval('#m-tarea', e => getComputedStyle(e).display) !== 'none');
    const proys = await p.$$eval('#mt-trabajo option', e => e.map(x => x.textContent));
    ok('y ofrece los proyectos de Central' + A,
       proys.length === 4 && proys.indexOf('PHH') !== -1, JSON.stringify(proys));
    ok('tiene horas estimadas y fecha de entrega' + A,
       (await p.$$('#mt-horas')).length === 1 && (await p.$$('#mt-fecha')).length === 1);
    ok('y pregunta qué pasa si no se entrega' + A,
       (await p.$$eval('#mt-riesgo option', e => e.length)) === 3);
    await p.fill('#mt-texto', '');
    await p.click('#mt-guardar');
    await p.waitForTimeout(80);
    ok('no deja guardar una tarea sin texto' + A,
       /Escribe qué hay que hacer/.test(await p.textContent('#mt-err')));
    await p.evaluate(() => cerrarModal('m-tarea'));

    // ══ Mindlab ══
    await p.evaluate(() => go('mindlab', null));
    const ml = await p.textContent('#v-mindlab');
    ok('Mindlab son doce semanas' + A,
       await p.$$eval('#ml-lista .fila', e => e.length) === 12);
    ok('con sus tres meses' + A,
       /MES 1 · TRÁFICO/.test(ml) && /MES 3 · VENTA/.test(ml));
    ok('dice cuánto pesa en total y que lo recorté' + A,
       /37 h en total/.test(ml) && /último trimestre/.test(ml), ml.replace(/\s+/g, ' ').slice(0, 160));
    ok('y por qué las dos últimas semanas del año quedan libres' + A,
       /finales y fiestas/.test(ml));

    // ══ Gastos fijos ══
    await p.evaluate(() => go('plata', null));
    await p.waitForTimeout(150);
    const fin = await p.textContent('#fin-lista');
    ok('están las nueve categorías que pidió' + A,
       await p.$$eval('#fin-lista .fin-row', e => e.length) === 10, // 9 + la cabecera
       await p.$$eval('#fin-lista .fin-row', e => e.length) + ' filas');
    ok('con sus nombres' + A,
       ['Arriendo','Mercado','Servicios','Internet','Crédito','Deudas','Móvil',
        'Gastos varios','Ahorro'].every(n => fin.indexOf(n) !== -1));
    ok('lo que no tiene monto dice «sin definir», no cero' + A,
       /sin definir/.test(fin) && !/COP 0/.test(fin));
    ok('y donde se pasó, lo dice' + A, /se pasó COP 150.000/.test(fin));
    ok('el ahorro no se suma con los gastos' + A,
       /va aparte, no es un gasto/.test(await p.textContent('#fin-kpis')));

    // ══ La rutina: se escribe una vez ══
    await p.evaluate(() => go('rutina', null));
    await p.waitForTimeout(180);
    const rut = await p.textContent('#v-rutina');
    ok('lista turnos y clases' + A, /Salsabor/.test(rut) && /Estadística/.test(rut));
    ok('dice que se repite solo' + A, /no hay que volver a escribirlo/.test(rut));
    ok('el turno de noche dura 8 h, no menos 16' + A, /8 h/.test(rut), rut.match(/[\d,]+ h/g) + '');
    ok('marca lo que no tiene fecha de fin' + A, /sin fecha de fin/.test(rut));
    ok('y la clase muestra hasta cuándo va' + A, /hasta 5 dic/.test(rut));

    await p.evaluate(() => nuevaRutina());
    ok('el formulario pregunta si es turno o clase' + A,
       (await p.$$eval('#mr-tipo option', e => e.length)) === 3);
    await p.selectOption('#mr-tipo', 'turno');
    ok('un turno sí pregunta cuánto pagan' + A,
       await p.$eval('#mr-pago', e => getComputedStyle(e).display) !== 'none');
    await p.selectOption('#mr-tipo', 'clase');
    ok('una clase NO pregunta paga' + A,
       await p.$eval('#mr-pago', e => getComputedStyle(e).display) === 'none');
    await p.selectOption('#mr-tipo', 'turno');
    await p.fill('#mr-inicio', '18:00');
    await p.fill('#mr-fin', '02:00');
    await p.waitForTimeout(80);
    ok('calcula la duración y avisa que cruza la medianoche' + A,
       /8 h/.test(await p.textContent('#mr-dura')) &&
       /cruza la medianoche/.test(await p.textContent('#mr-dura')),
       await p.textContent('#mr-dura'));
    await p.evaluate(() => cerrarModal('m-rutina'));

    // ══ Las horas útiles, con la resta a la vista ══
    await p.evaluate(() => abrirHoras());
    const mh = await p.textContent('#m-horas');
    ok('el modal avisa que cambió lo que pregunta' + A, /Cambió lo que te pregunto/.test(mh));
    ok('y muestra la resta de cada día' + A,
       /menos 2 h de turno o clase/.test(mh) && /menos 8 h de turno o clase/.test(mh), mh.slice(0, 200));
    await p.fill('#mh-1', '9');
    await p.waitForTimeout(80);
    ok('la resta se actualiza mientras escribe' + A,
       /7 h libres/.test(await p.textContent('#mh-r1')), await p.textContent('#mh-r1'));
    await p.evaluate(() => cerrarModal('m-horas'));

    // ══ La semana dibuja los bloques ══
    await p.evaluate(() => go('semana', null));
    const sem = await p.textContent('#sem-dias');
    ok('el viernes muestra el turno con su horario' + A,
       /Salsabor/.test(sem) && /18:00–02:00/.test(sem));
    ok('y el lunes la clase' + A, /Estadística/.test(sem) && /07:00–09:00/.test(sem));
    ok('cada día enseña la resta, no solo el resultado' + A,
       /12 h − 8 h = 4 h/.test(sem.replace(/\s+/g, ' ')), sem.replace(/\s+/g, ' ').slice(0, 220));

    // ══ Turnos sin cerrar ══
    await p.evaluate(() => go('hoy', null));
    const ht = await p.textContent('#hoy-turnos');
    ok('Hoy avisa que falta cerrar un turno' + A, /1 turno sin cerrar/.test(ht));
    ok('y dice por qué no lo puede hacer solo' + A, /solo lo sabes tú/.test(ht));
    ok('el menú lo marca' + A,
       await p.$eval('#nb-turnos', e => getComputedStyle(e).display) !== 'none');

    // ══ Gasto del día, en dos toques ══
    const hh = await p.textContent('#hoy-hormiga');
    ok('Hoy tiene los botones de gasto rápido' + A,
       /Bus/.test(hh) && /Uber/.test(hh) && /Antojo/.test(hh) && /Salida/.test(hh));
    await p.evaluate(() => { window.LLAMADAS.length = 0; abrirHormiga('antojos'); });
    await p.fill('#mho-monto', '12000');
    await p.click('#m-hormiga .card-cta');
    await p.waitForTimeout(150);
    const gh = (await p.evaluate(() => window.LLAMADAS))
      .filter(l => l.accion === 'nc_soul_hormiga')[0];
    ok('guarda el gasto con su categoría' + A,
       gh && gh.datos.categoria === 'antojos' && gh.datos.monto === '12000',
       JSON.stringify(gh && gh.datos));

    // ══ Mi plata: ¿sobra o falta? ══
    await p.evaluate(() => go('plata', null));
    await p.waitForTimeout(220);
    const pl = await p.textContent('#pl-mes');
    ok('dice lo que entró y lo que salió' + A,
       /ENTRÓ/.test(pl) && /114.000/.test(pl) && /97.900/.test(pl));
    ok('y lo que falta por pasar, aparte' + A,
       /Y FALTA QUE PASE ESTO/.test(pl) && /Turnos que faltan por trabajar/.test(pl));
    ok('no mezcla el hecho con la previsión, y lo explica' + A,
       /Lo de arriba ya pasó/.test(pl));
    ok('avisa que este mes faltaría plata' + A,
       /TE FALTARÍAN COP 1.023.900/.test(pl), pl.replace(/\s+/g, ' ').slice(0, 160));
    ok('y el bloque va en rojo, no en verde' + A,
       (await p.$$eval('#pl-mes .card-rojo', e => e.length)) === 1);

    const plt = await p.textContent('#pl-turnos');
    ok('los turnos sin cerrar salen con su base' + A,
       /TURNOS SIN CERRAR/.test(plt) && /COP 80.000 de base/.test(plt));
    await p.click('#pl-turnos .kbtn');
    await p.waitForTimeout(120);
    ok('al abrirlo, la base viene puesta' + A,
       (await p.inputValue('#mtu-paga')) === '80000');
    await p.fill('#mtu-propinas', '34000');
    await p.waitForTimeout(80);
    ok('y el total se arma solo' + A,
       /COP 114.000/.test(await p.textContent('#mtu-total')),
       await p.textContent('#mtu-total'));
    await p.evaluate(() => { window.LLAMADAS.length = 0; });
    await p.click('#m-turno .card-cta');
    await p.waitForTimeout(200);
    const gt = (await p.evaluate(() => window.LLAMADAS))
      .filter(l => l.accion === 'nc_soul_turno')[0];
    ok('guarda el turno con sus propinas' + A,
       gt && gt.datos.propinas === '34000' && gt.datos.fecha === '2026-09-18',
       JSON.stringify(gt && gt.datos));

    const ph = await p.textContent('#pl-hormiga');
    ok('lo hormiga se ve por categoría' + A,
       /Buses/.test(ph) && /Salidas/.test(ph) && /COP 65.000/.test(ph));
    ok('y una categoría sin gastos dice «nada», no cero' + A, /nada/.test(ph));
    await p.evaluate(() => avisoGlobal(''));

    // ══ Universidad ══
    await p.evaluate(() => go('uni', null));
    await p.waitForTimeout(180);
    const uni = await p.textContent('#v-uni');
    ok('lista las materias' + A, /Estadística/.test(uni) && /Química Ambiental/.test(uni));
    ok('con sus entregas abiertas y las vencidas' + A,
       /2 entregas abiertas/.test(uni) && /1 vencida/.test(uni));
    ok('dice dónde viven los archivos: en Drive, no aquí' + A,
       /guarda el enlace/.test(uni) && /nunca una copia/.test(uni));
    ok('y admite que todavía no entiende un PDF solo' + A,
       /todavía no hago/.test(uni) && /modelo de lenguaje/.test(uni));
    const carpeta = await p.$$eval('#uni-lista a', e => e.map(x => x.getAttribute('href')));
    ok('la carpeta abre en Drive, en otra pestaña' + A,
       carpeta.length === 1 && /drive\.google/.test(carpeta[0]), JSON.stringify(carpeta));
    ok('la materia sin carpeta ofrece enlazarla' + A, /Enlazar carpeta/.test(uni));

    // El lector: propone, no escribe
    await p.evaluate(() => { window.LLAMADAS.length = 0; abrirSilabo('m1'); });
    ok('el lector abre con la materia en el título' + A,
       /Estadística/.test(await p.textContent('#ms-titulo')));
    await p.fill('#ms-texto', 'Parcial 1 — 15 de octubre — 25%');
    await p.click('#ms-leer');
    await p.waitForTimeout(150);
    const prop = await p.textContent('#ms-paso2');
    ok('muestra lo que encontró' + A, /ENCONTRÉ 2 ENTREGAS/.test(prop));
    ok('avisa cuándo el año lo puso él' + A, /el año lo puse yo/.test(prop));
    ok('y cuándo la fecha venía de un rango' + A, /tomé el último día/.test(prop));
    ok('MUESTRA lo que ignoró, con el motivo' + A,
       /ESTAS LAS DEJÉ FUERA/.test(prop) && /Walpole/.test(prop) &&
       /porcentaje pero ninguna fecha/.test(prop));
    ok('y dice que no inventa fechas para completar la lista' + A,
       /No le invento una fecha/.test(prop));
    ok('leer NO guardó nada todavía' + A,
       (await p.evaluate(() => window.LLAMADAS)).filter(l => /guardar/.test(l.accion)).length === 0);
    ok('las dos vienen marcadas, y se pueden desmarcar' + A,
       await p.$$eval('.ms-ck', e => e.filter(x => x.checked).length) === 2);

    await p.evaluate(() => { document.getElementById('ms-ck1').checked = false; });
    await p.click('#ms-guardar');
    await p.waitForTimeout(200);
    const env = (await p.evaluate(() => window.LLAMADAS))
      .filter(l => l.accion === 'nc_soul_silabo_guardar')[0];
    ok('guarda SOLO la que quedó marcada' + A,
       env && env.datos.items.length === 1 && env.datos.items[0].titulo === 'Parcial 1',
       JSON.stringify(env && env.datos));
    ok('y la manda con su materia' + A, env && env.datos.materia === 'm1');
    ok('después avisa qué guardó y qué falta ponerle' + A,
       /Guardé 2/.test(await p.textContent('#av-global')) &&
       /horas que crees que cuestan/.test(await p.textContent('#av-global')),
       await p.textContent('#av-global'));
    await p.evaluate(() => avisoGlobal(''));

    // ══ El cielo ══
    await p.evaluate(() => go('cielo', null));
    await p.waitForTimeout(200);
    const ci = await p.textContent('#v-cielo');
    ok('dice qué momento es hoy' + A, /HOY ES TIEMPO DE/.test(ci) && /Cambiar/.test(ci));
    ok('y de dónde salió: su pensum, no la luna' + A,
       /tu pensum: Saturno por casa 10/.test(ci));
    ok('la luna de hoy, con su iluminación' + A, /96%/.test(ci) && /menguando/.test(ci));
    ok('la semana trae un momento por día' + A,
       (await p.$$eval('#ci-semana .sdia', e => e.length)) === 7);
    ok('dice que manda el pensum sobre la luna' + A,
       /tu marco antes que la lectura común/.test(ci));
    /**
     * Su división, con sus palabras. Si esto cambia, cambió su marco y
     * alguien tiene que haberlo decidido.
     */
    ok('la carta va por sus tres grupos' + A,
       /PERSONALES/.test(ci) && /SOCIALES/.test(ci) && /GENERACIONALES/.test(ci));
    ok('con lo que mira cada uno, en sus palabras' + A,
       /Lo inmediato, lo mío, lo propio/.test(ci) &&
       /Expansión y estructura/.test(ci) &&
       /Marcan época. Van más allá del ego/.test(ci));
    ok('Marte retrógrado se marca' + A, /\bR\b/.test(ci));
    ok('los tránsitos llevan su duración' + A, /91 días/.test(ci) && /3 días/.test(ci),
       (ci.match(/\d+ días?/g) || []).join(','));
    ok('la revolución solar sale con su ventana' + A,
       /28 años/.test(ci) && /14 sep/.test(ci) && /quedan 355 días/.test(ci));
    /**
     * La LECTURA la compone Nova, que es lo que ella pidió. No es el
     * texto que pegó: es la tabla aplicada a sus dos datos.
     */
    ok('Nova compone la lectura del año' + A,
       /EL AÑO PIDE/.test(ci) && /Descansar/.test(ci));
    ok('dice cómo entra al año' + A,
       /CÓMO ENTRAS AL AÑO/.test(ci) && /Se entra mostrándose/.test(ci));
    ok('y dónde va su atención' + A,
       /DÓNDE VA TU ATENCIÓN/.test(ci) && /Casa 11 · La gente/.test(ci));
    ok('los planetas del año van por sus grupos' + A,
       /Marte/.test(ci) && /Cómo te presentas/.test(ci));
    ok('y lo que ella escribió va aparte, marcado como suyo' + A,
       /Tú dijiste:/.test(ci) && /El año de mostrar/.test(ci));
    await p.evaluate(() => abrirRevolucion());
    ok('el formulario dice que no puede leer una captura' + A,
       /no hace nada/.test(await p.textContent('#m-revolucion')) &&
       /modelo con visión/.test(await p.textContent('#m-revolucion')));
    ok('y pide la casa de cada planeta' + A,
       (await p.$$eval('#mrv-planetas input', e => e.length)) >= 9);
    await p.evaluate(() => cerrarModal('m-revolucion'));
    /** Lo que ninguna app hace: medir si le funcionó A ELLA. */
    ok('mide si le funcionó a ella' + A,
       /SI TE FUNCIONÓ A TI/.test(ci) && /cuarto menguante las terminaste el 80%/.test(ci));
    ok('y dice que es su dato, no una creencia' + A,
       /no una creencia prestada/.test(ci));
    ok('una fase sin muestra dice cuántas faltan, no un porcentaje' + A,
       /faltan 5/.test(ci));

    // El cargador de la carta
    await p.evaluate(() => { window.LLAMADAS.length = 0; abrirCarta(); });
    ok('el cargador trae sus datos de nacimiento' + A,
       (await p.inputValue('#mc-fecha')) === '1998-09-14');
    ok('y dice que no calcula la carta, que la trae de Horus' + A,
       /Horus ya lo hace bien/.test(await p.textContent('#m-carta')));
    await p.fill('#mc-texto', 'Sol en Virgo 21°34 Casa 2');
    await p.click('#mc-leer');
    await p.waitForTimeout(150);
    const propC = await p.textContent('#mc-paso2');
    ok('propone lo que entendió' + A, /ENTENDÍ 2 CUERPOS/.test(propC));
    ok('avisa qué planetas le faltaron' + A,
       /Me faltaron: Luna, Mercurio/.test(propC) && /peor que una carta vacía/.test(propC));
    ok('y muestra lo que dejó fuera' + A, /Generado por Horus/.test(propC));
    ok('leer NO guardó nada' + A,
       (await p.evaluate(() => window.LLAMADAS)).filter(l => l.accion === 'nc_soul_carta').length === 0);
    await p.evaluate(() => cerrarModal('m-carta'));

    // ══ Nova Family ══
    await p.evaluate(() => go('family', null));
    await p.waitForTimeout(200);
    const fam = await p.textContent('#fam-cuerpo');
    ok('trae las alarmas de la tienda' + A, /4 novedades sin resolver/.test(fam));
    ok('y el semáforo de la pauta' + A,
       /EL SEMÁFORO DE LA PAUTA/.test(fam) && /CPA contra el techo/.test(fam));
    ok('dice que es el mismo del correo del lunes' + A, /correo los lunes/.test(fam));
    ok('resume Nova Central' + A, /Cobros atrasados/.test(fam) && /COP 800.000/.test(fam));
    ok('y dice la verdad de novAcademy en vez de un cero mudo' + A,
       /Todavía no has dado de alta/.test(fam));
    /**
     * La regla de toda Nova: nunca dos tiendas en la misma pantalla.
     * El selector las ofrece, el cuerpo muestra una.
     */
    ok('NUNCA muestra dos tiendas a la vez' + A,
       fam.indexOf('Nutrea EC') === -1 && /Nutrea GT/.test(fam));
    ok('pero deja cambiar de tienda' + A,
       (await p.$$eval('#fam-sel .kbtn', e => e.map(x => x.textContent))).join(',') === 'Nutrea GT,Nutrea EC');

    // ══ «Acción desconocida» se traduce a qué hacer ══
    const viejo = await p.evaluate(async () => {
      const antes = window.nc;
      window.nc = async () => ({ ok: false, error: 'Acción desconocida: nc_soul' });
      (0, eval)('nc = window.nc;');
      await cargar();
      const t = document.getElementById('av-global').textContent;
      window.nc = antes; (0, eval)('nc = window.nc;');
      await cargar();
      return t;
    });
    ok('«Acción desconocida» dice que falta publicar, y cómo' + A,
       /c(ó|o)digo viejo/.test(viejo) && /Nueva/.test(viejo) && /bootstrapTodo/.test(viejo),
       viejo.slice(0, 100));

    // ══ Un fallo al dibujar no se disfraza de falta de internet ══
    const msg = await p.evaluate(async () => {
      const viejo = window.pintarHoy;
      window.pintarHoy = () => { throw new Error('a propósito'); };
      (0, eval)('pintarHoy = window.pintarHoy;');
      await cargar();
      const t = document.getElementById('av-global').textContent;
      (0, eval)('pintarHoy = ' + viejo.toString() + ';');
      return t;
    });
    ok('un fallo al dibujar NO dice «sin conexión»' + A,
       /no los pude dibujar/.test(msg) && !/conexión/.test(msg) && !/hablar con la hoja/.test(msg),
       msg.slice(0, 90));
    await p.evaluate(() => cargar());
    await p.waitForTimeout(120);

    const lateral = await p.evaluate(() =>
      document.documentElement.scrollWidth - document.documentElement.clientWidth);
    ok('sin scroll lateral' + A, lateral === 0, lateral + 'px');

    if (ancho === 1200) {
      for (const [v, f] of [['cielo', 'soul-cielo'], ['rutina', 'soul-rutina'], ['uni', 'soul-uni'], ['hoy', 'soul-hoy'], ['pendientes', 'soul-tablero'],
                            ['semana', 'soul-semana'], ['plata', 'soul-plata'],
                            ['family', 'soul-family'], ['mindlab', 'soul-mindlab']]) {
        await p.evaluate((x) => go(x, null), v);
        await p.waitForTimeout(120);
        await p.locator('#v-' + v).screenshot({ path: OUT + f + '.png' });
      }
    }
    await p.close();
  }

  await b.close();
  console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
  process.exit(fallas ? 1 : 0);
})();
