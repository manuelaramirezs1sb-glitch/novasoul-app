/**
 * Un proyecto a fondo, y su plata.
 *
 * Lo que se comprueba: que la utilidad neta reste TODO lo que dijo que
 * resta; que un costo fijo vacío no se trate como cero en silencio; que
 * un mes en pérdida no invente una «parte» negativa; y —la que más
 * importa— que el texto de un proyecto confidencial NO suba a Central.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');

const LIBROS = { cen: {}, s: {}, a: {}, nutrea: {} };
function libro(id) {
  const hojas = LIBROS[id] || {};
  return { getSheetByName: (nombre) => {
    const m = hojas[nombre];
    if (!m) return null;
    return {
      getLastRow: () => m.length,
      getLastColumn: () => (m[0] ? m[0].length : 0),
      getDataRange: () => ({ getValues: () => m.map(f => f.slice()) }),
      appendRow: (f) => m.push(f.slice()),
      deleteRow: (n) => m.splice(n - 1, 1),
      getRange: (f, c, nf, nc) => ({
        getValues: () => {
          const out = [];
          for (let i = 0; i < (nf || 1); i++) {
            const fila = m[f - 1 + i] || [];
            out.push(fila.slice(c - 1, c - 1 + (nc || fila.length)));
          }
          return out;
        },
        setValues: (v) => { v.forEach((fila, i) => { m[f - 1 + i] = fila.slice(); }); },
        setValue: () => {},
      }),
    };
  } };
}
const PROPS = { ID_EMPRESARIAL: 'emp', ID_CENTRAL: 'cen', ID_SOUL: 's', ID_ACADEMY: 'a' };
const HOY = '2026-09-23T12:00:00Z';
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: (k) => PROPS[k] || '', getProperties: () => PROPS,
  setProperty: () => {}, deleteProperty: () => {} }) };
global.SpreadsheetApp = { openById: (id) => libro(id), flush: () => {} };
global.Logger = { log: () => {} };
global.ScriptApp = { getProjectTriggers: () => [], EventType: { CLOCK: 'CLOCK' }, WeekDay: { MONDAY: 'MONDAY' } };
global.Session = { getScriptTimeZone: () => 'UTC', getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {}, remove: () => {} }) };
global.MailApp = { sendEmail: () => {} };
global.UrlFetchApp = { fetch: () => ({ getContentText: () => '{"data":[]}' }) };
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

(0, eval)(src + '\n;globalThis.__F = { utilidadMes_, parteDelMes_, centralProyecto,' +
  ' centralFuenteGuardar, centralFuenteBorrar, centralProyectoLeer,' +
  ' centralProyectoGuardarTareas, centralMio, proyRol_, proyModalidad_,' +
  ' proyConfidencial_, tiendasParaProyecto_, PROY_BASES };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const C_TRAB = ['id','nombre','contraparte','tipo','estado','moneda','valor_acordado',
  'forma_cobro','fecha_inicio','fecha_entrega','horas_semana','especificacion','documento',
  'nota','mi_rol','modalidad','porcentaje','base_porcentaje','cliente_id','tienda_id','confidencial'];
const C_FUE = ['id','trabajo_id','nombre','tipo','enlace','nota','agregado_en'];
const C_COBR = ['id','trabajo_id','concepto','monto','moneda','fecha_esperada','fecha_cobrada','estado','nota'];
const C_FIN = ['id','fecha','flujo','categoria','concepto','monto','moneda','cuenta','recurrente','trabajo_id','nota'];
const C_PEND = ['id','usuario_id','texto','tipo','origen','fecha','hecho','hecho_en',
  'plataforma_id','trabajo_id','estado','prioridad','horas_estimadas','horas_reales',
  'riesgo','nota','materia_id'];
const YO = 'manuela@nova.com';

// Un pedido de Nutrea EC: valor, costo producto, costo envío, estado
const ped = (fecha, estado, valor, cp, ce) =>
  ['p', '', '', fecha, 'EC', '', '', '', '', '', '', '', '', '', '', 'Colágeno', '', 1,
   valor, cp, ce, '', '', estado, '', estado, '', '', '', '', '', '', '', '', '', '', '', '', '', '', ''];
const C_PED = ['id','fuente','id_externo','fecha','tienda','cliente','cedula','correo',
  'telefono','telefono_norm','telefono_2','telefono_2_norm','ciudad','departamento','direccion',
  'producto','sku','cantidad','valor','costo_producto','costo_envio','metodo_pago','bodega',
  'estado','estado_transportadora','estado_canonico','transportadora','guia','intentos',
  'gestora_asignada','fecha_promesa','fecha_entrega','razon_cancelacion','estado_nova','nota',
  'ultimo_movimiento','adelanto','acuerdo_oficina','confirmado_oficina','actualizado_en','actualizado_por'];
const C_PAU = ['id','fecha','fecha_fin','tienda','plataforma','cuenta','campana','conjunto',
  'entrega','presupuesto','gasto','moneda_gasto','gasto_normalizado','impresiones','alcance',
  'clics','ctr','cpc','cpm','frecuencia','resultados','costo_resultado','actualizado_en'];

function sembrarHojas(costosFijos) {
  /**
   * Cada caso de prueba es una PETICIÓN nueva, y una petición real
   * empieza soltando lo que Nova tenga en memoria de la anterior
   * (`manejar()` lo hace). Aquí hay que decirlo a mano porque las
   * pruebas llaman a las funciones directamente, sin pasar por ahí.
   *
   * Sin esto, el manejador del libro y las hojas ya leídas seguirían
   * apuntando a los datos del caso anterior, y la prueba mediría un
   * Nova que no existe.
   */
  libroOlvidar_(); soulOlvidar_();

  UUID = 0;
  LIBROS.nutrea = {
    Tiendas: [['id','nombre','pais','moneda','estado'],
              ['EC', 'Nutrea EC', 'EC', 'USD', 'activa'],
              ['GT', 'Nutrea GT', 'GT', 'GTQ', 'activa']],
    Pedidos: [C_PED,
      // Septiembre: 3 entregados y 1 devolución
      ped('2026-09-03', 'entregado', 100, 30, 8),
      ped('2026-09-08', 'entregado', 100, 30, 8),
      ped('2026-09-15', 'entregado', 100, 30, 8),
      ped('2026-09-18', 'devolucion', 100, 30, 8),
    ],
    Pauta: [C_PAU,
      ['g1', '2026-09-05', '', 'EC', 'meta', '', '', 'Conj A', '', 50, 40, 'USD', 40,
       1000, 900, 50, 5, 40, 1.1, 3, 13, ''],
      ['g2', '2026-09-12', '', 'EC', 'meta', '', '', 'Conj A', '', 50, 26, 'USD', 26,
       1000, 900, 50, 5, 40, 1.1, 2, 13, ''],
    ],
    Parametros: [['tienda','clave','valor','nota','actualizado_en'],
                 ['EC', 'costos_fijos_mes', costosFijos, '', '']],
    Tasas: [['fecha','origen','destino','tasa','fuente']],
    Equipo: [['id','nombre','correo','rol','tiendas','estado','permisos']],
  };
  LIBROS.cen = {
    Clientes: [['id','empresa','pais','plan','tarifa','costo','estado','fecha_alta',
                'fecha_corte','ultimo_pago','tickets_mes','usuarios','tiendas','sheet_id'],
               ['c1', 'Nutrea', 'EC', 'pro', 0, 0, 'activo', '', '', '', 0, 1, 2, 'nutrea']],
    Trabajos: [C_TRAB,
      ['t1', 'Nutrea EC', 'Sociedad', 'propio', 'activo', 'USD', 0, '', '', '', 0,
       'Mi 50% de la tienda de Ecuador', '', '',
       'socia', 'porcentaje', 50, 'utilidad_neta', 'c1', 'EC', 'no'],
      ['t2', 'PHH', 'Upwork', 'empleo', 'activo', 'USD', 0, '', '', '', 12, '', '', '',
       '', '', 0, '', '', '', ''],
      ['t3', 'Son de Sky', 'Salsabor', 'cliente', 'activo', 'COP', 2400000, '', '',
       '2026-09-30', 0, 'Carta, web y precios', 'https://ej.com/contrato', '',
       'trabajadora', 'fijo', 0, '', '', '', 'no'],
    ],
    Fuentes: [C_FUE],
    Cobros: [C_COBR,
      ['c1', 't3', 'Primer pago', 800000, 'COP', '2026-08-01', '2026-08-02', 'cobrado', ''],
      ['c2', 't3', 'Segundo pago', 800000, 'COP', '2026-09-05', '', 'pendiente', ''],
    ],
    Finanzas: [C_FIN,
      ['f1', '2026-08-02', 'ingreso', 'trabajo', 'Primer pago', 800000, 'COP', '', '', 't3', '']],
    Metas: [['id','tipo','nombre','con_quien','monto_meta','saldo','moneda','cuota',
             'dia_del_mes','fecha_meta','estado','nota']],
    Movimientos: [['fecha','quien','entidad','id','campo','antes','ahora']],
  };
  LIBROS.s = {
    Pendientes: [C_PEND,
      ['s1', YO, 'Ajustar precios de cocteles', '', '', '2026-09-25', '', '', '', 't3',
       'pendiente', 'media', 6, '', 'acordado', '', ''],
      ['s2', YO, 'Cerrar el informe secreto de PHH', '', '', '2026-09-24', '', '', '', 't2',
       'pendiente', 'alta', 5, '', 'acordado', '', ''],
    ],
    Materias: [['id','usuario_id','nombre','codigo','profesor','carpeta','semestre','trabajo_id','estado','nota']],
    Horas: [['usuario_id','dia_semana','horas_libres','nota']],
    Mindlab: [['id','usuario_id','semana','mes','tema','tarea','horas_estimadas','desde','hasta','estado','nota']],
    Fijos: [['id','usuario_id','categoria','concepto','monto','moneda','dia_del_mes','activo','nota']],
    Rutina: [['id','usuario_id','tipo','nombre','dia_semana','hora_inicio','hora_fin','lugar',
              'trabajo_id','materia_id','paga_fija','moneda','desde','hasta','activo','nota']],
    Turnos: [['id','usuario_id','rutina_id','fecha','paga','propinas','moneda','estado','finanza_id','nota']],
  };
  LIBROS.a = { Estudiantes: [['id','nombre']] };
}
const SOCIA = { correo: YO, nombre: 'Manuela', rol: 'socia' };
const OPERADORA = { correo: 'k@nova.com', nombre: 'Kat', rol: 'operadora' };

console.log('\n── La utilidad de la tienda, resta por resta ──');
sembrarHojas(200);
let u = F.utilidadMes_('nutrea', 'EC', '2026-09');
/**
 * 3 entregados × (100 − 30 − 8) = 186 de ganancia.
 * 1 devolución: el flete de 8 se pagó y no vuelve.
 * Pauta: 40 + 26 = 66.  Costos fijos: 200.
 */
igual('ganancia de los entregados', 186, u.ganancia);
igual('la devolución cuesta el flete', 8, u.costoDevoluciones);
igual('utilidad antes de pauta', 178, u.utilidad_antes_pauta);
igual('la pauta del mes', 66, u.pauta);
igual('utilidad después de pauta', 112, u.utilidad_post_pauta);
igual('y la neta, con costos fijos', -88, u.utilidad_neta);
igual('en la moneda de la tienda', 'USD', u.moneda);
ok('con costos fijos puestos, no falta nada',
   u.faltan.indexOf('costos_fijos_mes') === -1, JSON.stringify(u.faltan));

/**
 * Un costo fijo vacío NO es cero. Tratarlo como cero sube la utilidad,
 * y sobre esa utilidad se reparte plata de verdad: es el error que
 * nadie revisa porque da un número alto.
 */
sembrarHojas('');
u = F.utilidadMes_('nutrea', 'EC', '2026-09');
ok('un costo fijo vacío se DENUNCIA', u.faltan.indexOf('costos_fijos_mes') !== -1,
   JSON.stringify(u.faltan));
ok('y se dice que no los hay', u.hayCostosFijos === false);
igual('la neta queda igual a la de después de pauta, y se avisa', 112, u.utilidad_neta);

console.log('\n── Mi 50% ──');
sembrarHojas(50);
let t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-09' });
ok('el proyecto abre', t.ok === true, JSON.stringify(t.error));
igual('soy socia', 'socia', t.proyecto.rol);
igual('y cobro por porcentaje', 'porcentaje', t.proyecto.modalidad);
igual('el 50%', 50, t.proyecto.porcentaje);
igual('sobre la utilidad neta', 'utilidad_neta', t.proyecto.base);
igual('la base de este mes: 178 − 66 − 50', 62, t.parte.valorBase);
igual('mi parte son 31', 31, t.parte.parte);
igual('en dólares', 'USD', t.parte.moneda);
ok('y muestra la cuenta entera, no solo el resultado',
   t.parte.detalle.ganancia === 186 && t.parte.detalle.pauta === 66 &&
   t.parte.detalle.costosFijos === 50, JSON.stringify(t.parte.detalle).slice(0, 120));
ok('dice de qué tienda salió', t.parte.tiendaNombre === 'Nutrea EC', t.parte.tiendaNombre);

/**
 * Un mes en pérdida no reparte. Mostrar «tu parte: −44» invitaría a
 * restarlo de sus ingresos, y a ella no le cobran por un mes malo.
 */
sembrarHojas(300);
t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-09' });
igual('en pérdida, la base es negativa', -188, t.parte.valorBase);
igual('pero la parte es cero, no negativa', 0, t.parte.parte);
ok('y se dice que el mes fue en pérdida', t.parte.enPerdida === true);

sembrarHojas(50);
t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-05' });
ok('un mes sin pedidos dice que no hay datos', t.parte.hayDatos === false);
igual('y no inventa una parte', 0, t.parte.parte);

console.log('\n── Lo que falta para poder calcularlo ──');
sembrarHojas(50);
LIBROS.cen.Trabajos[1][19] = '';          // sin tienda enlazada
t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-09' });
ok('sin tienda enlazada lo dice, no calcula',
   t.parte.hay === false && /no está enlazado/.test(t.parte.porque), JSON.stringify(t.parte));
sembrarHojas(50);
LIBROS.cen.Trabajos[1][16] = 0;           // sin porcentaje
t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-09' });
ok('sin porcentaje, tampoco',
   t.parte.hay === false && /qué porcentaje/.test(t.parte.porque));

console.log('\n── La frontera: qué texto sube a Central ──');
sembrarHojas(50);
t = F.centralProyecto(SOCIA, { id: 't3', mes: '2026-09' });
ok('un proyecto normal SÍ muestra sus tareas',
   t.tareas.length === 1 && /cocteles/.test(t.tareas[0].texto), JSON.stringify(t.tareas));
ok('y marca cuál es de esta semana', t.tareas[0].estaSemana === true);
igual('con sus horas', 6, t.carga.horas);

const phh = F.centralProyecto(SOCIA, { id: 't2', mes: '2026-09' });
/**
 * PHH es un empleo y `confidencial` viene vacío. Vacío NO es «no»: para
 * un empleo es «sí». Un olvido no puede ser lo único que proteja esto.
 */
ok('un empleo nace confidencial aunque la celda esté vacía',
   phh.proyecto.confidencial === true);
igual('así que NO sube ni una tarea', 0, phh.tareas.length);
ok('pero sí sube cuánto pesa', phh.carga.abiertas === 1 && phh.carga.horas === 5,
   JSON.stringify(phh.carga));
ok('y el texto no aparece por ningún lado',
   JSON.stringify(phh).indexOf('secreto') === -1);

LIBROS.cen.Trabajos[2][20] = 'no';
const phh2 = F.centralProyecto(SOCIA, { id: 't2', mes: '2026-09' });
ok('si ella lo apaga a propósito, entonces sí sube',
   phh2.tareas.length === 1 && /secreto/.test(phh2.tareas[0].texto));

console.log('\n── Los cobros y la plata del proyecto ──');
sembrarHojas(50);
t = F.centralProyecto(SOCIA, { id: 't3', mes: '2026-09' });
igual('dos cobros', 2, t.cobros.length);
igual('uno cobrado', { COP: 800000 }, t.plata.cobrado);
igual('uno por cobrar', { COP: 800000 }, t.plata.porCobrar);
igual('y está atrasado', { COP: 800000 }, t.plata.atrasado);
igual('falta la mitad del valor acordado', 1600000, t.plata.falta);
igual('el movimiento del proyecto está ahí', 1, t.movimientos.length);

LIBROS.cen.Trabajos[3][6] = 0;
t = F.centralProyecto(SOCIA, { id: 't3', mes: '2026-09' });
igual('sin valor acordado, «falta» se deja vacío en vez de inventar', null, t.plata.falta);

console.log('\n── Las fuentes: enlaces, no archivos ──');
sembrarHojas(50);
ok('se agrega una fuente',
   F.centralFuenteGuardar(SOCIA, { datos: { trabajo_id: 't3', nombre: 'Carta actual',
     tipo: 'excel', enlace: 'https://docs.google.com/x' } }).ok);
t = F.centralProyecto(SOCIA, { id: 't3', mes: '2026-09' });
igual('queda en el proyecto', 1, t.fuentes.length);
ok('y dice que el Excel SÍ se puede leer', t.fuentes[0].lee === true);
ok('se agrega un PDF, y dice que no se lee todavía',
   F.centralFuenteGuardar(SOCIA, { datos: { trabajo_id: 't3', nombre: 'Contrato',
     tipo: 'pdf', enlace: 'https://drive.google.com/y' } }).ok &&
   F.centralProyecto(SOCIA, { id: 't3' }).fuentes.filter(f => f.tipo === 'pdf')[0].lee === false);
ok('un enlace que no es enlace se rechaza',
   /empezar por http/.test(F.centralFuenteGuardar(SOCIA,
     { datos: { trabajo_id: 't3', nombre: 'x', enlace: 'mi archivo' } }).error));
ok('un tipo inventado se rechaza',
   F.centralFuenteGuardar(SOCIA, { datos: { trabajo_id: 't3', nombre: 'x',
     tipo: 'cassette' } }).ok === false);
ok('sin proyecto no se crea',
   /de qué proyecto/.test(F.centralFuenteGuardar(SOCIA, { datos: { nombre: 'x' } }).error));

console.log('\n── Pegar el alcance y sacar las entregas ──');
sembrarHojas(50);
let r = F.centralProyectoLeer(SOCIA, { texto:
  'Entrega de la carta impresa: 10 de octubre\n' +
  'Recetario de cocteles 24 de octubre\n' +
  'El pago va contra entrega' });
igual('encuentra las dos con fecha', 2, r.encontradas.length);
ok('y muestra la que ignoró', r.ignoradas.length === 1,
   JSON.stringify(r.ignoradas));

let g = F.centralProyectoGuardarTareas(SOCIA, { trabajo: 't3', items: r.encontradas });
igual('guarda las dos', 2, g.creadas);
/**
 * Se guardan en NovaSoul aunque se creen desde Central: las entregas
 * con fecha son del día a día. Si vivieran en las dos, la semana se
 * contaría dos veces.
 */
igual('y viven en NovaSoul, no en Central', 4, LIBROS.s.Pendientes.length - 1);
t = F.centralProyecto(SOCIA, { id: 't3', mes: '2026-09' });
igual('el proyecto ya las ve', 3, t.tareas.length);
ok('con el riesgo de un cliente: se corre avisando',
   t.tareas.every(x => x.riesgo !== 'inamovible'));
igual('pegarlo otra vez no duplica', 0,
      F.centralProyectoGuardarTareas(SOCIA, { trabajo: 't3', items: r.encontradas }).creadas);

console.log('\n── La puerta ──');
ok('una operadora no abre un proyecto',
   F.centralProyecto(OPERADORA, { id: 't1' }).ok === false);
ok('ni agrega fuentes',
   F.centralFuenteGuardar(OPERADORA, { datos: { trabajo_id: 't3', nombre: 'x' } }).ok === false);
ok('ni pega alcances',
   F.centralProyectoLeer(OPERADORA, { texto: 'x' }).ok === false);
ok('un proyecto que no existe lo dice',
   /No encontré/.test(F.centralProyecto(SOCIA, { id: 'zz' }).error));

console.log('\n── Lo deducido se marca como deducido ──');
sembrarHojas(50);
t = F.centralProyecto(SOCIA, { id: 't2', mes: '2026-09' });
igual('un empleo sin rol escrito se deduce trabajadora', 'trabajadora', t.proyecto.rol);
ok('y se dice que fue deducido', t.proyecto.rolDeducido === true);
t = F.centralProyecto(SOCIA, { id: 't1', mes: '2026-09' });
ok('lo que sí está escrito NO se marca', t.proyecto.rolDeducido === false);

console.log('\n── La lista de Central lleva cómo cobra cada uno ──');
const mio = F.centralMio(SOCIA, {});
const nut = mio.trabajos.filter(x => x.id === 't1')[0];
igual('el porcentaje viaja a la lista', 50, nut.porcentaje);
igual('y el rol', 'socia', nut.rol);
ok('PHH sale marcado como confidencial',
   mio.trabajos.filter(x => x.id === 't2')[0].confidencial === true);

console.log('\n── Las tiendas que se pueden enlazar ──');
const tiendas = F.tiendasParaProyecto_();
igual('las dos de Nutrea', ['EC', 'GT'], tiendas.map(x => x.id));
ok('con su nombre de pantalla',
   tiendas[0].nombre === 'Nutrea EC' && tiendas[1].nombre === 'Nutrea GT',
   JSON.stringify(tiendas.map(x => x.nombre)));

console.log(fallas ? '\n' + fallas + ' FALLA(S)' : '\nTodo pasa.');
process.exit(fallas ? 1 : 0);
