/**
 * ═══════════════════════════════════════════════════════════
 *  NOVA CENTRAL · EL LADO DE ELLA
 * ═══════════════════════════════════════════════════════════
 *
 * Central era la consola del producto. Ahora también es el centro de
 * trabajo de Manuela, que además de Nova tiene clientes propios, un
 * empleo y la universidad.
 *
 * LA FRONTERA CON NOVASOUL, QUE ES DE PRIVACIDAD
 *
 * Aquí viven los trabajos y la plata: contratos, cotizaciones, lo que
 * debe entrar, lo que debe salir. En NovaSoul vive el día a día — el
 * ciclo, las comidas, el ánimo, las recompensas.
 *
 * El corte no es por función, es por quién puede ver la pantalla.
 * Central se abre delante de una socia o un contador; NovaSoul no se
 * abre delante de nadie. Por eso la información sube de Soul a Central
 * —entregas, fechas, plata— y nunca baja al revés.
 *
 * TODO ESTO ES SOLO DE LA SOCIA
 *
 * Una operadora crea clientes y los acompaña; no tiene por qué ver
 * cuánto le pagan a Manuela por un proyecto de una cafetería. El
 * permiso que manda es `facturacion`, el mismo que ya guardaba las
 * tarifas de los clientes.
 */

/** Las cuatro hojas de este módulo, y quién puede tocarlas. */
const MIO_HOJAS = {
  Trabajos: ['id','nombre','contraparte','tipo','estado','moneda',
             'valor_acordado','forma_cobro','fecha_inicio','fecha_entrega',
             'horas_semana','especificacion','documento','nota',
             'mi_rol','modalidad','porcentaje','base_porcentaje',
             'cliente_id','tienda_id','confidencial','padre_id',
             'periodicidad','dias_pago'],
  Fuentes:  ['id','trabajo_id','nombre','tipo','enlace','nota','agregado_en'],
  Cobros:   ['id','trabajo_id','concepto','monto','moneda',
             'fecha_esperada','fecha_cobrada','estado','nota'],
  Finanzas: ['id','fecha','flujo','categoria','concepto','monto','moneda',
             'cuenta','recurrente','trabajo_id','nota'],
  Metas:    ['id','tipo','nombre','con_quien','monto_meta','saldo','moneda',
             'cuota','dia_del_mes','fecha_meta','estado','nota'],
};

/**
 * Los tipos de trabajo, y si se cobran.
 *
 * La universidad y lo propio no facturan, pero ocupan las mismas horas
 * que un cliente que sí. Una lista que solo mira lo que entra deja
 * fuera justo lo que no se puede incumplir.
 */
const TIPOS_TRABAJO = {
  cliente:  { nombre: 'Cliente',        cobra: true  },
  empleo:   { nombre: 'Empleo',         cobra: true  },
  propio:   { nombre: 'Propio (Nova)',  cobra: false },
  estudio:  { nombre: 'Universidad',    cobra: false },
};

/**
 * ═══════════════════════════════════════════════════════════
 *  TRABAJOS Y PROYECTOS: DOS FAMILIAS, DOS PANTALLAS
 * ═══════════════════════════════════════════════════════════
 *
 * Ella lo dijo así: «entre proyectos pueden entrar los que no son
 * pagos, como los de Nova; en trabajos todo lo que me da ingresos y
 * beneficios económicos».
 *
 * La línea es una sola y es clara: ¿esto entra plata o no?
 *
 *   TRABAJO    Upwork, PHH, Nutrea, un cliente. Tiene moneda, forma de
 *              cobro y fecha de pago.
 *   PROYECTO   Nova, novAcademy, la universidad. Ocupa horas, no paga.
 *              No tiene moneda porque no hay nada que convertir.
 *
 * ── LO QUE NO SE DEDUCE, SE PREGUNTA ──
 *
 * Ella pidió repartir lo que ya está cargado «y lo que no esté claro
 * para Nova que me pregunte cuando entre». Eso es lo contrario de lo
 * fácil: lo fácil sería mandar todo lo dudoso a Trabajos y que ella lo
 * descubra un día mirando una cifra rara.
 *
 * Un caso dudoso de verdad: un «cliente» sin valor, sin porcentaje y
 * sin forma de cobro. Puede ser un cliente que todavía no negoció
 * precio, o un favor que nunca va a pagar. Nova no puede saberlo, y
 * adivinar mal cambia si esa fila suma o no a lo que le deben.
 */
const PROY_FAMILIAS = {
  trabajo:  { nombre: 'Trabajos',  que: 'Lo que te da ingresos' },
  proyecto: { nombre: 'Proyectos', que: 'Lo que no paga pero ocupa horas' },
};

function familiaDe_(t) {
  const tipo = norm(t.tipo);
  const mod = norm(t.modalidad);
  const tienePlata = num(t.valor_acordado) > 0 || num(t.porcentaje) > 0;

  // La universidad y lo propio no pagan, y eso no admite discusión.
  if (tipo === 'estudio') {
    return { familia: 'proyecto', universidad: true, claro: true,
             porque: 'La universidad no paga.' };
  }
  if (tipo === 'propio') {
    return { familia: 'proyecto', universidad: false, claro: true,
             porque: 'Lo propio no factura.' };
  }
  // Lo que dice explícitamente que no cobra, tampoco.
  if (mod === 'sin_cobro') {
    return { familia: 'proyecto', universidad: false, claro: true,
             porque: 'Está marcado como sin cobro.' };
  }
  // Un empleo paga siempre, por definición.
  if (tipo === 'empleo') {
    return { familia: 'trabajo', universidad: false, claro: true,
             porque: 'Un empleo paga.' };
  }
  // Un cliente con plata puesta, o con forma de cobro escrita, es trabajo.
  if (tipo === 'cliente' && (tienePlata || mod === 'porcentaje' ||
                             mod === 'fijo' || mod === 'por_hora')) {
    return { familia: 'trabajo', universidad: false, claro: true,
             porque: 'Es un cliente con forma de cobro.' };
  }
  /**
   * Y aquí es donde Nova se calla y pregunta.
   *
   * Se pone del lado de TRABAJO mientras ella decide —no desaparece de
   * la lista— pero marcado, para que la pregunta se vea antes de que
   * la cifra confunda.
   */
  return {
    familia: 'trabajo', universidad: false, claro: false,
    porque: !tipo
      ? 'No dice qué tipo de compromiso es.'
      : 'Es un cliente sin valor, sin porcentaje y sin forma de cobro: ' +
        'no sé si te va a pagar.',
  };
}

function mioSheet_(nombre) {
  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName(nombre);
  if (!sh) {
    throw new Error('Falta la hoja ' + nombre + ' en Nova_Central. ' +
                    'Corre bootstrapTodo() una vez.');
  }
  return sh;
}

/** Lee una hoja entera como objetos. */
function mioLeer_(nombre) {
  const sh = mioSheet_(nombre);
  if (sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const enc = d[0].map(norm);
  return d.slice(1).filter(function (f) { return String(f[0]).trim(); })
    .map(function (f) {
      const o = {};
      enc.forEach(function (c, i) { o[c] = f[i]; });
      return o;
    });
}

/**
 * Escribe o crea una fila, por id.
 *
 * Una sola puerta para las cuatro hojas: la pantalla manda la hoja y el
 * objeto, y el servidor decide qué columnas existen. Si la pantalla
 * mandara una columna que no está en el esquema, se ignora — no se
 * inventa una columna nueva porque alguien escribió mal un nombre.
 */
function mioGuardar_(nombre, datos) {
  const cols = MIO_HOJAS[nombre];
  if (!cols) throw new Error('No se puede escribir en ' + nombre + '.');

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) throw new Error('Hay otro cambio guardándose. Un segundo.');
  try {
    const sh = mioSheet_(nombre);
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const id = String(datos.id || '').trim();

    if (!id) {
      const nuevo = 'x' + Utilities.getUuid().slice(0, 8);
      datos.id = nuevo;
      sh.appendRow(enc.map(function (c) {
        return datos[c] !== undefined ? datos[c] : '';
      }));
      return datos;
    }

    const d = sh.getDataRange().getValues();
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][0]).trim() !== id) continue;
      enc.forEach(function (c, j) {
        // `undefined` es "no lo toques"; vacío sí borra, que es distinto.
        if (datos[c] !== undefined) d[i][j] = datos[c];
      });
      sh.getRange(i + 1, 1, 1, enc.length).setValues([d[i]]);
      const o = {};
      enc.forEach(function (c, j) { o[c] = d[i][j]; });
      return o;
    }
    throw new Error('No existe ' + nombre + ' con id ' + id + '.');
  } finally {
    lock.releaseLock();
  }
}

function mioBorrar_(nombre, id) {
  const sh = mioSheet_(nombre);
  const d = sh.getDataRange().getValues();
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][0]).trim() === String(id).trim()) {
      sh.deleteRow(i + 1);
      return true;
    }
  }
  return false;
}

// ─── LA FOTO ─────────────────────────────────────────────────

/**
 * Todo el lado de ella en una sola llamada.
 *
 * En una sola a propósito: son cuatro hojas de la misma carpeta y
 * pedirlas por separado serían cuatro viajes al servidor para pintar
 * una pantalla. El mismo error que hace lenta a Empresarial, que abre
 * con quince.
 */
function centralMio(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }

  const hoy = ahoraISO().slice(0, 10);
  const trabajos = mioLeer_('Trabajos');
  const cobros = mioLeer_('Cobros');
  const finanzas = mioLeer_('Finanzas');
  const metas = mioLeer_('Metas');

  /**
   * Cuánto pesa cada proyecto en NovaSoul: NÚMEROS, no tareas.
   *
   * Es la otra mitad del puente. Central le presta a Soul la lista de
   * proyectos y Soul le devuelve cuántas entregas abiertas tiene cada
   * uno y cuántas horas suman — nunca qué dicen. Lo de PHH es
   * confidencial y Central es la pantalla que algún día se abre delante
   * de una socia o un contador.
   */
  const carga = soulCargaPorTrabajo_(String(s.correo || '').toLowerCase());

  /**
   * Lo que falta cobrar, y lo que YA está tarde.
   *
   * Se separan porque son dos conversaciones distintas: una es esperar
   * y la otra es llamar. Juntarlas en un solo total esconde la segunda,
   * que es la única que pide hacer algo hoy.
   */
  const porCobrar = [], atrasados = [];
  cobros.forEach(function (c) {
    if (norm(c.estado) === 'cobrado' || c.fecha_cobrada) return;
    const esperada = aISO(c.fecha_esperada, 'UTC') || '';
    const fila = {
      id: c.id, trabajo_id: c.trabajo_id, concepto: c.concepto,
      monto: num(c.monto), moneda: String(c.moneda || '').toUpperCase(),
      esperada: esperada,
      dias: esperada ? Math.floor((new Date(hoy + 'T00:00:00Z') -
                                   new Date(esperada + 'T00:00:00Z')) / 86400000) : null,
      trabajo: (trabajos.filter(function (t) { return t.id === c.trabajo_id; })[0] || {}).nombre || '',
    };
    if (fila.dias !== null && fila.dias > 0) atrasados.push(fila);
    else porCobrar.push(fila);
  });

  // Los totales van POR MONEDA, sin sumarlas entre sí. Sumar dólares y
  // pesos con una tasa de hoy daría una cifra que cambia sola mañana.
  const totalPorMoneda = function (filas) {
    const m = {};
    filas.forEach(function (f) {
      const k = f.moneda || '?';
      m[k] = (m[k] || 0) + (f.monto || 0);
    });
    return m;
  };

  // El mes corriente de sus finanzas personales
  const mes = hoy.slice(0, 7);
  const delMes = finanzas.filter(function (f) {
    return String(aISO(f.fecha, 'UTC') || '').slice(0, 7) === mes;
  });
  const ingresos = {}, gastos = {};
  delMes.forEach(function (f) {
    const k = String(f.moneda || '?').toUpperCase();
    const destino = norm(f.flujo) === 'ingreso' ? ingresos : gastos;
    destino[k] = (destino[k] || 0) + Math.abs(num(f.monto));
  });

  return {
    ok: true,
    hoy: hoy, mes: mes,
    tipos: TIPOS_TRABAJO,
    trabajos: trabajos.map(function (t) {
      const suyos = cobros.filter(function (c) { return c.trabajo_id === t.id; });
      const cobrado = suyos.filter(function (c) { return norm(c.estado) === 'cobrado' || c.fecha_cobrada; })
                           .reduce(function (a, c) { return a + num(c.monto); }, 0);
      return {
        id: t.id, nombre: t.nombre, contraparte: t.contraparte,
        tipo: norm(t.tipo) || 'cliente', estado: norm(t.estado) || 'activo',
        moneda: String(t.moneda || '').toUpperCase(),
        valor: num(t.valor_acordado), cobrado: cobrado,
        // Lo que falta es un dato, no una promesa: si no hay valor
        // acordado se deja vacío en vez de restar contra cero.
        falta: num(t.valor_acordado) ? num(t.valor_acordado) - cobrado : null,
        entrega: aISO(t.fecha_entrega, 'UTC') || '',
        horasSemana: num(t.horas_semana),
        especificacion: t.especificacion || '',
        documento: t.documento || '',
        nota: t.nota || '',
        rol: proyRol_(t), modalidad: proyModalidad_(t),
        porcentaje: num(t.porcentaje), tiendaId: String(t.tienda_id || ''),
        confidencial: proyConfidencial_(t),
        // A cuál de las dos pantallas pertenece, y si Nova está segura.
        familia: familiaDe_(t).familia,
        universidad: familiaDe_(t).universidad,
        claro: familiaDe_(t).claro,
        porqueFamilia: familiaDe_(t).porque,
        // De quién cuelga, cuando es un proyecto dentro de un trabajo.
        padreId: String(t.padre_id || ''),
        tareas: carga[t.id] || { abiertas: 0, horas: 0, vencidas: 0, proxima: '' },
      };
    }),
    familias: PROY_FAMILIAS,
    /**
     * Lo que Nova no supo clasificar. Va en su propia lista para que la
     * pantalla lo pregunte al entrar, en vez de esconderlo en medio de
     * los demás donde nadie lo mira.
     */
    porClasificar: trabajos.filter(function (t) { return !familiaDe_(t).claro; })
      .map(function (t) {
        return { id: String(t.id || ''), nombre: String(t.nombre || ''),
                 porque: familiaDe_(t).porque };
      }),
    porCobrar: porCobrar.sort(function (a, b) { return (a.esperada || '9') < (b.esperada || '9') ? -1 : 1; }),
    atrasados: atrasados.sort(function (a, b) { return b.dias - a.dias; }),
    totales: {
      porCobrar: totalPorMoneda(porCobrar),
      atrasado: totalPorMoneda(atrasados),
      ingresosMes: ingresos,
      gastosMes: gastos,
    },
    finanzas: delMes.map(function (f) {
      return { id: f.id, fecha: aISO(f.fecha, 'UTC') || '', flujo: norm(f.flujo),
               categoria: f.categoria || '', concepto: f.concepto || '',
               monto: num(f.monto), moneda: String(f.moneda || '').toUpperCase(),
               cuenta: f.cuenta || '', trabajo_id: f.trabajo_id || '' };
    }).sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; }),
    metas: metas.map(function (m) {
      const meta = num(m.monto_meta), saldo = num(m.saldo);
      return {
        id: m.id, tipo: norm(m.tipo) || 'ahorro', nombre: m.nombre,
        conQuien: m.con_quien || '', meta: meta, saldo: saldo,
        moneda: String(m.moneda || '').toUpperCase(),
        cuota: num(m.cuota), dia: num(m.dia_del_mes),
        fechaMeta: aISO(m.fecha_meta, 'UTC') || '',
        // Sin meta no hay porcentaje. Un 0% inventado se lee como
        // "no has avanzado nada", que es una afirmación, no un vacío.
        pct: meta ? Math.round(saldo / meta * 100) : null,
        estado: norm(m.estado) || 'activa',
      };
    }),
  };
}

function centralMioGuardar(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const hoja = String(p.hoja || '');
  if (!MIO_HOJAS[hoja]) return { ok: false, error: 'No se puede escribir en ' + hoja + '.' };

  const datos = p.datos || {};
  if (hoja === 'Trabajos' && !String(datos.nombre || '').trim()) {
    return { ok: false, error: 'El trabajo necesita un nombre.' };
  }
  if (hoja === 'Cobros' && !num(datos.monto)) {
    return { ok: false, error: 'Un cobro sin monto no se puede seguir.' };
  }
  if (hoja === 'Finanzas' && !num(datos.monto)) {
    return { ok: false, error: 'Falta el monto.' };
  }

  try {
    const fila = mioGuardar_(hoja, datos);
    registrarCentral(s, hoja, fila.id, datos.id ? 'editado' : 'creado', '',
                     String(fila.nombre || fila.concepto || ''));
    return { ok: true, fila: fila };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function centralMioBorrar(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  const hoja = String(p.hoja || '');
  if (!MIO_HOJAS[hoja]) return { ok: false, error: 'No se puede borrar en ' + hoja + '.' };
  try {
    const fue = mioBorrar_(hoja, p.id);
    if (fue) registrarCentral(s, hoja, p.id, 'borrado', '', '');
    return { ok: fue, error: fue ? '' : 'No lo encontré.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

/** Marcar un cobro como recibido, con su fecha real. */
function centralMioCobrar(s, p) {
  if (!puedeCentral(s, 'facturacion')) {
    return { ok: false, error: 'Esta parte es solo de la socia.' };
  }
  try {
    const fila = mioGuardar_('Cobros', {
      id: p.id,
      estado: 'cobrado',
      fecha_cobrada: String(p.fecha || ahoraISO().slice(0, 10)),
    });
    /**
     * Y entra a Finanzas como ingreso, sin que haya que escribirlo dos
     * veces. Escribirlo a mano en dos sitios es la forma más segura de
     * que un día no cuadren.
     */
    mioGuardar_('Finanzas', {
      fecha: fila.fecha_cobrada,
      flujo: 'ingreso',
      categoria: 'trabajo',
      concepto: String(fila.concepto || 'Cobro'),
      monto: num(fila.monto),
      moneda: fila.moneda,
      trabajo_id: fila.trabajo_id,
      nota: 'Entró solo al marcar el cobro',
    });
    registrarCentral(s, 'Cobros', p.id, 'cobrado', '', String(num(fila.monto)));
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
