/**
 * ═══════════════════════════════════════════════════════════════
 *  MI PLATA · ordenada por lo que hace cada peso, no por categoría
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE ELLA DIJO, ENTERO ─────────────────────────────────┐
 * │                                                            │
 * │ «la parte de mi plata no la entiendo (…) no entiendo cómo  │
 * │  razona esta pantalla, porque los gastos fijos están       │
 * │  abajo, debería haber una distinción si es único pago,     │
 * │  pago mensual, pago por cuotas; las cuotas y deudas deben  │
 * │  estar en otra parte, separada; los ingresos y los gastos  │
 * │  fijos también deben estar separados, que tenga una        │
 * │  coherencia ahí; como lo veo no entiendo nada de lo que    │
 * │  me quiere transmitir la pantalla más que mi plata         │
 * │  entrante y ya.»                                           │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ QUÉ ESTABA MAL, DE VERDAD ────────────────────────────────┐
 * │                                                            │
 * │ No era el orden de los bloques. Era que la pantalla estaba │
 * │ ordenada por CATEGORÍA —arriendo, mercado, servicios,      │
 * │ internet, crédito, deudas— y la categoría no responde      │
 * │ ninguna pregunta que ella se haga con la plata.            │
 * │                                                            │
 * │ Las preguntas reales son cuatro, y son de otro eje:        │
 * │                                                            │
 * │   ¿cuánto entra?                                           │
 * │   ¿cuánto sale TODOS los meses, pase lo que pase?          │
 * │   ¿cuánto de eso es deuda, y cuándo se acaba?              │
 * │   ¿qué queda?                                              │
 * │                                                            │
 * │ «Deudas · COP 690.000» en la misma lista que «Internet ·   │
 * │ COP 93.000» hace imposible la tercera: el internet es para │
 * │ siempre y la deuda tiene fecha de salida, y esa fecha es   │
 * │ EL dato — es la diferencia entre deber y estar pagando.    │
 * │                                                            │
 * │ Así que esta pantalla no reordena: reagrupa por lo que     │
 * │ cada peso HACE. Las categorías siguen vivas, pero adentro  │
 * │ de su bloque, que es donde sí sirven.                      │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ NUNCA HAY UN TOTAL ÚNICO ─────────────────────────┐
 * │                                                            │
 * │ Cobra en dólares por Upwork, vive en pesos y se va a       │
 * │ España. Un total que sume USD, COP y EUR cambia solo de un │
 * │ día para otro sin que ella gaste ni cobre nada. Todo va    │
 * │ moneda por moneda, y eso no es una limitación: es lo       │
 * │ único honesto que se puede mostrar.                        │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Entra o sale. Una línea fija puede ser cualquiera de las dos. */
const PLATA_FLUJOS = {
  ingreso: { nombre: 'Entra', signo: 1 },
  gasto:   { nombre: 'Sale',  signo: -1 },
};

/**
 * Cómo se paga. Es la distinción que ella pidió por nombre, y la que
 * decide en qué bloque cae la línea.
 */
const PLATA_TIPOS = {
  mensual: { nombre: 'Todos los meses', orden: 1,
             ayuda: 'Se repite cada mes y no se acaba: arriendo, internet, servicios.' },
  cuotas:  { nombre: 'Por cuotas', orden: 2,
             ayuda: 'Tiene un número de cuotas y una fecha en que termina. Aquí van las deudas.' },
  unico:   { nombre: 'Pago único', orden: 3,
             ayuda: 'Pasa una vez, este mes. No se repite el siguiente.' },
};

/**
 * Los bloques de la pantalla, en el orden en que se leen.
 *
 * El orden no es estético: es el orden en que se contesta la pregunta
 * «¿me alcanza?». Primero lo que entra, después lo que sale sí o sí,
 * después lo que se puede negociar, y al final lo que queda.
 */
const PLATA_BLOQUES = [
  { id: 'entra',   nombre: 'Lo que entra',
    que: 'Turnos, propinas y todo lo que te pagan. Lo de este mes y lo que falta por entrar.' },
  { id: 'fijo',    nombre: 'Lo que sale todos los meses',
    que: 'Lo que se repite pase lo que pase. Esto es tu piso: por debajo de esto no se puede bajar el mes.' },
  { id: 'cuotas',  nombre: 'Cuotas y deudas',
    que: 'Lo que estás pagando a plazos. Esto SÍ se acaba, y aquí ves cuándo.' },
  { id: 'unico',   nombre: 'Pagos únicos de este mes',
    que: 'Lo que pasa una vez y no vuelve el mes que viene.' },
  { id: 'hormiga', nombre: 'Gastos hormiga',
    que: 'Lo pequeño que no estaba en el plan. Sumado, casi siempre pesa más de lo que parece.' },
  { id: 'ahorro',  nombre: 'Ahorro',
    que: 'No es un gasto. Va aparte para que no ensucie la cuenta de lo que sale.' },
];

/**
 * Una línea fija, leída con todo lo que hace falta para entenderla.
 *
 * Lo importante de aquí es `cuotas`: sin el mes de la primera cuota no
 * se puede decir cuándo termina, y decir «van 3 de 6» sin decir «acaba
 * en febrero» es justamente el dato a medias que no sirve para
 * decidir nada.
 */
function plataLinea_(f, mesISO, pagos) {
  const tipo = PLATA_TIPOS[norm(f.tipo_pago)] ? norm(f.tipo_pago) : '';
  const flujo = norm(f.flujo) === 'ingreso' ? 'ingreso' : 'gasto';
  const categoria = norm(f.categoria) || 'varios';

  const o = {
    id: String(f.id || ''),
    categoria: categoria,
    categoriaNombre: (SOUL_CATEGORIAS.filter(function (c) { return c.id === categoria; })[0]
                      || { nombre: 'Gastos varios' }).nombre,
    concepto: String(f.concepto || ''),
    monto: num(f.monto),
    moneda: String(f.moneda || 'COP').toUpperCase(),
    dia: num(f.dia_del_mes) || null,
    activo: norm(f.activo) !== 'no',
    nota: String(f.nota || ''),
    flujo: flujo,
    acreedor: String(f.acreedor || ''),
    /**
     * Sin tipo de pago declarado, se deduce en vez de dejar la línea
     * fuera de todos los bloques. Las que ella ya tenía cargadas no
     * traen la columna —es nueva— y desaparecer de la pantalla sería
     * el peor comportamiento posible para un dato viejo.
     */
    tipo: tipo || (categoria === 'ahorro' ? 'mensual'
                 : (categoria === 'deudas' || categoria === 'credito') ? 'cuotas'
                 : 'mensual'),
    tipoDeducido: !tipo,
  };

  if (o.tipo === 'cuotas') {
    o.deudaTotal = num(f.deuda_total) || null;
    const total = num(f.cuotas_total);
    const desde = String(f.cuota_desde || '').slice(0, 7);
    let pagadas = num(f.cuotas_pagadas);

    /**
     * Si hay mes de arranque, las pagadas se CUENTAN en vez de
     * creerle a la columna. Una columna que hay que actualizar a mano
     * cada mes es una columna que va a estar desactualizada, y una
     * deuda que dice «van 2 de 6» un año después es peor que no decir
     * nada.
     */
    let contadas = null;
    if (/^\d{4}-\d{2}$/.test(desde)) {
      const meses = (Number(mesISO.slice(0, 4)) - Number(desde.slice(0, 4))) * 12 +
                    (Number(mesISO.slice(5, 7)) - Number(desde.slice(5, 7)));
      contadas = Math.max(0, Math.min(total || 9999, meses));
      pagadas = contadas;
    }

    /**
     * ── LOS PAGOS DE VERDAD MANDAN SOBRE LA CUENTA ──
     *
     * Ella: «dar la oportunidad de cambiar el aporte cada que vaya a
     * subir un pago si pagué más o menos».
     *
     * Así que si hay abonos registrados, «cuánto llevas pagado» es la
     * SUMA DE ESOS ABONOS, no la cuota multiplicada por los meses. Un
     * mes que abonó el doble adelanta de verdad, y uno que abonó la
     * mitad no cuenta como uno completo.
     *
     * Si no hay ninguno registrado todavía, se cae al cálculo de antes
     * —meses transcurridos × cuota— y se DICE que es un estimado. Esa
     * distinción importa: un número calculado y uno contado se ven
     * iguales en pantalla, y solo uno de los dos aguanta una
     * discusión con el banco.
     */
    const abonos = (pagos || []).filter(function (x) { return x.fijoId === o.id; });
    const pagado = abonos.reduce(function (a, x) { return a + x.monto; }, 0);
    const porAbonos = abonos.length > 0;

    if (porAbonos) pagadas = abonos.length;
    const faltan = total ? Math.max(0, total - pagadas) : null;

    /**
     * Lo que falta por pagar. Con la deuda total es una resta exacta;
     * sin ella hay que estimar con la cuota, y se avisa.
     */
    let faltaPagar = null, exacto = false;
    if (o.deudaTotal) { faltaPagar = Math.max(0, o.deudaTotal - pagado); exacto = true; }
    else if (faltan !== null) { faltaPagar = faltan * o.monto; }

    o.cuotas = {
      total: total || null,
      pagadas: pagadas || 0,
      faltan: faltan,
      contadasSolas: contadas !== null && !porAbonos,
      porAbonos: porAbonos,
      abonos: abonos.length,
      pagado: pagado,
      deudaTotal: o.deudaTotal,
      desde: desde || '',
      // El dato que convierte una deuda en algo que se acaba.
      termina: (total && /^\d{4}-\d{2}$/.test(desde))
        ? plataMasMeses_(desde, total - 1) : '',
      faltaPagar: faltaPagar,
      exacto: exacto,
      /**
       * Y el mes en que se acaba de verdad, al ritmo al que va pagando.
       * Puede ser antes de lo planeado si abonó de más — que es
       * exactamente el premio por haberlo hecho, y no verlo desanima.
       */
      terminaAlRitmo: (function () {
        if (!o.deudaTotal || !porAbonos || !o.monto) return '';
        const queda = Math.max(0, o.deudaTotal - pagado);
        const meses = Math.ceil(queda / (pagado / abonos.length));
        return plataMasMeses_(mesISO, meses);
      })(),
      porque: total ? '' :
        'Sin saber cuántas cuotas son no puedo decirte cuándo se acaba. ' +
        'Es el dato que convierte una deuda en algo con fecha de salida.',
      porqueEstimado: (!o.deudaTotal && total)
        ? 'Esto es un estimado: multiplico la cuota por lo que falta. ' +
          'Si pones la deuda total, la resta es exacta.'
        : '',
    };
  }

  return o;
}

/** Sumar meses a un AAAA-MM sin pasar por Date, que aquí sobra. */
function plataMasMeses_(mesISO, n) {
  let a = Number(mesISO.slice(0, 4));
  let m = Number(mesISO.slice(5, 7)) + n;
  a += Math.floor((m - 1) / 12);
  m = ((m - 1) % 12 + 12) % 12 + 1;
  return a + '-' + (m < 10 ? '0' + m : m);
}

/** Suma por moneda. Nunca entre monedas. */
function plataSumar_(lineas) {
  const o = {};
  lineas.forEach(function (l) { o[l.moneda] = (o[l.moneda] || 0) + l.monto; });
  return o;
}

/**
 * La pantalla entera, agrupada por lo que hace cada peso.
 *
 * Devuelve bloques, no categorías. Cada bloque trae sus líneas, su
 * total por moneda y su propia frase de qué es — porque «Cuotas y
 * deudas» no se explica solo, y la pantalla anterior tampoco lo
 * intentaba.
 */
function plataOrdenada_(uid, mes, hoyISO) {
  /**
   * Los abonos se leen UNA vez y se pasan a cada línea. Leerlos dentro
   * de `plataLinea_` sería volver a abrir la hoja por cada gasto fijo.
   */
  const pagos = soulLeerSuave_('Pagos', uid, []).map(function (x) {
    return { id: String(x.id || ''), fijoId: String(x.fijo_id || ''),
             fecha: aISO(x.fecha, 'UTC') || '', monto: num(x.monto),
             moneda: String(x.moneda || 'COP').toUpperCase(),
             nota: String(x.nota || '') };
  }).filter(function (x) { return x.fijoId; })
    .sort(function (a, b) { return a.fecha < b.fecha ? 1 : -1; });

  const fijos = soulLeerSuave_('Fijos', uid, []).map(function (f) {
    return plataLinea_(f, mes, pagos);
  }).filter(function (l) { return l.activo; });

  const del = function (tipo, flujo) {
    return fijos.filter(function (l) {
      if (l.categoria === 'ahorro') return false;
      return l.tipo === tipo && l.flujo === flujo;
    });
  };

  const entraFijo = fijos.filter(function (l) { return l.flujo === 'ingreso'; });
  const saleMensual = del('mensual', 'gasto');
  const saleCuotas = del('cuotas', 'gasto');
  const saleUnico = del('unico', 'gasto');
  const ahorro = fijos.filter(function (l) { return l.categoria === 'ahorro'; });

  /**
   * Lo que YA entró y lo que FALTA por entrar, de los turnos.
   *
   * Se reutiliza `soulMes_`, que ya lo calculaba: no se vuelve a
   * escribir la cuenta de los turnos aquí. Dos cuentas del mismo
   * número en dos sitios es cómo se consiguen dos respuestas distintas
   * a la misma pregunta.
   */
  const mesResumen = soulMes_(uid, mes, hoyISO);

  const dia = hoyISO.slice(0, 7) === mes ? Number(hoyISO.slice(8, 10)) : 32;
  const yaPaso = function (l) { return l.dia && l.dia < dia; };

  const bloques = {
    entra: {
      lineas: entraFijo,
      total: plataSumar_(entraFijo),
      // Lo de los turnos va aparte porque no es un plan: ya pasó.
      turnos: mesResumen.entro || {},
      porVenir: {},
      turnosPorVenir: mesResumen.turnosPorVenir || 0,
    },
    fijo: {
      lineas: saleMensual,
      total: plataSumar_(saleMensual),
      pagado: plataSumar_(saleMensual.filter(yaPaso)),
      pendiente: plataSumar_(saleMensual.filter(function (l) { return !yaPaso(l); })),
    },
    cuotas: {
      lineas: saleCuotas,
      total: plataSumar_(saleCuotas),
      // Lo que falta por pagar EN TOTAL, no este mes. Es el número que
      // dice de qué tamaño es la deuda de verdad.
      faltaTodo: (function () {
        const o = {};
        saleCuotas.forEach(function (l) {
          if (l.cuotas && l.cuotas.faltaPagar !== null) {
            o[l.moneda] = (o[l.moneda] || 0) + l.cuotas.faltaPagar;
          }
        });
        return o;
      })(),
      /** La última en acabarse: el mes en que deja de doler. */
      ultima: (function () {
        const fs = saleCuotas.map(function (l) { return l.cuotas && l.cuotas.termina; })
                             .filter(function (x) { return x; }).sort();
        return fs.length ? fs[fs.length - 1] : '';
      })(),
      sinFecha: saleCuotas.filter(function (l) {
        return !l.cuotas || !l.cuotas.termina; }).length,
    },
    unico: { lineas: saleUnico, total: plataSumar_(saleUnico) },
    ahorro: { lineas: ahorro, total: plataSumar_(ahorro) },
  };

  // Lo que sale, sumado de verdad: fijo + cuotas + único. Sin ahorro.
  const sale = {};
  [saleMensual, saleCuotas, saleUnico].forEach(function (grupo) {
    grupo.forEach(function (l) { sale[l.moneda] = (sale[l.moneda] || 0) + l.monto; });
  });

  /**
   * El resultado, moneda por moneda. `queda` no es una proyección: es
   * lo que entró menos lo que está comprometido. Y se dice cuál de las
   * dos cosas es, porque confundirlas es cómo se gasta plata que ya
   * tenía dueño.
   */
  const monedas = {};
  [bloques.entra.total, bloques.entra.turnos, sale, bloques.ahorro.total]
    .forEach(function (o) { Object.keys(o || {}).forEach(function (k) { monedas[k] = 1; }); });

  const resultado = Object.keys(monedas).sort(function (a, b) {
    return a === 'COP' ? -1 : b === 'COP' ? 1 : (a < b ? -1 : 1);
  }).map(function (k) {
    const entra = (bloques.entra.total[k] || 0) + (bloques.entra.turnos[k] || 0);
    const salen = sale[k] || 0;
    return {
      moneda: k,
      entra: entra,
      sale: salen,
      ahorro: bloques.ahorro.total[k] || 0,
      queda: entra - salen,
      // Cuánto de lo que entra ya tiene dueño antes de llegar.
      comprometidoPct: entra > 0 ? Math.round(salen / entra * 100) : null,
    };
  });

  return {
    mes: mes,
    // Los abonos viajan para que la pantalla pueda listarlos y quitarlos
    // sin pedirlos otra vez.
    pagos: pagos,
    bloquesOrden: PLATA_BLOQUES,
    tipos: PLATA_TIPOS,
    bloques: bloques,
    sale: sale,
    resultado: resultado,
    /**
     * Cuántas líneas están en un bloque porque Nova lo dedujo y no
     * porque ella lo dijo. Se muestra para que pueda corregirlas: una
     * deducción en silencio es una mentira con buena intención.
     */
    porClasificar: fijos.filter(function (l) { return l.tipoDeducido; })
      .map(function (l) {
        return { id: l.id, concepto: l.concepto, categoria: l.categoriaNombre,
                 supuesto: PLATA_TIPOS[l.tipo].nombre };
      }),
  };
}

/**
 * `nc_soul_plata` · la pantalla, ya reagrupada.
 *
 * Se mantiene todo lo que la pantalla vieja devolvía —las categorías,
 * lo hormiga, los turnos pendientes— porque hay partes de la interfaz
 * que siguen leyéndolo, y se agrega `orden`, que es lo nuevo. Cambiar
 * la forma de la respuesta y la pantalla a la vez deja sin saber cuál
 * de las dos cosas rompió.
 */
function soulPlataOrdenada(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const hoy = ahoraISO().slice(0, 10);
  const mes = String(p.mes || hoy.slice(0, 7));

  const base = soulPlata(s, { mes: mes });
  if (!base.ok) return base;

  base.orden = plataOrdenada_(uid, mes, hoy);
  base.flujos = PLATA_FLUJOS;
  return base;
}

/**
 * ═══════════════════════════════════════════════════════════════
 *  LOS ABONOS · lo que de verdad pagó este mes
 * ═══════════════════════════════════════════════════════════════
 *
 * «dar la oportunidad de cambiar el aporte cada que vaya a subir un
 *  pago si pagué más o menos».
 *
 * La cuota del plan dice lo que debería pagar; el abono dice lo que
 * pagó. Cuando son distintos, el que manda es el segundo — y sin esta
 * hoja no había forma de decirlo: el plan se hacía pasar por hecho.
 */
function soulPagoGuardar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  const uid = soulUsuario_(s);
  const d = p.datos || {};

  const fijoId = String(d.fijoId || d.fijo_id || '').trim();
  if (!fijoId) return { ok: false, error: 'No sé a qué deuda es este abono.' };

  /**
   * El abono tiene que ir contra una deuda SUYA y que esté a plazos.
   * Sin esto se podrían acumular abonos contra el arriendo, y la
   * pantalla mostraría una deuda que se acaba donde no hay ninguna.
   */
  const suyo = soulLeerSuave_('Fijos', uid, []).filter(function (f) {
    return String(f.id || '').trim() === fijoId;
  })[0];
  if (!suyo) return { ok: false, error: 'No encuentro esa deuda entre tus gastos fijos.' };

  const monto = num(d.monto);
  if (!(monto > 0)) {
    return { ok: false, error: 'Un abono de cero no es un abono. ' +
                               'Si este mes no pagaste, simplemente no lo registres.' };
  }
  const fecha = String(d.fecha || '').trim() || ahoraISO().slice(0, 10);
  if (!/^\d{4}-\d{2}-\d{2}$/.test(fecha)) {
    return { ok: false, error: 'La fecha del abono no se entiende.' };
  }

  try {
    soulGuardar_('Pagos', {
      id: String(d.id || ''),
      fijo_id: fijoId,
      fecha: fecha,
      monto: monto,
      // La moneda la manda la deuda, no la pantalla: un abono en otra
      // moneda que la deuda no se puede restar de nada.
      moneda: String(suyo.moneda || 'COP').toUpperCase(),
      nota: String(d.nota || '').trim(),
    }, uid);
    return { ok: true };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}

function soulPagoBorrar(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };
  try {
    const fue = soulBorrar_('Pagos', p.id, soulUsuario_(s));
    return { ok: fue, error: fue ? '' : 'No encuentro ese abono.' };
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
