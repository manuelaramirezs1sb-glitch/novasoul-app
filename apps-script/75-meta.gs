/**
 * ═══════════════════════════════════════════════════════════
 *  META · LA PAUTA ENTRA SOLA
 * ═══════════════════════════════════════════════════════════
 *
 * Hasta hoy, para que Nova supiera cuánto se gastó en anuncios alguien
 * tenía que entrar a Meta, exportar un Excel y subirlo. Todos los días.
 * Y ese archivo trae fechas ambiguas, monedas mezcladas y una fila por
 * periodo en vez de por día — tres formas de equivocarse que no existen
 * si los datos se piden directamente.
 *
 * DÓNDE VIVE LA LLAVE
 *
 * En las Propiedades del Script, nunca en una hoja. Una llave en una
 * celda la ve cualquiera a quien le compartan el archivo, y la de Meta
 * no caduca: quien la copie la tiene para siempre. En las Propiedades no
 * la ve ni quien abra la hoja ni quien la descargue.
 *
 * Un mismo script atiende a todos los clientes, así que la clave lleva
 * el identificador de la hoja de cada uno: la llave de una tienda no
 * puede quedar al alcance de otra.
 *
 * LO QUE NOVA PUEDE HACER CON ELLA
 *
 * Leer. Nada más. El permiso que se pide es `ads_read` y el activo se
 * asigna como "ver rendimiento", así que ni esta función ni ninguna otra
 * puede crear un anuncio, pausarlo ni mover un presupuesto. Si esa llave
 * se filtrara, lo peor que alguien podría hacer es enterarse del gasto.
 */

const META_API = 'https://graph.facebook.com/v21.0/';

/** La llave de ESTE cliente. Nunca una global. */
function metaClave(s) { return 'META_TOKEN_' + String(s.sheetId || '').slice(0, 44); }

function metaToken(s) {
  return PropertiesService.getScriptProperties().getProperty(metaClave(s)) || '';
}

/**
 * Qué cuenta publicitaria le toca a cada tienda.
 * Vive en Parametros, que es donde ya viven los ajustes de la tienda.
 */
function metaCuenta(ss, tienda) {
  return String(ajustes(ss, tienda).meta_cuenta || '').replace(/^act_/, '').trim();
}

/**
 * El estado de la conexión, SIN devolver la llave.
 *
 * Nunca se manda de vuelta a la pantalla, ni recortada. Una llave que
 * viaja al navegador queda en la memoria del navegador, y de ahí a una
 * captura de pantalla hay un paso. Se dice si existe y cuándo se guardó;
 * para cambiarla se escribe una nueva.
 */
function apiMetaEstado(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña ve la conexión con Meta.' };
  const ss = SpreadsheetApp.openById(s.sheetId);
  const props = PropertiesService.getScriptProperties();
  /**
   * La última prueba que salió bien.
   *
   * Sin esto, lo único que decía si la conexión servía era el mensaje
   * que aparecía justo al probarla — y se perdía al recargar. Quien
   * volvía al día siguiente veía "hay una llave guardada" y no tenía
   * forma de saber si esa llave funcionaba.
   */
  let prueba = null;
  try {
    const crudo = props.getProperty(metaClave(s) + '_PRUEBA');
    if (crudo) prueba = JSON.parse(crudo);
  } catch (e) { prueba = null; }

  return {
    ok: true,
    hayLlave: !!metaToken(s),
    ultimaPrueba: prueba,
    guardadaEn: props.getProperty(metaClave(s) + '_FECHA') || '',
    cuentas: s.tiendas.map(function (t) {
      return { tienda: t, cuenta: metaCuenta(ss, t), moneda: monedaDeTienda(ss, t) };
    }),
  };
}

/** Guardar la llave y/o el número de cuenta de una tienda. */
function apiMetaGuardar(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña conecta Meta.' };
  const props = PropertiesService.getScriptProperties();

  /**
   * Se llama `llave` y no `token` a propósito.
   *
   * `token` es el de la sesión de Nova y viaja en cada petición. Cuando
   * este campo se llamaba igual, la llave de Meta ocupaba su lugar al
   * enviar: el servidor no reconocía la sesión y respondía cerrándola.
   * Guardar la llave expulsaba a quien la estaba guardando.
   */
  if (p.llave !== undefined) {
    const t = String(p.llave).trim();
    if (t && t.length < 50) {
      return { ok: false, error: 'Esa llave se ve incompleta. Las de Meta pasan de ' +
               'doscientos caracteres — puede que se haya cortado al copiarla.' };
    }
    if (t) {
      props.setProperty(metaClave(s), t);
      props.setProperty(metaClave(s) + '_FECHA', ahoraISO());
      // La prueba anterior era de la llave anterior. Dejarla puesta diría
      // "conectada" sobre una llave que nadie ha comprobado todavía.
      props.deleteProperty(metaClave(s) + '_PRUEBA');
    } else {
      props.deleteProperty(metaClave(s));
      props.deleteProperty(metaClave(s) + '_FECHA');
      props.deleteProperty(metaClave(s) + '_PRUEBA');
    }
    registrarMovimiento(s, 'Meta', 'llave', 'token', '(oculto)', t ? '(guardada)' : '(borrada)');
  }

  if (p.cuenta !== undefined && p.tienda) {
    if (s.tiendas.indexOf(p.tienda) === -1) {
      return { ok: false, error: 'No tienes acceso a esa tienda.' };
    }
    // Se escribe por la misma puerta que el resto de ajustes: valida,
    // registra el cambio y no acepta claves fuera del catálogo.
    const cta = String(p.cuenta).replace(/^act_/, '').replace(/\D/g, '');
    const rp = apiParametros(s, { tienda: p.tienda, cambios: { meta_cuenta: cta } });
    if (!rp.ok) return rp;
  }

  return apiMetaEstado(s, p);
}

/**
 * Probar la conexión y decir QUÉ falla.
 *
 * "No se pudo conectar" no sirve: quien lo lee no sabe si volver a Meta,
 * revisar el número de cuenta o esperar. Meta devuelve códigos distintos
 * para la llave vencida, el permiso que falta y la cuenta que no existe,
 * y cada uno tiene una respuesta distinta.
 */
function apiMetaProbar(s, p) {
  if (s.rol !== 'dueno') return { ok: false, error: 'Solo la dueña prueba la conexión.' };

  const token = metaToken(s);
  if (!token) return { ok: false, error: 'Todavía no has guardado la llave.' };

  const ss = SpreadsheetApp.openById(s.sheetId);
  const tienda = String(p.tienda || s.tiendas[0]);
  const cuenta = metaCuenta(ss, tienda);
  if (!cuenta) {
    return { ok: false, error: 'Falta el número de cuenta publicitaria de esta tienda. ' +
             'Está en business.facebook.com, junto al nombre de la cuenta.' };
  }

  const url = META_API + 'act_' + cuenta +
    '?fields=name,currency,account_status,amount_spent' +
    '&access_token=' + encodeURIComponent(token);

  let r;
  try {
    r = UrlFetchApp.fetch(url, { muteHttpExceptions: true });
  } catch (err) {
    return { ok: false, error: 'No se pudo hablar con Meta: ' + err.message };
  }

  let j;
  try { j = JSON.parse(r.getContentText()); }
  catch (err) { return { ok: false, error: 'Meta contestó algo que no entiendo.' }; }

  if (j.error) return Object.assign({ ok: false }, explicarErrorMeta(j.error));

  const monedaTienda = monedaDeTienda(ss, tienda);

  PropertiesService.getScriptProperties().setProperty(
    metaClave(s) + '_PRUEBA',
    JSON.stringify({ cuando: ahoraISO(), cuenta: cuenta,
                     nombre: j.name || '', moneda: j.currency || '', tienda: tienda }));

  return {
    ok: true, tienda: tienda, cuenta: cuenta,
    nombre: j.name || '', moneda: j.currency || '',
    // Si Meta cobra en una moneda y la tienda factura en otra, el gasto
    // no se puede sumar a las ventas sin convertirlo. Mejor decirlo aquí
    // que dejar que aparezca un margen absurdo dentro de tres semanas.
    avisoMoneda: (j.currency && monedaTienda && j.currency !== monedaTienda)
      ? 'Meta te cobra en ' + j.currency + ' y esta tienda factura en ' +
        monedaTienda + '. Nova va a convertir con la tasa del día de cada gasto.'
      : '',
  };
}

/** Traducir el error de Meta a algo que diga qué hacer. */
function explicarErrorMeta(e) {
  const cod = e.code, sub = e.error_subcode;

  if (cod === 190) {
    return { error: 'La llave ya no sirve.\n\n' +
      (sub === 463 ? 'Venció. ' : 'Puede que la hayan revocado, o que lleve 180 días sin usarse. ') +
      'Genera otra en Meta: Usuarios del sistema → Nova → Generar identificador.',
      accion: 'llave' };
  }
  if (cod === 200 || cod === 10) {
    return { error: 'La llave funciona, pero no tiene permiso para leer esta cuenta.\n\n' +
      'Revisa dos cosas en Meta: que la cuenta publicitaria esté asignada al usuario ' +
      'de sistema Nova con "ver rendimiento", y que la llave se haya generado con ' +
      'el permiso ads_read.', accion: 'permiso' };
  }
  if (cod === 100) {
    return { error: 'Meta no encuentra esa cuenta publicitaria.\n\n' +
      'Revisa el número: son solo dígitos, sin el "act_" adelante.', accion: 'cuenta' };
  }
  if (cod === 17 || cod === 4 || cod === 613) {
    return { error: 'Meta está limitando las consultas ahora mismo. ' +
      'Espera unos minutos y vuelve a intentar.', accion: 'esperar' };
  }
  return { error: 'Meta dijo: ' + (e.message || 'error ' + cod), accion: '' };
}


// ─── LO QUE CORRE SOLO ───────────────────────────────────────

/**
 * ═══════════════════════════════════════════════════════════
 *  UN SOLO INTERRUPTOR PARA TODO LO AUTOMÁTICO
 * ═══════════════════════════════════════════════════════════
 *
 * Nova tiene trabajos que deberían correr sin que nadie los pida: bajar
 * las tasas de cambio del día y revisar las alarmas. El código estaba,
 * pero cada uno se prendía con su propia función y había que acordarse de
 * las dos. Nadie se acuerda de dos cosas que se hacen una sola vez.
 *
 * El resultado era el peor posible: parecía automático y no lo era.
 * Alguien conectaba Meta, veía la pauta llegar en pesos contra una tienda
 * en dólares, y se quedaba esperando una conversión que nunca iba a pasar
 * porque la hoja Tasas estaba vacía — sin un error, sin una pista.
 *
 * Ahora es un interruptor. Se puede accionar desde el Apps Script
 * (`prenderAutomatico`) o desde la consola, con un botón — y se puede
 * comprobar por los dos lados, porque un automatismo que no se puede
 * comprobar es un automatismo en el que no se puede confiar.
 *
 * Una advertencia que vale por toda esta sección: los disparadores son
 * del PROYECTO, no de cada cliente. Un solo Apps Script atiende a todos,
 * así que se prenden una vez y sirven para todos. Por eso el botón vive
 * en Nova Central y no en la pantalla de la dueña: no es una decisión de
 * cada cliente, y ponerlo ahí le daría a cualquiera un interruptor que
 * afecta a los demás.
 * ═══════════════════════════════════════════════════════════
 */

/**
 * Los dos trabajos que Nova tiene que correr sola, y sus horas.
 *
 * Están en un solo sitio para que prender, revisar y mostrar hablen de
 * lo mismo. Cuando cada función tenía su propia lista, prender uno y
 * revisar otro era cuestión de tiempo.
 */
const TRABAJOS = [
  {
    fn: 'actualizarTasasDiario',
    nombre: 'Tasas de cambio',
    hora: 5,
    // Va primero a propósito: sin tasas el gasto de pauta en otra moneda
    // no se suma, así que las alarmas de CPA y margen estarían juzgando
    // una operación a la que le falta el gasto.
    porque: 'Sin esto, la pauta que Meta cobra en otra moneda no se puede sumar.',
  },
  {
    fn: 'leerMetaDiario',
    nombre: 'Lectura de Meta',
    hora: 6,
    porque: 'Sin esto, el gasto de anuncios solo entra si alguien sube el Excel.',
  },
  {
    fn: 'revisarAlarmasTodos',
    nombre: 'Revisión de alarmas',
    hora: 7,
    // De último: juzga el día con el gasto ya adentro y ya convertido.
    porque: 'Sin esto, las alarmas solo se calculan cuando alguien abre Nova.',
  },
  {
    fn: 'semaforoLunes',
    nombre: 'Semáforo semanal',
    hora: 8,
    // El único semanal. Va el lunes y a las 8, después de que las tasas,
    // Meta y las alarmas de esa misma mañana ya corrieron: si saliera
    // antes, juzgaría la semana con la pauta del viernes.
    dia: 'MONDAY',
    porque: 'Sin esto, el cierre de la semana solo existe si alguien lo pide.',
  },
];

/** Cómo se dice la frecuencia de un trabajo, en castellano. */
function cuandoCorre_(t) {
  return t.dia === 'MONDAY' ? 'los lunes a las ' + t.hora + ':00'
                            : 'todos los días a las ' + t.hora + ':00';
}

/** Qué disparadores hay puestos ahora mismo, por función. */
function trabajosPuestos_() {
  const puestos = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    puestos[t.getHandlerFunction()] = true;
  });
  return puestos;
}

/**
 * El orden de las horas importa, y por eso un disparador viejo se rehace.
 *
 * Las tasas tienen que estar antes de que entre el gasto de Meta, y las
 * alarmas después, para que juzguen el día con el gasto ya adentro. Si
 * una versión anterior dejó un disparador a otra hora, mantenerlo sería
 * dejar el orden al azar — así que se borra y se vuelve a crear.
 *
 * Google no dice a qué hora quedó un disparador (`atHour` define una
 * franja de una hora, no un minuto exacto), así que la hora se guarda
 * aparte, en las Propiedades. Sin eso no habría forma de saber si el que
 * está puesto es el de ahora o el de antes.
 */
const PROP_HORAS = 'NOVA_HORAS_TRABAJOS';

function horasGuardadas_() {
  try {
    return JSON.parse(PropertiesService.getScriptProperties()
      .getProperty(PROP_HORAS) || '{}');
  } catch (e) { return {}; }
}

/**
 * Prende los dos trabajos. No toca las tasas: eso se hace aparte.
 *
 * Separado a propósito. Instalar un disparador es instantáneo; bajar
 * noventa días de tasas de GOOGLEFINANCE tarda. Si fueran la misma
 * llamada, el botón de la consola se quedaría colgado hasta que
 * terminara lo lento, y quien lo aprieta no sabría si funcionó.
 */
function prenderTrabajos_() {
  const props = PropertiesService.getScriptProperties();
  const horas = horasGuardadas_();
  const disparadores = ScriptApp.getProjectTriggers();

  const hechos = TRABAJOS.map(function (t) {
    const mios = disparadores.filter(function (d) {
      return d.getHandlerFunction() === t.fn;
    });
    const correcta = horas[t.fn] === t.hora;

    if (mios.length === 1 && correcta) {
      return { fn: t.fn, nombre: t.nombre, ya: true, hora: t.hora,
               cuando: cuandoCorre_(t) };
    }
    // Sobrantes o a la hora equivocada: se rehace. Dos disparadores de la
    // misma función leerían Meta dos veces la misma mañana.
    mios.forEach(function (d) { try { ScriptApp.deleteTrigger(d); } catch (e) {} });
    const b = ScriptApp.newTrigger(t.fn).timeBased().atHour(t.hora);
    if (t.dia) b.onWeekDay(ScriptApp.WeekDay[t.dia]).create();
    else b.everyDays(1).create();
    horas[t.fn] = t.hora;
    return { fn: t.fn, nombre: t.nombre, ya: false, hora: t.hora,
             cuando: cuandoCorre_(t), rehecho: mios.length > 0 };
  });

  props.setProperty(PROP_HORAS, JSON.stringify(horas));
  return hechos;
}

/**
 * Las tasas de los últimos noventa días de cada cliente, ahora.
 *
 * El disparador diario solo tapa los huecos de aquí en adelante. Sin
 * esta primera carga, el gasto de pauta de los meses pasados se quedaría
 * sin convertir para siempre — y los cierres que ya se hicieron no
 * tendrían con qué cuadrar.
 *
 * Esta función es también la que corre el disparador de un solo uso que
 * arma la consola, así que no puede recibir parámetros ni depender de
 * una sesión.
 */
function cargaInicialTasas() {
  const hechos = [];
  let cargadas = 0;
  listarClientes().forEach(function (c) {
    if (!c.sheetId) return;
    try { actualizarTasas(c.sheetId, 90); cargadas++; }
    catch (e) { hechos.push('OJO · no pude cargar las tasas de ' + c.empresa + ': ' + e.message); }
  });
  hechos.unshift('Tasas de los últimos 90 días cargadas en ' + cargadas + ' cuenta(s).');

  // Y se borra el disparador de un solo uso que la trajo hasta aquí, si
  // lo hubo. Un disparador "after" que nadie limpia se queda ocupando
  // una de las veinte ranuras que da Google, para siempre.
  ScriptApp.getProjectTriggers().forEach(function (t) {
    if (t.getHandlerFunction() === 'cargaInicialTasas' &&
        t.getEventType() === ScriptApp.EventType.CLOCK) {
      try { ScriptApp.deleteTrigger(t); } catch (e) {}
    }
  });

  const msg = hechos.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Prender lo automático, desde el Apps Script.
 *
 * La consola hace lo mismo con un botón (`centralPrenderAutomatico`).
 * Esta versión existe para cuando la consola todavía no está montada, o
 * cuando hay que arreglar algo sin depender de que el sitio cargue.
 */
function prenderAutomatico() {
  const hechos = prenderTrabajos_().map(function (r) {
    if (r.ya) return '· ' + r.nombre + ': ya estaba corriendo.';
    return '· ' + r.nombre + ': ' + (r.rehecho ? 'reprogramado' : 'prendido') +
           ', ' + r.cuando + '.';
  });
  hechos.push('· ' + cargaInicialTasas());

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + hechos.join('\n') +
    '\n\nPara comprobarlo cuando quieras, corre verAutomatico().';
  Logger.log(msg);
  return msg;
}

// ─── COMPROBAR ───────────────────────────────────────────────

/**
 * El estado de lo automático, en datos.
 *
 * Devuelve objetos y no texto porque lo consume la consola. `verAutomatico()`
 * es esto mismo escrito para leer en el Apps Script — una sola fuente,
 * dos formas de mirarla, para que no puedan contradecirse.
 */
function estadoAutomatico() {
  const puestos = trabajosPuestos_();
  const horas = horasGuardadas_();
  const trabajos = TRABAJOS.map(function (t) {
    // "Puesto a otra hora" cuenta como apagado: el orden entre los tres
    // es lo que hace que el gasto entre convertido y las alarmas lo vean.
    return { nombre: t.nombre, hora: t.hora, porque: t.porque,
             cuando: cuandoCorre_(t),
             prendido: !!puestos[t.fn] && horas[t.fn] === t.hora };
  });

  const clientes = [];
  let error = '';
  try {
    listarClientes().forEach(function (c) {
      if (!c.sheetId) return;
      clientes.push(estadoTasasDe_(c.sheetId, c.empresa));
    });
  } catch (e) {
    error = e.message;
  }

  return {
    trabajos: trabajos,
    todoPrendido: trabajos.every(function (t) { return t.prendido; }),
    clientes: clientes,
    // Un cliente "en falta" es uno que necesita tasas y no las tiene al
    // día. Es el número que decide si la consola avisa o se calla.
    enFalta: clientes.filter(function (c) { return c.necesita && !c.alDia; }).length,
    error: error,
  };
}

/**
 * El estado de las tasas de una cuenta.
 *
 * Distingue tres cosas que se confunden: que la hoja esté vacía, que
 * esté vieja, y que no haga falta. Una tienda que factura en la misma
 * moneda en que le cobran no necesita ninguna tasa, y decirle que le
 * "faltan" sería mandarla a arreglar algo que no está roto.
 */
function estadoTasasDe_(sheetId, nombre) {
  const out = { empresa: nombre, necesita: false, alDia: true, pares: [], error: '' };
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const pares = paresNecesarios(ss);
    if (!pares.length) return out;          // no necesita: alDia se queda en true
    out.necesita = true;

    /**
     * Par por par, no la hoja entera.
     *
     * Mirar solo la fecha más reciente de toda la hoja diría "al día"
     * mientras a un par le falta hasta la primera fila: basta con que
     * otro par se esté actualizando bien para tapar el hueco.
     */
    const ultimaDe = {};
    const sh = ss.getSheetByName('Tasas');
    if (sh && sh.getLastRow() > 1) {
      sh.getDataRange().getValues().slice(1).forEach(function (f) {
        const fecha = aISO(f[0], 'UTC');
        if (!fecha) return;
        const o = String(f[1] || '').toUpperCase();
        const d = String(f[2] || '').toUpperCase();
        [o + '|' + d, d + '|' + o].forEach(function (k) {   // el inverso sirve igual
          if (!ultimaDe[k] || fecha > ultimaDe[k]) ultimaDe[k] = fecha;
        });
      });
    }

    const hoy = ahoraISO().slice(0, 10);
    out.pares = pares.map(function (p) {
      const ultima = ultimaDe[p.origen + '|' + p.destino] || '';
      const dias = ultima
        ? Math.floor((new Date(hoy + 'T00:00:00Z') - new Date(ultima + 'T00:00:00Z')) / 86400000)
        : -1;
      const bien = dias >= 0 && dias <= 3;
      if (!bien) out.alDia = false;
      return {
        par: p.origen + '→' + p.destino,
        ultima: ultima, dias: dias, bien: bien,
      };
    });
  } catch (e) {
    out.error = e.message;
    out.alDia = false;
  }
  return out;
}

/**
 * Lo mismo, escrito para leer en el Apps Script.
 *
 * Existe porque "instalado" y "funcionando" no son lo mismo. Un
 * disparador puede estar puesto y fallar todos los días en silencio —
 * Google lo reintenta, no avisa, y la hoja se queda vieja sin que nadie
 * lo note.
 */
function verAutomatico() {
  const e = estadoAutomatico();

  const lineas = e.trabajos.map(function (t) {
    return (t.prendido ? '✓ ' : '✗ ') + t.nombre +
           (t.prendido ? ' (' + t.cuando + ')' : '  ← APAGADO. ' + t.porque);
  });

  lineas.push('');
  if (e.error) {
    lineas.push('No pude revisar las tasas: ' + e.error);
  } else if (!e.clientes.length) {
    lineas.push('No hay clientes registrados todavía, así que no hay tasas que revisar.');
  }

  e.clientes.forEach(function (c) {
    if (c.error) { lineas.push('· ' + c.empresa + ': no pude revisar — ' + c.error); return; }
    if (!c.necesita) {
      lineas.push('· ' + c.empresa + ': no necesita tasas — factura y le cobran en la misma moneda.');
      return;
    }
    const detalle = c.pares.map(function (p) {
      if (p.dias < 0) return p.par + ' SIN NINGUNA TASA';
      return p.par + (p.bien ? ' ✓ al ' + p.ultima
                             : ' ' + p.dias + ' días atrasado (última: ' + p.ultima + ')');
    });
    lineas.push('· ' + c.empresa + ': ' + detalle.join(' · ') +
                (c.alDia ? '' : '\n    ← corre prenderAutomatico()'));
  });

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + lineas.join('\n');
  Logger.log(msg);
  return msg;
}
