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
 * Ahora es una función: `prenderAutomatico()`. Y `verAutomatico()` dice
 * qué está corriendo, porque un automatismo que no se puede comprobar es
 * un automatismo en el que no se puede confiar.
 */
function prenderAutomatico() {
  const hechos = [];

  // ── Tasas de cambio ──
  //
  // Va primero a propósito. Sin tasas, el gasto de pauta en otra moneda
  // no se suma —se cuenta aparte y se dice—, así que las alarmas de CPA
  // y margen estarían juzgando una operación a la que le falta el gasto.
  const tasas = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'actualizarTasasDiario';
  });
  if (tasas.length) {
    hechos.push('· Tasas de cambio: ya estaba corriendo.');
  } else {
    ScriptApp.newTrigger('actualizarTasasDiario').timeBased().atHour(6).everyDays(1).create();
    hechos.push('· Tasas de cambio: prendido, todos los días a las 6 a.m.');
  }

  // ── Alarmas ──
  const alarmas = ScriptApp.getProjectTriggers().filter(function (t) {
    return t.getHandlerFunction() === 'revisarAlarmasTodos';
  });
  const hora = Number(ALARMAS_DEFAULT.alarmas_hora) || 7;
  if (alarmas.length) {
    hechos.push('· Revisión de alarmas: ya estaba corriendo.');
  } else {
    ScriptApp.newTrigger('revisarAlarmasTodos').timeBased().atHour(hora).everyDays(1).create();
    hechos.push('· Revisión de alarmas: prendido, todos los días a las ' + hora + ':00.');
  }

  /**
   * Y las tasas de los últimos noventa días, ahora mismo.
   *
   * El disparador solo tapa los huecos de aquí en adelante. Sin esta
   * primera carga, el gasto de pauta de los meses pasados se quedaría
   * sin convertir para siempre — y los cierres que ya se hicieron no
   * tendrían con qué cuadrar.
   */
  let cargadas = 0;
  try {
    const clientes = listarClientes();
    clientes.forEach(function (c) {
      if (!c.sheetId) return;
      try { actualizarTasas(c.sheetId, 90); cargadas++; }
      catch (e) { hechos.push('· OJO · no pude cargar las tasas de ' +
                              c.empresa + ': ' + e.message); }
    });
    hechos.push('· Tasas de los últimos 90 días cargadas en ' + cargadas + ' cuenta(s).');
  } catch (e) {
    hechos.push('· OJO · no pude cargar las tasas: ' + e.message);
  }

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + hechos.join('\n') +
    '\n\nPara comprobarlo cuando quieras, corre verAutomatico().';
  Logger.log(msg);
  return msg;
}

/**
 * Qué está corriendo solo, y cuándo corrió por última vez.
 *
 * Existe porque "instalado" y "funcionando" no son lo mismo. Un
 * disparador puede estar puesto y fallar todos los días en silencio —
 * Google lo reintenta, no avisa, y la hoja se queda vieja sin que nadie
 * lo note.
 */
function verAutomatico() {
  const esperados = {
    actualizarTasasDiario: 'Tasas de cambio',
    revisarAlarmasTodos:   'Revisión de alarmas',
  };
  const puestos = {};
  ScriptApp.getProjectTriggers().forEach(function (t) {
    puestos[t.getHandlerFunction()] = true;
  });

  const lineas = Object.keys(esperados).map(function (fn) {
    return (puestos[fn] ? '✓ ' : '✗ ' ) + esperados[fn] +
           (puestos[fn] ? '' : '  ← APAGADO, corre prenderAutomatico()');
  });

  /**
   * Y si las tasas están al día de verdad, no solo "instaladas".
   *
   * Se revisa cuenta por cuenta, no una sola hoja: las tasas viven en la
   * hoja de cada cliente, que es donde actualizarTasas las escribe.
   * Mirar una sola diría "al día" mientras la de otro cliente lleva
   * meses vacía — la respuesta correcta a la pregunta equivocada.
   */
  lineas.push('');
  try {
    const clientes = listarClientes();
    if (!clientes.length) {
      lineas.push('No hay clientes registrados todavía, así que no hay tasas que revisar.');
    }
    clientes.forEach(function (c) {
      if (!c.sheetId) return;
      lineas.push(revisarTasasDe_(c.sheetId, c.empresa));
    });
  } catch (e) {
    lineas.push('No pude revisar las tasas: ' + e.message);
  }

  const msg = 'LO AUTOMÁTICO DE NOVA\n\n' + lineas.join('\n');
  Logger.log(msg);
  return msg;
}

/**
 * Una línea sobre las tasas de una cuenta.
 *
 * Distingue tres cosas que se confunden: que la hoja esté vacía, que
 * esté vieja, y que no haga falta. Una tienda que factura en la misma
 * moneda en que le cobran no necesita ninguna tasa, y decirle que le
 * "faltan" sería mandarla a arreglar algo que no está roto.
 */
function revisarTasasDe_(sheetId, nombre) {
  try {
    const ss = SpreadsheetApp.openById(sheetId);
    const pares = paresNecesarios(ss);
    if (!pares.length) {
      return '· ' + nombre + ': no necesita tasas — factura y le cobran en la misma moneda.';
    }
    const comoTexto = pares.map(function (p) {
      return p.origen + '→' + p.destino;
    }).join(', ');

    const sh = ss.getSheetByName('Tasas');
    if (!sh || sh.getLastRow() < 2) {
      return '· ' + nombre + ': OJO · la hoja Tasas está VACÍA y esta cuenta sí ' +
             'las necesita (' + comoTexto + '). El gasto de pauta no se va ' +
             'a sumar hasta que las tenga — corre prenderAutomatico().';
    }

    /**
     * Par por par, no la hoja entera.
     *
     * Mirar solo la fecha más reciente de toda la hoja diría "al día"
     * mientras a un par le falta hasta la primera fila: basta con que
     * otro par se esté actualizando bien para tapar el hueco.
     */
    const ultimaDe = {};
    sh.getDataRange().getValues().slice(1).forEach(function (f) {
      const fecha = aISO(f[0], 'UTC');
      if (!fecha) return;
      const o = String(f[1] || '').toUpperCase();
      const d = String(f[2] || '').toUpperCase();
      [o + '|' + d, d + '|' + o].forEach(function (k) {   // el inverso sirve igual
        if (!ultimaDe[k] || fecha > ultimaDe[k]) ultimaDe[k] = fecha;
      });
    });

    const hoy = ahoraISO().slice(0, 10);
    const detalle = pares.map(function (p) {
      const ultima = ultimaDe[p.origen + '|' + p.destino];
      if (!ultima) return p.origen + '→' + p.destino + ' SIN NINGUNA TASA';
      const dias = Math.floor(
        (new Date(hoy + 'T00:00:00Z') - new Date(ultima + 'T00:00:00Z')) / 86400000);
      return p.origen + '→' + p.destino +
             (dias > 3 ? ' ' + dias + ' días atrasado (última: ' + ultima + ')'
                       : ' ✓ al ' + ultima);
    });

    const malo = detalle.some(function (t) {
      return t.indexOf('✓') === -1;
    });
    return '· ' + nombre + ': ' + detalle.join(' · ') +
           (malo ? '\n    ← corre prenderAutomatico() en esta cuenta' : '');
  } catch (e) {
    return '· ' + nombre + ': no pude revisar las tasas — ' + e.message;
  }
}
