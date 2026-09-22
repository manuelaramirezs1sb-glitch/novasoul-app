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
