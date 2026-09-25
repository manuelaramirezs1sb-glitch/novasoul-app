/**
 * ═══════════════════════════════════════════════════════════════
 *  LOS ACCESOS DE LA TIENDA
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «los links de acceso, colocarlos en los lugares donde van   │
 * │  automáticamente».                                          │
 * │                                                            │
 * │ Y antes, explicando cuáles son:                            │
 * │                                                            │
 * │ «ahí está el link de la plataforma, la contraseña y correo, │
 * │  está el link de chat center o wpp: el canal que use el     │
 * │  cliente para comunicarse con los compradores de su         │
 * │  tienda».                                                   │
 * │                                                            │
 * │ Son las dos puertas que una gestora abre cada mañana, y     │
 * │ hasta hoy vivían en una hoja de Excel que hay que buscar.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ LA CONTRASEÑA SE GUARDA ──────────────────────────┐
 * │                                                            │
 * │ Levanté la mano por esto y ella lo resolvió con hechos:    │
 * │                                                            │
 * │ «el link que viste es el permitido para que la gestora      │
 * │  entre a hacer su trabajo y la contraseña pues debe de      │
 * │  estar para que la gestora entre sin problema. Ahora, la    │
 * │  persona que entre a la plataforma deberá colocar un        │
 * │  código que le llega al correo que aparece en pantalla: si  │
 * │  no tiene acceso a ese correo no entra aunque quiera. Y el  │
 * │  link que tiene ahí es uno que se crea dentro de Dropi, no  │
 * │  tiene permitido sacar dinero ni nada, solo confirmar y     │
 * │  gestionar novedades y pedidos».                           │
 * │                                                            │
 * │ Entonces sí se guarda, y con tres condiciones que este     │
 * │ archivo cumple y no depende de que nadie se acuerde:       │
 * │                                                            │
 * │  1. NO PASA POR LA BITÁCORA. `apiParametros` registra cada │
 * │     cambio con el valor anterior y el nuevo, así que meter │
 * │     la contraseña ahí la dejaría en texto plano en         │
 * │     Movimientos para siempre. Por eso los accesos tienen   │
 * │     su propia hoja y su propio guardado, y lo que se       │
 * │     registra es «cambió la contraseña», nunca cuál.        │
 * │                                                            │
 * │  2. SOLO LA VE QUIEN TIENE ESA TIENDA. Es el mismo filtro  │
 * │     que ya decide los pedidos; no hay un permiso nuevo que │
 * │     alguien pueda olvidar de marcar.                       │
 * │                                                            │
 * │  3. NO VIAJA DONDE NO HACE FALTA. `apiResumen` y las       │
 * │     alarmas mandan `ajustes` a la pantalla en cada carga.  │
 * │     Si los accesos fueran un ajuste más, la contraseña     │
 * │     iría en cada respuesta del día. Se piden aparte y      │
 * │     solo cuando se abre la pantalla de accesos.            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ DOS PUERTAS DISTINTAS, NO UNA ────────────────────────────┐
 * │                                                            │
 * │ `url/usuario/clave`  la PLATAFORMA donde se gestiona:      │
 * │                      Dropi, Shopify, Mastershop.          │
 * │                                                            │
 * │ `canal_url`          el canal donde la tienda habla con    │
 * │                      SUS COMPRADORES: Chat Center, WhatsApp│
 * │                                                            │
 * │ Estuvieron juntos un tiempo: el canal vivía en la pestaña  │
 * │ «Nova Chat» como si fuera el chat del equipo. No lo es —   │
 * │ el chat del equipo ahora está dentro de Nova (96-chat.gs), │
 * │ y este canal es una herramienta de trabajo más, al lado de │
 * │ la plataforma. Ella lo dijo exactamente así: «no es un     │
 * │ canal que ellos peguen».                                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Lo que Nova guarda de cada puerta. Uno por tienda. */
const ACCESOS_CAMPOS = ['plataforma', 'url', 'usuario', 'clave', 'correo_codigo',
                        'canal_nombre', 'canal_url', 'nota'];

/**
 * Los campos que son un enlace, y por tanto se validan y se completan.
 *
 * Un `javascript:` en cualquiera de los dos correría código en la sesión
 * de quien le diera clic, con su token al lado. Se valida en el servidor
 * porque validar solo en la pantalla deja la puerta abierta a quien llame
 * la API de frente.
 */
const ACCESOS_ENLACES = ['url', 'canal_url'];

/** Los accesos de una tienda, tal como están en la hoja. */
function accesosDe_(ss, tienda) {
  const out = { tienda: tienda };
  ACCESOS_CAMPOS.forEach(function (k) { out[k] = ''; });
  out.actualizado_en = ''; out.actualizado_por = '';

  const sh = ss.getSheetByName('Accesos');
  if (!sh || sh.getLastRow() < 2) return out;
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cT = e.indexOf('tienda');
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][cT] || '').trim() !== tienda) continue;
    e.forEach(function (k, j) {
      if (k && k !== 'tienda') out[k] = String(d[i][j] == null ? '' : d[i][j]).trim();
    });
    break;
  }
  return out;
}

/**
 * Los accesos de una tienda, para la pantalla.
 *
 * `puedeEditar` decide quién ve los campos y quién solo los botones. Lo
 * cambia la dueña: la contraseña de la plataforma es suya y una gestora
 * que la reescriba deja a todo el equipo afuera sin querer.
 */
function apiAccesos(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = libro_(s.sheetId);
  const a = accesosDe_(ss, tienda);

  /**
   * El aviso del código por correo va SIEMPRE, no solo si el correo está
   * puesto. Quien intenta entrar y se topa con la pantalla del código sin
   * saber que existe cree que la contraseña está mal, la cambia, y deja a
   * las demás afuera. Es el error más caro que puede cometer aquí.
   */
  return {
    ok: true, tienda: tienda,
    nombreTienda: nombreTienda(ss, tienda),
    accesos: a,
    puedeEditar: s.rol === 'dueno',
    hay: !!(a.url || a.canal_url),
  };
}

/**
 * Guardar los accesos. Solo la dueña.
 *
 * La bitácora dice QUÉ cambió y no CUÁL era: escribir la contraseña en
 * Movimientos la dejaría ahí en texto plano para siempre, y la bitácora
 * la puede leer cualquiera que abra la hoja.
 */
function apiAccesosGuardar(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  if (s.rol !== 'dueno') {
    return { ok: false, error: 'Solo la dueña cambia los accesos de la tienda.' };
  }
  const cambios = p.cambios || {};

  // Los dos enlaces se validan ANTES de tocar la hoja. Guardar la mitad
  // deja un botón con etiqueta y sin destino, que es peor que no tenerlo.
  for (let i = 0; i < ACCESOS_ENLACES.length; i++) {
    const k = ACCESOS_ENLACES[i];
    if (cambios[k] === undefined) continue;
    const err = validarCanal(cambios[k]);
    if (err) return { ok: false, error: err };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro guardado en curso.' };
  try {
    const ss = libro_(s.sheetId);
    const sh = ss.getSheetByName('Accesos');
    if (!sh) return { ok: false, error: 'Falta la hoja Accesos. Corre bootstrapTodo() una vez.' };

    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const antes = accesosDe_(ss, tienda);

    const nuevo = {};
    enc.forEach(function (k) { nuevo[k] = antes[k] !== undefined ? antes[k] : ''; });
    nuevo.tienda = tienda;
    ACCESOS_CAMPOS.forEach(function (k) {
      if (cambios[k] === undefined) return;
      nuevo[k] = ACCESOS_ENLACES.indexOf(k) !== -1
        ? normalizarEnlace(cambios[k])
        : String(cambios[k] == null ? '' : cambios[k]).trim();
    });
    nuevo.actualizado_en = ahoraISO();
    nuevo.actualizado_por = s.email;

    // Buscar la fila de esta tienda; si no existe, se agrega.
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];
    const cT = enc.indexOf('tienda');
    let fila = -1;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cT] || '').trim() === tienda) { fila = i + 1; break; }
    }
    const valores = enc.map(function (k) { return nuevo[k] !== undefined ? nuevo[k] : ''; });
    if (fila === -1) sh.appendRow(valores);
    else sh.getRange(fila, 1, 1, enc.length).setValues([valores]);

    /**
     * La bitácora, sin el secreto dentro.
     *
     * De la contraseña se anota que cambió y ya. De lo demás sí se anota
     * el valor, porque un enlace equivocado hay que poder rastrearlo — y
     * un enlace no es una llave.
     */
    ACCESOS_CAMPOS.forEach(function (k) {
      if (cambios[k] === undefined) return;
      if (String(antes[k] || '') === String(nuevo[k] || '')) return;
      const secreto = (k === 'clave');
      registrarMovimiento(s, 'Accesos', tienda, k,
        secreto ? (antes[k] ? '(había una)' : '(vacía)') : antes[k],
        secreto ? (nuevo[k] ? '(cambiada)' : '(borrada)') : nuevo[k]);
    });

    SpreadsheetApp.flush();
    libroOlvidar_();
    return { ok: true, accesos: accesosDe_(libro_(s.sheetId), tienda) };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}
