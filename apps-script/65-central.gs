/**
 * Nova · Central (la consola de la plataforma)
 * ─────────────────────────────────────────────────────────────
 * Lo que Nova Central hace y Nova Empresarial no: crear clientes, ver
 * cuántos hay y en qué plan, y controlar el demo.
 *
 * ┌─ POR QUÉ UNA SESIÓN APARTE ────────────────────────────────┐
 * │                                                            │
 * │ La sesión de Empresarial se resuelve contra la hoja Equipo │
 * │ de UN cliente: dice de qué tiendas es esa persona y qué    │
 * │ permisos le dio su dueña. Aquí la pregunta es otra: quién  │
 * │ opera la plataforma. Es quien crea cuentas, ve la          │
 * │ facturación de todos los clientes y enciende el demo.      │
 * │                                                            │
 * │ Si fueran la misma, darle un permiso de más a la admin de  │
 * │ un cliente podría, por un descuido, abrirle la consola de  │
 * │ todos los demás. Son dos preguntas distintas y tienen dos  │
 * │ hojas, dos tokens y dos caducidades.                       │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * Todas las acciones de aquí empiezan por `nc_`, y el enrutador las
 * manda a este archivo ANTES de buscar una sesión de cliente.
 */

/** Cuánto dura una sesión de la consola. Más corta que la de una dueña. */
const TTL_CENTRAL_H = 6;

/** Qué puede hacer cada quien en la consola. */
const ROLES_CENTRAL = {
  socia:      ['ver', 'crear_cliente', 'planes', 'demo', 'facturacion'],
  operadora:  ['ver', 'crear_cliente', 'demo'],
};

function puedeCentral(s, permiso) {
  return (ROLES_CENTRAL[s.rol] || []).indexOf(permiso) !== -1;
}

/**
 * Quién es esta persona en la plataforma.
 *
 * Solo la hoja Plataforma de Nova_Central. No se cae hacia Equipo: si
 * alguien no está aquí, no entra, aunque sea dueña de tres tiendas.
 */
function buscarOperadora(email) {
  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName('Plataforma');
  if (!sh || sh.getLastRow() < 2) return null;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const buscado = String(email || '').toLowerCase().trim();

  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    if (String(f[c('correo')] || '').toLowerCase().trim() !== buscado) continue;
    if (norm(f[c('estado')]) === 'inactivo') return null;
    const rol = norm(f[c('rol')]) === 'operadora' ? 'operadora' : 'socia';
    return { id: f[c('id')], nombre: f[c('nombre')], correo: buscado,
             rol: rol, fila: i + 1 };
  }
  return null;
}

function sesionCentral(token) {
  const crudo = CacheService.getScriptCache().get('nc_' + String(token || ''));
  if (!crudo) return null;
  const s = JSON.parse(crudo);
  if (Date.now() > s.vence) return null;
  return s;
}

/**
 * El enrutador de la consola.
 *
 * Va antes que el de clientes porque estas acciones no tienen sesión de
 * cliente que buscar: buscarla primero devolvería "sesión inválida" a
 * quien está intentando entrar a la consola, que es confuso.
 */
function manejarCentral(accion, p) {
  if (accion === 'nc_login')     return centralLogin(p);
  if (accion === 'nc_verificar') return centralVerificar(p);

  const s = sesionCentral(p.token);
  if (!s) return { ok: false, error: 'Sesión vencida o inválida.', reautenticar: true };

  switch (accion) {
    case 'nc_yo':       return { ok: true, sesion: publicoCentral(s) };
    case 'nc_clientes': return centralClientes(s, p);
    case 'nc_planes':   return centralPlanes(s, p);
    case 'nc_crear':    return centralCrearCliente(s, p);
    case 'nc_automatico':        return centralAutomatico(s, p);
    case 'nc_automatico_prender':return centralPrenderAutomatico(s, p);
    case 'nc_mio':          return centralMio(s, p);
    case 'nc_mio_guardar':  return centralMioGuardar(s, p);
    case 'nc_mio_borrar':   return centralMioBorrar(s, p);
    case 'nc_mio_cobrar':   return centralMioCobrar(s, p);

    /**
     * NovaSoul entra por aquí, con la misma sesión.
     *
     * Ella lo decidió: Central y Soul son suyas y de nadie más, y pedir
     * un segundo código para la misma persona en la misma máquina no
     * protege nada. Lo que sí separa las dos es el rol — `soulPuede_`
     * exige socia, así que una operadora con sesión de Central no abre
     * NovaSoul aunque escriba la acción a mano.
     */
    case 'nc_soul':            return soulHoy(s, p);
    case 'nc_soul_guardar':    return soulPendienteGuardar(s, p);
    case 'nc_soul_borrar':     return soulPendienteBorrar(s, p);
    case 'nc_soul_horas':      return soulHorasGuardar(s, p);
    case 'nc_soul_mindlab':    return soulMindlabGuardar(s, p);
    case 'nc_soul_mindlab_bajar': return soulMindlabAPendientes(s, p);
    case 'nc_soul_finanzas':   return soulFinanzas(s, p);
    case 'nc_soul_fijo':       return soulFijoGuardar(s, p);
    case 'nc_soul_fijo_borrar':return soulFijoBorrar(s, p);
    case 'nc_soul_family':     return soulFamily(s, p);
    case 'nc_soul_materias':       return soulMaterias(s, p);
    case 'nc_soul_materia':        return soulMateriaGuardar(s, p);
    case 'nc_soul_materia_borrar': return soulMateriaBorrar(s, p);
    case 'nc_soul_silabo':         return soulSilaboLeer(s, p);
    case 'nc_soul_silabo_guardar': return soulSilaboGuardar(s, p);
    case 'nc_salir':
      CacheService.getScriptCache().remove('nc_' + p.token);
      return { ok: true };
    default: return { ok: false, error: 'Acción desconocida: ' + accion };
  }
}

function publicoCentral(s) {
  return { correo: s.correo, nombre: s.nombre, rol: s.rol,
           vence: s.vence, horas: TTL_CENTRAL_H,
           permisos: ROLES_CENTRAL[s.rol] || [] };
}

/**
 * El código de acceso.
 *
 * Responde lo mismo exista la persona o no. Decir "ese correo no está en
 * la plataforma" convierte esta pantalla en una forma de averiguar quién
 * trabaja aquí, y esta consola crea cuentas de clientes.
 */
function centralLogin(p) {
  const email = String(p.email || '').toLowerCase().trim();
  if (!email || email.indexOf('@') === -1) {
    return { ok: false, error: 'Correo inválido.' };
  }
  const op = buscarOperadora(email);
  if (op) {
    const codigo = String(Math.floor(100000 + Math.random() * 900000));
    CacheService.getScriptCache().put('nccod_' + email, codigo, TTL_CODIGO_M * 60);
    try {
      MailApp.sendEmail({
        to: email,
        subject: 'Tu código de Nova Central: ' + codigo,
        body: 'Hola ' + (op.nombre || '') + ',\n\n' +
              'Tu código para entrar a Nova Central es: ' + codigo + '\n\n' +
              'Vence en ' + TTL_CODIGO_M + ' minutos.\n' +
              'Desde ahí se crean cuentas de clientes: si no fuiste tú, ' +
              'avísale a la otra socia.\n',
      });
    } catch (err) {
      return { ok: false, error: 'No se pudo enviar el código: ' + err.message };
    }
  }
  return { ok: true, enviado: true, vence_en_min: TTL_CODIGO_M };
}

function centralVerificar(p) {
  const email = String(p.email || '').toLowerCase().trim();
  const codigo = String(p.codigo || '').trim();
  const cache = CacheService.getScriptCache();
  const esperado = cache.get('nccod_' + email);

  if (!esperado || esperado !== codigo) {
    return { ok: false, error: 'Código incorrecto o vencido.' };
  }
  cache.remove('nccod_' + email);

  const op = buscarOperadora(email);
  if (!op) return { ok: false, error: 'Esta cuenta ya no tiene acceso.' };

  const token = Utilities.getUuid();
  const s = { correo: email, nombre: op.nombre, rol: op.rol, id: op.id,
              vence: Date.now() + TTL_CENTRAL_H * 3600000 };
  cache.put('nc_' + token, JSON.stringify(s), TTL_CENTRAL_H * 3600);

  // Rastro de quién entra a la consola, en la propia hoja
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Plataforma');
    const e = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const col = e.indexOf('ultima_conexion');
    if (col !== -1) sh.getRange(op.fila, col + 1).setValue(ahoraISO());
  } catch (err) { /* el rastro no puede impedir entrar */ }

  return { ok: true, token: token, sesion: publicoCentral(s) };
}

/**
 * Los clientes de la plataforma.
 *
 * Cada fila trae además lo que no está en Clientes y sí importa para
 * saber si esa cuenta está viva: cuántas personas tiene en su Equipo y
 * cuándo fue la última importación. Un cliente que pagó y no ha subido
 * un archivo en tres semanas es el que hay que llamar hoy.
 *
 * Abrir la hoja de cada cliente cuesta, así que se hace solo si lo
 * piden: la lista sola se dibuja al instante.
 */
function centralClientes(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve los clientes.' };

  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName('Clientes');
  if (!sh || sh.getLastRow() < 2) return { ok: true, clientes: [] };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const conDetalle = !!p.detalle;
  const verDinero = puedeCentral(s, 'facturacion');

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const f = d[i];
    const empresa = String(f[c('empresa')] || '').trim();
    if (!empresa) continue;

    const cl = {
      id: f[c('id')], empresa: empresa,
      pais: String(f[c('pais')] || ''),
      plan: String(f[c('plan')] || ''),
      estado: norm(f[c('estado')]) || 'activo',
      alta: aISO(f[c('fecha_alta')], 'UTC') || '',
      corte: aISO(f[c('fecha_corte')], 'UTC') || '',
      ultimoPago: aISO(f[c('ultimo_pago')], 'UTC') || '',
      tiendas: num(f[c('tiendas')]),
      usuarios: num(f[c('usuarios')]),
      sheetId: String(f[c('sheet_id')] || ''),
    };
    // La tarifa es dinero: solo la socia
    if (verDinero) {
      cl.tarifa = num(f[c('tarifa')]);
      cl.costo = num(f[c('costo')]);
    }

    if (conDetalle && cl.sheetId) {
      try {
        const cs = SpreadsheetApp.openById(cl.sheetId);
        const shE = cs.getSheetByName('Equipo');
        cl.personas = shE && shE.getLastRow() > 1 ? shE.getLastRow() - 1 : 0;
        const shF = cs.getSheetByName('Fuentes');
        if (shF && shF.getLastRow() > 1) {
          const df = shF.getDataRange().getValues();
          const ef = df[0].map(norm);
          const ci = ef.indexOf('ultima_importacion');
          let ult = '';
          for (let j = 1; j < df.length; j++) {
            const u = aISO(df[j][ci], 'UTC');
            if (u && u > ult) ult = u;
          }
          cl.ultimaImportacion = ult;
        }

        /**
         * Estados que ese cliente tiene sin clasificar.
         *
         * Aparece aquí porque el cliente no siempre se da cuenta: Nova se
         * lo dice en su pantalla, pero quien vende Nova necesita saberlo
         * antes de que le llamen diciendo que las cifras no cuadran. Un
         * estado nuevo de una transportadora suele afectar a varios
         * clientes del mismo país a la vez.
         */
        const shEst = cs.getSheetByName('Estados');
        if (shEst && shEst.getLastRow() > 1) {
          const de = shEst.getDataRange().getValues();
          const ee = de[0].map(norm);
          const cEst = ee.indexOf('estado_nova'), cTx = ee.indexOf('texto'),
                cPed = ee.indexOf('pedidos');
          const pend = [];
          for (let j = 1; j < de.length; j++) {
            if (String(de[j][cTx] || '').trim() && !norm(de[j][cEst])) {
              pend.push({ texto: String(de[j][cTx]), pedidos: num(de[j][cPed]) });
            }
          }
          pend.sort(function (a, b) { return b.pedidos - a.pedidos; });
          cl.estadosSinClasificar = pend.length;
          cl.pedidosSinClasificar = pend.reduce(function (t, x) { return t + x.pedidos; }, 0);
          cl.estadosNuevos = pend.slice(0, 5);
        }
      } catch (err) {
        // Una hoja borrada o sin permiso no puede tumbar la lista entera
        cl.problema = 'No se pudo abrir su hoja: ' + err.message;
      }
    }
    out.push(cl);
  }

  return { ok: true, clientes: out, verDinero: verDinero };
}

function centralPlanes(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve los planes.' };
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Planes');
  if (!sh || sh.getLastRow() < 2) return { ok: true, planes: [] };

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const verDinero = puedeCentral(s, 'facturacion');

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const nombre = String(d[i][c('nombre')] || '').trim();
    if (!nombre) continue;
    const pl = { id: d[i][c('id')], nombre: nombre,
      modulos: String(d[i][c('modulos')] || '').split(/[,;]/)
        .map(function (x) { return x.trim(); }).filter(String),
      limiteUsuarios: num(d[i][c('limite_usuarios')]),
      limiteTiendas: num(d[i][c('limite_tiendas')]) };
    if (verDinero) {
      pl.precio = num(d[i][c('precio_sugerido')]);
      pl.tarifa = num(d[i][c('tarifa_fijada')]);
    }
    out.push(pl);
  }
  return { ok: true, planes: out, verDinero: verDinero };
}

/**
 * Crear un cliente desde la pantalla.
 *
 * Hasta ahora crearCliente() solo se podía llamar desde el editor de
 * Apps Script. Eso significa que vender una cuenta requería abrir el
 * código, y que nadie más que quien lo escribió podía hacerlo.
 *
 * Lo que se valida aquí y no dentro de crearCliente: que quien pide
 * tenga permiso, y que lo que llega de la pantalla tenga la forma que
 * esa función espera. Las reglas del negocio —que no falte la dueña, que
 * no haya dos tiendas con el mismo id, que la empresa no exista ya—
 * siguen viviendo allá, porque también valen cuando se llama a mano.
 */
function centralCrearCliente(s, p) {
  if (!puedeCentral(s, 'crear_cliente')) {
    return { ok: false, error: 'Tu rol no puede crear clientes.' };
  }

  const empresa = String(p.empresa || '').trim();
  const pais = String(p.pais || '').trim().toUpperCase();
  const plan = String(p.plan || 'Base').trim();
  const dueno = p.dueno || {};

  if (!empresa) return { ok: false, error: 'Falta el nombre de la empresa.' };
  if (!String(dueno.correo || '').trim() || String(dueno.correo).indexOf('@') === -1) {
    return { ok: false, error: 'Falta el correo de la dueña. Sin una dueña en su ' +
             'hoja Equipo, nadie podría entrar a esa cuenta — ni ella.' };
  }

  const tiendas = (p.tiendas || []).filter(function (t) {
    return t && String(t.id || '').trim();
  }).map(function (t) {
    return {
      id: String(t.id).trim().toLowerCase(),
      nombre: String(t.nombre || t.id).trim(),
      pais: String(t.pais || pais || '').toUpperCase(),
      moneda: String(t.moneda || '').toUpperCase(),
      zona: String(t.zona || '').trim(),
      modalidad: String(t.modalidad || 'catalogo_publico').trim(),
    };
  });
  if (!tiendas.length) {
    return { ok: false, error: 'Hay que declarar al menos una tienda.' };
  }
  const sinMoneda = tiendas.filter(function (t) { return !t.moneda; });
  if (sinMoneda.length) {
    return { ok: false, error: 'Falta la moneda de: ' +
             sinMoneda.map(function (t) { return t.nombre; }).join(', ') +
             '. Sin moneda, Nova no puede sumar ni convertir nada de esa tienda.' };
  }

  const fuentes = (p.fuentes || []).filter(function (f) {
    return f && f.tienda && f.fuente;
  });

  let r;
  try {
    r = crearCliente(empresa, pais, tiendas, fuentes, plan, {
      nombre: String(dueno.nombre || '').trim() || 'Dueña',
      correo: String(dueno.correo).toLowerCase().trim(),
    });
  } catch (err) {
    return { ok: false, error: err.message };
  }

  registrarCentral(s, 'Clientes', r.clienteId, 'creado', '',
                   empresa + ' · ' + tiendas.length + ' tienda(s) · plan ' + plan);
  return { ok: true, clienteId: r.clienteId, url: r.url, empresa: empresa,
           dueno: String(dueno.correo).toLowerCase().trim() };
}

/** La bitácora de la consola vive en Nova_Central, no en la del cliente. */
function registrarCentral(s, entidad, id, campo, antes, ahora) {
  try {
    const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Movimientos');
    if (sh) sh.appendRow([ahoraISO(), s.correo, entidad, id, campo, antes, ahora]);
  } catch (err) {
    Logger.log('No se pudo registrar en Central: ' + err.message);
  }
}

/**
 * ★ Darse de alta en la consola, la primera vez ★
 *
 * Se entra a Nova Central con un correo que esté en la hoja Plataforma,
 * y esa hoja nace vacía: sin esto nadie podría entrar nunca.
 *
 * SE CORRE SIN ESCRIBIR NADA. Eliges primeraSocia en el editor, le das a
 * Ejecutar, y se da de alta la cuenta desde la que estás corriendo el
 * script — que es la dueña de las hojas, así que ya es tuya.
 *
 * El botón Ejecutar del editor no sabe pasar argumentos, y pedirle a
 * alguien que edite el código para darse de alta es pedirle que toque lo
 * único que no debería tener que tocar. Si quieres otro correo, se puede
 * pasar: primeraSocia("Nombre", "otro@correo.com")
 *
 * Solo funciona mientras la hoja esté vacía. Después, quien agrega gente
 * es quien ya está dentro.
 */
function primeraSocia(nombre, correo) {
  const sh = SpreadsheetApp.openById(IDS_().central).getSheetByName('Plataforma');
  if (!sh) throw new Error('Falta la hoja Plataforma. Corre bootstrapTodo().');
  if (sh.getLastRow() > 1) {
    const ya = sh.getRange(2, 1, sh.getLastRow() - 1, sh.getLastColumn()).getValues()
      .map(function (f) { return f[2]; }).filter(String).join(', ');
    throw new Error(
      'La hoja Plataforma ya tiene gente: ' + ya + '\n\n' +
      'Para agregar a alguien más, escribe su fila directamente ahí: ' +
      'nombre, correo, rol (socia u operadora) y estado "activo".');
  }

  // Sin correo, el de la cuenta que está corriendo esto
  const mail = String(correo || Session.getEffectiveUser().getEmail() || '')
    .toLowerCase().trim();
  if (!mail || mail.indexOf('@') === -1) {
    throw new Error(
      'No pude averiguar tu correo solo. Córrela así:\n\n' +
      '  primeraSocia("Manuela", "tucorreo@dominio.com")');
  }

  sembrar(sh.getParent(), 'Plataforma', [{
    id: 'pl-' + Utilities.getUuid().slice(0, 8),
    nombre: nombre || mail.split('@')[0],
    correo: mail,
    rol: 'socia', estado: 'activo', ultima_conexion: '', nota: 'primera socia',
  }], 2);

  const msg = [
    'Listo.',
    '',
    'Entra a Nova Central con este correo:  ' + mail,
    'Ahí te llega el código de 6 dígitos.',
    '',
    'Si prefieres otro correo, bórralo de la hoja Plataforma y vuelve a',
    'correr esto pasándole el que quieras.',
  ].join('\n');
  Logger.log(msg);
  return msg;
}


// ─── LO AUTOMÁTICO, DESDE LA CONSOLA ─────────────────────────

/**
 * ═══════════════════════════════════════════════════════════
 *  EL INTERRUPTOR DE LO AUTOMÁTICO
 * ═══════════════════════════════════════════════════════════
 *
 * Los disparadores que hacen correr a Nova sola —las tasas de cambio y
 * la revisión de alarmas— vivían únicamente en el Apps Script. Estaban
 * escritos desde el principio y nunca se prendieron, porque prenderlos
 * exigía abrir el editor, encontrar la función y ejecutarla a mano.
 *
 * Eso no se le puede pedir a nadie, y menos a un cliente. Pero tampoco
 * va en la pantalla de la dueña: los disparadores son del proyecto
 * entero, no de cada cuenta, y un botón ahí le daría a cualquier cliente
 * un interruptor que afecta a todos los demás.
 *
 * Así que va aquí, en la consola, que es de Manuela.
 */
function centralAutomatico(s, p) {
  if (!puedeCentral(s, 'ver')) return { ok: false, error: 'Tu rol no ve esto.' };
  return { ok: true, automatico: estadoAutomatico() };
}

/**
 * Prende los dos trabajos y programa la carga de tasas para dentro de un
 * minuto.
 *
 * La carga NO se hace aquí. Bajar noventa días de tasas de GOOGLEFINANCE
 * para varios clientes tarda más de lo que un navegador espera, y el
 * botón se quedaría colgado sin que nadie supiera si funcionó. En vez de
 * eso se arma un disparador de un solo uso: la respuesta vuelve al
 * instante y el trabajo pesado corre solo, por detrás.
 */
function centralPrenderAutomatico(s, p) {
  if (!puedeCentral(s, 'crear_cliente')) {
    return { ok: false, error: 'Tu rol no puede prender lo automático.' };
  }

  let hechos;
  try {
    hechos = prenderTrabajos_();
  } catch (e) {
    /**
     * Crear disparadores pide un permiso que la autorización vieja del
     * script puede no tener. No se disfraza de otro error: se dice qué
     * pasó y cuál es la salida, que es abrir el editor una vez.
     */
    return { ok: false, error:
      'No pude crear los disparadores: ' + e.message + '\n\n' +
      'Suele ser que la autorización del script es anterior a esta ' +
      'función. Abre el Apps Script, corre prenderAutomatico() una vez ' +
      'y acepta los permisos; después este botón ya funciona.' };
  }

  // Y la carga inicial, por detrás.
  let cargando = false;
  try {
    const yaHay = ScriptApp.getProjectTriggers().some(function (t) {
      return t.getHandlerFunction() === 'cargaInicialTasas';
    });
    if (!yaHay) {
      ScriptApp.newTrigger('cargaInicialTasas').timeBased().after(60 * 1000).create();
    }
    cargando = true;
  } catch (e) {
    cargando = false;
  }

  return {
    ok: true,
    prendidos: hechos,
    cargando: cargando,
    automatico: estadoAutomatico(),
    mensaje: cargando
      ? 'Listo. Las tasas de los últimos 90 días empiezan a bajar en un minuto; ' +
        'según cuántas cuentas haya puede tardar varios. Vuelve a mirar en un rato.'
      : 'Los trabajos quedaron prendidos, pero no pude programar la carga de ' +
        'las tasas viejas. Córrela a mano: cargaInicialTasas() en el Apps Script.',
  };
}
