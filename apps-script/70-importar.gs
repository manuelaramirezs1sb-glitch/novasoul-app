/**
 * Nova · Importar (el lado que escribe)
 * ─────────────────────────────────────────────────────────────
 * `leerCrudo()` normaliza pero no guarda nada. Esto es lo que convierte
 * un export pegado en filas que la app puede leer.
 *
 * Tres reglas que no se rompen:
 *
 *   1. NUNCA se sobrescribe lo que escribió el equipo.
 *      Si una gestora puso una nota o cambió el estado, la siguiente
 *      importación no lo borra. Se actualizan solo las columnas que
 *      vienen de la plataforma.
 *
 *   2. NUNCA se borran filas.
 *      Un pedido que desaparece del export no se elimina: el histórico
 *      es lo que hace posibles las tendencias.
 *
 *   3. Deduplicación por fuente + id_externo.
 *      Reimportar el mismo archivo dos veces no duplica nada, actualiza.
 *
 * USO:
 *   1. Pega el export en su pestaña _Import_*
 *   2. importar('dropi', 'ec')
 */

// Columnas que escribe la app y el importador jamás toca.
const COLUMNAS_DEL_EQUIPO = [
  'estado_nova', 'nota', 'solucion', 'gestora_asignada', 'fecha_promesa',
  'intentos', 'resuelta_en', 'telefono_2', 'telefono_2_norm',
  'actualizado_en', 'actualizado_por',
];

/**
 * Importa una fuente a su entidad.
 *
 * @param {string} fuenteId  'dropi' | 'mastershop' | 'effi_guias' | ...
 * @param {string} tienda    id de la tienda: 'gt' | 'ec'
 * @param {string} cliente   nombre del cliente (o vacío si solo hay uno)
 */
function importar(fuenteId, tienda, cliente) {
  const ss = libro_(hojaCliente(cliente));
  if (!tienda) throw new Error('Falta la tienda. Ej: importar("dropi","ec")');
  if (!monedaDeTienda(ss, tienda)) {
    throw new Error('La tienda "' + tienda + '" no está en la hoja Tiendas.');
  }

  // Todo el import corre con el formato de fecha que eligió la dueña
  return conFormatoFecha(ss, tienda, function () {
    return importarConFormato(ss, fuenteId, tienda);
  });
}

function importarConFormato(ss, fuenteId, tienda) {
  const r = leerCrudo(ss, fuenteId, tienda);
  if (!r.filas.length) {
    const msg = r.tab + ' está vacía. Pega el export ahí primero.';
    Logger.log(msg);
    return msg;
  }

  const destino = { pedidos: 'Pedidos', novedades: 'Novedades',
                    llamadas: 'Llamadas', pauta: 'Pauta',
                    facturacion: 'Facturacion', cartera: 'Cartera',
                    pedidos_secundario: 'Pedidos' }[r.tipo];
  if (!destino) throw new Error('No sé dónde guardar una fuente de tipo ' + r.tipo);

  let extra = '';
  const pais = paisDeTienda(ss, tienda);

  /**
   * Lo que la dueña ya clasificó, leído una sola vez.
   *
   * Si cada fila abriera la hoja Estados, un archivo de cuarenta mil
   * pedidos la abriría cuarenta mil veces y el script se quedaría sin sus
   * seis minutos antes de escribir nada.
   */
  const aprendidos = estadosAprendidos(ss);
  const vistos = {};

  const preparadas = r.filas.map(function (f) {
    return prepararFila(f, r.tipo, fuenteId, tienda, pais, ss, aprendidos, vistos);
  });

  // Lo que se vio queda anotado, se entendiera o no
  anotarEstados(ss, vistos);

  /**
   * Qué cambió respecto a la carga anterior.
   *
   * "206 actualizadas" no dice nada. Lo que alguien necesita saber al
   * subir el archivo del día es qué pasó desde ayer: cuántos llegaron,
   * cuántos se entregaron, cuántos se devolvieron y qué queda por
   * resolver. Eso se calcula comparando el estado que había en la hoja
   * contra el que trae el archivo, ANTES de escribir encima.
   */
  const antes = (r.tipo === 'pedidos') ? estadosActuales(ss, destino, tienda) : null;

  const res = escribirFilas(ss, destino, preparadas, fuenteId);
  registrarImportacion(ss, tienda, fuenteId, res.nuevas + res.actualizadas);

  if (antes) extra = novedadesDeLaCarga(antes, preparadas) + extra;

  // Las novedades que vienen dentro del export de pedidos se derivan aparte
  if (r.tipo === 'pedidos' && fuenteId === 'dropi') {
    const nov = derivarNovedades(preparadas, fuenteId, tienda);
    if (nov.length) {
      const rn = escribirFilas(ss, 'Novedades', nov, fuenteId);
      extra = '\nNovedades derivadas: ' + rn.nuevas + ' nuevas, ' +
              rn.actualizadas + ' actualizadas, ' + rn.iguales + ' sin cambios';
    }
  }

  /**
   * Una fila de pauta que cubre dos meses no se puede repartir sin
   * inventar: el gasto de una campaña no se distribuye parejo por día.
   *
   * Así que no se reparte — se avisa. Si no, todo el gasto de agosto a
   * septiembre se contaría en agosto y septiembre saldría en cero, que
   * es peor que un número que falta: es un número que miente.
   */
  if (r.tipo === 'pauta') {
    const cruzan = preparadas.filter(function (p) {
      return p.fecha_fin && String(p.fecha).slice(0, 7) !== String(p.fecha_fin).slice(0, 7);
    });
    if (cruzan.length) {
      extra += '\n⚠ ' + cruzan.length + ' fila(s) cubren más de un mes (' +
        preparadas[0].fecha + ' a ' + preparadas[0].fecha_fin + ').\n' +
        '   Todo ese gasto se contará en el primer mes. Para que cada mes ' +
        'reciba lo suyo,\n   vuelve a exportar en Meta con Desglose → Por día.';
    }
  }

  const msg = [
    'Importado: ' + fuenteId + ' → ' + destino + ' (tienda ' + tienda + ')',
    '  filas leídas   : ' + r.filas.length,
    '  nuevas         : ' + res.nuevas,
    '  actualizadas   : ' + res.actualizadas,
    '  sin cambios    : ' + res.iguales,
    r.sinMapear.length ? '  columnas sin mapear: ' + r.sinMapear.join(', ') : '',
    res.sinEstado.length
      ? '  ⚠ ' + res.sinEstado.length + ' estado(s) que no reconozco: ' +
        res.sinEstado.slice(0, 8).join(' · ') +
        '\n    Esos pedidos quedan SIN CLASIFICAR: no cuentan como entregados ' +
        'ni como devueltos.\n    Dinos qué significan en Configuración → Estados ' +
        'y las cifras se rehacen solas.'
      : '',
    extra,
  ].filter(String).join('\n');
  Logger.log(msg);
  return msg;
}

/** El estado que tiene hoy cada pedido en la hoja, por id. */
function estadosActuales(ss, hoja, tienda) {
  const sh = ss.getSheetByName(hoja);
  if (!sh || sh.getLastRow() < 2) return {};
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cId = e.indexOf('id'), cT = e.indexOf('tienda');
  const cE = e.indexOf('estado_canonico'), cN = e.indexOf('estado_nova');
  const out = {};
  for (let i = 1; i < d.length; i++) {
    if (cT !== -1 && String(d[i][cT]).trim() !== tienda) continue;
    const id = String(d[i][cId]).trim();
    if (id) out[id] = norm(d[i][cN] || d[i][cE]);
  }
  return out;
}

/**
 * El resumen de lo que cambió, en la lengua de quien opera.
 *
 * No dice "38 filas actualizadas": dice qué pasó con los pedidos.
 */
function novedadesDeLaCarga(antes, ahora) {
  let nuevos = 0, entregados = 0, devueltos = 0, cancelados = 0,
      aNovedad = 0, pendientes = 0, sinCambio = 0;

  ahora.forEach(function (p) {
    const previo = antes[p.id];
    const est = norm(p.estado_canonico);
    if (previo === undefined) {
      nuevos++;
    } else if (previo !== est) {
      if (est === 'entregado')  entregados++;
      else if (est === 'devolucion') devueltos++;
      else if (est === 'cancelado')  cancelados++;
      else if (est === 'novedad')    aNovedad++;
    } else {
      sinCambio++;
    }
    if (['entregado','devolucion','cancelado'].indexOf(est) === -1) pendientes++;
  });

  // Uno o varios: "se devolvieron 1" delata que lo escribió una máquina
  const pl = function (n, uno, varios) { return n === 1 ? uno : varios.replace('%', n); };

  const partes = [];
  if (nuevos)     partes.push(pl(nuevos, 'entró 1 pedido nuevo', 'entraron % pedidos nuevos'));
  if (entregados) partes.push(pl(entregados, 'se entregó 1', 'se entregaron %'));
  if (devueltos)  partes.push(pl(devueltos, 'se devolvió 1', 'se devolvieron %'));
  if (cancelados) partes.push(pl(cancelados, 'se canceló 1', 'se cancelaron %'));
  if (aNovedad)   partes.push(pl(aNovedad, '1 entró en novedad', '% entraron en novedad'));

  if (!partes.length && !pendientes) return '';

  let msg = '\n\nDesde la carga anterior: ';
  msg += partes.length ? partes.join(', ') + '.' : 'ningún pedido cambió de estado.';
  if (pendientes) {
    msg += '\n' + pl(pendientes, 'Queda 1 pedido sin desenlace.',
                     'Quedan % pedidos sin desenlace.');
  }
  return msg;
}

/** Convierte una fila normalizada en una fila lista para la entidad. */
/**
 * Las columnas que son fechas, en cualquier hoja.
 *
 * Todas se guardan en AAAA-MM-DD. Lo que llega del archivo puede venir
 * como 12-07-2026 o 07/12/2026 según la plataforma y el país, y esas dos
 * cosas se ven iguales: una es 12 de julio y la otra 7 de diciembre.
 * Mientras se guarde así, cada pantalla tiene que volver a adivinar, y
 * el navegador adivina a la americana —mes primero— sin avisar.
 */
const CAMPOS_FECHA = ['fecha', 'fecha_entrega', 'fecha_promesa', 'fecha_ingreso',
                      'fecha_solucion', 'fecha_fin', 'ultimo_movimiento',
                      'actualizado', 'creado_en', 'ultimo_conteo'];

function prepararFila(f, tipo, fuenteId, tienda, pais, ss, aprendidos, vistos) {
  const o = Object.assign({}, f);
  o.fuente = fuenteId;
  o.tienda = tienda;

  /**
   * La fecha se normaliza aquí, una vez, y ya nadie más tiene que dudar.
   *
   * Antes se guardaba tal como venía y cada lector la interpretaba por su
   * cuenta: el servidor con aISO —que sabe que en la región el día va
   * primero— y la pantalla con new Date(), que asume el formato de
   * Estados Unidos. Por eso un pedido del 12 de julio salía en pantalla
   * como 7 de diciembre.
   *
   * La hora se conserva cuando viene: el gráfico de la jornada y el
   * seguimiento de novedades la necesitan.
   */
  CAMPOS_FECHA.forEach(function (k) {
    if (o[k] === undefined || o[k] === '' || o[k] === null) return;
    const iso = aISO(o[k], 'UTC');
    if (!iso) return;
    const hora = String(o[k]).match(/\b(\d{1,2}:\d{2}(?::\d{2})?)\b/);
    o[k] = hora ? iso + ' ' + hora[1] : iso;
  });

  // El id es fuente + id externo: estable entre importaciones, y deja
  // ver de dónde salió cada fila sin abrir el export.
  const ext = String(o.id_externo || o.guia || '').trim();
  o.id_externo = ext;
  o.id = fuenteId + '-' + (ext || Utilities.getUuid().slice(0, 8));

  if (o.telefono) {
    o.telefono_norm = telefonoNorm(o.telefono, pais);
    // El número alterno suele estar escondido en la nota o la observación
    const libre = [o.nota, o.observacion, o.aclaracion].filter(String).join(' ');
    if (libre) {
      const t = extraerTelefonos(o.telefono, libre, pais);
      if (t.secundario) { o.telefono_2 = t.secundario; o.telefono_2_norm = t.secundario; }
    }
  }

  if (o.documento) {
    const d = partirDocumento(o.documento);
    o.cedula = d.numero;
    delete o.documento;
  }

  if (tipo === 'pedidos' || tipo === 'pedidos_secundario') {
    const r = estadoConOrigen(fuenteId, o.estado, aprendidos);
    o.estado_canonico = r.estado;
    if (vistos && r.clave) {
      const k = fuenteId + '|' + r.clave;
      if (!vistos[k]) {
        vistos[k] = { fuente: fuenteId, texto: r.clave,
                      estado: r.estado, origen: r.origen, n: 0 };
      }
      vistos[k].n++;
    }
  }
  if (tipo === 'novedades') {
    o.grupo = grupoNovedad(o.motivo, o.codigo);
    /**
     * ── LO DEL ARCHIVO NO SE MEZCLA CON LO DEL EQUIPO ──
     *
     * Antes la aclaración del archivo se copiaba dentro de `nota`, que
     * es la columna donde escribe el equipo, «si estaba vacía». Y el
     * histórico de soluciones caía dentro de `solucion`, igual.
     *
     * Parecía inofensivo porque el importador no pisa esas dos columnas
     * al actualizar. Pero significa que la primera importación deja el
     * texto de la plataforma haciéndose pasar por un apunte del equipo,
     * y después nadie puede distinguir cuál es cuál — ni saber si lo
     * que lee lo escribió una compañera aquí o venía en el Excel.
     *
     * Ahora cada uno tiene su columna. La pantalla las muestra por
     * separado y dice de dónde salió cada una.
     */
    if (o.solucion !== undefined) {
      if (!o.solucion_plataforma) o.solucion_plataforma = o.solucion;
      delete o.solucion;
    }
  }
  if (tipo === 'llamadas') {
    o.seg_conversado = aSegundos(o.seg_conversado);
    o.seg_espera     = aSegundos(o.seg_espera);
    o.seg_total      = aSegundos(o.seg_total);
    // IRIS no dice de qué tienda es la llamada: se cruza por teléfono
    o.tienda = '';
  }
  if (tipo === 'facturacion') {
    o.plataforma = fuenteId.replace('_facturacion', '');
    o.id = fuenteId + '-' + (o.id_externo || Utilities.getUuid().slice(0, 8));
    const mon = String(o.moneda_gasto || (FUENTES[fuenteId] || {}).moneda_default || '').toUpperCase();
    o.moneda_gasto = mon;
    const destino = monedaReporte(ss);
    o.moneda_reporte = destino;
    if (mon && o.fecha) {
      const c = convertir(ss, o.gasto, o.fecha, mon, destino);
      // Sin tasa no se inventa un número: queda vacío y visible
      o.gasto_normalizado = c.valor === null ? '' : c.valor;
    }
  }

  if (tipo === 'cartera') {
    o.id = fuenteId + '-' + (o.id_externo || Utilities.getUuid().slice(0, 8));
    o.tipo = String(o.tipo_movimiento || '').trim().toUpperCase();
    delete o.tipo_movimiento;
    o.clase = claseMovimiento(o.descripcion, o.concepto_retiro);
    o.importado_en = ahoraISO();
    /**
     * El signo se guarda en el monto, no en la cabeza de quien lee.
     *
     * Dropi manda todo positivo y dice aparte si fue ENTRADA o SALIDA.
     * Guardarlo así obliga a acordarse del signo cada vez que se suma, y
     * tarde o temprano alguien suma una devolución como ingreso.
     */
    if (o.tipo === 'SALIDA') o.monto = -Math.abs(num(o.monto));
    else o.monto = Math.abs(num(o.monto));
  }

  if (tipo === 'pauta') {
    o.plataforma = fuenteId;
    /**
     * Una fila de pauta no trae identificador propio, así que se arma uno
     * con lo que la hace única: plataforma, tienda, periodo y conjunto.
     *
     * El periodo entra entero —inicio y fin— porque Meta exporta el mismo
     * conjunto para rangos distintos. Sin el fin, volver a exportar con
     * otro rango pisaría la fila anterior y el gasto del mes cambiaría
     * solo, sin que nadie hubiera tocado nada.
     */
    o.id = [fuenteId, tienda, o.fecha || '', o.fecha_fin || '',
            norm(o.conjunto || o.campana || '')].join('-')
           .replace(/\s+/g, '_').slice(0, 180);
    const mon = String(o.moneda_gasto || (FUENTES[fuenteId] || {}).moneda_default || '').toUpperCase();
    o.moneda_gasto = mon;
    const destino = monedaReporte(ss);
    if (mon && o.fecha) {
      const c = convertir(ss, o.gasto, o.fecha, mon, destino);
      // Sin tasa no se inventa un número: queda vacío y visible
      o.gasto_normalizado = c.valor === null ? '' : c.valor;
    }
  }
  return o;
}

/** "00:01:23" → 83. Las llamadas vienen en hh:mm:ss. */
function aSegundos(v) {
  if (v === '' || v == null) return '';
  const s = String(v).trim();
  const m = s.match(/^(\d+):(\d{2}):(\d{2})$/);
  if (m) return (+m[1]) * 3600 + (+m[2]) * 60 + (+m[3]);
  const n = aNumero(s);
  return n === '' ? '' : n;
}

/**
 * Escribe las filas en su hoja: actualiza las que ya existen, agrega las
 * nuevas. Nunca borra, y nunca pisa las columnas del equipo.
 */
/**
 * Escribe las filas en su entidad, por bloques.
 *
 * Antes escribía celda por celda. Era lo más seguro sobre el papel —una
 * hoja no tiene bloqueo de fila, así que dos escrituras simultáneas se
 * pisan— pero con 206 pedidos y ocho columnas eran mil seiscientas
 * llamadas, una por una, y la importación tardaba más de un minuto.
 * Google corta la respuesta antes y el navegador solo ve un 404: el
 * archivo entraba a medias y parecía que había fallado la subida.
 *
 * Ahora se arma todo en memoria y se escribe por tramos de columnas
 * seguidas. Las columnas del equipo se quedan fuera del tramo, así que
 * siguen intocables; las de la plataforma solo las escribe esto, de modo
 * que devolver su valor actual a una fila que no cambió no pisa nada.
 * Y un candado impide que dos importaciones corran encima.
 */
function escribirFilas(ss, hoja, filas, fuenteId) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('No existe la hoja ' + hoja);

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(30000)) {
    throw new Error('Hay otra importación corriendo. Espera a que termine.');
  }

  try {
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const cId = enc.indexOf('id');
    if (cId === -1) throw new Error(hoja + ' no tiene columna id.');

    // Lo que ya está, indexado, para no recorrer la hoja por cada fila
    let datos = [];
    const existentes = {};
    if (sh.getLastRow() > 1) {
      datos = sh.getRange(2, 1, sh.getLastRow() - 1, enc.length).getValues();
      datos.forEach(function (f, i) {
        const k = String(f[cId]).trim();
        if (k) existentes[k] = i;
      });
    }

    const nuevas = [], sinEstado = {}, tocadas = {};
    let actualizadas = 0, iguales = 0;

    filas.forEach(function (o) {
      if (o.estado_canonico === ESTADOS.SIN_CLASIFICAR) {
        sinEstado[norm(o.estado) || '(vacio)'] = 1;
      }

      const i = existentes[o.id];
      if (i === undefined) {
        nuevas.push(enc.map(function (col) {
          return o[col] !== undefined ? o[col] : '';
        }));
        return;
      }

      let cambio = false;
      enc.forEach(function (col, c) {
        if (COLUMNAS_DEL_EQUIPO.indexOf(col) !== -1) return;
        if (o[col] === undefined || o[col] === '') return;
        if (String(datos[i][c]) === String(o[col])) return;
        datos[i][c] = o[col];
        tocadas[c] = true;
        cambio = true;
      });
      if (cambio) actualizadas++; else iguales++;
    });

    // Tramos de columnas seguidas que hay que reescribir
    if (datos.length && Object.keys(tocadas).length) {
      const cols = Object.keys(tocadas).map(Number).sort(function (a, b) { return a - b; });
      let ini = cols[0], fin = cols[0];
      const escribirTramo = function (a, b) {
        const ancho = b - a + 1;
        const sub = datos.map(function (f) { return f.slice(a, b + 1); });
        sh.getRange(2, a + 1, datos.length, ancho).setValues(sub);
      };
      for (let k = 1; k < cols.length; k++) {
        if (cols[k] === fin + 1) { fin = cols[k]; continue; }
        escribirTramo(ini, fin);
        ini = fin = cols[k];
      }
      escribirTramo(ini, fin);
    }

    if (nuevas.length) {
      sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, enc.length).setValues(nuevas);
    }

    return {
      nuevas: nuevas.length, actualizadas: actualizadas, iguales: iguales,
      sinEstado: Object.keys(sinEstado),
    };
  } finally {
    lock.releaseLock();
  }
}

/**
 * Dropi trae la novedad dentro de la misma fila del pedido, no en un
 * reporte aparte como Effi. Se extraen para que la alarma de patrón y
 * el cierre de mes puedan contarlas.
 */
/**
 * Saca las novedades que vienen dentro del export de pedidos.
 *
 * El estado de la novedad no es el estado del pedido, y confundirlos fue
 * un error caro: antes una novedad solo contaba como resuelta si el
 * pedido estaba EN ESE MOMENTO en "novedad solucionada". Pero un pedido
 * que tuvo novedad y después se entregó ya no está en ese estado, está
 * entregado — así que su novedad quedaba abierta para siempre. Setenta y
 * dos novedades y setenta y dos "abiertas", cuando de verdad quedaban dos.
 *
 * Dropi ya trae la respuesta en sus propias columnas, y ahora se leen.
 * Son tres hechos distintos y cada uno tiene su columna:
 *
 *   solucionada  lo que dice la plataforma: SI o NO
 *   desenlace    cómo terminó el pedido: entregado, devuelto, cancelado
 *   estado       qué hay que hacer hoy con ella
 *
 * Separarlos deja ver el caso que importa: la novedad que el equipo SÍ
 * resolvió y el pedido se devolvió igual. Eso es trabajo que no se
 * convirtió en venta, y mezclado con lo demás no se ve.
 */
function derivarNovedades(pedidos, fuenteId, tienda) {
  const TERMINALES = ['entregado', 'devolucion', 'cancelado'];

  return pedidos
    .filter(function (p) {
      const m = String(p.motivo_novedad || '').trim();
      return m && m !== '.' && m !== '-';
    })
    .map(function (p) {
      const sol = norm(p.solucionada || '');
      const solucionada = sol === 'si' || sol === 'sí' || sol === 'true' || sol === '1';
      const cerrado = TERMINALES.indexOf(p.estado_canonico) !== -1;

      let estado;
      if (solucionada || p.estado_canonico === 'novedad_resuelta') estado = 'resuelta';
      else if (cerrado) estado = 'cerrada';   // terminó sin resolverse
      else estado = 'abierta';                // sigue esperando a alguien

      return {
        id: fuenteId + '-nov-' + p.id_externo,
        fuente: fuenteId,
        id_externo: p.id_externo,
        pedido_id: p.id,
        fecha: p.ultimo_movimiento || p.fecha,
        tipo: 'novedad',
        motivo: p.motivo_novedad,
        grupo: grupoNovedad(p.motivo_novedad),
        estado: estado,
        solucionada: solucionada ? 'si' : 'no',
        fecha_solucion: p.fecha_solucion || '',
        desenlace: cerrado ? p.estado_canonico : '',
        /**
         * ── LO QUE YA HABÍA ESCRITO LA GESTORA ──
         *
         * Dropi trae una columna `solucion` y Effi un «histórico de
         * soluciones»: es lo que escribió quien atendió el caso, a
         * veces durante meses, antes de que existiera Nova.
         *
         * El alias del importador SIEMPRE lo leyó. Esta función nunca
         * lo copió a la fila, así que se leía y se tiraba: se
         * importaban setenta novedades resueltas y las setenta
         * llegaban sin una palabra de lo que se hizo.
         *
         * Va a `solucion_plataforma`, no a `solucion`: la segunda es la
         * que el equipo escribe DENTRO de Nova y el importador tiene
         * prohibido pisarla. Mezclarlas haría que una reimportación
         * borrara lo que alguien escribió aquí.
         */
        solucion_plataforma: String(p.solucion || '').trim(),
        aclaracion: String(p.aclaracion || '').trim(),
      };
    });
}

/** Deja constancia de cuándo se importó y cuántas filas entraron. */
function registrarImportacion(ss, tienda, fuenteId, filas) {
  const sh = ss.getSheetByName('Fuentes');
  if (!sh || sh.getLastRow() < 2) return;
  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cT = enc.indexOf('tienda'), cF = enc.indexOf('fuente');
  const cU = enc.indexOf('ultima_importacion'), cN = enc.indexOf('filas_ultima');

  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cT]).trim() !== tienda) continue;
    if (norm(datos[i][cF]) !== norm(fuenteId.split('_')[0])) continue;
    if (cU !== -1) sh.getRange(i + 1, cU + 1).setValue(ahoraISO());
    if (cN !== -1) sh.getRange(i + 1, cN + 1).setValue(filas);
    return;
  }
}

function paisDeTienda(ss, tienda) {
  const sh = ss.getSheetByName('Tiendas');
  if (!sh || sh.getLastRow() < 2) return '';
  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cId = enc.indexOf('id'), cP = enc.indexOf('pais');
  const ISO = { colombia:'CO', ecuador:'EC', guatemala:'GT', mexico:'MX',
                peru:'PE', chile:'CL', argentina:'AR', bolivia:'BO',
                paraguay:'PY', uruguay:'UY', venezuela:'VE', panama:'PA',
                'costa rica':'CR', honduras:'HN', nicaragua:'NI',
                'el salvador':'SV', brasil:'BR', espana:'ES' };
  for (let i = 1; i < datos.length; i++) {
    if (String(datos[i][cId]).trim() === tienda) {
      return ISO[norm(datos[i][cP])] || '';
    }
  }
  return '';
}

/**
 * Cruza las llamadas de IRIS con los pedidos, por teléfono normalizado.
 * Córrelo después de importar IRIS y los pedidos.
 *
 * Una llamada que no cruza con ningún pedido queda con tienda y pedido_id
 * vacíos — visible, no descartada. Son llamadas a números que no están en
 * ningún pedido, y vale la pena mirarlas.
 */
function cruzarLlamadas(cliente) {
  const ss = libro_(hojaCliente(cliente));
  const shL = ss.getSheetByName('Llamadas');
  const shP = ss.getSheetByName('Pedidos');
  if (!shL || shL.getLastRow() < 2) { Logger.log('No hay llamadas.'); return 'Sin llamadas.'; }
  if (!shP || shP.getLastRow() < 2) { Logger.log('No hay pedidos.'); return 'Sin pedidos.'; }

  const dP = shP.getDataRange().getValues();
  const eP = dP[0].map(norm);
  const cTel = eP.indexOf('telefono_norm'), cPid = eP.indexOf('id'), cPt = eP.indexOf('tienda');

  // Un teléfono puede tener varios pedidos: se queda el más reciente
  const porTel = {};
  const cFe = eP.indexOf('fecha');
  for (let i = 1; i < dP.length; i++) {
    const t = String(dP[i][cTel] || '').trim();
    if (!t) continue;
    const fe = aISO(dP[i][cFe], 'UTC') || '';
    if (!porTel[t] || fe > porTel[t].fecha) {
      porTel[t] = { id: dP[i][cPid], tienda: dP[i][cPt], fecha: fe };
    }
  }

  const dL = shL.getDataRange().getValues();
  const eL = dL[0].map(norm);
  const cLtel = eL.indexOf('telefono_norm'), cLpid = eL.indexOf('pedido_id'),
        cLt = eL.indexOf('tienda');

  let cruzadas = 0, huerfanas = 0;
  for (let i = 1; i < dL.length; i++) {
    const t = String(dL[i][cLtel] || '').trim();
    const m = t && porTel[t];
    if (m) {
      if (String(dL[i][cLpid]) !== String(m.id)) {
        shL.getRange(i + 1, cLpid + 1).setValue(m.id);
        shL.getRange(i + 1, cLt + 1).setValue(m.tienda);
      }
      cruzadas++;
    } else { huerfanas++; }
  }

  const msg = 'Llamadas cruzadas: ' + cruzadas + '\n' +
    'Sin pedido que las reciba: ' + huerfanas +
    (huerfanas ? '\n  (números que no están en ningún pedido — vale la pena revisarlos)' : '');
  Logger.log(msg);
  return msg;
}

/**
 * Importa todo lo que esté pegado y tenga fuente activa. Es lo que
 * conviene correr después de pegar los exports del día.
 */
function importarTodo(cliente) {
  const ss = libro_(hojaCliente(cliente));
  const sh = ss.getSheetByName('Fuentes');
  if (!sh || sh.getLastRow() < 2) return 'No hay fuentes configuradas.';

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cT = enc.indexOf('tienda'), cF = enc.indexOf('fuente'), cA = enc.indexOf('activa');

  // Cuántas tiendas activas usa cada plataforma. Si son dos y comparten
  // una sola pestaña, importar las dos etiquetaría las mismas filas
  // primero con una tienda y luego con la otra: los datos quedarían
  // asignados a la que corrió de último, sin que nadie lo note.
  const porFuente = {};
  datos.slice(1).forEach(function (f) {
    if (norm(f[cA]) !== 'si') return;
    const fu = String(f[cF]).trim();
    (porFuente[fu] = porFuente[fu] || []).push(String(f[cT]).trim());
  });

  const log = [];
  datos.slice(1).forEach(function (f) {
    if (norm(f[cA]) !== 'si') return;
    const fuente = String(f[cF]).trim(), tienda = String(f[cT]).trim();

    if (porFuente[fuente].length > 1) {
      const base = '_Import_' + fuente.charAt(0).toUpperCase() + fuente.slice(1);
      if (!ss.getSheetByName(base + '_' + tienda.toUpperCase())) {
        log.push('SALTADA: ' + fuente + ' / ' + tienda + '\n' +
          '  La usan ' + porFuente[fuente].length + ' tiendas (' +
          porFuente[fuente].join(', ') + ') y solo hay una pestaña ' + base + '.\n' +
          '  Importarlas juntas etiquetaría las mismas filas con la tienda\n' +
          '  equivocada. Elige una salida:\n' +
          '    a) Crea una pestaña por tienda: ' + base + '_GT y ' + base + '_EC\n' +
          '    b) O pon "no" en la columna activa de la tienda que no la usa');
        return;
      }
    }

    // Effi son dos reportes con nombres propios
    const ids = fuente === 'effi' ? ['effi_guias', 'effi_novedades'] : [fuente];
    ids.forEach(function (id) {
      try { log.push(importar(id, tienda, cliente)); }
      catch (err) { log.push('· ' + id + '/' + tienda + ': ' + err.message); }
    });
  });

  try { log.push(cruzarLlamadas(cliente)); } catch (err) { /* sin llamadas */ }

  const msg = log.join('\n\n');
  Logger.log(msg);
  return msg;
}


// ─── ATAJOS ──────────────────────────────────────────────────
/**
 * El botón Ejecutar de Apps Script no permite pasar argumentos, así que
 * para importar una sola fuente hace falta una función sin parámetros.
 *
 * No se listan por tienda a propósito: hacerlo ataría el código a las
 * tiendas de una cuenta, y Nova se vende a clientes de toda la región.
 * importarTodo() recorre la hoja Fuentes, que es donde vive esa lista.
 *
 * Si necesitas importar una fuente suelta, escribe la llamada en la
 * consola del editor:  importar('dropi', 'lima')
 */
function importarTodoAhora() { return importarTodo(); }
