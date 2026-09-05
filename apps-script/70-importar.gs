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
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  if (!tienda) throw new Error('Falta la tienda. Ej: importar("dropi","ec")');
  if (!monedaDeTienda(ss, tienda)) {
    throw new Error('La tienda "' + tienda + '" no está en la hoja Tiendas.');
  }

  const r = leerCrudo(ss, fuenteId, tienda);
  if (!r.filas.length) {
    const msg = r.tab + ' está vacía. Pega el export ahí primero.';
    Logger.log(msg);
    return msg;
  }

  const destino = { pedidos: 'Pedidos', novedades: 'Novedades',
                    llamadas: 'Llamadas', pauta: 'Pauta',
                    pedidos_secundario: 'Pedidos' }[r.tipo];
  if (!destino) throw new Error('No sé dónde guardar una fuente de tipo ' + r.tipo);

  const pais = paisDeTienda(ss, tienda);
  const preparadas = r.filas.map(function (f) {
    return prepararFila(f, r.tipo, fuenteId, tienda, pais, ss);
  });

  const res = escribirFilas(ss, destino, preparadas, fuenteId);
  registrarImportacion(ss, tienda, fuenteId, res.nuevas + res.actualizadas);

  // Las novedades que vienen dentro del export de pedidos se derivan aparte
  let extra = '';
  if (r.tipo === 'pedidos' && fuenteId === 'dropi') {
    const nov = derivarNovedades(preparadas, fuenteId, tienda);
    if (nov.length) {
      const rn = escribirFilas(ss, 'Novedades', nov, fuenteId);
      extra = '\nNovedades derivadas: ' + rn.nuevas + ' nuevas, ' +
              rn.actualizadas + ' actualizadas';
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
      ? '  ⚠ estados sin mapear: ' + res.sinEstado.slice(0, 8).join(' · ')
      : '',
    extra,
  ].filter(String).join('\n');
  Logger.log(msg);
  return msg;
}

/** Convierte una fila normalizada en una fila lista para la entidad. */
function prepararFila(f, tipo, fuenteId, tienda, pais, ss) {
  const o = Object.assign({}, f);
  o.fuente = fuenteId;
  o.tienda = tienda;

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
    o.estado_canonico = estadoCanonico(fuenteId, o.estado);
  }
  if (tipo === 'novedades') {
    o.grupo = grupoNovedad(o.motivo, o.codigo);
    if (o.aclaracion && !o.nota) { o.nota = o.aclaracion; }
  }
  if (tipo === 'llamadas') {
    o.seg_conversado = aSegundos(o.seg_conversado);
    o.seg_espera     = aSegundos(o.seg_espera);
    o.seg_total      = aSegundos(o.seg_total);
    // IRIS no dice de qué tienda es la llamada: se cruza por teléfono
    o.tienda = '';
  }
  if (tipo === 'pauta') {
    o.plataforma = fuenteId;
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
function escribirFilas(ss, hoja, filas, fuenteId) {
  const sh = ss.getSheetByName(hoja);
  if (!sh) throw new Error('No existe la hoja ' + hoja);

  const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
  const cId = enc.indexOf('id');
  if (cId === -1) throw new Error(hoja + ' no tiene columna id.');

  // Índice de lo que ya está, para no recorrer la hoja por cada fila
  const existentes = {};
  if (sh.getLastRow() > 1) {
    const datos = sh.getRange(2, 1, sh.getLastRow() - 1, enc.length).getValues();
    datos.forEach(function (f, i) {
      const k = String(f[cId]).trim();
      if (k) existentes[k] = { fila: i + 2, valores: f };
    });
  }

  const nuevas = [], sinEstado = {};
  let actualizadas = 0, iguales = 0;

  filas.forEach(function (o) {
    if (o.estado_canonico && o.estado_canonico.indexOf('__sin_mapear__') === 0) {
      sinEstado[o.estado_canonico.replace('__sin_mapear__:', '')] = 1;
    }

    const prev = existentes[o.id];
    if (!prev) {
      nuevas.push(enc.map(function (col) {
        return o[col] !== undefined ? o[col] : '';
      }));
      return;
    }

    // Ya existe: solo se tocan las celdas que cambiaron, y jamás las del equipo
    let cambio = false;
    enc.forEach(function (col, i) {
      if (COLUMNAS_DEL_EQUIPO.indexOf(col) !== -1) return;
      if (o[col] === undefined || o[col] === '') return;
      if (String(prev.valores[i]) === String(o[col])) return;
      sh.getRange(prev.fila, i + 1).setValue(o[col]);
      cambio = true;
    });
    if (cambio) actualizadas++; else iguales++;
  });

  if (nuevas.length) {
    sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, enc.length).setValues(nuevas);
  }

  return {
    nuevas: nuevas.length, actualizadas: actualizadas, iguales: iguales,
    sinEstado: Object.keys(sinEstado),
  };
}

/**
 * Dropi trae la novedad dentro de la misma fila del pedido, no en un
 * reporte aparte como Effi. Se extraen para que la alarma de patrón y
 * el cierre de mes puedan contarlas.
 */
function derivarNovedades(pedidos, fuenteId, tienda) {
  return pedidos
    .filter(function (p) {
      const m = String(p.motivo_novedad || '').trim();
      return m && m !== '.' && m !== '-';
    })
    .map(function (p) {
      return {
        id: fuenteId + '-nov-' + p.id_externo,
        fuente: fuenteId,
        id_externo: p.id_externo,
        pedido_id: p.id,
        fecha: p.ultimo_movimiento || p.fecha,
        tipo: 'novedad',
        motivo: p.motivo_novedad,
        grupo: grupoNovedad(p.motivo_novedad),
        estado: p.estado_canonico === 'novedad_resuelta' ? 'resuelta' : 'abierta',
        solucion: p.solucion || '',
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
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
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
  const ss = SpreadsheetApp.openById(hojaCliente(cliente));
  const sh = ss.getSheetByName('Fuentes');
  if (!sh || sh.getLastRow() < 2) return 'No hay fuentes configuradas.';

  const datos = sh.getDataRange().getValues();
  const enc = datos[0].map(norm);
  const cT = enc.indexOf('tienda'), cF = enc.indexOf('fuente'), cA = enc.indexOf('activa');

  const log = [];
  datos.slice(1).forEach(function (f) {
    if (norm(f[cA]) !== 'si') return;
    const fuente = String(f[cF]).trim(), tienda = String(f[cT]).trim();
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
