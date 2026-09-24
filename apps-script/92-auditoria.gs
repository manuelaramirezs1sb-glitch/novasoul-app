/**
 * ═══════════════════════════════════════════════════════════════
 *  AUDITORÍA · casos reales, veredicto que se guarda
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE ELLA ENCONTRÓ ─────────────────────────────────────┐
 * │                                                            │
 * │ «auditorías tampoco está guardando los hallazgos».         │
 * │                                                            │
 * │ No era un fallo de guardado. La pantalla ENTERA era una    │
 * │ maqueta: tres clientas inventadas —Andrea Morales, Luisa   │
 * │ Fernández, Carolina Jiménez—, dos gestoras inventadas,     │
 * │ transcripciones de llamadas inventadas, y un botón de      │
 * │ «Guardar veredicto» que no tenía ni un `onclick`. Se podía │
 * │ marcar OK o Hallazgo, escribir la nota, darle a guardar, y │
 * │ no pasaba absolutamente nada — sin un solo mensaje de      │
 * │ error, porque no había nada que fallara.                   │
 * │                                                            │
 * │ Un botón que no hace nada es peor que un botón que falla.  │
 * │ El que falla se arregla; el que no hace nada se usa        │
 * │ durante semanas creyendo que sí.                           │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ QUÉ ES AUDITAR AQUÍ, DE VERDAD ───────────────────────────┐
 * │                                                            │
 * │ Auditar un caso es contrastar DOS relatos del mismo hecho: │
 * │                                                            │
 * │   lo que escribió la gestora  ←→  lo que registró IRIS     │
 * │   (la solución de la novedad)     (la central telefónica)  │
 * │                                                            │
 * │ Cuando una gestora escribe «llamé tres veces» y la central │
 * │ registra dos llamadas, eso es un hallazgo. Cuando escribe  │
 * │ «entregado, la clienta confirmó» y no hay ni una llamada,  │
 * │ eso es otro. Y cuando calzan, se cierra el caso y ya.      │
 * │                                                            │
 * │ Los dos lados son datos que Nova YA tiene: Novedades trae  │
 * │ `gestora` y `solucion`; Llamadas trae quién llamó, cuándo  │
 * │ y cuánto habló. No había que inventar nada — había que     │
 * │ cruzarlo.                                                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ NOVA NO PONE EL VEREDICTO ────────────────────────┐
 * │                                                            │
 * │ Nova marca las DISCREPANCIAS —dijo tres llamadas, hay dos— │
 * │ y ahí se detiene. El veredicto lo pone una persona.        │
 * │                                                            │
 * │ Porque una discrepancia no es una falta: la clienta pudo   │
 * │ haber llamado ella, pudo ser por WhatsApp, la central pudo │
 * │ no haber registrado. Un sistema que reparte «hallazgos»    │
 * │ solo, sobre el trabajo de una persona concreta, se         │
 * │ equivoca en público y con nombre propio.                   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Los tres veredictos. No hay más, y no se aceptan inventados. */
const AUD_VEREDICTOS = {
  ok:        { nombre: 'OK', que: 'Lo que escribió calza con lo que hay registrado.' },
  hallazgo:  { nombre: 'Hallazgo', que: 'Hay algo que no calza y hay que hablarlo.' },
  pendiente: { nombre: 'Pendiente', que: 'Todavía no se revisó.' },
};

/** Cuántos días atrás se revisan por defecto. */
const AUD_VENTANA_DIAS = 14;

/**
 * Las señales que Nova sí puede levantar sola, sin juzgar.
 *
 * Cada una dice qué se comprobó y con qué números, para que quien
 * ponga el veredicto pueda estar en desacuerdo con el dato a la vista.
 * Una señal sin sus cifras es una acusación.
 */
function audSenales_(nov, llamadas) {
  const out = [];
  const texto = String(nov.solucion || '') + ' ' + String(nov.nota || '');
  const t = norm(texto);

  const hechas = llamadas.filter(function (l) {
    return norm(l.sentido) !== 'entrante';
  });
  const contestadas = llamadas.filter(function (l) {
    return num(l.seg_conversado) > 0;
  });

  /**
   * «Llamé N veces» contra las llamadas registradas.
   *
   * Se buscan números escritos con cifra y con letra, porque una
   * gestora escribe «llamé 3 veces» y también «llamé tres veces», y
   * una comprobación que solo entiende cifras deja pasar la mitad.
   */
  const letras = { un: 1, una: 1, dos: 2, tres: 3, cuatro: 4, cinco: 5, seis: 6 };
  let dichas = null;
  const m = t.match(/(\d+|un|una|dos|tres|cuatro|cinco|seis)\s+(veces|vez|intentos|intento|llamadas|llamada)/);
  if (m) dichas = /^\d+$/.test(m[1]) ? Number(m[1]) : letras[m[1]];

  if (dichas !== null && dichas > llamadas.length) {
    out.push({
      tipo: 'llamadas_de_menos', nivel: 'mal',
      que: 'Escribió ' + dichas + (dichas === 1 ? ' intento' : ' intentos') +
           ' y en la central hay ' + llamadas.length + '.',
      dato: { dijo: dichas, hay: llamadas.length },
    });
  }

  if (/entregad|confirm|acord|reprogram/.test(t) && llamadas.length === 0) {
    out.push({
      tipo: 'sin_llamada', nivel: 'ojo',
      que: 'Dice que hubo acuerdo con la clienta y no hay ninguna llamada registrada ' +
           'para este pedido. Puede haber sido por WhatsApp.',
      dato: { hay: 0 },
    });
  }

  if (llamadas.length > 0 && contestadas.length === 0 && /confirm|acord/.test(t)) {
    out.push({
      tipo: 'nadie_contesto', nivel: 'ojo',
      que: 'Hay ' + llamadas.length + ' llamada(s), pero ninguna con conversación. ' +
           'Dice que la clienta confirmó.',
      dato: { llamadas: llamadas.length, contestadas: 0 },
    });
  }

  if (!String(nov.solucion || '').trim() && norm(nov.estado) !== 'abierta') {
    out.push({
      tipo: 'sin_nota', nivel: 'mal',
      que: 'La novedad se cerró sin escribir qué se hizo.',
      dato: {},
    });
  }

  if (!out.length && llamadas.length && hechas.length) {
    out.push({
      tipo: 'calza', nivel: 'bien',
      que: 'Lo escrito y lo registrado calzan: ' + llamadas.length +
           ' llamada(s), ' + contestadas.length + ' con conversación.',
      dato: { llamadas: llamadas.length, contestadas: contestadas.length },
    });
  }

  return out;
}

/**
 * Los casos a revisar, con los dos relatos al lado.
 *
 * Solo novedades CERRADAS: una abierta todavía se está trabajando, y
 * auditar el trabajo de alguien mientras lo está haciendo no dice nada
 * salvo que va por la mitad.
 */
function apiAuditoriaCasos(s, p) {
  if (s.rol === 'gestora') return { ok: false, error: 'No tienes acceso a la auditoría.' };
  const tienda = String(p.tienda || '').trim();
  if (!tienda) return { ok: false, error: 'Falta decir de qué tienda.' };

  const ss = libro_(s.sheetId);
  const hoy = ahoraISO().slice(0, 10);
  const desde = masDias_(hoy, -(num(p.dias) || AUD_VENTANA_DIAS));

  const leer = function (nombre) {
    const sh = ss.getSheetByName(nombre);
    if (!sh || sh.getLastRow() < 2) return { enc: [], filas: [] };
    const d = sh.getDataRange().getValues();
    return { enc: d[0].map(norm), filas: d.slice(1) };
  };

  const nov = leer('Novedades');
  if (!nov.filas.length) {
    return { ok: true, casos: [], veredictos: AUD_VEREDICTOS, desde: desde, hoy: hoy,
             porque: 'Todavía no hay novedades importadas. La auditoría revisa novedades ' +
                     'cerradas, así que aparece en cuanto tu equipo cierre la primera.' };
  }

  const ped = leer('Pedidos');
  const lla = leer('Llamadas');
  const obj = function (enc, f) {
    const o = {};
    enc.forEach(function (c, i) { o[c] = f[i]; });
    return o;
  };

  // Los pedidos por id, para sacar cliente, teléfono y tienda.
  const porPedido = {};
  ped.filas.forEach(function (f) {
    const o = obj(ped.enc, f);
    if (o.id) porPedido[String(o.id)] = o;
  });

  /**
   * Las llamadas, indexadas por pedido Y por teléfono.
   *
   * IRIS no siempre trae `pedido_id` —el comentario del esquema ya lo
   * decía: «cruza con Pedidos por telefono_norm, no por id de orden»—.
   * Así que se indexa por los dos y se usa el que haya. Indexar solo
   * por pedido dejaría la mitad de las llamadas sin encontrar, y el
   * resultado sería una pantalla llena de «no hay llamadas» falsos.
   */
  const llaPorPedido = {}, llaPorTel = {};
  lla.filas.forEach(function (f) {
    const o = obj(lla.enc, f);
    const pid = String(o.pedido_id || '').trim();
    const tel = String(o.telefono_norm || '').trim();
    if (pid) (llaPorPedido[pid] = llaPorPedido[pid] || []).push(o);
    if (tel) (llaPorTel[tel] = llaPorTel[tel] || []).push(o);
  });

  // Lo ya auditado, para no volver a pedir el mismo veredicto.
  const yaHecho = {};
  const aud = leer('Auditorias');
  aud.filas.forEach(function (f) {
    const o = obj(aud.enc, f);
    if (o.novedad_id) yaHecho[String(o.novedad_id)] = o;
  });

  const casos = [];
  const tope = Math.min(120, Math.max(1, num(p.limite) || 40));
  let cerradas = 0;

  for (let i = nov.filas.length - 1; i >= 0 && casos.length < tope; i--) {
    const n = obj(nov.enc, nov.filas[i]);
    const abierta = norm(n.estado) === 'abierta' ||
                    (!norm(n.solucionada) && !String(n.fecha_solucion || '').trim());
    if (abierta) continue;

    const fecha = aISO(n.fecha_solucion, 'UTC') || aISO(n.fecha, 'UTC') || '';
    if (fecha && fecha < desde) continue;

    const pedido = porPedido[String(n.pedido_id || '')] || {};
    if (tienda && pedido.tienda && norm(pedido.tienda) !== norm(tienda)) continue;
    cerradas++;

    const tel = String(pedido.telefono_norm || '').trim();
    const suyas = (llaPorPedido[String(n.pedido_id || '')] ||
                   (tel ? llaPorTel[tel] : []) || []).slice();
    suyas.sort(function (a, b) {
      return String(a.fecha_hora) < String(b.fecha_hora) ? -1 : 1;
    });

    const guardado = yaHecho[String(n.id)] || null;

    casos.push({
      novedadId: String(n.id || ''),
      pedidoId: String(n.pedido_id || ''),
      fecha: fecha,
      cliente: String(pedido.cliente || ''),
      ciudad: String(pedido.ciudad || ''),
      motivo: String(n.motivo || ''),
      desenlace: String(n.desenlace || ''),
      gestora: String(n.gestora || '').trim(),
      /** Lo que escribió la persona. Tal cual, sin recortar. */
      dijoGestora: String(n.solucion || n.nota || '').trim(),
      /** Lo que registró la central. */
      llamadas: suyas.map(function (l) {
        return {
          cuando: String(l.fecha_hora || ''),
          agente: String(l.agente || ''),
          sentido: String(l.sentido || ''),
          estado: String(l.estado || ''),
          segundos: num(l.seg_conversado),
          observacion: String(l.observacion || ''),
          grabacion: String(l.grabacion || ''),
        };
      }),
      senales: audSenales_(n, suyas),
      veredicto: guardado ? norm(guardado.veredicto) : 'pendiente',
      notaAuditoria: guardado ? String(guardado.nota || '') : '',
      auditadoPor: guardado ? String(guardado.auditor || '') : '',
      auditadoEn: guardado ? String(guardado.creada_en || '') : '',
    });
  }

  // El recuento honesto: cuántas hay, cuántas se revisaron.
  const conVeredicto = casos.filter(function (c) { return c.veredicto !== 'pendiente'; });
  const hallazgos = casos.filter(function (c) { return c.veredicto === 'hallazgo'; });
  const conSenal = casos.filter(function (c) {
    return c.senales.filter(function (x) { return x.nivel !== 'bien'; }).length > 0;
  });

  return {
    ok: true, hoy: hoy, desde: desde, tienda: tienda,
    casos: casos,
    veredictos: AUD_VEREDICTOS,
    recuento: {
      cerradasEnVentana: cerradas,
      mostrados: casos.length,
      revisados: conVeredicto.length,
      sinRevisar: casos.length - conVeredicto.length,
      hallazgos: hallazgos.length,
      /**
       * Cuántos tienen una señal automática. NO es «cuántos están mal»:
       * es cuántos vale la pena mirar primero.
       */
      conSenal: conSenal.length,
    },
    /**
     * Si no hay ni una llamada importada, la mitad del cruce no existe
     * y hay que decirlo: si no, la pantalla mostraría «sin llamada» en
     * todos los casos y parecería que el equipo nunca llama.
     */
    sinLlamadas: lla.filas.length === 0,
    porqueSinLlamadas: lla.filas.length === 0
      ? 'No hay llamadas importadas de IRIS. Sin ellas solo se ve un lado del caso: ' +
        'lo que escribió la gestora. El cruce aparece en cuanto importes la central.'
      : '',
  };
}

/**
 * Guardar el veredicto. Esto es lo que el botón no hacía.
 *
 * La nota es OBLIGATORIA cuando hay hallazgo, y no por burocracia: un
 * hallazgo sin nota es una marca negra en el registro de una persona
 * sin nada que explique qué pasó. Dentro de un mes nadie se acuerda, y
 * lo único que queda es la marca.
 *
 * Para un OK la nota es opcional: cerrar cien casos correctos no puede
 * costar cien párrafos o nadie audita nada.
 */
function apiAuditoriaGuardar(s, p) {
  if (s.rol === 'gestora') return { ok: false, error: 'No tienes acceso a la auditoría.' };
  if (s.rol !== 'dueno' && s.rol !== 'admin' && s.rol !== 'supervisora') {
    return { ok: false, error: 'Solo quien supervisa puede poner un veredicto.' };
  }

  const novedadId = String(p.novedadId || p.novedad_id || '').trim();
  if (!novedadId) return { ok: false, error: 'No sé a qué caso le estás poniendo veredicto.' };

  const veredicto = norm(p.veredicto);
  if (!AUD_VEREDICTOS[veredicto]) {
    return { ok: false, error: 'El veredicto tiene que ser OK, Hallazgo o Pendiente.' };
  }
  const nota = String(p.nota || '').trim();
  if (veredicto === 'hallazgo' && nota.length < 10) {
    return { ok: false, error: 'Un hallazgo necesita una nota que explique qué pasó. ' +
                               'Queda en el registro de esa persona: sin la nota, dentro ' +
                               'de un mes solo queda la marca.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otro veredicto guardándose.' };
  try {
    const ss = libro_(s.sheetId);
    let sh = ss.getSheetByName('Auditorias');
    if (!sh) {
      return { ok: false, error: 'Falta la hoja Auditorias. Corre bootstrapTodo() una vez ' +
                                 'desde el editor de Apps Script y vuelve a intentar.' };
    }
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const ahora = ahoraISO();

    const fila = {
      id: 'au' + Utilities.getUuid().slice(0, 8),
      tienda: String(p.tienda || ''),
      novedad_id: novedadId,
      pedido_id: String(p.pedidoId || p.pedido_id || ''),
      gestora: String(p.gestora || ''),
      veredicto: veredicto,
      nota: nota,
      senales: Array.isArray(p.senales) ? p.senales.join(' · ') : String(p.senales || ''),
      auditor: String(s.email || s.nombre || ''),
      creada_en: ahora,
    };

    /**
     * Un caso tiene UN veredicto, no un historial de veredictos. Si ya
     * había uno, se reemplaza la fila en vez de apilar otra: de lo
     * contrario «revisados: 12» contaría doce veredictos sobre tres
     * casos y el número dejaría de significar nada.
     *
     * El cambio no se pierde: queda en Movimientos, que es el sitio del
     * rastro.
     */
    const d = sh.getLastRow() > 1 ? sh.getDataRange().getValues() : [enc];
    const cN = enc.indexOf('novedad_id');
    let reemplazo = 0;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cN] || '').trim() === novedadId) { reemplazo = i + 1; break; }
    }

    const valores = enc.map(function (c) { return fila[c] !== undefined ? fila[c] : ''; });
    if (reemplazo) {
      const antes = d[reemplazo - 1][enc.indexOf('veredicto')];
      // Se conserva el id original: es la misma auditoría, corregida.
      valores[enc.indexOf('id')] = d[reemplazo - 1][enc.indexOf('id')];
      sh.getRange(reemplazo, 1, 1, enc.length).setValues([valores]);
      registrarMovimiento(s, 'Auditorias', novedadId, 'veredicto',
                          String(antes || ''), veredicto);
    } else {
      sh.appendRow(valores);
      registrarMovimiento(s, 'Auditorias', novedadId, 'veredicto', '', veredicto);
    }

    return { ok: true, veredicto: veredicto, reemplazo: !!reemplazo, cuando: ahora };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}
