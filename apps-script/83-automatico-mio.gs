/**
 * ═══════════════════════════════════════════════════════════
 *  LO AUTOMÁTICO DEL LADO DE ELLA
 *  Nova Central y NovaSoul, sin que nadie abra nada
 * ═══════════════════════════════════════════════════════════
 *
 * Lo que ya corría solo era del lado del CLIENTE: las tasas, la lectura
 * de Meta, las alarmas de la tienda, el semáforo del lunes. Todo eso
 * mira las tiendas de Nutrea y le escribe a la dueña.
 *
 * De su lado —sus proyectos, sus cobros, su semana, su cielo— no corría
 * nada. Nova lo sabía todo y no decía nada hasta que ella abría la
 * pantalla, que es justo el día que no la abre.
 *
 * Aquí van los dos que faltaban:
 *
 *   soulLunes()     el lunes a las 9 · ¿cabe la semana? qué se puede
 *                   mover, qué turno quedó sin cerrar, qué momento es
 *   centralDiario() todos los días a las 9 · qué cobro está tarde, qué
 *                   entrega se vence, qué proyecto va con deuda
 *
 * ── DOS REGLAS QUE NO SE NEGOCIAN ──
 *
 * 1· CORREO VACÍO NO SE MANDA. Es la misma regla de `revisarAlarmas`.
 *    Un aviso que llega todos los días diciendo «todo bien» deja de
 *    leerse en dos semanas, y entonces tampoco se lee el día que sí
 *    traía algo. Si no hay nada que hacer, Nova se calla.
 *
 * 2· LO DE PHH NO SALE DE NOVASOUL. El correo del lunes puede decir
 *    cuántas horas ocupa un proyecto confidencial y que hay algo
 *    abierto; nunca QUÉ es. Es la misma pared que ya está en
 *    `soulCargaPorTrabajo_`, y aquí se repite porque un correo se
 *    reenvía, se imprime y queda en un buzón que no es solo suyo.
 */

/** Las socias activas de la plataforma. Son las dueñas de estos avisos. */
function sociasPlataforma_() {
  const ss = SpreadsheetApp.openById(IDS_().central);
  const sh = ss.getSheetByName('Plataforma');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const out = [];
  for (let i = 1; i < d.length; i++) {
    if (norm(d[i][c('rol')]) === 'operadora') continue;
    if (norm(d[i][c('estado')]) === 'inactivo') continue;
    const correo = String(d[i][c('correo')] || '').toLowerCase().trim();
    if (!correo) continue;
    out.push({ correo: correo, nombre: String(d[i][c('nombre')] || ''), rol: 'socia' });
  }
  return out;
}

/** Una fecha como se dice, no como se guarda. */
const AUTO_MESES = ['ene', 'feb', 'mar', 'abr', 'may', 'jun',
                    'jul', 'ago', 'sep', 'oct', 'nov', 'dic'];
function autoFecha_(iso) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ''))) return String(iso || '');
  return Number(iso.slice(8)) + ' ' + AUTO_MESES[Number(iso.slice(5, 7)) - 1];
}

/** Montos por moneda, sin sumarlas entre sí. */
function autoMonedas_(m) {
  const k = Object.keys(m || {});
  if (!k.length) return '';
  return k.map(function (x) { return redondear_(m[x]) + ' ' + x; }).join(' · ');
}

// ─── EL LUNES DE ELLA ────────────────────────────────────────

/**
 * ¿Cabe la semana?
 *
 * Es la única pregunta del correo. Todo lo demás está para responderla:
 * las horas que tiene, las que ya están comprometidas, lo que se puede
 * mover si no alcanza, y lo que NO se puede mover aunque quiera.
 *
 * Se manda el lunes y no el domingo a propósito: el domingo la respuesta
 * no se puede usar, y saber el domingo que la semana no cabe es una
 * mala noche sin nada a cambio.
 */
function soulLunesTexto(h, cielo, nuevasTemporadas) {
  const L = [];
  const r = h.riesgo || {};
  L.push('TU SEMANA · ' + autoFecha_(h.semana.lunes) + ' a ' + autoFecha_(h.semana.domingo));
  L.push('');

  // ── ¿Cabe? ──
  if (r.sinHoras) {
    L.push('NO SÉ SI CABE.');
    L.push(r.porque);
  } else {
    L.push('Tienes ' + redondear_(r.libres) + ' h libres esta semana.');
    L.push('Comprometidas: ' + redondear_(r.comprometidas) + ' h' +
           (r.fijas ? ' (' + redondear_(r.fijas) + ' h son horas fijas de proyectos)' : '') + '.');
    if (r.ocupadas) {
      L.push('Turnos y clases ya están descontados: ' + redondear_(r.ocupadas) + ' h.');
    }
    L.push('');
    if (r.sobra >= 0) {
      L.push('✓ CABE. Te sobran ' + redondear_(r.sobra) + ' h.');
    } else {
      L.push('✕ NO CABE. Faltan ' + redondear_(Math.abs(r.sobra)) + ' h.');
      if (r.candidatas && r.candidatas.length) {
        L.push('');
        L.push('Lo que se puede mover sin romper nada:');
        r.candidatas.slice(0, 5).forEach(function (t) {
          L.push('  · ' + t.texto + '  (' + redondear_(t.horas) + ' h' +
                 (t.fecha ? ', para el ' + autoFecha_(t.fecha) : '') + ')');
        });
      }
      if (r.noAlcanza > 0) L.push('', r.porque);
    }
  }

  // ── Lo que ya está tarde ──
  const vencidas = (h.pendientes || []).filter(function (t) {
    return t.estado !== 'hecho' && t.dias !== null && t.dias > 0;
  });
  if (vencidas.length) {
    L.push('');
    L.push('YA VENCIDAS (' + vencidas.length + ')');
    vencidas.slice(0, 8).forEach(function (t) {
      L.push('  · ' + t.texto + ' — ' + t.dias + (t.dias === 1 ? ' día' : ' días') + ' tarde');
    });
    if (vencidas.length > 8) L.push('  … y ' + (vencidas.length - 8) + ' más.');
  }

  // ── Lo de la semana, día por día ──
  const conFecha = (h.semana.dias || []).filter(function (d) { return d.entregas; });
  if (conFecha.length) {
    L.push('');
    L.push('ESTA SEMANA');
    conFecha.forEach(function (d) {
      L.push('  ' + autoFecha_(d.fecha) + ':  ' + d.entregas +
             (d.entregas === 1 ? ' entrega' : ' entregas') +
             (d.horas ? ', ' + redondear_(d.horas) + ' h' : ''));
    });
  }

  // ── Turnos sin cerrar ──
  if (h.turnosPendientes) {
    L.push('');
    L.push('PROPINAS SIN ANOTAR: ' + h.turnosPendientes +
           (h.turnosPendientes === 1 ? ' turno' : ' turnos') + '.');
    L.push('Sin eso, el cierre del mes va corto y no por culpa del mes.');
  }

  // ── Mindlab ──
  const ml = (h.mindlab || {}).semanaActual;
  if (ml) {
    L.push('');
    L.push('MINDLAB · semana ' + ml.semana + ': ' + ml.tema);
    L.push('  ' + ml.tarea + '  (' + redondear_(ml.horas) + ' h)');
  }

  // ── Qué momento es ──
  if (cielo && cielo.ok) {
    const hoyC = (cielo.semana.dias || []).filter(function (d) { return d.esHoy; })[0];
    const mom = hoyC ? (cielo.momentos[hoyC.momento] || {}) : null;
    const P = cielo.profecciones;
    if (mom || (P && P.hay)) {
      L.push('');
      L.push('QUÉ MOMENTO ES');
      if (mom && mom.nombre) {
        L.push('  Hoy: tiempo de ' + mom.nombre.toLowerCase() +
               ' (' + (hoyC.deDonde === 'pensum' ? 'tu pensum: ' : 'la luna: ') +
               hoyC.porque + ').');
      }
      if (P && P.hay) {
        L.push('  El mes: casa ' + P.mes.casa + ' · ' + P.mes.signoNombre + ' — ' +
               P.mes.area + '. Quedan ' + P.mes.diasRestantes +
               (P.mes.diasRestantes === 1 ? ' día.' : ' días.'));
        L.push('  El año: casa ' + P.anual.casa + ' · ' + P.anual.signoNombre + ' — ' +
               P.anual.area + '.');
      }
    }
  }

  // ── Lo que Nova puso sola en el pensum ──
  if (nuevasTemporadas && nuevasTemporadas.length) {
    L.push('');
    L.push('TEMPORADAS NUEVAS EN TU PENSUM (' + nuevasTemporadas.length + ')');
    L.push('Las puso Nova desde tus tránsitos. Bórralas o cámbialas cuando quieras.');
    nuevasTemporadas.slice(0, 6).forEach(function (x) {
      L.push('  · ' + x.titulo + '  (' + autoFecha_(x.desde) +
             (x.hasta ? ' a ' + autoFecha_(x.hasta) : '') + ')');
      /**
       * Si los dos sistemas de casas no coinciden, se dice. La casa
       * decide el momento, así que esa discrepancia cambia lo que Nova
       * le recomienda hacer con la semana: callarla sería decidir por
       * ella sin avisarle.
       */
      if (x.casas && x.casas.difieren) {
        L.push('      casa ' + x.casas.placidus + ' en Horus, ' +
               x.casas.entera + ' en casas enteras — no coinciden, míralo.');
      }
    });
    if (nuevasTemporadas.length > 6) {
      L.push('  … y ' + (nuevasTemporadas.length - 6) + ' más, en El cielo.');
    }
  }

  // ── Lo que Nova no pudo leer ──
  if (h.faltanHojas && h.faltanHojas.length) {
    L.push('');
    L.push('OJO: no pude leer ' + h.faltanHojas.join(', ') + '.');
    L.push('Corre bootstrapTodo() una vez y esto se arregla solo.');
  }

  L.push('');
  L.push('— NovaSoul');
  return L.join('\n');
}

/** ¿Hay algo que decir? Si no, el lunes no llega ningún correo. */
function soulLunesHayAlgo_(h) {
  const r = h.riesgo || {};
  if (r.sinHoras) return true;
  if (r.sobra < 0) return true;
  if (h.turnosPendientes) return true;
  if ((h.faltanHojas || []).length) return true;
  if ((h.resumen.semana || {}).entregas) return true;
  if ((h.pendientes || []).some(function (t) {
    return t.estado !== 'hecho' && t.dias !== null && t.dias > 0;
  })) return true;
  return false;
}

function soulLunes() {
  const log = [];
  sociasPlataforma_().forEach(function (s) {
    try {
      const h = soulHoy(s, {});
      if (!h || !h.ok) { log.push(s.correo + ': ' + ((h && h.error) || 'sin datos')); return; }
      if (!soulLunesHayAlgo_(h)) {
        log.push(s.correo + ': semana limpia, no se manda nada.');
        return;
      }
      /**
       * Antes del correo, el pensum del lunes.
       *
       * Ella pidió que Nova cree el pensum sola. El lunes es el día:
       * las temporadas que abren esta semana quedan puestas antes de
       * que el correo le diga qué momento es, y no al revés.
       *
       * Si falla, el correo sale igual — una temporada que no se pudo
       * crear no puede dejarla sin saber si la semana cabe.
       */
      let nuevasTemporadas = [];
      try { nuevasTemporadas = pensumAuto_(soulUsuario_(s), h.hoy); }
      catch (e) { nuevasTemporadas = []; }

      // El cielo es de adorno aquí: si falla, el correo sale igual.
      let cielo = null;
      try { cielo = soulCielo(s, {}); } catch (e) { cielo = null; }
      MailApp.sendEmail({
        to: s.correo,
        subject: 'Tu semana · ' + h.semana.lunes,
        body: soulLunesTexto(h, cielo, nuevasTemporadas),
      });
      log.push(s.correo + ' → enviado.');
    } catch (e) {
      log.push(s.correo + ': FALLÓ — ' + e.message);
    }
  });
  const msg = log.length ? log.join('\n') : 'No hay ninguna socia en la hoja Plataforma.';
  Logger.log(msg);
  return msg;
}

// ─── EL DÍA A DÍA DE CENTRAL ─────────────────────────────────

/** Cuántos días faltan (o sobran) para una fecha. */
function autoDias_(iso, hoy) {
  if (!/^\d{4}-\d{2}-\d{2}$/.test(String(iso || ''))) return null;
  return Math.round((new Date(iso + 'T00:00:00Z') -
                     new Date(hoy + 'T00:00:00Z')) / 86400000);
}

const CENTRAL_AVISO_DIAS = 3;

/**
 * Lo que hay que hacer hoy del lado del negocio.
 *
 * Tres cosas y nada más: un cobro que ya está tarde, un cobro que vence
 * esta semana, y una entrega de proyecto que se vence. Son las tres que
 * cuestan plata o cuestan un cliente si se pasan de largo.
 *
 * Un proyecto marcado confidencial entra por su nombre y su fecha —eso
 * es de Central, no de PHH— pero NUNCA con sus tareas.
 */
function centralDiarioArmar(m) {
  const hoy = m.hoy;
  const atrasados = m.atrasados || [];
  const porVencer = (m.porCobrar || []).filter(function (c) {
    const d = autoDias_(c.esperada, hoy);
    return d !== null && d >= 0 && d <= CENTRAL_AVISO_DIAS;
  });
  const entregas = (m.trabajos || []).filter(function (t) {
    if (t.estado !== 'activo') return false;
    const d = autoDias_(t.entrega, hoy);
    return d !== null && d <= CENTRAL_AVISO_DIAS;
  }).sort(function (a, b) { return a.entrega < b.entrega ? -1 : 1; });

  return { hoy: hoy, atrasados: atrasados, porVencer: porVencer, entregas: entregas,
           hayAlgo: !!(atrasados.length || porVencer.length || entregas.length),
           totales: m.totales || {} };
}

function centralDiarioTexto(a) {
  const L = [];
  L.push('NOVA CENTRAL · ' + autoFecha_(a.hoy));
  L.push('');

  if (a.atrasados.length) {
    L.push('COBROS TARDE (' + a.atrasados.length + ')');
    a.atrasados.slice(0, 10).forEach(function (c) {
      L.push('  · ' + (c.trabajo || 'Sin proyecto') + ' — ' + c.concepto + ': ' +
             redondear_(c.monto) + ' ' + c.moneda +
             '   ' + c.dias + (c.dias === 1 ? ' día' : ' días') + ' tarde');
    });
    const t = autoMonedas_(a.totales.atrasado);
    if (t) L.push('  Total atrasado: ' + t);
    L.push('');
  }

  if (a.porVencer.length) {
    L.push('COBROS QUE VENCEN EN ' + CENTRAL_AVISO_DIAS + ' DÍAS O MENOS');
    a.porVencer.forEach(function (c) {
      L.push('  · ' + (c.trabajo || 'Sin proyecto') + ' — ' + c.concepto + ': ' +
             redondear_(c.monto) + ' ' + c.moneda + '   el ' + autoFecha_(c.esperada));
    });
    L.push('');
  }

  if (a.entregas.length) {
    L.push('ENTREGAS');
    a.entregas.forEach(function (t) {
      const d = autoDias_(t.entrega, a.hoy);
      const cuando = d < 0 ? Math.abs(d) + (d === -1 ? ' día tarde' : ' días tarde')
                   : d === 0 ? 'HOY'
                   : 'en ' + d + (d === 1 ? ' día' : ' días');
      L.push('  · ' + t.nombre + ' (' + autoFecha_(t.entrega) + ') — ' + cuando);
      /**
       * Aquí está la pared. De un proyecto confidencial se dice que
       * tiene algo abierto y cuántas horas; jamás qué dice la tarea.
       * Un correo se reenvía.
       */
      const c = t.tareas || {};
      if (c.abiertas) {
        L.push('      ' + c.abiertas + (c.abiertas === 1 ? ' tarea abierta' : ' tareas abiertas') +
               (c.horas ? ', ' + redondear_(c.horas) + ' h' : '') +
               (c.vencidas ? ' · ' + c.vencidas + ' vencida' + (c.vencidas === 1 ? '' : 's') : '') +
               (t.confidencial ? '  (confidencial: el detalle solo en NovaSoul)' : ''));
      }
    });
    L.push('');
  }

  L.push('— Nova Central');
  return L.join('\n');
}

function centralDiario() {
  const log = [];
  sociasPlataforma_().forEach(function (s) {
    try {
      const m = centralMio(s, {});
      if (!m || !m.ok) { log.push(s.correo + ': ' + ((m && m.error) || 'sin datos')); return; }
      const a = centralDiarioArmar(m);
      if (!a.hayAlgo) { log.push(s.correo + ': nada que avisar hoy.'); return; }
      MailApp.sendEmail({
        to: s.correo,
        subject: 'Hoy en Nova Central · ' +
                 (a.atrasados.length ? a.atrasados.length + ' cobro(s) tarde'
                                     : a.entregas.length + ' entrega(s)'),
        body: centralDiarioTexto(a),
      });
      log.push(s.correo + ' → enviado.');
    } catch (e) {
      log.push(s.correo + ': FALLÓ — ' + e.message);
    }
  });
  const msg = log.length ? log.join('\n') : 'No hay ninguna socia en la hoja Plataforma.';
  Logger.log(msg);
  return msg;
}
