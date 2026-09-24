/**
 * ═══════════════════════════════════════════════════════════
 *  QUÉ QUEDÓ DE AYER
 * ═══════════════════════════════════════════════════════════
 *
 * Ella pidió «un reporte diario de lo que quedó pendiente del otro día»
 * y mandó como referencia una tabla de categorías de gestión con su
 * cantidad: chats activos, pendiente de confirmación, corregir
 * dirección, sin respuesta, total insumo.
 *
 * Eso no es un tablero de métricas: es una LISTA DE TRABAJO. Dice
 * cuánta gente hay que tocar hoy y por qué motivo. Por eso este reporte
 * no habla de ventas ni de márgenes — de eso ya hablan el semáforo y el
 * cierre de mes.
 *
 * Y eligió dos cosas como «pendiente»:
 *   · lo que sigue esperando a una persona
 *   · lo que se venció y nadie movió
 *
 * ── LA DIFERENCIA QUE HACE ÚTIL ESTO ──
 *
 * Un pedido sin confirmar de esta mañana y uno de hace tres días se ven
 * igual en cualquier lista. No son lo mismo: el segundo casi siempre
 * termina en devolución. Por eso todo lo que se cuenta aquí se cuenta
 * DOS veces — cuántos hay, y cuántos de esos llevan demasiado tiempo
 * quietos. Sin esa segunda cuenta, la lista crece y nadie nota cuál es
 * la parte que se está pudriendo.
 */

/**
 * Cuánto tiempo es «demasiado» para cada cosa.
 *
 * Una novedad se pudre más rápido que un pedido sin confirmar: el
 * courier tiene un intento diario y una novedad de tres días ya perdió
 * dos. Un pedido en oficina aguanta más, pero no mucho — pasados unos
 * días la clienta ya no va por él.
 */
const DIA_LIMITES = {
  novedad: 1,
  pendiente: 1,
  en_oficina: 3,
  sin_clasificar: 2,
};

/** Las categorías de trabajo, con sus palabras y no con las del sistema. */
const DIA_CATEGORIAS = [
  { id: 'novedad', nombre: 'Con novedad abierta',
    que: 'Hubo un problema y sigue sin resolverse.' },
  { id: 'pendiente', nombre: 'Sin confirmar',
    que: 'Entró el pedido y nadie ha hablado con la clienta.' },
  { id: 'en_oficina', nombre: 'Esperando en oficina',
    que: 'Llegó a la agencia y la clienta todavía no lo recoge.' },
  { id: 'confirmado', nombre: 'Confirmado, sin despachar',
    que: 'Ya se habló con la clienta y falta que salga.' },
  { id: 'sin_clasificar', nombre: 'Estado que Nova no entiende',
    que: 'La transportadora usó una palabra nueva. No cuenta como nada hasta que la traduzcas.' },
];

/**
 * El reporte de un día: qué quedó esperando y qué se pudrió.
 *
 * `dia` es el día que se mira (por defecto, ayer). Lo que se cuenta es
 * el estado de HOY de esos pedidos: un pedido de ayer que ya se entregó
 * esta mañana no es un pendiente, es un trabajo hecho.
 */
function reporteDelDia(sheetId, tienda, diaISO) {
  const ss = SpreadsheetApp.openById(sheetId);
  const hoy = ahoraISO().slice(0, 10);
  const dia = diaISO || masDias_(hoy, -1);

  const shP = ss.getSheetByName('Pedidos');
  const shN = ss.getSheetByName('Novedades');

  const abiertos = [];      // lo que sigue esperando a alguien
  const cerradosAyer = [];  // lo que sí se cerró
  let delDia = 0;

  if (shP && shP.getLastRow() > 1) {
    const d = shP.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (c('tienda') !== -1 && String(f[c('tienda')]).trim() !== tienda) continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      if (!fecha) continue;

      const estado = norm(f[c('estado_nova')] || f[c('estado_canonico')]) ||
                     ESTADOS.SIN_CLASIFICAR;
      const dias = Math.round((new Date(hoy + 'T00:00:00Z') -
                               new Date(fecha + 'T00:00:00Z')) / 86400000);

      if (fecha === dia) delDia++;

      const fila = {
        id: String(f[c('id')] || ''),
        idExterno: String(f[c('id_externo')] || ''),
        cliente: String(f[c('cliente')] || ''),
        telefono: String(f[c('telefono')] || f[c('telefono_norm')] || ''),
        ciudad: String(f[c('ciudad')] || ''),
        producto: String(f[c('producto')] || ''),
        valor: num(f[c('valor')]),
        gestora: String(f[c('gestora_asignada')] || ''),
        fecha: fecha, dias: dias, estado: estado,
      };

      if (ESTADOS_TERMINALES.indexOf(estado) !== -1) {
        if (fecha === dia) cerradosAyer.push(fila);
        continue;
      }
      abiertos.push(fila);
    }
  }

  // ── Por categoría, con la cuenta de los que llevan demasiado ──
  const categorias = DIA_CATEGORIAS.map(function (cat) {
    const suyos = abiertos.filter(function (p) { return p.estado === cat.id; });
    const limite = DIA_LIMITES[cat.id];
    const viejos = limite === undefined ? []
      : suyos.filter(function (p) { return p.dias > limite; });
    return {
      id: cat.id, nombre: cat.nombre, que: cat.que,
      n: suyos.length,
      viejos: viejos.length,
      limite: limite === undefined ? null : limite,
      valor: Math.round(suyos.reduce(function (a, p) { return a + p.valor; }, 0) * 100) / 100,
      // Los más viejos primero: es el orden en que hay que llamarlos.
      casos: suyos.sort(function (a, b) { return b.dias - a.dias; }).slice(0, 25),
    };
  });

  /**
   * En tránsito va aparte y NO como pendiente.
   *
   * Un pedido en camino no espera a nadie: espera al courier. Meterlo
   * en la lista de trabajo la infla con cosas que no se pueden hacer, y
   * una lista así deja de leerse.
   */
  const enCamino = abiertos.filter(function (p) {
    return ['en_transito', 'en_bodega', 'novedad_resuelta'].indexOf(p.estado) !== -1;
  });

  // ── Las novedades, por motivo ──
  const porMotivo = {};
  let novAbiertas = 0, novViejas = 0;
  if (shN && shN.getLastRow() > 1) {
    const d = shN.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const f = d[i];
      if (norm(f[c('estado')]) !== 'abierta') continue;
      const fecha = aISO(f[c('fecha')], 'UTC');
      const dias = fecha
        ? Math.round((new Date(hoy + 'T00:00:00Z') -
                      new Date(fecha + 'T00:00:00Z')) / 86400000) : null;
      const g = norm(f[c('grupo')]) || 'sin_grupo';
      if (!porMotivo[g]) porMotivo[g] = { grupo: g, n: 0, viejas: 0 };
      porMotivo[g].n++;
      novAbiertas++;
      if (dias !== null && dias > DIA_LIMITES.novedad) {
        porMotivo[g].viejas++;
        novViejas++;
      }
    }
  }

  const totalPendiente = categorias.reduce(function (a, c) { return a + c.n; }, 0);
  const totalViejo = categorias.reduce(function (a, c) { return a + c.viejos; }, 0);

  return {
    ok: true,
    tienda: tienda, nombreTienda: nombreTienda(ss, tienda),
    moneda: monedaDeTienda(ss, tienda) || '',
    dia: dia, hoy: hoy,
    // Cuántos entraron ese día, para saber si la lista es grande porque
    // hubo mucha venta o porque nadie la trabajó.
    entraronEseDia: delDia,
    cerradosEseDia: cerradosAyer.length,
    categorias: categorias,
    enCamino: enCamino.length,
    motivos: Object.keys(porMotivo).map(function (k) { return porMotivo[k]; })
      .sort(function (a, b) { return b.n - a.n; }),
    novedades: { abiertas: novAbiertas, viejas: novViejas },
    total: totalPendiente,
    totalViejo: totalViejo,
    valorPendiente: Math.round(
      categorias.reduce(function (a, c) { return a + c.valor; }, 0) * 100) / 100,
    hayAlgo: totalPendiente > 0,
  };
}

/** El reporte del día, en texto, para el correo y para imprimir. */
function reporteDelDiaTexto(r) {
  const L = [];
  L.push('QUÉ QUEDÓ DE AYER · ' + r.nombreTienda);
  L.push(r.dia);
  L.push('');

  if (!r.hayAlgo) {
    L.push('No quedó nada esperando. Todo lo de ese día está cerrado o en camino.');
    return L.join('\n');
  }

  L.push('HAY ' + r.total + ' PEDIDO' + (r.total === 1 ? '' : 'S') + ' ESPERANDO A ALGUIEN.');
  if (r.totalViejo) {
    L.push(r.totalViejo + ' de esos llevan más tiempo del que aguantan.');
  }
  L.push('');

  const ancho = 34;
  const pad = function (s) {
    s = String(s);
    return s.length >= ancho ? s.slice(0, ancho) : s + new Array(ancho - s.length + 1).join(' ');
  };

  L.push(pad('CATEGORÍA') + 'CANT   LLEVAN MUCHO');
  L.push(new Array(ancho + 22).join('─'));
  r.categorias.forEach(function (c) {
    if (!c.n) return;
    L.push(pad(c.nombre) + pad3_(c.n) + '    ' +
           (c.limite === null ? '—' : c.viejos + ' (+' + c.limite + ' d)'));
  });
  L.push(new Array(ancho + 22).join('─'));
  L.push(pad('TOTAL ESPERANDO') + pad3_(r.total) + '    ' + r.totalViejo);
  L.push('');
  L.push('En camino, sin nada que hacer: ' + r.enCamino);
  L.push('Entraron ese día: ' + r.entraronEseDia + ' · se cerraron: ' + r.cerradosEseDia);
  if (r.valorPendiente) {
    L.push('Plata detenida en lo que espera: ' + redondear_(r.valorPendiente) +
           ' ' + r.moneda);
  }

  if (r.motivos.length) {
    L.push('');
    L.push('POR QUÉ ESTÁN ABIERTAS LAS NOVEDADES');
    r.motivos.forEach(function (m) {
      L.push('  ' + pad(m.grupo.replace(/_/g, ' ')) + pad3_(m.n) +
             (m.viejas ? '    ' + m.viejas + ' de más de un día' : ''));
    });
  }

  /**
   * Lo más viejo, con nombre y teléfono.
   *
   * Un total no se puede trabajar. Lo que se puede trabajar es una lista
   * de personas a las que llamar, y por eso van los cinco más quietos
   * con su número al lado.
   */
  const viejos = [];
  r.categorias.forEach(function (c) {
    c.casos.forEach(function (p) {
      if (c.limite !== null && p.dias > c.limite) viejos.push({ cat: c.nombre, p: p });
    });
  });
  if (viejos.length) {
    L.push('');
    L.push('LOS QUE LLEVAN MÁS QUIETOS');
    viejos.sort(function (a, b) { return b.p.dias - a.p.dias; })
      .slice(0, 8).forEach(function (x) {
        L.push('  · ' + (x.p.cliente || 'Sin nombre') + ' · ' + x.p.dias + ' d · ' +
               x.cat + (x.p.telefono ? ' · ' + x.p.telefono : ''));
      });
  }

  L.push('');
  L.push('— Nova');
  return L.join('\n');
}

function pad3_(n) {
  const s = String(n);
  return s.length >= 4 ? s : new Array(4 - s.length + 1).join(' ') + s;
}

/** La acción de la pantalla. */
function apiReporteDia(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const dia = /^\d{4}-\d{2}-\d{2}$/.test(String(p.dia || '')) ? p.dia : '';
  try {
    const r = reporteDelDia(s.sheetId, tienda, dia);
    /**
     * El texto se arma AQUÍ y no en el navegador.
     *
     * Es el mismo que va al correo. Si la pantalla lo rearmara por su
     * cuenta, el día que uno de los dos cambie dirían cosas distintas
     * del mismo día, y no habría forma de saber a cuál creerle.
     */
    r.texto = reporteDelDiaTexto(r);
    return r;
  } catch (e) {
    return { ok: false, error: e.message };
  }
}
