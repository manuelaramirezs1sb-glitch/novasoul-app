/**
 * ═══════════════════════════════════════════════════════════════
 *  NOVA CENTRAL · HOY, LEYENDO LAS HOJAS DE VERDAD
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE ELLA DIJO ─────────────────────────────────────────┐
 * │                                                            │
 * │ «¿Cómo que aún no lee la hoja? Ni la hoja ni Nova           │
 * │  Empresarial. Recuerda que Central es la red que conecta   │
 * │  todo con todos. Así que ponte las pilas y no dejes que    │
 * │  eso se pierda, porque más adelante es duro arreglarlo.»   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ QUÉ HABÍA AHÍ ────────────────────────────────────────────┐
 * │                                                            │
 * │ Una tarjeta marcada «EJEMPLO · esta tarjeta todavía no lee │
 * │ la hoja» con cifras inventadas —Q 284.500, 342 pedidos,    │
 * │ 88% de efectividad, 5 novedades— y un equipo inventado:    │
 * │ Daniela R., Camila R., Katherin P.                         │
 * │                                                            │
 * │ El cartel de EJEMPLO era honesto, pero una etiqueta no     │
 * │ arregla el problema: la pantalla que se supone que es el   │
 * │ centro de todo era la única que no estaba conectada a      │
 * │ nada. Y tiene razón en lo otro: cuanto más tiempo pase,    │
 * │ más partes se cuelgan de la forma equivocada.              │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ QUÉ SIGNIFICA «LA RED QUE CONECTA TODO» ──────────────────┐
 * │                                                            │
 * │ Central no tiene datos propios de ninguna tienda. Lo que   │
 * │ tiene es la lista de clientes, y de cada uno el `sheet_id` │
 * │ de SU libro. Leer «la red» es abrir esos libros y contar.  │
 * │                                                            │
 * │   Central ──▶ Clientes ──▶ sheet_id ──▶ el libro del       │
 * │                                          cliente           │
 * │                                          (Pedidos,         │
 * │                                           Novedades,       │
 * │                                           Equipo)          │
 * │                                                            │
 * │ Por eso esta función no calcula nada nuevo: llama a        │
 * │ `agregarMes`, que es la MISMA cuenta que usa Nova          │
 * │ Empresarial. Si fueran dos cuentas distintas, Central y    │
 * │ Empresarial podrían decirle cifras distintas de la misma   │
 * │ tienda el mismo día, y no habría forma de saber cuál       │
 * │ tiene razón.                                               │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ UNA TIENDA A LA VEZ ──────────────────────────────────────┐
 * │                                                            │
 * │ Es la regla de toda Nova y aquí también: dos tiendas en la │
 * │ misma pantalla invitan a compararlas, y son negocios       │
 * │ distintos, en países distintos, con monedas distintas.     │
 * │ Se elige una y se ve esa.                                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * La lista de clientes de la red, con su libro.
 *
 * Se sacó a su propia función porque la usan el «hoy» y cualquier cosa
 * que venga después. Un cliente inactivo no entra: no es que tenga cero
 * pedidos, es que ya no está.
 */
function centralRed_() {
  const out = { clientes: [], error: '' };
  try {
    const sh = libro_(IDS_().central).getSheetByName('Clientes');
    if (!sh || sh.getLastRow() < 2) {
      out.error = 'Todavía no hay clientes en la hoja Clientes de Nova Central.';
      return out;
    }
    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    for (let i = 1; i < d.length; i++) {
      const empresa = String(d[i][c('empresa')] || '').trim();
      const sid = String(d[i][c('sheet_id')] || '').trim();
      if (!empresa) continue;
      if (norm(d[i][c('estado')]) === 'inactivo') continue;
      out.clientes.push({
        id: String(d[i][c('id')] || ''),
        empresa: empresa,
        sheetId: sid,
        plan: c('plan') === -1 ? '' : String(d[i][c('plan')] || ''),
        estado: c('estado') === -1 ? '' : norm(d[i][c('estado')]),
        // Un cliente sin libro no es un error de datos: es un cliente
        // creado a medias, y hay que poder verlo desde aquí.
        sinLibro: !sid,
      });
    }
  } catch (err) {
    out.error = 'No pude leer la lista de clientes: ' + err.message;
  }
  return out;
}

/**
 * El «Hoy» de Central: la red, y una tienda mirada de cerca.
 *
 * Todo lo que devuelve sale de una hoja. Cuando algo no se puede leer,
 * viaja el porqué en vez de un número — un cero que en realidad
 * significa «no pude abrir el libro» es peor que no mostrar nada.
 */
function centralHoy(s, p) {
  const hoy = ahoraISO().slice(0, 10);
  const out = {
    ok: true, hoy: hoy, mes: hoy.slice(0, 7),
    red: null, clientes: [], cliente: null, tiendas: [], tienda: '',
    kpis: null, equipo: [], alarmas: [], avisos: [],
  };

  const red = centralRed_();
  out.clientes = red.clientes.map(function (c) {
    return { id: c.id, empresa: c.empresa, plan: c.plan, sinLibro: c.sinLibro };
  });
  if (red.error) out.avisos.push(red.error);

  /** El estado de la red, en tres números que se pueden comprobar. */
  out.red = {
    clientes: red.clientes.length,
    conLibro: red.clientes.filter(function (c) { return !c.sinLibro; }).length,
    sinLibro: red.clientes.filter(function (c) { return c.sinLibro; }).length,
  };

  const elegido = String(p.cliente || '');
  const cl = red.clientes.filter(function (c) { return c.id === elegido && !c.sinLibro; })[0] ||
             red.clientes.filter(function (c) { return !c.sinLibro; })[0] || null;

  if (!cl) {
    out.avisos.push(red.clientes.length
      ? 'Ninguno de tus clientes tiene libro asignado todavía, así que no hay de dónde leer.'
      : 'Cuando crees tu primer cliente, aquí van a salir sus números de verdad.');
    return out;
  }
  out.cliente = { id: cl.id, empresa: cl.empresa, plan: cl.plan };

  let cs;
  try { cs = libro_(cl.sheetId); }
  catch (err) {
    out.avisos.push('No pude abrir el libro de ' + cl.empresa + ': ' + err.message);
    return out;
  }

  try {
    out.tiendas = tiendasActivas_(cs).map(function (t) {
      return { id: t, nombre: nombreTienda(cs, t) };
    });
  } catch (err) {
    out.avisos.push('No pude leer las tiendas de ' + cl.empresa + ': ' + err.message);
    return out;
  }

  const t = out.tiendas.filter(function (x) { return x.id === String(p.tienda || ''); })[0] ||
            out.tiendas[0] || null;
  if (!t) {
    out.avisos.push(cl.empresa + ' no tiene ninguna tienda activa.');
    return out;
  }
  out.tienda = t.id;
  out.tiendaNombre = t.nombre;

  /**
   * Los números del mes. Se usa `agregarMes`, la misma función que usa
   * Nova Empresarial — no una copia. Se le pasa una sesión de solo
   * lectura con el rol de dueña porque Central mira la cuenta entera,
   * no la de una gestora.
   */
  const sesionLectura = { rol: 'dueno', tiendas: [t.id], sheetId: cl.sheetId,
                          email: s && s.correo ? s.correo : '' };
  try {
    const m = agregarMes(cs, t.id, out.mes, sesionLectura, null);
    /**
     * La efectividad se cuenta sobre los pedidos que YA TERMINARON —
     * entregados más devueltos—, no sobre todos. Contarla sobre todos
     * castiga a una tienda por tener pedidos en camino, que es lo
     * normal a mitad de mes, y da un número que baja solo por vender.
     */
    const cerrados = m.entregados + m.devueltos;
    out.kpis = {
      moneda: monedaDeTienda(cs, t.id) || '',
      ventas: m.ventas,
      pedidos: m.pedidos,
      entregados: m.entregados,
      devueltos: m.devueltos,
      enCamino: m.pedidos - cerrados - m.cancelados - m.sinClasificar,
      cancelados: m.cancelados,
      efectividad: cerrados ? Math.round(m.entregados / cerrados * 100) : null,
      efectividadSobre: cerrados,
      novedades: m.novedades,
      sinMover: m.sinMover,
      sinClasificar: m.sinClasificar,
      vacio: m.pedidos === 0,
    };
    if (!cerrados) {
      out.avisos.push('Todavía no hay pedidos terminados este mes, así que la ' +
                      'efectividad no se puede calcular. No es 0%: es que aún no se sabe.');
    }
    if (m.sinClasificar) {
      out.avisos.push(m.sinClasificar + ' pedido(s) tienen un estado que Nova no entiende ' +
                      'y no entran en ninguna cuenta. Se arreglan en Empresarial → Estados.');
    }
  } catch (err) {
    out.avisos.push('No pude contar los pedidos de ' + t.nombre + ': ' + err.message);
  }

  /**
   * El equipo REAL de esa tienda. Aquí había tres personas inventadas.
   *
   * `ultima_conexion` sale de la hoja: si dice vacío, esa persona nunca
   * ha entrado, y eso es un dato —no un hueco que haya que rellenar con
   * algo bonito.
   */
  try {
    const sh = cs.getSheetByName('Equipo');
    if (sh && sh.getLastRow() > 1) {
      const d = sh.getDataRange().getValues();
      const e = d[0].map(norm);
      const c = function (n) { return e.indexOf(n); };
      for (let i = 1; i < d.length; i++) {
        const nombre = String(d[i][c('nombre')] || '').trim();
        if (!nombre) continue;
        if (norm(d[i][c('estado')]) === 'inactivo') continue;
        const suya = String(d[i][c('tienda')] || '').trim();
        if (suya && suya !== '*' && suya.split(/[,;]/).map(function (x) {
          return x.trim(); }).indexOf(t.id) === -1) continue;
        out.equipo.push({
          id: String(d[i][c('id')] || ''),
          nombre: nombre,
          correo: String(d[i][c('correo')] || ''),
          rol: rolCanonico(d[i][c('rol')]) || String(d[i][c('rol')] || ''),
          ultimaConexion: c('ultima_conexion') === -1 ? ''
            : String(aISO(d[i][c('ultima_conexion')], 'UTC') || ''),
        });
      }
    }
    if (!out.equipo.length) {
      out.avisos.push('En la hoja Equipo de ' + cl.empresa + ' no hay nadie asignado a ' +
                      t.nombre + '. Mientras no lo haya, esta tarjeta va a estar vacía.');
    }
  } catch (err) {
    out.avisos.push('No pude leer el equipo: ' + err.message);
  }

  /** Las alarmas de esa tienda, las mismas que ve ella en Empresarial. */
  try {
    const ev = evaluarAlarmas(cs, t.id);
    out.alarmas = (ev.alarmas || []).map(function (a) {
      return { id: a.id, nivel: a.nivel, nombre: a.nombre, titulo: a.titulo,
               detalle: a.detalle, casos: (a.casos || []).length };
    });
  } catch (err) {
    out.avisos.push('No pude leer las alarmas: ' + err.message);
  }

  return out;
}
