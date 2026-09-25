/**
 * ═══════════════════════════════════════════════════════════════
 *  ASIGNAR · repartir los pedidos entre quienes los trabajan
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ POR QUÉ ESTO FALTABA Y NADIE LO NOTÓ ─────────────────────┐
 * │                                                            │
 * │ Nova siempre supo filtrar: una gestora ve los pedidos      │
 * │ cuya columna `gestora_asignada` lleva su nombre, y no ve    │
 * │ los de las demás. Eso estaba escrito, probado y bien.      │
 * │                                                            │
 * │ Lo que no existía era la otra mitad: NO HABÍA NINGUNA      │
 * │ MANERA DE PONER ESE NOMBRE. Ni en la pantalla ni en el     │
 * │ servidor. La columna solo se pintaba —«Sin asignar»— y     │
 * │ nunca se escribía.                                         │
 * │                                                            │
 * │ Mientras la dueña trabajó sola no se notó: ella lo ve      │
 * │ todo. Se habría notado el primer día que entrara una       │
 * │ gestora, viera cero pedidos y cero novedades, y concluyera │
 * │ —con razón— que la herramienta no sirve.                   │
 * │                                                            │
 * │ Un filtro sin forma de llenar aquello que filtra no es     │
 * │ media función: es una función que hace daño.               │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ SE GUARDA EL NOMBRE Y NO EL id ───────────────────┐
 * │                                                            │
 * │ `gestora_asignada` es una columna que también llega desde  │
 * │ las plataformas: Dropi y Effi exportan el nombre de quien  │
 * │ gestionó. Si Nova guardara ahí un id suyo, la próxima      │
 * │ importación lo pisaría con un nombre y el filtro dejaría   │
 * │ de encontrar a nadie.                                      │
 * │                                                            │
 * │ Así que se guarda el NOMBRE CANÓNICO, el de la hoja        │
 * │ Equipo — no el que escriba quien asigna. Y al filtrar se   │
 * │ acepta también el correo, porque alguna plataforma exporta │
 * │ eso. Es la diferencia entre un emparejamiento que aguanta  │
 * │ la realidad y uno que aguanta el ejemplo.                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * Quién puede repartir trabajo. Una gestora no se asigna pedidos a sí
 * misma: eso convierte el reparto en una carrera por los fáciles.
 */
function puedeAsignar_(s) {
  return s && (s.rol === 'dueno' || s.rol === 'admin');
}

/**
 * Las personas a las que se les puede asignar algo en una tienda.
 *
 * Solo activas, y solo las que cubren esa tienda. Asignarle un pedido
 * de Guatemala a alguien que solo tiene Ecuador crea una fila que esa
 * persona no puede ni ver ni tocar — y el pedido queda en un limbo
 * peor que «sin asignar», porque parece que alguien lo está mirando.
 */
function asignablesDe_(ss, tienda) {
  const sh = ss.getSheetByName('Equipo');
  if (!sh || sh.getLastRow() < 2) return [];
  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  const out = [];
  for (let i = 1; i < d.length; i++) {
    const nombre = String(d[i][c('nombre')] || '').trim();
    if (!nombre) continue;
    if (norm(d[i][c('estado')]) === 'inactivo') continue;
    const suya = String(d[i][c('tienda')] || '').trim();
    if (tienda && suya && suya !== '*' &&
        suya.split(/[,;]/).map(function (x) { return x.trim(); }).indexOf(tienda) === -1) {
      continue;
    }
    out.push({
      id: String(d[i][c('id')] || ''),
      nombre: nombre,
      correo: String(d[i][c('correo')] || '').toLowerCase().trim(),
      rol: rolCanonico(d[i][c('rol')]) || String(d[i][c('rol')] || ''),
    });
  }
  return out;
}

/**
 * Cómo está repartido el trabajo hoy.
 *
 * El número que de verdad importa es `sinAsignar`: son los pedidos que
 * NADIE está mirando salvo la dueña. Mientras ese número sea todo el
 * total, tener equipo no sirve de nada.
 */
function apiReparto(s, p) {
  if (!puedeAsignar_(s)) return { ok: false, error: 'Tu rol no reparte trabajo.' };
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }

  const ss = libro_(s.sheetId);
  const gente = asignablesDe_(ss, tienda);
  const porNombre = {};
  gente.forEach(function (g) {
    porNombre[norm(g.nombre)] = g;
    if (g.correo) porNombre[norm(g.correo)] = g;
    g.total = 0; g.abiertos = 0;
  });

  const sh = ss.getSheetByName('Pedidos');
  const out = { ok: true, tienda: tienda, gente: gente,
                total: 0, sinAsignar: 0, aNadieConocido: [], huerfanos: {} };
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };
  const cG = c('gestora_asignada');
  if (cG === -1) {
    out.error = 'La hoja Pedidos no tiene columna gestora_asignada. ' +
                'Corre bootstrapTodo() una vez.';
    return out;
  }

  const cerrados = ['entregado', 'devolucion', 'cancelado'];
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')] || '').trim() !== tienda) continue;
    out.total++;
    const quien = String(d[i][cG] || '').trim();
    if (!quien) { out.sinAsignar++; continue; }

    const g = porNombre[norm(quien)];
    if (!g) {
      /**
       * Un nombre que no corresponde a nadie de Equipo. Pasa cuando la
       * plataforma exporta a alguien que ya no está, o escrito distinto.
       * Se cuenta y se nombra: son pedidos que NADIE ve, porque su
       * supuesta dueña no puede entrar.
       */
      out.huerfanos[quien] = (out.huerfanos[quien] || 0) + 1;
      continue;
    }
    g.total++;
    const est = norm(d[i][c('estado_nova')] || d[i][c('estado_canonico')] || d[i][c('estado')]);
    if (cerrados.indexOf(est) === -1) g.abiertos++;
  }

  out.aNadieConocido = Object.keys(out.huerfanos).map(function (k) {
    return { nombre: k, pedidos: out.huerfanos[k] };
  }).sort(function (a, b) { return b.pedidos - a.pedidos; });

  return out;
}

/**
 * Asignar pedidos a una persona. Uno o muchos, en una sola escritura.
 *
 * `a` vacío DESASIGNA — y eso es a propósito, no un descuido: repartir
 * mal y no poder deshacerlo es peor que no repartir.
 */
function apiAsignar(s, p) {
  if (!puedeAsignar_(s)) return { ok: false, error: 'Tu rol no reparte trabajo.' };

  const ids = (Array.isArray(p.ids) ? p.ids : [p.id])
    .map(function (x) { return String(x || '').trim(); })
    .filter(function (x) { return x; });
  if (!ids.length) return { ok: false, error: 'No me dijiste qué pedidos asignar.' };
  if (ids.length > 500) {
    return { ok: false, error: 'Son demasiados de una. Máximo 500 por vez.' };
  }

  const ss = libro_(s.sheetId);
  const quien = String(p.a || '').trim();

  /**
   * El nombre se canoniza contra la hoja Equipo, SIEMPRE.
   *
   * Si se guardara lo que venga de la pantalla, un día entraría
   * «Andrea» y otro «Andrea R.» y el filtro dejaría de encontrar la
   * mitad de sus pedidos sin que nadie entienda por qué.
   */
  let persona = null;
  if (quien) {
    const gente = asignablesDe_(ss, String(p.tienda || '').trim());
    persona = gente.filter(function (g) {
      return g.id === quien || norm(g.nombre) === norm(quien) ||
             (g.correo && norm(g.correo) === norm(quien));
    })[0];
    if (!persona) {
      return { ok: false, error: 'No encuentro a esa persona activa en tu equipo, ' +
                                 'o no tiene acceso a esta tienda.' };
    }
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(20000)) return { ok: false, error: 'Hay otro cambio guardándose.' };
  try {
    const sh = ss.getSheetByName('Pedidos');
    if (!sh || sh.getLastRow() < 2) return { ok: false, error: 'No hay pedidos.' };

    const d = sh.getDataRange().getValues();
    const e = d[0].map(norm);
    const c = function (n) { return e.indexOf(n); };
    const cG = c('gestora_asignada'), cId = c('id'), cT = c('tienda');
    if (cG === -1) {
      return { ok: false, error: 'La hoja Pedidos no tiene columna gestora_asignada. ' +
                                 'Corre bootstrapTodo() una vez.' };
    }

    const buscar = {};
    ids.forEach(function (x) { buscar[x] = 1; });

    const valor = persona ? persona.nombre : '';
    const tocados = [], pedidosTocados = [];
    let deOtraTienda = 0, iguales = 0;

    for (let i = 1; i < d.length; i++) {
      const id = String(d[i][cId] || '').trim();
      if (!buscar[id]) continue;
      // Nunca fuera de las tiendas de quien asigna.
      if (cT !== -1 && s.tiendas.indexOf(String(d[i][cT] || '').trim()) === -1) {
        deOtraTienda++; continue;
      }
      const antes = String(d[i][cG] || '').trim();
      if (norm(antes) === norm(valor)) { iguales++; continue; }
      tocados.push({ fila: i + 1, antes: antes });
      pedidosTocados.push(id);
      d[i][cG] = valor;
    }

    if (!tocados.length) {
      return { ok: true, asignados: 0, iguales: iguales, deOtraTienda: deOtraTienda,
               a: valor, porque: iguales
                 ? 'Ya estaban así.'
                 : (deOtraTienda ? 'Todos eran de otra tienda.' : 'No encontré esos pedidos.') };
    }

    /**
     * Se escribe la COLUMNA entera de una vez, no celda por celda.
     *
     * Trescientas escrituras sueltas son trescientas idas y vueltas a
     * Google y varios minutos; un solo `setValues` de una columna es
     * una. Y como solo se toca esa columna, ninguna otra se puede
     * pisar por accidente.
     */
    const columna = d.slice(1).map(function (f) { return [f[cG]]; });
    sh.getRange(2, cG + 1, columna.length, 1).setValues(columna);

    tocados.forEach(function (t) {
      registrarMovimiento(s, 'Pedidos', String(d[t.fila - 1][cId] || ''),
                          'gestora_asignada', t.antes, valor);
    });

    /**
     * Las novedades de esos pedidos van con ellos.
     *
     * Si no, la gestora vería el pedido y no la novedad que hay que
     * resolver — que es precisamente el trabajo. Son dos hojas y dos
     * columnas con nombres distintos (`gestora_asignada` y `gestora`),
     * y por eso esto se olvida tan fácil.
     */
    let novedades = 0;
    try { novedades = asignarNovedadesDe_(ss, pedidosTocados, valor, s); }
    catch (err) { /* el pedido ya quedó asignado; la novedad se reintenta */ }

    return { ok: true, asignados: tocados.length, iguales: iguales,
             deOtraTienda: deOtraTienda, novedades: novedades,
             a: valor, desasignado: !valor };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

/** Mueve la columna `gestora` de las novedades de esos pedidos. */
function asignarNovedadesDe_(ss, pedidoIds, valor, s) {
  const sh = ss.getSheetByName('Novedades');
  if (!sh || sh.getLastRow() < 2) return 0;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const cG = e.indexOf('gestora');
  const cP = e.indexOf('pedido_id');
  const cId = e.indexOf('id');
  if (cG === -1 || cP === -1) return 0;

  const buscar = {};
  pedidoIds.forEach(function (x) { buscar[x] = 1; });

  let n = 0;
  const cambios = [];
  for (let i = 1; i < d.length; i++) {
    if (!buscar[String(d[i][cP] || '').trim()]) continue;
    const antes = String(d[i][cG] || '').trim();
    if (norm(antes) === norm(valor)) continue;
    cambios.push({ id: cId === -1 ? '' : String(d[i][cId] || ''), antes: antes });
    d[i][cG] = valor;
    n++;
  }
  if (!n) return 0;

  const columna = d.slice(1).map(function (f) { return [f[cG]]; });
  sh.getRange(2, cG + 1, columna.length, 1).setValues(columna);
  cambios.forEach(function (x) {
    registrarMovimiento(s, 'Novedades', x.id, 'gestora', x.antes, valor);
  });
  return n;
}
