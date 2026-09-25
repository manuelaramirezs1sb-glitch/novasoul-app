/**
 * ═══════════════════════════════════════════════════════════════
 *  QUIÉN TRABAJA CADA TIENDA
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ SE ASIGNAN TIENDAS, NO PEDIDOS ───────────────────────────┐
 * │                                                            │
 * │ Este archivo empezó repartiendo pedidos uno por uno. Ella  │
 * │ lo corrigió de raíz:                                       │
 * │                                                            │
 * │   «de nada sirve asignar pedidos si ya asignaste la        │
 * │    tienda. La gestora de la tienda gestiona la tienda que  │
 * │    tiene asignada: todo lo que son pedidos, novedades y    │
 * │    CAS. Porque si la tienda tiene pedidos pendientes de    │
 * │    confirmar hace más de 3 días, alguien debe recibir esa  │
 * │    información — y no solo la admin o la dueña.            │
 * │    Y dos gestores o más pueden tener la misma tienda».     │
 * │                                                            │
 * │ El reparto caso por caso se borró. Lo que queda aquí es    │
 * │ la otra pregunta, que sí sirve todos los días: QUIÉN ESTÁ  │
 * │ TRABAJANDO ESTA TIENDA, y cuánto lleva hecho cada quien.   │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ DOS LISTAS DISTINTAS, Y POR QUÉ ──────────────────────────┐
 * │                                                            │
 * │ `gente`     quién tiene esta tienda asignada en Nova.      │
 * │             Entran, ven y trabajan.                        │
 * │                                                            │
 * │ `etiquetas` quién aparece como gestor en los datos, venga  │
 * │             del archivo histórico o de haber tocado un     │
 * │             caso. JAIME, ZULAY, APOYO… no necesitan tener  │
 * │             cuenta, y casi nunca la tienen.                │
 * │                                                            │
 * │ Separarlas es lo que le permite a una dueña probar Nova    │
 * │ una semana con su equipo entero sin crear a nadie, y ver   │
 * │ igual sus números por persona.                             │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * Quién puede ver el reparto de la tienda.
 *
 * La gestora no: saber cuánto lleva hecho cada compañera es información
 * de quien dirige, no de quien opera. Ella ve su tienda entera, que es
 * lo que necesita para trabajar.
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
  if (!puedeAsignar_(s)) return { ok: false, error: 'Tu rol no ve el reparto.' };
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
  /**
   * `sinTocar` ya NO quiere decir «nadie lo ve»: desde que se asignan
   * tiendas, todo lo de esta tienda lo ven todas sus gestoras. Quiere
   * decir «nadie le ha puesto la mano todavía», que es una pregunta
   * distinta y sigue siendo útil.
   */
  const out = { ok: true, tienda: tienda, gente: gente,
                total: 0, sinTocar: 0, aNadieConocido: [], huerfanos: {} };
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

  /**
   * ── LAS ETIQUETAS, QUE NO SON CUENTAS ──
   *
   * `gestionado_por` dice quién atendió el pedido según el archivo que
   * se importó: JAIME, ZULAY, APOYO. No hace falta que existan como
   * usuarios, y casi nunca existen.
   *
   * Se cuentan aparte para que la dueña tenga sus números por agente
   * desde el primer día, sin crear una sola cuenta. Es lo que le
   * permite probar Nova una semana con su equipo entero sin montarlo.
   */
  const cEt = c('gestionado_por');
  const porEtiqueta = {};

  const cerrados = ['entregado', 'devolucion', 'cancelado'];
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')] || '').trim() !== tienda) continue;
    out.total++;

    if (cEt !== -1) {
      const et = String(d[i][cEt] || '').trim();
      // Nadie le ha puesto la mano todavía. Se cuenta por ESTA columna,
      // no por la de acceso: desde que se asignan tiendas, «asignado» y
      // «trabajado» dejaron de ser lo mismo.
      if (!et) out.sinTocar++;
      if (et) {
        if (!porEtiqueta[et]) porEtiqueta[et] = { nombre: et, total: 0, abiertos: 0,
                                                  esUsuario: !!porNombre[norm(et)] };
        porEtiqueta[et].total++;
        const e2 = norm(d[i][c('estado_nova')] || d[i][c('estado_canonico')] || d[i][c('estado')]);
        if (cerrados.indexOf(e2) === -1) porEtiqueta[et].abiertos++;
      }
    }

    const quien = String(d[i][cG] || '').trim();
    if (!quien) continue;

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

  /**
   * Quién trabajó, según el archivo. Ordenado por volumen, y diciendo
   * de cada uno si además tiene cuenta en Nova — que es lo que separa
   * «esta persona ve sus pedidos» de «este nombre solo está anotado».
   */
  out.etiquetas = Object.keys(porEtiqueta).map(function (k) { return porEtiqueta[k]; })
    .sort(function (a, b) { return b.total - a.total; });
  out.sinEtiqueta = cEt === -1;

  return out;
}

/**
 * ── AQUÍ VIVÍA `apiAsignar`, Y SE BORRÓ ──
 *
 * Repartía pedidos uno por uno o en bloque. Funcionaba, estaba probado,
 * y sobraba. Ella:
 *
 *   «de nada sirve asignar pedidos si ya asignaste la tienda. No se
 *    asignan pedidos, se asignan tiendas. Y dos gestores o más pueden
 *    tener la misma tienda asignada».
 *
 * Con eso, repartir caso por caso no es una comodidad de menos: es una
 * tarea diaria inventada, y además dejaba huérfano lo que nadie tomaba
 * — que es justo lo que lleva más tiempo quieto.
 *
 * Se borró entero en vez de dejarlo apagado. Código muerto que todavía
 * funciona es una invitación a volver a engancharlo «mientras tanto»,
 * y de ahí no sale más.
 *
 * Las tiendas de cada persona se asignan en Permisos, con casillas.
 */
