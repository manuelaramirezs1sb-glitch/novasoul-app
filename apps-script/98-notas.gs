/**
 * ═══════════════════════════════════════════════════════════════
 *  LA BITÁCORA · una nota por cada intento de contacto
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «recuerda que deben dejar la nota por cada intento de       │
 * │  contacto. Deben aparecer todas, si se guardan que          │
 * │  aparezca en la info con fecha y nombre de quien puso la    │
 * │  nota».                                                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE PASABA, QUE ES PEOR QUE «NO SE GUARDÓ» ────────────┐
 * │                                                            │
 * │ Había UNA columna `nota` por pedido y una caja de texto     │
 * │ encima. Se guardaba bien. El problema es lo que pasaba la   │
 * │ segunda vez: quien anotaba el intento del jueves abría la   │
 * │ caja con lo del miércoles dentro, escribía encima, y el     │
 * │ intento del miércoles DEJABA DE EXISTIR.                    │
 * │                                                            │
 * │ Nadie ve ese borrado. No hay error, no hay aviso: la nota   │
 * │ anterior simplemente ya no está. Y después alguien mira el  │
 * │ pedido y cree que solo se llamó una vez.                    │
 * │                                                            │
 * │ Por eso esto no es «agregar historial»: es tapar una        │
 * │ pérdida de datos silenciosa.                               │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ NO SE EDITAN NI SE BORRAN ────────────────────────┐
 * │                                                            │
 * │ Igual que el chat del equipo. Una bitácora de intentos      │
 * │ donde se puede reescribir lo dicho no sirve para lo único   │
 * │ para lo que existe: saber qué se hizo, cuándo y quién.      │
 * │                                                            │
 * │ La auditoría se apoya en esto, y una auditoría sobre datos  │
 * │ que el auditado puede cambiar no es una auditoría.          │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/** Dónde se puede dejar una nota. Nada más: la lista es la puerta. */
const NOTAS_ENTIDADES = ['Pedidos', 'Novedades', 'CAS'];

/** Cuánto cabe en una nota, y cuántas trae una carga. */
const NOTA_LARGO = 1500;
const NOTAS_TOPE = 3000;

/**
 * Todas las notas de una tienda, de una sola lectura.
 *
 * La pantalla las indexa por entidad + id. Pedirlas de a un pedido serían
 * trescientas peticiones para abrir una lista que ya está en memoria, y
 * la hoja entera de notas es más pequeña que la de pedidos.
 */
function apiNotas(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const ss = libro_(s.sheetId);
  const out = { ok: true, tienda: tienda, notas: {}, total: 0 };

  const sh = ss.getSheetByName('Notas');
  if (!sh || sh.getLastRow() < 2) return out;

  const d = sh.getDataRange().getValues();
  const e = d[0].map(norm);
  const c = function (n) { return e.indexOf(n); };

  const todas = [];
  for (let i = 1; i < d.length; i++) {
    if (String(d[i][c('tienda')] || '').trim() !== tienda) continue;
    const ent = String(d[i][c('entidad')] || '').trim();
    const id = String(d[i][c('entidad_id')] || '').trim();
    if (!ent || !id) continue;
    todas.push({
      clave: ent + ':' + id,
      texto: String(d[i][c('texto')] || ''),
      autor: String(d[i][c('autor_nombre')] || d[i][c('autor')] || ''),
      cuando: String(d[i][c('creado_en')] || ''),
    });
  }

  /**
   * Si son muchísimas, se traen las más recientes. Nunca las más viejas:
   * lo que hace falta para trabajar hoy es el último intento, no el
   * primero. Y se dice cuántas quedaron fuera, para que nadie crea que
   * esas son todas.
   */
  todas.sort(function (a, b) { return a.cuando < b.cuando ? -1 : 1; });
  out.total = todas.length;
  const trozo = todas.slice(-NOTAS_TOPE);
  out.recortadas = todas.length - trozo.length;

  trozo.forEach(function (n) {
    if (!out.notas[n.clave]) out.notas[n.clave] = [];
    out.notas[n.clave].push({ texto: n.texto, autor: n.autor, cuando: n.cuando });
  });
  return out;
}

/**
 * Dejar una nota. Se agrega; nunca reemplaza a la anterior.
 *
 * Se comprueba que la fila exista Y que sea de una tienda suya antes de
 * escribir nada. Sin eso, mandar un id cualquiera dejaría notas colgando
 * de pedidos de otro cliente.
 */
function apiNotaAgregar(s, p) {
  const entidad = String(p.entidad || '').trim();
  if (NOTAS_ENTIDADES.indexOf(entidad) === -1) {
    return { ok: false, error: 'Ahí no se pueden dejar notas.' };
  }
  const id = String(p.id || '').trim();
  const texto = String(p.texto || '').trim();
  if (!id) return { ok: false, error: 'No sé de qué caso.' };
  if (!texto) return { ok: false, error: 'La nota está vacía.' };
  if (texto.length > NOTA_LARGO) {
    return { ok: false, error: 'La nota es muy larga. Máximo ' + NOTA_LARGO +
                               ' caracteres: son varias notas, no una.' };
  }

  const lock = LockService.getScriptLock();
  if (!lock.tryLock(15000)) return { ok: false, error: 'Hay otra nota guardándose.' };
  try {
    const ss = libro_(s.sheetId);
    const shE = ss.getSheetByName(entidad);
    if (!shE || shE.getLastRow() < 2) {
      return { ok: false, error: 'No encuentro ese ' + entidad.toLowerCase() + '.' };
    }
    const d = shE.getDataRange().getValues();
    const e = d[0].map(norm);
    const cId = e.indexOf('id'), cT = e.indexOf('tienda');

    let fila = -1;
    for (let i = 1; i < d.length; i++) {
      if (String(d[i][cId] || '').trim() === id) { fila = i; break; }
    }
    if (fila === -1) return { ok: false, error: 'No encuentro ese caso.' };

    const tienda = cT !== -1 ? String(d[fila][cT] || '').trim() : (s.tiendas[0] || '');
    if (s.tiendas.indexOf(tienda) === -1) {
      return { ok: false, error: 'Ese caso no es de una tienda tuya.' };
    }

    const sh = ss.getSheetByName('Notas');
    if (!sh) {
      return { ok: false, error: 'Falta la hoja Notas. Corre bootstrapTodo() una vez.' };
    }
    const enc = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0].map(norm);
    const nueva = {
      id: 'n' + Utilities.getUuid().slice(0, 10),
      tienda: tienda, entidad: entidad, entidad_id: id, texto: texto,
      autor: s.email || '', autor_nombre: s.nombre || s.email || '',
      creado_en: ahoraISO(),
    };
    sh.appendRow(enc.map(function (k) { return nueva[k] !== undefined ? nueva[k] : ''; }));

    /**
     * La columna `nota` de la fila se queda con la ÚLTIMA.
     *
     * No es duplicar el dato por comodidad: la auditoría levanta
     * «sin_nota» leyendo esa columna, las alarmas la miran, y el export
     * a Excel la lleva. Si dejara de escribirse, todo eso empezaría a
     * decir que nadie anotó nada justo cuando el equipo por fin anota.
     *
     * La bitácora completa vive en Notas; esto es el resumen de una
     * línea que el resto de Nova ya sabía leer.
     */
    const cN = e.indexOf('nota');
    if (cN !== -1) shE.getRange(fila + 1, cN + 1).setValue(texto);

    // Y quien anota es quien lo trabajó, si nadie lo había tocado antes.
    const cHizo = e.indexOf('gestionado_por');
    if (cHizo !== -1 && !String(d[fila][cHizo] || '').trim()) {
      shE.getRange(fila + 1, cHizo + 1).setValue(s.nombre || s.email || '');
    }

    return { ok: true, nota: { texto: texto, autor: nueva.autor_nombre,
                               cuando: nueva.creado_en },
             gestionadoPor: cHizo !== -1
               ? String(d[fila][cHizo] || '').trim() || (s.nombre || s.email || '') : '' };
  } catch (e) {
    return { ok: false, error: e.message };
  } finally { lock.releaseLock(); }
}

/**
 * ── LO QUE YA ESTABA ESCRITO ANTES DE QUE HUBIERA BITÁCORA ──
 *
 * Las notas de una sola columna que el equipo lleva meses escribiendo no
 * se tiran ni se migran: la pantalla las muestra como la primera entrada
 * de la bitácora, sin autor ni fecha porque nunca los tuvo.
 *
 * Se hace al pintar y no con una migración porque una migración que se
 * corre dos veces duplica todo, y ésta no tendría forma de saber si ya
 * corrió. La pantalla ya tiene la fila en la mano; no cuesta nada.
 */
