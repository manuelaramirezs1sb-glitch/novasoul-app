/**
 * ═══════════════════════════════════════════════════════════════
 *  NOVASOUL EN UNA SOLA PETICIÓN
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ POR QUÉ SE DEMORA, EN UNA FRASE ──────────────────────────┐
 * │                                                            │
 * │ No son los datos. Son los VIAJES.                          │
 * │                                                            │
 * │ Cada `nc(...)` de la pantalla es una petición HTTPS entera  │
 * │ a Apps Script: abrir conexión, arrancar el motor de Google, │
 * │ comprobar la sesión, correr, devolver, cerrar. Eso cuesta   │
 * │ entre medio segundo y dos segundos SIEMPRE, traiga tres     │
 * │ filas o tres mil.                                          │
 * │                                                            │
 * │ NovaSoul hacía cinco de esos viajes para pintar lo que ella │
 * │ mira al entrar. Cinco viajes son cinco arranques de motor,  │
 * │ y ahí están los segundos que ella siente.                  │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y POR QUÉ n8n NO ARREGLA ESTO ────────────────────────────┐
 * │                                                            │
 * │ Ella preguntó si n8n serviría para que no fuera tan lento.  │
 * │ La respuesta honesta es que no, y por una razón que no      │
 * │ tiene que ver con si n8n es bueno o malo:                   │
 * │                                                            │
 * │ n8n no quita viajes. Los AGREGA. La pantalla llamaría a     │
 * │ n8n, n8n llamaría a Google Sheets, Google respondería, n8n  │
 * │ respondería. Donde había un salto, hay dos. Y el dato sigue │
 * │ viviendo en la misma hoja, leído por la misma API.          │
 * │                                                            │
 * │ Lo que n8n sí hace bien —conectar servicios ajenos entre    │
 * │ sí sin escribir código— es otro problema, y ese día se      │
 * │ mira. Para ESTA demora, lo que sirve es hacer un viaje en   │
 * │ vez de cinco. Que es lo que hay en este archivo.            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE NO SE JUNTA, Y POR QUÉ ────────────────────────────┐
 * │                                                            │
 * │ `soulFamily` NO entra aquí. Esa abre los libros de las      │
 * │ tiendas —Nutrea EC y GT, miles de pedidos— y calcula el     │
 * │ semáforo. Son segundos de verdad, y meterla en el arranque  │
 * │ haría que la pantalla de Pendientes esperara por datos de   │
 * │ una tienda que ella ni está mirando.                        │
 * │                                                            │
 * │ La regla: se junta lo que se pinta AL ENTRAR. Lo que cuesta │
 * │ segundos y se mira a veces, se pide cuando se mira.         │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * Todo lo de NovaSoul que se pinta al entrar, de una.
 *
 * Cada sección va envuelta en su propio try. Una hoja que falta no
 * puede dejar la pantalla entera en blanco: eso ya pasó una vez con
 * Tránsitos, y el síntoma —una pantalla vacía sin ningún mensaje— es
 * el peor de todos porque no se distingue de «todavía está cargando».
 */
function soulArranque(s, p) {
  if (!soulPuede_(s)) return { ok: false, error: 'NovaSoul es de Manuela.' };

  const out = { ok: true, hoy: ahoraISO().slice(0, 10), fallaron: [] };

  const parte = function (nombre, clave, fn) {
    try {
      const r = fn();
      // Una sección que devuelve su propio error lo dice y no se pierde.
      if (r && r.ok === false) {
        out.fallaron.push({ seccion: nombre, porque: r.error || 'no respondió' });
      }
      out[clave] = r;
    } catch (e) {
      out.fallaron.push({ seccion: nombre, porque: e.message });
      out[clave] = null;
    }
  };

  /**
   * `hoy_` con guion bajo porque `hoy` ya es la fecha. Dos cosas con el
   * mismo nombre en la misma respuesta es cómo se consigue que la
   * pantalla pinte una fecha donde iba una lista.
   */
  parte('Hoy', 'hoy_', function () { return soulHoy(s, p); });
  parte('Mi plata', 'plata', function () { return soulPlataOrdenada(s, p); });
  parte('El cielo', 'cielo', function () { return soulCielo(s, p); });
  parte('Mi rutina', 'rutina', function () { return soulRutina(s, p); });
  parte('Universidad', 'materias', function () { return soulMaterias(s, p); });

  return out;
}
