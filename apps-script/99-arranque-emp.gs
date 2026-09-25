/**
 * ═══════════════════════════════════════════════════════════════
 *  NOVA EMPRESARIAL EN UNA SOLA PETICIÓN
 * ═══════════════════════════════════════════════════════════════
 *
 * ┌─ LO QUE SE MIDIÓ ──────────────────────────────────────────┐
 * │                                                            │
 * │ «se demora mucho en cargar».                                │
 * │                                                            │
 * │ Conté las peticiones que hace la pantalla al abrirse, en    │
 * │ un navegador de verdad: DIECINUEVE.                        │
 * │                                                            │
 * │   accesos · alarmas · auditoria · auditoria_casos · cas ·   │
 * │   equipo · estados · fuentes · historial · listar ×2 ·      │
 * │   meta_estado · notas · productos · recuento · reparto ·    │
 * │   reporte_dia · resumen · semaforo                         │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ POR QUÉ DIECINUEVE CUESTAN TANTO ─────────────────────────┐
 * │                                                            │
 * │ No son los datos: son los VIAJES. Cada petición a Apps      │
 * │ Script arranca un motor de Google, comprueba la sesión,     │
 * │ corre y devuelve. Eso cuesta entre medio segundo y dos      │
 * │ SIEMPRE, traiga tres filas o tres mil.                     │
 * │                                                            │
 * │ Y hay algo peor, que es lo que de verdad duele: Google      │
 * │ SERIALIZA las peticiones del mismo usuario. Las diecinueve  │
 * │ que el navegador manda «a la vez» se ejecutan una detrás de │
 * │ otra. Diecinueve arranques, diecinueve `openById`,          │
 * │ diecinueve lecturas de las mismas hojas.                    │
 * │                                                            │
 * │ Juntarlas en una no solo quita dieciocho viajes: dentro de  │
 * │ UNA ejecución, `libro_()` memoriza el libro y cada hoja se  │
 * │ abre una vez para todas las secciones.                     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ LO QUE NO SE JUNTA, Y POR QUÉ ────────────────────────────┐
 * │                                                            │
 * │ Nada que ESCRIBA. Esto es solo lectura: si una petición de  │
 * │ lectura se pierde, se reintenta y ya. Meter una escritura   │
 * │ aquí la haría depender de que las otras dieciocho salgan    │
 * │ bien.                                                      │
 * │                                                            │
 * │ Y nada que llame afuera: `meta_traer` habla con Facebook y  │
 * │ puede tardar lo que quiera. Un arranque que espera a un     │
 * │ servidor ajeno deja la pantalla en blanco por culpa de otro.│
 * │ `meta_estado` sí entra: solo mira si hay llave guardada.    │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y SI UNA SECCIÓN FALLA ───────────────────────────────────┐
 * │                                                            │
 * │ Cada una va en su propio try. Una hoja que falta no puede   │
 * │ dejar la pantalla entera en blanco — el peor síntoma de     │
 * │ todos, porque no se distingue de «todavía está cargando».   │
 * │                                                            │
 * │ Lo que falló va en `fallaron`, con su nombre y su motivo, y │
 * │ la pantalla vuelve a pedir ESAS por separado.              │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */

/**
 * Qué trae el arranque, y con qué se llama cada cosa.
 *
 * Está como tabla y no como diecisiete líneas de código para que agregar
 * una sección sea agregar una fila — y para que se vea de un vistazo qué
 * viaja al abrir la pantalla, que es la lista que hay que vigilar para
 * que esto no vuelva a crecer a diecinueve.
 */
const ARRANQUE_EMP = [
  ['resumen',          function (s, p) { return apiResumen(s, p); }],
  ['pedidos',          function (s, p) {
    return apiListar(s, { entidad: 'Pedidos', tienda: p.tienda, limite: p.limite });
  }],
  ['novedades',        function (s, p) {
    return apiListar(s, { entidad: 'Novedades', tienda: p.tienda, limite: p.limite });
  }],
  ['notas',            function (s, p) { return apiNotas(s, p); }],
  ['accesos',          function (s, p) { return apiAccesos(s, p); }],
  ['equipo',           function (s, p) { return apiEquipo(s, p); }],
  ['alarmas',          function (s, p) { return apiAlarmas(s, p); }],
  ['productos',        function (s, p) { return apiProductos(s, p); }],
  ['cas',              function (s, p) { return apiCas(s, p); }],
  ['estados',          function (s, p) { return apiEstados(s, p); }],
  ['historial',        function (s, p) { return apiHistorial(s, p); }],
  ['fuentes',          function (s, p) { return apiFuentes(s, p); }],
  ['recuento',         function (s, p) { return apiRecuento(s, p); }],
  ['reporte_dia',      function (s, p) { return apiReporteDia(s, p); }],
  ['semaforo',         function (s, p) { return apiSemaforo(s, p); }],
  ['auditoria',        function (s, p) { return apiAuditoria(s, p); }],
  ['auditoria_casos',  function (s, p) { return apiAuditoriaCasos(s, p); }],
  ['reparto',          function (s, p) { return apiReparto(s, p); }],
  ['meta_estado',      function (s, p) { return apiMetaEstado(s, p); }],
  /**
   * El chat entra porque su punto rojo tiene que estar puesto ANTES de
   * que a nadie se le ocurra abrir el panel. Sin él aquí, entrar a Nova
   * costaba dos viajes: el arranque y el chat.
   */
  ['chat',             function (s, p) { return apiChat(s, { con: '' }); }],
];

/**
 * Todo lo que la pantalla pinta al entrar, de una.
 *
 * Devuelve `partes`, con la respuesta tal cual de cada acción: la
 * pantalla no tiene que aprender un formato nuevo, le da a cada pintor
 * exactamente lo que ya esperaba recibir de su propia petición.
 */
function apiArranque(s, p) {
  const tienda = String(p.tienda || s.tiendas[0] || '').trim();
  if (s.tiendas.indexOf(tienda) === -1) {
    return { ok: false, error: 'No tienes acceso a esa tienda.' };
  }
  const args = { tienda: tienda, mes: p.mes || '', limite: p.limite || 0 };

  const partes = {};
  const fallaron = [];
  const t0 = Date.now();

  ARRANQUE_EMP.forEach(function (par) {
    const nombre = par[0];
    try {
      const r = par[1](s, args);
      /**
       * Un «no tienes permiso» NO es un fallo del arranque.
       *
       * La admin no ve el reparto y la gestora no ve la auditoría: esas
       * secciones responden `ok:false` a propósito. Ponerlas en
       * `fallaron` haría que la pantalla las volviera a pedir una por
       * una —diecinueve viajes otra vez, y el mismo no— por cada
       * persona que no es la dueña.
       */
      partes[nombre] = r;
    } catch (e) {
      fallaron.push({ parte: nombre, error: String(e && e.message || e) });
    }
  });

  return {
    ok: true, tienda: tienda, partes: partes, fallaron: fallaron,
    // Cuánto tardó de verdad, para no volver a discutirlo de memoria.
    ms: Date.now() - t0,
  };
}
