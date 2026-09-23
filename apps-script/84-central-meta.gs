/**
 * ═══════════════════════════════════════════════════════════
 *  META, VISTO DESDE LA CONSOLA
 * ═══════════════════════════════════════════════════════════
 *
 * Lo que había en Nova Central era una maqueta: un cuadro para subir un
 * CSV y cuatro cifras de alcance, CPA, ROAS e inversión escritas a mano
 * en el HTML, que no salían de ninguna parte y no cambiaban nunca. Eso
 * es peor que una pantalla vacía: una pantalla vacía se nota, y un
 * número inventado se cree.
 *
 * Lo que va en su lugar es lo único que Central puede decir de verdad
 * sobre Meta: QUIÉN ESTÁ CONECTADO Y QUIÉN NO. Central es la consola
 * que mira a todos los clientes a la vez; el gasto, las campañas y el
 * CPA de cada tienda se miran en Nova Empresarial, adentro de esa
 * cuenta, que es donde está el dato.
 *
 * ── POR QUÉ AQUÍ NO SE GUARDA NINGUNA LLAVE ──
 *
 * La llave de Meta es POR CLIENTE y se guarda en las Propiedades del
 * script con el sheetId de ese cliente en el nombre. La dueña de cada
 * cuenta la escribe en SU Nova Empresarial. Poner aquí un campo para
 * escribirla sería darle a la consola una forma de conectar cuentas
 * ajenas desde una sola pantalla — y el día que alguien más opere
 * Central, esa pantalla es la que hay que no haber construido.
 *
 * Esto solo LEE si existe. El servidor nunca devuelve la llave, ni
 * recortada, igual que en Empresarial.
 */

function centralMeta(s, p) {
  if (!puedeCentral(s, 'crear_cliente')) {
    return { ok: false, error: 'Tu rol no puede ver las conexiones de los clientes.' };
  }

  const props = PropertiesService.getScriptProperties();
  const clientes = [];
  let error = '';

  try {
    listarClientes().forEach(function (c) {
      if (!c.sheetId) return;

      const clave = 'META_TOKEN_' + String(c.sheetId).slice(0, 44);
      const modulos = modulosDeCliente_(c.sheetId);
      const tienePauta = modulos.indexOf('pauta') !== -1;

      let prueba = null;
      try {
        const crudo = props.getProperty(clave + '_PRUEBA');
        if (crudo) prueba = JSON.parse(crudo);
      } catch (e) { prueba = null; }

      /**
       * Las tiendas y su cuenta publicitaria. Una llave guardada sin
       * número de cuenta no trae nada, y desde fuera las dos cosas se
       * ven igual de «conectadas»: por eso se cuentan aparte.
       */
      const tiendas = [];
      let sinCuenta = 0;
      try {
        const ss = SpreadsheetApp.openById(c.sheetId);
        tiendasActivas_(ss).forEach(function (t) {
          const cuenta = metaCuenta(ss, t);
          if (!cuenta) sinCuenta++;
          tiendas.push({
            id: t,
            // Nunca el nombre de la hoja de gastos: el de pantalla.
            nombre: nombreTienda(ss, t),
            cuenta: cuenta,
            moneda: monedaDeTienda(ss, t) || '',
          });
        });
      } catch (e) {
        // Una cuenta que no se puede abrir no tumba la lista entera.
        tiendas.length = 0;
      }

      const hayLlave = !!props.getProperty(clave);
      clientes.push({
        id: c.id, empresa: c.empresa, estado: norm(c.estado) || 'activo',
        tienePauta: tienePauta,
        hayLlave: hayLlave,
        guardadaEn: props.getProperty(clave + '_FECHA') || '',
        ultimaPrueba: prueba,
        tiendas: tiendas,
        sinCuenta: sinCuenta,
        /**
         * Tres estados, no dos. «Conectado» de verdad es llave + cuenta
         * en todas las tiendas; con llave y sin cuenta el gasto no
         * entra, y decir que está conectado sería mentir despacio.
         */
        listo: hayLlave && tiendas.length > 0 && sinCuenta === 0,
      });
    });
  } catch (e) {
    error = e.message;
  }

  const conPauta = clientes.filter(function (c) { return c.tienePauta; });

  return {
    ok: true,
    hoy: ahoraISO().slice(0, 10),
    clientes: clientes,
    resumen: {
      total: clientes.length,
      conPlanDePauta: conPauta.length,
      listos: conPauta.filter(function (c) { return c.listo; }).length,
      aMedias: conPauta.filter(function (c) { return c.hayLlave && !c.listo; }).length,
      sinLlave: conPauta.filter(function (c) { return !c.hayLlave; }).length,
    },
    /**
     * Lo de ella. Nova y novAcademy todavía no pautan, y eso se dice
     * con todas las letras en vez de dibujar un panel de campañas en
     * cero que parece una mala semana.
     */
    propias: CENTRAL_PROPIAS.map(function (x) {
      return { id: x.id, nombre: x.nombre, que: x.que, pauta: false };
    }),
    error: error,
  };
}

/** Los dos negocios de ella que algún día van a pautar. */
const CENTRAL_PROPIAS = [
  { id: 'nova', nombre: 'Nova', que: 'La plataforma que le vende a las tiendas.' },
  { id: 'academy', nombre: 'novAcademy', que: 'Los cursos y el acompañamiento.' },
];
