/**
 * Nova · Monedas y conversión
 * ─────────────────────────────────────────────────────────────
 * Dos cosas distintas que se confunden fácil:
 *
 *   1. El CATÁLOGO — cuántos decimales tiene cada moneda y cómo se muestra.
 *      Es una tabla estática. Barata. Está completa para Latinoamérica.
 *
 *   2. La TASA — cuánto vale esa moneda un día concreto. Vive en la hoja
 *      `Tasas` y hay que alimentarla. Sin tasa NO hay conversión posible,
 *      por más completo que esté el catálogo.
 *
 * El spec es explícito: se convierte con "la tasa del DÍA DE LA TRANSACCIÓN,
 * no la de hoy". Si no, los márgenes salen mal y todas las alarmas de dinero
 * disparan en falso.
 *
 * Por eso, cuando falta una tasa, la conversión NO inventa un número:
 * devuelve null y marca la fila. Un margen vacío se ve; un margen calculado
 * con la tasa equivocada no se ve y miente.
 */

// ─── CATÁLOGO ────────────────────────────────────────────────
// dec: decimales que usa la moneda en la práctica.
// CLP y PYG son 0: tratarlas con centavos infla los montos por 100.

const MONEDAS = {
  // Latinoamérica
  ARS: { nom: 'Peso argentino',        sim: '$',    dec: 2, pais: 'Argentina' },
  BOB: { nom: 'Boliviano',             sim: 'Bs',   dec: 2, pais: 'Bolivia' },
  BRL: { nom: 'Real brasileño',        sim: 'R$',   dec: 2, pais: 'Brasil' },
  CLP: { nom: 'Peso chileno',          sim: '$',    dec: 0, pais: 'Chile' },
  COP: { nom: 'Peso colombiano',       sim: '$',    dec: 0, pais: 'Colombia' },
  CRC: { nom: 'Colón costarricense',   sim: '₡',    dec: 2, pais: 'Costa Rica' },
  CUP: { nom: 'Peso cubano',           sim: '$',    dec: 2, pais: 'Cuba' },
  DOP: { nom: 'Peso dominicano',       sim: 'RD$',  dec: 2, pais: 'República Dominicana' },
  GTQ: { nom: 'Quetzal',               sim: 'Q',    dec: 2, pais: 'Guatemala' },
  HNL: { nom: 'Lempira',               sim: 'L',    dec: 2, pais: 'Honduras' },
  MXN: { nom: 'Peso mexicano',         sim: '$',    dec: 2, pais: 'México' },
  NIO: { nom: 'Córdoba',               sim: 'C$',   dec: 2, pais: 'Nicaragua' },
  PAB: { nom: 'Balboa',                sim: 'B/.',  dec: 2, pais: 'Panamá' },
  PEN: { nom: 'Sol',                   sim: 'S/',   dec: 2, pais: 'Perú' },
  PYG: { nom: 'Guaraní',               sim: '₲',    dec: 0, pais: 'Paraguay' },
  UYU: { nom: 'Peso uruguayo',         sim: '$U',   dec: 2, pais: 'Uruguay' },
  VES: { nom: 'Bolívar',               sim: 'Bs.',  dec: 2, pais: 'Venezuela' },
  // Caribe y Centroamérica
  BZD: { nom: 'Dólar beliceño',        sim: 'BZ$',  dec: 2, pais: 'Belice' },
  GYD: { nom: 'Dólar guyanés',         sim: 'G$',   dec: 2, pais: 'Guyana' },
  SRD: { nom: 'Dólar surinamés',       sim: '$',    dec: 2, pais: 'Surinam' },
  TTD: { nom: 'Dólar trinitense',      sim: 'TT$',  dec: 2, pais: 'Trinidad y Tobago' },
  JMD: { nom: 'Dólar jamaiquino',      sim: 'J$',   dec: 2, pais: 'Jamaica' },
  HTG: { nom: 'Gourde',                sim: 'G',    dec: 2, pais: 'Haití' },
  BSD: { nom: 'Dólar bahameño',        sim: '$',    dec: 2, pais: 'Bahamas' },
  BBD: { nom: 'Dólar barbadense',      sim: '$',    dec: 2, pais: 'Barbados' },
  AWG: { nom: 'Florín arubeño',        sim: 'ƒ',    dec: 2, pais: 'Aruba' },
  XCD: { nom: 'Dólar del Caribe Or.',  sim: '$',    dec: 2, pais: 'Caribe Oriental' },
  // Usadas en la región
  USD: { nom: 'Dólar estadounidense',  sim: '$',    dec: 2, pais: 'Ecuador · Panamá · El Salvador' },
  // Resto del mundo
  EUR: { nom: 'Euro',                  sim: '€',    dec: 2, pais: 'Zona euro' },
  GBP: { nom: 'Libra esterlina',       sim: '£',    dec: 2, pais: 'Reino Unido' },
  CAD: { nom: 'Dólar canadiense',      sim: 'C$',   dec: 2, pais: 'Canadá' },
  CHF: { nom: 'Franco suizo',          sim: 'CHF',  dec: 2, pais: 'Suiza' },
  JPY: { nom: 'Yen',                   sim: '¥',    dec: 0, pais: 'Japón' },
  CNY: { nom: 'Yuan',                  sim: '¥',    dec: 2, pais: 'China' },
  AUD: { nom: 'Dólar australiano',     sim: 'A$',   dec: 2, pais: 'Australia' },
};

/**
 * Devuelve la ficha de una moneda. Una moneda desconocida no rompe nada:
 * cae en 2 decimales y usa su propio código como símbolo.
 */
function moneda(cod) {
  const c = String(cod || '').trim().toUpperCase();
  return MONEDAS[c] || { nom: c || 'desconocida', sim: c, dec: 2, pais: '', desconocida: true };
}

/**
 * Formatea un monto en su moneda, con convención latinoamericana:
 * punto para miles, coma para decimales. Ej: Q 1.250,50
 *
 * El separador de miles se aplica SOLO a la parte entera; aplicarlo a
 * la cadena completa produce "1.250.50", que ya no es un número.
 */
function fmtMoneda(monto, cod) {
  const m = moneda(cod);
  const n = aNumero(monto);
  if (n === '') return '';

  const signo = n < 0 ? '-' : '';
  const partes = Math.abs(n).toFixed(m.dec).split('.');
  const entera = partes[0].replace(/\B(?=(\d{3})+(?!\d))/g, '.');
  const dec = partes.length > 1 ? ',' + partes[1] : '';

  return signo + m.sim + ' ' + entera + dec;
}

// ─── TASAS ───────────────────────────────────────────────────

/**
 * Busca la tasa de un día concreto en la hoja `Tasas`.
 *
 * Si no hay tasa exacta para esa fecha, usa la más reciente ANTERIOR
 * (nunca una posterior: sería usar información del futuro). Si tampoco
 * hay ninguna anterior, devuelve null.
 *
 * @return {Object|null} { tasa, fecha_usada, exacta }
 */
function buscarTasa(ss, fecha, origen, destino) {
  origen = String(origen || '').toUpperCase();
  destino = String(destino || '').toUpperCase();
  if (!origen || !destino) return null;
  if (origen === destino) return { tasa: 1, fecha_usada: fecha, exacta: true };

  const sh = ss.getSheetByName('Tasas');
  if (!sh || sh.getLastRow() < 2) return null;

  const filas = sh.getDataRange().getValues().slice(1);
  let mejor = null;

  for (let i = 0; i < filas.length; i++) {
    const f = filas[i];
    const fFecha = aISO(f[0], 'UTC');
    const fOri = String(f[1] || '').toUpperCase();
    const fDes = String(f[2] || '').toUpperCase();
    const fTasa = aNumero(f[3]);
    if (!fFecha || fTasa === '' || fTasa <= 0) continue;

    let tasa = null;
    if (fOri === origen && fDes === destino) tasa = fTasa;
    else if (fOri === destino && fDes === origen) tasa = 1 / fTasa; // el par inverso sirve igual
    if (tasa === null) continue;

    if (fFecha === fecha) return { tasa: tasa, fecha_usada: fFecha, exacta: true };
    if (fFecha < fecha && (!mejor || fFecha > mejor.fecha_usada)) {
      mejor = { tasa: tasa, fecha_usada: fFecha, exacta: false };
    }
  }
  return mejor;
}

/**
 * Convierte un monto usando la tasa del día de la transacción.
 *
 * NO inventa un número cuando falta la tasa: devuelve valor null y un
 * motivo. Quien llama decide qué hacer, pero nunca recibe una cifra
 * inventada que se vea como buena.
 *
 * @return {Object} { valor, tasa, fecha_usada, exacta, motivo }
 */
function convertir(ss, monto, fecha, origen, destino) {
  const n = aNumero(monto);
  if (n === '') return { valor: null, motivo: 'monto no numérico' };

  const o = String(origen || '').toUpperCase();
  const d = String(destino || '').toUpperCase();
  if (!o || !d) return { valor: null, motivo: 'falta moneda origen o destino' };
  if (o === d) return { valor: n, tasa: 1, fecha_usada: fecha, exacta: true };

  const t = buscarTasa(ss, fecha, o, d);
  if (!t) {
    return {
      valor: null,
      motivo: 'sin tasa ' + o + '->' + d + ' para ' + fecha +
              ' ni ninguna anterior. Agrega la fila en la hoja Tasas.',
    };
  }

  const dec = moneda(d).dec;
  return {
    valor: Number((n * t.tasa).toFixed(dec)),
    tasa: t.tasa,
    fecha_usada: t.fecha_usada,
    exacta: t.exacta,
    motivo: t.exacta ? '' : 'tasa del ' + t.fecha_usada + ', no había del ' + fecha,
  };
}

/**
 * Revisa qué tasas hacen falta para poder convertir todo lo que ya está
 * cargado. Córrela antes de confiar en cualquier número de dinero.
 */
function tasasFaltantes(fileId, monedaDestino) {
  const ss = SpreadsheetApp.openById(fileId || IDS.empresarial);
  const destino = String(monedaDestino || 'COP').toUpperCase();
  const faltan = {};

  ['Pedidos', 'Pauta'].forEach(function (tab) {
    const sh = ss.getSheetByName(tab);
    if (!sh || sh.getLastRow() < 2) return;
    const datos = sh.getDataRange().getValues();
    const enc = datos[0].map(norm);
    const cFecha = enc.indexOf('fecha');
    const cMon = enc.indexOf('moneda_gasto') !== -1
      ? enc.indexOf('moneda_gasto') : enc.indexOf('moneda');
    if (cFecha === -1 || cMon === -1) return;

    datos.slice(1).forEach(function (f) {
      const fecha = aISO(f[cFecha], 'UTC');
      const mon = String(f[cMon] || '').toUpperCase();
      if (!fecha || !mon || mon === destino) return;
      if (!buscarTasa(ss, fecha, mon, destino)) {
        faltan[mon + ' -> ' + destino + '  desde ' + fecha] = true;
      }
    });
  });

  const lista = Object.keys(faltan).sort();
  const msg = lista.length
    ? 'FALTAN TASAS (' + lista.length + '):\n  ' + lista.join('\n  ') +
      '\n\nAgrégalas en la hoja Tasas: fecha · moneda_origen · moneda_destino · tasa'
    : 'No falta ninguna tasa para convertir a ' + destino + '. ✓';
  Logger.log(msg);
  return msg;
}
