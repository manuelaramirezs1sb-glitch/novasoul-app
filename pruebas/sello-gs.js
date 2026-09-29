/**
 * EL SELLO DEL BACKEND.
 *
 * ┌─ POR QUÉ EXISTE ───────────────────────────────────────────┐
 * │                                                            │
 * │ Le mandé cuatro veces NOVA-COMPLETO.gs en una tarde. Todos  │
 * │ se llaman igual y ninguno decía cuál era, así que pegó el   │
 * │ tercero creyendo que era el cuarto y reportó como nuevo un  │
 * │ error que ya estaba arreglado.                              │
 * │                                                            │
 * │ Lo único que lo delató fue que el número de línea del error │
 * │ —23076, en un archivo que terminaba en 23145— no cuadraba   │
 * │ con el mío. Un diagnóstico que dependió de que yo me fijara │
 * │ en un número de línea no es un diagnóstico: es suerte.      │
 * │                                                            │
 * │ La PANTALLA lleva su sello desde hace meses, exactamente    │
 * │ por esto. El backend no, y lo pagó ella con su tiempo.      │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * ┌─ Y LO QUE ESTA PRUEBA VIGILA ──────────────────────────────┐
 * │                                                            │
 * │ Que el sello NO PUEDA MENTIR. Un sello que dice un número   │
 * │ y el archivo tiene otro es peor que no tener sello: se      │
 * │ compara, cuadra, y se descarta la hipótesis correcta.       │
 * │                                                            │
 * │ Por eso lo escribe `construir.py` y no una mano.            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const { execFileSync } = require('child_process');
const RUTA = __dirname + '/../apps-script/NOVA-COMPLETO.gs';

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

const texto = fs.readFileSync(RUTA, 'utf8');

console.log('\n══ 1 · EL ARCHIVO DICE QUÉ ES ══');
const m = texto.match(/^const NOVA_GS = '([^']+)';/m);
ok('trae el sello', !!m, texto.slice(0, 120));
if (!m) { console.log('\n1 FALLAS\n'); process.exit(1); }
const sello = m[1];
ok('con fecha, líneas y huella',
   /^\d{4}-\d{2}-\d{2} · \d+ líneas · [0-9a-f]{6}$/.test(sello), sello);

console.log('\n══ 2 · Y EL NÚMERO ES EL DE VERDAD ══');
/**
 * `count('\n') + 1` y no `wc -l`: el editor de Apps Script numera la
 * última línea vacía, así que este es el número que ella ve abajo del
 * todo. Es el que se compara, así que es el que tiene que cuadrar.
 */
const lineasReales = texto.split('\n').length;
const lineasSello = Number(sello.match(/· (\d+) líneas/)[1]);
igual('el sello dice las líneas que el archivo tiene', lineasReales, lineasSello);
ok('y es el número que enseña el editor de Apps Script abajo del todo',
   texto.endsWith('\n'), 'el archivo debería terminar en salto de línea');

console.log('\n══ 3 · NO LO ESCRIBE UNA MANO ══');
/**
 * Un sello puesto a mano miente el día que alguien olvida actualizarlo,
 * y ese día es el peor de todos: se compara, cuadra, y se descarta la
 * hipótesis correcta.
 */
ok('el sello no está en ningún archivo fuente',
   fs.readdirSync(__dirname + '/../apps-script')
     .filter((f) => /\.gs$/.test(f) && f !== 'NOVA-COMPLETO.gs')
     .every((f) => fs.readFileSync(__dirname + '/../apps-script/' + f, 'utf8')
       .indexOf('const NOVA_GS') === -1));

console.log('\n══ 4 · RECONSTRUIRLO SIN CAMBIOS DA EL MISMO SELLO ══');
/**
 * Si la huella cambiara sola en cada build, no serviría para comparar
 * dos archivos: todo parecería distinto siempre.
 */
execFileSync('python3', [__dirname + '/../apps-script/construir.py']);
const nuevo = fs.readFileSync(RUTA, 'utf8');
const m2 = nuevo.match(/^const NOVA_GS = '([^']+)';/m);
igual('mismo sello', sello.split(' · ').slice(1).join(' · '),
      m2 && m2[1].split(' · ').slice(1).join(' · '));
igual('y el archivo es idéntico', texto.length, nuevo.length);

console.log('\n══ 5 · CAMBIAR UNA LÍNEA CAMBIA LA HUELLA ══');
/**
 * Lo contrario también tiene que ser cierto, o el sello no distingue
 * dos versiones — que es justo para lo que se hizo.
 */
const FUENTE = __dirname + '/../apps-script/00-bootstrap.gs';
const original = fs.readFileSync(FUENTE, 'utf8');
try {
  fs.writeFileSync(FUENTE, original + '\n// una línea de más, a propósito\n');
  execFileSync('python3', [__dirname + '/../apps-script/construir.py']);
  const m3 = fs.readFileSync(RUTA, 'utf8').match(/^const NOVA_GS = '([^']+)';/m);
  const huella = (s) => s.split(' · ')[2];
  const lin = (s) => Number(s.match(/· (\d+) líneas/)[1]);
  ok('la huella es otra', huella(m3[1]) !== huella(sello),
     huella(sello) + ' → ' + huella(m3[1]));
  ok('y las líneas también', lin(m3[1]) > lin(sello),
     lin(sello) + ' → ' + lin(m3[1]));
} finally {
  // Pase lo que pase, la fuente vuelve como estaba: una prueba que deja
  // el repositorio sucio es una trampa para la siguiente.
  fs.writeFileSync(FUENTE, original);
  execFileSync('python3', [__dirname + '/../apps-script/construir.py']);
}
const m4 = fs.readFileSync(RUTA, 'utf8').match(/^const NOVA_GS = '([^']+)';/m);
igual('y al deshacer, vuelve el de antes', sello, m4 && m4[1]);

console.log('\n══ 6 · QUEVERSION() LO SABE DECIR ══');
/**
 * Es la función que ella corre desde el editor cuando algo no cuadra.
 * Si no existiera, la única forma de saber qué está publicado sería
 * mirar números de línea, que es como lo supimos esta vez.
 */
global.Logger = { log: () => {} };
global.PropertiesService = { getScriptProperties: () => ({
  getProperty: () => '', getProperties: () => ({}) }) };
global.SpreadsheetApp = { openById: () => null, flush: () => {} };
global.ScriptApp = { getProjectTriggers: () => [] };
global.Session = { getScriptTimeZone: () => 'UTC',
                   getActiveUser: () => ({ getEmail: () => '' }),
                   getEffectiveUser: () => ({ getEmail: () => '' }) };
global.LockService = { getScriptLock: () => ({ tryLock: () => true, releaseLock: () => {} }) };
global.CacheService = { getScriptCache: () => ({ get: () => null, put: () => {} }) };
global.UrlFetchApp = {};
global.MailApp = { sendEmail: () => {} };
global.Utilities = { sleep: () => {}, getUuid: () => 'u',
  formatDate: (d) => new Date(d).toISOString() };
(0, eval)(fs.readFileSync(RUTA, 'utf8') +
  '\n;globalThis.__V = { queVersion, NOVA_GS };');
const dicho = globalThis.__V.queVersion();
ok('devuelve el sello, no un texto vacío', dicho.indexOf(sello) !== -1, dicho);
ok('dice qué hacer si no coincide', /es otro/.test(dicho), dicho);
/**
 * «¿hay alguna diferencia si primero lo corro y luego lo implemento?»
 *
 * El botón «Ejecutar» corre lo GUARDADO; las pantallas hablan con
 * `/exec`, que sirve lo IMPLEMENTADO. Esta función solo sabe de lo
 * primero, y la primera versión del mensaje decía «lo que está
 * PUBLICADO es otro archivo» — justo la palabra que confunde las dos.
 * Un mensaje de diagnóstico que usa mal la palabra clave manda a
 * arreglar lo que no está roto.
 */
ok('deja claro que habla de lo GUARDADO, no de lo implementado',
   /GUARDADO/.test(dicho) && /NO dice qué versión están usando las pantallas/.test(dicho),
   dicho);
ok('y no usa «publicado», que es la palabra ambigua',
   !/publicad/i.test(dicho), dicho);

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
