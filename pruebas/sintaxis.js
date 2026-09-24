/**
 * QUE EL JAVASCRIPT DE CADA PANTALLA COMPILE. Nada más, y no es poco.
 *
 * ┌─ POR QUÉ EXISTE ESTE ARCHIVO ──────────────────────────────┐
 * │                                                            │
 * │ Una vez quedaron dos `let bandaActual` en empresarial.html. │
 * │ Un `let` repetido no rompe esa línea: rompe el ARCHIVO      │
 * │ ENTERO —el navegador no ejecuta ni una función del bloque—  │
 * │ y la pantalla queda en blanco sin un solo mensaje.          │
 * │                                                            │
 * │ Se descubrió de casualidad, días después, porque falló otra │
 * │ prueba que no tenía nada que ver. Esto lo atrapa en dos     │
 * │ segundos y sin navegador.                                   │
 * │                                                            │
 * │ No comprueba que la pantalla se vea bien ni que los datos   │
 * │ lleguen: para eso están las pruebas de pantalla. Comprueba  │
 * │ lo único que, si falla, deja todo lo demás sin sentido.     │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 */
const fs = require('fs');
const path = require('path');

const RAIZ = path.join(__dirname, '..');
const PANTALLAS = fs.readdirSync(RAIZ)
  .filter(function (f) { return f.endsWith('.html'); })
  .sort();

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };

console.log('\n── El JavaScript de cada pantalla compila ──');

PANTALLAS.forEach(function (archivo) {
  const h = fs.readFileSync(path.join(RAIZ, archivo), 'utf8');
  const re = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g;
  let m, bloques = 0, error = '';

  while ((m = re.exec(h))) {
    const js = m[1];
    if (!js.trim()) continue;
    bloques++;
    try {
      // `new Function` compila sin ejecutar: si hay un error de
      // sintaxis lo dice, y si no, no toca nada.
      new Function(js);
    } catch (e) {
      const linea = h.slice(0, m.index).split('\n').length;
      error = e.message + ' (el bloque empieza en la línea ' + linea + ')';
      break;
    }
  }

  if (!bloques) { console.log('  —      ' + archivo + ' (sin script)'); return; }
  ok(archivo + ' · ' + bloques + ' bloque(s)', !error, error);
});

/**
 * Y lo que de verdad pasó aquella vez: una declaración repetida en el
 * mismo ámbito. `new Function` ya la atrapa, pero se comprueba también
 * a la vista para que el día que falle, el mensaje diga QUÉ nombre está
 * repetido y no solo «Identifier has already been declared».
 */
console.log('\n── Ninguna declaración repetida en el ámbito de arriba ──');

PANTALLAS.forEach(function (archivo) {
  const h = fs.readFileSync(path.join(RAIZ, archivo), 'utf8');
  const re = /<script(?:\s[^>]*)?>([\s\S]*?)<\/script>/g;
  const vistos = {}, repes = [];
  let m;
  while ((m = re.exec(h))) {
    m[1].split('\n').forEach(function (l) {
      // Solo las del margen izquierdo: las de dentro de una función
      // están en otro ámbito y repetirlas es legal.
      const d = l.match(/^(let|const|var|function)\s+([A-Za-z_$][\w$]*)/);
      if (!d) return;
      const nombre = d[2];
      if (vistos[nombre] && (d[1] !== 'var' && d[1] !== 'function')) repes.push(nombre);
      vistos[nombre] = d[1];
    });
  }
  ok(archivo, repes.length === 0, 'repetidos: ' + repes.join(', '));
});

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien.\n');
process.exit(fallas ? 1 : 0);
