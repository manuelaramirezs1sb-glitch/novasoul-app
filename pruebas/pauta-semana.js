/**
 * EL GASTO DE PAUTA POR SEMANA.
 *
 * ┌─ QUÉ PIDIÓ ────────────────────────────────────────────────┐
 * │                                                            │
 * │ «la parte de gastos y pauta tampoco está dándome la         │
 * │  gráfica del gasto, cómo no va a sacar el gasto semanal si  │
 * │  ahí tiene los datos».                                      │
 * │                                                            │
 * │ El cuadro salía vacío con una explicación: el informe de    │
 * │ conjuntos de Meta trae UNA fila por todo el periodo y no    │
 * │ hay curva diaria que dibujar. Eso era cierto para el DÍA y  │
 * │ falso para la SEMANA, que es lo que ella pedía.            │
 * │                                                            │
 * └────────────────────────────────────────────────────────────┘
 *
 * La línea que separa calcular de inventar:
 *
 *  · un periodo dentro de una semana  → exacto, sin marca
 *  · un periodo que cruza semanas     → repartido por días, MARCADO
 *
 * Marcar no es un adorno: es lo que permite enseñar el número sin que
 * se lea como un dato del archivo.
 */
const fs = require('fs');
const src = fs.readFileSync(__dirname + '/../apps-script/NOVA-COMPLETO.gs', 'utf8');
(0, eval)(src.replace(/^function on(Open|Edit)/gm, 'function _no$1') +
  '\n;globalThis.__F = { semanasDePeriodo_, lunesDe_ };');
const F = globalThis.__F;

let fallas = 0;
const ok = (n, c, d) => { if (c) console.log('  ok     ' + n);
  else { fallas++; console.log('  FALLA  ' + n + (d !== undefined ? '\n         ' + d : '')); } };
const igual = (n, esp, real) => ok(n, JSON.stringify(esp) === JSON.stringify(real),
  'esperaba ' + JSON.stringify(esp) + ', obtuve ' + JSON.stringify(real));

console.log('\n── 1 · un periodo dentro de una semana es exacto ──');
// Miércoles 2026-09-23 a viernes 2026-09-25: la misma semana.
let r = F.semanasDePeriodo_('2026-09-23', '2026-09-25', 300);
igual('cae en una sola semana', 1, r.length);
igual('la del lunes 21', '2026-09-21', r[0].semana);
igual('con el gasto entero', 300, r[0].gasto);
igual('y NO se marca como repartido, porque no lo está', false, r[0].repartido);

console.log('\n── 2 · una fila de un solo día ──');
r = F.semanasDePeriodo_('2026-09-25', '', 50);
igual('una semana', 1, r.length);
igual('exacta', 50, r[0].gasto);
igual('sin marca', false, r[0].repartido);

console.log('\n── 3 · el periodo de su reporte: 21 ago a 25 sep ──');
r = F.semanasDePeriodo_('2026-08-21', '2026-09-25', 244);
ok('toca varias semanas', r.length > 4, String(r.length));
const suma = r.reduce((a, x) => a + x.gasto, 0);
ok('y la suma es exactamente el gasto, sin perder centavos',
   Math.abs(suma - 244) < 0.000001, String(suma));
ok('TODAS marcadas como repartidas', r.every(x => x.repartido),
   JSON.stringify(r.map(x => x.repartido)));
ok('en orden', r.map(x => x.semana).join() ===
   r.map(x => x.semana).sort().join(), JSON.stringify(r.map(x => x.semana)));
ok('todas empiezan en lunes',
   r.every(x => new Date(x.semana + 'T00:00:00Z').getUTCDay() === 1),
   JSON.stringify(r.map(x => x.semana)));

console.log('\n── 4 · el domingo cierra la semana, no la abre ──');
// Domingo 2026-09-27 pertenece a la semana del lunes 21.
igual('el domingo va con el lunes anterior', '2026-09-21',
      F.semanasDePeriodo_('2026-09-27', '', 10)[0].semana);
igual('y el lunes 28 abre la suya', '2026-09-28',
      F.semanasDePeriodo_('2026-09-28', '', 10)[0].semana);

console.log('\n── 5 · un periodo a caballo entre dos semanas ──');
// Domingo 27 sep (semana del 21) y lunes 28 (semana del 28): 2 días.
r = F.semanasDePeriodo_('2026-09-27', '2026-09-28', 100);
igual('dos semanas', 2, r.length);
igual('mitad y mitad', [50, 50], r.map(x => x.gasto));
ok('las dos marcadas', r.every(x => x.repartido));

console.log('\n── 6 · lo que no se puede calcular no se calcula ──');
igual('sin fecha, nada', [], F.semanasDePeriodo_('', '2026-09-25', 100));
// Una fila con un periodo de años es un archivo mal leído, no un gasto.
igual('un periodo absurdo no dibuja nada', [],
      F.semanasDePeriodo_('2020-01-01', '2026-09-25', 100));
// Una fecha final anterior a la inicial se trata como un solo día en vez
// de devolver un rango al revés.
r = F.semanasDePeriodo_('2026-09-25', '2026-09-01', 80);
igual('un rango invertido cuenta como un día', 1, r.length);
igual('con todo el gasto', 80, r[0].gasto);

console.log('\n── 7 · el gasto en cero no rompe nada ──');
r = F.semanasDePeriodo_('2026-09-21', '2026-09-30', 0);
ok('devuelve semanas', r.length >= 2, String(r.length));
igual('todas en cero', true, r.every(x => x.gasto === 0));

console.log(fallas ? '\n' + fallas + ' FALLAS\n' : '\nTodo bien\n');
process.exit(fallas ? 1 : 0);
