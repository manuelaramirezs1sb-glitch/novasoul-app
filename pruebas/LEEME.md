# Pruebas

Se corren a mano, desde la raíz del repositorio. No se publican:
`.assetsignore` deja fuera esta carpeta entera.

```
node apps-script/pruebas/tasas.js     # lógica de tasas, sin abrir una hoja
node pruebas/central-automatico.js    # la tarjeta de lo automático, en Chromium
node pruebas/hub-puerta.js            # que el hub no se vea sin sesión
node pruebas/guia-meta.js             # que la guía renderice en ancho y en móvil
node pruebas/semaforo.js              # las reglas de cálculo del semáforo
node pruebas/semaforo-pantalla.js     # el semáforo en Hoy y en Pauta
node pruebas/meta-traer.js            # la lectura de Meta, con Meta remedado
node pruebas/meta-tarjeta.js          # la tarjeta de conexión
node pruebas/contar-llamadas.js       # cuántas llamadas le hace Nova a Meta
node pruebas/chat.js                  # el chat del equipo: quién lee qué
node pruebas/chat-pantalla.js         # el chat en la pantalla, en Chromium
node pruebas/accesos.js               # los accesos, y la contraseña fuera de la bitácora
node pruebas/accesos-pantalla.js      # los accesos donde van, y tapados
node pruebas/equilibrio.js            # el punto de equilibrio y su invariante
node pruebas/notas.js                 # que una nota no borre a la otra
node pruebas/pauta-semana.js          # el gasto por semana, y qué se reparte
node pruebas/arranque-empresarial.js  # cuántas peticiones cuesta abrir la pantalla
node pruebas/caches.js                # que guardar y volver a leer nunca dé lo viejo
node pruebas/puerta-tiendas.js        # que cualquier cliente pueda entrar, no solo Nutrea
node pruebas/primer-dia.js            # una cuenta recién nacida: una tienda, cero de todo
node pruebas/horario.js               # que la semana que Nova arma sea POSIBLE
node pruebas/hub-reintento.js         # la puerta, cuando Google contesta raro
node pruebas/cambio-tienda.js         # saltar de tienda sin que se invente nada
node pruebas/novedades-tienda.js      # una novedad es de UNA tienda, en las 5 puertas
node pruebas/family-tiendas.js        # las tiendas de NovaSoul, en chico y sin plata
node pruebas/cierre-previo.js         # qué falta para un cierre limpio (servidor)
node pruebas/cierre-revision.js       # revisar antes de cerrar (pantalla)
```

`primer-dia.js` no escribe a mano lo que contesta el servidor: lo CALCULA
con el bundle de verdad sobre una hoja recién creada, y se lo da a la
pantalla tal cual. Mi primer intento sí lo inventaba, y me dio tres
fallos que no existían —campos que mi simulador no mandaba y el servidor
sí— mientras tapaba el único que era real. Una prueba que se inventa la
respuesta prueba mi imaginación, no el producto.

`puerta-tiendas.js` existe porque Sara no pudo entrar a su propia cuenta:
el paso 3 del login tenía las dos tiendas de Nutrea escritas a mano en el
HTML y el código solo sabía esconderlas, nunca crear una. La puerta de
Nova solo se abría para una cuenta cuyas tiendas se llamaran `ec` y `gt`.
Y detrás había un segundo bloqueo: `POR_TIENDA[ST]` es `undefined` para
cualquier otra tienda, así que `buildAll()` reventaba al entrar.

`caches.js` es la contraparte de la velocidad. Desde que `libro_()`
guarda lo leído mientras dura la petición —lo que bajó el arranque de 75
lecturas de pestaña a 8— hay un error posible que antes no existía:
guardar algo y que la siguiente lectura devuelva lo de antes. En pantalla
eso se ve como «no se guardó», que es el peor síntoma porque invita a
guardar otra vez encima. Cada caso del archivo escribe de verdad por la
API y vuelve a leer por la API, en la misma ejecución.

Si una prueba cambia las pestañas POR DEBAJO (reemplazando los arreglos
del arnés en vez de escribir por la API), tiene que llamar a
`libroOlvidar_()`: está simulando a alguien editando el Google Sheet a
mano, que es algo que entre dos peticiones sí puede pasar.

Ajusté seis arneses así y me faltó `apps-script/pruebas/tasas.js`, que es
el único que vive en otra carpeta. Quedó rojo en `main` varios días: su
segundo escenario leía las pestañas del primero y `estadoTasasDe_` veía
una tienda donde había dos. **Una prueba roja en main no avisa de nada**
— hay que correr la lista entera, no la que se acaba de tocar.

`cambio-tienda.js` es la prueba que me enseñó a no creerle a mi propia
explicación. Yo estaba seguro de que la mezcla de tiendas venía de que
`cambiarTienda` limpiaba cinco de diecisiete variables. Lo arreglé,
escribí la prueba, y **al desactivar el arreglo a propósito siguió en
verde**: mi explicación no era la causa. Al imprimir la pantalla de
Ecuador en vez de razonar sobre ella apareció la de verdad —

    if (!r.ok || !r.filas.length) return false;

Cero pedidos se trataba como «no hay respuesta», no se pintaba nada, y
quedaba el HTML de la maqueta: seis pedidos con nombres inventados, dos
novedades con su transportadora, «3 gestoras · agosto 2026», «VENTAS
$ 68.420.000». Con el cartel verde «Datos reales de tu hoja» encima.

Por eso la prueba no busca variables sino TEXTOS DE LA MAQUETA: esos
nombres solo existen escritos a mano en el HTML, así que uno de ellos en
pantalla con la sesión abierta significa que un pintor se calló. Y por
eso rompe un cargador a propósito: los dieciocho iban en un
`Promise.all` desnudo y uno que se cayera sellaba la pantalla entera
como EJEMPLO, devolviendo las tablas falsas de golpe.

`equilibrio.js` afirma una propiedad, no un número:

    faltan > 0   ⟺   utilidad < 0

sobre 400 meses generados al azar con semilla fija. Es lo que hace
imposible que «ya cubriste tus gastos» y «estás perdiendo plata»
aparezcan juntos otra vez, como aparecían en la foto del 25 de
septiembre. Buscando ese invariante encontró un caso real que yo no
había pensado: con la contribución negativa decía «faltan 0».

`arranque-empresarial.js` mide peticiones de RED, no llamadas a `api()`.
La primera versión remedaba `api` —justo la función que sirve de la
despensa del arranque— y por eso contaba 20 cuando ya eran 1.

`accesos.js` tiene una aserción que sostiene una decisión y no un detalle:
que la contraseña de la plataforma NO aparezca en ninguna celda de
Movimientos. Es la razón entera de que los accesos tengan su propia hoja
en vez de ser un parámetro más — `apiParametros` registra cada cambio con
el valor viejo y el nuevo. Rompí el enmascarado a propósito para
comprobar que la prueba lo nota, y lo nota.

Las de navegador usan el Chromium de este entorno. En otro, se apunta con
`PLAYWRIGHT=/ruta/a/playwright` y, si hace falta, cambiando `executablePath`.

La de `tasas.js` no necesita navegador ni internet: remeda los servicios de
Google y le da hojas falsas al código real del bundle. Por eso corre en
menos de un segundo y se puede correr siempre.

`novedades-tienda.js` monta DOS tiendas, cada una con su novedad, y
pregunta por cada una a las cinco puertas que leen novedades. Monta dos
porque el error no era un filtro mal escrito: era que **no había
filtro**. `Novedades` era la única hoja con filas de una tienda y sin
columna para decir de cuál, y `apiListar` filtra por tienda solo si la
columna existe. Una prueba de una sola tienda pasa igual con filtro y
sin filtro — que es exactamente cómo esto vivió meses sin que nada
avisara, mezclando la bandeja de novedades, el KPI del mes, la alarma de
novedades sin gestionar, «qué quedó de ayer» y las novedades por gestora.

Tiene una sección para la ventana entre pegar el código y correr
`bootstrapTodo()`: en ese rato la columna todavía no existe y un filtro
que la exija deja la bandeja vacía en las dos tiendas, que es el error
contrario y molesta igual.

Y afirma que el relleno **no vuelve a escribir**, no lo que dice. Una
novedad huérfana —su pedido no está en la hoja— no se puede resolver
nunca, así que el contador de «faltan» no baja a cero: la primera
versión reescribía la columna entera de cuatro mil filas en cada
`bootstrapTodo()` para no cambiar una celda. Lo encontró la prueba al
correr el relleno dos veces seguidas.

`cierre-revision.js` y `cierre-previo.js` sostienen una decisión que fue
al revés de lo que ella pidió. Pidió **descargar un archivo** porque el
cierre «me saca la info en un cuadrito no más». El diagnóstico era
correcto; el archivo no, porque es un callejón sin salida para una tarea
que es ARREGLAR COSAS: se abre en otra parte, no tiene botones, y los
catorce pedidos siguen sin corregirse. La aserción que sostiene eso es
«se cambia el estado desde la revisión y las cifras de arriba se rehacen
con ese pedido dentro» — lo que un archivo no puede hacer.

La otra que importa: **tocar «Cerrar el mes» ya no cierra nada**, solo
abre la revisión. Antes era un `prompt` pidiendo AAAA-MM y un `confirm`
que decía cuántos pedidos faltaban y nunca cuáles. Cerrar congela las
cifras para siempre y solo se reabre editando la hoja a mano, así que un
dedo torcido en ese `prompt` costaba caro.

`family-tiendas.js` mide el ANCHO y el COLOR en pantalla, no en el CSS.
Las dos primeras versiones estaban mal y el CSS se veía razonable en las
dos: las tarjetas salían de 570px —«más pequeños», había pedido ella—
porque `flex:1 1 150px` con dos tiendas las estira a media pantalla; y el
verde salía DORADO, porque lo pinté con `var(--acc)`, que cambia con el
tema. Un semáforo donde «bien» y «regular» son del mismo color no es un
semáforo.

Tiene además la aserción que protege una regla suya —«dentro de Nova
nunca se comparan dos tiendas al mismo tiempo en una misma pantalla»—
afinada: lo que la regla protegía era comparar CIFRAS, no ver estados.
Las tarjetas no llevan dinero, y si alguien le mete una cifra, falla.

Una advertencia de arnés que costó una aserción falsa: **si la vista está
oculta, `getBoundingClientRect()` devuelve 0** y un «ninguna pasa de
240px» pasa por la razón equivocada. Por eso hay también un mínimo.

Y una de shell: el contador de rojas tiene que ir FUERA de cualquier
`$( )`. Dentro es una subshell, `fail=1` no sale de ella, y casi reporto
«todas verdes» con dos rojas.

`contar-llamadas.js` no afirma ni falla: imprime. Existe porque una vez
afirmé de memoria que traer un año de anuncios «llenaría la cuota de la
hora», puse un tope de 90 días por esa razón, y al contarlo resultó
falso — eran decenas de llamadas contra un presupuesto de miles. Lo que
sí estaba roto era otra cosa que el tope tapaba. Cuando haya una duda de
volumen, se corre en vez de estimarla.
