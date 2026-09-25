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
```

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

`contar-llamadas.js` no afirma ni falla: imprime. Existe porque una vez
afirmé de memoria que traer un año de anuncios «llenaría la cuota de la
hora», puse un tope de 90 días por esa razón, y al contarlo resultó
falso — eran decenas de llamadas contra un presupuesto de miles. Lo que
sí estaba roto era otra cosa que el tope tapaba. Cuando haya una duda de
volumen, se corre en vez de estimarla.
