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
```

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
