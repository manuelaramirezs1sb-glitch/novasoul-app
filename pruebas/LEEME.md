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
```

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
