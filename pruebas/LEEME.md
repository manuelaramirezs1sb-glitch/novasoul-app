# Pruebas

Se corren a mano, desde la raíz del repositorio. No se publican:
`.assetsignore` deja fuera esta carpeta entera.

```
node apps-script/pruebas/tasas.js     # lógica de tasas, sin abrir una hoja
node pruebas/central-automatico.js    # la tarjeta de lo automático, en Chromium
node pruebas/guia-meta.js             # que la guía renderice en ancho y en móvil
```

Las de navegador usan el Chromium de este entorno. En otro, se apunta con
`PLAYWRIGHT=/ruta/a/playwright` y, si hace falta, cambiando `executablePath`.

La de `tasas.js` no necesita navegador ni internet: remeda los servicios de
Google y le da hojas falsas al código real del bundle. Por eso corre en
menos de un segundo y se puede correr siempre.
