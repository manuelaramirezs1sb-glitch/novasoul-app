# Conectar tu cuenta publicitaria a Nova

Para que Nova traiga sola las cifras de Meta todos los días, hay que
darle una llave de lectura. Se hace **una sola vez** y toma unos quince
minutos.

**Lo que Nova va a poder hacer con esa llave:** leer cuánto gastaste, en
qué conjunto y con qué resultados.

**Lo que NO va a poder hacer:** crear anuncios, pausarlos, cambiar
presupuestos ni gastar un peso. El permiso que se pide es de lectura y
nada más. Aunque la llave se perdiera, nadie podría tocar tu cuenta.

Puedes revocarla cuando quieras desde tu propio Business Manager, sin
avisarle a nadie.

---

## Antes de empezar

Necesitas ser **administradora de tu Business Manager** — el panel donde
vive tu cuenta publicitaria. Si tu cuenta de anuncios es personal y nunca
pasó por un Business Manager, primero hay que crear uno; no se puede
hacer lo de abajo sin él.

Ten a mano una pestaña con cada uno:

- `business.facebook.com` — tu Business Manager
- `developers.facebook.com` — donde se crea la app

---

## Paso 1 · Crear la app

Una "app" de Meta aquí no es un programa: es un registro que sirve para
emitir la llave. Se crea en dos minutos y no hace nada por su cuenta.

1. Entra a **developers.facebook.com** con tu cuenta de Facebook
2. Arriba a la derecha: **Mis aplicaciones** → **Crear aplicación**
3. Nombre: `Nova` (o el que quieras — solo lo ves tú)
4. Cuando pregunte qué quieres hacer, elige la opción de **negocio /
   empresa**
5. **Importante:** cuando pida vincular un Business Manager, **elige el
   tuyo**. Si lo dejas sin vincular, el resto no funciona.
6. Crear

> **Por qué la app la creas tú y no Nova.** Meta exige que la llave y la
> app vivan en el mismo Business Manager. Si la app fuera de Nova, Meta
> pediría una revisión formal que tarda semanas. Creándola tú, Meta ve a
> una dueña leyendo su propia cuenta con su propia app, y no hay nada
> que revisar. Es más rápido y te deja a ti el control.

---

## Paso 2 · Crear el usuario de sistema

Un "usuario de sistema" es un usuario que no es una persona. Existe para
que la conexión no dependa de que tú tengas la sesión abierta.

1. Entra a **business.facebook.com**
2. **Configuración del negocio** (el engranaje)
3. Menú izquierdo: **Usuarios** → **Usuarios del sistema**
4. **Agregar** → nombre: `Nova` → rol: **Empleado**
5. Crear

---

## Paso 3 · Darle acceso a tu cuenta publicitaria

Con el usuario `Nova` seleccionado:

1. **Agregar activos**
2. Pestaña **Cuentas publicitarias**
3. Marca tu cuenta
4. Enciende **solo** el permiso de **ver rendimiento**
   (lectura). **No** actives administrar campañas.
5. Guardar cambios

---

## Paso 4 · Generar la llave

Con el usuario `Nova` todavía seleccionado:

1. **Generar nuevo token**
2. Elige la app que creaste en el paso 1
3. En la lista de permisos marca **`ads_read`** — solo ese
4. Generar

**Copia la llave y guárdala de una vez.** Meta la muestra **una sola
vez**. Si cierras la ventana sin copiarla, toca generar otra (no pasa
nada malo, pero es volver a empezar este paso).

---

## Paso 5 · El número de la cuenta

Nova también necesita saber cuál cuenta leer.

En **Configuración del negocio** → **Cuentas publicitarias**, junto al
nombre de tu cuenta aparece un número largo. Ese es el identificador.
Cópialo. A veces se muestra con `act_` adelante; con o sin eso sirve.

---

## Paso 6 · Pegarlo en Nova

En Nova: **Configuración** → **Meta** → pegas la llave y el número de
cuenta → **Probar conexión**.

Nova hace una consulta de prueba ahí mismo y te dice qué encontró. Si
algo está mal, te dice **cuál** de los dos, no un error genérico.

A partir de ahí trae las cifras del día anterior todas las mañanas, sola.

---

## Si algo se atasca

Los menús de Meta cambian de nombre cada tanto y las traducciones al
español no siempre coinciden con lo de arriba. **Si una pantalla no se
parece a lo que dice esta guía, manda una captura antes de tocar nada.**

Errores comunes:

| Lo que pasa | Qué suele ser |
|---|---|
| La app no deja elegir Business Manager | La app se creó sin vincular. Se vincula después en la configuración de la app. |
| El token se genera pero Nova dice que no ve la cuenta | Falta el paso 3: el usuario de sistema no tiene la cuenta asignada. |
| Nova dice que el permiso no alcanza | En el paso 4 se marcó otro permiso, o ninguno. Genera otro token con `ads_read`. |
| Funcionaba y dejó de funcionar | Meta revoca llaves de apps que llevan 180 días sin usarse. Se genera otra con el paso 4. |
