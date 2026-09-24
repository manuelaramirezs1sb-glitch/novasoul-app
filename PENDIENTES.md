# Nova · lo que sigue

Lo acordado, en orden. Se escribe aquí y no en el chat porque un acuerdo
que solo vive en una conversación se pierde con la conversación.

---

## 1 · Mudanza a Cloudflare — HECHO (21-09-2026)

**En vivo:** https://nova.novasoul959.workers.dev

| | |
|---|---|
| Hub | `/` |
| Empresarial | `/empresarial.html` |
| Central | `/novacentral.html` |
| Demo | `/empresarial.html?demo=1` |

Cloudflare unificó Pages dentro de Workers, así que el proyecto quedó
como Worker de solo archivos: sin `main`, sin código de servidor. El
servidor de Nova sigue siendo Apps Script.

Dos cosas que salieron de esto y no estaban previstas:

- `wrangler.jsonc` hacía falta o el despliegue no arrancaba.
- **`apps-script/` estaba publicado en Netlify desde siempre.** Nadie lo
  había mirado. No había claves (los identificadores viven en las
  Propiedades del Script), pero era el producto entero a un clic de quien
  adivinara la dirección. `.assetsignore` lo cerró — verificado: da 404.

Netlify se puede dar de baja cuando se quiera.

<details>
<summary>Por qué Cloudflare y no las otras</summary>



**Por qué.** Netlify se quedó sin cuota. Cloudflare Pages da ancho de
banda ilimitado, **permite uso comercial** en el plan gratis y funciona
con el repositorio **privado** — las tres cosas que Nova necesita para
poder venderse.

Se descartaron:

- **Vercel** — su plan gratis (Hobby) prohíbe el uso comercial en los
  términos. No es un límite técnico que se pueda esquivar: en el momento
  en que Nova le cobre a un cliente, toca Pro a 20 USD por persona.
- **GitHub Pages** — gratis solo con el repositorio público. Publicar el
  código de un producto que se va a vender es una decisión que no hay
  por qué tomar cuando existe una alternativa que no la exige.
- **Neon** — no es hosting, es una base de datos PostgreSQL. Serviría si
  algún día Nova se muda fuera de Google Sheets, no para servir páginas.

</details>

---

## 2 · Meta conectado — LA LLAVE YA ESTÁ (21-09-2026)

Nutrea EC conectada con la cuenta «CP Nutrea Ecuador» por el camino B:
app y usuario de sistema creados en el portfolio de Manuela, permiso de
activo "ver rendimiento", token con `ads_read` únicamente.

**Ojo con la moneda.** Meta le cobra en COP y esa tienda factura en USD.
Sin la tasa del día de cada gasto ese gasto NO se suma: se cuenta aparte
como "sin convertir" y la pantalla lo dice. Eso es lo correcto —es la
misma protección que evitó el margen de −181.817% de antes— pero deja el
CPA y el margen incompletos hasta que haya tasas.

### Lo automático ya se prende con una sola función

`prenderAutomatico()` en el Apps Script, una vez por cuenta. Instala los
dos disparadores diarios (tasas 6 a.m., alarmas 7 a.m.) y carga de una
vez los últimos 90 días de tasas. `verAutomatico()` dice qué quedó
prendido y, aparte, si las tasas están de verdad al día.

Dos cosas que había que arreglar para que eso sirviera:

- **Los disparadores nunca se habían prendido.** `instalarTriggerTasas()`
  e `instalarTriggerAlarmas()` existían desde siempre, cada uno con su
  instalador aparte que había que correr a mano. Nadie los corrió. Eso
  es peor que estar roto: *parecía* automático. La hoja Tasas vacía no da
  error, y quien conecta Meta se queda esperando una conversión que no
  va a llegar.
- **`paresEnUso` no miraba dónde estaba el problema.** Decidía qué tasas
  traer comparando la moneda de cada tienda con la del reporte — nunca
  la moneda en que Meta cobra. A Nutrea le funcionaba de rebote (reporta
  en pesos, le cobran en pesos, y el par USD→COP sirve invertido); a una
  tienda que facture y reporte en la misma moneda y le cobren en otra le
  habría respondido "no hacen falta tasas". Ahora `paresDeGasto()` lee la
  moneda real del gasto en Pauta contra la moneda de esa tienda, y
  `paresNecesarios()` junta las dos listas sin repetir inversos.
  Probado en `apps-script/pruebas/tasas.js`.

### La lectura diaria — HECHA (22-09-2026)

`76-meta-leer.gs`. Nova le pide a Meta el gasto y lo escribe en Pauta:
sola cada mañana a las 6, o con el botón **Traer ahora** de la pantalla
de la dueña, que además sirve para traer el historial la primera vez
(hasta 400 días).

Se pide a nivel de **conjunto** y con `time_increment=1` —una fila por
día— porque el export manual trae una sola fila por todo el periodo, y
repartirla entre los días dibuja una curva que nunca existió.

Tres cosas que dejan de ser un problema al pedirlo así: la moneda la
declara Meta (`account_currency`) en vez de adivinarse por el nombre de
una columna; las fechas vienen en ISO, sin el 03/04 que es marzo o abril
según el país; y el gasto entra como número, sin el punto y la coma
latinoamericanos.

**Lo que no se inventa.** Meta no entrega presupuesto ni estado de
entrega en este informe: esas columnas quedan vacías, no en cero. Y si
los conjuntos no reportan compras sino registros, se cuenta el registro
y **se dice** — ese CPA no es por venta, y quien lo mire tiene que
saberlo.

**Volver a leer no duplica.** El identificador de cada fila es el mismo
que arma el importador del Excel, así que traer dos veces el mismo día
corrige la fila en vez de sumar otra. Eso también importa por una razón
menos obvia: Meta sigue atribuyendo conversiones días después, así que
las cifras de ayer cambian pasado mañana. Se vuelven a pedir siete días
cada mañana y se reescriben encima.

**La única fuga posible, y está tapada.** Una fila subida a mano que
cubre un RANGO de varios días no comparte identificador con las diarias,
así que se sumaría además de ellas. Nova las detecta y las cuenta en el
informe — pero no las borra: borrar datos que alguien subió es decisión
suya.

Probado en `pruebas/meta-traer.js` con Meta remedado: 27 casos, entre
ellos la doble lectura, la corrección de conversiones, la respuesta
partida en páginas, el error de llave vencida y la fila que se solapa.

### Lo que se aprendió montándolo

- La opción de crear la app sale EN GRIS hasta que pasas el mouse por
  encima y aparece "Confirmar cuenta". No lo dice ninguna documentación.
- El usuario de sistema necesita rol sobre la app —"Desarrollar la
  aplicación"— o la lista de permisos sale vacía sin explicar por qué.
- Publicar una versión nueva del Apps Script mata todas las sesiones
  abiertas. Conviene publicar cuando nadie esté gestionando pedidos.

---

## 2b · Meta, lo que sigue

Que Nova le pida las cifras a Meta todos los días, en vez de que alguien
baje un Excel.

### Cómo, y por qué así

La regla de Meta: **la llave (token) y la app tienen que vivir en el
mismo Business Manager.** De ahí salen dos caminos:

| | Quién crea la app | ¿App Review? | Tiempo |
|---|---|---|---|
| **A · una app de Nova** | Nova | sí — documentos, video, espera | semanas |
| **B · cada cliente la suya** | el cliente, en su BM | **no** | el mismo día |

Se arranca por **B**. Meta ve a un dueño leyendo su propia cuenta con su
propia app: no hay tercero que revisar. Los límites de ese nivel
(*Limited Access*: 300 llamadas/hora + 40 por anuncio activo) sobran para
una lectura diaria.

Y **B no es un callejón sin salida**: para dar el acceso completo, Meta
pide historial real de uso de la API. Cada cliente conectado por token
lo va construyendo. B es la rampa hacia A, no su sustituto.

### Reglas que no se negocian

- **Solo `ads_read`.** Permiso de lectura. Si una llave se filtra, nadie
  puede gastar la plata de nadie.
- **Las llaves van en las Propiedades del Script, nunca en una celda.**
  Quien abra la hoja no debe poder verlas.
- **Un token muerto se dice.** Meta revoca tokens de apps inactivas a los
  180 días. Mostrar la pauta en cero como si no se hubiera gastado sería
  mentir con forma de dato.

### Falta

- Guía con capturas, paso por paso, para el cliente. Son ~10 pasos entre
  `developers.facebook.com` y `business.facebook.com`, y varios se van a
  perder. A algunos habrá que acompañarlos por videollamada.

---

## 3 · Análisis de pauta contra el techo

Lo que Meta no puede hacer, y Nova sí.

**El argumento.** Meta cuenta una compra cuando alguien confirma en el
checkout. En contraentrega, ~30 de cada 100 nunca se entregan: el ROAS
que muestra Meta está inflado y con él se deciden presupuestos. Nova sabe
cuáles llegaron. Y sabe el **techo** — cuánto se puede pagar por entrega
después del producto, el flete, las devoluciones que arrastra cada venta
y el 3% del retiro. Ese número no lo tiene nadie más.

```
"El conjunto de retargeting está a $6,80 por entrega. Tu techo es $9,24.
 Le quedan $2,44 de margen y solo se lleva el 28% del presupuesto."

"El conjunto frío entrega al 52%, doce puntos bajo tu promedio. Su CPA en
 Meta se ve bien ($7,10) pero por entrega real cuesta $13,65 — sobre tu
 techo. Cada dólar ahí se pierde."
```

### Las tres preguntas

Son decisiones, no datos: ningún historial las revela.

1. **¿Qué quieres de este mes?** crecer aunque apriete la caja ·
   sostener el margen · recuperar caja
2. **¿Cuánto puedes perder probando al mes, sin que te duela?**
3. **¿Cuántos días tarda tu proveedor en reponer?**

### Lo que Nova deduce sola

Se pregunta solo mientras no haya con qué deducirlo, y se le muestra al
cliente lo deducido para que lo corrija.

4. **Cuántos pedidos al día aguanta el equipo** — de los días de más
   volumen, viendo si ahí se cayó la tasa de entrega. Es la que más plata
   salva: el error clásico es subir 40% el presupuesto, que el equipo no
   alcance a confirmar, que la entrega caiga ocho puntos, y que el
   escalado se coma a sí mismo.
5. **En qué punto está cada producto** — por la fecha de la primera venta
   y hacia dónde vienen el CPA y los pedidos.

### Las dos reglas que lo hacen fiable

- **Cada consejo muestra la cuenta.** Nunca "escala este conjunto", sino
  la aritmética que lleva ahí. Un consejo que no se puede auditar no se
  puede corregir.
- **Callarse cuando no hay con qué.** Un conjunto con 6 entregas no dice
  nada: el azar manda sobre esa muestra. Por debajo del mínimo, Nova dice
  cuántas entregas faltan en vez de opinar. Va a ser la respuesta más
  frecuente el primer mes, y está bien.

Esto último es lo que protege el cobro por uso: el primer consejo seguro
que le haga perder plata a un cliente cuesta el cliente.

---

## 4 · Una IA que redacte los consejos — MÁS ADELANTE

Anotado a pedido de Manuela, explícitamente **para después**.

**Qué sería.** Un modelo de lenguaje que, sobre los números ya
calculados, escriba el consejo en prosa y proponga metodología de pauta.

**Por qué no ahora.** Dos razones, y la segunda pesa más:

1. **Cuesta por consulta.** Todo lo demás de esta lista es aritmética en
   Apps Script: $0. Esto no.
2. **Un modelo responde con seguridad aunque los datos no alcancen.** Es
   exactamente lo contrario de la regla de callarse. Meterlo antes de que
   la capa determinista esté firme sería construir al revés.

### Y NovaBot, que ya existe a medias

Hoy vive solo en `novacademy.html`: un panel de chat con preguntas y
respuestas ESCRITAS A MANO (`BOT_QS`, `BOT_R`), emparejadas por palabra
clave. No lee ninguna hoja, no hay servidor y no hay modelo. Responde
sobre el programa de la academia, nada más.

Conectarlo a los datos de Meta es posible y hay dos caminos, muy
distintos de precio:

| | Qué contesta | Cuesta |
|---|---|---|
| **Determinista** | una lista fija de preguntas, con las cifras reales de las hojas | $0 |
| **Con modelo** | cualquier cosa, en prosa | por consulta |

El determinista se puede hacer ya: «¿cuánto gasté esta semana?»,
«¿cuál anuncio va mejor?», «¿estoy sobre mi techo?» — Nova ya tiene
todas esas cifras calculadas, el bot solo las leería.

**Pero conviene decir lo que no cambia.** El semáforo del lunes ya
responde «qué hago con la pauta» sin que nadie pregunte. Un bot es otra
forma de PREGUNTAR lo mismo, no una capacidad nueva. Y la pregunta que
importa —«¿qué hago?»— es justo la que no hay que tener que hacer.

Por eso va después de NovaSoul, y el determinista antes que el modelo.

**Cuándo tendría sentido.** Cuando el análisis del punto 3 lleve meses
funcionando y el límite sea la redacción, no el cálculo. Entonces el
modelo se apoya sobre números que ya están bien, que es el orden
correcto.

---

## 5 · Lo que quedó pendiente de antes

- **Nova Central** — Dashboard y Demo siguen en ejemplo. Equipo depende
  de resolver las comisiones (cada dueño configura por persona, con un
  estimado sugerido, no una regla fija).
- **NovaSoul y novAcademy** — maquetas completas, cero servidor. Ningún
  dato sale de ninguna parte; los números están escritos en el HTML.
- **Información de producto desde Dropi** — Manuela ya la tiene escrita
  en la plataforma. Falta el export para importarla en vez de que la
  vuelva a escribir.

---

## 6 · NovaSoul y Nova Central, como de verdad son (22-09-2026)

Manuela describió para qué las necesita. No es lo que decían las maquetas,
así que se escribe aquí antes de construir nada.

### Lo que hay encima de ella ahora mismo

| Compromiso | Qué forma tiene | Quién pone la fecha |
|---|---|---|
| **PHH** (EE. UU., por Upwork) | tareas sueltas, por horas | ellos |
| **Salsabor** — mesera | turnos fijos | el local |
| **Son de Sky** — carta, web, cocteles, precios, marketing | proyecto con entregables y precio | acordada |
| **Nova** — Empresarial hoy, Academy pronto | producto propio, sin fecha externa | ella |
| **Universidad** + prácticas el año entrante | entregas con fecha inamovible | la U |

Cinco compromisos sobre las mismas veinticuatro horas. **El problema no
es recordar: es repartir.** Una lista de tareas no lo resuelve — lo
esconde, porque una lista siempre cabe en una pantalla aunque no quepa
en una semana.

### El reparto

Nova Central y NovaSoul comparten un solo objeto, el **proyecto**, y lo
miran por lados distintos:

**Nova Central — el negocio**
Contratos y su especificación · cotizaciones · qué se cobró, qué falta
cobrar y de quién · finanzas de Nova (Empresarial, Academy) · **y sus
finanzas personales**. Es la pantalla que algún día se le muestra a una
socia o a un contador.

**NovaSoul — el día a día**
Los entregables de esos proyectos con su fecha · tareas de la
universidad · **límites y riesgo de entrega** · cumplirse a sí misma ·
comidas, ejercicio, astrología, tranquilidad.

### La frontera, que es de privacidad y no de función

**De Soul a Central sube solo lo que es trabajo**: entregas, fechas,
plata. Las comidas, el ciclo y el ánimo no salen de Soul.

Es de una sola vía a propósito. Central se abre delante de otra gente;
Soul no se abre delante de nadie. Si se mezclan una vez, ya no se puede
abrir Central sin pensarlo.

### Lo que hay que construir primero, y por qué

El **riesgo de entrega**. Es el mismo argumento del semáforo de pauta
aplicado a ella: decirle **antes** que la semana no cabe, no después de
que incumplió.

Para eso hacen falta tres cosas y ninguna se puede adivinar:

1. **Cuántas horas tiene libres cada día**, descontados turnos y clases.
   Es el equivalente del techo de CPA: sin ese número no hay con qué
   comparar.
2. **Cuánto cree que cuesta cada entregable.** Estimado por ella, y
   corregido con lo que de verdad tardó.
3. **Qué pasa si no se entrega.** No es lo mismo un parcial que un
   ajuste de carta que se puede correr una semana.

Con eso, Nova puede decir: *«esta semana tienes 22 horas libres y 31
comprometidas. Algo se va a caer; estas son las tres candidatas y esto
cuesta cada una.»* Eso sí sirve. Una lista de pendientes no.

**Y la misma regla de siempre:** si no hay con qué calcularlo, se dice
que falta — no se dibuja una barra de progreso inventada.

### 6b · La astrología no es adorno: es el criterio de CUÁNDO

Manuela lo dijo claro y cambia el diseño: NovaSoul no es un calendario
con horóscopo al lado. La carta y los tránsitos son **la regla que decide
qué va en qué día**.

> «Los planetas marcan el tiempo: hay tiempo para aprender, tiempo para
> descansar y tiempo para cambiar.»

Así que el repartidor de la semana no ordena solo por fecha e
importancia. Ordena por fecha, importancia **y ventana** — y la ventana
sale de su carta, de los tránsitos vigentes y de su ciclo.

Piezas:

- **Su carta natal**, cargada una vez. Es fija.
- **Los tránsitos**, con su DURACIÓN: no es lo mismo la Luna, que dura
  dos días y medio, que Saturno, que dura meses. Un consejo que no
  distingue eso es ruido.
- **Su ciclo menstrual**, que es fisiología medible y tiene su propio
  patrón de energía.
- **Las recompensas**, que son suyas y las define ella.

**Lo que Nova puede afirmar y lo que no.** Nova puede decir «hoy tienes
Luna en tu casa 10, que en tu marco es día de empujar lo público» —eso
es aritmética sobre su carta, y es verdad dentro de su marco. Lo que
Nova NO va a hacer es afirmar que por eso va a salir bien.

Lo que sí va a hacer, y no hace ninguna app de astrología: **medir si le
funcionó a ella**. Cada tarea queda con la ventana en que se hizo y con
si se cumplió. A los tres meses Nova puede decirle «las entregas que
pusiste en días de Luna menguante las terminaste el 80% de las veces; en
Mercurio retrógrado, el 40%». Eso es su propio dato, no una creencia
prestada — y si el patrón no aparece, también se lo dice.

Es la misma regla de siempre: se muestra la cuenta, y se calla cuando no
hay muestra.

### 6c · Los accesos, decididos (22-09-2026)

| Pantalla | Quién entra | Cómo |
|---|---|---|
| **Central** | solo Manuela | su propio código, lista Plataforma |
| **Hub** | desde Central | sin volver a pedir código |
| **Empresarial** | Manuela y clientes | correo + código, igual para todos |
| **NovaSoul** | solo Manuela | **misma sesión de Central**, sin segundo código |
| **novAcademy** | estudiantes y docentes | su propio usuario, que ella genera |
| novAcademy · panel de ella | solo Manuela | misma sesión de Central |

**Por qué NovaSoul NO puede ser «sin login».** Manuela pidió que no le
pida contraseña. Eso está bien y así queda — pero una página sin sesión
en un sitio público la ve cualquiera que escriba la dirección, y ahí
adentro están su ciclo, sus finanzas personales y su ánimo. Lo que
resuelve el pedido sin dejar la puerta abierta es **compartir la sesión
de Central**: ella no escribe ningún código, y quien no tenga esa sesión
no ve nada. Es lo mismo que ya se hizo entre el hub y Empresarial.

---

## 7 · Lo que entró el 22-09-2026 por la tarde

### La paleta de Nova Central — HECHA

Manuela: «no me gustan esos colores, no tiene el modo claro, muy oscuro
para mi gusto». Tenía razón a medias y la mitad importa: **el modo claro
existía, pero Central arrancaba en oscuro y había que pedirlo.** Nadie va
a buscar el interruptor de algo que no sabe que está apagado.

Ahora el claro es la casa y el oscuro la excepción que se marca.

La paleta nueva sale del tema de la pantalla, que es plata: **un libro de
cuentas**. Papel tibio (`#F3F1E9`), tinta (`#23272B`), el verde profundo
de las rayas contables como acento (`#2F5E4E`) y latón para lo que hay
que mirar (`#B8873B`).

Descartados y por qué: el rosa del tablero que mandó de referencia —dijo
que «tampoco me mata»—; el azul marino, que ya había descartado para
NovaSoul; y el verde amarillento de Nova Empresarial, que es otra
pantalla y no deben confundirse.

**El rojo y el verde de dinero van aparte del acento.** Si el acento
fuera verde y lo positivo también, una cifra buena no se distinguiría de
un botón. Y lo ATRASADO va en rojo: una cifra que pide llamar hoy no
puede verse igual que una que está bien.

### PHH es confidencial — queda por escrito

Manuela: «los de PHH normalmente son confidenciales, entonces para que
dejes eso por escrito, y sepas que lo comparto con Nova solo por
organización».

Queda dicho y manda sobre lo que se construya:

- Lo de PHH se guarda **solo para organizarse**: qué hay que entregar y
  cuándo. No es material para analizar, ni para entrenar nada, ni para
  mostrar en ninguna pantalla compartida.
- Vive en **NovaSoul**, que es la que no se abre delante de nadie. A
  Central sube únicamente que existe un compromiso y cuánto ocupa — el
  nombre del trabajo y sus horas. **El contenido no sube.**
- Los documentos que suba de PHH no se copian a ninguna otra hoja ni se
  mandan a ningún servicio externo.

### Mindlab: su meta más alta ahora mismo

Un programa de 3 meses y 12 semanas para monetizar su marca personal.
NovaSoul tiene que tenerlo presente: es la meta contra la que se mide si
una semana sirvió.

| Mes | | Semanas |
|---|---|---|
| 1 · **Tráfico** | atraer a las personas correctas | 01 Fundamentos de marca · 02 Investigación y estrategia · 03 Producción · 04 Edición |
| 2 · **Nutrición** | convertir seguidores en audiencia | 05 Historias que conectan · 06 Ecosistema de confianza · 07 Sistemas de conversión · 08 Ventas desde contenido |
| 3 · **Venta** | monetizar la marca | 09 Diseño de oferta · 10 Creación de producto · 11 Infraestructura · 12 Landing pages |

Cada semana trae UNA tarea concreta —«optimizar perfil y definir
posicionamiento», «crear 30 ideas y 15 guiones», «grabar entre 8 y 12
piezas», «lanzar tu sistema de adquisición»— y esas doce tareas son las
que entran a NovaSoul como entregables con fecha.

**Por qué importa para el diseño.** Sin esto, Mindlab compite en igualdad
con Salsabor y con PHH por sus horas, y pierde siempre: no tiene cliente
que reclame. Marcarla como meta es lo que permite que NovaSoul diga «esta
semana no tocaste lo único que es tuyo».

### 7b · Central se organiza por PROYECTO (22-09-2026)

Era una lista plana —Dashboard, Clientes, Equipo, Pipeline, Onboarding,
Meta, Demo— que describía las partes de Nova. Pero **Nova es uno de los
proyectos de Manuela**, y los demás no tenían dónde caer: la carta de un
bar, un encargo de PHH, un parcial.

Ahora el primer nivel son sus proyectos, y Nova Family es el primero de
ellos. Lo que antes era el menú entero vive adentro de ese grupo.

```
Hoy
Mis finanzas
── MIS PROYECTOS ──   Todos + cada proyecto, con su punto de urgencia
── NOVA FAMILY ──     Clientes · Nuevo cliente · Equipo · Pipeline ·
                      Meta Ads · Demo para clientes
── IR A ──            NovaSoul ↗ · Nova Empresarial ↗ · novAcademy ↗
```

**Hoy** toma la forma del tablero que mandó de referencia —cifras
arriba, repartición a la izquierda, el centro, y lo que se debe a la
derecha— con sus temas en vez de los del ejemplo.

Dos decisiones de esa pantalla:

- **La repartición es por HORAS, no por plata.** La plata dice qué paga;
  las horas dicen qué ocupa, y son cosas distintas — la universidad no
  paga nada y se lleva un pedazo grande. Repartir por plata la haría
  desaparecer justo del sitio donde hay que verla.
- **La semana solo dibuja entregas.** Las clases y los turnos no están
  en ninguna hoja todavía, así que no se pintan, y la pantalla dice por
  qué: dibujarlos sin saberlos sería inventarle la semana.

### Los proyectos, y qué falta de cada uno

| Proyecto | Artefacto | Qué falta |
|---|---|---|
| **Nova Family** | las cuatro pantallas | NovaSoul entera |
| **Carta Skyblue** | ella ya lo tiene | volverlo PLANTILLA: costeo de cartas nuevas con fotos de carta, costos de licores y recetarios |
| **Marketing** | no existe | cuánto cobrar por pieza de pauta en FB/TikTok |
| **PHH** | no existe | lo crea cuando le asignen el primero |
| **Universidad** | uno por materia | conectar horarios de clase |

**El Demo cambia de significado y hay que decidirlo.** Hoy «Demo» es
mostrarle Nova Empresarial a un cliente potencial. Manuela quiere que
abra la plantilla de trabajo del proyecto. Son dos cosas distintas con
el mismo nombre: por ahora el de clientes se llama «Demo para clientes»
y el otro será «Plantilla» dentro de cada proyecto.

---

## 8 · Correcciones y encargos del 22-09-2026 (noche)

### NovaBot y NovaChat son DOS cosas, y yo las mezclé

Corrección de Manuela, y es de fondo:

| | Quién lo usa | Para qué |
|---|---|---|
| **NovaChat** | la dueña y su equipo, adentro | hablar entre ellos |
| **NovaBot** | la dueña | leer, analizar y avisar |

Lo que hay hoy en Empresarial es un solo panel que hace de las dos cosas
a medias. Se separan.

**NovaChat** vuelve a ser lo que era: el chat interno del equipo.

**NovaBot** tiene pantalla propia, y ahí vive todo lo de la operación
que hoy está regado:

- El **enlace a la plataforma** de esa tienda —Dropi, Effi, Mastershop,
  la que sea— para entrar a confirmar y gestionar pedidos.
- Las **condiciones de la tienda**: lo que el equipo tiene que saber de
  memoria y hoy pregunta cada vez.
- El **enlace al chat** de la tienda.
- Las **plantillas** para los casos de siempre. Si el cliente ya las
  tiene en un Drive, que comparta el enlace — y **hay que insistir en
  que lo comparta como «cualquiera con el enlace»**, porque un Drive
  privado se ve perfecto desde la cuenta de quien lo subió y sale en
  blanco para todo el equipo. Es el error más común y el más silencioso.

Y lo que la convierte en NovaBot y no en una página de enlaces:
**análisis por pantalla, sin que nadie lo pida.** Fallos, olvidos,
errores que se repiten, incumplimientos. Cada pantalla con el suyo.

### Lo demás que pidió, en orden de lo que bloquea

1. **Equipo sale de Central y entra a novAcademy**, enlazado con
   Empresarial y con las tiendas que tengan vacante de admin o gestora.
2. **Meta Ads en Central** con su propia llave, igual que en
   Empresarial: si pauta para Nova o para novAcademy, esas cifras van en
   su centro de mando. **Sin mezclar funciones entre pantallas.**
3. **El Demo es por artefacto, no uno solo.** Cada proyecto que suba
   genera su demo: Empresarial, novAcademy (estudiante y docente), Son
   de Sky, Procesos Químicos, Química Ambiental. Que se dupliquen, lean
   los documentos que les dé y reflejen eso en el demo.
   *(Yo lo había separado en «Demo para clientes» y «Plantilla». Ella
   dice que es lo mismo. Tiene razón: es un artefacto que se duplica.)*
4. **El panel de NovaSoul no la convence.** Hay que rehacerlo con las
   referencias que mandó: kanban de pendientes, calendario semanal,
   plan de seis semanas.

### Una pregunta que hizo y vale escribir

*«¿Esta pantalla se guarda para siempre ya activada la automatización?»*

**Sí.** Los disparadores viven en el proyecto de Apps Script, no en el
navegador ni en la sesión. Sobreviven a cerrar la página, a cerrar
sesión, y a publicar una versión nueva del código.

Solo dejan de correr si alguien los borra, o si Google los desactiva
después de que fallen muchos días seguidos. Para eso está
`verAutomatico()`, y la tarjeta de Clientes que los muestra: «instalado»
y «funcionando» no son lo mismo.

---

## 9 · NovaSoul, construida (23-09-2026)

Dejó de ser maqueta. Tiene servidor (`78-soul.gs`), cuatro hojas nuevas
en Nova_Soul y una pantalla que lee y escribe de verdad.

**Hay que correr `bootstrapTodo()` una vez** para que aparezcan las
columnas y las hojas nuevas: `Pendientes` crece con `trabajo_id`,
`estado`, `prioridad`, `horas_estimadas`, `horas_reales`, `riesgo` y
`nota`; y nacen `Horas`, `Mindlab` y `Fijos`.

### Lo que hace, y por qué así

**Entra con la sesión de Central.** Sin segundo código, como quedó
decidido en 6c. El servidor igual exige rol `socia` en cada llamada: una
operadora con sesión de Central no abre esto aunque escriba la acción a
mano.

**El riesgo de entrega, que era lo primero que había que construir.**
Compara las horas libres de la semana contra lo comprometido y, si no
cabe, dice qué se puede correr. Tres decisiones dentro:

- **Sin horas libres no calcula.** Mientras la hoja `Horas` esté vacía,
  la pantalla pide el dato en vez de pintar una semana holgada. Es el
  techo de todo lo demás y no se puede suponer.
- **Las horas de un proyecto con horas fijas no se suman dos veces.**
  Las doce semanales de PHH ya incluyen la tarea de PHH; sumarlas daría
  una semana imposible, y una alarma falsa se apaga sola en una semana.
- **Solo puede caerse lo que de verdad sumó.** Correr el encargo de PHH
  no libera una hora, así que no aparece como candidata. Si corriendo
  todo lo corrible sigue sin caber, lo dice: *«eso no se arregla
  moviendo una entrega, se arregla hablando con alguien».*
  Lo inamovible —un parcial— nunca se ofrece.

**La frontera con Central es código, no una nota.**
`soulCargaPorTrabajo_` devuelve **conteos y horas, nunca texto**, y
Central pinta ese peso en cada proyecto. Lo de PHH es confidencial:
Central se abre delante de una socia o un contador, Soul no se abre
delante de nadie. Hay una prueba que falla si algún día el texto se
cuela.

### Mindlab, recortado para el trimestre

Doce semanas, del **28 de septiembre al 20 de diciembre**. Las dos
últimas del año quedan libres —son finales y fiestas, y un plan que las
ocupa se incumple el primer día. Los doce temas son los suyos; lo que
ajusté fue el calendario y el peso: **37 horas en total**, unas tres por
semana, y las dos de diciembre a dos horas. Cada semana baja a
Pendientes con un botón, con sus horas y su fecha.

El plan vive en una constante, no en filas sembradas: la hoja guarda
solo lo que ella marque. Así se puede corregir en una línea sin dejar
doce filas viejas contando otra historia.

### Gastos fijos

Las nueve categorías que nombró: arriendo, mercado, servicios, internet,
crédito, deudas, móvil, varios y ahorro. Todas modificables.

**El plan vive en Soul y los movimientos en Central**, y por eso se
pueden comparar. Si el presupuesto se reescribiera solo con lo que se
gastó, siempre cuadraría y nunca serviría. Donde no hay monto dice «sin
definir», no cero. El ahorro va aparte, no sumado a los gastos. Las
monedas no se suman entre sí.

### Nova Family, en vez del botón directo

Quitado el botón grande de «Entrar a NOVA». En su lugar una sección con
el resumen de las tres: las alarmas de la tienda, **el mismo semáforo de
pauta que le llega por correo los lunes** —una sola cuenta, no dos—, los
cobros atrasados y las entregas vencidas de Central, y novAcademy
diciendo la verdad en vez de un cero mudo.

**Una tienda a la vez**, con selector. Es la regla de toda Nova y aquí
también: dos tiendas en la misma pantalla invitan a compararlas.

Se carga cuando ella abre la sección, no al arrancar: abre la hoja del
cliente y calcula el semáforo, y eso son segundos.

### El cielo, dicho como es

La sección sigue, pero ahora dice que **los tránsitos que muestra son
inventados** y qué falta para que sean suyos: su carta cargada una vez,
los tránsitos con su duración y su ciclo. Dejarlos pasar por reales era
lo único que esa pantalla no podía hacer.

### Lo que sigue de NovaSoul

- **Los colores**: quedaron como estaban, a propósito. Ella decide
  después de verlo.
- **Horarios de clase y turnos de Salsabor**: hoy las horas libres se
  escriben a mano día por día. Con los horarios, el cálculo sería solo.
- **Comidas, movimiento, recompensas y ciclo**: siguen sin servidor.
- **Medir si la astrología le funciona a ella**: guardar la ventana en
  que se hizo cada tarea, que es lo que hace posible el dato de los tres
  meses.

### 9b · Dónde se leen los archivos de la universidad (23-09-2026)

Ella preguntó, y la respuesta honesta era «en ninguna parte todavía».
No estaba construido. Ahora sí, con una frontera explícita.

**Los archivos no se leen y no se copian.** Viven en su Drive, en la
carpeta de cada materia, y Nova guarda el **enlace**. Misma regla que
con PHH: Nova sabe dónde están las cosas, no qué dicen.

**Lo que sí se lee es el texto que ella pegue.** Un cronograma de sílabo
casi siempre es texto —«Parcial 1 · 15 de octubre · 25%»— y de ahí salen
fechas y porcentajes con aritmética, no con adivinanza. Formatos que
entiende: `15 de octubre`, `octubre 15`, `15/10/2026`, `03/11`, y
rangos (`del 12 al 16 de octubre`, del que toma el **último** día porque
es cuando se entrega).

**Lo que NO hace: entender un PDF.** Subir el archivo y que se lea solo
necesita un modelo de lenguaje, que cuesta por cada lectura. La decisión
sigue aparcada, y la pantalla lo dice en vez de disimularlo.

**Tres reglas que lo hacen confiable:**

1. **Nunca escribe sola.** Propone, ella confirma casilla por casilla, y
   puede corregir el título y la fecha antes de guardar.
2. **Muestra lo que ignoró y por qué.** Un lector que solo enseña sus
   aciertos parece infalible y no lo es — lo peligroso de un parcial no
   es que quede mal escrito, es que no quede.
3. **Dice cuándo adivinó.** Un sílabo no escribe el año; se elige el más
   cercano a hoy y se marca «el año lo puse yo». Una fecha adivinada que
   no se anuncia es una fecha inventada.

**Un parcial entra como `inamovible`**, y eso conecta con el repartidor
de la semana: nunca se lo va a proponer como candidato a caerse.

**Borrar una materia NO borra sus entregas**, y el aviso lo dice antes.
Un parcial sigue teniendo fecha aunque se borre la carpeta que lo
mencionaba.

**Dos errores que encontré corriéndolo**, y que vale dejar escritos
porque son del mismo tipo: el encabezado `ESTADÍSTICA 2026-2` entraba
como «26 de febrero», y pegar el mismo sílabo dos veces duplicaba los
parciales porque el anti-duplicados comparaba «Parcial 1» contra
«Estadística · Parcial 1». Los dos habrían pasado desapercibidos hasta
el día del examen.

**Falta**: los horarios de clase. Con ellos, las horas libres de cada
día se calcularían solas en vez de escribirse a mano.

### 9c · Turnos, clases y la plata chiquita (23-09-2026)

Tres cosas que pidió y que resultaron ser la misma: lo que ocupa el
tiempo y lo que mueve la plata.

**Lo que se repite se escribe UNA vez.** La hoja `Rutina` guarda turnos
y clases con su día, su horario y —si es turno— su paga. Se repiten
solos. `desde`–`hasta` acotan el semestre o el contrato: sin eso, una
clase de este semestre seguiría ocupando el jueves en marzo.

**Cambió el significado de las horas.** Antes ella escribía las horas ya
descontadas. Ahora escribe las **útiles** y Nova resta:

```
libres = útiles − (turnos + clases de ESE día)
```

El cambio no puede ser silencioso, así que el formulario muestra la
resta día por día **mientras escribe**, y la cuadrícula de la semana la
muestra también (`12 h − 8 h = 4 h`). Un texto explicando el cambio no
habría bastado; el número nuevo tiene que verse en el momento.

**La paga es fija, las propinas no.** La base vive en la rutina; las
propinas se escriben al día siguiente. Hasta que no se escriban, el
turno aparece pendiente en Hoy y en Mi plata. Cerrar un turno escribe el
ingreso en Finanzas de Nova_Central, y **corregir las propinas actualiza
ese mismo movimiento** en vez de crear otro: por eso existe
`finanza_id`. Un día sin propinas se cierra con cero — un día malo es un
dato, no un olvido.

**Gastos hormiga**: buses, transporte, Uber, antojos, salidas y otros
sueltos. Dos toques desde Hoy. Van a la misma hoja Finanzas con su
categoría.

**UNA SOLA CONTABILIDAD.** Turnos, propinas y hormiga se escriben en
Finanzas de Nova_Central, que es donde ya viven sus movimientos. Si Soul
llevara su propia caja, a fin de mes habría dos respuestas a «¿cuánto me
queda?» y ninguna forma de saber cuál es la buena.

**El fin de mes, en dos mitades que no se mezclan:**

| Ya pasó (hecho) | Falta que pase (previsión) |
|---|---|
| entró · salió · queda | turnos por trabajar · fijos sin pagar |

Un solo número que junta las dos se lee como un hecho, y con eso se
deciden compras. Se muestran aparte y la pantalla dice por qué.

**Tres errores que salieron al correrlo:**

1. Un turno de 18:00 a 02:00 daba **−16 horas**, así que la semana salía
   con horas de sobra justo los días que trabaja de noche.
2. `soulHoy` se caía **entera** si faltaba una hoja nueva: entre
   publicar el código y correr `bootstrapTodo()` pasan minutos, y en
   esos minutos NovaSoul no mostraba una sección incompleta, no mostraba
   nada. Ahora degrada y dice qué hojas faltan.
3. Un «guardé 2 entregas» lo borraba la recarga que venía justo detrás.

**Falta**: que los horarios de clase entren desde el sílabo, junto con
las fechas de entrega. Hoy la clase se escribe a mano una vez.

---

## 10 · Los proyectos a fondo, y cómo cobra cada uno (23-09-2026)

### Dónde va la información profunda, que era la pregunta

En **Nova Central → Mis proyectos → A fondo**. Cada proyecto tiene:

| Bloque | Qué hay |
|---|---|
| Especificación | qué hay que hacer, el contrato |
| **Mi parte** | la cuenta entera del porcentaje, resta por resta |
| La plata | cobrado, por cobrar, atrasado, lo que falta del acuerdo |
| **Fuentes** | Excel, PDF, Word, PPT, Drive o artefactos de Claude |
| **Lo que hay que entregar** | las tareas de ESTA semana y lo vencido |
| Pegar un texto | saca las fechas y las vuelve entregas |

**Las fuentes son ENLACES.** El archivo se queda donde está. Es la misma
regla de las materias y de PHH: Nova sabe dónde están las cosas, no
guarda una copia de lo que dicen. La pantalla dice cuáles Nova **puede
leer de verdad** (Excel/CSV, que ya sabe leer) y cuáles solo guarda el
enlace. Entender un PDF, un Word o una presentación por su cuenta sigue
necesitando el modelo de lenguaje, y la pantalla lo dice en vez de
disimularlo.

**Las entregas se crean desde Central y se guardan en NovaSoul.** El día
a día vive en Soul; si vivieran en las dos, la semana se contaría dos
veces.

### Cómo cobra cada proyecto

Tres columnas nuevas, porque son tres preguntas distintas:

- **`tipo`** — qué ES el proyecto: cliente, empleo, propio, universidad.
- **`mi_rol`** — qué soy YO adentro: socia, trabajadora, propio, estudio.
- **`modalidad`** — de dónde sale la plata: precio fijo, porcentaje,
  por hora, o no se cobra.

Cuando es porcentaje, `tienda_id` apunta a la tienda de Empresarial y
`base_porcentaje` dice sobre qué. **Lo deducido se marca como deducido**
en pantalla: un valor supuesto que parece escrito es el que nadie revisa.

### Su 50% de Nutrea EC y GT

Ella eligió: **utilidad neta del mes**, después de pauta y costos fijos.

```
  ganancia de lo entregado
  − devoluciones (el flete se pagó y no vuelve)
  − pauta
  − costos fijos del mes
  = utilidad neta   →   × 50%
```

La pantalla muestra **la cuenta entera**, no solo el resultado. Es el
número con el que decide si el mes alcanza, y un número sin su cuenta no
se puede discutir.

Tres reglas que la hacen fiable:

1. **Un costo fijo vacío NO es cero.** Se avisa que no se descontó. Un
   vacío tratado como cero sube la utilidad, y sobre esa utilidad se
   reparte plata de verdad.
2. **Un mes en pérdida da parte CERO, no negativa.** A ella no le cobran
   por un mes malo, y un número negativo invita a restarlo de sus
   ingresos.
3. **Un mes sin pedidos dice que no hay datos**, que es distinto de un
   mes malo.

### La frontera, ahora con interruptor

`confidencial` decide si el TEXTO de las tareas sube a Central. **Nace
encendido para los empleos**, porque lo de PHH lo es y ella lo dejó
dicho. Vacío no es «no»: para un empleo es «sí». Un olvido no puede ser
lo único que proteja eso. Para un proyecto confidencial Central ve
«3 entregas, 14 horas» y nada más.

### Un error que salió al correrlo

`costos_fijos_mes` se leía con `ajustes()`, que **solo devuelve las
claves que ya conoce** — y esa no es una de ellas. Devolvía vacío
siempre, así que los costos fijos nunca se descontaban y la utilidad
salía alta. Es exactamente el error que nadie revisa porque da un número
que gusta.

### Lo que sigue, en el orden que ella eligió

1. ~~Proyectos a fondo + cómo cobra cada uno~~ — HECHO
2. **Meta Ads en Nova Central**, igual que en Empresarial, y la misma
   conexión para novAcademy.
3. **El lugar para cargar la carta astral.** Ella sube lo que saca de
   Horus y Nova completa lo que falte para leer los tránsitos de la
   semana, el mes, el año y su revolución solar. Su marco, que manda
   sobre el diseño: *los planetas personales afectan lo inmediato y lo
   propio; los sociales hablan de expansión y estructura; los
   generacionales marcan época y van más allá del ego.* El planeta tiene
   que decirle **qué momento es**.
4. novAcademy: Equipo sale de Central y entra ahí.

---

## 11 · El cielo de NovaSoul (23-09-2026)

Su marco, con sus palabras, porque manda sobre todo el diseño:

> «Los planetas marcan el tiempo: hay tiempo para aprender, tiempo para
> descansar y tiempo para cambiar. Los personales afectan lo inmediato,
> lo mío, lo propio. Los sociales hablan de expansión y estructura. Los
> generacionales marcan época, van más allá del ego. **El planeta me
> debe mostrar qué momento es.**»

### El reparto: qué calcula Nova y qué trae ella

**NOVA CALCULA**, sin internet y sin inventar:

- **La fase de la Luna** de cualquier día. Es la única astronomía que
  Nova hace sola: una cuenta de dos líneas desde una luna nueva conocida
  (6 de enero de 2000, 18:14 UTC) y el mes sinódico. Comprobada contra
  seis lunas reales de 2026 — si algún día se corre, la prueba falla.
- **La ventana de su revolución solar**: de cumpleaños a cumpleaños. Es
  aritmética de calendario.
- **Qué temporada del pensum está abierta hoy.**

**ELLA TRAE**, de Horus, que ya se lo da bien: su carta natal, los
tránsitos con sus fechas, y la carta de la revolución. Calcular
efemérides aquí sería rehacer mal algo que ya está bien hecho, y con una
precisión que no se puede verificar desde una hoja de cálculo.

### Su pensum manda sobre la luna

Cada día dice **qué momento es**: aprender, descansar o cambiar. Si hay
una temporada del pensum abierta, manda esa. Si no, habla la luna. Su
marco antes que la lectura común, siempre — y la pantalla dice de cuál
de las dos salió.

### Lo único que aquí es un dato y no una creencia

`cieloMedir_` cuenta, **fase por fase**, cuántas de sus entregas
cumplió. *«Las entregas que pusiste en cuarto menguante las terminaste
el 80% de las veces, contra un 62% en general.»* Eso es su propio
número, no una creencia prestada.

Tres guardas para que sea honesto:

1. **No se habla sin muestra.** Con menos de 8 entregas en una fase se
   dice cuántas faltan. Un 100% sobre dos casos no es un patrón: es una
   casualidad con decimales.
2. **El patrón se mide contra SU promedio**, no contra cero. Si cumple
   el 70% de todo, un 72% en menguante no es nada.
3. **Si no hay patrón, se dice.** «Miré tus entregas fase por fase y no
   encontré diferencia que valga la pena nombrar. Eso también es un
   resultado.»

Y una regla de conteo: una entrega **del futuro** no cuenta como
incumplida. Todavía no se ha cumplido ni incumplido.

### El lector de la carta

Pega el listado de Horus y Nova propone: planeta, signo, grado, casa y
retrógrado. Como con el sílabo, **muestra lo que no entendió** y además
**avisa qué planetas le faltaron** — una carta a la que le falta Saturno
en silencio es peor que una carta vacía. Volver a cargarla actualiza las
filas, no las duplica.

Una guarda que salió de la prueba: «Urano Acuario 12» **no** inventa una
casa. El 12 es el grado. La casa solo se toma si la palabra «casa» está,
porque una casa inventada mueve el tema entero.

### Lo que sigue

1. **Meta Ads en Nova Central y novAcademy** — solo la pantalla, con la
   misma lectura que Empresarial. Sin token y sin conectar todavía.
2. **Equipo sale de Central.** Hoy en Central es una maqueta con nombres
   inventados que duplica lo que Empresarial ya hace de verdad.
3. Que los tránsitos entren pegando el texto de Horus de una vez, como
   el sílabo, en vez de uno por uno.

### 11b · La carta de Manuela, y dónde está de verdad la frontera (23-09-2026)

Ella mandó las capturas de Horus y preguntó si Nova podía leerlas. La
respuesta separa dos cosas que se venían confundiendo:

- **Claude (este chat) SÍ lee imágenes.** Le mandó cuatro capturas y de
  ahí salió su carta completa, transcrita.
- **Nova (el Apps Script) NO.** Corre en los servidores de Google y no
  tiene modelo de visión. Ahí la frontera no se mueve.

Así que el camino que funciona es: ella manda la captura por el chat, se
transcribe, y el texto se pega en NovaSoul. El archivo
`carta-manuela.txt` tiene su carta lista para pegar.

**Su carta, para que quede escrita:** Sol en Virgo 27 (casa 9), Luna en
Leo 10 (casa 7), Ascendente en Capricornio 15, Mediocielo en Libra 19.
Saturno, Urano, Neptuno y los dos Nodos, retrógrados.

*(La maqueta decía «Asc Leo, Luna Capricornio». Eso me lo inventé yo
cuando no había datos. Era al revés.)*

**Tres cosas que el lector aprendió con la carta de verdad:**

1. **Horus abrevia los signos** — «Vir», «Lib», «Esc», «Cán». Sin esas
   abreviaturas, pegar su pantalla tal cual no encontraba NI UN signo,
   que es exactamente lo que ella iba a hacer.
2. **«Vir 27 Casa 9» tiene dos números.** El grado es 27, no 9.
   Confundirlos le ponía el Sol a 9 grados y movía medio análisis.
3. **«Nodo Sur» empieza por «Nodo».** Buscando palabra por palabra, el
   Nodo Sur se volvía Nodo Norte y después se descartaba por repetido:
   desaparecía de la carta sin decir nada. Ahora se buscan los grupos
   largos de palabras primero.

**Los asteroides van en su propio cajón.** Ella fue explícita con sus
tres grupos —personales de Sol a Marte, sociales Júpiter y Saturno,
generacionales Urano, Neptuno y Plutón—. Meter a Quirón o a Ceres entre
los generacionales diluiría su marco, y el marco es suyo. Hay un quinto
grupo, «Puntos y asteroides: matices, no el marco».

### La lectura de la revolución la compone Nova

De una tabla, no de una interpretación nueva cada vez. Ella teclea dos
datos —el ascendente del año y en qué casa cae su Sol— y Nova arma:

- **Cómo entra al año** (el ascendente).
- **Dónde va su atención** (la casa del Sol).
- **Qué momento pide el año**, de la división tradicional de las casas,
  que resulta ser exactamente sus tres momentos:

| Casas | Clase | Momento |
|---|---|---|
| 1, 4, 7, 10 | angulares | **cambiar** |
| 2, 5, 8, 11 | sucedentes | **descansar** |
| 3, 6, 9, 12 | cadentes | **aprender** |

No es una regla inventada para que cuadrara: así se leen las casas desde
hace siglos, y da la casualidad de que sus tres momentos son esas tres
clases.

Con los mismos datos dice siempre lo mismo — y eso es justo lo que
permite que ella la discuta y la corrija. Una lectura generada de nuevo
cada vez diría algo distinto el martes que el jueves.

**Y propone temporadas de pensum** desde los tránsitos que ella cargó:
usa las fechas de Horus, agrega el momento según la casa, y **propone**
— ella confirma. Solo los que duran catorce días o más: un tránsito de
la Luna dura dos días y medio y llenaría la lista de ruido tapando a
Saturno, que es el que marca meses.

### Equipo salió de Nova Central

Era una maqueta con nombres inventados que duplicaba lo que Empresarial
ya hace de verdad. Fuera: el menú, la vista, `buildEquipo` e
`inviteUser`.

**Lo que ella decidió:** el equipo de Nutrea se ve en **Nova
Empresarial**, que es donde vive. Y en **novAcademy** NO va esa pantalla
tal cual — hay que rehacerla en torno a **profesorado y estudiantes**,
que es otra cosa.

### 11c · Su fecha de nacimiento, y un dato que casi se carga mal

**Nació el 20 de septiembre de 1995.** Su año solar vigente va del
**20-09-2026 al 19-09-2027** — arrancó tres días antes de esta
conversación. Cumplió 31.

*(La maqueta decía 14 sep 1998. Eso me lo inventé yo. El Sol a 27° de
Virgo encaja con el 20 de septiembre, así que la carta y la fecha se
confirman entre sí.)*

**El error que estuvo a punto de pasar.** Ella dijo: «en qué casa cae mi
sol ya lo tienes». Y sí, en su carta NATAL el Sol está en casa 9 — pero
la revolución solar es una carta **distinta**, levantada para el momento
en que el Sol vuelve a su grado. Tiene su propio ascendente y sus
propias casas. Cargar el 9 de la natal habría dado una lectura del año
entera sobre un dato falso, y nada en pantalla lo habría delatado.

El formulario ahora lo dice antes de que escriba: que es otra carta, que
en su natal es casa 9 y que en la revolución casi seguro es otra, y
dónde sacarlo en Horus. Un formulario que pide un número sin decir de
dónde sale invita a poner el que uno tenga a mano.

### 11d · Lo que rige su año: la profección

Ella pidió «algo que rija mi año», y en tiempos distintos: hoy completo
e informativo, y algo más grande encima.

La respuesta es la **profección**: una casa por año cumplido, una casa
por mes solar. Es lo único del cielo que es **aritmética pura** — no
hace falta ninguna efeméride, no hay nada que adivinar, y con los mismos
datos da siempre lo mismo. Por eso puede estar en pantalla sabiendo solo
dos cosas suyas: su **Ascendente** y su **fecha de nacimiento**.

Para ella, hoy:

| | |
|---|---|
| **El año** | Cumplió 31 → casa **8**. Con Ascendente Capricornio, casas enteras, la 8 cae en **Leo**. Lo rige el **Sol**, que en su carta está en **Virgo, casa 9** — ahí es donde se le va a notar el año. Casa sucedente: el año pide **descansar**. |
| **El mes** | Mes solar 1 de 12, del **20 sep al 19 oct**, casa 8 otra vez. El mes solar cuenta desde su cumpleaños, no desde el 1° del calendario. |
| **El día** | La luna y su pensum, que ya estaban. |

En pantalla: bloque ámbar `ci-anio`, con el año, el mes y **los doce
meses de un vistazo**. Dice explícitamente que **es una cuenta, no una
lectura** — para que no se confunda con lo que sí es interpretación.

**Su revolución solar 2026–2027**, del PDF que ella encontró (Horus no
la dejaba sacarla). Momento exacto 20/09/2026 01:09, residencia Armenia:

- **Ascendente Cáncer 16°20'** ← este es «el ascendente del año» que preguntaba
- **Sol en Virgo 27°19' → casa 3** ← NO es la casa 9 de su natal

Esos son los **dos únicos números** que hay que escribir en el
formulario de revolución. Lo demás lo compone Nova.

### 12 · Lo automático de Nova Central y NovaSoul

Lo que ya corría solo era todo del lado del **cliente**: tasas (5h),
Meta (6h), alarmas (7h), semáforo el lunes (8h). Todo eso mira las
tiendas de Nutrea y le escribe a la dueña.

De **su** lado no corría nada. Nova sabía que la semana no cabía y que
un cobro llevaba once días tarde, y no lo decía hasta que ella abriera
la pantalla — que es justo lo que no pasa esos días.

`apps-script/83-automatico-mio.gs`, registrados en `TRABAJOS` para que
aparezcan solos en la tarjeta de «lo automático» de Central:

| Trabajo | Cuándo | Qué dice |
|---|---|---|
| `soulLunes` | lunes 9h | ¿cabe la semana?, qué se puede mover, qué ya está vencido, turnos sin propinas, la semana de Mindlab, qué momento es |
| `centralDiario` | diario 9h | cobros tarde, cobros que vencen en ≤3 días, entregas de proyecto |

**Dos reglas que no se negocian, y que están en la prueba:**

1. **Correo vacío no se manda.** La misma regla de `revisarAlarmas`. Un
   aviso diario que dice «todo bien» deja de leerse en dos semanas, y
   entonces tampoco se lee el día que sí traía algo.
2. **Lo de PHH no sale de NovaSoul.** El correo de Central dice el
   nombre del proyecto y **cuántas** tareas abiertas tiene; nunca qué
   dicen. Un correo se reenvía y se imprime. `pruebas/automatico-mio.js`
   lo comprueba por las dos puntas.

Y una tercera que ya existía y aquí se nota más: el correo del lunes
**no ofrece correr lo que no libera horas**. Ni el parcial (inamovible),
ni lo de PHH (ya está dentro de sus 12 h fijas). Ofrecerlo sería
proponerle un sacrificio que no arregla nada y hacerle creer que la
semana ya cabe.

**Falta prenderlos:** correr `prenderAutomatico()` una vez, o el botón
de la consola en Nova Central. Los disparadores son del proyecto, no de
cada cliente.

**Una prueba se corrigió, y vale decir por qué.** `pruebas/semaforo.js`
exigía que el semáforo corriera *después de todos los trabajos
diarios*. Esa regla es más ancha de lo que hace falta: lo que importa es
que corra después de los **tres que lo alimentan** (tasas → Meta →
alarmas), para juzgar con el gasto ya adentro y convertido.
`centralDiario`, que lee el libro de ella y no toca la tienda, la
rompía sin que hubiera nada roto. Ahora los tres se nombran uno por uno.

### 13 · Meta Ads en Nova Central: la maqueta fuera

Ella dijo: *«solo necesito las pantallas, no la integración completa
porque aún no pauto»*. Y antes: *«el meta de nova central sigue igual,
no está igual a la de empresarial»*.

**Lo que había** era una maqueta: un cuadro para subir un CSV y cuatro
cifras —alcance, CPA, ROAS, inversión— **escritas a mano en el HTML**,
que no salían de ningún lado y no cambiaban nunca. Eso es peor que una
pantalla vacía: una pantalla vacía se nota, un número inventado se cree.

**Lo que va en su lugar** es lo único que Central puede decir de verdad
sobre Meta: **quién está conectado y quién no**. El gasto y el CPA de
cada tienda se miran adentro de esa cuenta, en Nova Empresarial, que es
donde está el dato.

**Tres estados, no dos.** «Conectada» de verdad es llave + número de
cuenta en todas las tiendas. El de en medio —llave guardada, cuenta sin
poner— es el que se pierde si solo hay sí y no, y es justo el que deja a
alguien creyendo que el gasto está entrando.

**Central no tiene dónde escribir una llave, y es a propósito.** La
llave es de cada cliente y la escribe su dueña en SU Nova Empresarial.
Un campo aquí sería una pantalla desde la que se conectan cuentas
ajenas — y el día que alguien más opere la consola, esa es la pantalla
que hay que no haber construido. `pruebas/central-meta.js` cuenta los
`input` de la vista y exige **cero**.

**Nova y novAcademy** dicen «todavía no pauta», en vez de dibujar un
panel de campañas en cero que se lee como una mala semana.

**Un detalle del guardia.** La prueba busca cada cifra inventada en el
archivo, no en la pantalla — un número escondido en un div oculto
también pasaría una revisión visual. Al escribir el comentario que
explica esto, repetí las cifras, y la prueba se cayó con razón: un
comentario que las nombra deja el guardia sin filo. El comentario ahora
las describe sin escribirlas.

**Pendiente:** novAcademy sigue sin su pantalla de Meta y sin la de
profesorado/estudiantes.

### 14 · El token de Meta, cerrado

Ella dijo que creó como cinco tokens después de ese, porque la
integración no cargaba bien. El que pasó por el chat quedó enterrado
entre los que vinieron después y además **no hay integración conectada
todavía** — no pauta. Se deja de insistir con eso.

Lo que sigue en pie, para cuando conecte: `ads_read` únicamente, permiso
de activo «ver rendimiento», nunca «Administrar cuentas publicitarias»,
y el token va a Script Properties por `sheetId`, nunca a una celda.

### 15 · Pedidos y Novedades: la tabla con panel al lado

Ella lo dijo así: *«hay algo que no hiciste ni aplicaste, y es el nuevo
diseño de la pantalla de pedidos y novedades»*. Y tenía razón: mandó seis
imágenes de referencia y la número 4 decía exactamente esto —

> **Imagen 4 (Polytrox)** → Nova Empresarial, pantalla Pedidos y pantalla
> novedades con los estados (estructura, no colores). La tabla de pedidos
> con columnas + panel derecho de detalle del pedido seleccionado es lo
> que le falta a tu pantalla de Pedidos actual. Los colores naranja/coral
> los cambias por el oliva+dorado de Nova.

**Hecho.** Tabla a la izquierda (Cliente · Producto · Destino · Valor ·
Estado), panel de detalle a la derecha, con el color de Nova y una franja
del color del estado en el borde de cada fila. En teléfono se apilan: la
fila pasa a ser una ficha de tres renglones y el panel cae debajo.

Las bandas por urgencia, los filtros y el sello de EJEMPLO siguen igual.
El ejemplo usa la MISMA tabla que los datos reales — si se dibujara
distinto, lo que ella aprende mirándolo no le serviría después.

### 16 · Los dos fallos de «no me actualiza los estados»

Eran dos cosas distintas con el mismo síntoma, y las dos decían que
habían funcionado.

**a · Clasificar un estado no tocaba los pedidos que ya estaban.**
La traducción se guardaba en la hoja `Estados` —el diccionario— y los
pedidos importados conservaban el `estado_canonico` del día que entraron.
La pantalla decía *«las cifras se están rehaciendo»* y no se rehacía
nada: solo volvían a contarse bien si reimportaba el archivo entero.

Peor: el código lo afirmaba en un comentario —«cada pantalla recalcula
desde Pedidos, y el estado se vuelve a traducir al leer»— y la función
que haría eso, `estadoCanonico()`, **no la llamaba nadie**. Era código
muerto que servía de coartada. Se borró.

Ahora `reaplicarEstado_()` reescribe los pedidos que ya estaban, devuelve
cuántos movió, y la pantalla dice el número: *«Guardado · 2 pedidos
cambiaron de cuenta»*. Nunca pisa el `estado_nova` de un pedido corregido
a mano: una regla general no manda sobre una decisión puntual.

De paso, el aviso de mes cerrado contaba de más. Contaba todo pedido con
ese texto, incluidos los de otra plataforma y los corregidos a mano —que
no cambian de cifra. Ahora el aviso sale de lo que **de verdad** cambió.

**b · Cambiar el estado de UN pedido no se veía.**
La etiqueta de la fila miraba solo `estado_canonico`, nunca
`estado_nova`. Ella elegía el estado, guardaba, salía «Guardado en la
hoja» y la etiqueta seguía igual, el pedido no cambiaba de banda y la
barra de arriba no se movía. Se guardaba de verdad: la que no se
enteraba era la pantalla. Y encima no repintaba.

Se arregló con **un solo lugar de verdad**: `estadoDe(p)` devuelve
`estado_nova || estado_canonico` —la misma precedencia del servidor— y
lo usan la etiqueta, las bandas y los diez filtros. Guardar ahora
repinta tabla, filtros y barra.

**Una prueba que faltaba.** Al armar la tabla dejé un `let` repetido que
tumbaba TODO el JavaScript de `empresarial.html`. Ninguna prueba dibujaba
Pedidos, así que no se supo hasta que reventó otra pantalla.
`pruebas/pedidos-pantalla.js` ahora dibuja las dos.

---

## LO QUE SIGUE · pedido el 23 de septiembre

Lo apunta aquí porque ella lo pidió así: *«lo que te diga aquí y no esté
en contexto con lo que haces en este momento guárdalo en la carpeta de
pendientes»*. **Nada de esto está empezado, y hay preguntas que hacerle
antes de empezar.**

### P1 · El clima astral de verdad, y el pensum automático

Sus palabras: *«hay algo que no entiendo y es el pensum kármico, la idea
es que con la información que yo brindo, y la que debes cruzar con
internet… solo necesito que lea y profundice en el clima astral con toda
la información que hay de mí, que en el sheet haga un enlace o algo para
leer directo los tránsitos de la semana, mes y año, cruzar esa
información con mi revolución solar y mi carta natal y sacar un análisis
detallado, no lo quiero solo en una ventana, y el pensum kármico debe
crearlo Nova automáticamente»*.

Hoy el pensum se escribe a mano, o se acepta lo que Nova propone desde
los tránsitos que ella pegó. Ella quiere que Nova **los traiga sola**.

**La frontera real:** Apps Script SÍ puede llamar a una API por internet
(`UrlFetchApp`), pero **no hay efemérides dentro de Nova**. Calcular
dónde está Saturno el martes que viene pide o una librería astronómica
—que en Apps Script no existe— o una API externa. Esa es la pregunta
que hay que hacerle antes de escribir una línea.

### P2 · Cómo se cobra: por hora, día, semana o mes

Sus palabras: *«me estás pidiendo en Nova Central al poner el trabajo
como un pago único y no me das a elegir si es único o es por día, por
mes o por semana, ni por hora… dámelo también por hora porque el PHH es
por horas»*.

Hoy `modalidad` en `Trabajos` tiene: precio fijo, porcentaje, por hora y
nada. Falta que la PANTALLA deje elegir la **periodicidad**: única, por
hora, por día, por semana, por mes. Y con PHH por horas, hay que decidir
de dónde salen esas horas.

### P3 · Archivos por proyecto que Nova lee y recuerda

Sus palabras: *«necesito que me des la opción de subir archivos para que
lean y tengan guardados para cada proyecto, de esos archivos pueden sacar
fechas de entrega, tareas, tips, libros, temas de evaluaciones!! todo eso
necesito que Nova lo recuerde y cree automáticamente el pendiente y le
saque horario… no solo el enlace del documento en Drive, porque son
muchos documentos»*.

Hoy `Fuentes` guarda **enlaces**, no contenido. Y Apps Script no lee un
PDF, un Word ni un PowerPoint: eso necesita un LLM. Ella ya propuso la
salida: *«yo voy a hacer el asistente virtual en n8n si es necesario»*.

### P4 · Los artefactos, y que se vean desde NovaSoul

Sus palabras: *«voy a organizar los artefactos para que lean el sílabus y
saquen el plan de estudio con fechas y tiempos estimados… cosas y
pendientes que debo hacer que Nova Soul me debe preguntar y yo responder
esa pregunta y poder editar las respuestas, que no todas sean
predeterminadas, para ubicar las tareas con mayor urgencia, también los
trabajos… que en Nova Soul me aparezca el link del artefacto para
trabajar ahí y poder encontrarlo fácil»*.

Dos piezas: **el link del artefacto** guardado por proyecto/materia y
visible en NovaSoul, y **preguntas con respuesta editable** para ordenar
la urgencia, en vez de que todo sea predeterminado.

### P5 · De antes, sin tocar

- novAcademy: su pantalla de Meta, y la de profesorado/estudiantes.
- El equipo de Nutrea, visible en Nova Empresarial.
- Demo por artefacto (uno por proyecto, seleccionable).
- NovaBot separado de NovaChat.
- Onboarding: 3 planes, módulos por cliente, moneda de residencia.
- Empresarial: 15 viajes al servidor al cargar → una sola llamada.

### P1-bis · Lo que se averiguó de las efemérides (23 sep)

Ella pidió comparar las dos APIs. **No se pudo: la política de red del
contenedor bloquea `freeastroapi.com`, `openephemeris.com` y `rapidapi.com`.**
No se inventó una comparación.

Pero sí se pudo probar **el motor que hay debajo** de una de ellas
(Astrologer API → Kerykeion → **Swiss Ephemeris**), instalando
`pyswisseph` directo. Y la prueba que vale es contra SU carta:

**Nacimiento: 20/09/1995 13:20, Palmira (3.5394 N, 76.3036 W), UTC−5.**

| | Swiss Ephemeris | Horus |
|---|---|---|
| Sol | Virgo 27° | Virgo 27 ✓ |
| Luna | Leo 10° | Leo 10 ✓ |
| Mercurio | Libra 20° | Libra 20 ✓ |
| Venus | Libra 5° | Libra 5 ✓ |
| Marte | Escorpio 9° | Escorpio 9 ✓ |
| Júpiter | Sagitario 8° | Sagitario 8 ✓ |
| Saturno Rx | Piscis 20° | Piscis 20 ✓ |
| Urano Rx | Capricornio 26° | Capricornio 26 ✓ |
| Neptuno Rx | Capricornio 22° | Capricornio 22 ✓ |
| Plutón | Escorpio 28° | Escorpio 28 ✓ |
| Nodo Norte | Libra 26.86° | Libra 27 ✓ (redondeo) |
| **Ascendente** | **Capricornio 15.10°** | **Capricornio 15 ✓** |
| **Mediocielo** | **Libra 19.36°** | **Libra 19 ✓** |

**Reproduce su carta entera.** Horus usa el mismo motor.

**La consecuencia: puede que no haga falta ninguna API.** Un año completo
de sus temporadas —cada vez que un planeta social o generacional toca
algo de su carta, con orbe y con fecha de entrada y salida— son
**33 filas**. Cinco años caben en unas 165. Eso entra en la hoja
`Transitos` sin despeinarse: sin llave, sin factura, sin depender de que
un servicio de terceros siga vivo, y exacto al mismo motor que ella ya
usa.

El costo honesto de ese camino: la tabla hay que regenerarla cuando se
acabe el horizonte. No es «Nova la llama sola», que es lo que ella
eligió. Por eso hay que preguntárselo.

### P1-ter · Horus usa PLACIDUS, Nova usa casas enteras

Se descubrió comprobando lo anterior. Contra las diez casas que dice su
captura de Horus:

- **Placidus acierta 10 de 10.**
- **Casas enteras aciertan 6 de 10.**

Las cuatro que se van: Luna (7 vs 8), Venus (9 vs 10), Marte (10 vs 11),
Júpiter (11 vs 12).

Esto **no** es un error en las profecciones: las profecciones son
tradicionalmente de casas enteras y así deben quedar. Pero el día que
Nova diga «este tránsito te cae en la casa N», si usa casas enteras va a
contradecir lo que ella ve en Horus cuatro veces de cada diez — y va a
creer, con razón, que Nova está mal.

**Propuesta a confirmar con ella:** casas enteras para la profección
(porque es la técnica) y Placidus para dónde cae un tránsito (porque es
lo que coincide con su carta), diciéndolo en pantalla para que los dos
números nunca parezcan una contradicción.

### 17 · Los tránsitos se calculan, y el pensum se crea solo

Sus decisiones: **tabla ahora con la puerta abierta a una API después**,
y **las dos casas a la vista, ella decide cada vez**.

**El generador.** `efemerides/generar-transitos.py`, con Swiss Ephemeris.
Lee un JSON con los datos de nacimiento y saca un TSV con las columnas
exactas de la hoja `Transitos`. Para ella, cinco años = **135
temporadas**. No hay llave que cuidar, no hay factura, y su pantalla no
depende de que un servicio ajeno siga vivo. El costo honesto: hay que
volver a correrlo cuando se acabe el horizonte.

Su tabla ya está generada en `efemerides/manuela-transitos.tsv`.

**Las dos casas, en todas partes.** `Transitos` gana `casa_placidus` y
`Pensum` gana `casa_alterna`. La propuesta trae las dos lecturas con su
momento —no los dos números, los dos CONSEJOS, que es lo que de verdad
cambia— y en pantalla hay un botón para cambiar de sistema. El mismo
botón deshace: si cambiar fuera de ida, no lo probaría.

**El pensum automático.** `pensumAuto_()` corre los lunes, dentro de
`soulLunes()`, antes del correo. Tres cuidados:

1. **No crea todo.** Solo lo abierto hoy o lo que empieza en 120 días.
   Volcar cinco años de una sentaría a Saturno de 2031 al lado de lo de
   esta semana, y el pensum dejaría de decir qué momento es.
2. **Lo que ella escribió no se toca.** Se compara por título y fecha.
3. **Queda dicho quién la hizo.** Chip «la puso Nova» en pantalla. Sin
   eso no podría distinguir lo que escribió de lo que le apareció, y lo
   segundo pesa menos.

El correo del lunes trae las temporadas nuevas y, cuando los dos
sistemas discrepan, lo dice.

**OJO · hay que correr `bootstrapTodo()`** esta vez: son dos columnas
nuevas (`Transitos.casa_placidus` y `Pensum.casa_alterna`).

### 18 · Lo que ella vio y no cuadraba

**a · Los pedidos se veían planos.** Sus palabras: «no se ven separados
los pedidos, se ven planos como si fueran todos los mismos». Tenía
razón: la tabla tenía 2px de aire entre filas y una franja de color de
3px. Ahora: 7px de aire, franja de 5px, sombra propia y un leve
levantarse al pasar por encima. Se probó además una franja de color al
pie de cada fila y **se quitó**: con el texto encima se leía como un
subrayado, no como el estado.

**b · Novedades había perdido a la clienta.** «Ya no aparecen los
nombres de los clientes en novedad ni la información para contactarse».

No era un descuido del rediseño: **la hoja `Novedades` nunca ha guardado
nombre, teléfono ni dirección.** Solo el `pedido_id`. Así llega de las
plataformas. La bandeja mostraba códigos y motivos, y nada con lo que
llamar a nadie — que es lo único para lo que existe una bandeja de
novedades.

Ahora se cruza contra los pedidos cargados: nombre de la clienta,
teléfono con enlace para llamar y para WhatsApp, dirección, producto,
valor, transportadora y en qué va el pedido. Y si ese pedido no está
entre los cargados, **lo dice** en vez de dejar el hueco.

### 19 · El inventario: lo que de verdad pasaba

Ella lleva semanas diciendo «sin poder editar lo del TAG RECEDE, no he
podido montar la información, ni el link de la landing page».

Se probó todo el camino, servidor y pantalla (`pruebas/inventario.js` y
`pruebas/inventario-pantalla.js`): **guardar funciona**, incluida la
landing, los precios de combo y las cuatro respuestas.

Lo que sí estaba roto era otra cosa, y es peor:

**Un campo que se deja escribir, descarta lo escrito y responde que
guardó.** Si la hoja `Inventario` de la cuenta no tiene todavía la
columna —`landing`, `precio_2`, `categoria`, `resp_1`…— la pantalla la
pintaba igual, `leerFicha()` la descartaba en silencio y el guardado
contestaba «Ficha actualizada». Ella escribía la landing, leía que se
guardó, y no se guardaba nada.

Ahora ese campo sale **desactivado**, en otro color, y dice en la propia
casilla «no se puede guardar». Y al guardar no se dice que quedó todo:
se dice qué columnas faltan y que hay que correr `bootstrapTodo()`.

**Y un detalle del mismo tipo que ya había aparecido en NovaSoul:** el
mensaje de error se escribía ANTES de repintar la ficha, así que el
repintado se lo llevaba. Quien guardaba a medias no veía ni rastro de
por qué. Ahora va después.

**Queda por confirmar con ella** si su hoja tiene las columnas. Si al
abrir una ficha ve campos apagados, ese es el diagnóstico y la cura está
escrita ahí mismo.

### 20 · META, una sola pantalla · EL SERVIDOR (hecho)

Sus palabras: *«no entiendo lo que hiciste en la parte de pauta, no veo
dónde se ven los datos que lee Nova de Meta […] todo está regado […] que
cargue la información en la misma pantalla, no que me toque ir a
configuración para poder cargar o pedir la info a Meta, nooo eso no me
gusta es desorden»*.

**Sus decisiones:** una sola pantalla llamada META (Configuración deja de
tener lo de Meta); los cuatro grupos de números (gasto/CPA/pedidos,
ROAS/utilidad, alcance/impresiones/clics, por campaña y por anuncio);
**todos** los periodos; y además —esto lo agregó ella—:

> *«del día de hoy, cuánto lleva gastado en el día, qué hacer con las
> ventas y sacar un semáforo diario, son 3 revisiones por día al
> mediodía, a las 5pm y a las 11:30pm»*.

`apis-script/85-meta-panel.gs` → acción `meta_panel`. Una sola llamada
devuelve todo: conexión, periodos, revisiones, qué hacer y desglose.

**Las tres revisiones no dicen lo mismo, y ese es el punto.** No son tres
horas cualesquiera: son los tres momentos en que todavía se puede hacer
algo distinto. Al mediodía se corrige, a las 5 es la última hora útil
para mover presupuesto, y a las 11:30 ya no se corrige nada. **La de la
noche no propone mover presupuesto**, porque sería un consejo que no se
puede seguir.

Tres cuidados más, todos probados:

- **El presupuesto se juzga contra la PARTE del día que va.** A mediodía
  tocan 50 de los 100, no los 100. Compararlo con el día entero sería
  felicitarla por no haber llegado todavía.
- **Gasto sin pedidos NO es rojo.** Es «sin datos», con su explicación.
  Un rojo a las doce por un día que termina bien enseña a no mirar el
  semáforo.
- **La pauta sin convertir va aparte y se dice primero.** Si Meta cobró
  en otra moneda y falta la tasa, ese gasto no está dentro del CPA — y
  el CPA es justo el número que decide si escalar o frenar.

**Falta la pantalla.** Cuando se haga:
- renombrar «Pauta y gastos» → META, mover ahí la llave y «Traer ahora»,
  y quitar esa parte de Configuración;
- la tarjeta «Que el gasto entre solo» **debe desaparecer cuando Meta ya
  está activo** — ella lo señaló: *«que el gasto entre solo sigue
  apareciendo pero Meta ya está activo»*.

### 21 · «¿Dónde está la pantalla de los permisos, me la escondiste?»

**No, y vale decirlo con precisión:** «Permisos» en Nova Empresarial
nunca se tocó — se ve en su propia captura, en el menú bajo EQUIPO. Lo
que sí se quitó, con su visto bueno ese mismo día, fue **Equipo en Nova
Central**, que era una maqueta con nombres inventados.

Lo que ella buscaba era otra cosa: **«Prender lo automático»**. Y ahí sí
había un problema real, aunque no fuera el que ella creía: la tarjeta
estaba metida **arriba de la lista de Clientes**. No estaba escondida,
pero nadie busca el interruptor de lo que corre solo dentro de la lista
de clientes. Que lo diera por desaparecido es la prueba.

Ahora tiene **su propia pantalla y su propio sitio en el menú**, y el
texto ya no habla de dos trabajos sino de los seis que hay.

### 22 · Los dos reportes que pidió

Sus palabras: *«en el cierre de mes, des la opción de sacar 2 reportes
más, el semáforo o diagnóstico o autopsia semanal y un reporte diario de
lo que quedó pendiente del otro día […] que no llegue solo al correo
sino también que aparezca en la pantalla principal»*.

Mandó como referencia una tabla de categorías de gestión con su
cantidad. Eso no es un tablero de métricas: es una **lista de trabajo**.
Dice cuánta gente hay que tocar hoy y por qué. Por eso el reporte no
habla de ventas ni de márgenes — de eso ya hablan el semáforo y el
cierre.

**Ella eligió** que «pendiente» sea: lo que sigue esperando a una
persona, y lo que se venció y nadie movió.

`apps-script/86-reporte-dia.gs` → acción `reporte_dia`. Tres decisiones
que son las que lo hacen útil:

1. **Todo se cuenta DOS veces**: cuántos hay, y cuántos de esos llevan
   demasiado quietos. «12 sin confirmar» no dice si hay que
   preocuparse; «12, y 5 llevan más de un día» sí — esos cinco son los
   que se van a perder. Cada categoría tiene su propio límite: una
   novedad se pudre en un día, un pedido en oficina aguanta tres.
2. **Lo que va en camino NO cuenta como pendiente.** No espera a nadie,
   espera al courier. Meterlo infla la lista con cosas que no se pueden
   hacer, y una lista así deja de leerse. Va aparte, con su número.
3. **Un total no se puede trabajar.** El texto termina con los ocho más
   quietos, con nombre y teléfono, que es lo que sí se puede hacer.

**Dónde se ve:** en **Hoy**, arriba de las alarmas —las alarmas dicen
qué se rompió, esto dice a quién llamar— y en **Cierre de mes**, junto
al semáforo de la semana que se elija, los dos imprimibles.

**El texto lo arma el servidor**, no la pantalla: es el mismo que va al
correo. Si la pantalla lo rearmara, el día que uno de los dos cambie
dirían cosas distintas del mismo día y no habría a cuál creerle.

### 23 · El puente Central ↔ Soul: revisado a fondo

Ella preguntó: *«los proyectos en Nova Central y los pendientes que
coloqué… a cada pendiente le puse el proyecto pero no sé si ya hilan y
toman info de Nova Central o qué»*.

Es una pregunta justa, porque el puente es **invisible**: un proyecto
vive en el libro de Central y un pendiente en el de Soul, y lo único que
los une es un `trabajo_id` en una celda. Si ese hilo se corta, nada
falla — el pendiente sale sin nombre de proyecto y las horas fijas dejan
de descontarse, y las dos pantallas siguen viéndose perfectas mientras
mienten.

`pruebas/puente.js` lo recorre en los dos sentidos y **sí funciona**:

| | |
|---|---|
| **Central → Soul** | el pendiente trae el nombre del proyecto y su tipo; las horas fijas del proyecto entran al cálculo de si la semana cabe; y una tarea de un proyecto con horas fijas NO se cuenta dos veces |
| **Soul → Central** | cada proyecto muestra cuántas tareas abiertas tiene, cuántas horas y cuál es la próxima entrega |
| **La pared** | Central no ve el texto de ninguna tarea. La prueba busca las cadenas en toda la respuesta |

Y las tres formas de romperse en silencio, todas cubiertas: un
`trabajo_id` de un proyecto borrado (el pendiente queda sin nombre pero
conserva el id, para poder arreglarlo), un proyecto cerrado (deja de
descontar horas) y Central caída (Soul abre igual, sin horas fijas
inventadas).

**Una trampa encontrada al escribir la prueba, y caí yo mismo en ella:**
la hoja guarda `trabajo_id` pero todo lo que Nova DEVUELVE lo llama
`trabajoId`. Mandé el segundo y la tarea quedó sin proyecto sin que nada
fallara. Ahora el servidor **acepta los dos nombres**. Un enlace que se
pierde en silencio no se nota hasta que las horas dejan de cuadrar.

### 24 · Varias veces por semana, en una sola fila

Sus palabras: *«hay horarios que son varias veces por semana, también
arregla eso, por ejemplo […] las 6 veces que debo de hacer ejercicio en
la semana»*.

Antes **una fila era un día**: el gimnasio seis veces eran seis turnos, y
cambiar la hora había que cambiarla seis veces.

Ahora `dia_semana` acepta `1,3,5`, `1-5`, `1-5,7`, `L,X,V`, `lunes y
jueves` o `diario`. En la pantalla son **casillas**, una por día. Lo que
no se entiende **se descarta en vez de volverse lunes**: poner el
gimnasio un día que ella no dijo es peor que no ponerlo.

**Un fallo que salió de ahí:** el resumen de horas por día sumaba solo
en `r.dia` —el primero—. Con una fila por día daba igual; con el
gimnasio de seis veces en una sola fila, cinco días habrían quedado
pareciendo libres. Y esa es justo la cuenta que decide si la semana
cabe.
