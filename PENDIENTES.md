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
