/**
 * ══════════════════════════════════════════════════════════════════
 *  NOVA · MODO DEMOSTRACIÓN
 * ══════════════════════════════════════════════════════════════════
 *
 * Una copia de Nova Empresarial que funciona entera sin tocar ninguna
 * hoja de cálculo. Se activa solo cuando la dirección lo pide:
 *
 *     empresarial.html?demo=1
 *
 * POR QUÉ ASÍ Y NO CON UNA TIENDA DE PRUEBA EN LA HOJA
 *
 * Un demo conectado al servidor es un demo que puede escribir. Basta un
 * parámetro mal puesto, una sesión que se cuela o un cliente que abre la
 * pestaña equivocada para que una prueba termine dentro de los datos
 * reales de alguien. Aquí eso es imposible por construcción: en modo
 * demostración la pantalla NUNCA abre una conexión. `api()` queda
 * reemplazada por una función que contesta desde la memoria del
 * navegador, y la dirección del servidor no se usa ni una vez.
 *
 * "SIN MEMORIA" EN SERIO
 *
 * Todo vive en variables de JavaScript. Escribir, crear, borrar, cerrar
 * un mes, subir un archivo: todo se ve reflejado al instante y todo
 * desaparece al recargar. No se guarda nada en el servidor, ni en el
 * disco, ni en el almacenamiento del navegador. La única forma de que
 * algo sobreviva sería que el visitante lo copiara a mano.
 *
 * LOS NÚMEROS SON INVENTADOS, PERO NO CUALQUIERA
 *
 * Se generan con una semilla fija, así que son los mismos en cada
 * visita y dos personas mirando la misma pantalla ven lo mismo. Y están
 * calculados con las mismas reglas que usa Nova de verdad —el costo se
 * cuenta cuando se incurrió, la tasa de entrega se mide sobre lo
 * resuelto, el cobro de devolución sale de la cartera— porque un demo
 * que enseña cuentas distintas a las reales enseña mal.
 *
 * La tienda se llama Luma y no existe. Es a propósito: nadie debe poder
 * confundir estas cifras con las de una tienda de verdad.
 */
(function () {
  'use strict';

  /**
   * Dos formas de encenderlo.
   *
   * `?demo=1` en la dirección, cuando Nova está publicada en algún lado.
   * Y la bandera, que la pone el archivo de un solo pedazo: ese no tiene
   * dirección que mirar —se abre desde el disco, o llega por correo— y
   * si dependiera del `?demo` no arrancaría nunca.
   */
  if (!/[?&]demo\b/.test(location.search) && !window.NOVA_DEMO_FORZADO) return;

  // ─── 1 · AZAR REPETIBLE ───────────────────────────────────────
  /**
   * Mismo número siempre, dada la misma semilla.
   *
   * Si el demo usara Math.random, cada recarga cambiaría las cifras y
   * nadie podría decir "mira el 12 de agosto" por teléfono. Un demo que
   * no se puede señalar no se puede explicar.
   */
  function dado(semilla) {
    let a = semilla >>> 0;
    return function () {
      a = (a + 0x6D2B79F5) >>> 0;
      let t = a;
      t = Math.imul(t ^ (t >>> 15), t | 1);
      t ^= t + Math.imul(t ^ (t >>> 7), t | 61);
      return ((t ^ (t >>> 14)) >>> 0) / 4294967296;
    };
  }

  const dosDec = function (n) { return Math.round(n * 100) / 100; };

  // ─── 2 · EL CALENDARIO DEL DEMO ───────────────────────────────
  /**
   * El demo se mueve con el reloj de quien lo abre.
   *
   * Si los meses estuvieran escritos a mano, el demo envejecería: en
   * enero mostraría "septiembre" como mes en curso y parecería
   * abandonado. Así, quien lo abra siempre ve su propio mes a medias y
   * los cuatro anteriores completos, que es como se ve una operación de
   * verdad.
   */
  const HOY = new Date();
  const HOY_ISO = iso(HOY);

  function iso(d) {
    return d.getFullYear() + '-' + dd(d.getMonth() + 1) + '-' + dd(d.getDate());
  }
  function dd(n) { return (n < 10 ? '0' : '') + n; }
  function mesDe(isoStr) { return String(isoStr).slice(0, 7); }
  function mesMenos(mes, n) {
    const a = +mes.slice(0, 4), m = +mes.slice(5, 7) - 1 - n;
    const d = new Date(Date.UTC(a, m, 1));
    return d.getUTCFullYear() + '-' + dd(d.getUTCMonth() + 1);
  }
  function diasDelMes(mes) {
    return new Date(Date.UTC(+mes.slice(0, 4), +mes.slice(5, 7), 0)).getUTCDate();
  }
  function sumarDias(isoStr, n) {
    const d = new Date(isoStr + 'T00:00:00Z');
    d.setUTCDate(d.getUTCDate() + n);
    return d.toISOString().slice(0, 10);
  }
  function diasEntre(a, b) {
    return Math.floor((new Date(b + 'T00:00:00Z') - new Date(a + 'T00:00:00Z')) / 86400000);
  }

  const MES_ACTUAL = mesDe(HOY_ISO);

  /**
   * Ocho meses, y no tres.
   *
   * El recuento del inicio compara un trimestre contra el trimestre
   * anterior, y el histórico mira ocho meses atrás. Con pocos meses
   * sembrados, el trimestre anterior quedaba a medias y el demo anunciaba
   * cosas como "ingresos +222%" comparando tres meses contra uno. La
   * cuenta era correcta; la lectura, absurda. Sembrar el histórico
   * completo cuesta memoria y arregla las dos pantallas.
   */
  const N_MESES = 8;
  const MESES = [];
  for (let i = N_MESES - 1; i >= 0; i--) MESES.push(mesMenos(MES_ACTUAL, i));

  // ─── 3 · LA TIENDA QUE NO EXISTE ──────────────────────────────
  const TIENDAS = {
    ec: { id: 'ec', nombre: 'Luma EC', sub: 'Shopify + Dropi · USD',
          moneda: 'USD', pais: 'Ecuador',
          ticket: 26, costo: 8.4, flete: 2.9, cobroDev: 5.1, cpa: 8.2, entrega: 0.70 },
    gt: { id: 'gt', nombre: 'Luma GT', sub: 'Dropi · GTQ',
          moneda: 'GTQ', pais: 'Guatemala',
          ticket: 195, costo: 62, flete: 22, cobroDev: 38, cpa: 55, entrega: 0.65 },
  };

  const CATALOGO = [
    { sku: 'LUM-01', nombre: 'Colágeno Luma 300g',    cat: 'suplementos', peso: 0.34 },
    { sku: 'LUM-02', nombre: 'Serum Facial Noche',    cat: 'cuidado facial', peso: 0.26 },
    { sku: 'LUM-03', nombre: 'Faja Moldeadora Alta',  cat: 'moda', peso: 0.20 },
    { sku: 'LUM-04', nombre: 'Kit Cejas Perfectas',   cat: 'maquillaje', peso: 0.12 },
    { sku: 'LUM-05', nombre: 'Aceite de Romero 120ml', cat: 'cabello', peso: 0.08 },
  ];

  const CIUDADES = {
    ec: [['Guayaquil','Guayas'],['Quito','Pichincha'],['Cuenca','Azuay'],
         ['Machala','El Oro'],['Ambato','Tungurahua'],['Manta','Manabí']],
    gt: [['Ciudad de Guatemala','Guatemala'],['Quetzaltenango','Quetzaltenango'],
         ['Escuintla','Escuintla'],['Cobán','Alta Verapaz'],['Antigua','Sacatepéquez']],
  };

  const TRANSPORTADORAS = {
    ec: ['Servientrega', 'Laar Courier', 'Speed'],
    gt: ['Cargo Express', 'Guatex', 'Forza'],
  };

  const NOMBRES = ['María Fernanda Ruiz','Jorge Andrés Peña','Lucía Villalba',
    'Carlos Mendoza','Ana Sofía Cortés','Diego Salgado','Paola Jiménez',
    'Rubén Castillo','Daniela Orozco','Héctor Ramírez','Valeria Nieto',
    'Andrés Montoya','Camila Estrada','Óscar Herrera','Natalia Rojas'];

  const EQUIPO = [
    { id: 'eq1', nombre: 'Elena Duarte',  correo: 'elena@luma.demo',
      rol: 'dueno',   tienda: '*',  estado: 'activo', permisos: '' },
    { id: 'eq2', nombre: 'Sara Moreno',   correo: 'sara@luma.demo',
      rol: 'gestora', tienda: 'ec', estado: 'activo', permisos: '' },
    { id: 'eq3', nombre: 'Kevin Aguilar', correo: 'kevin@luma.demo',
      rol: 'gestora', tienda: 'gt', estado: 'activo', permisos: '' },
    { id: 'eq4', nombre: 'Marcela Ríos',  correo: 'marcela@luma.demo',
      rol: 'admin',   tienda: '*',  estado: 'activo', permisos: 'subir_pauta' },
  ];

  const GESTORAS = { ec: ['Sara Moreno', 'Elena Duarte'], gt: ['Kevin Aguilar', 'Elena Duarte'] };

  // ─── 4 · LAS HOJAS, EN MEMORIA ────────────────────────────────
  /**
   * Cada "hoja" es un arreglo de objetos con exactamente los mismos
   * nombres de columna que la hoja de cálculo real. No es cosmético: es
   * lo que permite que el demo use el mismo cálculo que el servidor, y
   * que lo que se aprende aquí sirva allá.
   */
  const HOJA = {
    Pedidos: [], Novedades: [], Pauta: [], Gastos: [], Cartera: [],
    Inventario: [], Equipo: EQUIPO.slice(), CAS: [], Cierres: [],
    Auditoria: [], Fuentes: [],
  };

  const PARAMS = {
    ec: {}, gt: {},
  };

  const UMBRALES_BASE = {
    dias_sin_mover: 3, horas_novedad: 24, efectividad_min: 65,
    devoluciones_max: '', cpa_aviso_pct: 85, cpa_subida_pct: 25,
    stock_dias_min: '', alarmas_a: '', alarmas_hora: 7,
  };

  const AJUSTES_BASE = {
    canal_nombre: 'Grupo de pedidos', canal_url: 'https://chat.whatsapp.com/demo',
    formato_fecha: 'dia_primero',
    preguntas_producto: '¿Para qué sirve?|¿Cómo se usa?|' +
                        '¿En cuánto tiempo se ven resultados?|¿Tiene contraindicaciones?',
    retiro_pct: 3, comision_intl_pct: 2, comision_intl_a: 'meta',
  };

  function umbrales(t) {
    const o = {};
    Object.keys(UMBRALES_BASE).forEach(function (k) {
      o[k] = PARAMS[t] && PARAMS[t][k] !== undefined ? PARAMS[t][k] : UMBRALES_BASE[k];
    });
    return o;
  }
  function ajustes(t) {
    const o = {};
    Object.keys(AJUSTES_BASE).forEach(function (k) {
      o[k] = PARAMS[t] && PARAMS[t][k] !== undefined ? PARAMS[t][k] : AJUSTES_BASE[k];
    });
    return o;
  }

  // ─── 5 · SEMBRAR LA OPERACIÓN ─────────────────────────────────
  /**
   * Cinco meses de una tienda que va razonablemente bien pero no
   * perfecta: 66% de entrega en Ecuador, 61% en Guatemala, y un CPA que
   * sube en los últimos meses.
   *
   * Que suba es a propósito. Un demo donde todo mejora no enseña a usar
   * la herramienta —lo que hay que saber leer es el mes en que algo se
   * empieza a torcer, y eso es justo lo que Nova está para avisar.
   */
  let SEQ = 0;
  function nid(pre) { return pre + (++SEQ).toString(36).padStart(4, '0'); }

  function sembrar() {
    Object.keys(TIENDAS).forEach(function (tid, ti) {
      const T = TIENDAS[tid];
      const r = dado(9187 + ti * 1000);

      MESES.forEach(function (mes, mi) {
        const ultimo = mes === MES_ACTUAL ? HOY.getDate() : diasDelMes(mes);

        /**
         * El CPA se va deteriorando mes a mes: empieza 12% abajo y
         * termina 10% arriba.
         *
         * Es la historia que el demo cuenta. Un demo donde todo mejora no
         * enseña a usar la herramienta: lo que hay que aprender a leer es
         * el mes en que algo se empieza a torcer, y eso es justo lo que
         * Nova está para avisar.
         */
        const deriva = 0.88 + mi * (0.22 / (N_MESES - 1));
        const porDia = tid === 'ec' ? 9 : 6;

        for (let dia = 1; dia <= ultimo; dia++) {
          const fecha = mes + '-' + dd(dia);
          const cuantos = Math.max(1, Math.round(porDia * (0.6 + r() * 0.85)));

          for (let k = 0; k < cuantos; k++) {
            const prod = CATALOGO[Math.floor(r() * CATALOGO.length)];
            const ciudad = CIUDADES[tid][Math.floor(r() * CIUDADES[tid].length)];
            const trans = TRANSPORTADORAS[tid][Math.floor(r() * TRANSPORTADORAS[tid].length)];
            const gest = GESTORAS[tid][r() < 0.78 ? 0 : 1];

            // Promociones: la mayoría lleva una unidad, algunas dos o tres
            const u = r() < 0.70 ? 1 : (r() < 0.75 ? 2 : 3);
            const valor = dosDec(T.ticket * (u === 1 ? 1 : u === 2 ? 1.72 : 2.35)
                                 * (0.93 + r() * 0.14));
            const cProd = dosDec(T.costo * u * (0.96 + r() * 0.08));
            const cEnv = dosDec(T.flete * (0.9 + r() * 0.25));

            // El desenlace depende de cuánto tiempo ha pasado
            const edad = diasEntre(fecha, HOY_ISO);
            let est, fEntrega = '', ultMov = fecha;

            if (edad < 2) {
              est = r() < 0.5 ? 'pendiente' : 'confirmado';
              ultMov = fecha;
            } else if (edad < 4) {
              const x = r();
              est = x < 0.45 ? 'entregado' : x < 0.55 ? 'novedad'
                  : x < 0.72 ? 'en_transito' : 'en_bodega';
              if (est === 'entregado') fEntrega = sumarDias(fecha, 2);
              ultMov = sumarDias(fecha, Math.min(edad, 1 + Math.floor(r() * 3)));
            } else {
              const x = r();
              if (x < T.entrega) {
                est = 'entregado';
                fEntrega = sumarDias(fecha, 2 + Math.floor(r() * 3));
                ultMov = fEntrega;
              } else if (x < T.entrega + 0.26) {
                est = 'devolucion';
                ultMov = sumarDias(fecha, 6 + Math.floor(r() * 8));
              } else if (x < T.entrega + 0.31) {
                est = 'cancelado';
                ultMov = sumarDias(fecha, 1 + Math.floor(r() * 2));
              } else {
                // Los que se quedaron trabados: la materia prima del CAS
                est = r() < 0.5 ? 'en_transito' : 'en_oficina';
                ultMov = sumarDias(fecha, 2 + Math.floor(r() * 4));
              }
            }
            if (ultMov > HOY_ISO) ultMov = HOY_ISO;

            const pid = nid('p');
            HOJA.Pedidos.push({
              id: pid, fuente: 'dropi', id_externo: String(400000 + SEQ),
              fecha: fecha, tienda: tid,
              cliente: NOMBRES[Math.floor(r() * NOMBRES.length)],
              cedula: '', correo: '', telefono: '09' + Math.floor(10000000 + r() * 89999999),
              telefono_norm: '', telefono_2: '', telefono_2_norm: '',
              ciudad: ciudad[0], departamento: ciudad[1], direccion: '',
              producto: prod.nombre, sku: prod.sku, cantidad: u,
              valor: valor, costo_producto: cProd, costo_envio: cEnv,
              metodo_pago: 'contraentrega', bodega: '',
              estado: est, estado_transportadora: '', estado_canonico: est,
              transportadora: trans, guia: 'G' + (900000 + SEQ),
              intentos: '', gestora_asignada: gest,
              fecha_promesa: '', fecha_entrega: fEntrega,
              razon_cancelacion: '', estado_nova: '', nota: '',
              ultimo_movimiento: ultMov,
              adelanto: '', acuerdo_oficina: '', confirmado_oficina: '',
              actualizado_en: ultMov, actualizado_por: 'dropi',
            });

            // Una novedad por cada pedido que pasó por ahí
            if (est === 'novedad' || (est === 'devolucion' && r() < 0.45)) {
              const motivos = [
                ['No contesta el teléfono', 'no_contacta'],
                ['Cliente rechaza el pedido', 'rechaza'],
                ['Dirección incorrecta', 'direccion'],
                ['No tiene el dinero completo', 'dinero'],
                ['Reprogramar entrega', 'reprograma'],
              ];
              const mo = motivos[Math.floor(r() * motivos.length)];
              const resuelta = est !== 'novedad' || r() < 0.55;
              HOJA.Novedades.push({
                id: nid('n'), fuente: 'dropi', id_externo: '', pedido_id: pid,
                fecha: sumarDias(fecha, 1 + Math.floor(r() * 3)),
                tipo: 'novedad', motivo: mo[0], grupo: mo[1],
                estado: resuelta ? 'resuelta' : 'abierta',
                solucionada: resuelta ? 'si' : 'no', fecha_solucion: '',
                desenlace: est === 'entregado' ? 'entregado'
                         : est === 'devolucion' ? 'devolucion' : '',
                gestora: gest, solucion: '', nota: '', intentos: 1 + Math.floor(r() * 2),
                resuelta_en: '', actualizado_en: fecha, actualizado_por: 'dropi',
              });
            }
          }

          // Pauta del día
          const entregadosEse = 1;
          const gasto = dosDec(T.cpa * deriva * cuantos * T.entrega * (0.8 + r() * 0.45));
          ['Conjunto frío · video', 'Conjunto retargeting', 'Conjunto prueba creativo']
            .forEach(function (conj, ci) {
              const parte = ci === 0 ? 0.6 : ci === 1 ? 0.28 : 0.12;
              HOJA.Pauta.push({
                id: nid('a'), fecha: fecha, fecha_fin: '', tienda: tid,
                plataforma: 'meta', cuenta: 'Luma ' + tid.toUpperCase(),
                campana: 'Ventas ' + mes, conjunto: conj,
                entrega: 'activa', presupuesto: '',
                gasto: dosDec(gasto * parte), moneda_gasto: T.moneda,
                gasto_normalizado: dosDec(gasto * parte),
                impresiones: Math.round(gasto * parte * 120),
                alcance: Math.round(gasto * parte * 80), frecuencia: 1.4,
                clics: Math.round(gasto * parte * 3.2), ctr: 2.1,
                cpc: 0.31, cpm: 8.4,
                resultados: Math.round(cuantos * parte), compras: '',
                cpa: '', roas: '', valor_conv: '', visitas_lp: '',
              });
              void entregadosEse;
            });
        }

        // ── Cartera: el extracto de la billetera ──
        /**
         * Aquí está la lección más cara de todo el demo.
         *
         * El export de pedidos trae un flete de lista. La cartera trae lo
         * que la plataforma cobró de verdad por devolver, que es más
         * alto. Nova usa el promedio real de la cartera, y por eso el
         * cierre del demo no cuadra con una cuenta hecha a ojo desde el
         * export — igual que en la vida real.
         */
        const delMes = HOJA.Pedidos.filter(function (p) {
          return p.tienda === tid && mesDe(p.fecha) === mes;
        });
        delMes.forEach(function (p) {
          if (p.estado_canonico === 'entregado') {
            HOJA.Cartera.push({
              id: nid('c'), fuente: 'dropi_cartera', tienda: tid,
              fecha: p.fecha_entrega || p.fecha, tipo: 'ENTRADA', clase: 'ganancia',
              monto: dosDec(p.valor - p.costo_producto - p.costo_envio),
              saldo_previo: '', orden_id: p.id_externo, guia: p.guia,
              descripcion: 'ENTRADA POR GANANCIA DE ORDEN', cuenta: '',
              concepto_retiro: '', importado_en: HOY_ISO,
            });
          }
          if (p.estado_canonico === 'devolucion') {
            HOJA.Cartera.push({
              id: nid('c'), fuente: 'dropi_cartera', tienda: tid,
              fecha: p.ultimo_movimiento, tipo: 'SALIDA', clase: 'devolucion',
              monto: dosDec(-T.cobroDev * (0.88 + r() * 0.24)),
              saldo_previo: '', orden_id: p.id_externo, guia: p.guia,
              descripcion: 'SALIDA DE COBRO DE DEVOLUCIÓN', cuenta: '',
              concepto_retiro: '', importado_en: HOY_ISO,
            });
          }
        });

        // Retiros: dos por mes, con su comisión implícita
        [8, 22].forEach(function (dia) {
          if (dia > ultimo) return;
          const monto = tid === 'ec' ? 1400 + r() * 900 : 9000 + r() * 6000;
          HOJA.Cartera.push({
            id: nid('c'), fuente: 'dropi_cartera', tienda: tid,
            fecha: mes + '-' + dd(dia), tipo: 'SALIDA', clase: 'retiro',
            monto: -dosDec(monto), saldo_previo: '', orden_id: '', guia: '',
            descripcion: 'SALIDA POR RETIRO A CUENTA BANCARIA', cuenta: '****4417',
            concepto_retiro: 'Retiro programado', importado_en: HOY_ISO,
          });
        });
      });

      // ── Gastos fijos ──
      const fijos = tid === 'ec'
        ? [['Gestora (medio tiempo)', 420], ['Shopify', 39], ['Central telefónica', 55]]
        : [['Gestor (medio tiempo)', 3100], ['Central telefónica', 410]];
      fijos.forEach(function (g) {
        HOJA.Gastos.push({
          id: nid('g'), tienda: tid, mes: '', tipo: 'fijo', nombre: g[0],
          valor: g[1], moneda: TIENDAS[tid].moneda, nota: '', activo: 'si',
          actualizado_en: HOY_ISO, actualizado_por: 'elena@luma.demo',
        });
      });

      // ── Inventario: fichas completas, menos una a medias ──
      /**
       * La última ficha va sin costo a propósito. Nova no inventa un
       * margen cuando falta el costo: lo dice. Que el demo tenga un
       * producto así es lo que deja ver ese comportamiento.
       */
      CATALOGO.forEach(function (prod, i) {
        const T2 = TIENDAS[tid];
        const completo = i < CATALOGO.length - 1;
        HOJA.Inventario.push({
          id: nid('i'), sku: prod.sku, producto: prod.nombre, tienda: tid,
          fuente: 'manual', origen: 'proveedor', categoria: prod.cat,
          proveedor: completo ? 'Distribuidora Andina' : '',
          landing: completo ? 'https://luma.demo/' + prod.sku.toLowerCase() : '',
          stock: 40 + Math.floor(dado(i + 7)() * 160),
          costo_unitario: completo ? dosDec(T2.costo * (0.85 + i * 0.09)) : '',
          precio: dosDec(T2.ticket * (0.9 + i * 0.07)),
          precio_2: dosDec(T2.ticket * (0.9 + i * 0.07) * 1.72),
          precio_3: dosDec(T2.ticket * (0.9 + i * 0.07) * 2.35),
          minimo: 25, dias_cobertura: '', ultimo_conteo: '',
          resp_1: completo ? 'Aporta colágeno hidrolizado que el cuerpo absorbe rápido; ayuda a piel, uñas y articulaciones.' : '',
          resp_2: completo ? 'Una medida rasa en agua o jugo, en ayunas o antes de dormir. No necesita batidora.' : '',
          resp_3: completo ? 'Las primeras señales entre la segunda y la tercera semana; el cambio firme, al segundo frasco.' : '',
          resp_4: completo ? 'No en embarazo ni lactancia. Si toma medicación diaria, consultar primero.' : '',
          nota: '', activo: 'si', actualizado_en: HOY_ISO,
          actualizado_por: 'elena@luma.demo',
        });
      });

      // ── Un mes ya cerrado, para que se vea qué significa ──
      const mesViejo = MESES[0];
      HOJA.Cierres.push({
        tienda: tid, mes: mesViejo, estado: 'cerrado',
        cerrado_en: mesViejo + '-28', cerrado_por: 'elena@luma.demo',
        snapshot: '', _datos: null,
      });

      // ── Qué se ha importado ──
      [['dropi', 'pedidos'], ['meta', 'pauta'], ['dropi_cartera', 'cartera']]
        .forEach(function (f) {
          HOJA.Fuentes.push({
            tienda: tid, fuente: f[0], tipo: f[1],
            ultima: sumarDias(HOY_ISO, -1) + ' 08:40',
            filas: f[0] === 'dropi' ? 1240 : f[0] === 'meta' ? 450 : 610,
          });
        });
    });

    // El cierre congelado se calcula una vez sembrado todo
    HOJA.Cierres.forEach(function (c) {
      c._datos = mesDe_(c.tienda, c.mes, 'dueno');
    });
  }

  // ─── 6 · EL CÁLCULO, IGUAL QUE EN EL SERVIDOR ─────────────────
  function num(v) { const n = Number(v); return isFinite(n) ? n : 0; }
  function norm(s) {
    return String(s == null ? '' : s).trim().toLowerCase()
      .normalize('NFD').replace(/[̀-ͯ]/g, '');
  }

  const TERMINALES = ['entregado', 'devolucion', 'cancelado'];

  /**
   * El mes, calculado una sola vez.
   *
   * agregarMes recorre TODOS los pedidos. Con los ocho meses sembrados eso
   * no se nota, pero el demo está para que alguien suba su histórico de
   * verdad —hay tiendas con cuatro años— y ahí sí: una sola apertura de
   * pantalla llama a agregarMes unas treinta veces entre el resumen, el
   * cierre, el histórico, el recuento y las alarmas. Con cincuenta mil
   * pedidos son millón y medio de vueltas para responder lo mismo.
   *
   * Se devuelve una copia y no el original porque quien llama le agrega
   * cosas encima —el resumen le cuelga recaudo7, el cierre lo congela— y
   * eso ensuciaría lo guardado para el siguiente.
   */
  const CACHE_MES = {};
  let SELLO_CACHE = 0;

  function invalidarCache() { SELLO_CACHE++; Object.keys(CACHE_MES).forEach(function (k) { delete CACHE_MES[k]; }); }

  function copia(o) {
    try { return structuredClone(o); }
    catch (e) { return JSON.parse(JSON.stringify(o)); }
  }

  function mesDe_(tienda, mes, rol) {
    const k = tienda + '|' + mes + '|' + rol;
    if (!CACHE_MES[k]) CACHE_MES[k] = agregarMesCrudo(tienda, mes, rol);
    return copia(CACHE_MES[k]);
  }

  /**
   * La cartera del mes, resumida.
   *
   * `devolucionPromedio` es lo único que se usa para corregir el costo, y
   * se usa como promedio y no como total a propósito: los cobros de un
   * mes incluyen devoluciones de pedidos del mes anterior. El promedio sí
   * representa lo que cuesta una devolución; el total sería de otro
   * conjunto de pedidos.
   */
  function carteraDelMes(tienda, mes) {
    const filas = HOJA.Cartera.filter(function (x) {
      return x.tienda === tienda && mesDe(x.fecha) === mes;
    });
    const out = { hay: filas.length > 0, ganancia: 0, devoluciones: 0, fletes: 0,
                  retiros: 0, conceptosRetiro: {}, netoOperativo: 0, saldo: 0,
                  devolucionPromedio: 0, costoRetiros: 0, retiroPct: 0, nDevoluciones: 0 };
    if (!out.hay) return out;

    filas.forEach(function (x) {
      const m = num(x.monto);
      if (x.clase === 'ganancia') out.ganancia += m;
      else if (x.clase === 'devolucion') { out.devoluciones += Math.abs(m); out.nDevoluciones++; }
      else if (x.clase === 'flete') out.fletes += Math.abs(m);
      else if (x.clase === 'retiro') {
        out.retiros += Math.abs(m);
        const k = x.concepto_retiro || 'Sin concepto';
        out.conceptosRetiro[k] = (out.conceptosRetiro[k] || 0) + Math.abs(m);
      }
      out.saldo += m;
    });

    out.netoOperativo = out.ganancia - out.devoluciones - out.fletes;
    out.devolucionPromedio = out.nDevoluciones ? out.devoluciones / out.nDevoluciones : 0;
    out.retiroPct = num(ajustes(tienda).retiro_pct);
    out.costoRetiros = out.retiros * (out.retiroPct / 100);
    return out;
  }

  /**
   * Un mes entero, con las mismas reglas que usa Nova de verdad.
   *
   * Las tres que más se equivocan a mano, y que aquí están escritas:
   *  · el costo del producto solo cuenta si se entregó
   *  · el flete cuenta si salió de bodega, entregado o devuelto
   *  · lo cancelado y lo pendiente no cuestan nada todavía
   */
  function agregarMesCrudo(tienda, mes, rol) {
    const out = {
      pedidos: 0, despachados: 0, entregados: 0, devueltos: 0, cancelados: 0,
      pendientes: 0, ventas: 0, costoProducto: 0, costoEnvio: 0,
      costoDevolucion: 0, valorAbierto: 0,
      costosPorEstado: {
        producto: { entregado: 0, devolucion: 0, cancelado: 0, pendiente: 0 },
        envio:    { entregado: 0, devolucion: 0, cancelado: 0, pendiente: 0 },
      },
      novedades: 0, sinMover: 0, grupos: {}, transportadoras: {}, productos: {},
    };

    HOJA.Pedidos.forEach(function (f) {
      if (f.tienda !== tienda) return;
      if (mesDe(f.fecha) !== mes) return;

      out.pedidos++;
      const est = norm(f.estado_nova || f.estado_canonico || f.estado);
      if (est === 'entregado') { out.entregados++; out.ventas += num(f.valor); }
      if (est === 'devolucion') out.devueltos++;
      if (est === 'cancelado') out.cancelados++;
      if (['cancelado', 'pendiente'].indexOf(est) === -1) out.despachados++;

      const cProd = num(f.costo_producto), cEnv = num(f.costo_envio);
      const grupo = TERMINALES.indexOf(est) !== -1 ? est : 'pendiente';
      out.costosPorEstado.producto[grupo] += cProd;
      out.costosPorEstado.envio[grupo] += cEnv;

      if (est === 'entregado') out.costoProducto += cProd;
      if (['cancelado', 'pendiente'].indexOf(est) === -1) out.costoEnvio += cEnv;
      if (est === 'devolucion') out.costoDevolucion += cEnv;

      const prod = String(f.producto || 'Sin producto').trim();
      if (!out.productos[prod]) out.productos[prod] = { pedidos: 0, entregados: 0, ventas: 0 };
      out.productos[prod].pedidos++;
      if (est === 'entregado') {
        out.productos[prod].entregados++;
        out.productos[prod].ventas += num(f.valor);
      }

      const t = String(f.transportadora || '').trim();
      if (t) {
        if (!out.transportadoras[t]) out.transportadoras[t] = { n: 0, entregados: 0 };
        out.transportadoras[t].n++;
        if (est === 'entregado') out.transportadoras[t].entregados++;
      }

      if (TERMINALES.indexOf(est) === -1) {
        out.pendientes++;
        out.valorAbierto += num(f.valor);
        const ult = f.ultimo_movimiento || f.actualizado_en || f.fecha;
        if (diasEntre(ult, HOY_ISO) > 3) out.sinMover++;
      }
    });

    HOJA.Novedades.forEach(function (n) {
      if (mesDe(n.fecha) !== mes) return;
      const ped = HOJA.Pedidos.filter(function (p) { return p.id === n.pedido_id; })[0];
      if (ped && ped.tienda !== tienda) return;
      out.novedades++;
      const g = n.grupo || 'otro';
      out.grupos[g] = (out.grupos[g] || 0) + 1;
    });

    out.resueltos = out.entregados + out.devueltos;
    out.efectividad = out.resueltos ? out.entregados / out.resueltos * 100 : 0;
    out.efectividadDespacho = out.despachados ? out.entregados / out.despachados * 100 : 0;
    out.tasaDevolucion = out.despachados ? out.devueltos / out.despachados * 100 : 0;
    out.ticket = out.entregados ? out.ventas / out.entregados : 0;

    // La plata solo se le muestra a la dueña, igual que en el servidor
    if (rol === 'dueno') {
      out.gasto = 0; out.campanas = {}; out.gastoPorDia = {};
      out.gastoSinConvertir = 0; out.monedasSinTasa = {};
      HOJA.Pauta.forEach(function (f) {
        if (f.tienda !== tienda || mesDe(f.fecha) !== mes) return;
        const g = num(f.gasto);
        out.gasto += g;
        const nom = f.conjunto || f.campana || 'Sin nombre';
        if (!out.campanas[nom]) out.campanas[nom] = { gasto: 0, resultados: 0, sinTasa: 0 };
        out.campanas[nom].gasto += g;
        out.campanas[nom].resultados += num(f.resultados);
        out.gastoPorDia[f.fecha] = (out.gastoPorDia[f.fecha] || 0) + g;
      });

      out.fijos = 0; out.detalleFijos = [];
      HOJA.Gastos.forEach(function (f) {
        if (f.tienda !== tienda) return;
        if (norm(f.activo) === 'no') return;
        if (f.mes && f.mes !== mes) return;
        const v = num(f.valor);
        out.fijos += v;
        out.detalleFijos.push({ nombre: f.nombre, valor: v, tipo: f.tipo || 'fijo',
                                recurrente: !f.mes, id: f.id });
      });

      out.cartera = carteraDelMes(tienda, mes);
      out.costoDevolucionEstimado = out.costoDevolucion;
      if (out.cartera.hay && out.cartera.devolucionPromedio && out.devueltos) {
        out.costoDevolucion = out.cartera.devolucionPromedio * out.devueltos;
        out.costoDevolucionFuente = 'cartera';
      } else {
        out.costoDevolucionFuente = 'export';
      }

      out.cpa = out.entregados ? out.gasto / out.entregados : 0;
      out.roas = out.gasto ? out.ventas / out.gasto : 0;

      const cobroRetorno = out.costoDevolucionFuente === 'cartera' ? out.costoDevolucion : 0;
      const aj = ajustes(tienda);
      out.comisionRetiro = (out.cartera && out.cartera.costoRetiros) || 0;
      const pct = num(aj.comision_intl_pct);
      const plats = String(aj.comision_intl_a || '').split(/[,;]/).map(norm).filter(String);
      out.comisionIntl = (pct && plats.indexOf('meta') !== -1) ? out.gasto * (pct / 100) : 0;
      out.comisiones = out.comisionRetiro + out.comisionIntl;
      out.margen = out.ventas - out.gasto - out.costoProducto - out.costoEnvio - cobroRetorno;
      out.utilidad = out.margen - out.fijos - out.comisiones;
    }

    return out;
  }

  function recaudoUltimosDias(tienda, dias) {
    const serie = [], idx = {};
    for (let i = dias - 1; i >= 0; i--) {
      const f = sumarDias(HOY_ISO, -i);
      idx[f] = serie.length;
      serie.push({ fecha: f, total: 0, entregas: 0 });
    }
    const out = { serie: serie, sinFecha: 0, pedidosSinFecha: 0,
                  moneda: TIENDAS[tienda].moneda };
    const desde = serie[0].fecha;
    HOJA.Pedidos.forEach(function (f) {
      if (f.tienda !== tienda) return;
      if (norm(f.estado_nova || f.estado_canonico) !== 'entregado') return;
      const e = f.fecha_entrega;
      if (!e) {
        if (f.fecha >= desde) { out.sinFecha += num(f.valor); out.pedidosSinFecha++; }
        return;
      }
      if (idx[e] === undefined) return;
      serie[idx[e]].total += num(f.valor);
      serie[idx[e]].entregas++;
    });
    return out;
  }

  // ─── 7 · QUIÉN ESTÁ MIRANDO ───────────────────────────────────
  const PERMISOS_TODOS = ['subir_pedidos', 'subir_novedades', 'subir_pauta', 'subir_cartera'];
  const PERMISOS_ROL = {
    admin:   ['subir_pedidos', 'subir_novedades'],
    gestora: ['subir_pedidos', 'subir_novedades'],
  };

  let SES = null;

  function abrirSesion(rol) {
    const p = EQUIPO.filter(function (x) { return x.rol === rol; })[0];
    const permisos = rol === 'dueno' ? PERMISOS_TODOS.slice()
      : (PERMISOS_ROL[rol] || []).concat(
          String(p.permisos || '').split(/[,;]/).map(norm).filter(function (x) {
            return PERMISOS_TODOS.indexOf(x) !== -1;
          }));
    const tiendas = rol === 'gestora'
      ? [p.tienda]
      : Object.keys(TIENDAS);

    SES = {
      email: p.correo, nombre: p.nombre, rol: rol, rolReal: rol,
      tiendas: tiendas, permisos: permisos,
    };
    return SES;
  }

  /**
   * La sesión tal como la vería la pantalla si viniera del servidor.
   *
   * Las tiendas van como fichas completas y no como códigos sueltos: es
   * lo que hace que arriba diga "Luma EC" y no "EC", sin tocar el HTML.
   */
  function publico(vista) {
    const rol = vista || SES.rol;
    return {
      email: SES.email, nombre: SES.nombre, rol: rol,
      tiendas: SES.tiendas.map(function (id) {
        return { id: id, nombre: TIENDAS[id].nombre, sub: TIENDAS[id].sub,
                 moneda: TIENDAS[id].moneda, pais: TIENDAS[id].pais };
      }),
      modulos: ['empresarial'],
      permisos: rol === 'dueno' ? PERMISOS_TODOS.slice()
                                : (PERMISOS_ROL[rol] || []).slice(),
      // Lejos: un demo que se corta a la mitad de una explicación es un
      // demo que deja a quien lo muestra pidiendo disculpas.
      vence: Date.now() + 1000 * 60 * 60 * 12,
      horas: 12, inactividad_min: 720,
      puede_ver_como: SES.rol !== 'gestora',
      rol_real: SES.rol,
      demo: true,
    };
  }

  function rolEfectivo(p) {
    // La vista sombra nunca da MÁS permisos de los que ya se tienen
    const orden = { gestora: 1, admin: 2, dueno: 3 };
    const v = p && p.vista;
    if (!v || !SES) return SES ? SES.rol : 'gestora';
    if (orden[v] >= orden[SES.rol]) return SES.rol;
    return v;
  }

  function tiendaDe(p) {
    const t = String((p && p.tienda) || SES.tiendas[0] || '').trim();
    return TIENDAS[t] ? t : SES.tiendas[0];
  }

  function accesible(t) { return SES && SES.tiendas.indexOf(t) !== -1; }

  // ─── 8 · LAS ACCIONES ─────────────────────────────────────────
  function anotar(entidad, id, campo, antes, despues) {
    HOJA.Auditoria.unshift({
      cuando: new Date().toISOString().slice(0, 16).replace('T', ' '),
      quien: SES ? SES.email : '', entidad: entidad, entidad_id: id,
      campo: campo, antes: antes, despues: despues,
    });
  }

  const ACCIONES = {

    yo: function () {
      if (!SES) return { ok: false, error: 'Sin sesión.' };
      return { ok: true, sesion: publico() };
    },

    salir: function () { SES = null; return { ok: true }; },

    resumen: function (p) {
      const t = tiendaDe(p);
      if (!accesible(t)) return { ok: false, error: 'No tienes acceso a esa tienda.' };
      const mes = p.mes || MES_ACTUAL;
      const d = mesDe_(t, mes, rolEfectivo(p));
      d.recaudo7 = recaudoUltimosDias(t, 7);
      return { ok: true, tienda: t, mes: mes, datos: d };
    },

    listar: function (p) {
      const ent = String(p.entidad || '');
      const hoja = HOJA[ent];
      if (!hoja) return { ok: false, error: 'No existe la hoja ' + ent + '.' };
      const rol = rolEfectivo(p);
      if (ent === 'Gastos' && rol !== 'dueno') {
        return { ok: false, error: 'Solo la dueña ve ' + ent + '.' };
      }
      let filas = hoja.slice();
      if (p.tienda && ent !== 'Novedades') {
        filas = filas.filter(function (f) { return !f.tienda || f.tienda === p.tienda; });
      }
      if (ent === 'Novedades' && p.tienda) {
        const ids = {};
        HOJA.Pedidos.forEach(function (x) { if (x.tienda === p.tienda) ids[x.id] = 1; });
        filas = filas.filter(function (f) { return ids[f.pedido_id]; });
      }
      // Lo más nuevo primero, que es lo que se va a gestionar
      filas.sort(function (a, b) {
        return String(b.fecha || '').localeCompare(String(a.fecha || ''));
      });
      const desde = Math.max(0, parseInt(p.offset || 0, 10));
      const cuantas = Math.min(500, Math.max(1, parseInt(p.limite || 200, 10)));
      return {
        ok: true, total: filas.length, offset: desde,
        columnas: Object.keys(hoja[0] || {}),
        filas: filas.slice(desde, desde + cuantas).map(function (f) {
          const o = {};
          Object.keys(f).forEach(function (k) { if (k.charAt(0) !== '_') o[k] = f[k]; });
          return o;
        }),
      };
    },

    escribir: function (p) {
      const ent = String(p.entidad || '');
      const hoja = HOJA[ent];
      if (!hoja) return { ok: false, error: 'No se puede escribir en ' + ent + '.' };
      const rol = rolEfectivo(p);
      if (['Gastos', 'Equipo', 'Inventario'].indexOf(ent) !== -1 && rol !== 'dueno') {
        return { ok: false, error: 'Solo la dueña cambia ' + ent + '.' };
      }
      const fila = hoja.filter(function (f) { return String(f.id) === String(p.id); })[0];
      if (!fila) return { ok: false, error: 'No encuentro ' + ent + ' con id ' + p.id + '.' };

      const campos = p.campos || {};
      const escritos = [], rechazados = [];
      Object.keys(campos).forEach(function (k) {
        if (!(k in fila)) { rechazados.push(k); return; }
        anotar(ent, p.id, k, fila[k], campos[k]);
        fila[k] = campos[k];
        escritos.push(k);
      });
      fila.actualizado_en = HOY_ISO;
      fila.actualizado_por = SES.email;
      return { ok: true, escritos: escritos, rechazados: rechazados };
    },

    crear: function (p) {
      const ent = String(p.entidad || '');
      const hoja = HOJA[ent];
      if (!hoja) return { ok: false, error: 'No se pueden crear filas en ' + ent + '.' };
      if (rolEfectivo(p) !== 'dueno') return { ok: false, error: 'Solo la dueña.' };

      const modelo = hoja[0] || {};
      const fila = {};
      Object.keys(modelo).forEach(function (k) { if (k.charAt(0) !== '_') fila[k] = ''; });
      Object.keys(p.datos || {}).forEach(function (k) { fila[k] = p.datos[k]; });
      fila.id = nid(ent.charAt(0).toLowerCase());
      // Una fila sin tienda no la ve ningún cálculo: quedaría guardada y
      // sin efecto, que es la forma más confusa de "funcionar".
      if ('tienda' in fila && !fila.tienda) fila.tienda = tiendaDe(p);
      if ('activo' in fila && !fila.activo) fila.activo = 'si';
      if (ent === 'Equipo' && !fila.estado) fila.estado = 'activo';
      fila.actualizado_en = HOY_ISO;
      fila.actualizado_por = SES.email;
      hoja.push(fila);
      anotar(ent, fila.id, '(alta)', '', fila.nombre || fila.producto || fila.id);
      return { ok: true, id: fila.id };
    },

    borrar: function (p) {
      const ent = String(p.entidad || '');
      const hoja = HOJA[ent];
      if (!hoja) return { ok: false, error: 'No se puede borrar en ' + ent + '.' };
      if (rolEfectivo(p) !== 'dueno') return { ok: false, error: 'Solo la dueña.' };
      if (ent === 'Equipo') {
        const duenas = hoja.filter(function (x) {
          return x.rol === 'dueno' && String(x.id) !== String(p.id);
        });
        if (!duenas.length) {
          return { ok: false, error: 'No puedes quitar a la única dueña: la cuenta ' +
                   'quedaría sin nadie que pueda devolver permisos.' };
        }
      }
      const i = hoja.map(function (f) { return String(f.id); }).indexOf(String(p.id));
      if (i === -1) return { ok: false, error: 'No encuentro esa fila.' };
      anotar(ent, p.id, '(baja)', hoja[i].nombre || hoja[i].producto || p.id, '');
      hoja.splice(i, 1);
      return { ok: true };
    },

    productos: function (p) {
      const t = tiendaDe(p);
      const cat = {};
      HOJA.Pedidos.forEach(function (f) {
        if (f.tienda !== t) return;
        const k = norm(f.producto);
        if (!cat[k]) {
          cat[k] = { nombre: f.producto, sku: f.sku, pedidos: 0, entregados: 0,
                     devueltos: 0, cancelados: 0, unidades: 0, unidadesEntregadas: 0,
                     ventas: 0, primera: f.fecha, ultima: f.fecha };
        }
        const x = cat[k];
        const est = norm(f.estado_nova || f.estado_canonico);
        x.pedidos++; x.unidades += num(f.cantidad);
        if (est === 'entregado') {
          x.entregados++; x.ventas += num(f.valor);
          x.unidadesEntregadas += num(f.cantidad);
        }
        if (est === 'devolucion') x.devueltos++;
        if (est === 'cancelado') x.cancelados++;
        if (f.fecha < x.primera) x.primera = f.fecha;
        if (f.fecha > x.ultima) x.ultima = f.fecha;
      });

      const fichas = {};
      HOJA.Inventario.forEach(function (f) {
        if (f.tienda !== t) return;
        fichas[norm(f.producto)] = {
          id: f.id, sku: f.sku, nombre: f.producto, categoria: f.categoria,
          proveedor: f.proveedor, landing: f.landing, stock: num(f.stock),
          costo: f.costo_unitario === '' ? 0 : num(f.costo_unitario),
          precio: num(f.precio), precio_2: num(f.precio_2), precio_3: num(f.precio_3),
          minimo: num(f.minimo), nota: f.nota,
          resp_1: f.resp_1, resp_2: f.resp_2, resp_3: f.resp_3, resp_4: f.resp_4,
        };
      });

      const salida = Object.keys(cat).map(function (k) {
        const x = cat[k], ficha = fichas[k] || null;
        const resueltos = x.entregados + x.devueltos;
        const ventana = diasEntre(x.primera, HOY_ISO) + 1;
        const ritmo = (ventana >= 14 && x.unidadesEntregadas)
          ? x.unidadesEntregadas / ventana : null;
        const falta = [];
        if (!ficha) falta.push('ficha');
        else {
          if (!ficha.costo) falta.push('costo');
          if (!ficha.precio) falta.push('precio');
          if (!ficha.categoria) falta.push('categoría');
        }
        return {
          clave: k, nombre: x.nombre, sku: x.sku || (ficha ? ficha.sku : ''),
          pedidos: x.pedidos, entregados: x.entregados, devueltos: x.devueltos,
          cancelados: x.cancelados, unidades: x.unidades, ventas: x.ventas,
          unidadesEntregadas: x.unidadesEntregadas,
          ticket: x.entregados ? x.ventas / x.entregados : 0,
          entrega: resueltos ? x.entregados / resueltos * 100 : 0,
          primera: x.primera, ultima: x.ultima, ficha: ficha, falta: falta,
          ritmo: ritmo, ventanaDias: ventana,
          coberturaDias: (ritmo && ficha && ficha.stock > 0)
            ? Math.floor(ficha.stock / ritmo) : null,
          margen: (ficha && ficha.costo && x.entregados)
            ? (x.ventas / x.entregados) - ficha.costo : null,
        };
      }).sort(function (a, b) { return b.pedidos - a.pedidos; });

      Object.keys(fichas).forEach(function (k) {
        if (cat[k]) return;
        const fi = fichas[k];
        salida.push({ clave: k, nombre: fi.nombre, sku: fi.sku, pedidos: 0,
          entregados: 0, devueltos: 0, cancelados: 0, unidades: 0,
          unidadesEntregadas: 0, ventas: 0, ticket: 0, entrega: 0,
          primera: '', ultima: '', ficha: fi, falta: [], margen: null,
          ritmo: null, coberturaDias: null, ventanaDias: 0, sinVentas: true });
      });

      return { ok: true, tienda: t, productos: salida,
               preguntas: String(ajustes(t).preguntas_producto).split('|')
                 .map(function (x) { return x.trim(); }).filter(String).slice(0, 4),
               // La hoja del demo está siempre al día: no falta ninguna columna
               columnas: Object.keys(HOJA.Inventario[0] || {}),
               modalidad: 'catalogo_publico', moneda: TIENDAS[t].moneda };
    },

    equipo: function (p) {
      const t = tiendaDe(p);
      const mes = p.mes || MES_ACTUAL;
      const rol = rolEfectivo(p);
      const personas = HOJA.Equipo.map(function (f) {
        return { id: f.id, nombre: f.nombre, correo: f.correo, rol: f.rol,
                 rolLegible: '', tienda: f.tienda || '*',
                 estado: f.estado || 'activo', permisos: f.permisos || '',
                 ultima_conexion: '', pedidos: 0, entregados: 0,
                 novedades: 0, resueltas: 0, sinMover: 0 };
      });
      const porNombre = {};
      personas.forEach(function (x) { porNombre[norm(x.nombre)] = x; });

      HOJA.Pedidos.forEach(function (f) {
        if (f.tienda !== t || mesDe(f.fecha) !== mes) return;
        const g = porNombre[norm(f.gestora_asignada)];
        if (!g) return;
        g.pedidos++;
        const est = norm(f.estado_nova || f.estado_canonico);
        if (est === 'entregado') g.entregados++;
        if (TERMINALES.indexOf(est) === -1 &&
            diasEntre(f.ultimo_movimiento || f.fecha, HOY_ISO) > 3) g.sinMover++;
      });
      HOJA.Novedades.forEach(function (n) {
        if (mesDe(n.fecha) !== mes) return;
        const g = porNombre[norm(n.gestora)];
        if (!g) return;
        g.novedades++;
        if (norm(n.estado) === 'resuelta') g.resueltas++;
      });
      personas.forEach(function (x) {
        x.efectividad = x.pedidos ? x.entregados / x.pedidos * 100 : 0;
        x.hoy = null;
      });

      const salida = rol === 'gestora'
        ? personas.filter(function (x) { return norm(x.nombre) === norm(SES.nombre); })
        : personas;
      return { ok: true, mes: mes, personas: salida, puedeEditar: rol === 'dueno' };
    },

    fuentes: function (p) {
      const t = tiendaDe(p);
      return { ok: true, tienda: t,
               configuradas: HOJA.Fuentes.filter(function (f) { return f.tienda === t; }),
               catalogo: [] };
    },

    cierre: function (p) {
      const t = tiendaDe(p);
      if (!accesible(t)) return { ok: false, error: 'No tienes acceso a esa tienda.' };
      const mes = p.mes || MES_ACTUAL;
      const rol = rolEfectivo(p);
      const cong = HOJA.Cierres.filter(function (c) {
        return c.tienda === t && c.mes === mes;
      })[0];
      const r = {
        ok: true, tienda: t, mes: mes, moneda: TIENDAS[t].moneda,
        actual: cong ? cong._datos : mesDe_(t, mes, rol),
        anterior: mesDe_(t, mesMenos(mes, 1), rol),
        cerrado: !!cong, cerrado_en: cong ? cong.cerrado_en : '',
      };
      if (!cong) { r.provisional = r.actual.pendientes > 0; r.pendientes = r.actual.pendientes; }
      r.monedaReporte = TIENDAS[t].moneda;   // el demo reporta en su propia moneda
      return r;
    },

    cerrarmes: function (p) {
      if (rolEfectivo(p) !== 'dueno') return { ok: false, error: 'Solo la dueña cierra un mes.' };
      const t = tiendaDe(p), mes = p.mes || MES_ACTUAL;
      if (!/^\d{4}-\d{2}$/.test(mes)) return { ok: false, error: 'El mes va como AAAA-MM.' };
      const ya = HOJA.Cierres.filter(function (c) { return c.tienda === t && c.mes === mes; })[0];
      if (ya) {
        return { ok: false, error: 'Ese mes ya está cerrado. Para rehacerlo, ' +
                 'borra la fila en la hoja Cierres.' };
      }
      const d = mesDe_(t, mes, 'dueno');
      if (d.pendientes > 0 && !p.forzar) {
        return { ok: false, avisar: true, pendientes: d.pendientes,
                 valorAbierto: d.valorAbierto, mes: mes, tienda: t,
                 error: 'Todavía hay ' + d.pendientes + ' pedidos sin desenlace.' };
      }
      HOJA.Cierres.push({ tienda: t, mes: mes, estado: 'cerrado',
        cerrado_en: HOY_ISO, cerrado_por: SES.email, snapshot: '', _datos: d });
      anotar('Cierres', t + ' ' + mes, 'estado', 'abierto', 'cerrado');
      return { ok: true, mes: mes, tienda: t, datos: d, cerrado_en: HOY_ISO };
    },

    historial: function (p) {
      const t = tiendaDe(p);
      if (rolEfectivo(p) !== 'dueno') {
        return { ok: false, error: 'El estado del histórico lo ve la dueña.' };
      }
      const cuantos = Math.min(12, Math.max(1, parseInt(p.meses || 8, 10)));
      const meses = [];
      for (let i = 0; i < cuantos; i++) meses.push(mesMenos(MES_ACTUAL, i));

      const out = meses.map(function (mm) {
        const d = mesDe_(t, mm, 'dueno');
        const car = carteraDelMes(t, mm);
        const b = {
          mes: mm, pedidos: d.pedidos, entregados: d.entregados, ventas: d.ventas,
          pautaFilas: HOJA.Pauta.filter(function (f) {
            return f.tienda === t && mesDe(f.fecha) === mm;
          }).length,
          pautaSinTasa: 0, monedasSinTasa: {},
          fijos: d.fijos || 0,
          nFijos: (d.detalleFijos || []).length,
          cartera: car.hay ? 1 : 0,
          cerrado: HOJA.Cierres.some(function (c) { return c.tienda === t && c.mes === mm; }),
          falta: [],
        };
        if (!b.pedidos) { b.falta.push('pedidos'); return b; }
        if (!b.pautaFilas) b.falta.push('pauta');
        if (!b.nFijos) b.falta.push('gastos fijos');
        b.listo = !b.falta.length;
        b.enCurso = mm === MES_ACTUAL;
        return b;
      });
      return { ok: true, tienda: t, moneda: TIENDAS[t].moneda, meses: out };
    },

    cas: function (p) {
      const t = tiendaDe(p);
      const minDias = num(umbrales(t).dias_sin_mover) || 3;
      const ventana = p.todos ? 3650 : 10;
      const abiertos = {};
      HOJA.CAS.forEach(function (c) {
        if (c.tienda === t && c.estado !== 'resuelto') abiertos[c.pedido_id] = c;
      });

      const candidatos = [];
      let viejos = 0, valorViejos = 0;
      HOJA.Pedidos.forEach(function (f) {
        if (f.tienda !== t) return;
        const est = norm(f.estado_nova || f.estado_canonico);
        if (TERMINALES.indexOf(est) !== -1) return;
        const quieto = diasEntre(f.ultimo_movimiento || f.fecha, HOY_ISO);
        if (quieto < minDias) return;
        const edad = diasEntre(f.fecha, HOY_ISO);
        if (edad > ventana) { viejos++; valorViejos += num(f.valor); return; }
        candidatos.push({
          pedidoId: f.id, idExterno: f.id_externo, cliente: f.cliente,
          ciudad: f.ciudad, guia: f.guia, transportadora: f.transportadora,
          estado: est, valor: num(f.valor), fecha: f.fecha,
          diasQuieto: quieto, edad: edad,
          yaTiene: !!abiertos[f.id], cas: abiertos[f.id] || null,
        });
      });
      candidatos.sort(function (a, b) { return b.diasQuieto - a.diasQuieto; });

      return { ok: true, tienda: t, minDias: minDias, ventana: ventana,
               candidatos: candidatos,
               cas: HOJA.CAS.filter(function (c) { return c.tienda === t; }),
               viejos: viejos, valorViejos: valorViejos,
               sinRadicar: candidatos.filter(function (x) { return !x.yaTiene; }).length };
    },

    cas_escribir: function (p) {
      const t = tiendaDe(p);
      if (p.id) {
        const c = HOJA.CAS.filter(function (x) { return String(x.id) === String(p.id); })[0];
        if (!c) return { ok: false, error: 'No encuentro ese CAS.' };
        Object.keys(p.campos || {}).forEach(function (k) { c[k] = p.campos[k]; });
        if (c.estado === 'resuelto' && !c.cerrado_en) c.cerrado_en = HOY_ISO;
        c.ultima_gestion = HOY_ISO;
        return { ok: true, id: c.id };
      }
      const ped = HOJA.Pedidos.filter(function (x) { return x.id === p.pedidoId; })[0];
      if (!ped) return { ok: false, error: 'No encuentro ese pedido.' };
      const c = {
        id: nid('cas'), tienda: t, pedido_id: ped.id, id_externo: ped.id_externo,
        guia: ped.guia, transportadora: ped.transportadora,
        abierto_en: HOY_ISO, abierto_por: SES.email,
        ticket: (p.campos && p.campos.ticket) || '', estado: 'abierto',
        dias_quieto: diasEntre(ped.ultimo_movimiento || ped.fecha, HOY_ISO),
        ultima_gestion: HOY_ISO, respuesta: '', cerrado_en: '',
        nota: (p.campos && p.campos.nota) || '',
      };
      HOJA.CAS.push(c);
      anotar('CAS', c.id, '(alta)', '', ped.id_externo);
      return { ok: true, id: c.id };
    },

    alarmas: function (p) {
      const t = tiendaDe(p);
      const rol = rolEfectivo(p);
      const u = umbrales(t);
      const m = mesDe_(t, MES_ACTUAL, 'dueno');
      const out = [];

      const mkAlarma = function (id, nombre, nivel, titulo, detalle, casos, ir) {
        return { id: id, nivel: nivel, nombre: nombre, titulo: titulo,
                 detalle: detalle, casos: casos || [], ir: ir || '' };
      };

      const dias = num(u.dias_sin_mover);
      if (dias > 0) {
        const det = [];
        HOJA.Pedidos.forEach(function (f) {
          if (f.tienda !== t) return;
          const est = norm(f.estado_nova || f.estado_canonico);
          if (TERMINALES.indexOf(est) !== -1) return;
          const d2 = diasEntre(f.ultimo_movimiento || f.fecha, HOY_ISO);
          if (d2 >= dias) det.push({ id: f.id_externo, cliente: f.cliente,
                                     dias: d2, gestora: f.gestora_asignada });
        });
        if (det.length) {
          det.sort(function (a, b) { return b.dias - a.dias; });
          out.push(mkAlarma('sin_mover', 'Pedidos detenidos', 'mal',
            det.length + ' pedidos llevan ' + dias + ' días o más sin moverse',
            'El más viejo lleva ' + det[0].dias + ' días. Un pedido detenido no avisa ' +
            'solo: o se gestiona, o se convierte en devolución.',
            det.slice(0, 10), 'Pedidos'));
        }
      }

      const horas = num(u.horas_novedad);
      if (horas > 0) {
        const viejas = [];
        HOJA.Novedades.forEach(function (n) {
          if (norm(n.estado) !== 'abierta') return;
          const h = diasEntre(n.fecha, HOY_ISO) * 24;
          if (h >= horas) viejas.push({ id: n.pedido_id, motivo: n.motivo,
                                        horas: h, gestora: n.gestora });
        });
        if (viejas.length) {
          viejas.sort(function (a, b) { return b.horas - a.horas; });
          out.push(mkAlarma('novedad_vieja', 'Novedades sin gestionar', 'mal',
            viejas.length + ' novedades llevan más de ' + horas + ' horas abiertas',
            'Una novedad sin contactar a la clienta en el primer día se vuelve ' +
            'devolución en la mayoría de los casos.', viejas.slice(0, 10), 'Novedades'));
        }
      }

      const efMin = num(u.efectividad_min);
      if (efMin > 0 && m.resueltos >= 10 && m.efectividad < efMin) {
        out.push(mkAlarma('efectividad', 'Efectividad baja', 'mal',
          'Efectividad en ' + m.efectividad.toFixed(1) + '%, bajo tu meta de ' + efMin + '%',
          'De ' + m.resueltos + ' pedidos ya resueltos este mes llegaron ' + m.entregados +
          '. Cada punto por debajo son ventas que ya pagaste en pauta y no entraron.',
          [], 'Pedidos'));
      }

      if (rol === 'dueno') {
        const ant = mesDe_(t, mesMenos(MES_ACTUAL, 1), 'dueno');
        const subida = num(u.cpa_subida_pct);
        if (subida > 0 && ant.cpa && m.cpa && m.entregados >= 10) {
          const pct = (m.cpa - ant.cpa) / ant.cpa * 100;
          if (pct >= subida) {
            out.push(mkAlarma('cpa_sube', 'CPA subiendo', 'ojo',
              'El CPA subió ' + pct.toFixed(0) + '% contra el mes pasado',
              'Pasó de ' + ant.cpa.toFixed(2) + ' a ' + m.cpa.toFixed(2) + '. ' +
              'Conseguir la misma venta te está costando más.', [], 'Pauta'));
          }
        }
      }

      return { ok: true, tienda: t, alarmas: out,
               umbrales: rol === 'dueno' ? u : null,
               ajustes: ajustes(t),
               catalogo: [
                 { id: 'sin_mover', nombre: 'Pedidos detenidos', param: 'dias_sin_mover', unidad: 'días' },
                 { id: 'novedad_vieja', nombre: 'Novedades sin gestionar', param: 'horas_novedad', unidad: 'horas' },
                 { id: 'efectividad', nombre: 'Efectividad baja', param: 'efectividad_min', unidad: '%' },
                 { id: 'devoluciones', nombre: 'Devoluciones altas', param: 'devoluciones_max', unidad: '%', opcional: true },
                 { id: 'cpa', nombre: 'CPA cerca del techo', param: 'cpa_aviso_pct', unidad: '% del techo' },
                 { id: 'cpa_sube', nombre: 'CPA subiendo', param: 'cpa_subida_pct', unidad: '% vs. mes pasado', opcional: true },
                 { id: 'stock', nombre: 'Stock por agotarse', param: 'stock_dias_min', unidad: 'días de cobertura', opcional: true },
               ] };
    },

    parametros: function (p) {
      const t = tiendaDe(p);
      if (!p.cambios) return { ok: true, tienda: t, umbrales: umbrales(t), ajustes: ajustes(t) };
      if (rolEfectivo(p) !== 'dueno') {
        return { ok: false, error: 'Solo la dueña cambia la configuración de la tienda.' };
      }
      // Mismo control que el servidor: un enlace que no es enlace se rechaza
      if (p.cambios.canal_url) {
        const v = String(p.cambios.canal_url).trim();
        if (!/^(https?:\/\/)?[\w.-]+\.\w{2,}(\/|$)/.test(v)) {
          return { ok: false, error: 'Ese enlace no se entiende. Pega la dirección ' +
                   'completa del grupo, por ejemplo chat.whatsapp.com/AbCd.' };
        }
        p.cambios.canal_url = /^https?:\/\//.test(v) ? v : 'https://' + v;
      }
      PARAMS[t] = PARAMS[t] || {};
      Object.keys(p.cambios).forEach(function (k) {
        anotar('Parametros', t, k, PARAMS[t][k], p.cambios[k]);
        PARAMS[t][k] = p.cambios[k];
      });
      return { ok: true, umbrales: umbrales(t), ajustes: ajustes(t) };
    },

    recuento: function (p) {
      const t = tiendaDe(p);
      const rol = rolEfectivo(p);
      const conDatos = [];
      let m = mesMenos(MES_ACTUAL, 1);
      for (let i = 0; i < 6 && conDatos.length < 3; i++) {
        const d = mesDe_(t, m, rol);
        if (d.pedidos > 0) conDatos.push({ mes: m, d: d });
        m = mesMenos(m, 1);
      }
      if (!conDatos.length) return { ok: true, tienda: t, vacio: true, mesActual: MES_ACTUAL };
      conDatos.reverse();
      const periodo = conDatos.length >= 3 ? 'trimestre' : 'mes';
      const usados = periodo === 'trimestre' ? conDatos : [conDatos[conDatos.length - 1]];

      const claves = ['pedidos','entregados','devueltos','resueltos','ventas',
                      'gasto','fijos','costoProducto','costoEnvio'];
      const tot = {};
      claves.forEach(function (k) { tot[k] = 0; });
      usados.forEach(function (x) { claves.forEach(function (k) { tot[k] += x.d[k] || 0; }); });
      tot.efectividad = tot.resueltos ? tot.entregados / tot.resueltos * 100 : 0;
      tot.roas = tot.gasto ? tot.ventas / tot.gasto : 0;

      const ant = { ventas: 0, entregados: 0, resueltos: 0, gasto: 0, fijos: 0 };
      let pm = mesMenos(usados[0].mes, 1);
      for (let i = 0; i < usados.length; i++) {
        const d = mesDe_(t, pm, rol);
        ant.ventas += d.ventas || 0; ant.entregados += d.entregados || 0;
        ant.resueltos += d.resueltos || 0; ant.gasto += d.gasto || 0; ant.fijos += d.fijos || 0;
        pm = mesMenos(pm, 1);
      }
      ant.efectividad = ant.resueltos ? ant.entregados / ant.resueltos * 100 : 0;

      const avances = [];
      const pri = usados[0].d, ult = usados[usados.length - 1].d;
      if (usados.length > 1) {
        const dEf = (ult.efectividad || 0) - (pri.efectividad || 0);
        if (Math.abs(dEf) >= 3) {
          avances.push({ signo: dEf > 0 ? '+' : '', valor: dEf.toFixed(0) + ' pts',
            titulo: 'La tasa de entrega ' + (dEf > 0 ? 'mejoró' : 'cayó') + ' de ' +
              (pri.efectividad || 0).toFixed(0) + '% a ' + (ult.efectividad || 0).toFixed(0) + '%',
            detalle: 'Entre ' + usados[0].mes + ' y ' + usados[usados.length - 1].mes + '.',
            bueno: dEf > 0 });
        }
        const dV = (ult.ventas || 0) - (pri.ventas || 0);
        if (pri.ventas && Math.abs(dV / pri.ventas) >= 0.1) {
          avances.push({ signo: dV > 0 ? '+' : '', valor: Math.round(dV / pri.ventas * 100) + '%',
            titulo: 'Las ventas ' + (dV > 0 ? 'subieron' : 'bajaron') + ' de ' +
              Math.round(pri.ventas) + ' a ' + Math.round(ult.ventas),
            detalle: 'Mes a mes dentro del periodo.', bueno: dV > 0 });
        }
      }
      if (ant.ventas) {
        const dv = (tot.ventas - ant.ventas) / ant.ventas * 100;
        if (Math.abs(dv) >= 5) {
          avances.push({ signo: dv > 0 ? '+' : '', valor: dv.toFixed(0) + '%',
            titulo: 'Ingresos ' + (dv > 0 ? 'por encima' : 'por debajo') + ' del ' +
              (periodo === 'mes' ? 'mes' : 'trimestre') + ' anterior',
            detalle: Math.round(tot.ventas) + ' contra ' + Math.round(ant.ventas) + '.',
            bueno: dv > 0 });
        }
      }

      return { ok: true, tienda: t, periodo: periodo,
               meses: usados.map(function (x) { return x.mes; }),
               total: tot, anterior: ant, avances: avances.slice(0, 3),
               moneda: TIENDAS[t].moneda, mesActual: MES_ACTUAL,
               mesesConDatos: conDatos.length };
    },

    auditoria: function (p) {
      const lim = Math.min(300, parseInt(p.limite || 150, 10));
      return { ok: true, movimientos: HOJA.Auditoria.slice(0, lim) };
    },

    trozo: function (p) {
      // El archivo llega partido igual que en producción, pero aquí los
      // pedazos se guardan en memoria en vez de en el caché de Google.
      TROZOS[p.clave] = TROZOS[p.clave] || [];
      TROZOS[p.clave][p.indice] = p.trozo;
      return { ok: true, recibido: p.indice };
    },

    importar: function (p) { return importarDemo(p); },
  };

  /**
   * Subir un archivo, sin el techo de 8 MB.
   *
   * Ese techo existe por Apps Script: el archivo viaja en base64, partido
   * en pedazos de 30 KB, y del otro lado hay un script con seis minutos
   * para terminar. Ocho megas ya son casi cuatrocientas peticiones.
   *
   * En el demo no hay otro lado. El archivo se lee aquí mismo, en la
   * pestaña. Aplicarle el límite del servidor a un camino que no pasa por
   * el servidor era heredar una restricción sin heredar su motivo — y
   * dejaba fuera justo lo que un demo tiene que poder mostrar: el
   * histórico de un cliente con tres o cuatro años encima.
   *
   * Tampoco pasa por base64: el archivo va como bytes, directo. Convertir
   * cincuenta megas a texto y devolverlos a bytes es duplicar la memoria
   * y el tiempo para llegar al mismo sitio.
   */
  /**
   * Qué tienda está mirando la pantalla.
   *
   * `ST` está declarada con `let` dentro del script de la página: vive en
   * el ámbito léxico global y no cuelga de `window`. Leerla como
   * `window.ST` devuelve undefined en silencio, y el archivo terminaría
   * importado en la tienda equivocada sin que nadie lo note.
   */
  function tiendaActiva() {
    try {
      const t = (0, eval)('ST');
      if (t && TIENDAS[t]) return t;
    } catch (e) { /* la pantalla todavía no arrancó */ }
    return SES ? SES.tiendas[0] : 'ec';
  }

  window.subirArchivo = async function (input, z) {
    const file = input.files && input.files[0];
    input.value = '';
    if (!file) return;

    const fuente = document.getElementById('imp-fuente-' + z).value;
    const btn = document.getElementById('imp-btn-' + z);
    const mb = (file.size / 1048576).toFixed(1);

    window.impEstado(z, 'Leyendo ' + file.name + ' (' + mb + ' MB)…', 'ojo');
    btn.style.opacity = '.6'; btn.textContent = 'Leyendo…';

    try {
      const bytes = new Uint8Array(await file.arrayBuffer());

      window.impEstado(z, 'Procesando ' + mb + ' MB…\n' +
        (bytes[0] === 0x50 && bytes[1] === 0x4B
          ? 'Un Excel grande puede tardar y dejar la pantalla quieta un momento.'
          : 'Leyendo las filas.'), 'ojo');

      // Un respiro para que el navegador alcance a pintar el aviso antes
      // de bloquearse con el archivo. Sin esto el mensaje nunca se ve.
      await new Promise(function (ok) { setTimeout(ok, 60); });

      const t0 = performance.now();
      const r = await ACCIONES.importar({ fuente: fuente, tienda: tiendaActiva(),
                                          nombre: file.name, bytes: bytes });
      const seg = ((performance.now() - t0) / 1000).toFixed(1);
      invalidarCache();

      if (!r.ok) { window.impEstado(z, r.error, 'mal'); return; }

      window.impEstado(z, r.resumen + '\nTardó ' + seg + ' s.\n\nActualizando…', 'ok');
      await window.cargarReales();
      await window.cargarHistorialFuentes();
      window.impEstado(z,
        r.nuevas + ' filas entraron al demo · ' +
        (r.repetidas ? r.repetidas + ' repetidas se omitieron · ' : '') +
        (r.ignoradas ? r.ignoradas + ' sin fecha se saltaron · ' : '') +
        'en ' + seg + ' s.\n\n' +
        'Se ven reflejadas al instante y desaparecen cuando recargues.', 'ok');
      window.showToast('Pantalla actualizada');
    } catch (e) {
      console.error('[demo] importar', e);
      window.impEstado(z, 'No se pudo leer el archivo: ' + e.message, 'mal');
    } finally {
      btn.style.opacity = '1'; btn.textContent = 'Elegir archivo';
    }
  };

  const TROZOS = {};

  // ─── 9 · SUBIR UN ARCHIVO DE VERDAD ───────────────────────────
  /**
   * El demo lee el archivo que le den, no finge leerlo.
   *
   * Un demo que contesta "importados 240 pedidos" sin mirar el archivo
   * es una mentira con forma de demostración: la primera pregunta que
   * hace cualquiera es "¿y si subo el mío?", y ahí se cae todo. Así que
   * aquí se abre el archivo, se leen sus columnas y lo que entra cambia
   * de verdad las cifras de la pantalla.
   *
   * CSV se lee solo. Para Excel se carga un lector desde cdnjs, y si no
   * hay internet se dice con claridad en vez de fallar en silencio.
   */
  function b64aBytes(b64) {
    const bin = atob(b64);
    const out = new Uint8Array(bin.length);
    for (let i = 0; i < bin.length; i++) out[i] = bin.charCodeAt(i);
    return out;
  }

  /**
   * El lector de Excel viaja con Nova, no se descarga.
   *
   * Antes venía de un CDN. Funcionaba en la oficina y fallaba justo donde
   * importa: la conexión lenta de quien abre el demo por primera vez, la
   * red de empresa que bloquea dominios ajenos, el avión. Un demo que
   * depende de que un tercero esté disponible no es un demo confiable —y
   * el archivo suelto, que se abre sin internet, no podía leer Excel
   * en absoluto.
   *
   * Pesa 860 KB. Se carga solo cuando alguien sube un .xlsx, así que
   * quien no suba ninguno no lo descarga nunca.
   *
   * Si el archivo local faltara, queda el CDN como último recurso. No al
   * revés: lo propio primero.
   */
  const LECTOR_LOCAL = 'vendor/xlsx.min.js';
  const LECTOR_CDN = 'https://cdnjs.cloudflare.com/ajax/libs/xlsx/0.18.5/xlsx.full.min.js';

  function cargarGuion(src) {
    return new Promise(function (ok, mal) {
      const s = document.createElement('script');
      s.src = src;
      s.onload = function () { ok(window.XLSX); };
      s.onerror = function () { mal(new Error('no se pudo cargar ' + src)); };
      document.head.appendChild(s);
    });
  }

  async function cargarLectorExcel() {
    if (window.XLSX) return window.XLSX;

    /**
     * En el archivo de un solo pedazo el lector viene adentro, guardado
     * en una etiqueta que el navegador no ejecuta. Se despierta aquí, la
     * primera vez que alguien sube un Excel: así el archivo abre rápido
     * para quien solo viene a mirar.
     */
    const dentro = document.getElementById('nova-xlsx');
    if (dentro && dentro.textContent.length > 1000) {
      (0, eval)(dentro.textContent);
      if (window.XLSX) return window.XLSX;
    }

    try { return await cargarGuion(LECTOR_LOCAL); }
    catch (e) { /* se intenta el de afuera */ }
    try { return await cargarGuion(LECTOR_CDN); }
    catch (e) {
      throw new Error('El demo no pudo cargar el lector de Excel.\n\n' +
        'Guarda el archivo como CSV y súbelo así: ese lo lee sin depender de nada.');
    }
  }

  function leerCSV(texto) {
    const filas = [];
    let campo = '', fila = [], enCom = false;
    for (let i = 0; i < texto.length; i++) {
      const ch = texto[i];
      if (enCom) {
        if (ch === '"' && texto[i + 1] === '"') { campo += '"'; i++; }
        else if (ch === '"') enCom = false;
        else campo += ch;
      } else if (ch === '"') enCom = true;
      else if (ch === ',' || ch === ';') { fila.push(campo); campo = ''; }
      else if (ch === '\n') { fila.push(campo); filas.push(fila); fila = []; campo = ''; }
      else if (ch !== '\r') campo += ch;
    }
    if (campo || fila.length) { fila.push(campo); filas.push(fila); }
    return filas.filter(function (f) {
      return f.some(function (c) { return String(c).trim() !== ''; });
    });
  }

  /** Los mismos alias que usa el importador de verdad, recortados a lo útil. */
  const ALIAS = {
    dropi: {
      id_externo: ['id', 'no. orden', 'numero orden', 'orden'],
      fecha: ['fecha', 'fecha creacion', 'created at'],
      cliente: ['nombre cliente', 'cliente', 'nombre'],
      telefono: ['telefono', 'celular', 'phone'],
      guia: ['numero guia', 'no guia', 'guia', 'tracking'],
      estado: ['estatus', 'estado', 'status'],
      ciudad: ['ciudad destino', 'ciudad', 'city'],
      transportadora: ['transportadora', 'courier', 'operador'],
      valor: ['total de la orden', 'valor', 'total', 'monto'],
      producto: ['producto', 'productos', 'nombre producto'],
      cantidad: ['cantidad', 'qty', 'unidades'],
      costo_envio: ['flete', 'costo envio', 'envio', 'precio flete'],
      costo_producto: ['precio proveedor', 'costo proveedor'],
      departamento: ['departamento destino', 'departamento', 'provincia'],
      sku: ['sku', 'codigo'],
      ultimo_movimiento: ['ultimo movimiento'],
    },
    meta: {
      fecha: ['dia', 'fecha', 'inicio del informe', 'date'],
      fecha_fin: ['fin del informe'],
      campana: ['nombre de la campana', 'campana'],
      conjunto: ['nombre del conjunto de anuncios', 'conjunto de anuncios'],
      gasto: ['importe gastado (usd)', 'importe gastado', 'gasto', 'amount spent'],
      resultados: ['resultados', 'results', 'compras'],
      impresiones: ['impresiones'], clics: ['clics', 'clics en el enlace'],
    },
  };

  /**
   * Los mismos estados que reconoce Nova de verdad.
   *
   * VERIFICADO contra un archivo real de Dropi Colombia: 871 órdenes,
   * abril a septiembre, 18 estados distintos. Un estado que no se
   * reconoce cae en "pendiente" y eso no da error — solo deja pedidos
   * resueltos contados como si siguieran en camino, el mes sin cerrar y
   * la tasa de entrega más baja de lo que fue.
   */
  const ESTADO_DEMO = {
    'entregado': 'entregado', 'entregada': 'entregado',

    'devolucion': 'devolucion', 'devuelto': 'devolucion', 'devuelta': 'devolucion',
    'en proceso de devolucion': 'devolucion',
    // Rechazado en la puerta: el paquete vuelve y el flete se paga igual
    'rechazado': 'devolucion', 'rechazada': 'devolucion',

    'cancelado': 'cancelado', 'cancelada': 'cancelado',
    'guia_anulada': 'cancelado', 'guia anulada': 'cancelado',

    'novedad': 'novedad', 'novedad solucionada': 'novedad_resuelta',
    'solucion aprobada': 'novedad_resuelta',

    'pendiente': 'pendiente', 'pendiente confirmacion': 'pendiente',

    'guia generada': 'confirmado', 'guia_generada': 'confirmado',
    'en procesamiento': 'confirmado',

    'en bodega origen': 'en_bodega', 'preparado para transportadora': 'en_bodega',
    'en bodega transportadora': 'en_bodega',
    'ingresando operativo a bodega': 'en_bodega',

    'en reparto': 'en_transito', 'en transito': 'en_transito',
    'despachada': 'en_transito', 'despachado': 'en_transito',
    'en bodega destino': 'en_transito', 'en terminal destino': 'en_transito',
    'en reexpedicion': 'en_transito', 'en espera de ruta domestica': 'en_transito',
    'zona de entrega': 'en_transito', 'en distribucion a cliente': 'en_transito',

    'reclame en oficina': 'en_oficina', 'reclamo en oficina': 'en_oficina',
    'para retiro en agencia': 'en_oficina',
    'para retiro en agencia servientrega': 'en_oficina',
  };

  /**
   * Un número escrito como se escribe en América Latina.
   *
   * "27044,5" son veintisiete mil con cincuenta centavos, y "1.234,56"
   * son mil doscientos treinta y cuatro. Cambiar la coma por punto a
   * secas funciona con el primero y con el segundo produce "1.234.56",
   * que no es ningún número — y el costo de ese pedido se vuelve cero
   * sin que nadie lo note.
   *
   * La regla: el último separador que aparezca es el decimal, siempre
   * que le sigan una o dos cifras. Lo demás son separadores de miles.
   */
  function numLatam(v) {
    if (typeof v === 'number') return isFinite(v) ? v : 0;
    let s = String(v == null ? '' : v).trim().replace(/[^\d.,-]/g, '');
    if (!s) return 0;

    const coma = s.lastIndexOf(','), punto = s.lastIndexOf('.');
    const corte = Math.max(coma, punto);

    if (corte !== -1 && s.length - corte - 1 <= 2 && s.length - corte - 1 > 0) {
      s = s.slice(0, corte).replace(/[.,]/g, '') + '.' + s.slice(corte + 1);
    } else {
      s = s.replace(/[.,]/g, '');          // todos eran de miles
    }

    const n = parseFloat(s);
    return isFinite(n) ? n : 0;
  }

  /** Día primero, que es como escribe toda América Latina. */
  function fechaDemo(v) {
    const s = String(v == null ? '' : v).trim();
    if (!s) return '';
    let m = s.match(/^(\d{4})-(\d{2})-(\d{2})/);
    if (m) return m[1] + '-' + m[2] + '-' + m[3];
    m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})/);
    if (m) {
      const a = m[3].length === 2 ? '20' + m[3] : m[3];
      return a + '-' + dd(+m[2]) + '-' + dd(+m[1]);
    }
    return '';
  }

  /**
   * Cuál de las pestañas trae los pedidos.
   *
   * Antes se leía siempre la primera. Con un export de Dropi da igual —
   * solo tiene una— pero el control que lleva un equipo de verdad no es
   * un export: es un cuaderno con diecisiete pestañas. En el primer
   * archivo real que probamos, la primera pestaña se llamaba ACCESOS y
   * era la lista de usuarios; los pedidos estaban en la trece.
   *
   * Así que se miran todas y gana la que más columnas reconocidas tenga,
   * con el número de filas como desempate. Y se dice cuál se eligió: si
   * acierta, tranquiliza; si se equivoca, se ve dónde.
   */
  function elegirHoja(wb, fuente) {
    const alias = ALIAS[fuente] || ALIAS[fuente === 'meta_facturacion' ? 'meta' : 'dropi'];
    const claves = Object.keys(alias);
    let mejor = null;

    wb.SheetNames.forEach(function (nombre) {
      const m = XLSX.utils.sheet_to_json(wb.Sheets[nombre],
        { header: 1, raw: false, defval: '' });
      if (!m || m.length < 2) return;

      const enc = (m[0] || []).map(norm);
      let puntos = 0;
      claves.forEach(function (campo) {
        if (alias[campo].some(function (a) { return enc.indexOf(a) !== -1; })) puntos++;
      });

      // Sin fecha no hay nada que importar, por muchas columnas que tenga
      const hayFecha = alias.fecha.some(function (a) { return enc.indexOf(a) !== -1; });
      if (!hayFecha || puntos < 3) return;

      const filas = m.filter(function (f) {
        return f.some(function (c) { return String(c).trim() !== ''; });
      }).length;

      if (!mejor || puntos > mejor.puntos ||
          (puntos === mejor.puntos && filas > mejor.filas)) {
        mejor = { nombre: nombre, matriz: m, puntos: puntos, filas: filas };
      }
    });

    return mejor;
  }

  async function importarDemo(p) {
    const t = tiendaDe(p);
    const fuente = String(p.fuente || 'dropi');
    let hojaUsada = '';

    let bytes = p.bytes;
    if (!bytes) {
      let b64 = p.contenido;
      if (!b64 && p.clave) b64 = (TROZOS[p.clave] || []).join('');
      if (!b64) return { ok: false, error: 'No llegó el archivo.' };
      bytes = b64aBytes(b64);
    }
    let matriz;

    // Los .xlsx empiezan por PK: es un zip
    const esZip = bytes[0] === 0x50 && bytes[1] === 0x4B;
    if (esZip) {
      let XLSX;
      try { XLSX = await cargarLectorExcel(); }
      catch (e) { return { ok: false, error: e.message }; }
      const wb = XLSX.read(bytes, { type: 'array', cellDates: false });
      const elegida = elegirHoja(wb, fuente);
      if (!elegida) {
        return { ok: false, error: 'Ese Excel tiene ' + wb.SheetNames.length +
          ' pestañas y en ninguna encontré columnas de pedidos.\n\n' +
          'Pestañas: ' + wb.SheetNames.join(', ') };
      }
      hojaUsada = elegida.nombre;
      matriz = elegida.matriz;
    } else {
      matriz = leerCSV(new TextDecoder('utf-8').decode(bytes));
    }

    if (!matriz || matriz.length < 2) {
      return { ok: false, error: 'El archivo no tiene filas debajo del encabezado.' };
    }

    const enc = matriz[0].map(norm);
    const alias = ALIAS[fuente] || ALIAS[fuente === 'meta_facturacion' ? 'meta' : 'dropi'];
    const col = {};
    Object.keys(alias).forEach(function (campo) {
      for (let i = 0; i < alias[campo].length; i++) {
        const j = enc.indexOf(alias[campo][i]);
        if (j !== -1) { col[campo] = j; return; }
      }
    });

    if (col.fecha === undefined) {
      return { ok: false, error: 'No encuentro una columna de fecha en ese archivo.\n\n' +
        'Las columnas que vi: ' + enc.filter(String).slice(0, 12).join(', ') + '…' };
    }

    const esPauta = fuente === 'meta' || fuente === 'meta_facturacion' || fuente === 'tiktok';
    let nuevas = 0, repetidas = 0, sinFecha = 0;

    /**
     * Los números de orden que ya están, en un conjunto.
     *
     * Antes esto se resolvía recorriendo TODOS los pedidos por cada fila
     * del archivo. Con un export de cuatrocientas filas ni se notaba; con
     * el histórico de cuatro años de una tienda —cincuenta mil pedidos
     * contra cincuenta mil ya cargados— son dos mil quinientos millones
     * de comparaciones y la pestaña se cuelga sin decir nada.
     *
     * Un conjunto responde en tiempo fijo. El archivo entra en una sola
     * pasada, y además se va llenando: así un archivo que trae la misma
     * orden dos veces tampoco la duplica.
     */
    const yaEstan = new Set();
    if (!esPauta) {
      for (let i = 0; i < HOJA.Pedidos.length; i++) {
        const x = HOJA.Pedidos[i];
        if (x.tienda === t && x.id_externo) yaEstan.add(String(x.id_externo));
      }
    }

    for (let i = 1; i < matriz.length; i++) {
      const f = matriz[i];
      const fecha = fechaDemo(f[col.fecha]);
      if (!fecha) { sinFecha++; continue; }

      if (esPauta) {
        HOJA.Pauta.push({
          id: nid('a'), fecha: fecha,
          fecha_fin: col.fecha_fin !== undefined ? fechaDemo(f[col.fecha_fin]) : '',
          tienda: t, plataforma: fuente === 'tiktok' ? 'tiktok' : 'meta', cuenta: '',
          campana: col.campana !== undefined ? f[col.campana] : 'Importada',
          conjunto: col.conjunto !== undefined ? f[col.conjunto] : '',
          entrega: '', presupuesto: '',
          gasto: numLatam(f[col.gasto]),
          moneda_gasto: TIENDAS[t].moneda,
          gasto_normalizado: '', impresiones: col.impresiones !== undefined ? f[col.impresiones] : '',
          alcance: '', frecuencia: '', clics: col.clics !== undefined ? f[col.clics] : '',
          ctr: '', cpc: '', cpm: '',
          resultados: col.resultados !== undefined ? numLatam(f[col.resultados]) : 0,
          compras: '', cpa: '', roas: '', valor_conv: '', visitas_lp: '',
        });
        nuevas++;
        continue;
      }

      const ext = col.id_externo !== undefined ? String(f[col.id_externo]).trim() : '';
      if (ext && yaEstan.has(ext)) { repetidas++; continue; }
      if (ext) yaEstan.add(ext);

      const crudo = col.estado !== undefined ? norm(f[col.estado]) : '';
      const est = ESTADO_DEMO[crudo] || 'pendiente';
      HOJA.Pedidos.push({
        id: nid('p'), fuente: fuente, id_externo: ext || String(500000 + SEQ),
        fecha: fecha, tienda: t,
        cliente: col.cliente !== undefined ? f[col.cliente] : '',
        cedula: '', correo: '',
        telefono: col.telefono !== undefined ? f[col.telefono] : '',
        telefono_norm: '', telefono_2: '', telefono_2_norm: '',
        ciudad: col.ciudad !== undefined ? f[col.ciudad] : '',
        departamento: col.departamento !== undefined ? f[col.departamento] : '',
        direccion: '',
        producto: col.producto !== undefined ? f[col.producto] : 'Sin producto',
        sku: col.sku !== undefined ? f[col.sku] : '',
        cantidad: col.cantidad !== undefined ? numLatam(f[col.cantidad]) || 1 : 1,
        valor: col.valor !== undefined ? numLatam(f[col.valor]) : 0,
        costo_producto: col.costo_producto !== undefined ? numLatam(f[col.costo_producto]) : 0,
        costo_envio: col.costo_envio !== undefined ? numLatam(f[col.costo_envio]) : 0,
        metodo_pago: 'contraentrega', bodega: '',
        estado: crudo, estado_transportadora: '', estado_canonico: est,
        transportadora: col.transportadora !== undefined ? f[col.transportadora] : '',
        guia: col.guia !== undefined ? f[col.guia] : '',
        intentos: '', gestora_asignada: GESTORAS[t][0],
        fecha_promesa: '', fecha_entrega: est === 'entregado' ? fecha : '',
        razon_cancelacion: '', estado_nova: '', nota: '',
        ultimo_movimiento: col.ultimo_movimiento !== undefined
          ? (fechaDemo(f[col.ultimo_movimiento]) || fecha) : fecha,
        adelanto: '', acuerdo_oficina: '', confirmado_oficina: '',
        actualizado_en: HOY_ISO, actualizado_por: SES.email,
      });
      nuevas++;
    }

    const ya = HOJA.Fuentes.filter(function (x) {
      return x.tienda === t && x.fuente === fuente;
    })[0];
    const sello = new Date().toISOString().slice(0, 16).replace('T', ' ');
    if (ya) { ya.ultima = sello; ya.filas = (ya.filas || 0) + nuevas; }
    else HOJA.Fuentes.push({ tienda: t, fuente: fuente, tipo: esPauta ? 'pauta' : 'pedidos',
                             ultima: sello, filas: nuevas });

    anotar('Importación', fuente, 'filas', '', nuevas);

    return {
      ok: true, fuente: fuente, tienda: t,
      nuevas: nuevas, actualizadas: 0, repetidas: repetidas,
      ignoradas: sinFecha, total: matriz.length - 1,
      hoja: hojaUsada,
      resumen: nuevas + ' filas entraron al demo' +
               (hojaUsada ? ' (pestaña "' + hojaUsada + '")' : '') + '.',
      demo: true,
    };
  }

  // ─── 10 · LA PUERTA DEL DEMO ──────────────────────────────────
  const ESTILO = `
  #demo-gate{position:fixed;inset:0;z-index:9000;display:flex;align-items:center;
    justify-content:center;padding:24px 16px;overflow-y:auto;
    background:radial-gradient(ellipse at 50% 18%,#1d3350 0%,#0a121d 72%);}
  #demo-gate .dg-box{width:100%;max-width:440px;display:flex;flex-direction:column;gap:22px;}
  #demo-gate h1{font:600 26px/1.18 'Fraunces',Georgia,serif;color:#f2ede3;margin:0;
    text-wrap:balance;}
  #demo-gate .dg-lead{font:400 14px/1.6 'DM Sans',system-ui,sans-serif;color:#9fb0c4;margin:0;}
  #demo-gate .dg-tag{font:600 10px/1 'DM Sans',sans-serif;letter-spacing:.18em;
    text-transform:uppercase;color:#d8b65a;}
  #demo-gate .dg-roles{display:flex;flex-direction:column;gap:10px;}
  #demo-gate .dg-rol{display:flex;gap:13px;align-items:flex-start;text-align:left;
    background:rgba(255,255,255,.045);border:1px solid rgba(216,182,90,.22);
    border-radius:14px;padding:15px 16px;cursor:pointer;color:inherit;
    transition:border-color .15s,background .15s;font-family:inherit;}
  #demo-gate .dg-rol:hover{border-color:#d8b65a;background:rgba(216,182,90,.08);}
  #demo-gate .dg-ico{font-size:19px;line-height:1.35;flex-shrink:0;color:#d8b65a;}
  #demo-gate .dg-txt{display:flex;flex-direction:column;gap:4px;min-width:0;}
  #demo-gate .dg-nm{display:block;font:600 14.5px/1.25 'DM Sans',sans-serif;color:#f2ede3;}
  #demo-gate .dg-ds{display:block;font:400 12.5px/1.5 'DM Sans',sans-serif;color:#93a4b8;}
  #demo-gate .dg-logo{max-width:168px;height:auto;display:block;margin-bottom:20px;}
  #demo-gate .dg-nota{font:400 11.5px/1.6 'DM Sans',sans-serif;color:#7c8ca0;
    border-top:1px solid rgba(255,255,255,.09);padding-top:14px;margin:0;}
  #demo-cinta{position:fixed;left:0;right:0;bottom:0;z-index:8000;
    background:#d8b65a;color:#221d10;padding:7px 14px;
    font:600 11.5px/1.45 'DM Sans',system-ui,sans-serif;text-align:center;}
  #demo-cinta button{background:rgba(0,0,0,.14);border:none;border-radius:20px;
    color:inherit;font:600 11px 'DM Sans',sans-serif;padding:4px 11px;margin-left:9px;
    cursor:pointer;font-family:inherit;}
  /* Que la cinta no se coma la última fila de nada */
  body.demo-on .sidebar{padding-bottom:50px;}
  body.demo-on .content{padding-bottom:54px;}
  body.demo-on .chat-fab{bottom:52px;}
  body.demo-on .chat-panel{bottom:30px;}
  @media(max-width:520px){ #demo-gate h1{font-size:22px;} }
  `;

  const ROLES_DEMO = [
    ['dueno', '◆', 'La dueña',
     'Lo ve todo: dinero, margen, utilidad del mes, pauta y cierre. Es la pantalla ' +
     'de quien decide.'],
    ['gestora', '✦', 'La gestora',
     'Pedidos, novedades y su propio rendimiento. Sin margen ni gasto: no es su ' +
     'trabajo y no debería distraerla.'],
    ['admin', '◈', 'La administradora',
     'Opera todo el día y ve el equipo entero, pero la plata sigue siendo de la dueña.'],
  ];

  function puerta() {
    const st = document.createElement('style');
    st.textContent = ESTILO;
    document.head.appendChild(st);

    const g = document.createElement('div');
    g.id = 'demo-gate';
    g.innerHTML =
      '<div class="dg-box">' +
        '<div>' +
          // El de tinta blanca: el fondo de esta puerta es oscuro siempre
          '<img class="dg-logo" src="logo-nova-empresarial-claro.png" alt="Nova Empresarial">' +
          '<div class="dg-tag">Demostración</div>' +
          '<h1 style="margin-top:9px">Nova, con una tienda que no existe</h1>' +
          '<p class="dg-lead" style="margin-top:10px">Todo funciona: los números, ' +
          'los filtros, editar un pedido, subir un archivo, cerrar un mes. Nada se ' +
          'guarda. Al recargar la página vuelve a empezar de cero.</p>' +
        '</div>' +
        '<div>' +
          '<p class="dg-lead" style="margin-bottom:11px"><strong style="color:#d8b65a">' +
          'Entra como quieras y cambia cuando quieras.</strong> Cada rol ve cosas ' +
          'distintas, y esa es media herramienta.</p>' +
          '<div class="dg-roles">' +
            ROLES_DEMO.map(function (r) {
              return '<button class="dg-rol" data-rol="' + r[0] + '">' +
                '<span class="dg-ico">' + r[1] + '</span>' +
                '<span class="dg-txt"><span class="dg-nm">' + r[2] + '</span>' +
                '<span class="dg-ds">' + r[3] + '</span></span></button>';
            }).join('') +
          '</div>' +
        '</div>' +
        '<p class="dg-nota">La tienda se llama <strong>Luma</strong> y es inventada, ' +
        'igual que sus clientes y sus cifras. Se parecen a las de una tienda real ' +
        'porque están calculadas con las mismas reglas, no porque sean de nadie.</p>' +
      '</div>';
    document.body.appendChild(g);

    g.querySelectorAll('.dg-rol').forEach(function (b) {
      b.addEventListener('click', function () { entrar(b.getAttribute('data-rol')); });
    });
  }

  /**
   * "Datos reales de tu hoja" no puede decirse en el demo.
   *
   * Esa etiqueta existe para distinguir lo que salió de la hoja de lo que
   * es guion de ejemplo, y es de las cosas más útiles de Nova. Aquí los
   * números SÍ vienen del cálculo de verdad —no son el guion— pero la
   * hoja no existe. Dejar la frase tal cual sería usar la etiqueta que
   * enseña a confiar justo donde no hay nada en qué confiar.
   */
  function reetiquetarOrigen() {
    const b = document.getElementById('origen-datos');
    if (!b || b.dataset.demo) return;
    if (b.textContent.indexOf('EJEMPLO') !== -1) return;
    b.style.background = '#f3e9cf';
    b.style.color = '#6f5410';
    b.textContent = '● Tienda de demostración · los números salen del mismo ' +
      'cálculo que Nova usa de verdad, sobre una tienda inventada';
  }

  /**
   * Dejar la tienda vacía para subir los archivos propios.
   *
   * La tienda inventada sirve para entender la herramienta. Pero quien
   * quiere saber si Nova le sirve a SU negocio necesita ver SUS números,
   * y mezclados con los de Luma no se puede: un agosto de 666 pedidos
   * donde 281 son de mentira no dice nada de nadie.
   *
   * Se borran los pedidos, las novedades, la pauta y la cartera. Se
   * quedan el equipo, el inventario y los gastos fijos: son la estructura
   * de la pantalla, y sin ellos la app se ve rota en vez de vacía. Los
   * gastos fijos además se pueden editar, que es justo lo que alguien
   * probando querría cambiar por los suyos.
   */
  function vaciar() {
    if (!confirm('Se borran los pedidos, las novedades y la pauta de la tienda ' +
                 'de ejemplo, para que subas los tuyos y veas solo tus números.\n\n' +
                 'El equipo y los gastos fijos se quedan para que la pantalla ' +
                 'siga teniendo forma. Todo vuelve al recargar la página.')) return;

    ['Pedidos', 'Novedades', 'Pauta', 'Cartera', 'CAS', 'Cierres', 'Fuentes']
      .forEach(function (h) { HOJA[h].length = 0; });
    HOJA.Auditoria.length = 0;
    invalidarCache();

    const c = document.getElementById('demo-cinta');
    if (c) c.firstChild.textContent =
      'DEMOSTRACIÓN · vacía y lista para tus archivos · nada se guarda';

    if (window.cargarReales) window.cargarReales();
    if (window.showToast) window.showToast('Listo: sube tu archivo en Pedidos');
  }

  function cinta() {
    if (document.getElementById('demo-cinta')) return;
    const c = document.createElement('div');
    c.id = 'demo-cinta';
    c.innerHTML = 'MODO DEMOSTRACIÓN · datos inventados, nada se guarda' +
      '<button id="demo-vaciar">Vaciar y usar mis archivos</button>' +
      '<button id="demo-reiniciar">Empezar de nuevo</button>';
    document.body.appendChild(c);
    document.body.classList.add('demo-on');
    document.getElementById('demo-reiniciar').onclick = function () { location.reload(); };
    document.getElementById('demo-vaciar').onclick = vaciar;
  }

  /**
   * Entrar por la misma puerta que usa la aplicación de verdad.
   *
   * La tentación era asignar la sesión a mano —CONECTADO, SESION, ST— y
   * llamar a pintar. No sirve, y además es peor. No sirve porque esas
   * variables están declaradas con `let` dentro del script de la página:
   * viven en el ámbito léxico global, no cuelgan de `window`, y
   * escribirles desde aquí creaba unas variables nuevas con el mismo
   * nombre mientras las de la pantalla seguían en falso. El síntoma era
   * silencioso y perfecto: la app abría, se veía completa, y mostraba
   * los datos de ejemplo del HTML en vez de los del demo.
   *
   * Y es peor porque replicar el arranque a mano significa mantener dos
   * arranques. Lo único que hace falta es poner el token —con `eval`
   * indirecto, que es la forma de tocar una variable léxica global— y
   * dejar que `reanudarSesion()` haga exactamente lo que hace cada día:
   * preguntar quién es, y entrar con lo que le contesten. Como `api` ya
   * está reemplazada, quien contesta es el demo.
   */
  function entrar(rol) {
    abrirSesion(rol);
    const g = document.getElementById('demo-gate');
    if (g) g.remove();
    cinta();
    (0, eval)('TOKEN = "demo-sin-servidor";');
    window.reanudarSesion();
    // La etiqueta de origen la pinta cargarReales() cuando termina, y
    // se vuelve a pintar cada vez que se cambia de tienda
    setInterval(reetiquetarOrigen, 600);
  }

  // ─── 11 · REEMPLAZAR LA CONEXIÓN ──────────────────────────────
  /**
   * Desde aquí la pantalla ya no tiene a dónde llamar.
   *
   * `api` es el único punto por donde Nova habla con el servidor. Al
   * reemplazarla, el modo demostración queda aislado por construcción:
   * no es que se evite escribir, es que no hay nada a dónde escribir.
   */
  window.api = async function (accion, datos) {
    const p = Object.assign({}, datos || {});
    if (window.VISTA) p.vista = window.VISTA;

    if (accion === 'login') {
      return { ok: false, error: 'Estás en el demo: no hace falta correo ni código. ' +
               'Recarga la página y elige con qué rol quieres entrar.' };
    }
    if (accion === 'verificar') {
      return { ok: false, error: 'Estás en el demo. Recarga y elige un rol.' };
    }
    if (!SES && accion !== 'yo') return { ok: false, error: 'Elige un rol para empezar.' };

    const fn = ACCIONES[accion];
    if (!fn) return { ok: false, error: 'El demo todavía no responde "' + accion + '".' };

    // Una pausa corta: sin ella todo aparece instantáneo y no se entiende
    // que hay una consulta detrás. Con ella se ve el "cargando" real.
    await new Promise(function (ok) { setTimeout(ok, 120 + Math.random() * 130); });

    /**
     * Lo que cambia datos tira el mes guardado.
     *
     * Se decide aquí y no dentro de cada acción a propósito: si cada una
     * tuviera que acordarse de invalidar, el día que se agregue una
     * acción nueva alguien lo va a olvidar y la pantalla va a mostrar
     * cifras viejas sin ningún síntoma. Es más seguro que la lista diga
     * qué NO cambia nada.
     */
    const SOLO_LEEN = ['yo', 'resumen', 'listar', 'productos', 'equipo', 'fuentes',
                       'cierre', 'historial', 'cas', 'alarmas', 'recuento',
                       'auditoria', 'trozo'];

    try {
      const r = await fn(p);
      if (SOLO_LEEN.indexOf(accion) === -1) invalidarCache();
      return r;
    }
    catch (e) {
      console.error('[demo]', accion, e);
      return { ok: false, error: 'El demo se atoró en "' + accion + '": ' + e.message };
    }
  };

  /**
   * Que una visita al demo no deje ni se lleve nada de una sesión real.
   *
   * Se borra el token guardado ANTES de que la página intente reanudar
   * con él. Sin esto, alguien que usa Nova de verdad y abre el enlace
   * del demo en la misma pestaña entraría con su propia sesión a medias.
   */
  try {
    localStorage.removeItem('ne_token');
    localStorage.removeItem('ne_email');
  } catch (e) { /* navegador sin almacenamiento: mejor todavía */ }
  try { (0, eval)('TOKEN = "";'); } catch (e) { /* aún no existe: se pone al entrar */ }

  // ─── 12 · ARRANCAR ────────────────────────────────────────────
  sembrar();

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', puerta);
  } else {
    puerta();
  }

  // Para poder mirar por dentro desde la consola si algo no cuadra
  window.NOVA_DEMO = { HOJA: HOJA, meses: MESES,
                       agregarMes: mesDe_, sinCache: agregarMesCrudo,
                       invalidar: invalidarCache };
})();
