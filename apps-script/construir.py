#!/usr/bin/env python3
"""
Arma NOVA-COMPLETO.gs desde los archivos fuente.

Apps Script no deja importar archivos, así que hay que pegar todo junto.
Durante semanas mantuve las dos cosas a mano —editar la fuente y luego
parchear el combinado— y eso produjo tres errores del mismo tipo: código
que existía en la fuente y no llegaba al archivo que de verdad corre.
El último dejó a la dueña sin poder agregar a nadie a su equipo.

Un archivo generado no puede divergir de su fuente. Se corre y ya.
"""
import pathlib
import sys

AQUI = pathlib.Path(__file__).parent

# El orden importa: las constantes y utilidades antes de quien las usa.
ORDEN = [
    ('00-bootstrap.gs',    '1 · INSTALACIÓN'),
    ('10-estados.gs',      '2 · ESTADOS Y TELÉFONOS'),
    ('15-monedas.gs',      '3 · MONEDAS'),
    ('20-importadores.gs', '4 · FUENTES E IMPORTADORES'),
    ('30-automapeo.gs',    '5 · MAPEO AUTOMÁTICO'),
    ('30-alarmas.gs',      '6 · ALARMAS'),
    ('40-provisionar.gs',  '7 · CLIENTES'),
    ('50-tasas.gs',        '8 · TASAS DE CAMBIO'),
    ('60-api.gs',          '9 · API WEB'),
    ('65-central.gs',      '10 · NOVA CENTRAL'),
    ('66-central-mio.gs',  '10b · CENTRAL · TRABAJOS Y PLATA'),
    ('70-importar.gs',     '11 · ESCRITURA'),
    ('75-meta.gs',         '12 · META'),
    ('76-meta-leer.gs',    '13 · META · LECTURA DIARIA'),
    ('77-semaforo.gs',     '14 · SEMÁFORO SEMANAL'),
    ('78-soul.gs',         '15 · NOVASOUL · EL DÍA A DÍA'),
    ('79-silabo.gs',       '16 · NOVASOUL · LA UNIVERSIDAD'),
    ('80-rutina.gs',       '17 · NOVASOUL · RUTINA, TURNOS Y PLATA'),
    ('81-proyecto.gs',     '18 · CENTRAL · UN PROYECTO A FONDO'),
    ('82-cielo.gs',        '19 · NOVASOUL · EL CIELO'),
    ('83-automatico-mio.gs', '20 · LO AUTOMÁTICO DE CENTRAL Y SOUL'),
    ('84-central-meta.gs',  '21 · CENTRAL · META, DESDE LA CONSOLA'),
    ('85-meta-panel.gs',   '22 · META · TODO EN UNA PANTALLA'),
    ('86-reporte-dia.gs',  '23 · QUÉ QUEDÓ DE AYER'),
    ('87-lectura.gs',      '24 · EL CIELO · EL VOCABULARIO DE LA LECTURA'),
    ('88-efemerides.gs',   '25 · EL CIELO · LAS EFEMÉRIDES YA CALCULADAS'),
    ('89-cielo-rangos.gs', '26 · EL CIELO · SEMANA, MES Y AÑO'),
    ('90-plata.gs',        '27 · MI PLATA · ENTRA, SALE, CUOTAS'),
    ('91-arranque-soul.gs','28 · NOVASOUL EN UNA SOLA PETICIÓN'),
    ('92-auditoria.gs',    '29 · AUDITORÍA · CASOS REALES'),
    ('93-central-hoy.gs',  '30 · CENTRAL · HOY, LEYENDO LA RED'),
    ('94-asignar.gs',      '31 · ASIGNAR · REPARTIR EL TRABAJO'),
    ('95-propio.gs',       '32 · EL CONTROL DIARIO · NOVEDADES Y CAS'),
    ('96-chat.gs',         '33 · EL CHAT DEL EQUIPO, DENTRO DE NOVA'),
    ('97-accesos.gs',      '34 · LOS ACCESOS DE LA TIENDA'),
    ('98-notas.gs',        '35 · LA BITÁCORA DE INTENTOS'),
    ('99-arranque-emp.gs', '36 · EMPRESARIAL EN UNA SOLA PETICIÓN'),
    ('A0-horario.gs',      '37 · QUE NOVA ORGANICE LA SEMANA'),
]

SALIDA = AQUI / 'NOVA-COMPLETO.gs'


def banner(titulo):
    linea = '═' * 63
    return f'/* {linea}\n   {titulo}\n   {linea} */\n\n'


def main():
    partes = []
    faltan = []
    for nombre, titulo in ORDEN:
        ruta = AQUI / nombre
        if not ruta.exists():
            faltan.append(nombre)
            continue
        partes.append(banner(titulo) + ruta.read_text(encoding='utf-8').rstrip() + '\n')

    if faltan:
        print('FALTAN archivos: ' + ', '.join(faltan), file=sys.stderr)
        return 1

    texto = '\n\n'.join(partes)

    # ── EL SELLO DEL BACKEND ──────────────────────────────────
    #
    # La pantalla lleva `VERSION_PANTALLA` desde hace meses, justo por
    # esto: sin una marca visible no se puede saber si lo que corre es
    # lo nuevo o lo de ayer. El .gs no lo llevaba, y se pagó.
    #
    # Le mandé cuatro veces el mismo archivo en una tarde —todos
    # llamados NOVA-COMPLETO.gs, sin forma de distinguirlos— y pegó el
    # tercero. El error que reportó como nuevo era el viejo, en la línea
    # exacta del archivo anterior, y lo único que lo delató fue que el
    # número de línea del error no cuadraba con el mío.
    #
    # Con el sello se pregunta y se responde en diez segundos.
    import datetime
    import hashlib

    # La huella es del CUERPO, sin la cabecera: si se calculara sobre el
    # archivo entero cambiaría al escribirse a sí misma.
    huella = hashlib.sha256(texto.encode('utf-8')).hexdigest()[:6]

    PLANTILLA = (
        '/* NOVA-COMPLETO.gs · generado por apps-script/construir.py\n'
        '   NO SE EDITA A MANO: los cambios van en los archivos numerados.\n'
        '   {sello}\n'
        '   Para saber qué versión está corriendo: ejecuta queVersion() */\n\n'
        "const NOVA_GS = '{sello}';\n\n"
    )
    # Y las líneas son las del archivo FINAL, cabecera incluida. Es el
    # número con el que se compara contra el que alguien tiene abierto:
    # un conteo que no cuadra con lo que se ve en pantalla no sirve para
    # lo único que existe.
    #
    # `count('\n') + 1` y no `wc -l`: el editor de Apps Script numera la
    # última línea vacía, así que este es el número que ella ve abajo del
    # todo. `wc -l` daría uno menos y la comparación fallaría por uno.
    lineas = texto.count('\n') + 1 + PLANTILLA.count('\n')
    sello = f"{datetime.date.today().isoformat()} · {lineas} líneas · {huella}"
    texto = PLANTILLA.format(sello=sello) + texto
    SALIDA.write_text(texto, encoding='utf-8')

    real = texto.count('\n') + 1
    if real != lineas:                      # el sello no puede mentir
        print(f'ERROR: el sello dice {lineas} líneas y el archivo tiene {real}',
              file=sys.stderr)
        return 1

    print(f'NOVA-COMPLETO.gs: {len(ORDEN)} archivos, {lineas} líneas · {huella}')

    # Una función declarada dos veces no falla al pegar, pero la segunda
    # gana en silencio. Vale la pena avisar.
    import re
    nombres = re.findall(r'^function\s+([A-Za-z_$][\w$]*)\s*\(', texto, re.M)
    repes = {n for n in nombres if nombres.count(n) > 1}
    if repes:
        print('OJO, funciones repetidas: ' + ', '.join(sorted(repes)), file=sys.stderr)

    consts = re.findall(r'^const\s+([A-Z_][A-Z0-9_]*)\s*=', texto, re.M)
    repesC = {n for n in consts if consts.count(n) > 1}
    if repesC:
        print('ERROR, constantes repetidas (no carga): ' + ', '.join(sorted(repesC)),
              file=sys.stderr)
        return 1
    return 0


if __name__ == '__main__':
    sys.exit(main())
