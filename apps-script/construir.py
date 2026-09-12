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
    ('70-importar.gs',     '10 · ESCRITURA'),
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
    SALIDA.write_text(texto, encoding='utf-8')

    lineas = texto.count('\n') + 1
    print(f'NOVA-COMPLETO.gs: {len(ORDEN)} archivos, {lineas} líneas')

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
