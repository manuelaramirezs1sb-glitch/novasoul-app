#!/usr/bin/env python3
"""
Arma el demo en UN SOLO ARCHIVO.

Por qué existe
──────────────
El demo normal se enciende con ?demo=1 sobre la Nova publicada. Eso
supone que hay dónde publicarla. Cuando no lo hay —el hosting se acabó,
o simplemente quieres mandárselo a alguien por correo— hace falta algo
que funcione solo: un archivo que se abre con doble clic, sin internet,
sin servidor y sin instalar nada.

Este script produce ese archivo. Mete dentro el guion del demo, convierte
los logos a texto para que viajen con él, y le quita la dirección del
servidor de Apps Script.

Lo de la dirección no es cosmético. El demo nunca la llama —api() está
reemplazada— pero dejarla escrita en un archivo que va a circular por
correo sería repartir el número de la puerta de la casa con la excusa de
que nadie la va a tocar. Si no se usa, no viaja.

    python3 construir-demo.py    →    nova-demo.html
"""
import base64
import pathlib
import re
import sys

AQUI = pathlib.Path(__file__).parent
SALIDA = AQUI / 'nova-demo.html'


def texto(nombre):
    return (AQUI / nombre).read_text(encoding='utf-8')


def imagen_en_texto(nombre):
    """Una imagen convertida a algo que quepa dentro del HTML."""
    datos = base64.b64encode((AQUI / nombre).read_bytes()).decode('ascii')
    return 'data:image/png;base64,' + datos


def main():
    html = texto('empresarial.html')
    demo = texto('demo.js')

    cambios = []

    # ── 1 · Los logos viajan dentro ──
    for png in ['logo-nova-empresarial-claro.png', 'logo-nova-empresarial.png']:
        uri = imagen_en_texto(png)
        antes = html.count(png) + demo.count(png)
        if not antes:
            sys.exit('No encontré referencias a ' + png)
        html = html.replace(png, uri)
        demo = demo.replace(png, uri)
        cambios.append('%-34s %d referencia(s), %d KB' % (png, antes, len(uri) // 1024))

    # ── 2 · Fuera la dirección del servidor ──
    patron = re.compile(r"const NOVA_API = '[^']*';")
    if not patron.search(html):
        sys.exit('No encontré la constante NOVA_API: revisa empresarial.html')
    html = patron.sub(
        "const NOVA_API = '';   // el demo no habla con ningún servidor",
        html, count=1)
    cambios.append('%-34s borrada' % 'dirección del servidor')

    # ── 3 · El guion del demo, adentro y encendido ──
    cargador = re.compile(
        r"<script>\s*\(function \(\) \{\s*if \(!/\[\?&\]demo.*?</script>",
        re.S)
    if not cargador.search(html):
        sys.exit('No encontré el cargador del demo al final de empresarial.html')
    # Con una función y no con una cadena: el guion trae barras invertidas
    # (las expresiones regulares de las fechas) y como texto de reemplazo
    # Python intentaría interpretarlas.
    adentro = ('<script>window.NOVA_DEMO_FORZADO = true;</script>\n'
               '<script>\n' + demo + '\n</script>')
    html = cargador.sub(lambda _: adentro, html, count=1)
    cambios.append('%-34s %d KB adentro' % ('demo.js', len(demo) // 1024))

    # ── 4 · Que se note qué es, desde la pestaña ──
    html = html.replace('<title>Nova Empresarial</title>',
                        '<title>Nova Empresarial · Demostración</title>', 1)

    SALIDA.write_text(html, encoding='utf-8')

    print('\n'.join('  ' + c for c in cambios))
    print('\n  %-34s %d KB' % (SALIDA.name, len(html.encode('utf-8')) // 1024))
    print('  Se abre con doble clic. No necesita internet ni servidor.')

    # Una última revisión: que no se haya quedado nada por fuera
    sueltos = re.findall(r'(?:src|href)="(?!https?:|data:|#)([^"]+\.(?:png|js|css))"', html)
    if sueltos:
        print('\n  OJO · quedaron archivos sueltos que no viajan: ' +
              ', '.join(sorted(set(sueltos))))
    if 'script.google.com' in html:
        sys.exit('\n  ALTO · todavía hay una dirección de Apps Script adentro.')


if __name__ == '__main__':
    main()
