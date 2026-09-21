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

    # ── 4 · El lector de Excel, dormido adentro ──
    #
    # Va en una etiqueta que el navegador NO ejecuta. Si fuera un <script>
    # normal, el archivo tardaría casi un segundo más en abrir para todo el
    # mundo, incluida la mayoría que solo viene a mirar y nunca sube nada.
    # demo.js lo despierta la primera vez que alguien sube un .xlsx.
    #
    # Solo se escapa </script, que es lo único que corta la etiqueta.
    #
    # El primer intento escapaba TODO "</" y eso rompió la librería: entre
    # sus 329 apariciones hay expresiones regulares como /\s+</g, que al
    # volverse /\s+<\/g dejan de terminar donde deben. El archivo se
    # construía sin quejarse y reventaba al leer el primer Excel.
    #
    # De ahí la comprobación de más abajo: escapar código ajeno es fácil
    # de hacer mal y el error no aparece hasta que alguien lo usa.
    lector = re.sub(r'</(?=script)', r'<\\/', texto('vendor/xlsx.min.js'),
                    flags=re.I)
    etiqueta = ('<script type="text/plain" id="nova-xlsx">'
                + lector + '</' + 'script>\n')
    html = html.replace('<script>window.NOVA_DEMO_FORZADO = true;</script>',
                        etiqueta + '<script>window.NOVA_DEMO_FORZADO = true;</script>', 1)
    cambios.append('%-34s %d KB, dormido' % ('lector de Excel', len(lector) // 1024))

    # ── 5 · Que se note qué es, desde la pestaña ──
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

    comprobar_lector(html)


def comprobar_lector(html):
    """
    Que el lector de Excel siga siendo JavaScript válido después de meterlo.

    Existe porque ya falló una vez: un escape de más lo dejó con una
    expresión regular sin cerrar. El archivo se construyó igual, se veía
    perfecto, y solo reventaba cuando alguien subía un Excel — es decir,
    en manos del cliente y no aquí.
    """
    import subprocess
    import tempfile

    m = re.search(r'<script type="text/plain" id="nova-xlsx">(.*?)</script>',
                  html, re.S)
    if not m:
        sys.exit('\n  ALTO · el lector de Excel no quedó adentro.')

    cuerpo = m.group(1).replace('<\\/script', '</script')
    with tempfile.NamedTemporaryFile('w', suffix='.js', delete=False,
                                     encoding='utf-8') as f:
        f.write(cuerpo)
        tmp = f.name

    try:
        r = subprocess.run(['node', '--check', tmp],
                           capture_output=True, text=True, timeout=60)
    except (FileNotFoundError, subprocess.TimeoutExpired):
        print('  (sin node: no se pudo comprobar el lector)')
        return

    if r.returncode != 0:
        sys.exit('\n  ALTO · el lector de Excel quedó roto al meterlo:\n  ' +
                 r.stderr.strip().split('\n')[0])
    print('  %-34s sigue siendo válido' % 'lector comprobado')


if __name__ == '__main__':
    main()
