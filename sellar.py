#!/usr/bin/env python3
"""
Sella cada pantalla con la fecha y una huella de su contenido.

"No me carga el cambio" es imposible de diagnosticar a ciegas: el
navegador guarda el HTML y lo sigue sirviendo aunque el servidor ya
tenga otro. Con un sello visible, saber si lo que está abierto es lo
último toma dos segundos en vez de media hora.

La huella sale del contenido del archivo, así que cambia sola cada vez
que el archivo cambia — aunque sean dos versiones el mismo día.

    python3 sellar.py     (antes de cada commit que toque un HTML)
"""
import hashlib, re, datetime, pathlib

HOY = datetime.date.today().isoformat()
RAIZ = pathlib.Path(__file__).parent

for nombre in ['empresarial.html', 'novacentral.html', 'index.html', 'novasoul.html']:
    f = RAIZ / nombre
    if not f.exists():
        continue
    txt = f.read_text(encoding='utf-8')
    # La huella se calcula SIN el sello, o cambiaría cada vez que se sella
    limpio = re.sub(r"const VERSION_PANTALLA = '[^']*';", '', txt)
    huella = hashlib.sha256(limpio.encode('utf-8')).hexdigest()[:4]
    sello = HOY + ' · ' + huella

    nuevo, n = re.subn(r"const VERSION_PANTALLA = '[^']*';",
                       "const VERSION_PANTALLA = '" + sello + "';", txt)
    if n:
        f.write_text(nuevo, encoding='utf-8')
        print('%-20s %s' % (nombre, sello))
    else:
        print('%-20s (sin sello que actualizar)' % nombre)
