#!/usr/bin/env python3
"""
Prepara las capturas de la guía de Meta.

Las capturas vienen de la pantalla de Manuela mientras hacía el proceso
de verdad — por eso muestran pasos reales de 2026 y no una simulación.
Pero también muestran su nombre, su foto, sus identificadores de negocio
y sus cuentas publicitarias con el gasto.

Nada de eso puede salir en una guía que van a leer los clientes. Este
script recorta lo que importa y tapa lo demás antes de publicar.

Se tapa con un rectángulo opaco, no con desenfoque: un desenfoque se
puede deshacer y además invita a entrecerrar los ojos. Lo que se tapa,
se tapa.

    python3 guia/limpiar.py
"""
import pathlib
from PIL import Image, ImageDraw

AQUI = pathlib.Path(__file__).parent
ORIGEN = pathlib.Path('/tmp/claude-0/-home-claude-repo/'
                      '94f95c50-2e20-5c67-9501-7ae933f16b73/images')
DESTINO = AQUI / 'img'
TAPA = (232, 234, 238)     # gris claro, del color de los paneles de Meta

# nombre de salida · archivo · recorte (x1,y1,x2,y2) o None · zonas a tapar
RECETAS = [
    ('01-aplicaciones.png', '22.png', (280, 150, 1010, 530), [
        (280, 150, 1010, 200),      # el encabezado con el nombre del portfolio
    ]),
    ('02-registro-rol.png', '23.png', (170, 90, 1180, 630), []),
    ('03-crear-app.png', '24.png', (20, 90, 1340, 420), [
        (1280, 95, 1345, 140),      # la foto de perfil
    ]),
    ('04-aviso-nuevo.png', '25.png', (340, 155, 990, 600), []),
    ('05-casos-de-uso.png', '26.png', (170, 210, 1180, 700), []),
    ('06-marketing-api.png', '28.png', None, []),
]


def main():
    DESTINO.mkdir(parents=True, exist_ok=True)
    for salida, fuente, recorte, zonas in RECETAS:
        ruta = ORIGEN / fuente
        if not ruta.exists():
            print('  falta', fuente, '— se salta')
            continue

        im = Image.open(ruta).convert('RGB')
        d = ImageDraw.Draw(im)
        for z in zonas:
            d.rectangle(z, fill=TAPA)
        if recorte:
            im = im.crop(recorte)

        # Un borde tenue para que la captura se despegue del fondo
        borde = Image.new('RGB', (im.width + 2, im.height + 2), (214, 216, 220))
        borde.paste(im, (1, 1))

        borde.save(DESTINO / salida, optimize=True)
        print('  %-26s %sx%s' % (salida, borde.width, borde.height))


if __name__ == '__main__':
    main()
