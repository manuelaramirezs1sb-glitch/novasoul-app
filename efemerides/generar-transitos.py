#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Genera la tabla de tránsitos de una carta, para pegar en la hoja Transitos.

┌─ POR QUÉ ESTO NO ES UNA API ───────────────────────────────┐
│                                                            │
│ Nova vive en Apps Script y no sabe dónde está Saturno el   │
│ martes que viene. Eso pide efemérides, y las efemérides no │
│ están dentro de Google.                                    │
│                                                            │
│ La salida obvia era pagar una API. Antes de hacerlo se     │
│ probó el motor que usan esas APIs por dentro —Swiss        │
│ Ephemeris— contra la carta REAL de ella, la que le da su   │
│ app Horus. Reproduce los once planetas, el Ascendente y el │
│ Mediocielo, al grado.                                      │
│                                                            │
│ Y un año entero de sus temporadas son 33 filas. Cinco      │
│ años, unas 165. Eso cabe en una hoja sin despeinarse.      │
│                                                            │
│ Así que no hay llave que cuidar, no hay factura, y no hay  │
│ un servicio ajeno del que dependa su pantalla. El costo    │
│ es este archivo: hay que volver a correrlo cuando se       │
│ acabe el horizonte. Es un rato cada varios años.           │
│                                                            │
│ La puerta a una API queda abierta: la hoja es la misma, y  │
│ quien la llene puede ser esto o puede ser otra cosa.       │
│                                                            │
└────────────────────────────────────────────────────────────┘

┌─ LAS DOS CASAS ────────────────────────────────────────────┐
│                                                            │
│ Horus usa PLACIDUS. Se comprobó contra las diez casas de   │
│ su captura: Placidus acierta 10 de 10, casas enteras 6.    │
│                                                            │
│ Pero las PROFECCIONES son de casas enteras por técnica, y  │
│ así están hechas en Nova. Los dos sistemas son correctos   │
│ en su sitio, y para el mismo planeta dan números distintos │
│ cuatro veces de cada diez.                                 │
│                                                            │
│ Por eso se escriben LAS DOS. Ella decide cuál mirar, y no  │
│ se entera por una contradicción que nadie le explicó.      │
│                                                            │
└────────────────────────────────────────────────────────────┘

USO:
    python3 generar-transitos.py --config manuela.json --anios 5

Necesita:  pip install pyswisseph
"""
import argparse
import json
import sys
from datetime import date, timedelta

try:
    import swisseph as swe
except ImportError:
    sys.exit('Falta pyswisseph.  pip install pyswisseph')

SIGNOS = ['Aries', 'Tauro', 'Géminis', 'Cáncer', 'Leo', 'Virgo', 'Libra',
          'Escorpio', 'Sagitario', 'Capricornio', 'Acuario', 'Piscis']
SIGNOS_ID = ['aries', 'tauro', 'geminis', 'cancer', 'leo', 'virgo', 'libra',
             'escorpio', 'sagitario', 'capricornio', 'acuario', 'piscis']

# Los grupos, como los piensa ella: los personales afectan lo inmediato,
# los sociales hablan de expansión y estructura, los generacionales
# marcan época. Quirón queda fuera: pide un archivo de efemérides aparte
# que no viene con la librería, y meterlo a medias sería peor.
PERSONALES = [('sol', 'Sol', swe.SUN), ('mercurio', 'Mercurio', swe.MERCURY),
              ('venus', 'Venus', swe.VENUS), ('marte', 'Marte', swe.MARS)]
SOCIALES = [('jupiter', 'Júpiter', swe.JUPITER), ('saturno', 'Saturno', swe.SATURN)]
GENERACIONALES = [('urano', 'Urano', swe.URANUS), ('neptuno', 'Neptuno', swe.NEPTUNE),
                  ('pluton', 'Plutón', swe.PLUTO)]

# Los cuerpos natales que vale la pena que alguien toque. La Luna natal
# entra; la Luna EN TRÁNSITO no, porque cambia de signo cada dos días y
# media y Nova ya calcula su fase sola.
NATALES = [('sol', 'Sol', swe.SUN), ('luna', 'Luna', swe.MOON),
           ('mercurio', 'Mercurio', swe.MERCURY), ('venus', 'Venus', swe.VENUS),
           ('marte', 'Marte', swe.MARS), ('jupiter', 'Júpiter', swe.JUPITER),
           ('saturno', 'Saturno', swe.SATURN), ('urano', 'Urano', swe.URANUS),
           ('neptuno', 'Neptuno', swe.NEPTUNE), ('pluton', 'Plutón', swe.PLUTO)]

ASPECTOS = [(0, 'conjunción', 6), (60, 'sextil', 4), (90, 'cuadratura', 5),
            (120, 'trígono', 5), (180, 'oposición', 6)]

# Cuántos días tiene que durar algo para llamarse temporada. Un planeta
# personal pasa en días y sirve para la semana; uno lento dura meses y es
# lo que de verdad puede regir una temporada del pensum.
MINIMO = {'personal': 3, 'social': 14, 'generacional': 14}


def separacion(a, b):
    d = abs(a - b) % 360
    return 360 - d if d > 180 else d


class Carta:
    """La carta natal, y las dos formas de contar sus casas."""

    def __init__(self, cfg):
        h = cfg['hora_utc']
        self.jd = swe.julday(cfg['anio'], cfg['mes'], cfg['dia'],
                             h[0] + h[1] / 60.0)
        self.lat, self.lon = cfg['lat'], cfg['lon']
        self.cuspides, ascmc = swe.houses(self.jd, self.lat, self.lon, b'P')
        self.asc = ascmc[0]
        self.mc = ascmc[1]
        self.asc_signo = int(self.asc // 30)

        self.natal = {}
        for cid, nom, cuerpo in NATALES:
            self.natal[cid] = (nom, swe.calc_ut(self.jd, cuerpo)[0][0])
        self.natal['ascendente'] = ('Ascendente', self.asc)
        self.natal['medio_cielo'] = ('Mediocielo', self.mc)

    def casa_entera(self, lon):
        """Un signo, una casa. Es como se hacen las profecciones."""
        return ((int(lon // 30) - self.asc_signo) % 12) + 1

    def casa_placidus(self, lon):
        """El sistema que usa Horus. Es el que coincide con su carta."""
        for i in range(12):
            a, b = self.cuspides[i], self.cuspides[(i + 1) % 12]
            if a < b:
                if a <= lon < b:
                    return i + 1
            elif lon >= a or lon < b:
                return i + 1
        return None


def generar(carta, desde, hasta, incluir_personales):
    grupos = [('social', SOCIALES), ('generacional', GENERACIONALES)]
    if incluir_personales:
        grupos.insert(0, ('personal', PERSONALES))

    abiertos, eventos = {}, []
    d = desde
    while d <= hasta:
        jd = swe.julday(d.year, d.month, d.day, 12.0)
        for grupo, lista in grupos:
            for tid, tnom, tcuerpo in lista:
                pos = swe.calc_ut(jd, tcuerpo)[0]
                tlon, retro = pos[0], pos[3] < 0
                for nid, (nnom, nlon) in carta.natal.items():
                    if nid == tid and grupo != 'personal':
                        # El retorno de un planeta a su propio sitio sí
                        # cuenta, y es de los que más significan.
                        pass
                    s = separacion(tlon, nlon)
                    for ang, anom, orbe in ASPECTOS:
                        clave = (tid, nid, anom)
                        dentro = abs(s - ang) <= orbe
                        if dentro and clave not in abiertos:
                            abiertos[clave] = {
                                'desde': d, 'grupo': grupo, 'tnom': tnom,
                                'nnom': nnom, 'aspecto': anom,
                                'casa_e': carta.casa_entera(tlon),
                                'casa_p': carta.casa_placidus(tlon),
                                'retro': retro, 'pico': abs(s - ang),
                                'fpico': d,
                            }
                        elif dentro:
                            a = abiertos[clave]
                            if abs(s - ang) < a['pico']:
                                a['pico'], a['fpico'] = abs(s - ang), d
                                a['casa_e'] = carta.casa_entera(tlon)
                                a['casa_p'] = carta.casa_placidus(tlon)
                        elif clave in abiertos:
                            a = abiertos.pop(clave)
                            a['hasta'] = d - timedelta(days=1)
                            if (a['hasta'] - a['desde']).days >= MINIMO[a['grupo']]:
                                eventos.append(a)
        d += timedelta(days=1)

    # Lo que sigue abierto al final del horizonte se cierra ahí, y se dice.
    for a in abiertos.values():
        a['hasta'] = hasta
        a['sigue'] = True
        if (a['hasta'] - a['desde']).days >= MINIMO[a['grupo']]:
            eventos.append(a)

    eventos.sort(key=lambda e: (e['desde'], e['tnom']))
    return eventos


# Las columnas de la hoja Transitos, en su orden. `casa` es la de casas
# enteras, que es lo que el código de Nova ya usa; `casa_placidus` es la
# que coincide con Horus. Van las dos porque ella pidió ver las dos.
COLUMNAS = ['usuario_id', 'fecha', 'casa', 'casa_placidus', 'tema',
            'intensidad_pct', 'texto_transito', 'por_que', 'como_trabajarlo',
            'el_otro_lado', 'cuerpo', 'aspecto', 'a_natal', 'desde', 'hasta',
            'fuente']


def fila(e, uid):
    titulo = '%s %s a mi %s' % (e['tnom'], e['aspecto'], e['nnom'])
    return {
        'usuario_id': uid,
        'fecha': e['fpico'].isoformat(),      # el día que aprieta más
        'casa': e['casa_e'],
        'casa_placidus': e['casa_p'],
        'tema': titulo,
        'intensidad_pct': '',
        'texto_transito': titulo + (' (retrógrado)' if e['retro'] else ''),
        'por_que': '',
        'como_trabajarlo': '',
        'el_otro_lado': '',
        'cuerpo': e['tnom'].lower().replace('ó', 'o').replace('ú', 'u'),
        'aspecto': e['aspecto'],
        'a_natal': e['nnom'],
        'desde': e['desde'].isoformat(),
        'hasta': e['hasta'].isoformat(),
        'fuente': 'Swiss Ephemeris',
    }


def main():
    ap = argparse.ArgumentParser(description=__doc__,
                                 formatter_class=argparse.RawDescriptionHelpFormatter)
    ap.add_argument('--config', required=True, help='JSON con los datos de nacimiento')
    ap.add_argument('--anios', type=int, default=5, help='horizonte en años (5 por defecto)')
    ap.add_argument('--desde', help='AAAA-MM-DD; por defecto, hoy')
    ap.add_argument('--personales', action='store_true',
                    help='incluir Sol, Mercurio, Venus y Marte (muchas más filas)')
    ap.add_argument('--salida', default='transitos.tsv')
    args = ap.parse_args()

    cfg = json.load(open(args.config, encoding='utf-8'))
    carta = Carta(cfg)

    desde = date.fromisoformat(args.desde) if args.desde else date.today()
    hasta = date(desde.year + args.anios, desde.month, desde.day) - timedelta(days=1)

    print('Carta de %s' % cfg.get('nombre', '—'))
    print('  Ascendente  %s %.2f°' % (SIGNOS[int(carta.asc // 30)], carta.asc % 30))
    print('  Mediocielo  %s %.2f°' % (SIGNOS[int(carta.mc // 30)], carta.mc % 30))
    print('  Horizonte   %s → %s' % (desde, hasta))
    print()

    eventos = generar(carta, desde, hasta, args.personales)
    filas = [fila(e, cfg.get('usuario_id', '')) for e in eventos]

    with open(args.salida, 'w', encoding='utf-8') as f:
        f.write('\t'.join(COLUMNAS) + '\n')
        for r in filas:
            f.write('\t'.join(str(r[c]) for c in COLUMNAS) + '\n')

    difieren = sum(1 for e in eventos if e['casa_e'] != e['casa_p'])
    por_grupo = {}
    for e in eventos:
        por_grupo[e['grupo']] = por_grupo.get(e['grupo'], 0) + 1

    print('%d temporadas → %s' % (len(filas), args.salida))
    for g in ('personal', 'social', 'generacional'):
        if g in por_grupo:
            print('   %-14s %d' % (g, por_grupo[g]))
    print()
    print('En %d de %d, las casas enteras y Placidus NO coinciden.' % (difieren, len(eventos)))
    print('Por eso van las dos columnas: que lo decida ella, viéndolas.')


if __name__ == '__main__':
    main()
