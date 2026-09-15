# -*- coding: utf-8 -*-
"""Formato de celda del mapa de PROGRAMACION.

Reproduce el mismo estilo que 'Cristour 2026 planilha.xlsx' (planilla modelo):
  - IN/OUT con ciudad del vuelo (IQQ/FLN/SCL...), horas reales si existen.
  - Dia con 'Salida p/<ATRACCION>', 'as HH:MMH', sub-atracciones con '+',
    restaurante del almuerzo, cena, noche y '5ª COMIDA'.
  - Vehiculo al final (BUS / MICRO / VAN / n BUS).

DICCIONARIO INTERNO POR ATRACCION: cuando el Excel/Word no trae el dato
(hora, sub-atraccion, restaurante...), se usa este diccionario y se deja
'____' para completar a mano si falta.
"""

import re
from datetime import datetime

__all__ = ['texto_in', 'texto_out', 'texto_dia', 'HORAS_DIC',
           'PERFILES', 'TEXTOS_AUX']


def _hh(v):
    """'09:30' / '09:30H' / '09.30' / '09' / datetime -> '09:30H'. None -> '____H'."""
    if v is None:
        return '____H'
    if isinstance(v, datetime):
        return v.strftime('%H:%M') + 'H'
    s = str(v).strip().lower().replace('hs', '').replace('h', '').rstrip('.')
    s = s.replace(';', ':').replace('.', ':')
    if ':' in s:
        p = s.split(':')[:2]
        try:
            return f"{int(p[0]):02d}:{int(p[1]):02d}H"
        except ValueError:
            return '____H'
    m = re.match(r'^(\d{1,4})$', s)
    if m:
        s2 = s.zfill(4)
        return f"{s2[:2]}:{s2[2:]}H"
    return '____H'


def _txt(v):
    """Devuelve el valor o '____' para llenar a mano."""
    if v is None or (isinstance(v, str) and not v.strip()):
        return '____'
    return str(v).strip()


def _vuelo(v):
    if v is None or not str(v).strip():
        return ' ____'
    return ' ' + str(v).strip()


def _ciudad(sigla):
    """Sigla de aeropuerto (IQQ, FLN, SCL...). None/'' -> '' (sin ciudad)."""
    return (sigla or '').strip().upper()


def texto_in(vuelo, sal_cod, sal_hora, leg_cod, leg_hora, marco_in=None,
             veh='BUS'):
    """Celda del dia IN (igual que la planilla modelo).
    marco_in: (titulo, [subtitulo...]) opcional para 'Noche bajo las estrellas'."""
    L = ['IN - Vuelo:', _vuelo(vuelo), '']
    L += [f'Salida {_ciudad(sal_cod)}:', _hh(sal_hora), ''] if sal_cod \
        else ['Salida:', _hh(sal_hora), '']
    L += [f'Llegada {_ciudad(leg_cod)}:', _hh(leg_hora)] if leg_cod \
        else ['Llegada:', _hh(leg_hora)]
    L += ['', '+', 'Almuerzo', '', '+', 'Cena hotel ']
    if marco_in:
        L += ['', '+', marco_in[0].upper(), *(marco_in[1:] or []),
              '+', '5ª COMIDA']
    L += ['', '', '', veh]
    return '\n'.join(L)


def texto_out(vuelo, bus_hora=None, sal_cod=None, sal_hora=None,
              leg_cod=None, leg_hora=None, fech_leg=None, refri=None,
              veh='BUS'):
    """Celda del dia OUT (igual que la planilla modelo).
    bus_hora: salida del bus hacia el aeropuerto.
    sal_cod/sal_hora: ciudad y hora de salida del vuelo.
    leg_cod/leg_hora: ciudad y hora de llegada del vuelo."""
    L = ['Box Lunch', '', '', 'Salida p/aeropuerto', _hh(bus_hora), '',
         'OUT - Vuelo:', _vuelo(vuelo), '',
         f'Salida {_ciudad(sal_cod)}:', _hh(sal_hora), ''] if sal_cod \
        else ['Box Lunch', '', '', 'Salida p/aeropuerto', _hh(bus_hora), '',
              'OUT - Vuelo:', _vuelo(vuelo), '',
              'Salida:', _hh(sal_hora), '']
    if not sal_cod:
        pass
    L += [f'Llegada {_ciudad(leg_cod)}:', _hh(leg_hora)] if leg_cod \
        else ['Llegada:', _hh(leg_hora)]
    if fech_leg:
        L.append(f'({fech_leg})')
    if refri:
        L += ['', '+', 'Refrigerio', _txt(refri)]
    L += ['', '', '', veh]
    return '\n'.join(L)


# ---------------------------------------------------------------------------
# DICCIONARIO INTERNO POR ATRACCION
# ---------------------------------------------------------------------------
# Horas tipicas de salida estimadas solo cuando el generador no trae el dato.
HORAS_DIC = {
    'Beto Carrero': '09:00',
    'Teleferico': '09:00',
    'Youhooo': '09:00',
    'Cascata Carolina': '10:00',
    'Zacarias': '10:00',
    'Barco': '09:00',
    'Panam Sport': '14:00',
    'Multiparque': '10:00',
    'Rueda': '14:00',
    'Panam Day': '09:00',
}

# Perfil por atraccion (estructura del dia como en la planilla modelo).
#   'sub': sub-atracciones de la manana/antes del almuerzo
#   'almuerzo': lista de lineas (None -> 'Almuerzo')
#   'cena': lista de lineas (None -> 'Cena hotel')
#   'salida': prefijo de la linea de salida (None -> 'Salida')
#   'noche_extra': True si la cena se escribe dentro del bloque de atraccion
PERFILES = {
    'Teleferico': dict(sub=['Youhooo '],
                       almuerzo=['Almuerzo Rest.', 'Pedra da Baleia'],
                       cena=['Teleférico', 'Cena hotel']),
    'Cascata Carolina': dict(sub=['Cascata ', 'Carolina'],
                             frases=['(Pool Parthy com espuma)'],
                             almuerzo=['Almuerzo en', 'parque'],
                             cena=['Cena hotel']),
    'Zacarias': dict(sub=['Pq. Zacarias ', 'hasta a las '],
                     almuerzo=['Almuerzo en', 'parque'],
                     cena=['Cena hotel']),
    'Beto Carrero': dict(sub=['Beto ', 'Carrero '],
                         almuerzo=['Almuerzo', '+', '+'],
                         cena=['Cena hotel']),
    'Rueda': dict(sub=['Rueda Gigante '], almuerzo=['Almuerzo', 'hotel '],
                  cena=['Cena hotel']),
    'Panam Day': dict(sub=['Mariscal'], almuerzo=['Almuerzo'],
                      cena=['Cena hotel']),
    'Barco': dict(sub=['Barco'], almuerzo=['Almuerzo'], cena=['Cena hotel']),
    'Panam Sport': dict(sub=['Panam Sport'], almuerzo=['Almuerzo'],
                        cena=['Cena hotel']),
    'Multiparque': dict(sub=['Multiparque'], almuerzo=['Almuerzo'],
                        cena=['Cena hotel']),
    'Campamento Americano': dict(sub=['Campamento Americano'],
                                 almuerzo=['Almuerzo'], cena=['Cena hotel']),
}


def _perfil(atr):
    return PERFILES.get(str(atr or '').strip())


def _noche(noche):
    """Convierte una noche ('Eclipse (Espuma)') a sus lineas.
    'Cristo Luz'/'Cristo Luz con Cena' -> Salida/Visita (sin 5ª COMIDA ni BUS)."""
    nn = (noche or '').strip()
    if not nn or nn.lower() == 'noche libre':
        return None
    if 'cristo luz' in nn.lower():
        return ['Salida:', '____H', 'Visita', 'CRISTO LUZ']
    m = re.match(r'^(.*?)\s*\((.*?)\)\s*$', nn)
    if m:
        return ['+', m.group(1).strip().upper(), '(' + m.group(2).strip() + ')',
                '', '+', '5ª COMIDA']
    return ['+', nn.upper(), '', '+', '5ª COMIDA']


def texto_dia(atr, hora=None, sub=None, almuerzo=None, cena=None,
              noche=None, visita=None, veh='BUS'):
    """Celda de un dia normal. El modelo (template) es solo referencia de
    estructura; el contenido va con la ATRACCION PRINCIPAL que viene de la
    marcacion. Los sub-textos (Youhooo, restaurantes, etc.) solo se
    agregan si vienen dados (sub/almuerzo/cena); si no, no se inventan.
    hora es un dato de la marcacion; no se inventa (HORAS_DIC queda solo de
    referencia). noche: disco del dia. visita: Cristo Luz (sin 5ª COMIDA)."""
    atr_str = (atr or '').strip()
    visit_lines = None
    if visita:
        vv = visita if isinstance(visita, dict) else {'nombre': visita}
        visit_lines = [' Salida:', _hh(vv.get('hora')),
                       'Visita', _txt(vv.get('nombre')).upper()]

    # -- dia libre ---------------------------------------------------------
    if not atr_str or 'libre' in atr_str.lower():
        base = ['Mañana libre', '+', 'Almuerzo']
        if sub:
            for s in sub:
                base += ['+', str(s)]
        base += ['+', 'Cena hotel']
        nbl = _noche(noche)
        if nbl:
            base.append('')
            base += nbl
            if nbl[0] != '+':          # Cristo Luz: visita nocturna, sin BUS
                return '\n'.join(base)
        base += ['', '', veh]
        return '\n'.join(base)

    # -- dia normal (solo atraccion principal de la marcacion) -------------
    L = [f'Salida p/{atr_str}', f'as {_hh(hora)}']
    if sub:
        for s in sub:
            L += ['+', str(s).upper()]
    L += ['+', 'Almuerzo', '', '+', 'Cena hotel']
    nbl = _noche(noche)
    if nbl:
        L.append('')
        L += nbl
        if nbl[0] != '+':              # Cristo Luz: visita nocturna, sin BUS
            return '\n'.join(L)
        L += ['', '', veh]
    elif visit_lines:
        L += ['', '']
        L += visit_lines
    else:
        L += ['', '', '', veh]
    return '\n'.join(L)


TEXTOS_AUX = {
    'almuerzo_en_parque': ['Almuerzo en', 'parque'],
    'teleferico': ['Salida p/Teleférico ', 'Youhooo '],
}