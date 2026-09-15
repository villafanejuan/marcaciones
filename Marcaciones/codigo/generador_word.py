# -*- coding: utf-8 -*-
"""
Generador de planillas Excel a partir de un itinerario en Word (.docx), del
tipo que la empresa le entrega a sus pasajeros.

El Word es texto explicativo: se extrae todo lo aprovechable por dia
    - la escuela (nombre) y la cantidad de personas,
    - las cabeceras de dia (ej. 'Dom. 11 de Oct. :') y el titulo de actividad,
    - los vuelos (aerolinea + numero) y los horarios de salida/arribo,
    - desayuno / almuerzo / cena con su restaurante o lugar,
    - las atracciones diurnas (mapeadas a los nombres estandar),
    - las noches de discoteca.

Todo el texto puramente descriptivo se ignora. Produce la MISMA salida que
generador_receptivo.py (hojas por atraccion, 5 COMIDA, PROGRAMACION y
COLEGIOS MAESTRO) para que despues el usuario ajuste a mano los campos que
el Word no trae.

Uso:
    python generador_word.py [archivo.docx] [archivo_salida.xlsx]
"""

import os
import re
import sys
import zipfile
import html as _html
import glob
import unicodedata
from datetime import datetime

import openpyxl

import generador_receptivo as base
import formato_modelo as modelo

if getattr(sys, "frozen", False):
    RAIZ = os.path.dirname(sys.executable)
    BASE = RAIZ
else:
    BASE = os.path.dirname(os.path.abspath(__file__))
    RAIZ = os.path.abspath(os.path.join(BASE, ".."))
ENTRADAS = os.path.join(RAIZ, "entradas")
PLANTILLAS = os.path.join(RAIZ, "plantillas")
SALIDAS = os.path.join(RAIZ, "salidas")
for _dir in (ENTRADAS, PLANTILLAS, SALIDAS):
    os.makedirs(_dir, exist_ok=True)
PLANTILLA = os.path.join(PLANTILLAS, "cristour_template.xlsx")

# ---------------------------------------------------------------------------
# Mapeos (mismos nombres estandar que los otros generadores)
# ---------------------------------------------------------------------------
MAPEO_ATRACCION = {
    'cristo luz':       'Cristo Luz',
    'parque unipraias': 'Teleferico',
    'unipraias':        'Teleferico',
    'teleferico':       'Teleferico',
    'beto carrero':     'Beto Carrero',
    'barco pirata':     'Barco',
    'porto belo':       'Barco',
    'campamento americano': 'Campamento Americano',
    'zacarias':         'Zacarias',
    'zacarías':         'Zacarias',
    'multiparque':      'Multiparque',
    'parque acuatico':  'Multiparque',
    'panam sport':      'Panam Sport',
    'rueda':            'Rueda',
    'cascata':          'Cascata Carolina',
}

MAPEO_NOCHE = {
    'fiesta eclipse': 'Eclipse (Espuma)',
    'eclipse':        'Eclipse (Espuma)',
    'espuma':         'Eclipse (Espuma)',
    'green valley':   'Green Valley (Blanco)',
    'duff':           "D'uff (Negro)",
    'disco 2':        'Disco 2 (Azul)',
    'fiesta privada': 'Fiesta Privada',
    'shed':           'Shed (Disfraces)',
    'cristo luz con cena': 'Cristo Luz con Cena',
}

MESES = {
    'ene': 1, 'jan': 1, 'janeiro': 1, 'enero': 1,
    'feb': 2, 'fev': 2, 'febrero': 2, 'fevereiro': 2,
    'mar': 3, 'marzo': 3, 'marco': 3,
    'abr': 4, 'abril': 4,
    'may': 5, 'maio': 5, 'mayo': 5,
    'jun': 6, 'junho': 6, 'junio': 6,
    'jul': 7, 'julho': 7, 'julio': 7,
    'ago': 8, 'agosto': 8,
    'sep': 9, 'set': 9, 'sept': 9, 'setembro': 9, 'septiembre': 9,
    'oct': 10, 'out': 10, 'outubro': 10, 'octubre': 10,
    'nov': 11, 'novembro': 11, 'noviembre': 11,
    'dic': 12, 'dez': 12, 'diciembre': 12, 'dezembro': 12,
}

RE_CABECERA_DIA = re.compile(
    r'^\s*(Lun|Mar|Mi[eé]|Jue|Vie|S[aá]b|Dom)\.?\s+(\d{1,2})\s+de\s+([A-Za-záéíóúçãõâêôü.]+)',
    re.IGNORECASE
)
RE_HORA = re.compile(r'(\d{1,2}):(\d{2})\s*(hrs?\.?|hr|Hr|HR)')
RE_NUM_VUELO = re.compile(r'\b(H2\s*\d{4}|LA\s*\d{4}|JJ\s*\d{4}|G3\s*\d{4})\b', re.I)
RE_AERO = re.compile(r'\b(SKY\s*Airlines|LATAM|Azul Linhas|Aereas|GOL\b)', re.I)


def normalizar(t):
    """Minusculas sin acentos (para buscar keywords de forma segura)."""
    s = ''.join(c for c in unicodedata.normalize('NFD', str(t))
                if unicodedata.category(c) != 'Mn')
    return s.lower()


# ---------------------------------------------------------------------------
# Extraccion del .docx
# ---------------------------------------------------------------------------
def extraer_texto_docx(path):
    """Devuelve las lineas de texto del documento Word (sin marcas, quien
    conserva los sectores de parrafo y tabs)."""
    with zipfile.ZipFile(path) as z:
        xml = z.read('word/document.xml').decode('utf-8', errors='replace')
    xml = xml.replace('</w:p>', '\n')
    xml = re.sub(r'<w:tab[^>]*/>', '\t', xml)
    xml = re.sub(r'<[^>]+>', '', xml)
    texto = _html.unescape(xml)
    lineas = []
    for l in texto.split('\n'):
        l = l.strip()
        if l:
            lineas.append(l)
    return lineas


def anio_desde_nombre(path):
    m = re.search(r'(20\d{2})', os.path.basename(path))
    return int(m.group(1)) if m else 2026


def mes_desde_texto(t):
    clave = re.sub(r'[^A-Za-z]', '', t).lower()
    return MESES.get(clave)


def limpiar_hora(linea):
    """Extrae todas las horas mencionadas en una linea (HH:MM),
    respetando am/pm."""
    horas = []
    for m in re.finditer(r'\b(\d{1,2}):(\d{2})\b\s*(am|pm)?\b', linea, re.I):
        h = int(m.group(1))
        suf = m.group(3)
        if suf and suf.lower() == 'pm' and h < 12:
            h += 12
        elif suf and suf.lower() == 'am' and h == 12:
            h = 0
        horas.append(f"{h:02d}:{m.group(2)}")
    return horas


def separar_hora(linea):
    """Devuelve (hora primera o None, texto de la linea sin la hora),
    respetando am/pm."""
    m = re.search(r'\b(\d{1,2}):(\d{2})\b\s*(am|pm)?\b', linea, re.I)
    if m:
        h = int(m.group(1))
        suf = m.group(3)
        if suf and suf.lower() == 'pm' and h < 12:
            h += 12
        elif suf and suf.lower() == 'am' and h == 12:
            h = 0
        hora = f"{h:02d}:{m.group(2)}"
        resto = linea[m.end():].strip()
        resto = re.sub(r'^\s*(hrs?\.?|hr|Hr|HR|pm|am)\b', '', resto, flags=re.I).strip(' .,:')
        return hora, resto
    return None, linea.strip(' .,:')


def parsear_itinerario(lineas, anio):
    """Devuelve (colegio_info, bloques_de_dias).

    colegio_info: dict con nombre, lineas de cantidad de personas.
    bloques: lista de dicts {'fecha', 'titulo', 'lineas': [...]}
    """
    nombre_escuela = None
    cantidades = None
    dias = []
    actual = None

    for l in lineas:
        m = RE_CABECERA_DIA.match(l)
        if m:
            dia, mes_txt = int(m.group(2)), m.group(3)
            mes = mes_desde_texto(mes_txt)
            if mes is None:
                continue
            try:
                fecha = datetime(anio, mes, dia)
            except ValueError:
                continue
            resto = l[m.end():].strip(' \t:')
            actual = {'fecha': fecha, 'titulo': resto or None, 'lineas': []}
            if actual['titulo'] and len(actual['titulo']) > 80:
                actual['titulo'] = actual['titulo'][:80]
            dias.append(actual)
            continue
        if actual is None:
            # antes del primer dia: solo el nombre y las cantidades
            if nombre_escuela is None and l.isupper() and not re.search(r'\d', l[:6]):
                nombre_escuela = l.strip()
            elif re.search(r'\d', l) and re.search(r'[MH]|PADRES|ADULTOS|TOTAL|=\s*\d', l, re.I):
                cantidades = l.strip()
            continue
        actual['lineas'].append(l)

    return {'nombre': nombre_escuela, 'cantidades': cantidades}, dias


def attract(en_lapso):
    """Detecta las atracciones diurnas de un bloque de dia (nombre estandar)."""
    texto = normalizar(' '.join(en_lapso))
    encontradas = []
    for clave, estandar in MAPEO_ATRACCION.items():
        if clave in texto and estandar not in encontradas:
            encontradas.append(estandar)
    return encontradas


def noche_del_dia(lineas):
    """Detecta la noche (discoteca / atraccion nocturna) de un bloque."""
    texto = normalizar(' '.join(lineas))
    mejor = None
    for clave, estandar in MAPEO_NOCHE.items():
        if clave in texto:
            if mejor is None or len(clave) > len(mejor[0]):
                mejor = (clave, estandar)
    if mejor:
        return mejor[1]
    if 'discoteca' in texto or 'disco' in texto or 'fiesta' in texto:
        return 'DISCOTECA'
    return None


def horario_atraccion(en_lapso, nombre_atraccion):
    """Busca la primera hora de la linea donde aparece la atraccion."""
    clave = next((k for k in MAPEO_ATRACCION if MAPEO_ATRACCION[k] == nombre_atraccion), None)
    if not clave:
        return None
    for l in en_lapso:
        if clave in normalizar(l):
            horas = limpiar_hora(l)
            if horas:
                return horas[0]
    return None


# ---------------------------------------------------------------------------
# Analisis de detalle por dia (vuelos, comidas, salidas, retornos, noches)
# ---------------------------------------------------------------------------
def extraer_vuelo(lista_textos):
    """Nombre corto del vuelo: 'SKY AIRLINES H2 0809' o None."""
    t = ' '.join(lista_textos or [])
    partes = []
    m = RE_NUM_VUELO.search(t)
    if m:
        partes.append(m.group(1).replace('  ', ' ').upper())
    m = RE_AERO.search(t)
    if m:
        partes.append(m.group(1).title())
    if partes:
        return ' '.join(partes)
    # si hay 'vuelo' pero no aerolinea/numero, devolver texto corto
    return None


def extraer_hotel(lineas):
    """Busca el nombre del hotel (ej. 'D SINTRA'), ignorando textos como
    'hotel ubicado en...' (empiezan en minuscula)."""
    pat = (r'hotel\s*[^\w\s]+[^\w]*\s*'
           r'([A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúüñ]*(?:\s+[A-Za-zÁÉÍÓÚÑáéíóúüñ]+){0,3})')
    for l in lineas:
        m = re.search(pat, l)
        if m:
            candidato = re.sub(r'[\uFFFD\x00-\x1f\x7f]', '', m.group(1)).strip()
            if 2 <= len(candidato) <= 40:
                return candidato
    return None


def analizar_bloque(lineas):
    """Clasifica las lineas de un dia en vuelos, comidas, salidas, etc.

    Devuelve un dict con listas de {'hora', 'texto'} para cada categoria.
    """
    res = {'titulo': None, 'vuelos': [], 'salida': [], 'desayuno': [],
           'almuerzo': [], 'cena': [], 'retorno': [], 'alojamiento': [],
           'info': [], 'noche_item': None}
    for l in lineas:
        hora, resto = separar_hora(l)
        if not resto:
            continue
        rn = normalizar(resto)
        if any(k in rn for k in ('vuelo', 'sky airlines', 'latam', 'salida en vuelo',
                                 'salida hacia', 'arribo a', 'arribo al', 'llegada')):
            if 'aeropuerto' in rn or 'vuelo' in rn or 'arribo' in rn or \
                    any(a in rn for a in ('salida hacia', 'salida en vuelo', 'llegada')):
                res['vuelos'].append({'hora': hora, 'texto': resto})
                continue
        if 'desayuno' in rn:
            res['desayuno'].append({'hora': hora, 'texto': resto})
            continue
        if 'almuerzo' in rn and 'buffet' in rn or 'almuerzo' in rn:
            res['almuerzo'].append({'hora': hora, 'texto': resto})
            continue
        if 'cena' in rn:
            res['cena'].append({'hora': hora, 'texto': resto})
            continue
        if any(k in rn for k in ('discoteca', 'fiesta', 'noche de disco')):
            res['noche_item'] = {'hora': hora, 'texto': resto}
            continue
        if any(k in rn for k in ('alojamiento', 'check-in', 'distribucion de habitaciones')):
            res['alojamiento'].append({'hora': hora, 'texto': resto})
            continue
        if any(k in rn for k in ('retornar', 'regreso al hotel', 'arribo al hotel',
                                 'dejando el sector', 'pernocte')):
            res['retorno'].append({'hora': hora, 'texto': resto})
            continue
        if any(k in rn for k in ('reuni', 'lobby', 'abordar', 'partiremos',
                                 'salida', 'trasladaremos', 'trasladara')):
            res['salida'].append({'hora': hora, 'texto': resto})
            continue
        res['info'].append({'hora': hora, 'texto': resto})
    return res


def lineas_completas(b):
    """Todas las lineas de un bloque incluido el titulo de la cabecera."""
    return ([b.get('titulo')] if b.get('titulo') else []) + b['lineas']


def construir_itinerario(bloques):
    """Lista de pasos {dia, fecha, tipo, titulo, atracciones, noche, horarios, ana}."""
    if not bloques:
        return []
    idx_out = indice_dia_regreso(bloques)
    bloques, idx_out = bloques_sin_llegada_madrugada(bloques, idx_out)
    pasos = []
    for i, b in enumerate(bloques):
        lc = lineas_completas(b)
        ana = analizar_bloque(lc)
        paso = {
            'dia': i + 1,
            'fecha': b['fecha'],
            'titulo': b.get('titulo'),
            'tipo': 'IN' if i == 0 else ('OUT' if i == idx_out else 'DIA'),
            'atracciones': [],
            'noche': None,
            'horarios': {},
            'ana': ana,
        }
        if paso['tipo'] == 'DIA':
            paso['atracciones'] = attract(lc)
            for a in paso['atracciones']:
                paso['horarios'][a] = horario_atraccion(lc, a)
            paso['noche'] = noche_del_dia(lc)
        pasos.append(paso)
    return pasos


def indice_dia_regreso(bloques):
    """Busca el indice del dia de regreso (OUT).

    Pistas fuertes: 'local del colegio', 'traslado hacia el', 'regreso',
    'retorno' o paso por aeropuerto con salida/vuelo. Si no se encuentra,
    el ultimo dia.
    """
    cand = None
    for i, b in enumerate(bloques):
        texto = normalizar(' '.join(b['lineas']))
        if any(p in texto for p in ('local del colegio', 'traslado hacia el',
                                    'regreso', 'retorno')):
            cand = i
        elif 'aeropuerto' in texto and ('salida' in texto or 'vuelo' in texto):
            cand = i
    return cand if cand is not None else len(bloques) - 1


def bloques_sin_llegada_madrugada(bloques, idx_out):
    """Borra los dias posteriores al OUT que solo son la llegada de la
    madrugada (todas sus horas antes de las 06:00 -> continuacion del viaje)."""
    if idx_out + 1 >= len(bloques):
        return bloques[:idx_out + 1], idx_out
    ultimo_util = idx_out
    for i in range(idx_out + 1, len(bloques)):
        horas = []
        for l in bloques[i]['lineas']:
            horas += limpiar_hora(l)
        if horas and min(int(h.split(':')[0]) for h in horas) >= 6:
            ultimo_util = i  # tiene contenido real en horario de dia
        else:
            break
    return bloques[:ultimo_util + 1], idx_out


def parsear_cantidades(texto):
    """Extrae alumnos (M+H), adultos (PADRES/ADULTOS...) y total de la linea
    de cantidades. Si no se encuentra, devuelve todo None."""
    if not texto:
        return None, None, None
    t = texto.upper()
    m = re.search(r'(\d+)\s*M\s*\+\s*(\d+)\s*H', t)
    alunos = (int(m.group(1)) + int(m.group(2))) if m else None
    a = re.search(r'\+\s*(\d+)\s*(PADRES|ADULTOS|ACOMP|MAYORES|ADT)', t)
    adultos = int(a.group(1)) if a else None
    total = None
    mt = re.search(r'=\s*(\d+)', t)
    if mt:
        total = int(mt.group(1))
    if total is None and alunos is not None:
        total = alunos + (adultos or 0)
    if alunos is None and total is not None:
        # formato "13 + 01 ADT.H  = 14"  -> 13 alumnos, 1 adulto, 14 total
        pa = re.search(r'(\d+)\s*\+\s*(\d+)\s*ADT', t)
        if pa:
            alunos, adultos = int(pa.group(1)), int(pa.group(2))
    return alunos, adultos, total


def construir_colegio(info, bloques, anio):
    """Arma un colegio con la misma estructura que usa generador_receptivo."""
    it = construir_itinerario(bloques)
    alunos, adultos, total = parsear_cantidades(info.get('cantidades'))
    if not it:
        salida = llegada = None
    else:
        salida = it[0]['fecha']
        llegada = it[-1]['fecha']
    noche_count = sum(1 for s in it if s.get('noche'))

    todo_lineas = [l for b in bloques for l in b['lineas']]
    hotel = extraer_hotel(todo_lineas)

    vuelos_txt = []
    for s in it:
        ana = s['ana']
        if ana.get('vuelos'):
            vuelos_txt += [v['texto'] for v in ana['vuelos']]
    aerea = extraer_vuelo(vuelos_txt) or None

    colg = {
        'num': 1,
        'file': None,
        'colegio': info.get('nombre') or 'SIN NOMBRE',
        'curso': '',
        'alunos': alunos,
        'adultos': adultos,
        'guias': None,
        'total': total,
        'hotel': hotel,
        'salida': salida,
        'llegada': llegada,
        'cant_discos': noche_count,
        'aerea': aerea,
        'obs': None,
        '_it': it,
    }
    return colg


# ---------------------------------------------------------------------------
# Distribucion (misma estructura de eventos que receptivo, + horario inicial)
# ---------------------------------------------------------------------------
def distribuir_word(colg):
    atracciones = {}
    eventos_5comida = []
    for step in colg['_it']:
        if step['tipo'] in ('IN', 'OUT'):
            continue
        for atr in step['atracciones']:
            base_ev = {
                'fecha': step['fecha'],
                'num': colg['num'],
                'colegio': colg['colegio'],
                'curso': colg['curso'],
                'alunos': colg['alunos'],
                'adultos': colg['adultos'],
                'guias': colg['guias'],
                'total': colg['total'],
                'hotel': colg['hotel'],
                'cant_discos': colg['cant_discos'],
                'obs': colg['obs'],
                'horario': step.get('horarios', {}).get(atr),
            }
            atracciones.setdefault(atr, []).append(base_ev)
        noche = step.get('noche')
        if not noche:
            continue
        if noche.lower().strip() in base.ATRACCIONES_NOCTURNAS:
            base_ev = {
                'fecha': step['fecha'],
                'num': colg['num'],
                'colegio': colg['colegio'],
                'curso': colg['curso'],
                'alunos': colg['alunos'],
                'adultos': colg['adultos'],
                'guias': colg['guias'],
                'total': colg['total'],
                'hotel': colg['hotel'],
                'cant_discos': colg['cant_discos'],
                'obs': colg['obs'],
                'horario': None,
            }
            atracciones.setdefault('Cristo Luz', []).append(base_ev)
            continue
        if noche not in ('Noche Libre', 'Noche libre'):
            evento = {
                'fecha': step['fecha'],
                'num': colg['num'],
                'colegio': colg['colegio'],
                'curso': colg['curso'],
                'alunos': colg['alunos'],
                'adultos': colg['adultos'],
                'guias': colg['guias'],
                'total': colg['total'],
                'hotel': colg['hotel'],
                'cant_discos': colg['cant_discos'],
                'obs': colg['obs'],
                'boliche': noche,
            }
            eventos_5comida.append(evento)
    return atracciones, eventos_5comida


# ---------------------------------------------------------------------------
# Textos de PROGRAMACION (nutridos con la informacion real del Word)
# ---------------------------------------------------------------------------
def _hh(h):
    return (h + 'H') if h else '____H'


def _cortar(t, n=26):
    """Recorta una frase a n chars sin cortar palabras."""
    t = re.sub(r'\s+', ' ', (t or '').strip())
    if len(t) <= n:
        return t
    i = t[:n].rfind(' ')
    return (t[:i] + '...') if i > 0 else t[:n] + '...'


def _quitar_prefijo(texto, *prefijos):
    t = re.sub(r'\s+', ' ', (texto or '').strip())
    for p in prefijos:
        t2 = re.sub('^' + re.escape(p) + r'\s*', '', t, flags=re.I)
        if t2 != t:
            t = t2
            break
    return t.strip()


def _lugar(texto):
    """Extrae la frase completa del restaurante/lugar de una comida.
    Ej: 'ALMUERZO buffet en DON ALBERTO' -> 'en DON ALBERTO'."""
    t = _quitar_prefijo(texto, 'ALMUERZO', 'CENA', 'DESAYUNO')
    t = re.sub(r'\s+', ' ', t).strip(' .')
    t = re.sub(r'^buffet\s+', '', t, re.I)
    nt = normalizar(t)
    if not t or 'comedor del hotel' in nt or nt in ('hotel', 'el hotel'):
        return None
    return t


_CIUDADES = [('iquique', 'IQQ'), ('florianopolis', 'FLN'),
             ('santiago', 'SANTIAGO'), ('lima', 'LIMA'),
             ('camboriu', 'CAMBORIU')]


def _ciudad_de(texto):
    n = normalizar(texto)
    for nombre, codigo in _CIUDADES:
        if nombre in n:
            return codigo
    return None


def _noche_bloque(noche, noche_item):
    nn = normalizar(noche)
    if 'noche libre' in nn:
        return []
    hora = (noche_item or {}).get('hora')
    txt = normalizar((noche_item or {}).get('texto') or '')
    L = ['+']
    if 'cristo luz' in nn:
        L += ['Visita', 'CRISTO LUZ']
    else:
        m = re.match(r'^(.*?)\s*\((.*?)\)\s*$', noche)
        if m:
            L += [m.group(1).strip().upper(),
                  '(' + m.group(2).strip() + ')']
        elif 'discoteca' in nn and 'confirmar' in txt:
            L += ['DISCOTECA', '(d\xeda por confirmar)']
        else:
            L.append(noche.strip().upper())
    if 'cristo luz' in nn and hora:
        L.append(_hh(hora))
    return L


def texto_dia_prog(step, sig=None):
    """Texto de celda del mapa de dias, mismo formato que el archivo modelo
    del usuario: bloques cortos separados por '+', horas 'HH:MMH'."""
    ana = step.get('ana') or {}
    vuelos = ana.get('vuelos') or []
    salida = ana.get('salida') or []
    desayuno = ana.get('desayuno') or []
    almuerzo = ana.get('almuerzo') or []
    cena = ana.get('cena') or []
    retorno = ana.get('retorno') or []
    aloj = ana.get('alojamiento') or []
    noche_item = ana.get('noche_item') or {}
    noche = step.get('noche')

    # ---------------- IN ----------------
    if step['tipo'] == 'IN':
        vuelo = extraer_vuelo([v['texto'] for v in vuelos]) or None
        hs = next((v['hora'] for v in vuelos if v['hora']), None)
        c_ids = next((_ciudad_de(v['texto']) for v in vuelos), None)
        hl = None
        c_lleg = None
        if sig:
            sv = (sig.get('ana') or {}).get('vuelos') or []
            arribos = [(v['hora'], _ciudad_de(v['texto']))
                       for v in sv if v['hora']
                       and 'arribo' in normalizar(v['texto'])]
            if arribos:
                hl, c_lleg = arribos[-1]
        return modelo.texto_in(vuelo, c_ids, hs,
                               c_lleg or _ciudad_de('florianopolis'), hl)

    # ---------------- OUT ----------------
    if step['tipo'] == 'OUT':
        ap = next((it for it in salida if 'aeropuerto'
                   in normalizar(it['texto'])), None)
        vuelo = extraer_vuelo([v['texto'] for v in vuelos]) or '____'
        L = ['Box Lunch']
        L += ['', '', 'Salida p/aeropuerto',
              _hh(ap['hora'] if ap else None)]
        L += ['', 'OUT - Vuelo: ', ' ' + vuelo, '']
        tramos = [v for v in vuelos if v['hora']]
        for i, v in enumerate(tramos):
            rn = normalizar(v['texto'])
            acc = 'Salida' if 'salida' in rn else (
                'Llegada' if 'arribo' in rn else 'Vuelo')
            ci = _ciudad_de(v['texto'])
            L.append((acc + ' ' + ci if ci else acc) + ':')
            L.append(_hh(v['hora']))
            if i < len(tramos) - 1:
                L.append('+')
        for it in ana.get('info') or []:
            if 'refrigerio' in normalizar(it['texto']):
                L.append('+')
                L.append('Refrigerio')
                L.append(_cortar(_quitar_prefijo(it['texto'], 'REFRIGERIO'),
                                 24))
                if it['hora']:
                    L.append(_hh(it['hora']))
        L += ['', '', 'BUS']
        return '\n'.join(L)

    # ---------------- DIA NORMAL ----------------
    atrs = step['atracciones']
    titulo = (step.get('titulo') or '').strip()
    L = []

    hs = next((it['hora'] for it in salida if it['hora']), None)
    if hs:
        L += ['Salida', 'as ' + _hh(hs)]
    else:
        L.append(_cortar(titulo or 'DÍA LIBRE', 40).upper())

    if atrs:
        L.append('+')
        for a in atrs:
            L.append(a.upper())
        for a in atrs:
            hx = step.get('horarios', {}).get(a)
            if hx:
                L.append(_hh(hx) + ' ' + a.upper())
        ht = next((v['hora'] for v in vuelos if v['hora']
                   and 'arribo al hotel' in normalizar(v['texto'])), None)
        if ht:
            L.append(_hh(ht) + ' Arribo al hotel')
    elif any(it.get('hora') for it in retorno):
        L.append('+')
        r = next(it for it in retorno if it.get('hora'))
        L.append(_hh(r['hora']) + ' ' + _cortar(r['texto'], 22))

    if almuerzo:
        L += ['+', 'Almuerzo']
        lug = _lugar(almuerzo[0]['texto'])
        if lug:
            L.append(lug)
    else:
        L += ['+', 'Almuerzo']

    if cena:
        L += ['+', 'Cena hotel']
        lug = _lugar(cena[0]['texto'])
        if lug:
            L.append(lug)
    else:
        L += ['+', 'Cena hotel']

    if noche:
        L.append('')
        L += _noche_bloque(noche, noche_item)
        L += ['+', '5ª COMIDA']
    else:
        # eventos nocturnos no-disco (ej. "Noche bajo las estrellas en la playa")
        eventos_noche = [
            it for it in ana.get('info') or []
            if 'noche' in normalizar(it['texto'])
        ]
        for it in eventos_noche:
            L.append('')
            L.append('+')
            L.append(_cortar(it['texto'].strip(), 30).upper())
            if it['hora']:
                L.append(_hh(it['hora']))

    L += ['', '', 'BUS']
    return '\n'.join(L)


def texto_bloque_col(word_colg):
    return base.texto_bloque_colegio(word_colg)


def rellenar_programacion_word(ws, P, colg):
    n = 1
    base.quitar_merges_excepto(ws, {"A1:J1"})
    base.limpiar_valores(ws, 3)

    headers = base.cabeceras_semana(colg['salida'])
    for i, texto in enumerate(headers[:8]):
        if texto:
            ws.cell(row=2, column=3 + i).value = texto

    ws.cell(row=3, column=1).value = colg['num']
    ws.cell(row=3, column=2).value = texto_bloque_col(colg)
    texto_max = 1
    pasos = colg['_it']
    for idx, step in enumerate(pasos):
        dia = step['dia']
        if dia > 8:
            continue
        sig = pasos[idx + 1] if idx + 1 < len(pasos) else None
        t = texto_dia_prog(step, sig=sig)
        ws.cell(row=3, column=2 + dia).value = t
        texto_max = max(texto_max, t.count('\n') + 1)

    for r in range(3, n + 3):
        cel = ws.cell(row=r, column=2)
        if not cel.has_style:
            cel._style = __import__('copy').copy(P.p_model[1]._style)
        cel.fill = openpyxl.styles.PatternFill(
            fill_type="solid", start_color=base.COLORES_FILA_B[(r - 3) % 6],
            end_color=base.COLORES_FILA_B[(r - 3) % 6])

    # altura de fila igual que el modelo del usuario (fixed)
    altura_base = ws.row_dimensions[3].height or 277.2
    ws.row_dimensions[3].height = altura_base


# ---------------------------------------------------------------------------
# Atraccion: igual que el receptivo pero con HORARIO en la columna I
# ---------------------------------------------------------------------------
def rellenar_atraccion_word(ws, P, nombre, eventos, titulo_sufijo):
    base.quitar_merges_excepto(ws, {"B1:K1", "B2:C2"})
    base.limpiar_valores(ws, 3)
    ws["B1"] = f"{nombre.upper()} - {titulo_sufijo}"

    fila = 3
    grupos = base.agrupar_por_fecha(eventos)
    for i, (fecha, grupo) in enumerate(sorted(grupos.items(), key=lambda kv: kv[0])):
        r0, r1 = fila, fila + len(grupo) - 1
        for j, ev in enumerate(grupo):
            r = fila + j
            base.estilar_celdas(ws, P.a_data, "BCDEFGHIJK", r)
            ws.cell(row=r, column=2).value = ev["num"]           # B Nº
            ws.cell(row=r, column=3).value = base.texto_nombre(ev)  # C colegio+curso
            ws.cell(row=r, column=4).value = ev["alunos"]        # D ALUNOS
            ws.cell(row=r, column=5).value = ev.get("adultos")   # E ADULTOS
            ws.cell(row=r, column=6).value = ev["total"]         # F TOTAL
            ws.cell(row=r, column=7).value = ev["guias"]         # G GUIA
            ws.cell(row=r, column=8).value = ev["hotel"]         # H HOTEL
            ws.cell(row=r, column=9).value = ev.get("horario")   # I HORARIO (hora del Word)
            ws.cell(row=r, column=10).value = fecha              # J DATA
            ws.cell(row=r, column=10).number_format = P.fmt_fecha
        if r1 == r0:
            ws.cell(row=r0, column=11).value = f"=F{r0}+G{r0}"
        else:
            ws.cell(row=r0, column=11).value = f"=SUM(F{r0}:F{r1})+SUM(G{r0}:G{r1})"
            ws.merge_cells(start_row=r0, start_column=10, end_row=r1, end_column=10)
            ws.merge_cells(start_row=r0, start_column=11, end_row=r1, end_column=11)
        color = base.COLORES_FILA_B[i % 6]
        for col in range(2, 12):
            for r in range(r0, r1 + 1):
                try:
                    ws.cell(row=r, column=col).fill = openpyxl.styles.PatternFill(
                        fill_type="solid", start_color=color, end_color=color)
                except AttributeError:
                    pass
        fila = r1 + 1


def rellenar_comida_word(ws, P, eventos, titulo_sufijo):
    """5 COMIDA: mismo formato que la plantilla (grupos por fecha + boliche)."""
    conservar = {"B1:J1", "B2:C3", "D2:D3", "E2:E3", "F2:F3", "G2:G3", "H2:H3", "I2:J3"}
    base.quitar_merges_excepto(ws, conservar)
    base.limpiar_valores(ws, 4)
    ws["B1"] = f"5ª COMIDA - {titulo_sufijo}"

    grupos = {}
    for ev in eventos:
        grupos.setdefault((ev["fecha"], ev["boliche"]), []).append(ev)
    orden = sorted(grupos.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[1][0]["num"]))

    fila = 4
    for i, ((fecha, boliche), grupo) in enumerate(orden):
        r0, r1 = fila, fila + len(grupo) - 1
        for j, ev in enumerate(grupo):
            r = fila + j
            base.estilar_celdas(ws, P.c_data, "BCDEFGHIJ", r)
            ws.cell(row=r, column=2).value = j + 1                          # B Nº
            ws.cell(row=r, column=3).value = base.texto_nombre(ev)          # C colegio
            ws.cell(row=r, column=4).value = ev["alunos"]                   # D ALUNOS
            ws.cell(row=r, column=5).value = ev.get("adultos")              # E ADULTOS
            ws.cell(row=r, column=6).value = f"=D{r}+E{r}"                  # F TOTAL
            ws.cell(row=r, column=7).value = ev["guias"]                    # G GUIA
            ws.cell(row=r, column=9).value = fecha                          # I DATA
            ws.cell(row=r, column=9).number_format = P.fmt_fecha
        ws.cell(row=r0, column=8).value = f"=SUM(F{r0}:F{r1})+SUM(G{r0}:G{r1})"
        ws.cell(row=r0, column=10).value = boliche
        if r1 > r0:
            ws.merge_cells(start_row=r0, start_column=8, end_row=r1, end_column=8)    # H
            ws.merge_cells(start_row=r0, start_column=9, end_row=r1, end_column=9)    # I
            ws.merge_cells(start_row=r0, start_column=10, end_row=r1, end_column=10)  # J
        color = base.COLORES_FILA_B[i % 6]
        for col in range(2, 11):
            for r in range(r0, r1 + 1):
                try:
                    ws.cell(row=r, column=col).fill = openpyxl.styles.PatternFill(
                        fill_type="solid", start_color=color, end_color=color)
                except AttributeError:
                    pass
        fila = r1 + 1


# ---------------------------------------------------------------------------
# Generacion
# ---------------------------------------------------------------------------
def generar(docx_path, plantilla_path, out_path=None, anio=None, titulo_sufijo=None):
    if anio is None:
        anio = anio_desde_nombre(docx_path)
    lineas = extraer_texto_docx(docx_path)
    if not lineas:
        raise ValueError(f"No se pudo leer texto del archivo: {docx_path}")

    info, bloques = parsear_itinerario(lineas, anio)
    print(f"informacion: {info}")
    print(f"dias detectados: {len(bloques)}")
    for b in bloques:
        lc = lineas_completas(b)
        print(f"  {b['fecha'].strftime('%d/%m/%Y')}  tit={b.get('titulo')!r}  "
              f"atr={attract(lc)}  noche={noche_del_dia(lc)}")

    colg = construir_colegio(info, bloques, anio)
    for s in colg['_it']:
        a = s['ana']
        print(f"  paso {s['dia']} {s['fecha'].strftime('%d/%m/%Y')} {s['tipo']} "
              f"atr={s['atracciones']} noche={s['noche']} "
              f"vuelos={len(a['vuelos'])} sal={len(a['salida'])} "
              f"desa={len(a['desayuno'])} alm={len(a['almuerzo'])} cena={len(a['cena'])}")

    P = base.Plantilla(plantilla_path)
    wb = P.wb

    for extra in ("Cristo Luz  (2)", "Cristo Luz  (3)"):
        if extra in wb.sheetnames:
            wb.remove(wb[extra])

    if titulo_sufijo is None:
        titulo_sufijo = f"{colg['colegio']} - {anio}/{anio + 1}"

    atracciones, eventos_5comida = distribuir_word(colg)
    for atr in sorted(atracciones.keys()):
        f = sorted({e["fecha"] for e in atracciones[atr]})
        print(f"atraccion: {atr:<16} visitas={len(atracciones[atr]):<4} "
              f"{[d.strftime('%d/%m') for d in f]}")
    print(f"5 COMIDA: {len(eventos_5comida)} eventos")

    # PROGRAMACION (una hoja)
    base_prog = wb["Programacion "]
    ws = wb.copy_worksheet(base_prog)
    ws.title = "PROGRAMACION " + colg['salida'].strftime("%d-%m")
    ws["A1"] = f"PLANILLA DE LOS COLEGIOS - 2026/2027 - {colg['colegio']}"
    rellenar_programacion_word(ws, P, colg)
    wb.remove(base_prog)

    # atracciones
    base_atr = wb["Cristo Luz "]
    if "Cristo Luz" not in atracciones:
        # la hoja base no se usa: se limpia para no dejar los datos de la plantilla
        rellenar_atraccion_word(base_atr, P, "Cristo Luz", [], titulo_sufijo)
    for atr in sorted(atracciones.keys()):
        if atr == "Cristo Luz":
            ws_a = base_atr
            ws_a.title = atr
        else:
            ws_a = wb.copy_worksheet(base_atr)
            ws_a.title = atr
        rellenar_atraccion_word(ws_a, P, atr, atracciones[atr], titulo_sufijo)

    # 5 COMIDA
    rellenar_comida_word(wb["5 COMIDA"], P, eventos_5comida, titulo_sufijo)

    # COLEGIOS MAESTRO
    base.hoja_colegios(wb, [colg])

    if out_path:
        wb.save(out_path)
        print("guardado:", out_path)
        return out_path
    return wb


def main():
    import sys
    args = [a for a in sys.argv[1:]]
    docx = args[0] if args else None
    if not docx:
        encontrados = glob.glob(os.path.join(ENTRADAS, "*.docx"))
        if not encontrados:
            print("No hay archivos .docx en la carpeta.")
            return
        docx = encontrados[0]
    out = args[1] if len(args) > 1 else os.path.join(
        SALIDAS, "SALIDA_WORD_" + os.path.splitext(os.path.basename(docx))[0][:30] + ".xlsx")
    generar(docx, PLANTILLA, out)


if __name__ == "__main__":
    main()