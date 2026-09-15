# -*- coding: utf-8 -*-
"""
Generador de planillas INKA - Lee del Control Operativo (Reservas) y genera
el Excel con el MISMO formato que 'inka grupo 1 e 2.xlsx'.

Fuentes de verdad:
  - Itinerario Ideal (filas 8-10 de Reservas): UNICA fuente de itinerario.
  - NO inventa boliches, horarios, nombres de hoja, ni nada que no este en la fuente.
"""
import openpyxl, re, os, sys
from datetime import datetime, timedelta
from copy import copy
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

# ----------------------------------------------------------------------
# CONFIGURACION
# ----------------------------------------------------------------------
if getattr(sys, "frozen", False):
    _RAIZ_ = os.path.dirname(sys.executable)
else:
    _BASE_ = os.path.dirname(os.path.abspath(__file__))
    _RAIZ_ = os.path.abspath(os.path.join(_BASE_, ".."))
ENTRADAS = os.path.join(_RAIZ_, "entradas")
PLANTILLAS = os.path.join(_RAIZ_, "plantillas")
SALIDAS = os.path.join(_RAIZ_, "salidas")
for _dir_ in (ENTRADAS, PLANTILLAS, SALIDAS):
    os.makedirs(_dir_, exist_ok=True)
ARCHIVO_SALIDA = os.path.join(SALIDAS, "SALIDA_OPERATIVO_2026.xlsx")
PLANTILLA = os.path.join(PLANTILLAS, "inka_template.xlsx")


def _buscar_entrada():
    """Localiza el Control Operativo en entradas/ (primer .xlsx con hoja Reservas)."""
    if os.path.exists(os.path.join(ENTRADAS, "Control Operativo.xlsx")):
        return os.path.join(ENTRADAS, "Control Operativo.xlsx")
    import glob
    for f in glob.glob(os.path.join(ENTRADAS, "*.xlsx")):
        try:
            if 'Reservas' in openpyxl.load_workbook(f, read_only=True).sheetnames:
                return f
        except Exception:
            continue
    return os.path.join(ENTRADAS, "Control Operativo.xlsx")


ARCHIVO_ENTRADA = _buscar_entrada()
HOJA_ENTRADA = 'Reservas'

# Mapeo de nombres en Itinerario Ideal -> nombre de hoja INKA
MAPEO_HOJA = {
    'Parque Unipraias':  'UNIPRAIA',
    'Beto Carrero World': 'BETO CARRERO',
    'Porto Belo':         'PORTO BELO',
    'Campamento Americano': 'ZACARIAS',
    'Cristo Luz':         'CRISTO LUZ',
    'Barco Pirata':       'BARCO PIRATA',
}

DIAS_ES = ["Lunes", "Martes", "Miercoles", "Jueves",
           "Viernes", "Sabado", "Domingo"]

# ----------------------------------------------------------------------
# LECTURA DEL ARCHIVO FUENTE
# ----------------------------------------------------------------------
def _num(txt):
    if not txt:
        return 0
    m = re.match(r'(\d+)', str(txt).strip())
    return int(m.group(1)) if m else 0


def leer_grupos(path, hoja=HOJA_ENTRADA):
    """Lee grupos + Itinerario Ideal + hora salida de las notas."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb[hoja]

    # --- grupos (filas 4+, columna A != None y no es fila de itinerario) ---
    grupos = []
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        if a is None or str(a).strip().lower() in ('grupo', ''):
            continue
        b = ws.cell(r, 2).value
        if isinstance(a, (int, float)) and b and str(b).strip().startswith('IN'):
            continue
        g = dict(
            fila=r,
            nombre_raw=str(a).strip(),
            estudiantes=_num(ws.cell(r, 3).value),
            adultos=_num(ws.cell(r, 4).value),
            total=ws.cell(r, 5).value or 0,
            coordinadores=ws.cell(r, 6).value or 0,
            in_date=ws.cell(r, 7).value,
            out_date=ws.cell(r, 8).value,
        )
        grupos.append(g)

    # re-asignar nombres secuenciales
    for i, g in enumerate(grupos, start=1):
        g['nombre'] = f'Grupo {i}'

    # --- itinerario ideal ---
    itinerarios = {}
    for r in range(4, ws.max_row + 1):
        a = ws.cell(r, 1).value
        b = ws.cell(r, 2).value
        if isinstance(a, (int, float)) and a >= 1 and b and str(b).strip().startswith('IN'):
            lineas = []
            for c in range(2, 10):
                v = ws.cell(r, c).value
                if v is not None:
                    lineas.append(str(v).strip())
            itinerarios[int(a)] = lineas
    for i, g in enumerate(grupos, start=1):
        g['itinerario'] = itinerarios.get(i, [])

    # --- hora salida desde notas (fila 11+) ---
    hora_salida = None
    for r in range(11, ws.max_row + 1):
        for c in range(1, 11):
            v = ws.cell(r, c).value
            if v and isinstance(v, str):
                m = re.search(r'salida\s*(\d{1,2}:\d{2})', v, re.IGNORECASE)
                if m:
                    hora_salida = m.group(1)
                    break
        if hora_salida:
            break

    wb.close()
    return grupos, hora_salida


def _extraer_bus(nombre_raw):
    m = re.search(r'(\d+)\s*van\s*(?:de\s*)?(\d+)', nombre_raw, re.IGNORECASE)
    return f'1 VAN ({m.group(2)})' if m else '1 VAN'


# ----------------------------------------------------------------------
# PARSEO DE UN DIA DEL ITINERARIO IDEAL
# ----------------------------------------------------------------------
def _parse_dia(texto):
    """Devuelve (main, items, has_fiesta, has_quinta)."""
    if not texto:
        return None, [], False, False
    lineas = [l.strip() for l in str(texto).split('\n') if l.strip()]
    if not lineas:
        return None, [], False, False
    main = lineas[0]
    items = [l.lstrip('- ').strip() for l in lineas[1:] if l.strip()]
    txt_lower = ' '.join(items).lower() + ' ' + main.lower()
    has_fiesta = 'fiesta' in txt_lower
    has_quinta = 'quinta' in txt_lower
    return main, items, has_fiesta, has_quinta


def _hoja_atraccion(main):
    for clave, nome in MAPEO_HOJA.items():
        if clave.lower() in main.lower():
            return nome
    return None


# ----------------------------------------------------------------------
# GENERACION DE CELDAS DE PROGRAMACION  (texto exacto de la referencia)
# ----------------------------------------------------------------------
def _celda_in(hora_salida, items):
    """Celda del dia IN."""
    atr = [i for i in items
           if i.lower() not in ('in', 'transfer almuerzo', '')]
    texto_atr = atr[-1] if atr else 'Cena Cristo Luz'
    h = hora_salida or '____'
    return (f'IN - Vuelo: \n- \n\n'
            f'Salida IQQ:\n{h}H\n\n'
            f'Llegada FLN:\n-H\n\n'
            f'{texto_atr}                \n'
            f'1 VAN                            ')


def _celda_out():
    return ('10:00 h \nCheck-out \n+ \n'
            'Almuerzo                       \n+\n'
            'Salida p/aeropuerto\n- H\n\n'
            'OUT - Vuelo: \n____\n\n'
            'Salida FLN:\n- H\n'
            'Llegada IQQ:\n - H                                                     \n'
            '1 VAN')


def _celda_unipraia(fiesta, quinta):
    base = ('Salida / Unipraia \nas - H\n'
            '                                              \n'
            'Teleferico\n+\n'
            '                                       Unipraias\n+\n'
            '                                 Almuerzo              \n+\n'
            '                                       Cena \n +\n'
            '                              \n')
    if fiesta:
        base += 'Fiesta\n+\n'
        if quinta:
            base += '\n5\u00aa COMIDA \n\n 1 VAN'
        else:
            base += '\n 1 VAN'
    else:
        base += ' \n\n 1 VAN'
    return base


def _celda_beto():
    return ('Salida p/Beto Carrero  \nas - H\n+ \n\n'
            'Almuerzo                                 +\n'
            '                                   Cena\n\n\n1 VAN')


def _celda_porto(fiesta, quinta):
    base = ('Salida p/Porto Belo  \nas - H\n+ \n\n'
            'Almuerzo.\n'
            '                                  +\n'
            '                                   Cena')
    if fiesta:
        base += ('                                    +\n'
                 'Fiesta\n+\n\n'
                 '5\u00aa COMIDA\n\n1 VAN')
    else:
        base += '\n\n\n1 VAN'
    return base


def _celda_libre(fiesta, quinta):
    base = ('Salida p/ \n'
            '  (Dia libre)                                                            +\n'
            'Almuerzo                     +\n'
            ' \n+\nCena\n+\n')
    if fiesta:
        base += ('\nFiesta  \n\n'
                 '+                                    5\u00aa COMIDA\n\n'
                 '1 VAN')
    else:
        base += '\n\n1 VAN'
    return base


def _celda_zacarias(fiesta, quinta):
    if fiesta:
        return ('Salida p/ Zacarias as -H \n\n+\n'
                'Parque acuatico                \n\n+\n'
                'Almuerzo  \n                          \n'
                ' Cena                               +                        Fiesta  \n\n'
                '+                                    5\u00aa COMIDA\n \n\n'
                ' 1 VAN ')
    return ('Salida p/ Zacarias as -H \n\n+\n'
            'Parque acuatico                \n\n+\n'
            'Almuerzo  \n                          \n'
            ' Cena\n\n+\n \n\n'
            ' 1 VAN ')


def celda_dia(main, items, fiesta, quinta, hora_salida):
    if not main:
        return ''
    low = main.lower()
    if main.strip().upper() == 'IN':
        return _celda_in(hora_salida, items)
    if any(i.upper() == 'OUT' for i in items):
        return _celda_out()
    if 'parque unipraias' in low:
        return _celda_unipraia(fiesta, quinta)
    if 'beto carrero' in low:
        return _celda_beto()
    if 'porto belo' in low:
        return _celda_porto(fiesta, quinta)
    if 'campamento americano' in low:
        return _celda_zacarias(fiesta, quinta)
    if 'dia libre' in low or 'día libre' in low:
        return _celda_libre(fiesta, quinta)
    # atraccion conocida sin plantilla propia (ej. Barco Pirata):
    # celda generica estilo Beto/Porto con el nombre REAL del itinerario.
    nome = _hoja_atraccion(main)
    if nome:
        base = (f'Salida p/{main.strip()}  \nas - H\n+ \n\n'
                'Almuerzo                                 +\n'
                '                                   Cena')
        if fiesta:
            base += ('                                    +\n'
                     'Fiesta\n+\n\n'
                     '5\u00aa COMIDA\n\n1 VAN')
        else:
            base += '\n\n\n1 VAN'
        return base
    return ''


# ----------------------------------------------------------------------
# BLOQUE DE LA COLUMNA B (PROGRAMACION)
# ----------------------------------------------------------------------
def _bloque_grupo(colg, bus):
    return (f"Colégio:\n{colg['nombre']} "
            f"\n \nAlunos: {colg['estudiantes']}"
            f"\nAdultos: {colg['adultos']}"
            f"\nGuia: {colg['coordinadores']} "
            f"\n\nTotal: {colg['total']}"
            f"\n\nHotel: - "
            f"\n\nGuia:\n----"
            f"\n \n\n\nBus: {bus}\n")


# ----------------------------------------------------------------------
# GENERACION
# ----------------------------------------------------------------------
def generar(in_path=ARCHIVO_ENTRADA, template_path=PLANTILLA,
            out_path=ARCHIVO_SALIDA):
    grupos, hora_salida = leer_grupos(in_path)

    for i, g in enumerate(grupos, start=1):
        g['num'] = i
        g['nombre'] = f'INKA GRUPO {i}'
        g['bus'] = _extraer_bus(g['nombre_raw'])
        ini = g['in_date']
        out = g['out_date']
        ini_s = ini.date() if isinstance(ini, datetime) else ini
        out_s = out.date() if isinstance(out, datetime) else out
        print(f"  {g['nombre']} ({g['nombre_raw']}): "
              f"IN={ini_s}  OUT={out_s}  "
              f"Alunos={g['estudiantes']}  Adultos={g['adultos']}  "
              f"Dias={len(g['itinerario'])}")

    # ---- repartir en atracciones y 5 COMIDA ----
    atracciones = {}          # nome_hoja -> [{fecha, nombre, estudiantes, adultos, total, coordinadores}]
    eventos_5comida = []      # [{fecha, ...}]
    for g in grupos:
        fecha_in = g['in_date']
        if isinstance(fecha_in, datetime):
            fecha_in = fecha_in.date()
        for idx, dia_txt in enumerate(g['itinerario']):
            main, items, fiesta, quinta = _parse_dia(dia_txt)
            if not main:
                continue
            fecha = fecha_in + timedelta(days=idx)
            info = dict(fecha=fecha, nombre=f"GRUPO {g['num']}",
                        estudiantes=g['estudiantes'], adultos=g['adultos'],
                        total=g['total'], coordinadores=g['coordinadores'])
            low = main.lower()
            # IN -> Cristo Luz
            if main.strip().upper() == 'IN':
                atracciones.setdefault('CRISTO LUZ', []).append(info)
                continue
            # OUT
            if any(i.upper() == 'OUT' for i in items):
                continue
            # atraccion normal
            nome = _hoja_atraccion(main)
            if nome:
                atracciones.setdefault(nome, []).append(info)
            # 5 COMIDA
            if quinta:
                eventos_5comida.append(info)

    print(f"\n  Atracciones: {', '.join(f'{k}({len(v)})' for k,v in atracciones.items())}")
    print(f"  5 COMIDA: {len(eventos_5comida)} eventos")

    # ---- abrir plantilla ----
    wb = openpyxl.load_workbook(template_path)

    # ================================================================
    # PROGRAMACION
    # ================================================================
    ws = wb['PROGRAMACION']
    ws['A1'] = 'PLANILLA DE COLEGIOS INKA - DEZEMBRO - 2026'

    # rango de fechas
    fechas_in = [g['in_date'] for g in grupos if g['in_date']]
    fechas_out = [g['out_date'] for g in grupos if g['out_date']]
    f_min = min(fechas_in)
    f_max = max(fechas_out)
    if isinstance(f_min, datetime): f_min = f_min.date()
    if isinstance(f_max, datetime): f_max = f_max.date()
    n_dias = (f_max - f_min).days + 1

    # headers de dia (fila 2, columna C..I)
    for i in range(min(n_dias, 7)):
        f = f_min + timedelta(days=i)
        ws.cell(row=2, column=3 + i).value = f'{DIAS_ES[f.weekday()]} {f.day}'

    # limpiar datos viejos
    for r in range(3, ws.max_row + 1):
        for c in range(1, 10):
            ws.cell(row=r, column=c).value = None

    # modelo de estilo (fila 3 de la plantilla)
    style_a3 = copy(ws['A3']._style)
    style_b3 = copy(ws['B3']._style)
    style_c3 = copy(ws['C3']._style)
    style_di3 = copy(ws['D3']._style)  # estilo comun dia

    # filas de grupos
    for g in grupos:
        r = 2 + g['num']  # fila 3 para grupo 1, fila 4 para grupo 2
        ws.cell(row=r, column=1).value = g['num']
        ws.cell(row=r, column=1)._style = style_a3
        ws.cell(row=r, column=2).value = _bloque_grupo(g, g['bus'])
        ws.cell(row=r, column=2)._style = style_b3

        # alinear por fecha: cada dia del itinerario cae en la columna
        # correspondiente a su fecha real (dias globales compartidos).
        g_in = g['in_date']
        if isinstance(g_in, datetime):
            g_in = g_in.date()
        offset = (g_in - f_min).days

        for idx, dia_txt in enumerate(g['itinerario'][:7]):
            col = 3 + offset + idx
            if col > 9:  # solo columnas C..I
                break
            main, items, fiesta, quinta = _parse_dia(dia_txt)
            ws.cell(row=r, column=col).value = celda_dia(
                main, items, fiesta, quinta, hora_salida)
            # columna C (dia IN) usa estilo rojo
            if col == 3:
                ws.cell(row=r, column=col)._style = style_c3
            else:
                ws.cell(row=r, column=col)._style = style_di3

        ws.row_dimensions[r].height = 340.2

    # ================================================================
    # HOJAS DE ATRACCION
    # ================================================================
    for nome_atr in ['CRISTO LUZ', 'UNIPRAIA', 'BETO CARRERO',
                     'PORTO BELO', 'ZACARIAS', 'BARCO PIRATA']:
        ws = wb[nome_atr]
        ws['A1'] = f'{nome_atr} - INKA - DEZEMBRO/2026'
        # limpiar datos
        for r in range(3, ws.max_row + 1):
            for c in range(1, 12):
                ws.cell(row=r, column=c).value = None

        visitas = atracciones.get(nome_atr, [])
        # estilos de fila de datos (filas 3-5 de la plantilla)
        if visitas:
            modelo_row = 3
            try:
                estilos = [copy(ws.cell(row=modelo_row, column=c)._style)
                           for c in range(1, 11)]
            except Exception:
                estilos = None
        else:
            estilos = None
        for i, v in enumerate(visitas):
            r = 3 + i
            if estilos:
                for c in range(1, 11):
                    ws.cell(row=r, column=c)._style = copy(estilos[c - 1])
            ws.cell(row=r, column=1).value = i + 1
            ws.cell(row=r, column=2).value = v['nombre']
            ws.cell(row=r, column=3).value = v['estudiantes']
            ws.cell(row=r, column=4).value = v['adultos']
            ws.cell(row=r, column=5).value = v['total']
            ws.cell(row=r, column=6).value = v['coordinadores']
            ws.cell(row=r, column=9).value = v['fecha']
            ws.cell(row=r, column=10).value = f'=E{r}+F{r}'

    # ================================================================
    # 5 COMIDA
    # ================================================================
    ws = wb['5 COMIDA']
    ws['A1'] = '5\u00aa COMIDA - INKA - DEZEMBRO/2026'
    # asegurar merge H2:J3 para DATA e SERVICIO
    try:
        ws.merge_cells('H2:J3')
    except Exception:
        pass
    ws['H2'] = 'DATA e SERVIC\u00c7O'

    for r in range(4, ws.max_row + 1):
        for c in range(1, 11):
            ws.cell(row=r, column=c).value = None

    for i, ev in enumerate(eventos_5comida):
        r = 4 + i
        if i == 0:
            try:
                estilos_comida = [copy(ws.cell(row=4, column=c)._style)
                                  for c in range(1, 10)]
            except Exception:
                estilos_comida = None
        if estilos_comida:
            for c in range(1, 10):
                ws.cell(row=r, column=c)._style = copy(estilos_comida[c - 1])
        ws.cell(row=r, column=1).value = i + 1
        ws.cell(row=r, column=2).value = ev['nombre']
        ws.cell(row=r, column=3).value = ev['estudiantes']
        ws.cell(row=r, column=4).value = ev['adultos']
        ws.cell(row=r, column=5).value = f'=C{r}+D{r}'
        ws.cell(row=r, column=6).value = ev['coordinadores']
        ws.cell(row=r, column=7).value = f'=E{r}+F{r}'
        ws.cell(row=r, column=8).value = ev['fecha']
        # I y J quedan vacios (servicio se llena a mano)

    # ---- guardar ----
    if out_path:
        wb.save(out_path)
        print(f"\n  Guardado: {out_path}")

    return wb, dict(grupos=len(grupos),
                    atracciones={a: len(v) for a, v in atracciones.items()},
                    comida5=len(eventos_5comida))


# ----------------------------------------------------------------------
if __name__ == '__main__':
    generar()
