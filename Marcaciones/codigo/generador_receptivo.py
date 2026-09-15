# -*- coding: utf-8 -*-
"""Generador de planillas para empresa receptiva de viajes.

REUTILIZA el archivo 'Cristour 2026 planilha.xlsx' como BASE: las hojas
'Cristo Luz ', '5 COMIDA' y 'Programacion ' se limpian y rellenan, y las
demas atracciones se crean copiando la hoja 'Cristo Luz ' completa, de modo
que TODO el formato (bordes, colores, anchos, celdas unidas, hojas con 860
filas, gridlines apagadas, titulos) queda exactamente igual al archivo que
entrega el usuario.

Datos de entrada:
  - 'Colegios totales' (lista maestra de colegios)
  - 'Marcaciones' (itinerario semanal por grupo) -> ver PATRONES_HARDCODED
"""

import openpyxl
from openpyxl.utils import get_column_letter
from datetime import datetime, timedelta
from copy import copy
import os, re, sys

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
IN_XLSX = os.path.join(ENTRADAS, "COLEGIOS CANTIDAD DE PASAJEROS, DISCOS, MARCACIONES (3).xlsx")
ROOMING_XLSX = os.path.join(ENTRADAS, "Rooming 2026.xlsx")
PLANTILLA = os.path.join(PLANTILLAS, "cristour_template.xlsx")
OUT_XLSX = os.path.join(SALIDAS, "SALIDA_SISTEMA_RECEPTIVO.xlsx")

TITULO_SUFIJO = "CAMBORIU - 2026/2027"

COLORES_FILA_B = ["FFFFFFCC", "FFCCECFF", "FFFFCCFF", "FFCCFFCC", "FFFFCCCC", "FFCCCCFF"]


# ---------------------------------------------------------------------------
# 1. LECTURA DE 'COLEGIOS TOTALES'
# ---------------------------------------------------------------------------
def leer_colegios(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Colegios totales"]
    lista = []
    for r in range(2, ws.max_row + 1):
        num = ws.cell(row=r, column=1).value
        if num is None:
            continue
        lista.append({
            "num": int(num),
            "file": ws.cell(row=r, column=2).value,
            "colegio": ws.cell(row=r, column=3).value or "",
            "curso": ws.cell(row=r, column=4).value or "",
            "alunos": ws.cell(row=r, column=5).value or 0,
            "adultos": None,
            "guias": ws.cell(row=r, column=6).value or 0,
            "total": ws.cell(row=r, column=7).value or 0,
            "destino": ws.cell(row=r, column=8).value or "",
            "programa": ws.cell(row=r, column=9).value or "",
            "salida": ws.cell(row=r, column=10).value,
            "llegada": ws.cell(row=r, column=11).value,
            "hotel": ws.cell(row=r, column=12).value or "",
            "cant_discos": ws.cell(row=r, column=13).value or "",
            "aerea": ws.cell(row=r, column=14).value or "",
            "obs": ws.cell(row=r, column=15).value or "",
        })
    return lista


# ---------------------------------------------------------------------------
# 1b. ROOMING LIST: cantidades reales de alumnos y adultos por colegio (File)
# ---------------------------------------------------------------------------
def leer_rooming(path):
    """Por File: estudiantes = H+M (col 7-8), adultos = I+J (col 9-10) y
    guias = col 11 (Guías Contrato), desde las hojas de hotel."""
    wb = openpyxl.load_workbook(path, data_only=True)
    res = {}
    for sh in wb.sheetnames:
        if sh in ("Hoja 14", "DSintra"):
            continue
        ws = wb[sh]
        for r in range(2, ws.max_row + 1):
            f = ws.cell(row=r, column=1).value
            if f is None:
                continue
            try:
                fi = int(float(f))
            except (TypeError, ValueError):
                continue
            e = int(ws.cell(row=r, column=7).value or 0)
            m = int(ws.cell(row=r, column=8).value or 0)
            ah = int(ws.cell(row=r, column=9).value or 0)
            am = int(ws.cell(row=r, column=10).value or 0)
            g = int(ws.cell(row=r, column=11).value or 0)
            t = res.get(fi, [0, 0, 0])
            res[fi] = [t[0] + e + m, t[1] + ah + am, t[2] + g]
    return res


def aplicar_rooming(colegios, rooming):
    """Actualiza alumnos/adultos/guias/total de cada colegio con los totales
    del rooming. Alumnos = estudiantes (H+M), adultos = adultos (I+J),
    guias = columna de Guías Contrato (sin separar por género)."""
    usados = 0
    for c in colegios:
        try:
            fi = int(float(c["file"]))
        except (TypeError, ValueError):
            fi = None
        if fi is not None and fi in rooming:
            est, ad, guias = rooming[fi]
            c["alunos"] = est
            c["adultos"] = ad
            c["guias"] = guias
            c["total"] = est + ad + guias
            usados += 1
        else:
            c["adultos"] = None
    return usados


# ---------------------------------------------------------------------------
# 2. MARCACIONES SEMANALES (leidas directo de la hoja 'Marcaciones')
# ---------------------------------------------------------------------------
# Cada bloque = una semana: fila con numeros de dia + fila con 'IN'..'OUT'.
# Debajo pueden venir filas de sub-atracciones y de noches.
# Noches: {'1based': boliche/visita}; 'Cristo Luz'/'Cristo Luz con Cena' =
# atraccion nocturna. Subs: {'1based': [sub-atracciones del dia]}.
MESES = {"septiembre": 9, "noviembre": 11, "diciembre": 12}
DIAS_SEMANA = {"lunes", "martes", "miercoles", "miércoles", "jueves",
               "viernes", "sabado", "sábado", "domingo"}
PALABRAS_NOCHE = ("cristo luz", "noche", "disco", "green valley", "lounge",
                  "eclipse", "fiesta", "shed")
ATRACCIONES_NOCTURNAS = {"cristo luz", "cristo luz con cena"}


def normalizar_atraccion(texto):
    t = (texto or "").strip()
    return re.sub(r"[?¿]+$", "", t).strip()


def _es_noche(texto):
    t = (texto or "").lower()
    return any(p in t for p in PALABRAS_NOCHE)


def _runes_numericas(ws, fila):
    """Ranges (col_inicio, col_fin) con numeros de dia consecutivos 1-31."""
    cols = []
    for c in range(1, ws.max_column + 1):
        v = ws.cell(row=fila, column=c).value
        if isinstance(v, (int, float)) and 1 <= int(v) <= 31:
            cols.append(c)
    runes = []
    inicio = prev = None
    for c in cols:
        if inicio is None:
            inicio = prev = c
        elif c == prev + 1:
            prev = c
        else:
            runes.append((inicio, prev))
            inicio = prev = c
    if inicio is not None:
        runes.append((inicio, prev))
    return [r for r in runes if r[1] - r[0] >= 3]


def leer_marcaciones(path):
    """Lee los bloques semanales de la hoja 'Marcaciones'."""
    wb = openpyxl.load_workbook(path, data_only=True)
    ws = wb["Marcaciones"]
    fila_num = {}
    for fila in range(1, ws.max_row + 1):
        if any(ws.cell(row=fila, column=c).value is not None
               for c in range(1, ws.max_column + 1)):
            runes = _runes_numericas(ws, fila)
            if runes:
                fila_num[fila] = runes

    # nombre del mes (titulo) mas cercano arriba de la fila num
    def mes_de(fila_num_):
        for r in range(fila_num_ - 1, 0, -1):
            for c in range(1, ws.max_column + 1):
                v = ws.cell(row=r, column=c).value
                if isinstance(v, str) and v.strip().lower() in MESES:
                    return MESES[v.strip().lower()]
        return None

    bloques = []
    for fn, runes in sorted(fila_num.items()):
        atr_row = fn + 1
        for c0, c1 in runes:
            atr = [ws.cell(row=atr_row, column=c).value for c in range(c0, c1 + 1)]
            if not atr or not str(atr[0]).strip().upper() == "IN":
                continue
            mes = mes_de(fn)
            if mes is None:
                continue
            dia0 = int(ws.cell(row=fn, column=c0).value)
            try:
                inicio = datetime(2026, mes, dia0)
            except ValueError:
                continue
            dias = [normalizar_atraccion(a) for a in atr]
            subs, noches = {}, {}
            # filas de datos debajo de la atraccion: sub y noche
            for r in range(atr_row + 1, atr_row + 4):
                for i, c in enumerate(range(c0, c1 + 1)):
                    v = ws.cell(row=r, column=c).value
                    if v is None or not isinstance(v, str):
                        continue
                    t = v.strip()
                    if not t or t.lower() in DIAS_SEMANA or t.lower() in MESES:
                        continue
                    if _es_noche(t):
                        noches.setdefault(i + 1, t)
                    else:
                        subs.setdefault(i + 1, [])
                        if t not in subs[i + 1]:
                            subs[i + 1].append(t)
            bloques.append({
                "inicio": inicio,
                "dias": dias,
                "subs": subs,
                "noches": noches,
                "horas": {},
            })
    bloques.sort(key=lambda b: b["inicio"])
    return bloques


def asignar_semana(colg, bloques):
    """Bloque de semana para un colegio: fecha de salida + duracion."""
    sal = colg["salida"]
    if isinstance(sal, str):
        try:
            sal = datetime.strptime(sal[:10], "%Y-%m-%d")
        except Exception:
            return None
    if sal is None:
        return None
    lle = colg["llegada"]
    if isinstance(lle, str):
        try:
            lle = datetime.strptime(lle[:10], "%Y-%m-%d")
        except Exception:
            lle = None
    dur = (lle - sal).days + 1 if lle is not None and lle >= sal else None

    cand = [b for b in bloques if b["inicio"] == sal]
    if not cand:
        return None
    if dur is not None:
        exactos = [b for b in cand if len(b["dias"]) == dur]
        if exactos:
            return exactos[0]
    return cand[0]  # sin duracion clara: primer bloque de esa salida


def construir_itinerario(colegio, patron):
    salida = colegio["salida"]
    if isinstance(salida, str):
        try:
            salida = datetime.strptime(salida[:10], "%Y-%m-%d")
        except Exception:
            return []
    if patron is None:
        return []
    llegada = colegio["llegada"]
    if isinstance(llegada, str):
        try:
            llegada = datetime.strptime(llegada[:10], "%Y-%m-%d")
        except Exception:
            llegada = None
    n_dias = len(patron["dias"])
    if llegada is not None:
        n_dias = min(n_dias, (llegada - salida).days + 1 if llegada >= salida else n_dias)
    it = []
    for i in range(n_dias):
        fecha = salida + timedelta(days=i)
        ultimo = (i == n_dias - 1)
        atraccion = "OUT" if ultimo and patron["dias"][i] != "OUT" else normalizar_atraccion(patron["dias"][i])
        noche = patron["noches"].get(i + 1)
        sub = (patron.get("subs") or {}).get(i + 1)
        hora = None
        horas_dia = (patron.get("horas") or {}).get(i + 1)
        if isinstance(horas_dia, dict):
            hora = horas_dia.get("sa")
        elif horas_dia:
            hora = horas_dia
        it.append({"dia": i + 1, "fecha": fecha, "atraccion": atraccion,
                   "noche": noche, "hora": hora, "sub": sub})
    return it


def distribuir(colegios):
    atracciones = {}
    eventos_5comida = []
    for colg in colegios:
        for step in colg["_it"]:
            atr = step["atraccion"]
            base = {"fecha": step["fecha"], "num": colg["num"], "colegio": colg["colegio"],
                    "curso": colg["curso"], "alunos": colg["alunos"], "adultos": colg["adultos"],
                    "guias": colg["guias"],
                    "total": colg["total"], "hotel": colg["hotel"], "cant_discos": colg["cant_discos"],
                    "obs": colg["obs"]}
            if atr and atr not in ("IN", "OUT"):
                atracciones.setdefault(atr, []).append(base)
            if atr == "OUT":
                break
            noche = step["noche"]
            if not noche:
                continue
            if noche.lower().strip() in ATRACCIONES_NOCTURNAS:
                atracciones.setdefault("Cristo Luz", []).append(base)
                continue
            if noche not in ("Noche Libre", "Noche libre"):
                eventos_5comida.append({**base, "boliche": noche})
    return atracciones, eventos_5comida


# ---------------------------------------------------------------------------
# 3. PLANTILLA (referencias de estilo HOY: la base es el archivo del usuario)
# ---------------------------------------------------------------------------
class Plantilla:
    def __init__(self, path):
        self.wb = openpyxl.load_workbook(path)
        self.atr = self.wb["Cristo Luz "]       # hoja por atraccion (modelo)
        self.comida = self.wb["5 COMIDA"]
        self.prog = self.wb["Programacion "]
        # celdas modelo de una fila de datos
        self.a_data = [self.atr[cell + "3"] for cell in "BCDEFGHIJK"]
        self.c_data = [self.comida[cell + "4"] for cell in "BCDEFGHIJK"]
        self.p_model = [self.prog[cell + "3"] for cell in "ABCDEFGHIJ"]
        self.fmt_fecha = self.atr["J3"].number_format or "DD/MM"

    def copiar_estilo(self, src, dst):
        if src.has_style:
            dst._style = copy(src._style)


def estilar_celdas(ws, modelos, letras, fila):
    for letra, src in zip(letras, modelos):
        ws[letra + str(fila)]._style = copy(src._style)


# ---------------------------------------------------------------------------
# 4. LIMPIEZA y RELLENO (mantiene TODO el formato existente de la plantilla)
# ---------------------------------------------------------------------------
def quitar_merges_excepto(ws, conservar):
    for mr in list(ws.merged_cells.ranges):
        if str(mr) not in conservar:
            ws.unmerge_cells(str(mr))


def limpiar_valores(ws, fila_inicio):
    for fila in ws.iter_rows(min_row=fila_inicio, min_col=1, max_col=12):
        for cel in fila:
            if cel.value is not None:
                cel.value = None


def agrupar_por_fecha(eventos):
    grupos = {}
    for ev in eventos:
        grupos.setdefault(ev["fecha"], []).append(ev)
    for fecha in grupos:
        grupos[fecha].sort(key=lambda e: e["num"])
    return grupos


def texto_nombre(ev):
    return f"{ev['colegio'].strip()} {ev['curso'].strip()}".strip()


def rellenar_atraccion(ws, P, nombre, eventos):
    """Hoja de atraccion: reutiliza el formato de 'Cristo Luz ' de la plantilla."""
    quitar_merges_excepto(ws, {"B1:K1", "B2:C2"})
    limpiar_valores(ws, 3)
    ws["B1"] = f"{nombre.upper()} - {TITULO_SUFIJO}"

    fila = 3
    grupos = agrupar_por_fecha(eventos)
    for i, (fecha, grupo) in enumerate(sorted(grupos.items(), key=lambda kv: kv[0])):
        r0, r1 = fila, fila + len(grupo) - 1
        for j, ev in enumerate(grupo):
            r = fila + j
            estilar_celdas(ws, P.a_data, "BCDEFGHIJK", r)
            ws.cell(row=r, column=2).value = ev["num"]          # B Nº
            ws.cell(row=r, column=3).value = texto_nombre(ev)   # C colegio + curso
            ws.cell(row=r, column=4).value = ev["alunos"]       # D ALUNOS (estudiantes)
            ws.cell(row=r, column=5).value = ev.get("adultos")  # E ADULTOS
            ws.cell(row=r, column=6).value = (ev["total"] or 0) - (ev["guias"] or 0)  # F TOAL (sin guias)
            ws.cell(row=r, column=7).value = ev["guias"]        # G GUIA
            ws.cell(row=r, column=8).value = ev["hotel"]        # H HOTEL
            ws.cell(row=r, column=9).value = None               # I HORARIO (queda vacia)
            ws.cell(row=r, column=10).value = fecha             # J DATA
            ws.cell(row=r, column=10).number_format = P.fmt_fecha
        # K TOTAL P/DIA (celda unida en todo el grupo, formula en la 1ª fila)
        if r1 == r0:
            ws.cell(row=r0, column=11).value = f"=F{r0}+G{r0}"
        else:
            ws.cell(row=r0, column=11).value = f"=SUM(F{r0}:F{r1})+SUM(G{r0}:G{r1})"
            ws.merge_cells(start_row=r0, start_column=10, end_row=r1, end_column=10)  # J
            ws.merge_cells(start_row=r0, start_column=11, end_row=r1, end_column=11)  # K
        # color pastel de grupo para TODAS las celdas de datos (B..K)
        color = COLORES_FILA_B[i % 6]
        for col in range(2, 12):
            for r in range(r0, r1 + 1):
                try:
                    cel = ws.cell(row=r, column=col)
                    cel.fill = openpyxl.styles.PatternFill(
                        fill_type="solid",
                        start_color=color,
                        end_color=color,
                    )
                except AttributeError:  # celdas unidas (J/K) heredan de su ancla
                    pass
        fila = r1 + 1


def rellenar_comida(ws, P, eventos):
    """5 COMIDA: mismo formato que la plantilla (grupos por fecha + boliche)."""
    conservar = {"B1:J1", "B2:C3", "D2:D3", "E2:E3", "F2:F3", "G2:G3", "H2:H3", "I2:J3"}
    quitar_merges_excepto(ws, conservar)
    limpiar_valores(ws, 4)
    ws["B1"] = f"5ª COMIDA - {TITULO_SUFIJO}"

    grupos = {}
    for ev in eventos:
        grupos.setdefault((ev["fecha"], ev["boliche"]), []).append(ev)
    orden = sorted(grupos.items(), key=lambda kv: (kv[0][0], kv[0][1], kv[1][0]["num"]))

    fila = 4
    for i, ((fecha, boliche), grupo) in enumerate(orden):
        r0, r1 = fila, fila + len(grupo) - 1
        for j, ev in enumerate(grupo):
            r = fila + j
            estilar_celdas(ws, P.c_data, "BCDEFGHIJ", r)
            ws.cell(row=r, column=2).value = j + 1                          # B Nº
            ws.cell(row=r, column=3).value = texto_nombre(ev)               # C colegio
            ws.cell(row=r, column=4).value = ev["alunos"]                   # D ALUNOS
            ws.cell(row=r, column=5).value = ev.get("adultos")              # E ADULTOS
            ws.cell(row=r, column=6).value = f"=D{r}+E{r}"                  # F TOTAL
            ws.cell(row=r, column=7).value = ev["guias"]                    # G GUIA
            ws.cell(row=r, column=9).value = fecha                          # I DATA
            ws.cell(row=r, column=9).number_format = P.fmt_fecha
        # H TOTAL P/DIA, I y J unidas en el grupo
        ws.cell(row=r0, column=8).value = f"=SUM(F{r0}:F{r1})+SUM(G{r0}:G{r1})"
        ws.cell(row=r0, column=10).value = boliche
        if r1 > r0:
            ws.merge_cells(start_row=r0, start_column=8, end_row=r1, end_column=8)    # H
            ws.merge_cells(start_row=r0, start_column=9, end_row=r1, end_column=9)    # I
            ws.merge_cells(start_row=r0, start_column=10, end_row=r1, end_column=10)  # J
        # color pastel de grupo para TODAS las celdas de datos (B..J)
        color = COLORES_FILA_B[i % 6]
        for col in range(2, 11):
            for r in range(r0, r1 + 1):
                try:
                    ws.cell(row=r, column=col).fill = openpyxl.styles.PatternFill(
                        fill_type="solid",
                        start_color=color,
                        end_color=color,
                    )
                except AttributeError:  # celdas unidas (H/I/J) heredan de su ancla
                    pass
        fila = r1 + 1


# ---------------------------------------------------------------------------
# 4b. PROGRAMACION POR SEMANAS (cada hoja = una salida, headers con fecha real)
# ---------------------------------------------------------------------------
DIAS_ES = ["Lunes", "Martes", "Miercoles", "Jueves", "Viernes", "Sabado", "Domingo"]


def nom_header_dia(fecha):
    return f"{DIAS_ES[fecha.weekday()]} {fecha.day:02d}"


def _num(v):
    try:
        if v is None:
            return "____"
        return str(int(v)) if float(v).is_integer() else str(v)
    except Exception:
        return "____" if v is None else str(v)


def texto_bloque_colegio(colg):
    """Bloque de la celda B de PROGRAMACION, igual al formato del original."""
    nombre = f"{colg['colegio'].strip()} {colg['curso'].strip()}"
    return (f"Colégio:\n{nombre}"
            f"\n \nAlunos: {_num(colg.get('alunos'))}"
            f"\nAdultos: {_num(colg.get('adultos'))}"
            f"\nGuia: {_num(colg.get('guias'))}"
            f"\n \nTotal:{_num(colg.get('total'))}"
            f"\n\nHotel: {colg.get('hotel') or '____'}"
            f"\n\nGuia:\n____"
            f"\n \n\nBus: ____")


def _noche_excel(noche):
    return modelo._noche(noche)


def _hh_receptivo(hora):
    return modelo._hh(hora)


def hora_dia(step):
    """Salida del día (as HH:MMH) usando step['hora'] o el diccionario interno."""
    h = step.get("hora") or step.get("salida")
    return modelo._hh(h) if h is not None else '____H'


def _parsear_aerea(aerea):
    """'LA_ANF | FLN_Charter1_01/12/2026' -> (aerolinea, origen, destino)."""
    texto = (aerea or '').strip()
    if not texto:
        return None, None, None
    parte_a, _, parte_b = texto.partition('|')
    origen = None
    aerolinea = None
    for tok in parte_a.split('_'):
        tok = tok.strip().upper()
        if not tok:
            continue
        if tok in ('LA', 'H2', 'JJ', 'G3', 'SKY'):
            aerolinea = tok
        elif len(tok) == 3:
            origen = tok
    destino = None
    if parte_b:
        dest = parte_b.split('_')[0].strip().upper()
        if len(dest) == 3:
            destino = dest
    return aerolinea, origen, destino


def texto_dia_programacion(step, colg):
    """Celda del mapa de dias con el MISMO formato que la planilla modelo
    del usuario (bloques cortos con '+', horas 'HH:MMH', sub-atracciones,
    restaurante y aeropuertos desde los datos / diccionario interno)."""
    atr = step["atraccion"]
    noche = (step["noche"] or "").strip()
    aero, origen, destino = _parsear_aerea(colg.get("aerea"))
    hraw = step.get("hora") or step.get("salida")

    if atr == "IN":
        return modelo.texto_in(aero, origen, hraw, destino or 'FLN', None)
    if atr == "OUT":
        return modelo.texto_out(aero, bus_hora=hraw,
                                sal_cod='FLN', leg_cod=origen)
    return modelo.texto_dia(atr, hraw, sub=step.get("sub"), noche=noche)


def cabeceras_semana(salida):
    return [nom_header_dia(salida + timedelta(days=i)) for i in range(8)]


def agrupar_por_salida(colegios):
    grupos = {}
    sin_marcacion = []
    for c in colegios:
        sal = c["salida"]
        if isinstance(sal, str):
            try:
                sal = datetime.strptime(sal[:10], "%Y-%m-%d")
            except Exception:
                sal = None
        if sal is None or not (c.get("_it") or []):
            sin_marcacion.append(c)
        else:
            grupos.setdefault(sal, []).append(c)
    return grupos, sin_marcacion


def rellenar_programacion_semana(ws, P, colegios, headers, sin_marcacion=False):
    n = len(colegios)
    limpiar_valores(ws, 3)                       # conserva fila 2 y A1
    quitar_merges_excepto(ws, {"A1:J1"})

    # headers de los dias con la fecha real de la semana
    for i, texto in enumerate(headers[:8]):
        if texto:
            ws.cell(row=2, column=3 + i).value = texto

    if n > 18:  # la plantilla solo tiene estilo de datos hasta la fila 21
        for r in range(22, n + 3):
            estilar_celdas(ws, P.p_model, "ABCDEFGHIJ", r)

    for r in range(3, n + 3):  # columna B: 6 colores pasteles ciclicos como el original
        cel = ws.cell(row=r, column=2)
        if not cel.has_style:
            cel._style = copy(P.p_model[1]._style)
        cel.fill = openpyxl.styles.PatternFill(
            fill_type="solid",
            start_color=COLORES_FILA_B[(r - 3) % 6],
            end_color=COLORES_FILA_B[(r - 3) % 6],
        )

    altura = ws.row_dimensions[3].height or 277.2
    for r in range(3, n + 3):
        ws.row_dimensions[r].height = altura

    r = 3
    for colg in sorted(colegios, key=lambda c: c["num"]):
        ws.cell(row=r, column=1).value = colg["num"]
        bloque = texto_bloque_colegio(colg)
        ws.cell(row=r, column=2).value = bloque
        if sin_marcacion:
            for d in range(8):
                cel = ws.cell(row=r, column=3 + d)
                cel.value = "SIN\nMARCACIÓN"
                cel.font = openpyxl.styles.Font(color="FFFF0000", bold=True)
        else:
            for step in (colg.get("_it") or []):
                dia = step["dia"]
                if dia > 8:
                    continue
                ws.cell(row=r, column=2 + dia).value = texto_dia_programacion(step, colg)
        r += 1


def hoja_colegios(wb, colegios):
    ws = wb.create_sheet("COLEGIOS MAESTRO")
    hdrs = ["Nº", "FILE", "COLEGIO", "CURSO", "ALUMNOS", "ADULTOS", "GUÍAS", "TOTAL", "HOTEL", "SALIDA", "LLEGADA",
            "CANT DISCOS", "PRE ASIG AEREA", "OBSERVACIONES"]
    for i, h in enumerate(hdrs, start=1):
        c = ws.cell(row=1, column=i, value=h)
        c.font = openpyxl.styles.Font(bold=True)
    r = 2
    for colg in sorted(colegios, key=lambda c: c["num"]):
        row = [colg["num"], colg["file"], colg["colegio"], colg["curso"], colg["alunos"], colg["adultos"],
               colg["guias"], colg["total"], colg["hotel"], colg["salida"], colg["llegada"], colg["cant_discos"],
               colg["aerea"], colg["obs"]]
        for i, v in enumerate(row, start=1):
            ws.cell(row=r, column=i, value=v)
        r += 1
    widths = [6, 10, 36, 20, 10, 9, 8, 9, 16, 12, 12, 26, 26, 40]
    for i, w in enumerate(widths, start=1):
        ws.column_dimensions[get_column_letter(i)].width = w


# ---------------------------------------------------------------------------
# 5. GENERACIÓN
# ---------------------------------------------------------------------------
def generar(in_path, plantilla_path, out_path=None):
    colegios = leer_colegios(in_path)
    print(f"colegios leidos: {len(colegios)}")
    if os.path.exists(ROOMING_XLSX):
        n_room = aplicar_rooming(colegios, leer_rooming(ROOMING_XLSX))
        print(f"rooming aplicado a {n_room} colegios")
    P = Plantilla(plantilla_path)
    wb = P.wb  # trabajamos SOBRE el archivo del usuario

    for extra in ("Cristo Luz  (2)", "Cristo Luz  (3)"):
        if extra in wb.sheetnames:
            wb.remove(wb[extra])

    bloques = leer_marcaciones(in_path)
    for colg in colegios:
        colg["_it"] = construir_itinerario(colg, asignar_semana(colg, bloques))

    atracciones, eventos_5comida = distribuir(colegios)
    for atr in sorted(atracciones.keys()):
        f = sorted({e["fecha"] for e in atracciones[atr]})
        print(f"atraccion: {atr:<16} visitas={len(atracciones[atr]):<4} fechas={len(f)} {[d.strftime('%d/%m') for d in f]}")
    print(f"5 COMIDA: {len(eventos_5comida)} eventos")

    # 1) PROGRAMACION: una hoja por semana (copias de la hoja original)
    base_prog = wb["Programacion "]
    grupos, sin_marcacion = agrupar_por_salida(colegios)
    for salida in sorted(grupos.keys()):
        ws = wb.copy_worksheet(base_prog)
        ws.title = "PROGRAMACION " + salida.strftime("%d-%m")
        rellenar_programacion_semana(ws, P, grupos[salida], cabeceras_semana(salida))
    if sin_marcacion:
        ws = wb.copy_worksheet(base_prog)
        ws.title = "PROGRAMACION SIN MARCACION"
        rellenar_programacion_semana(ws, P, sin_marcacion,
                                     cabeceras_semana(datetime(2026, 12, 1)),
                                     sin_marcacion=True)
    wb.remove(base_prog)

    # 2) atracciones: 'Cristo Luz' usa la hoja base; las demas = copia identica
    base = wb["Cristo Luz "]
    for atr in sorted(atracciones.keys()):
        if atr == "Cristo Luz":
            ws = base
        else:
            ws = wb.copy_worksheet(base)
            ws.title = atr
        rellenar_atraccion(ws, P, atr, atracciones[atr])

    # 3) 5 COMIDA
    rellenar_comida(wb["5 COMIDA"], P, eventos_5comida)

    # 4) hoja extra de referencia
    hoja_colegios(wb, colegios)

    if out_path:
        wb.save(out_path)
        print("guardado:", out_path)

    return wb, {"colegios": len(colegios),
                "atracciones": {a: len(v) for a, v in atracciones.items()},
                "5comida": len(eventos_5comida)}


def main():
    generar(IN_XLSX, PLANTILLA, OUT_XLSX)


if __name__ == "__main__":
    main()