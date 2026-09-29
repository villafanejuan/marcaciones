from __future__ import annotations

import math
import os
import re
from collections import Counter
from datetime import datetime, timedelta

from openpyxl import Workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

from modelos import (  # noqa: E402
    ErrorDistribucion,
    Habitacion,
    Sexo,
    TIPOS_TRIPULACION,
    TipoPax,
)
from motor import OpcionesMotor, ReporteGeneral, clave_orden_habitacion, distribuir

PALETA = [
    "C2410C", "0E7490", "A21CAF", "4D7C0F", "B45309",
    "BE185D", "4338CA", "15803D", "7C3AED", "0369A1",
]
COLOR_LIBRE = "D6EADF"
COLOR_BLOQUEADA = "F2D7D7"
BORDE = Side(style="thin", color="9CA3AF")
BORDE_GRILLA = Border(left=BORDE, right=BORDE, top=BORDE, bottom=BORDE)
POR_FILA = 8


def _piso_de(numero: str) -> str:
    limpio = numero.strip()
    if limpio.isdigit():
        return str(int(limpio) // 100) if int(limpio) >= 100 else "0"
    return limpio[:1].upper()


def _formatear(d) -> str:
    return d.strftime("%d/%m/%Y")


def _dias_rango(desde, hasta) -> list:
    dias = []
    d = desde
    while d <= hasta:
        dias.append(d)
        d += timedelta(days=1)
    return dias


def _camas_limite_por_dia(hotel, dia):
    if hotel.lotes_camas:
        return int(hotel.camas_inventario_en(dia))
    return int(hotel.total_camas())


def exportar_mapa_grupos(ruta: str, hotel, grupos, reporte, desde, hasta, modo: str,
                         titulo_empresa: str = "Tudo Bem Eventos") -> int:
    """Matriz 'empresa x noches' en el formato del mapa-pessoas/apartamentos.

    modo: "Pessoas" (camas) o "Apartamentos". Un grupo ocupa su PAX entero
    cada noche; los apartamentos por noche salen de reporte (si se simulo).
    Devuelve cuantos grupos quedaron pendientes de simular (solo modo apartamentos).
    """
    import datetime
    dias = _dias_rango(desde, hasta)
    dist_por_grupo = {}
    if reporte is not None:
        for d in reporte.distribuciones:
            dist_por_grupo[id(d.grupo)] = d
    libro = Workbook()
    hoja = libro.active
    hoja.title = "Mapa"
    hoja["A1"] = titulo_empresa
    hoja["A1"].font = Font(size=13, bold=True, color="F59E0B")
    hoja["A2"] = (f"Mapa de {modo} - {hotel.nombre} - "
                  f"{desde.strftime('%d/%m')} a {hasta.strftime('%d/%m')}")
    hoja["A2"].font = Font(size=13, bold=True)
    hoja["A3"] = (f"Hotel: {hotel.nombre}  |  "
                  f"Período: {desde.strftime('%d/%m')} a {hasta.strftime('%d/%m')}")
    hoja["A3"].font = Font(size=10, color="6B7280")
    hoja["A4"] = f"Relatório gerado em: {datetime.datetime.now().strftime('%d/%m/%Y, %H:%M:%S')}"
    hoja["A4"].font = Font(size=10, italic=True, color="6B7280")

    total_col = 5 + len(dias)
    fila = 5
    for c, nombre in enumerate(["Empresa", "PAX", "IN", "OUT"] + list(range(1, len(dias) + 1)) + ["Total"],
                               start=1):
        celda = hoja.cell(row=fila, column=c, value=nombre)
        celda.font = Font(bold=True)
        celda.fill = PatternFill("solid", start_color="E7E5E4")
        celda.alignment = Alignment(horizontal="center")
        celda.border = BORDE_GRILLA
    fila = 7

    ocupados_por_dia = [0] * len(dias)
    pendientes = 0
    for g in grupos:
        if g.total_pax() < 1 or not g.tiene_fechas:
            continue
        for c, valor in ((1, g.nombre), (2, g.total_pax()),
                         (3, g.fecha_ingreso.strftime("%d/%m")),
                         (4, g.fecha_egreso.strftime("%d/%m"))):
            hoja.cell(row=fila, column=c, value=valor).border = BORDE_GRILLA
        dist = dist_por_grupo.get(id(g))
        aptos = dist.habitaciones_usadas if dist else None
        total = 0
        for c, d in enumerate(dias):
            if not (g.fecha_ingreso <= d < g.fecha_egreso):
                continue
            if modo == "Pessoas":
                valor = g.total_pax()
                ocupados_por_dia[c] += valor
            else:
                if aptos is None:
                    pendientes += 1
                    continue
                valor = aptos
                ocupados_por_dia[c] += valor
            celda = hoja.cell(row=fila, column=5 + c, value=valor)
            celda.alignment = Alignment(horizontal="center")
            celda.border = BORDE_GRILLA
            total += valor
        celda = hoja.cell(row=fila, column=total_col, value=total or None)
        celda.font = Font(bold=True)
        celda.border = BORDE_GRILLA
        fila += 1

    fila += 1
    if modo == "Pessoas":
        limites = [_camas_limite_por_dia(hotel, d) for d in dias]
    else:
        limites = [len(hotel.habitaciones)] * len(dias)
    for nombre, valores in (("Limite Diário", limites),
                            ("Ocupados", ocupados_por_dia),
                            ("Disponíveis", [l - o for l, o in zip(limites, ocupados_por_dia)])):
        celda = hoja.cell(row=fila, column=1, value=nombre)
        celda.font = Font(bold=True, color="374151")
        celda.fill = PatternFill("solid", start_color="F3F4F6")
        for c, v in enumerate(valores):
            celda_d = hoja.cell(row=fila, column=5 + c, value=v)
            celda_d.border = BORDE_GRILLA
            if nombre == "Disponíveis":
                celda_d.font = Font(color="DC2626" if v < 0 else "059669", bold=v < 0)
                celda_d.alignment = Alignment(horizontal="center")
            else:
                celda_d.font = Font(bold=True)
                celda_d.alignment = Alignment(horizontal="center")
        fila += 1

    hoja.column_dimensions["A"].width = 42
    for c in (2, 3, 4):
        hoja.column_dimensions[get_column_letter(c)].width = 7
    for c in range(5, 5 + len(dias)):
        hoja.column_dimensions[get_column_letter(c)].width = 6
    hoja.column_dimensions[get_column_letter(total_col)].width = 8
    hoja.freeze_panes = "E7"
    try:
        libro.save(ruta)
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo escribir el Excel '{ruta}': {exc}")
    return pendientes


def exportar_excel(reporte: ReporteGeneral, ruta: str) -> int:
    try:
        libro = Workbook()
        filas_pax = _hoja_mapa(libro, reporte)
        _hoja_resumen(libro, reporte)
        _hoja_pasajeros(libro, reporte)
        libro.save(ruta)
        return filas_pax
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo escribir el Excel '{ruta}': {exc}")


def _hoja_mapa(libro: Workbook, reporte: ReporteGeneral) -> int:
    hoja = libro.active
    hoja.title = "Mapa del hotel"
    hotel = reporte.hotel
    asignadas: dict[str, tuple] = {}
    for indice, dist in enumerate(reporte.distribuciones):
        for asig in dist.asignaciones:
            asignadas[asig.habitacion.numero] = (indice, dist, asig)

    hoja["B2"] = hotel.nombre
    hoja["B2"].font = Font(size=18, bold=True, color="1F2937")
    generado = datetime.now().strftime("%d/%m/%Y %H:%M")
    grupos_txt = ", ".join(d.grupo.nombre for d in reporte.distribuciones) or "sin simulacion"
    hoja["B3"] = f"Distribucion generada el {generado} - Grupos: {grupos_txt}"
    hoja["B3"].font = Font(size=10, italic=True, color="6B7280")

    fila = 5
    fila = _leyenda(hoja, fila, reporte)

    pisos: dict[str, list[Habitacion]] = {}
    for h in hotel.habitaciones:
        pisos.setdefault(_piso_de(h.numero), []).append(h)
    for piso in sorted(pisos.keys(), key=lambda k: (0, int(k)) if k.isdigit() else (1, k)):
        hoja.cell(row=fila, column=2, value=f"PISO {piso}").font = Font(
            size=12, bold=True, color="374151")
        fila += 1
        rooms = sorted(pisos[piso], key=lambda h: clave_orden_habitacion(h.numero))
        for tanda in range(math.ceil(len(rooms) / POR_FILA)):
            lote = rooms[tanda * POR_FILA:(tanda + 1) * POR_FILA]
            for offset, h in enumerate(lote):
                celda = hoja.cell(row=fila, column=2 + offset * 2)
                par = asignadas.get(h.numero)
                if par is not None:
                    indice, dist, asig = par
                    relleno = PALETA[indice % len(PALETA)]
                    etiqueta = f"{dist.grupo.nombre[:14]}\n{len(asig.nombres_pax)} pax"
                elif h.bloqueos:
                    relleno = COLOR_BLOQUEADA
                    etiqueta = f"Mant.\n{h.bloqueos[0].titulo[:14]}"
                else:
                    relleno = COLOR_LIBRE
                    etiqueta = "Libre"
                celda.value = f"{h.numero}\n{etiqueta}"
                celda.fill = PatternFill("solid", fgColor=relleno)
                celda.border = BORDE_GRILLA
                celda.alignment = Alignment(horizontal="center", vertical="center",
                                            wrap_text=True)
                vecina = hoja.cell(row=fila, column=3 + offset * 2)
                vecina.fill = PatternFill("solid", fgColor=relleno)
                vecina.border = BORDE_GRILLA
            hoja.row_dimensions[fila].height = 46
            fila += 1
        fila += 1

    hoja.column_dimensions["A"].width = 2
    for col in range(2, 2 + POR_FILA * 2):
        hoja.column_dimensions[get_column_letter(col)].width = 11
    return sum(len(a.nombres_pax) for d in reporte.distribuciones for a in d.asignaciones)


def _leyenda(hoja, fila: int, reporte: ReporteGeneral) -> int:
    def chip(columna: int, color_hex: str, texto: str):
        muestra = hoja.cell(row=fila, column=columna)
        muestra.fill = PatternFill("solid", fgColor=color_hex)
        muestra.border = BORDE_GRILLA
        texto_celda = hoja.cell(row=fila, column=columna + 1, value=texto)
        texto_celda.font = Font(size=10, color="374151")

    columna = 2
    chip(columna, COLOR_LIBRE, "Libre")
    chip(columna + 3, COLOR_BLOQUEADA, "Bloqueada / mantenimiento")
    fila += 1
    for indice, dist in enumerate(reporte.distribuciones):
        g = dist.grupo
        chip(2, PALETA[indice % len(PALETA)],
             f"{g.nombre}: {_formatear(g.fecha_ingreso)} a {_formatear(g.fecha_egreso)}"
             f" ({g.total_pax()} pax)")
        fila += 1
    return fila + 1


def _hoja_resumen(libro: Workbook, reporte: ReporteGeneral) -> None:
    hoja = libro.create_sheet("Resumen")
    encabezados = ["Grupo", "Ingreso", "Egreso", "Noches", "PaxTotal", "PaxAlojados",
                   "Faltantes", "HabitacionesUsadas", "CamasOcupadas", "CamasTotales",
                   "Desperdicio", "Estrategia", "Optimo", "Completo"]
    for col, titulo in enumerate(encabezados, start=1):
        celda = hoja.cell(row=1, column=col, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
    for num, dist in enumerate(reporte.distribuciones, start=2):
        g = dist.grupo
        camas_totales = sum(a.habitacion.capacidad for a in dist.asignaciones)
        camas_ocupadas = sum(a.carga for a in dist.asignaciones)
        valores = [g.nombre, _formatear(g.fecha_ingreso), _formatear(g.fecha_egreso),
                   g.noches(), g.total_pax(), dist.pax_alojados, len(dist.faltantes),
                   dist.habitaciones_usadas, camas_ocupadas, camas_totales,
                   camas_totales - camas_ocupadas, dist.estrategia_usada,
                   "SI" if dist.optimo else "NO", "SI" if dist.completo else "NO"]
        for col, valor in enumerate(valores, start=1):
            hoja.cell(row=num, column=col, value=valor)
    anchos = [22, 11, 11, 7, 9, 12, 9, 18, 13, 12, 11, 12, 7, 9]
    for col, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(col)].width = ancho
    hoja.freeze_panes = "A2"


def _hoja_pasajeros(libro: Workbook, reporte: ReporteGeneral) -> int:
    hoja = libro.create_sheet("Pasajeros")
    encabezados = ["Grupo", "Ingreso", "Egreso", "Noches", "Habitacion", "Capacidad",
                   "Ocupacion", "Nombre", "Tipo", "Sexo", "Edad", "Familia", "ParejaDe",
                   "Single", "Notas"]
    for col, titulo in enumerate(encabezados, start=1):
        celda = hoja.cell(row=1, column=col, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
    filas = 1
    for dist in reporte.distribuciones:
        grupo = dist.grupo
        for asig in dist.asignaciones:
            hab = asig.habitacion
            for atomo in asig.atomos:
                for p in atomo.miembros:
                    valores = [grupo.nombre, _formatear(grupo.fecha_ingreso),
                               _formatear(grupo.fecha_egreso), grupo.noches(), hab.numero,
                               hab.capacidad, asig.carga, p.nombre, p.tipo.value,
                               p.sexo.value, p.edad if p.edad is not None else "",
                               p.familia or "", p.pareja_de or "",
                               "SI" if (p.requiere_single or atomo.requiere_single) else "",
                               p.notas]
                    for col, valor in enumerate(valores, start=1):
                        hoja.cell(row=filas + 1, column=col, value=valor)
                    filas += 1
        for p in dist.faltantes:
            valores = [grupo.nombre, _formatear(grupo.fecha_ingreso),
                       _formatear(grupo.fecha_egreso), grupo.noches(), "SIN HABITACION",
                       "", "", p.nombre, p.tipo.value, p.sexo.value,
                       p.edad if p.edad is not None else "", p.familia or "",
                       p.pareja_de or "", "SI" if p.requiere_single else "", p.notas]
            for col, valor in enumerate(valores, start=1):
                hoja.cell(row=filas + 1, column=col, value=valor)
            filas += 1
    anchos = [20, 11, 11, 7, 11, 9, 10, 24, 11, 6, 6, 10, 12, 7, 18]
    for col, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(col)].width = ancho
    hoja.freeze_panes = "A2"
    hoja.auto_filter.ref = f"A1:{get_column_letter(len(encabezados))}{max(filas, 2)}"
    return filas


def _categoria_pax(p) -> str:
    if p.tipo in TIPOS_TRIPULACION:
        return "Guías (mixto)"
    if p.tipo == TipoPax.ESTUDIANTE:
        if p.sexo == Sexo.M:
            return "Estudiantes - Hombres"
        if p.sexo == Sexo.F:
            return "Estudiantes - Mujeres"
        return "Estudiantes (mixto)"
    if p.tipo in (TipoPax.NINO, TipoPax.BEBE):
        if p.sexo == Sexo.M:
            return "Niños - Hombres"
        if p.sexo == Sexo.F:
            return "Niños - Mujeres"
        return "Niños (mixto)"
    if p.sexo == Sexo.M:
        return "Adultos - Hombres"
    if p.sexo == Sexo.F:
        return "Adultos - Mujeres"
    return "Adultos (mixto)"


_CATEGORIA_ORDEN = [
    "Estudiantes - Hombres",
    "Estudiantes - Mujeres",
    "Estudiantes (mixto)",
    "Adultos - Hombres",
    "Adultos - Mujeres",
    "Adultos (mixto)",
    "Niños - Hombres",
    "Niños - Mujeres",
    "Niños (mixto)",
    "Guías (mixto)",
]


def _categoria_habitacion(asig) -> str:
    conteo: Counter = Counter()
    for atomo in asig.atomos:
        for p in atomo.miembros:
            conteo[_categoria_pax(p)] += 1
    if not conteo:
        return "Sin categoría"
    return max(conteo, key=lambda c: (conteo[c], -_CATEGORIA_ORDEN.index(c) if c in _CATEGORIA_ORDEN else 0))


def _etiqueta_capacidad(cap: int) -> str:
    mapa = {1: "Single", 2: "Doble", 3: "Triple", 4: "Cuádruple", 5: "Quíntuple", 6: "Séxtuple"}
    return mapa.get(cap, f"{cap} plazas")


def _header_capacidades(max_cap: int) -> list[str]:
    return [_etiqueta_capacidad(c) for c in range(1, max(max_cap + 1, 6))]


def _conteo_capacidades(asignaciones) -> Counter:
    return Counter(a.habitacion.capacidad for a in asignaciones)


def _libro_rooming(hotel, grupos, reporte, destino: str, titulo_empresa: str) -> Workbook:
    """Construye el Workbook Rooming (Resumen + Detalle por Categoría + Cambios)."""
    libro = Workbook()

    max_cap = max(
        (a.habitacion.capacidad for d in reporte.distribuciones for a in d.asignaciones),
        default=5,
    )
    cap_headers = _header_capacidades(max_cap)
    cap_cols = list(range(1, len(cap_headers) + 1))

    archivo_counter = 0
    total_filas_resumen = 0
    total_filas_detalle = 0

    # ── Hoja Resumen por Grupo ────────────────────────────────────────────
    hoja = libro.active
    hoja.title = "Resumen por Grupo"
    encabezados = ["File", "Colegio", "Hotel", "Destino", "Check-in", "Check-out",
                   "Pasajeros", "Habitaciones"] + cap_headers + ["Sin habitación", "Estado"]
    for c, titulo in enumerate(encabezados, start=1):
        celda = hoja.cell(row=1, column=c, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
        celda.alignment = Alignment(horizontal="center")
        celda.border = BORDE_GRILLA

    for dist in reporte.distribuciones:
        archivo_counter += 1
        archivo_counter_str = str(archivo_counter)
        g = dist.grupo
        conteo = _conteo_capacidades(dist.asignaciones)
        estado = "COMPLETO" if dist.completo else "INCOMPLETO"
        if g.confirmado:
            estado += " · CONFIRMADO"
        elif g.provisional:
            estado += " · ESTIMADO"
        fila_valores = [
            archivo_counter_str,
            g.nombre,
            hotel.nombre,
            destino,
            _formatear(g.fecha_ingreso),
            _formatear(g.fecha_egreso),
            g.total_pax(),
            dist.habitaciones_usadas,
        ] + [conteo.get(cap, 0) for cap in cap_cols] + [
            len(dist.faltantes),
            estado,
        ]
        fila = 2 + total_filas_resumen
        for c, valor in enumerate(fila_valores, start=1):
            celda = hoja.cell(row=fila, column=c, value=valor)
            celda.border = BORDE_GRILLA
            celda.alignment = Alignment(horizontal="center")
        total_filas_resumen += 1

    anchos = [8, 38, hotel.nombre and len(hotel.nombre) or 12, len(destino) or 12,
              12, 12, 9, 10] + [9] * len(cap_headers) + [14, 11]
    for c, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(c)].width = min(ancho, 42)
    hoja.freeze_panes = "A2"

    # ── Hoja Detalle por Categoría ────────────────────────────────────────
    hoja2 = libro.create_sheet("Detalle por Categoría")
    encabezados2 = ["File", "Colegio", "Hotel", "Check-in", "Categoría",
                    "Pasajeros", "Habitaciones"] + cap_headers + ["Sin habitación", "Excepción"]
    for c, titulo in enumerate(encabezados2, start=1):
        celda = hoja2.cell(row=1, column=c, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
        celda.alignment = Alignment(horizontal="center")
        celda.border = BORDE_GRILLA

    archivo_counter = 0
    for dist in reporte.distribuciones:
        archivo_counter += 1
        archivo_counter_str = str(archivo_counter)
        g = dist.grupo

        cat_hab: dict[str, int] = Counter()
        cat_cap: dict[str, Counter] = {}
        for asig in dist.asignaciones:
            cat = _categoria_habitacion(asig)
            cat_hab[cat] += 1
            cat_cap.setdefault(cat, Counter())[asig.habitacion.capacidad] += 1

        cat_falt: dict[str, int] = Counter()
        for p in dist.faltantes:
            cat_falt[_categoria_pax(p)] += 1

        todas_cats = set(cat_hab.keys()) | set(cat_falt.keys())
        for cat in sorted(todas_cats, key=lambda c: _CATEGORIA_ORDEN.index(c) if c in _CATEGORIA_ORDEN else 99):
            falt = cat_falt.get(cat, 0)
            pax_cat = cat_hab.get(cat, 0) + falt
            hab_cat = cat_hab.get(cat, 0)
            excepcion = (
                f"Queda {falt} persona(s) en '{cat}' sin posibilidad de "
                f"compartir habitación dentro de las reglas."
                if falt > 0 else ""
            )
            fila_valores = [
                archivo_counter_str,
                g.nombre,
                hotel.nombre,
                _formatear(g.fecha_ingreso),
                cat,
                pax_cat,
                hab_cat,
            ] + [
                cat_cap.get(cat, Counter()).get(cap, 0) for cap in cap_cols
            ] + [falt, excepcion]
            fila = 2 + total_filas_detalle
            for c, valor in enumerate(fila_valores, start=1):
                celda = hoja2.cell(row=fila, column=c, value=valor)
                celda.border = BORDE_GRILLA
                celda.alignment = Alignment(horizontal="center")
            total_filas_detalle += 1

    anchos2 = [8, 38, 12, 12, 22, 9, 10] + [9] * len(cap_headers) + [14, 55]
    for c, ancho in enumerate(anchos2, start=1):
        hoja2.column_dimensions[get_column_letter(c)].width = min(ancho, 56)
    hoja2.freeze_panes = "A2"
    hoja2.auto_filter.ref = f"A1:{get_column_letter(len(encabezados2))}{max(total_filas_detalle + 1, 2)}"

    # ── Hoja Cambios (encabezados vacíos para auditoría) ──────────────────
    hoja3 = libro.create_sheet("Cambios")
    for c, titulo in enumerate(["File", "Colegio", "Hotel", "Tipo", "Qué cambió", "Antes", "Después"], start=1):
        celda = hoja3.cell(row=1, column=c, value=titulo)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
        celda.border = BORDE_GRILLA
    avisos_rows = 0
    for dist in reporte.distribuciones:
        if not dist.avisos:
            continue
        avisos_rows += 1
        idx = reporte.distribuciones.index(dist) + 1
        aviso_texto = "; ".join(dist.avisos[:3])
        fila_valores = [str(idx), dist.grupo.nombre, hotel.nombre, "Aviso", aviso_texto, "", ""]
        for c, valor in enumerate(fila_valores, start=1):
            celda = hoja3.cell(row=1 + avisos_rows, column=c, value=valor)
            celda.border = BORDE_GRILLA
    for c, w in enumerate([8, 38, 12, 10, 55, 20, 20], start=1):
        hoja3.column_dimensions[get_column_letter(c)].width = w

    return libro


def _rango_grupos(grupos) -> tuple:
    con_fechas = [g for g in grupos if g.tiene_fechas]
    if not con_fechas:
        return None, None
    inicios = [g.fecha_ingreso for g in con_fechas]
    egresos = [g.fecha_egreso for g in con_fechas]
    return min(inicios), max(egresos) - timedelta(days=1)


def _hoja_limites(libro: Workbook, hotel, grupos, reporte, desde, hasta, titulo_empresa: str = "Distribución Hotelera"):
    hoja = libro.create_sheet("Límites por noche")
    dias = _dias_rango(desde, hasta)
    dist_por_grupo = {}
    if reporte is not None:
        for d in reporte.distribuciones:
            dist_por_grupo[id(d.grupo)] = d

    encabezados = ["Fecha", "Hotel", "Límite personas", "Ocupadas", "Disponibles",
                   "Límite aptos", "Ocupados", "Disponibles"]
    for c, txt in enumerate(encabezados, start=1):
        celda = hoja.cell(row=1, column=c, value=txt)
        celda.font = Font(bold=True, color="FFFFFF")
        celda.fill = PatternFill("solid", fgColor="374151")
        celda.alignment = Alignment(horizontal="center")
        celda.border = BORDE_GRILLA

    lim_camas = [_camas_limite_por_dia(hotel, d) for d in dias]
    lim_aptos = [len(hotel.habitaciones)] * len(dias)
    oc_camas = [0] * len(dias)
    oc_aptos = [0] * len(dias)
    for g in grupos:
        if g.total_pax() < 1 or not g.tiene_fechas:
            continue
        dist = dist_por_grupo.get(id(g))
        aptos = dist.habitaciones_usadas if dist else 0
        for i, d in enumerate(dias):
            if g.fecha_ingreso <= d < g.fecha_egreso:
                oc_camas[i] += g.total_pax()
                oc_aptos[i] += aptos

    neg_fondo = PatternFill("solid", fgColor="FCA5A5")
    neg_font = Font(color="991B1B", bold=True)
    ok_fondo = PatternFill("solid", fgColor="D1FAE5")
    ok_font = Font(color="065F46", bold=True)
    fila = 2
    for i, d in enumerate(dias):
        disp_camas = lim_camas[i] - oc_camas[i]
        disp_aptos = lim_aptos[i] - oc_aptos[i]
        valores = [d, hotel.nombre, lim_camas[i], oc_camas[i], disp_camas,
                   lim_aptos[i], oc_aptos[i], disp_aptos]
        for c, v in enumerate(valores, start=1):
            celda = hoja.cell(row=fila, column=c, value=v)
            celda.border = BORDE_GRILLA
            celda.alignment = Alignment(horizontal="center" if c > 1 else "left")
        for col in (5, 8):
            v = hoja.cell(row=fila, column=col).value
            if isinstance(v, int) and v < 0:
                hoja.cell(row=fila, column=col).fill = neg_fondo
                hoja.cell(row=fila, column=col).font = neg_font
            else:
                hoja.cell(row=fila, column=col).fill = ok_fondo
                hoja.cell(row=fila, column=col).font = ok_font
        fila += 1
    if any(lim_camas[i] - oc_camas[i] < 0 for i in range(len(dias))):
        hoja.cell(row=fila, column=1, value="⚠ SOBREVENTA").font = Font(bold=True, color="DC2626")
        fila += 1

    anchos = [12, 20, 14, 10, 12, 12, 10, 12]
    for c, ancho in enumerate(anchos, start=1):
        hoja.column_dimensions[get_column_letter(c)].width = ancho
    hoja.freeze_panes = "C2"
    return hoja


def exportar_rooming_excel(
    ruta: str,
    reporte: ReporteGeneral | None = None,
    destino: str = "",
    titulo_empresa: str = "Distribución Hotelera",
    desde=None,
    hasta=None,
    hotel=None,
    grupos=None,
) -> int:
    """Exporta Excel Rooming. Si reporte es None, requiere hotel+grupos para calcularlo."""
    if reporte is None and (hotel is None or not grupos):
        raise ErrorDistribucion("Se necesita reporte o hotel+grupos para exportar.")
    if reporte is not None:
        hotel = reporte.hotel
        grupos = [d.grupo for d in reporte.distribuciones]
    else:
        reporte = distribuir(hotel, list(grupos), OpcionesMotor())

    libro = _libro_rooming(hotel, grupos, reporte, destino, titulo_empresa)
    desde_h, hasta_h = desde, hasta
    if desde_h is None:
        desde_h, hasta_h = _rango_grupos(grupos)
    if desde_h is not None and hasta_h is not None:
        _hoja_limites(libro, hotel, grupos, reporte, desde_h, hasta_h, titulo_empresa)

    try:
        libro.save(ruta)
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo escribir el Excel Rooming '{ruta}': {exc}")
    return len(reporte.distribuciones)


def exportar_rooming_hoteles(
    carpeta: str,
    hoteles: list,
    grupos_por_hotel: list,
    reportes: list | None = None,
    desde=None,
    hasta=None,
    destino: str = "",
    titulo_empresa: str = "Distribución Hotelera",
) -> list[str]:
    """Genera UN archivo Excel Rooming por hotel dentro de `carpeta`.

    Devuelve la lista de rutas generadas. Los archivos incluyen la hoja
    'Límites por noche' con rangos por fecha y sobreventa marcada en rojo.
    """
    if reportes is None:
        reportes = [None] * len(hoteles)
    if not os.path.isdir(carpeta):
        os.makedirs(carpeta, exist_ok=True)
    rutas: list[str] = []
    if desde is None or hasta is None:
        desde, hasta = None, None
        for grupos in grupos_por_hotel:
            a, b = _rango_grupos(grupos)
            if a is not None:
                desde = min(desde, a) if desde else a
                hasta = max(hasta, b) if hasta else b
    for indice, hotel in enumerate(hoteles):
        grupos = grupos_por_hotel[indice]
        if not grupos:
            continue
        reporte = reportes[indice]
        if reporte is None:
            try:
                reporte = distribuir(hotel, list(grupos), OpcionesMotor())
            except (ErrorDistribucion, ValueError, Exception):
                continue
        nombre = re.sub(r'[\\/:*?"<>|]', "_", hotel.nombre) or f"Hotel{indice + 1}"
        ruta = os.path.join(carpeta, f"Rooming_{nombre}.xlsx")
        exportar_rooming_excel(ruta, reporte=reporte, destino=destino,
                               titulo_empresa=titulo_empresa,
                               hotel=hotel, grupos=grupos,
                               desde=desde, hasta=hasta)
        rutas.append(ruta)
    return rutas
