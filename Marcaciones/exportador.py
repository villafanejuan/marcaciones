from __future__ import annotations

import csv
import os

from modelos import ErrorDistribucion
from motor import DistribucionGrupo, ReporteGeneral


def exportar_csv(reporte: ReporteGeneral, ruta: str) -> int:
    try:
        with open(ruta, "w", newline="", encoding="utf-8-sig") as archivo:
            escritor = csv.writer(archivo)
            escritor.writerow(
                [
                    "Grupo",
                    "Ingreso",
                    "Egreso",
                    "Noches",
                    "Habitacion",
                    "Capacidad",
                    "Ocupacion",
                    "Nombre",
                    "Tipo",
                    "Sexo",
                    "Edad",
                    "Familia",
                    "ParejaDe",
                    "Single",
                    "Notas",
                ]
            )
            filas = 0
            for dist in reporte.distribuciones:
                grupo = dist.grupo
                for asig in dist.asignaciones:
                    hab = asig.habitacion
                    for atomo in asig.atomos:
                        for p in atomo.miembros:
                            escritor.writerow(
                                [
                                    grupo.nombre,
                                    grupo.fecha_ingreso.isoformat(),
                                    grupo.fecha_egreso.isoformat(),
                                    grupo.noches(),
                                    hab.numero,
                                    hab.capacidad,
                                    asig.carga,
                                    p.nombre,
                                    p.tipo.value,
                                    p.sexo.value,
                                    p.edad if p.edad is not None else "",
                                    p.familia or "",
                                    p.pareja_de or "",
                                    "SI" if (p.requiere_single or atomo.requiere_single) else "",
                                    p.notas,
                                ]
                            )
                            filas += 1
                for p in dist.faltantes:
                    escritor.writerow(
                        [
                            grupo.nombre,
                            grupo.fecha_ingreso.isoformat(),
                            grupo.fecha_egreso.isoformat(),
                            grupo.noches(),
                            "SIN HABITACION",
                            "",
                            "",
                            p.nombre,
                            p.tipo.value,
                            p.sexo.value,
                            p.edad if p.edad is not None else "",
                            p.familia or "",
                            p.pareja_de or "",
                            "SI" if p.requiere_single else "",
                            p.notas,
                        ]
                    )
                    filas += 1
        return filas
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo escribir el CSV '{ruta}': {exc}")


def exportar_resumen_csv(reporte: ReporteGeneral, ruta: str) -> int:
    try:
        with open(ruta, "w", newline="", encoding="utf-8-sig") as archivo:
            escritor = csv.writer(archivo)
            escritor.writerow(
                [
                    "Grupo",
                    "Ingreso",
                    "Egreso",
                    "Noches",
                    "PaxTotal",
                    "PaxAlojados",
                    "Faltantes",
                    "HabitacionesUsadas",
                    "CamasOcupadas",
                    "CamasTotales",
                    "Desperdicio",
                    "Estrategia",
                    "Optimo",
                    "Completo",
                ]
            )
            for dist in reporte.distribuciones:
                g = dist.grupo
                camas_totales = sum(a.habitacion.capacidad for a in dist.asignaciones)
                camas_ocupadas = sum(a.carga for a in dist.asignaciones)
                escritor.writerow(
                    [
                        g.nombre,
                        g.fecha_ingreso.isoformat(),
                        g.fecha_egreso.isoformat(),
                        g.noches(),
                        g.total_pax(),
                        dist.pax_alojados,
                        len(dist.faltantes),
                        dist.habitaciones_usadas,
                        camas_ocupadas,
                        camas_totales,
                        camas_totales - camas_ocupadas,
                        dist.estrategia_usada,
                        "SI" if dist.optimo else "NO",
                        "SI" if dist.completo else "NO",
                    ]
                )
            return len(reporte.distribuciones)
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo escribir el resumen '{ruta}': {exc}")
