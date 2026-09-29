from __future__ import annotations

import os
import random
import sys
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from modelos import ConfigGrupo, Grupo, Habitacion, Hotel, Pax, Sexo, TipoPax

FECHA_BASE = date(2026, 6, 1)


def mk_hotel(capacidades: list[int], desde: date | None = None, hasta: date | None = None) -> Hotel:
    habitaciones = [
        Habitacion(
            numero=f"{100 + i}",
            capacidad=cap,
            disponible_desde=desde,
            disponible_hasta=hasta,
        )
        for i, cap in enumerate(capacidades)
    ]
    return Hotel(nombre="Hotel Prueba", habitaciones=habitaciones)


def mk_pax(nombre: str, tipo: TipoPax = TipoPax.ADULTO, sexo: Sexo = Sexo.X, edad=None, **kwargs) -> Pax:
    return Pax(nombre=nombre, tipo=tipo, sexo=sexo, edad=edad, **kwargs)


def mk_grupo(
    nombre: str,
    pax=None,
    dia_inicio: int | None = 1,
    noches: int = 3,
    config: ConfigGrupo | None = None,
    base: date = FECHA_BASE,
    empresa: str = "",
    provisional: bool = False,
    confirmado: bool = False,
) -> Grupo:
    ingreso = base + timedelta(days=dia_inicio) if dia_inicio is not None else None
    egreso = (base + timedelta(days=dia_inicio + noches)
              if dia_inicio is not None else None)
    return Grupo(
        nombre=nombre,
        empresa=empresa,
        fecha_ingreso=ingreso,
        fecha_egreso=egreso,
        provisional=provisional,
        confirmado=confirmado,
        pax=list(pax or []),
        config=config or ConfigGrupo(),
    )


def grupo_aleatorio(
    rng: random.Random,
    nombre: str,
    cantidad: int,
    config: ConfigGrupo | None = None,
    con_familias: bool = True,
    con_parejas: bool = True,
    dia_inicio: int | None = None,
    noches: int | None = None,
) -> Grupo:
    config = config or ConfigGrupo()
    pax: list[Pax] = []
    contador = {"n": 0}

    def nuevo(prefijo: str) -> str:
        contador["n"] += 1
        return f"{prefijo}-{contador['n']:03d}"

    restantes = cantidad
    while restantes > 0:
        eleccion = rng.random()
        if con_familias and eleccion < 0.15 and restantes >= 3:
            tamano = rng.randint(2, min(4, restantes))
            id_familia = nuevo("FAM")
            adultos = max(1, min(tamano - 1, rng.randint(1, 2)))
            for j in range(tamano):
                es_menor = j >= adultos
                sexo = rng.choice([Sexo.M, Sexo.F])
                if es_menor:
                    pax.append(Pax(nombre=nuevo("MEN"), tipo=TipoPax.NINO, sexo=sexo,
                                   edad=rng.randint(4, 15), familia=id_familia))
                else:
                    pax.append(Pax(nombre=nuevo("ADU"), sexo=sexo,
                                   edad=rng.randint(25, 60), familia=id_familia))
            restantes -= tamano
        elif con_parejas and eleccion < 0.30 and restantes >= 2:
            a = nuevo("PAR")
            b = nuevo("PAR")
            sexo_a = rng.choice([Sexo.M, Sexo.F])
            sexo_b = Sexo.F if sexo_a == Sexo.M else Sexo.M
            pax.append(Pax(nombre=a, sexo=sexo_a, pareja_de=b))
            pax.append(Pax(nombre=b, sexo=sexo_b, pareja_de=a))
            restantes -= 2
        else:
            dado = rng.random()
            sexo = rng.choice([Sexo.M, Sexo.F])
            if dado < 0.08:
                pax.append(Pax(nombre=nuevo("VIP"), sexo=sexo, requiere_single=True))
            elif dado < 0.16:
                pax.append(Pax(nombre=nuevo("GUI"), tipo=TipoPax.GUIA, sexo=sexo))
            elif dado < 0.24:
                pax.append(Pax(nombre=nuevo("CHO"), tipo=TipoPax.CHOFER, sexo=sexo))
            elif dado < 0.34:
                pax.append(Pax(nombre=nuevo("EST"), tipo=TipoPax.ESTUDIANTE, sexo=sexo,
                               edad=rng.randint(12, 19)))
            elif dado < 0.40:
                id_familia = nuevo("FAMX")
                sexo_ninio = rng.choice([Sexo.M, Sexo.F])
                pax.append(Pax(nombre=nuevo("NIN"), tipo=TipoPax.NINO, sexo=sexo_ninio,
                               edad=rng.randint(4, 12), familia=id_familia))
                if restantes >= 2:
                    sexo_adulto = Sexo.F if sexo_ninio == Sexo.M else Sexo.M
                    pax.append(Pax(nombre=nuevo("ADU"), sexo=sexo_adulto,
                                   edad=rng.randint(30, 55), familia=id_familia))
                    restantes -= 1
            else:
                pax.append(Pax(nombre=nuevo("ADU"), sexo=sexo))
            restantes -= 1
    inicio = dia_inicio if dia_inicio is not None else rng.randint(0, 10)
    duracion = noches if noches is not None else rng.randint(2, 7)

    if config.ninos_con_adulto_obligatorio:
        familias_con_menores: dict[str, int] = {}
        for p in pax:
            if p.familia and p.es_menor():
                familias_con_menores[p.familia] = familias_con_menores.get(p.familia, 0) + 1
        adultos_libres = [p for p in pax if not p.familia and p.es_adulto()]
        for id_familia, menores in familias_con_menores.items():
            tiene_adulto = any(p.familia == id_familia and p.es_adulto() for p in pax)
            if not tiene_adulto:
                if adultos_libres:
                    adoptante = adultos_libres.pop()
                    adoptante.familia = id_familia
                else:
                    raise ValueError(f"fabrica: familia {id_familia} sin adulto disponible")

    return mk_grupo(nombre, pax, dia_inicio=inicio, noches=duracion, config=config)
