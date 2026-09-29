from __future__ import annotations

import os
import random
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fabrica import grupo_aleatorio  # noqa: E402
from modelos import ConfigGrupo, Pax, Sexo, TipoPax  # noqa: E402
from motor import OpcionesMotor, distribuir, verificar_reporte  # noqa: E402


def expandir(conteo_por_capacidad: dict[int, int]) -> list[int]:
    caps: list[int] = []
    for capacidad, cantidad in conteo_por_capacidad.items():
        caps.extend([capacidad] * cantidad)
    return caps


def grupo_colegio(nombre: str, varones: int, mujeres: int, guias: int, choferes: int,
                  dia_inicio: int, noches: int):
    from fabrica import mk_grupo
    pax = (
        [Pax(nombre=f"Est-M-{i:03d}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.M, edad=17)
         for i in range(varones)]
        + [Pax(nombre=f"Est-F-{i:03d}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.F, edad=17)
           for i in range(mujeres)]
        + [Pax(nombre=f"Guia-{i}", tipo=TipoPax.GUIA,
               sexo=Sexo.F if i % 2 == 0 else Sexo.M) for i in range(guias)]
        + [Pax(nombre=f"Chofer-{i}", tipo=TipoPax.CHOFER, sexo=Sexo.M) for i in range(choferes)]
    )
    config = ConfigGrupo(estudiantes_por_habitacion=4, estudiantes_estricto=True)
    return mk_grupo(nombre, pax=pax, dia_inicio=dia_inicio, noches=noches, config=config)


ESCENARIOS = [
    {
        "nombre": "S1 pequeno mixto",
        "hotel": {1: 10, 2: 12, 3: 6, 4: 6},
        "grupos": [("Agencia Sur", 20, 0, None), ("Corporativo Este", 15, 3, None)],
        "tiempo_max": 10.0,
    },
    {
        "nombre": "S2 colegio grande",
        "hotel": {1: 15, 2: 20, 3: 10, 4: 14},
        "grupos": [],
        "tiempo_max": 15.0,
        "especiales": [
            ("Colegio Norte", dict(varones=32, mujeres=16, guias=6, choferes=2,
                                   dia_inicio=1, noches=4)),
            ("Congreso Medico", dict(cantidad=40, dia_inicio=3, noches=4)),
        ],
    },
    {
        "nombre": "S3 hotel casi lleno",
        "hotel": {2: 20},
        "grupos": [("Lote A", 18, 0, 4), ("Lote B", 18, 2, 5)],
        "tiempo_max": 10.0,
        "relajado": True,
    },
    {
        "nombre": "S4 mega multi-empresa",
        "hotel": {1: 60, 2: 80, 3: 40, 4: 60},
        "grupos": [
            ("Empresa 1", 60, 0, None),
            ("Empresa 2", 70, 2, None),
            ("Empresa 3", 50, 5, None),
            ("Delegacion 4", 45, 8, None),
            ("Ultima Hora 5", 35, 11, None),
        ],
        "tiempo_max": 30.0,
    },
    {
        "nombre": "S5 estres reglas duras",
        "hotel": {1: 25, 2: 40, 3: 25, 4: 30},
        "grupos": [(f"Duro-{i}", 35, i * 2, None) for i in range(5)],
        "config": ConfigGrupo(separar_sexos=True, max_ninos_por_habitacion=2),
        "semilla_base": 900,
        "tiempo_max": 30.0,
    },
]


def correr_escenario(escenario: dict) -> dict:
    rng = random.Random(escenario.get("semilla_base", 2026))
    hotel_spec = expandir(escenario["hotel"])

    from fabrica import mk_hotel
    hotel = mk_hotel(hotel_spec)
    grupos = []
    config = escenario.get("config")
    for nombre, cantidad, dia_inicio, noches in escenario["grupos"]:
        grupos.append(
            grupo_aleatorio(rng, nombre, cantidad, config=config,
                            dia_inicio=dia_inicio, noches=noches)
        )
    for nombre, kwargs in escenario.get("especiales", []):
        if "cantidad" in kwargs:
            config_especial = ConfigGrupo(separar_sexos=False)
            g = grupo_aleatorio(rng, nombre, kwargs["cantidad"], config=config_especial,
                                dia_inicio=kwargs.get("dia_inicio"),
                                noches=kwargs.get("noches"))
        else:
            g = grupo_colegio(nombre, **kwargs)
        grupos.append(g)

    inicio = time.perf_counter()
    reporte = distribuir(hotel, grupos, OpcionesMotor(tiempo_max_por_grupo=8.0))
    tiempo = time.perf_counter() - inicio

    violaciones = verificar_reporte(reporte)
    total_pax = sum(g.total_pax() for g in grupos)
    alojados = sum(d.pax_alojados for d in reporte.distribuciones)

    return {
        "nombre": escenario["nombre"],
        "habitaciones": len(hotel.habitaciones),
        "camas": hotel.total_camas(),
        "grupos": len(grupos),
        "pax_total": total_pax,
        "pax_alojados": alojados,
        "pico_pax": max(reporte.pax_por_noche.values()) if reporte.pax_por_noche else 0,
        "faltantes": total_pax - alojados,
        "violaciones": violaciones,
        "tiempo": tiempo,
        "limite_tiempo": escenario["tiempo_max"],
        "todo_alojado": reporte.todo_alojado,
    }


def main() -> bool:
    print("=" * 88)
    print("PRUEBAS DE ESTRES - distintas escalas, solapamientos y reglas")
    print("=" * 88)
    exito_global = True
    for escenario in ESCENARIOS:
        resultado = correr_escenario(escenario)
        ok_sin_violaciones = not resultado["violaciones"]
        ok_tiempo = resultado["tiempo"] <= resultado["limite_tiempo"]
        exito = ok_sin_violaciones and ok_tiempo
        exito_global = exito_global and exito
        estado = "OK" if exito else "FALLO"
        print(
            f"[{estado}] {resultado['nombre']:<24} hab={resultado['habitaciones']:<4} "
            f"camas={resultado['camas']:<5} grupos={resultado['grupos']} "
            f"pax={resultado['pax_alojados']}/{resultado['pax_total']} "
            f"pico={resultado['pico_pax']:<3} "
            f"tiempo={resultado['tiempo']:.2f}s (max {resultado['limite_tiempo']}s)"
        )
        if resultado["violaciones"]:
            for v in resultado["violaciones"][:5]:
                print(f"        VIOLACION: {v}")
    print("=" * 88)
    print("ESTRES:", "TODO OK" if exito_global else "HUBO FALLOS")
    return exito_global


if __name__ == "__main__":
    sys.exit(0 if main() else 1)
