from __future__ import annotations

from datetime import date

from informes import renderizar_reporte
from modelos import (
    ConfigGrupo,
    Grupo,
    Habitacion,
    Hotel,
    Pax,
    Sexo,
    TipoPax,
)
from motor import OpcionesMotor, distribuir


def _construir_habitaciones() -> list[Habitacion]:
    salida: list[Habitacion] = []

    def lote(inicial: int, cantidad: int, capacidad: int) -> None:
        for i in range(cantidad):
            salida.append(Habitacion(numero=f"{inicial + i}", capacidad=capacidad))

    lote(101, 12, 1)
    lote(201, 15, 2)
    lote(301, 8, 3)
    lote(401, 8, 4)
    return salida


def construir_datos_demo() -> tuple[Hotel, list[Grupo]]:
    hotel = Hotel(nombre="Cabanas del Sol", habitaciones=_construir_habitaciones())
    habitacion_205 = hotel.habitacion("205")
    habitacion_205.agregar_bloqueo(date(2026, 2, 20), date(2026, 3, 20), "Mantenimiento de calefaccion")

    andina = Grupo(
        nombre="Bachilleres 3° A",
        empresa="Empresa Andina",
        fecha_ingreso=date(2026, 3, 1),
        fecha_egreso=date(2026, 3, 6),
        provisional=True,
        pax=[
            Pax(nombre="Sosa Martin", tipo=TipoPax.ADULTO, sexo=Sexo.M,
                requiere_single=True, notas="Gerente"),
            Pax(nombre="Lopez Carla", tipo=TipoPax.ADULTO, sexo=Sexo.F, pareja_de="Diaz Pablo"),
            Pax(nombre="Diaz Pablo", tipo=TipoPax.ADULTO, sexo=Sexo.M, pareja_de="Lopez Carla"),
            Pax(nombre="Ferreyra Julia", tipo=TipoPax.ADULTO, sexo=Sexo.F, pareja_de="Molina Raul"),
            Pax(nombre="Molina Raul", tipo=TipoPax.ADULTO, sexo=Sexo.M, pareja_de="Ferreyra Julia"),
            Pax(nombre="Acosta Bruno", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Rios Damian", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Paredes Nico", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Gomez Ivan", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Tevez Mauro", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Blanco Silvina", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Noble Micaela", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Cruz Veronica", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Aguirre Lorena", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Suarez Paula", tipo=TipoPax.ADULTO, sexo=Sexo.F),
        ],
    )

    colegio = Grupo(
        nombre="Primaria 5° B",
        empresa="Colegio Norte",
        fecha_ingreso=date(2026, 3, 6),
        fecha_egreso=date(2026, 3, 10),
        confirmado=True,
        config=ConfigGrupo(estudiantes_por_habitacion=4, estudiantes_estricto=True),
        pax=(
            [
                Pax(nombre=f"Alumno-M-{i:02d}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.M, edad=17)
                for i in range(1, 17)
            ]
            + [
                Pax(nombre=f"Alumna-F-{i:02d}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.F, edad=17)
                for i in range(1, 9)
            ]
            + [
                Pax(nombre="Vega Laura", tipo=TipoPax.GUIA, sexo=Sexo.F),
                Pax(nombre="Ramos Diego", tipo=TipoPax.GUIA, sexo=Sexo.M),
                Pax(nombre="Ibarra Jose", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            ]
        ),
    )

    familias = Grupo(
        nombre="Escuela de Formación",
        empresa="Instituto Sur",
        fecha_ingreso=date(2026, 3, 8),
        fecha_egreso=date(2026, 3, 12),
        pax=[
            Pax(nombre="Garcia Hugo", tipo=TipoPax.ADULTO, sexo=Sexo.M, familia="GARCIA"),
            Pax(nombre="Garcia Maria", tipo=TipoPax.ADULTO, sexo=Sexo.F, familia="GARCIA"),
            Pax(nombre="Garcia Tomas", tipo=TipoPax.NINO, sexo=Sexo.M, edad=9, familia="GARCIA"),
            Pax(nombre="Garcia Sofia", tipo=TipoPax.NINO, sexo=Sexo.F, edad=6, familia="GARCIA"),
            Pax(nombre="Perez Ana", tipo=TipoPax.ADULTO, sexo=Sexo.F, familia="PEREZ"),
            Pax(nombre="Perez Rosa", tipo=TipoPax.ADULTO, sexo=Sexo.F, familia="PEREZ", notas="Abuela"),
            Pax(nombre="Perez Emma", tipo=TipoPax.NINO, sexo=Sexo.F, edad=7, familia="PEREZ"),
        ],
    )

    pendiente = Grupo(
        nombre="Turno Tarde",
        empresa="Colegio Norte",
        pax=[
            Pax(nombre=f"Pendiente-{i:02d}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.M, edad=16)
            for i in range(1, 11)
        ],
        provisional=True,
    )

    return hotel, [andina, colegio, familias, pendiente]


def main() -> None:
    hotel, grupos = construir_datos_demo()
    reporte = distribuir(hotel, grupos, OpcionesMotor())
    print(renderizar_reporte(reporte))


if __name__ == "__main__":
    main()
