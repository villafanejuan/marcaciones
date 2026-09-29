from __future__ import annotations

import os
import sys
import unittest
from collections import Counter
from datetime import timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fabrica import FECHA_BASE, mk_grupo, mk_hotel, mk_pax  # noqa: E402
from modelos import (  # noqa: E402
    ConfigGrupo,
    ErrorDistribucion,
    Grupo,
    Pax,
    Sexo,
    TipoPax,
)
from motor import (  # noqa: E402
    AsignacionHabitacion,
    DistribucionGrupo,
    OpcionesMotor,
    ReporteGeneral,
    distribuir,
    verificar_reporte,
)


class TestMotor(unittest.TestCase):
    def test_grupos_mismas_fechas_habitaciones_distintas(self):
        hotel = mk_hotel([1, 1])
        g1 = mk_grupo("A", [mk_pax("A1", sexo=Sexo.M)])
        g2 = mk_grupo("B", [mk_pax("B1", sexo=Sexo.M)])
        reporte = distribuir(hotel, [g1, g2])
        self.assertTrue(reporte.todo_alojado)
        numeros = {
            a.habitacion.numero
            for d in reporte.distribuciones
            for a in d.asignaciones
        }
        self.assertEqual(len(numeros), 2)

    def test_checkout_e_ingreso_mismo_dia_comparten(self):
        hotel = mk_hotel([1])
        g1 = mk_grupo("A", [mk_pax("A1")], dia_inicio=1, noches=3)
        g2 = mk_grupo("B", [mk_pax("B1")], dia_inicio=4, noches=3)
        reporte = distribuir(hotel, [g1, g2])
        self.assertTrue(reporte.todo_alojado)
        numeros = [
            a.habitacion.numero
            for d in reporte.distribuciones
            for a in d.asignaciones
        ]
        self.assertEqual(len(numeros), 2)
        self.assertEqual(len(set(numeros)), 1)

    def test_bloqueo_previo_excluye_y_calcula_deficit(self):
        hotel = mk_hotel([2, 2])
        hotel.habitaciones[0].agregar_bloqueo(FECHA_BASE, FECHA_BASE + timedelta(days=30), "previa")
        grupo = mk_grupo("G", [mk_pax(f"P{i}") for i in range(3)],
                         config=ConfigGrupo(separar_sexos=False))
        reporte = distribuir(hotel, [grupo])
        dist = reporte.distribuciones[0]
        self.assertEqual(len(dist.faltantes), 1)
        self.assertEqual(dist.deficit.get(2, 0), 1)
        self.assertEqual(len(dist.asignaciones), 1)

    def test_ventana_de_disponibilidad_excluye(self):
        desde = FECHA_BASE + timedelta(days=20)
        hasta = FECHA_BASE + timedelta(days=60)
        hotel = mk_hotel([2], desde=desde, hasta=hasta)
        grupo = mk_grupo("G", [mk_pax("X")], dia_inicio=1, noches=2)
        reporte = distribuir(hotel, [grupo])
        dist = reporte.distribuciones[0]
        self.assertEqual(len(dist.faltantes), 1)
        self.assertFalse(dist.asignaciones)

    def test_capacidad_efectiva_baja(self):
        hotel = mk_hotel([4, 4])
        grupo = mk_grupo("G", [mk_pax(f"P{i}") for i in range(6)],
                         config=ConfigGrupo(separar_sexos=False))
        opciones = OpcionesMotor(capacidades_efectivas={4: 3})
        reporte = distribuir(hotel, [grupo], opciones)
        self.assertTrue(reporte.todo_alojado)
        for asig in reporte.distribuciones[0].asignaciones:
            self.assertLessEqual(asig.carga, 3)

    def test_capacidad_efectiva_suba_error(self):
        hotel = mk_hotel([4])
        grupo = mk_grupo("G", [mk_pax("P")])
        with self.assertRaises(ErrorDistribucion):
            distribuir(hotel, [grupo], OpcionesMotor(capacidades_efectivas={4: 5}))

    def test_reserva_previa_restringe_segunda_corrida(self):
        hotel = mk_hotel([2, 2])
        g1 = mk_grupo(
            "A",
            [mk_pax("a"), mk_pax("b"), mk_pax("c"), mk_pax("d")],
            dia_inicio=1, noches=2,
            config=ConfigGrupo(separar_sexos=False),
        )
        primer = distribuir(hotel, [g1])
        self.assertTrue(primer.todo_alojado)
        for asig in primer.distribuciones[0].asignaciones:
            asig.habitacion.agregar_bloqueo(g1.fecha_ingreso, g1.fecha_egreso, titulo="A")
        g2 = mk_grupo("B", [mk_pax("c"), mk_pax("d")], dia_inicio=1, noches=2,
                      config=ConfigGrupo(separar_sexos=False))
        segundo = distribuir(hotel, [g2])
        dist = segundo.distribuciones[0]
        self.assertEqual(len(dist.faltantes), 2)
        self.assertFalse(dist.asignaciones)

    def test_estudiantes_cupo_fijo_estricto_completo(self):
        hotel = mk_hotel([4] * 6 + [1] * 3)
        pax = (
            [mk_pax(f"M-{i}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.M, edad=17) for i in range(16)]
            + [mk_pax(f"F-{i}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.F, edad=17) for i in range(8)]
            + [mk_pax("Guia Uno", tipo=TipoPax.GUIA, sexo=Sexo.F),
               mk_pax("Chofer Uno", tipo=TipoPax.CHOFER, sexo=Sexo.M)]
        )
        config = ConfigGrupo(estudiantes_por_habitacion=4, estudiantes_estricto=True)
        grupo = mk_grupo("Colegio", pax, config=config)
        reporte = distribuir(hotel, [grupo])
        self.assertTrue(reporte.todo_alojado)
        dist = reporte.distribuciones[0]
        self.assertEqual(len(dist.asignaciones), 8)
        for asig in dist.asignaciones:
            tipos = {m.tipo for a in asig.atomos for m in a.miembros}
            if tipos == {TipoPax.ESTUDIANTE}:
                self.assertEqual(asig.carga, 4)

    def test_estudiantes_estricto_no_divisible_error(self):
        hotel = mk_hotel([4] * 3)
        pax = [mk_pax(f"E-{i}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.M, edad=17) for i in range(5)]
        config = ConfigGrupo(estudiantes_por_habitacion=4, estudiantes_estricto=True)
        grupo = mk_grupo("Colegio Chico", pax, config=config)
        with self.assertRaises(ErrorDistribucion) as contexto:
            distribuir(hotel, [grupo])
        self.assertIn("estricto", str(contexto.exception))

    def test_estudiantes_sin_habitacion_de_capacidad_error(self):
        hotel = mk_hotel([2, 2])
        pax = [mk_pax(f"E-{i}", tipo=TipoPax.ESTUDIANTE, sexo=Sexo.X, edad=17) for i in range(4)]
        config = ConfigGrupo(estudiantes_por_habitacion=4)
        grupo = mk_grupo("C", pax, config=config)
        with self.assertRaises(ErrorDistribucion) as contexto:
            distribuir(hotel, [grupo])
        self.assertIn("capacidad", str(contexto.exception))

    def test_familias_juntas_aunque_se_solen_grupos(self):
        hotel = mk_hotel([4, 2, 2])
        otro = mk_grupo("Otros", [mk_pax(f"O-{i}", sexo=Sexo.M) for i in range(4)],
                        dia_inicio=1, noches=5)
        familia_pax = [
            mk_pax("Papa", sexo=Sexo.M, familia="LOPEZ"),
            mk_pax("Mama", sexo=Sexo.F, familia="LOPEZ"),
            mk_pax("Hijo", tipo=TipoPax.NINO, edad=8, sexo=Sexo.M, familia="LOPEZ"),
        ]
        familia = mk_grupo("Familia Lopez", familia_pax, dia_inicio=1, noches=5)
        reporte = distribuir(hotel, [otro, familia])
        self.assertTrue(reporte.todo_alojado)
        habitacion_familia = None
        for d in reporte.distribuciones:
            if d.grupo.nombre == "Familia Lopez":
                self.assertEqual(len(d.asignaciones), 1)
                habitacion_familia = d.asignaciones[0].habitacion
                nombres = set(d.asignaciones[0].nombres_pax)
                self.assertEqual(nombres, {"Papa", "Mama", "Hijo"})
        self.assertIsNotNone(habitacion_familia)

    def test_ninos_maximo_por_habitacion(self):
        hotel = mk_hotel([4, 2, 2, 1])
        pax = [
            mk_pax("A1", sexo=Sexo.M, familia="F1"),
            mk_pax("M1", tipo=TipoPax.NINO, edad=7, sexo=Sexo.M, familia="F1"),
            mk_pax("A2", sexo=Sexo.M, familia="F2"),
            mk_pax("M2", tipo=TipoPax.NINO, edad=8, sexo=Sexo.M, familia="F2"),
        ]
        config = ConfigGrupo(max_ninos_por_habitacion=1)
        grupo = mk_grupo("G", pax, config=config)
        reporte = distribuir(hotel, [grupo])
        self.assertTrue(reporte.todo_alojado)
        for asig in reporte.distribuciones[0].asignaciones:
            menores = sum(1 for m in asig.nombres_pax if m.startswith("M"))
            self.assertLessEqual(menores, 1)

    def test_fecha_invalida_error(self):
        hotel = mk_hotel([2])
        grupo = Grupo(nombre="Malas fechas",
                      fecha_ingreso=FECHA_BASE + timedelta(days=5),
                      fecha_egreso=FECHA_BASE + timedelta(days=5),
                      pax=[mk_pax("X")])
        with self.assertRaises(ErrorDistribucion):
            distribuir(hotel, [grupo])

    def test_grupo_vacio_error(self):
        hotel = mk_hotel([2])
        with self.assertRaises(ErrorDistribucion):
            distribuir(hotel, [mk_grupo("Vacio", [])])

    def test_hotel_duplicado_error(self):
        hotel = mk_hotel([2, 2])
        hotel.habitaciones[1].numero = "100"
        with self.assertRaises(ErrorDistribucion):
            distribuir(hotel, [])

    def test_verificador_detecta_sobrecarga_manipulada(self):
        hotel = mk_hotel([2])
        grupo = mk_grupo("G", [mk_pax("p1"), mk_pax("p2")],
                         config=ConfigGrupo(separar_sexos=False))
        hab = hotel.habitaciones[0]
        hab.capacidad = 1
        atomo_falso = None
        from reglas import atomizar
        atomos, _ = atomizar(grupo)
        dist = DistribucionGrupo(
            grupo=grupo,
            asignaciones=[AsignacionHabitacion(habitacion=hab, atomos=list(atomos))],
        )
        reporte = ReporteGeneral(hotel=hotel, distribuciones=[dist])
        violaciones = verificar_reporte(reporte)
        self.assertTrue(any("capacidad" in v.lower() for v in violaciones))

    def test_contabilidad_nocturna(self):
        hotel = mk_hotel([1])
        g1 = mk_grupo("A", [mk_pax("A1")], dia_inicio=0, noches=2)
        g2 = mk_grupo("B", [mk_pax("B1")], dia_inicio=2, noches=2)
        reporte = distribuir(hotel, [g1, g2])
        self.assertEqual(sum(reporte.pax_por_noche.values()), 4)
        self.assertEqual(min(reporte.pax_por_noche.values()), 1)

    def test_grupo_sin_fechas_no_se_distribuye(self):
        hotel = mk_hotel([4])
        g1 = mk_grupo("Con fecha", [mk_pax("A")], dia_inicio=1, noches=2,
                      config=ConfigGrupo(separar_sexos=False))
        g2 = mk_grupo("Sin fecha", [mk_pax("B")], dia_inicio=None,
                      config=ConfigGrupo(separar_sexos=False))
        self.assertFalse(g2.tiene_fechas)
        reporte = distribuir(hotel, [g1, g2])
        nombres = [d.grupo.nombre for d in reporte.distribuciones]
        self.assertNotIn("Sin fecha", nombres)
        self.assertTrue(any("Sin fecha" in a for a in reporte.avisos))
        self.assertTrue(reporte.todo_alojado)

    def test_confirmado_no_se_mueve_con_previo(self):
        hotel = mk_hotel([2, 2, 2, 2])
        config = ConfigGrupo(separar_sexos=False)
        g1 = mk_grupo("KEA", [mk_pax("1"), mk_pax("2"), mk_pax("3"), mk_pax("4")],
                      dia_inicio=1, noches=4, config=config, provisional=True)
        g2 = mk_grupo("KEB", [mk_pax("5"), mk_pax("6"), mk_pax("7"), mk_pax("8")],
                      dia_inicio=2, noches=4, config=config, confirmado=True)
        primero = distribuir(hotel, [g1, g2])
        self.assertTrue(primero.todo_alojado)

        def habs_de(r, nombre):
            for d in r.distribuciones:
                if d.grupo.nombre == nombre:
                    return sorted(a.habitacion.numero for a in d.asignaciones)
            return None

        confirmado1 = habs_de(primero, "KEB")
        self.assertTrue(confirmado1)
        segundo = distribuir(hotel, [g1, g2], previo=primero, reoptimizar=False)
        self.assertTrue(segundo.todo_alojado)
        self.assertEqual(confirmado1, habs_de(segundo, "KEB"))
        provisorias = [a.habitacion.numero for d in segundo.distribuciones
                       if d.grupo.nombre == "KEA" for a in d.asignaciones]
        self.assertTrue(all(h not in set(confirmado1) for h in provisorias))


if __name__ == "__main__":
    unittest.main()
