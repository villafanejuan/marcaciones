from __future__ import annotations

import os
import sys
import unittest
from datetime import date

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fabrica import FECHA_BASE, mk_grupo, mk_pax  # noqa: E402
from modelos import ConfigGrupo, ErrorDistribucion, Sexo  # noqa: E402
from motor import (  # noqa: E402
    crear_grupo_generico,
    distribucion_estandar,
    distribuir,
    generar_hotel_estimado,
    sugerir_mezcla,
)
from reglas import atomizar  # noqa: E402


class TestDistribucionEstandar(unittest.TestCase):
    def test_ratio_respetado(self):
        caps = distribucion_estandar(55, 4, 3.5)
        self.assertEqual(len(caps), 55)
        self.assertTrue(all(1 <= c <= 4 for c in caps))
        total = sum(caps)
        self.assertAlmostEqual(total / 55, 3.5, delta=0.35)

    def test_capacidad_max_no_superada(self):
        caps = distribucion_estandar(10, 3, 2.9)
        self.assertTrue(all(c <= 3 for c in caps))

    def test_cantidad_minima_una_cama(self):
        caps = distribucion_estandar(8, 4, 0.5)
        self.assertTrue(all(c >= 1 for c in caps))

    def test_errores_entrada(self):
        with self.assertRaises(Exception):
            distribucion_estandar(0, 4, 3)
        with self.assertRaises(Exception):
            distribucion_estandar(5, 0, 3)


class TestGenerarHotelEstimado(unittest.TestCase):
    def test_hotel_generado_con_habitaciones(self):
        hotel = generar_hotel_estimado("Estimado", 20, 4, 3.2)
        self.assertEqual(len(hotel.habitaciones), 20)
        self.assertEqual(hotel.nombre, "Estimado")
        self.assertTrue(all(h.capacidad >= 1 for h in hotel.habitaciones))
        self.assertEqual(hotel.total_camas(), sum(h.capacidad for h in hotel.habitaciones))

    def test_hotel_sin_ratio_usa_capacidad_max(self):
        hotel = generar_hotel_estimado("X", 6, 2)
        self.assertEqual(hotel.total_camas(), 12)


class TestCrearGrupoGenerico(unittest.TestCase):
    def test_total_pax_respetado(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=3), 10)
        self.assertEqual(g.total_pax(), 10)
        atomos, avisos = atomizar(g)
        self.assertEqual(len(atomos), 10)
        self.assertTrue(all(len(a.miembros) == 1 for a in atomos))

    def test_varones_mujeres_indistintos(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=4),
                                 total_pax=10, varones=3, mujeres=4)
        varones = sum(1 for p in g.pax if p.sexo == Sexo.M)
        mujeres = sum(1 for p in g.pax if p.sexo == Sexo.F)
        indistintos = sum(1 for p in g.pax if p.sexo == Sexo.X)
        self.assertEqual(varones, 3)
        self.assertEqual(mujeres, 4)
        self.assertEqual(indistintos, 3)

    def test_mixed_sexos_config_activo(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=4),
                                 total_pax=6, varones=3, mujeres=3)
        self.assertTrue(g.config.separar_sexos)

    def test_total_cero_e_superado(self):
        with self.assertRaises(Exception):
            crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=2), 0)
        with self.assertRaises(Exception):
            crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=2), 5, varones=4, mujeres=4)


class TestSugerirMezcla(unittest.TestCase):
    def test_mezcla_optima_sin_tope(self):
        # 10 adultos indistintos con capacidades 1-4 -> mejor llenado en cuádruples
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=3), total_pax=10)
        atomos, _ = atomizar(g)
        sug = sugerir_mezcla(atomos, [1, 2, 3, 4], ConfigGrupo(separar_sexos=False))
        self.assertTrue(sug.completo)
        self.assertGreaterEqual(sug.habitaciones, 0)
        # 10 personas -> mínimo 3 habitaciones de 4 (12 camas, desperdicio 2)
        self.assertEqual(sug.habitaciones, 3)
        self.assertEqual(sug.camas, 12)
        self.assertEqual(sug.pax_alojados, 10)
        self.assertEqual(sug.desperdicio, 2)

    def test_con_tope_reporta_incompleto(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=3), total_pax=20)
        atomos, _ = atomizar(g)
        sug = sugerir_mezcla(atomos, [4] * 2 + [2] * 2, ConfigGrupo(separar_sexos=False),
                             total_habitaciones=4)
        self.assertFalse(sug.completo)
        self.assertEqual(sug.pax_alojados, sug.pax_total - sum(a.tamano for a in sug.pendientes))

    def test_sin_capacidades_error(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=2), total_pax=4)
        atomos, _ = atomizar(g)
        with self.assertRaises(Exception):
            sugerir_mezcla(atomos, [], ConfigGrupo())

    def test_atomos_demasiado_grandes(self):
        g = crear_grupo_generico("G", FECHA_BASE, FECHA_BASE.replace(day=4),
                                 total_pax=3, varones=1, mujeres=1)
        pax = g.pax[0]
        pax.familia = "F"
        g.pax[0].familia = "F"
        g.pax[1].familia = "F"
        g.pax[2].familia = "F"
        atomos, _ = atomizar(g)
        # familia de 3 personas necesita >= 3; con cant. max 2 no puede
        sug = sugerir_mezcla(atomos, [1, 2], ConfigGrupo(separar_sexos=False))
        self.assertFalse(sug.completo)


class TestE2EPlanificador(unittest.TestCase):
    def test_planificador_end_to_end(self):
        # Flujo completo: generar hotel estimado + grupo generico + distribuir
        hotel = generar_hotel_estimado("Hotel X", 40, 4, 3.5)
        g = crear_grupo_generico("Empresa", FECHA_BASE, FECHA_BASE.replace(day=6), total_pax=60)
        reporte = distribuir(hotel, [g])
        self.assertTrue(reporte.todo_alojado)


if __name__ == "__main__":
    unittest.main()