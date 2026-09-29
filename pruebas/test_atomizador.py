from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fabrica import mk_grupo, mk_pax  # noqa: E402
from modelos import ConfigGrupo, ErrorDistribucion, Pax, Sexo, TipoPax  # noqa: E402
from reglas import atomizar  # noqa: E402


def grupo(pax, **kwargs_config) -> object:
    config = ConfigGrupo(**kwargs_config)
    return mk_grupo("G", pax=pax, config=config)


class PruebasAtomizador(unittest.TestCase):
    def test_pareja_fusionada_en_un_atomo(self):
        g = grupo([
            Pax(nombre="Ana", sexo=Sexo.F, pareja_de="Beto"),
            Pax(nombre="Beto", sexo=Sexo.M, pareja_de="Ana"),
            Pax(nombre="Carl", sexo=Sexo.X),
        ])
        atomos, avisos = atomizar(g)
        self.assertEqual(len(atomos), 2)
        duos = [a for a in atomos if len(a.miembros) == 2]
        self.assertEqual(len(duos), 1)
        nombres = {m.nombre for m in duos[0].miembros}
        self.assertEqual(nombres, {"Ana", "Beto"})
        self.assertTrue(duos[0].es_mixto)

    def test_pareja_de_un_solo_lado_funciona(self):
        g = grupo([
            Pax(nombre="Ana", sexo=Sexo.F, pareja_de="Beto"),
            Pax(nombre="Beto", sexo=Sexo.M),
        ])
        atomos, _ = atomizar(g)
        self.assertEqual(len(atomos), 1)
        self.assertEqual(len(atomos[0].miembros), 2)

    def test_pareja_inexistente_error(self):
        g = grupo([Pax(nombre="Ana", sexo=Sexo.F, pareja_de="Fantasma")])
        with self.assertRaises(ErrorDistribucion):
            atomizar(g)

    def test_pareja_con_menor_error(self):
        g = grupo([
            Pax(nombre="Ana", sexo=Sexo.F, edad=15, pareja_de="Beto"),
            Pax(nombre="Beto", sexo=Sexo.M),
        ])
        with self.assertRaises(ErrorDistribucion):
            atomizar(g)

    def test_pareja_consigo_misma_error(self):
        g = grupo([Pax(nombre="Ana", sexo=Sexo.F, pareja_de="ana")])
        with self.assertRaises(ErrorDistribucion):
            atomizar(g)

    def test_familia_intacta_en_un_atomo(self):
        g = grupo([
            Pax(nombre="Papa", sexo=Sexo.M, familia="GARCIA"),
            Pax(nombre="Mama", sexo=Sexo.F, familia="GARCIA"),
            Pax(nombre="Hijo", tipo=TipoPax.NINO, edad=8, sexo=Sexo.M, familia="GARCIA"),
        ])
        atomos, _ = atomizar(g)
        self.assertEqual(len(atomos), 1)
        self.assertEqual(len(atomos[0].miembros), 3)
        self.assertTrue(atomos[0].tiene_menor)
        self.assertTrue(atomos[0].tiene_adulto)
        self.assertTrue(atomos[0].es_mixto)

    def test_nino_sin_adulto_error(self):
        g = grupo([Pax(nombre="Nini", tipo=TipoPax.NINO, edad=7)])
        with self.assertRaises(ErrorDistribucion) as contexto:
            atomizar(g)
        self.assertIn("adulto", str(contexto.exception).lower())

    def test_nino_sin_adulto_permitido_con_regla_apagada(self):
        g = grupo(
            [Pax(nombre="Nini", tipo=TipoPax.NINO, edad=7)],
            ninos_con_adulto_obligatorio=False,
        )
        atomos, _ = atomizar(g)
        self.assertEqual(len(atomos), 1)
        self.assertTrue(atomos[0].tiene_menor)

    def test_familia_de_menores_sin_adulto_error(self):
        g = grupo([
            Pax(nombre="Hijo1", tipo=TipoPax.NINO, edad=8, familia="X"),
            Pax(nombre="Hija2", tipo=TipoPax.NINO, edad=9, familia="X"),
        ])
        with self.assertRaises(ErrorDistribucion):
            atomizar(g)

    def test_guia_single_cuando_config_activa(self):
        g = grupo(
            [Pax(nombre="Guia Uno", tipo=TipoPax.GUIA, sexo=Sexo.F)],
            guias_choferes_single=True,
        )
        atomos, _ = atomizar(g)
        self.assertTrue(atomos[0].requiere_single)

    def test_guia_comparte_cuando_config_apagada(self):
        g = grupo(
            [Pax(nombre="Guia Uno", tipo=TipoPax.GUIA, sexo=Sexo.F)],
            guias_choferes_single=False,
        )
        atomos, _ = atomizar(g)
        self.assertFalse(atomos[0].requiere_single)

    def test_tripulacion_agrupa_por_defecto(self):
        """Desde el cambio a modo estudiantil la opcion de single es opt-in: dos choferes
        pueden terminar en un duplo salvo que se marque la casilla."""
        g = grupo([
            Pax(nombre="Chofer Uno", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Chofer Dos", tipo=TipoPax.CHOFER, sexo=Sexo.M),
        ])
        atomos, _ = atomizar(g)
        self.assertFalse(any(a.requiere_single for a in atomos))
        self.assertEqual(ConfigGrupo().guias_choferes_single, False)

    def test_chofer_y_coordinador_tambien_single(self):
        g = grupo([
            Pax(nombre="Chofer Uno", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Coord Uno", tipo=TipoPax.COORDINADOR, sexo=Sexo.X),
        ], guias_choferes_single=True)
        atomos, _ = atomizar(g)
        self.assertTrue(all(a.requiere_single for a in atomos))

    def test_flag_manual_single(self):
        g = grupo([Pax(nombre="VIP", sexo=Sexo.M, requiere_single=True)])
        atomos, _ = atomizar(g)
        self.assertTrue(atomos[0].requiere_single)

    def test_nombres_duplicados_error(self):
        g = grupo([
            Pax(nombre="Doble", sexo=Sexo.M),
            Pax(nombre="doble ", sexo=Sexo.F),
        ])
        with self.assertRaises(ErrorDistribucion):
            atomizar(g)

    def test_bebes_sin_cupo(self):
        g = grupo(
            [
                Pax(nombre="Mama", sexo=Sexo.F, familia="F"),
                Pax(nombre="Bebe", tipo=TipoPax.BEBE, edad=1, sexo=Sexo.X, familia="F"),
            ],
            bebes_no_cuentan_cupo=True,
        )
        atomos, _ = atomizar(g)
        self.assertEqual(atomos[0].tamano, 1)
        self.assertEqual(atomos[0].miembros.__len__(), 2)

    def test_bebes_cuentan_cupo_por_defecto(self):
        g = grupo([
            Pax(nombre="Mama", sexo=Sexo.F, familia="F"),
            Pax(nombre="Bebe", tipo=TipoPax.BEBE, edad=1, sexo=Sexo.X, familia="F"),
        ])
        atomos, _ = atomizar(g)
        self.assertEqual(atomos[0].tamano, 2)

    def test_advertencia_edad_incoherente(self):
        g = grupo(
            [Pax(nombre="Grande", tipo=TipoPax.NINO, edad=30, sexo=Sexo.X)],
            ninos_con_adulto_obligatorio=False,
        )
        _, avisos = atomizar(g)
        self.assertTrue(any("ADULTO" in aviso or "edad" in aviso.lower() for aviso in avisos))

    def test_orden_determinista(self):
        pax = [Pax(nombre=f"P{i}", sexo=Sexo.M) for i in range(5)]
        g1 = grupo(list(pax))
        g2 = grupo(list(reversed(pax)))
        a1, _ = atomizar(g1)
        a2, _ = atomizar(g2)
        self.assertEqual([a.descripcion for a in a1], [a.descripcion for a in a2])


if __name__ == "__main__":
    unittest.main()
