from __future__ import annotations

import os
import random
import sys
import unittest
from collections import Counter
from itertools import product

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from empaquetador import (  # noqa: E402
    Estrategia,
    calcular_deficit,
    empaquetar,
)
from modelos import Atomo, ConfigGrupo, ErrorDistribucion  # noqa: E402

CONFIG_ABIERTA = ConfigGrupo(separar_sexos=False, ninos_con_adulto_obligatorio=False)


def atomo(tamano: int = 1, sexo=None, requiere_single: bool = False) -> Atomo:
    return Atomo(
        miembros=(),
        tamano=tamano,
        sexos_no_x=frozenset({sexo}) if sexo else frozenset(),
        tiene_menor=False,
        tiene_adulto=True,
        requiere_single=requiere_single,
    )


def _optimo_bruto(tamanos: list[int], caps: list[int]) -> tuple[int, int] | None:
    mejor = None
    cantidad_bins = len(caps)
    for asignacion in product(range(cantidad_bins), repeat=len(tamanos)):
        cargas = [0] * cantidad_bins
        ok = True
        for tamano, indice in zip(tamanos, asignacion):
            cargas[indice] += tamano
            if cargas[indice] > caps[indice]:
                ok = False
                break
        if not ok:
            continue
        usadas = sum(1 for c in cargas if c > 0)
        desperdicio = sum(
            caps[i] - cargas[i] for i in range(cantidad_bins) if cargas[i] > 0
        )
        score = (usadas, desperdicio)
        if mejor is None or score < mejor:
            mejor = score
    return mejor


def verificar_empaque(resultado, atomos_originales, config: ConfigGrupo) -> list[str]:
    violaciones: list[str] = []
    colocadas = [a for h in resultado.habitaciones for a in h.atomos]
    colocadas += list(resultado.pendientes)
    ids = [id(a) for a in colocadas]
    originales = [id(a) for a in atomos_originales]
    if len(ids) != len(set(ids)):
        violaciones.append("atomos duplicados en la salida")
    if sorted(ids) != sorted(originales):
        violaciones.append("faltan o sobran atomos")
        for h in resultado.habitaciones:
            if h.carga > h.capacidad:
                violaciones.append(f"sobrecarga {h.carga}/{h.capacidad}")
            if len(h.atomos) > 1:
                for a in h.atomos:
                    if a.requiere_single:
                        violaciones.append(f"single acompanado ({a.descripcion})")
                sexos = set()
                for a in h.atomos:
                    sexos |= set(a.sexos_no_x)
                mixtos = [a for a in h.atomos if a.es_mixto]
                if len(mixtos) > 1 and config.separar_sexos:
                    violaciones.append("dos atomos mixtos juntos")
                if len(sexos) > 1 and not mixtos and config.separar_sexos:
                    violaciones.append("mezcla sin atomo exento")
    return violaciones


class PruebasEmpaquetador(unittest.TestCase):
    def test_caso_simple_optimo(self):
        atomos = [atomo(2), atomo(2), atomo(2), atomo(1)]
        caps = [3, 2, 2, 2]
        resultado = empaquetar(atomos, caps, CONFIG_ABIERTA, estrategia="EXACTA", semilla=1)
        self.assertTrue(resultado.completo)
        self.assertEqual(resultado.score(), (3, 0))

    def test_todo_en_una_habitacion(self):
        atomos = [atomo(1)] * 4
        resultado = empaquetar(atomos, [4], CONFIG_ABIERTA, estrategia="AUTO", semilla=1)
        self.assertTrue(resultado.completo)
        self.assertEqual(len(resultado.habitaciones), 1)

    def test_infactible_reportado(self):
        atomos = [atomo(1)] * 5
        resultado = empaquetar(atomos, [2, 2], CONFIG_ABIERTA, estrategia="AUTO", semilla=1)
        self.assertFalse(resultado.completo)
        self.assertEqual(len(resultado.pendientes), 1)

    def test_sin_capacidades(self):
        atomos = [atomo(1)]
        resultado = empaquetar(atomos, [], CONFIG_ABIERTA, estrategia="AUTO", semilla=1)
        self.assertFalse(resultado.completo)
        self.assertEqual(len(resultado.pendientes), 1)

    def test_atomo_mas_grande_que_toda_habitacion_error(self):
        with self.assertRaises(ErrorDistribucion):
            empaquetar([atomo(5)], [4, 3], CONFIG_ABIERTA, estrategia="AUTO", semilla=1)

    def test_optimalidad_contra_fuerza_bruta(self):
        rng = random.Random(123)
        comparados = 0
        for caso in range(60):
            cantidad_bins = rng.randint(2, 4)
            caps = sorted((rng.randint(2, 4) for _ in range(cantidad_bins)), reverse=True)
            n = rng.randint(2, 6)
            tamanos = [rng.randint(1, 3) for _ in range(n)]
            if any(t > max(caps) for t in tamanos):
                continue
            esperado = _optimo_bruto(tamanos, caps)
            piezas = [atomo(t) for t in tamanos]
            resultado = empaquetar(piezas, caps, CONFIG_ABIERTA, estrategia="EXACTA",
                                   semilla=caso + 500)
            if esperado is None:
                self.assertFalse(resultado.completo, f"caso {caso}: debia ser infactible")
            else:
                self.assertTrue(resultado.completo, f"caso {caso}: debia tener solucion")
                self.assertEqual(
                    tuple(resultado.score()), esperado,
                    f"caso {caso}: obtuvo {resultado.score()} y el optimo es {esperado}",
                )
            comparados += 1
        self.assertGreater(comparados, 30)

    def test_invariantes_todas_las_estrategias(self):
        estrategias = ["AUTO", "BFD", "FFD", "MULTISTART", "EXACTA", "TODAS"]
        configs = [
            ConfigGrupo(),
            ConfigGrupo(separar_sexos=True),
            CONFIG_ABIERTA,
        ]
        rng = random.Random(777)
        for semilla_caso in range(12):
            caps = [rng.choice([1, 2, 2, 3, 4]) for _ in range(rng.randint(3, 10))]
            piezas = []
            for _ in range(rng.randint(2, 8)):
                dado = rng.random()
                if dado < 0.15:
                    piezas.append(atomo(1, sexo=rng.choice(["M", "F"]), requiere_single=True))
                elif dado < 0.55:
                    piezas.append(atomo(1, sexo=rng.choice(["M", "F"])))
                else:
                    piezas.append(atomo(1))
            for nombre_estrategia in estrategias:
                for config in configs:
                    copia = [
                        atomo(a.tamano,
                              next(iter(a.sexos_no_x)) if a.sexos_no_x else None,
                              a.requiere_single)
                        for a in piezas
                    ]
                    resultado = empaquetar(
                        copia, list(caps), config,
                        estrategia=nombre_estrategia, semilla=semilla_caso,
                    )
                    problemas = verificar_empaque(resultado, copia, config)
                    self.assertEqual(
                        problemas, [],
                        f"estrategia {nombre_estrategia} config {config}: {problemas}",
                    )

    def test_separacion_de_sexos_respetada(self):
        piezas = [
            atomo(1, sexo="M"), atomo(1, sexo="M"),
            atomo(1, sexo="F"), atomo(1, sexo="F"),
        ]
        caps = [4, 4]
        resultado = empaquetar(piezas, caps, ConfigGrupo(separar_sexos=True),
                               estrategia="MULTISTART", semilla=3)
        self.assertTrue(resultado.completo)
        for h in resultado.habitaciones:
            sexos = set()
            for a in h.atomos:
                sexos |= set(a.sexos_no_x)
            self.assertLessEqual(len(sexos), 1)

    def test_single_siempre_solo(self):
        piezas = [
            atomo(1, sexo="M", requiere_single=True),
            atomo(1, sexo="M"),
            atomo(1, sexo="M"),
        ]
        caps = [4, 4, 4]
        resultado = empaquetar(piezas, caps, ConfigGrupo(), estrategia="AUTO", semilla=9)
        self.assertTrue(resultado.completo)
        for h in resultado.habitaciones:
            singles = [a for a in h.atomos if a.requiere_single]
            if singles:
                self.assertEqual(len(h.atomos), 1)

    def test_determinismo_con_semilla(self):
        piezas = [atomo(1, sexo=s) for s in ("M", "M", "F", "F", "X", "X")]
        caps = [3, 3, 3, 2]
        r1 = empaquetar(piezas[:], caps, ConfigGrupo(), estrategia="MULTISTART", semilla=99)
        r2 = empaquetar(piezas[:], caps, ConfigGrupo(), estrategia="MULTISTART", semilla=99)
        firma1 = [(h.capacidad, tuple(sorted(a.descripcion for a in h.atomos)))
                  for h in r1.habitaciones]
        firma2 = [(h.capacidad, tuple(sorted(a.descripcion for a in h.atomos)))
                  for h in r2.habitaciones]
        self.assertEqual(firma1, firma2)

    def test_calcular_deficit_basico(self):
        piezas = [atomo(1) for _ in range(5)]
        deficit = calcular_deficit(piezas, [2], CONFIG_ABIERTA)
        self.assertEqual(deficit[2], 3)

    def test_calcular_deficit_mixto(self):
        piezas = [atomo(1) for _ in range(6)]
        deficit = calcular_deficit(piezas, [4, 2], CONFIG_ABIERTA)
        total_camas = sum(cap * cant for cap, cant in deficit.items())
        self.assertGreaterEqual(total_camas, 6)

    def test_multistart_nunca_peor_que_bfd_en_promedio(self):
        rng = random.Random(42)
        peor = 0
        for caso in range(25):
            caps = [rng.randint(2, 4) for _ in range(rng.randint(3, 8))]
            max_cap = max(caps) if caps else 0
            piezas = [atomo(rng.randint(1, min(3, max_cap or 1))) for _ in range(rng.randint(3, 10))]
            if not caps:
                continue
            bfd = empaquetar(piezas, caps, CONFIG_ABIERTA, estrategia="BFD", semilla=caso)
            multi = empaquetar(piezas, caps, CONFIG_ABIERTA, estrategia="MULTISTART", semilla=caso)
            if bfd.completo and multi.completo:
                peor = max(peor, multi.score()[0] - bfd.score()[0])
        self.assertLessEqual(peor, 0)


if __name__ == "__main__":
    unittest.main()
