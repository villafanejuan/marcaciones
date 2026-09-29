"""Modo de hospedaje estudiantil: alumnos aparte de los acompañantes, y adultos
agrupados por rol y sexo en la habitación más justa que entra."""

from __future__ import annotations

import os
import sys
import unittest

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from fabrica import mk_grupo, mk_hotel  # noqa: E402
from modelos import (  # noqa: E402
    ConfigGrupo,
    ErrorDistribucion,
    ModoHospedaje,
    Pax,
    Sexo,
    TipoPax,
)
from motor import distribuir, verificar_reporte  # noqa: E402
from persistencia import (  # noqa: E402
    cargar_datos,
    config_a_dict,
    config_de_dict,
    guardar_datos,
)


def alumnos(cantidad: int, sexo: Sexo) -> list[Pax]:
    return [
        Pax(nombre=f"Est-{sexo.value}-{i:02d}", tipo=TipoPax.ESTUDIANTE, sexo=sexo, edad=17)
        for i in range(1, cantidad + 1)
    ]


def ejecucion(pax, config, capacidades=(4,) * 12):
    hotel = mk_hotel(list(capacidades))
    grupo = mk_grupo("G", pax=pax, config=config)
    reporte = distribuir(hotel, [grupo])
    self_codigo = verificar_reporte(reporte)
    if self_codigo:
        raise AssertionError(f"invariantes rotos: {self_codigo}")
    return reporte.distribuciones[0]


def ocupantes(dist):
    """Devuelve la lista de occupants por habitacion, en el orden de las asignaciones."""
    return [sorted(m.nombre for a in asig.atomos for m in a.miembros) for asig in dist.asignaciones]


def es_alumno(nombre: str) -> bool:
    return nombre.startswith("Est-")


class PruebasModoEstudiantil(unittest.TestCase):
    def _hab_de(self, nombre, dist):
        for habitacion in ocupantes(dist):
            if nombre in habitacion:
                return habitacion
        raise AssertionError(f"{nombre} no aparece en ninguna habitacion")

    def test_adultos_y_alumnos_nunca_comparten_habitacion(self):
        pax = alumnos(8, Sexo.M) + alumnos(8, Sexo.F) + [
            Pax(nombre="Apod-1", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Apod-2", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Apod-3", tipo=TipoPax.ADULTO, sexo=Sexo.M),
        ]
        dist = ejecucion(pax, ConfigGrupo(estudiantes_por_habitacion=4))
        for habitacion in ocupantes(dist):
            hay_alumno = any(es_alumno(n) for n in habitacion)
            hay_adulto = any(not es_alumno(n) for n in habitacion)
            self.assertFalse(hay_alumno and hay_adulto,
                             f"se mezclaron alumnos y adultos: {habitacion}")

    def test_adultos_agrupados_por_rol_y_sexo(self):
        """2 apoderadas = un duplo, 3 apoderados = un triplo, 2 choferes = un duplo,
        1 guia = un single. Cada rol y sexo en su propia habitacion."""
        pax = alumnos(4, Sexo.M) + [
            Pax(nombre="Apod-F1", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Apod-F2", tipo=TipoPax.ADULTO, sexo=Sexo.F),
            Pax(nombre="Apod-M1", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Apod-M2", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Apod-M3", tipo=TipoPax.ADULTO, sexo=Sexo.M),
            Pax(nombre="Chof-1", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Chof-2", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Guia-1", tipo=TipoPax.GUIA, sexo=Sexo.F),
        ]
        dist = ejecucion(pax, ConfigGrupo(estudiantes_por_habitacion=4))

        self.assertEqual(set(self._hab_de("Apod-F1", dist)), {"Apod-F1", "Apod-F2"})
        self.assertEqual(set(self._hab_de("Chof-1", dist)), {"Chof-1", "Chof-2"})
        self.assertEqual(set(self._hab_de("Guia-1", dist)), {"Guia-1"})
        self.assertEqual(set(self._hab_de("Apod-M1", dist)), {"Apod-M1", "Apod-M2", "Apod-M3"})

    def test_habitacion_es_la_mas_justa_que_ente(self):
        """Un trio de acompanantes tiene que ir a un room de 4 plazas, no a uno de 6."""
        pax = alumnos(4, Sexo.M) + [
            Pax(nombre=f"Apod-M{i}", tipo=TipoPax.ADULTO, sexo=Sexo.M) for i in range(1, 4)
        ]
        capacidades = (4,) * 4 + (6,) * 4
        dist = ejecucion(pax, ConfigGrupo(estudiantes_por_habitacion=4), capacidades)
        trio = self._hab_de("Apod-M1", dist)
        self.assertEqual(len(trio), 3)
        for asig in dist.asignaciones:
            if trio[0] in asig.nombres_pax:
                self.assertEqual(asig.habitacion.capacidad, 4)

    def test_sin_rol_se_juntan_los_del_mismo_sexo(self):
        pax = alumnos(4, Sexo.M) + [
            Pax(nombre="Chof-1", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Guia-1", tipo=TipoPax.GUIA, sexo=Sexo.M),
        ]
        config = ConfigGrupo(estudiantes_por_habitacion=4, adultos_por_rol=False)
        dist = ejecucion(pax, config)
        juntos = set(self._hab_de("Chof-1", dist))
        self.assertEqual(juntos, {"Chof-1", "Guia-1"})

    def test_opcion_single_parte_uno_por_habitacion(self):
        pax = alumnos(4, Sexo.M) + [
            Pax(nombre="Chof-1", tipo=TipoPax.CHOFER, sexo=Sexo.M),
            Pax(nombre="Chof-2", tipo=TipoPax.CHOFER, sexo=Sexo.M),
        ]
        config = ConfigGrupo(estudiantes_por_habitacion=4, guias_choferes_single=True)
        dist = ejecucion(pax, config)
        self.assertEqual(set(self._hab_de("Chof-1", dist)), {"Chof-1"})
        self.assertEqual(set(self._hab_de("Chof-2", dist)), {"Chof-2"})

    def test_max_plazas_evita_habitaciones_grandes(self):
        pax = alumnos(8, Sexo.M) + [Pax(nombre="Chof-1", tipo=TipoPax.CHOFER, sexo=Sexo.M)]
        config = ConfigGrupo(estudiantes_por_habitacion=4, max_plazas_estudiantes=4)
        dist = ejecucion(pax, config, capacidades=(4,) * 3 + (6,) * 4)
        for asig in dist.asignaciones:
            es_sala_de_alumnos = any(m.tipo == TipoPax.ESTUDIANTE for m in asig.atomos[0].miembros)
            if es_sala_de_alumnos:
                self.assertEqual(asig.habitacion.capacidad, 4, "un alumno quedó en una sextuple")

    def test_max_plazas_ignorado_con_aviso_si_solo_hay_grandes(self):
        pax = alumnos(8, Sexo.M)
        config = ConfigGrupo(estudiantes_por_habitacion=4, max_plazas_estudiantes=4)
        dist = ejecucion(pax, config, capacidades=(6,) * 4)
        self.assertTrue(dist.completo)
        self.assertTrue(any("plazas o menos" in a for a in dist.avisos), dist.avisos)

    def test_cupo_mayor_que_el_tope_se_ajusta_con_aviso(self):
        pax = alumnos(8, Sexo.M)
        config = ConfigGrupo(estudiantes_por_habitacion=6, max_plazas_estudiantes=4)
        dist = ejecucion(pax, config, capacidades=(4,) * 4)
        self.assertTrue(dist.completo)
        for asig in dist.asignaciones:
            self.assertLessEqual(asig.carga, 4)
        self.assertTrue(any("supera el máximo" in a for a in dist.avisos), dist.avisos)

    def test_modo_adultos_no_separa_alumnos_de_acompanantes(self):
        """Sin cupo no hay ninguna razon para separar: el reparto es libre como antes."""
        pax = alumnos(3, Sexo.M) + [Pax(nombre="Apod-1", tipo=TipoPax.ADULTO, sexo=Sexo.M)]
        config = ConfigGrupo(modo_hospedaje=ModoHospedaje.ADULTOS.value)
        dist = ejecucion(pax, config)
        habitacion = self._hab_de("Apod-1", dist)
        self.assertTrue(any(es_alumno(n) for n in habitacion),
                        "el modo adultos debe repartir libre, sin cohortes")
        self.assertEqual(len(ocupantes(dist)), 1)

    def test_modo_adultos_conserva_el_cupo_si_esta_puesto(self):
        """El cupo sigue funcionando en modo adultos: es el comportamiento previo."""
        pax = alumnos(4, Sexo.M) + [Pax(nombre="Apod-1", tipo=TipoPax.ADULTO, sexo=Sexo.M)]
        config = ConfigGrupo(estudiantes_por_habitacion=4,
                             modo_hospedaje=ModoHospedaje.ADULTOS.value)
        dist = ejecucion(pax, config)
        sala_alumnos = self._hab_de("Est-M-01", dist)
        self.assertEqual(sala_alumnos, ["Est-M-01", "Est-M-02", "Est-M-03", "Est-M-04"])
        self.assertEqual(self._hab_de("Apod-1", dist), ["Apod-1"])

    def test_modo_estudiantil_sin_alumnos_cae_a_adultos(self):
        pax = [Pax(nombre=f"Apod-M{i}", tipo=TipoPax.ADULTO, sexo=Sexo.M) for i in range(1, 4)]
        config = ConfigGrupo(modo_hospedaje=ModoHospedaje.ESTUDIANTIL.value,
                             estudiantes_por_habitacion=4)
        dist = ejecucion(pax, config)
        self.assertEqual(len(ocupantes(dist)), 1)
        self.assertNotIn("ESTUDIANTIL", dist.estrategia_usada)


class PruebasPersistenciaModo(unittest.TestCase):
    def test_round_trip_de_los_campos_nuevos(self):
        original = ConfigGrupo(
            modo_hospedaje=ModoHospedaje.ADULTOS.value,
            max_plazas_estudiantes=3,
            adultos_por_rol=False,
            guias_choferes_single=True,
        )
        datos = config_a_dict(original)
        self.assertEqual(datos["modo_hospedaje"], "ADULTOS")
        recuperado = config_de_dict(datos, "prueba")
        self.assertEqual(recuperado.modo_hospedaje, original.modo_hospedaje)
        self.assertEqual(recuperado.max_plazas_estudiantes, 3)
        self.assertFalse(recuperado.adultos_por_rol)
        self.assertTrue(recuperado.guias_choferes_single)

    def test_archivo_viejo_sin_los_campos_nuevos(self):
        """Un JSON guardado antes del cambio tiene que seguir cargando."""
        config = config_de_dict({"separar_sexos": True}, "prueba")
        self.assertEqual(config.modo_hospedaje, ModoHospedaje.ESTUDIANTIL.value)
        self.assertEqual(config.max_plazas_estudiantes, 4)
        self.assertTrue(config.adultos_por_rol)

    def test_modo_desconocido_falla_con_mensaje_claro(self):
        with self.assertRaises(ErrorDistribucion) as ctx:
            config_de_dict({"modo_hospedaje": "TURISMO"}, "Grupo X")
        self.assertIn("modo de hospedaje desconocido", str(ctx.exception))

    def test_guardar_y_abrir_conserva_el_modo(self):
        import tempfile

        config = ConfigGrupo(estudiantes_por_habitacion=4, max_plazas_estudiantes=3)
        grupo = mk_grupo("Colegio", pax=alumnos(4, Sexo.M), config=config)
        ruta = os.path.join(tempfile.mkdtemp(), "modo.json")
        guardar_datos(ruta, mk_hotel([4] * 4), [grupo])
        _hotel, grupos = cargar_datos(ruta)
        self.assertEqual(grupos[0].config.modo_hospedaje, ModoHospedaje.ESTUDIANTIL.value)
        self.assertEqual(grupos[0].config.max_plazas_estudiantes, 3)


if __name__ == "__main__":
    unittest.main()
