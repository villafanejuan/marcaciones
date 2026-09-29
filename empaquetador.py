from __future__ import annotations

import math
import random
import sys
import time
from collections import Counter
from dataclasses import dataclass, field
from datetime import date
from enum import Enum

from modelos import (
    Atomo,
    ConfigGrupo,
    ErrorDistribucion,
    EstadoHabitacion,
)


class Estrategia(str, Enum):
    AUTO = "AUTO"
    EXACTA = "EXACTA"
    BFD = "BFD"
    FFD = "FFD"
    MULTISTART = "MULTISTART"
    TODAS = "TODAS"


LIMITE_ATOMOS_EXACTA = 12


@dataclass
class HabitacionVirtual:
    capacidad: int
    atomos: list[Atomo] = field(default_factory=list)

    @property
    def carga(self) -> int:
        return sum(a.tamano for a in self.atomos)


@dataclass
class ResultadoEmpaque:
    habitaciones: list[HabitacionVirtual]
    pendientes: list[Atomo]
    estrategia: str
    nodos_explorados: int = 0
    tiempo: float = 0.0
    optimo: bool = False

    @property
    def completo(self) -> bool:
        return not self.pendientes

    @property
    def camas_ocupadas(self) -> int:
        return sum(h.carga for h in self.habitaciones)

    @property
    def camas_totales(self) -> int:
        return sum(h.capacidad for h in self.habitaciones)

    @property
    def desperdicio(self) -> int:
        return self.camas_totales - self.camas_ocupadas

    def score(self) -> tuple[int, int]:
        return (len(self.habitaciones), self.desperdicio)


def _ordenar_atomos(atomos: list[Atomo], rng: random.Random | None = None, aleatorio: bool = False) -> list[Atomo]:
    base = sorted(atomos, key=lambda a: a.clave_orden)
    if not aleatorio or rng is None:
        return base
    grupos: dict[tuple, list[Atomo]] = {}
    for a in base:
        clave = (a.requiere_single, a.tamano, a.menores, len(a.sexos_no_x))
        grupos.setdefault(clave, []).append(a)
    resultado: list[Atomo] = []
    for clave in sorted(grupos):
        lista = grupos[clave][:]
        rng.shuffle(lista)
        resultado.extend(lista)
    return resultado


def _elegir_apertura(disponibles: Counter, tamano_minimo: int, mayor_primero: bool = True) -> int | None:
    candidatos = [c for c in disponibles if disponibles[c] > 0 and c >= tamano_minimo]
    if not candidatos:
        return None
    return max(candidatos) if mayor_primero else min(candidatos)


def _colocar_heuristico(
    atomos_ordenados: list[Atomo],
    capacidades_disponibles: Counter,
    config: ConfigGrupo,
    modo_best_fit: bool = True,
    apertura_mayor: bool = True,
    rng: random.Random | None = None,
    aleatorio: bool = False,
) -> tuple[list[EstadoHabitacion], list[Atomo]]:
    bins: list[EstadoHabitacion] = []
    disponibles = Counter(capacidades_disponibles)
    pendientes: list[Atomo] = []
    for atomo in atomos_ordenados:
        destino = None
        if modo_best_fit:
            candidatos = [b for b in bins if b.puede_recibir(atomo, config)]
            if candidatos:
                if aleatorio and rng is not None and len(candidatos) > 1:
                    candidatos.sort(key=lambda b: (b.espacio_libre - atomo.tamano, b.capacidad))
                    destino = rng.choice(candidatos[:2])
                else:
                    destino = min(candidatos, key=lambda b: (b.espacio_libre - atomo.tamano, b.capacidad))
        else:
            destino = next((b for b in bins if b.puede_recibir(atomo, config)), None)
        if destino is not None:
            destino.recibir(atomo)
            continue
        cap = _elegir_apertura(disponibles, atomo.tamano, mayor_primero=apertura_mayor)
        if cap is None:
            pendientes.append(atomo)
            continue
        disponibles[cap] -= 1
        nuevo = EstadoHabitacion(cap)
        nuevo.recibir(atomo)
        bins.append(nuevo)
    bins.sort(key=lambda b: (-b.carga, -b.capacidad))
    return bins, pendientes


def _snapshot(bins: list[EstadoHabitacion]) -> list[HabitacionVirtual]:
    ordenados = sorted(bins, key=lambda b: (-b.carga, -b.capacidad))
    return [HabitacionVirtual(capacidad=b.capacidad, atomos=list(b.atomos)) for b in ordenados]


def _score_bins(bins: list[EstadoHabitacion]) -> tuple[int, int]:
    return (len(bins), sum(b.capacidad - b.carga for b in bins))


def _reconstruir(snapshot: list[tuple[int, list[Atomo]]]) -> list[EstadoHabitacion]:
    bins = []
    for cap, atomos in snapshot:
        e = EstadoHabitacion(cap)
        for a in atomos:
            e.recibir(a)
        bins.append(e)
    bins.sort(key=lambda b: (-b.carga, -b.capacidad))
    return bins


def _busqueda_exacta(
    atomos: list[Atomo],
    capacidades: list[int],
    config: ConfigGrupo,
    presupuesto_nodos: int = 400000,
) -> tuple[list[EstadoHabitacion] | None, int, bool]:
    orden = _ordenar_atomos(atomos)
    n = len(orden)
    caps_desc = sorted(capacidades, reverse=True)
    cap_maxima = caps_desc[0] if caps_desc else 0
    disponibles = Counter(capacidades)

    sufijos = [0] * (n + 1)
    for i in range(n - 1, -1, -1):
        sufijos[i] = sufijos[i + 1] + orden[i].tamano

    limite_recursion = sys.getrecursionlimit()
    sys.setrecursionlimit(max(limite_recursion, n * 5 + 300))

    estado = {"nodos": 0, "agotado": False, "score": None, "solucion": None}
    abiertas: Counter = Counter()

    def dfs(i: int, bins: list[EstadoHabitacion]) -> None:
        if estado["agotado"]:
            return
        estado["nodos"] += 1
        if estado["nodos"] > presupuesto_nodos:
            estado["agotado"] = True
            return
        if i == n:
            score = _score_bins(bins)
            if estado["score"] is None or score < estado["score"]:
                estado["score"] = score
                estado["solucion"] = [(b.capacidad, list(b.atomos)) for b in bins]
            return
        restante = sufijos[i]
        if estado["score"] is not None and cap_maxima > 0:
            libre_abierto = sum(b.espacio_libre for b in bins)
            excedente = restante - libre_abierto
            cota = len(bins) + (math.ceil(excedente / cap_maxima) if excedente > 0 else 0)
            if cota > estado["score"][0]:
                return
        atomo = orden[i]
        vistos: set[tuple] = set()
        candidatos = [b for b in bins if b.puede_recibir(atomo, config)]
        candidatos.sort(key=lambda b: (b.espacio_libre - atomo.tamano, b.capacidad))
        for b in candidatos:
            firma = b.firma()
            if firma in vistos:
                continue
            vistos.add(firma)
            b.recibir(atomo)
            dfs(i + 1, bins)
            b.quitar(atomo)
            if estado["agotado"]:
                return
        probadas: set[int] = set()
        for cap in caps_desc:
            if cap < atomo.tamano or cap in probadas:
                continue
            if abiertas[cap] >= disponibles[cap]:
                continue
            probadas.add(cap)
            nuevo = EstadoHabitacion(cap)
            nuevo.recibir(atomo)
            bins.append(nuevo)
            abiertas[cap] += 1
            dfs(i + 1, bins)
            bins.pop()
            abiertas[cap] -= 1
            if estado["agotado"]:
                return

    dfs(0, [])
    solucion = _reconstruir(estado["solucion"]) if estado["solucion"] is not None else None
    return solucion, estado["nodos"], estado["agotado"]


def _mejora_local(bins: list[EstadoHabitacion], capacidades_disponibles: Counter, config: ConfigGrupo) -> None:
    mejoro = True
    pasadas = 0
    while mejoro and pasadas < 8:
        mejoro = False
        pasadas += 1
        for origen in list(bins):
            if len(bins) <= 1:
                break
            for atomo in list(origen.atomos):
                origen.quitar(atomo)
                destino = next(
                    (b for b in bins if b is not origen and b.puede_recibir(atomo, config)),
                    None,
                )
                if destino is not None:
                    destino.recibir(atomo)
                    if origen.vacia:
                        bins.remove(origen)
                        capacidades_disponibles[origen.capacidad] += 1
                        mejoro = True
                        break
                    destino.quitar(atomo)
                    origen.recibir(atomo)
                else:
                    origen.recibir(atomo)


def _multistart(
    atomos: list[Atomo],
    capacidades_disponibles: Counter,
    config: ConfigGrupo,
    iteraciones: int,
    tiempo_max: float,
    semilla: int,
) -> list[EstadoHabitacion] | None:
    rng = random.Random(semilla)
    inicio = time.monotonic()
    mejor_score: tuple[int, int] | None = None
    mejor_snapshot: list[tuple[int, list[Atomo]]] | None = None
    for indice_iteracion in range(iteraciones):
        if indice_iteracion == 0:
            bins, pendientes = _colocar_heuristico(
                _ordenar_atomos(atomos), capacidades_disponibles, config
            )
        else:
            orden = _ordenar_atomos(atomos, rng=rng, aleatorio=True)
            bins, pendientes = _colocar_heuristico(
                orden, capacidades_disponibles, config, rng=rng, aleatorio=True
            )
        if pendientes:
            continue
        _mejora_local(bins, Counter(capacidades_disponibles), config)
        score = _score_bins(bins)
        if mejor_score is None or score < mejor_score:
            mejor_score = score
            mejor_snapshot = [(b.capacidad, list(b.atomos)) for b in bins]
        if time.monotonic() - inicio > tiempo_max:
            break
    return _reconstruir(mejor_snapshot) if mejor_snapshot is not None else None


def empaquetar(
    atomos: list[Atomo],
    capacidades: list[int],
    config: ConfigGrupo,
    estrategia: str | Estrategia = Estrategia.AUTO,
    semilla: int | None = None,
    tiempo_max: float = 4.0,
    presupuesto_nodos: int = 400000,
) -> ResultadoEmpaque:
    inicio = time.monotonic()
    nombre_estrategia = estrategia.value if isinstance(estrategia, Estrategia) else str(estrategia).upper()
    try:
        estrategia_enum = Estrategia(nombre_estrategia)
    except ValueError:
        raise ErrorDistribucion(
            f"Estrategia desconocida: '{nombre_estrategia}'. "
            "Opciones: AUTO, EXACTA, BFD, FFD, MULTISTART, TODAS."
        )

    if semilla is None:
        base = ",".join(sorted(a.descripcion.lower() for a in atomos))
        semilla = (hash(base) & 0x7FFFFFFF) or 42

    caps_lista = sorted(capacidades, reverse=True)
    caps_counter = Counter(caps_lista)

    if atomos and caps_lista:
        cap_maxima = caps_lista[0]
        excedidos = [a for a in atomos if a.tamano > cap_maxima]
        if excedidos:
            nombres = "; ".join(f"{a.descripcion} (necesita {a.tamano} lugares)" for a in excedidos)
            raise ErrorDistribucion(
                f"No existe habitación con capacidad suficiente para: {nombres}. "
                f"La habitación más grande tiene {cap_maxima} lugares."
            )

    intentos: list[str] = []
    if estrategia_enum in (Estrategia.AUTO, Estrategia.TODAS, Estrategia.BFD):
        intentos.append("BFD")
    if estrategia_enum in (Estrategia.AUTO, Estrategia.TODAS):
        intentos.append("BFD-AJUSTADA")
    if estrategia_enum in (Estrategia.AUTO, Estrategia.TODAS, Estrategia.FFD):
        intentos.append("FFD")
    if estrategia_enum == Estrategia.EXACTA:
        intentos.append("EXACTA")
        intentos.append("MULTISTART")
        intentos.append("BFD")
    if estrategia_enum in (Estrategia.AUTO, Estrategia.TODAS, Estrategia.MULTISTART):
        intentos.append("MULTISTART")
    if estrategia_enum in (Estrategia.AUTO, Estrategia.EXACTA, Estrategia.TODAS):
        if len(atomos) <= LIMITE_ATOMOS_EXACTA:
            intentos.append("EXACTA")

    resultados: dict[str, ResultadoEmpaque] = {}
    nodos_acumulados = 0

    def registrar(nombre: str, bins: list[EstadoHabitacion] | None, pendientes: list[Atomo], optimo: bool) -> None:
        habitaciones_virtuales = (
            [HabitacionVirtual(capacidad=b.capacidad, atomos=list(b.atomos)) for b in _snapshot(bins)]
            if bins is not None
            else []
        )
        resultado = ResultadoEmpaque(
            habitaciones=habitaciones_virtuales,
            pendientes=list(pendientes),
            estrategia=nombre,
            nodos_explorados=nodos_acumulados,
            optimo=optimo,
        )
        previo = resultados.get(nombre)
        if previo is None or _mejor(resultado, previo):
            resultados[nombre] = resultado

    def _mejor(nuevo: ResultadoEmpaque, actual: ResultadoEmpaque) -> bool:
        if nuevo.completo != actual.completo:
            return nuevo.completo
        if not nuevo.completo:
            return len(nuevo.pendientes) < len(actual.pendientes)
        return nuevo.score() < actual.score()

    for nombre in intentos:
        if nombre == "BFD":
            bins, pendientes = _colocar_heuristico(
                _ordenar_atomos(atomos), caps_counter, config, modo_best_fit=True, apertura_mayor=True
            )
            registrar(nombre, bins, pendientes, optimo=False)
        elif nombre == "BFD-AJUSTADA":
            bins, pendientes = _colocar_heuristico(
                _ordenar_atomos(atomos), caps_counter, config, modo_best_fit=True, apertura_mayor=False
            )
            registrar(nombre, bins, pendientes, optimo=False)
        elif nombre == "FFD":
            bins, pendientes = _colocar_heuristico(
                _ordenar_atomos(atomos), caps_counter, config, modo_best_fit=False, apertura_mayor=True
            )
            registrar(nombre, bins, pendientes, optimo=False)
        elif nombre == "MULTISTART":
            n = len(atomos)
            iteraciones = max(60, min(400, 60000 // max(1, n)))
            bins = _multistart(atomos, caps_counter, config, iteraciones, tiempo_max, semilla)
            registrar(nombre, bins, [] if bins is not None else list(atomos), optimo=False)
        elif nombre == "EXACTA":
            presupuesto = presupuesto_nodos * (10 if estrategia_enum == Estrategia.EXACTA else 1)
            bins, nodos, agotado = _busqueda_exacta(atomos, caps_lista, config, presupuesto)
            nodos_acumulados += nodos
            registrar(nombre, bins, [] if bins is not None else list(atomos), optimo=(bins is not None and not agotado))

    if not resultados:
        vacio = ResultadoEmpaque(
            habitaciones=[],
            pendientes=list(atomos),
            estrategia="SIN-CAPACIDAD",
        )
        vacio.tiempo = time.monotonic() - inicio
        return vacio

    completos = [r for r in resultados.values() if r.completo]
    ganador = min(completos, key=lambda r: r.score()) if completos else min(
        resultados.values(), key=lambda r: len(r.pendientes)
    )
    ganador.nodos_explorados = nodos_acumulados
    ganador.tiempo = time.monotonic() - inicio
    return ganador


def empaquetar_todas(
    atomos: list[Atomo],
    capacidades: list[int],
    config: ConfigGrupo,
    semilla: int | None = None,
    tiempo_max: float = 4.0,
) -> list[ResultadoEmpaque]:
    nombres = ["BFD", "BFD-AJUSTADA", "FFD", "MULTISTART"]
    if len(atomos) <= LIMITE_ATOMOS_EXACTA:
        nombres.append("EXACTA")
    salida = []
    for nombre in nombres:
        resultado = empaquetar(
            atomos, capacidades, config, estrategia=nombre, semilla=semilla, tiempo_max=tiempo_max
        )
        salida.append(resultado)
    return salida


def calcular_deficit(
    atomos: list[Atomo],
    capacidades_distintas: list[int],
    config: ConfigGrupo,
) -> Counter:
    caps = sorted({c for c in capacidades_distintas if c >= 1}, reverse=True)
    conteo: Counter = Counter()
    if not atomos:
        return conteo
    if not caps:
        raise ErrorDistribucion("No hay tipos de habitación definidos en el hotel.")
    menor_cap = caps[-1]
    imposibles = [a for a in atomos if a.tamano > caps[0]]
    if imposibles:
        nombres = ", ".join(a.descripcion for a in imposibles)
        raise ErrorDistribucion(
            f"Ninguna habitación puede alojar a: {nombres} (capacidad máxima disponible {caps[0]})."
        )
    orden = _ordenar_atomos(atomos)
    bins: list[EstadoHabitacion] = []
    for atomo in orden:
        candidatos = [b for b in bins if b.puede_recibir(atomo, config)]
        if candidatos:
            elegido = min(candidatos, key=lambda b: (b.espacio_libre - atomo.tamano, b.capacidad))
            elegido.recibir(atomo)
        else:
            aptas = [c for c in caps if c >= atomo.tamano]
            cap = max(aptas)
            nuevo = EstadoHabitacion(cap)
            nuevo.recibir(atomo)
            bins.append(nuevo)
    for b in bins:
        conteo[b.capacidad] += 1
    return conteo
