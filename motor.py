from __future__ import annotations

import re
from collections import Counter
from dataclasses import dataclass, field
from datetime import date

from empaquetador import (
    Estrategia,
    HabitacionVirtual,
    ResultadoEmpaque,
    calcular_deficit,
    empaquetar,
)
from modelos import (
    SEXO_MIXTO,
    Atomo,
    ConfigGrupo,
    ErrorDistribucion,
    ErrorValidacion,
    EstadoHabitacion,
    Grupo,
    Habitacion,
    Hotel,
    ModoHospedaje,
    Pax,
    Sexo,
    TipoPax,
    noches_entdesde_hasta,
)
from reglas import atomizar


def clave_orden_habitacion(numero: str) -> tuple:
    partes = re.split(r"(\d+)", numero.strip().lower())
    piezas = []
    for p in partes:
        if p.isdigit():
            piezas.append((0, int(p), ""))
        elif p:
            piezas.append((1, 0, p))
    return tuple(piezas)


@dataclass
class ReservaSesion:
    desde: date
    hasta: date
    grupo: str

    def solapa_con(self, ini: date, fin: date) -> bool:
        return ini < self.hasta and self.desde < fin


@dataclass
class AsignacionHabitacion:
    habitacion: Habitacion
    atomos: list[Atomo] = field(default_factory=list)

    @property
    def carga(self) -> int:
        return sum(a.tamano for a in self.atomos)

    @property
    def nombres_pax(self) -> list[str]:
        return [m.nombre for a in self.atomos for m in a.miembros]


@dataclass
class DistribucionGrupo:
    grupo: Grupo
    asignaciones: list[AsignacionHabitacion] = field(default_factory=list)
    faltantes: list[Pax] = field(default_factory=list)
    deficit: Counter = field(default_factory=Counter)
    avisos: list[str] = field(default_factory=list)
    estrategia_usada: str = ""
    tiempo: float = 0.0
    optimo: bool = False

    @property
    def completo(self) -> bool:
        return not self.faltantes

    @property
    def pax_alojados(self) -> int:
        return sum(len(a.nombres_pax) for a in self.asignaciones)

    @property
    def habitaciones_usadas(self) -> int:
        return len(self.asignaciones)


@dataclass
class OpcionesMotor:
    estrategia: str = Estrategia.AUTO.value
    semilla: int | None = None
    tiempo_max_por_grupo: float = 4.0
    capacidades_efectivas: dict[int, int] = field(default_factory=dict)

    def validar(self, hotel: Hotel) -> list[str]:
        errores = []
        caps_reales = {h.capacidad for h in hotel.habitaciones}
        for original, efectiva in self.capacidades_efectivas.items():
            if original not in caps_reales:
                errores.append(
                    f"Capacidad efectiva definida para habitaciones de {original} lugares, "
                    "pero el hotel no tiene habitaciones con esa capacidad."
                )
            elif efectiva < 1:
                errores.append(f"Capacidad efectiva inválida {original}->{efectiva} (debe ser >= 1).")
            elif efectiva > original:
                errores.append(
                    f"Capacidad efectiva {original}->{efectiva} no permitida: "
                    "no se puede aumentar la capacidad física de una habitación."
                )
        return errores


@dataclass
class ReporteGeneral:
    hotel: Hotel
    distribuciones: list[DistribucionGrupo] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    pax_por_noche: dict[date, int] = field(default_factory=dict)
    habitaciones_por_noche: dict[date, int] = field(default_factory=dict)

    @property
    def grupos_con_faltantes(self) -> list[DistribucionGrupo]:
        return [d for d in self.distribuciones if not d.completo]

    @property
    def todo_alojado(self) -> bool:
        return not self.grupos_con_faltantes


def _habitacion_libre(
    hab: Habitacion,
    ini: date,
    fin: date,
    sesion: dict[str, list[ReservaSesion]],
    ignorar_grupo: str | None = None,
) -> bool:
    if not hab.libre_en_rango(ini, fin):
        return False
    for reserva in sesion.get(hab.numero, []):
        if reserva.grupo == ignorar_grupo:
            continue
        if reserva.solapa_con(ini, fin):
            return False
    return True


def _es_alumno(atomo: Atomo) -> bool:
    return bool(atomo.miembros) and all(m.tipo == TipoPax.ESTUDIANTE for m in atomo.miembros)


def _modo_efectivo(atomos: list[Atomo], config: ConfigGrupo) -> str:
    """El modo estudiantil solo aplica si el grupo tiene alumnos; si no, se comporta
    como receptivo para que un grupo de adultos nunca se rompa por tener el toggle puesto."""
    if not config.es_estudiantil:
        return ModoHospedaje.ADULTOS.value
    if not any(_es_alumno(a) for a in atomos):
        return ModoHospedaje.ADULTOS.value
    return ModoHospedaje.ESTUDIANTIL.value


def _sexo_clave(atomo: Atomo) -> str:
    if not atomo.sexos_no_x:
        return Sexo.X.value
    if len(atomo.sexos_no_x) == 1:
        return next(iter(atomo.sexos_no_x)).value
    return SEXO_MIXTO


def _cohortes_adultas(adultos: list[Atomo], config: ConfigGrupo) -> list[list[Atomo]]:
    """Parte a los acompanantes en cohortes: mismo rol y mismo sexo van juntos.
    Con `adultos_por_rol=False` el rol se ignora y solo se separa por sexo."""
    grupos: dict[tuple, list[Atomo]] = {}
    for atomo in sorted(adultos, key=lambda a: a.clave_orden):
        roles = tuple(sorted({m.tipo.value for m in atomo.miembros})) if config.adultos_por_rol else ()
        grupos.setdefault((roles, _sexo_clave(atomo)), []).append(atomo)
    return [grupos[clave] for clave in sorted(grupos)]


def _caben_en_uno(estado: EstadoHabitacion, orden: list[Atomo], config: ConfigGrupo) -> bool:
    for atomo in orden:
        if not estado.puede_recibir(atomo, config):
            return False
        estado.recibir(atomo)
    return True


def _empaquetar_cohorte(
    pool: list[Atomo],
    capacidades_disponibles: Counter,
    config: ConfigGrupo,
) -> tuple[list[HabitacionVirtual], Counter, list[Atomo]]:
    """Una cohorte (mismo rol y sexo) se aloja junta en la habitacion mas justa que
    entre: 2 apoderadas van a un duplo, 3 a un triplo, 1 a un single. Si no entra en
    una sola habitacion se reparte por mejor ajuste."""
    if not pool:
        return [], Counter(capacidades_disponibles), []

    disponibles = Counter(capacidades_disponibles)
    orden = sorted(pool, key=lambda a: a.clave_orden)
    total = sum(a.tamano for a in orden)

    for cap in sorted(c for c in disponibles if disponibles[c] > 0 and c >= total):
        estado = EstadoHabitacion(cap)
        if _caben_en_uno(estado, orden, config):
            disponibles[cap] -= 1
            return [HabitacionVirtual(capacidad=cap, atomos=list(orden))], disponibles, []

    bins: list[EstadoHabitacion] = []
    pendientes: list[Atomo] = []
    for atomo in orden:
        destino = min(
            (b for b in bins if b.puede_recibir(atomo, config)),
            key=lambda b: (b.espacio_libre - atomo.tamano, b.capacidad),
            default=None,
        )
        if destino is not None:
            destino.recibir(atomo)
            continue
        cap = max((c for c in disponibles if disponibles[c] > 0 and c >= atomo.tamano), default=None)
        if cap is None:
            pendientes.append(atomo)
            continue
        disponibles[cap] -= 1
        nuevo = EstadoHabitacion(cap)
        nuevo.recibir(atomo)
        bins.append(nuevo)

    bins.sort(key=lambda b: (-b.carga, -b.capacidad))
    virtuales = [HabitacionVirtual(capacidad=b.capacidad, atomos=list(b.atomos)) for b in bins]
    return virtuales, disponibles, pendientes


def _separar_estudiantes(atomos: list[Atomo], config: ConfigGrupo) -> tuple[list[Atomo], list[Atomo]]:
    if config.estudiantes_por_habitacion is None:
        return [], list(atomos)
    pool: list[Atomo] = []
    resto: list[Atomo] = []
    for a in atomos:
        if _es_alumno(a):
            pool.append(a)
        else:
            resto.append(a)
    return pool, resto


def _validar_divisibilidad_estudiantes(pool: list[Atomo], k: int, config: ConfigGrupo, nombre_grupo: str) -> None:
    if not config.estudiantes_estricto or not pool:
        return
    if config.separar_sexos:
        por_sexo: dict[str, int] = {}
        for a in pool:
            for m in a.miembros:
                por_sexo[m.sexo.value] = por_sexo.get(m.sexo.value, 0) + 1
        for sexo, cantidad in sorted(por_sexo.items()):
            resto = cantidad % k
            if resto != 0:
                raise ErrorValidacion(
                    f"Grupo '{nombre_grupo}': modo estricto exige habitaciones completas de {k} estudiantes, "
                    f"pero hay {cantidad} estudiantes sexo {sexo} ({cantidad} % {k} = sobran {resto}). "
                    "Ajustá las cantidades o desactivá el modo estricto."
                )
    else:
        total = sum(a.tamano for a in pool)
        resto = total % k
        if resto != 0:
            raise ErrorValidacion(
                f"Grupo '{nombre_grupo}': modo estricto exige habitaciones completas de {k} estudiantes, "
                f"pero el total es {total} (sobran {resto}). Ajustá cantidades o desactivá el modo estricto."
            )


def _empaquetar_estudiantes(
    pool: list[Atomo],
    capacidades_disponibles: Counter,
    config: ConfigGrupo,
    nombre_grupo: str,
    cupo: int | None = None,
    tope_plazas: int | None = None,
    avisos: list[str] | None = None,
) -> tuple[list[HabitacionVirtual], Counter, list[Atomo]]:
    if cupo is None or not pool:
        return [], Counter(capacidades_disponibles), []

    avisos = avisos if avisos is not None else []
    _validar_divisibilidad_estudiantes(pool, cupo, config, nombre_grupo)

    disponibles = Counter(capacidades_disponibles)
    bins: list[EstadoHabitacion] = []
    ordenados = sorted(pool, key=lambda a: (-a.tamano, a.descripcion))
    relajado = False

    for atomo in ordenados:
        destino = next(
            (
                b
                for b in bins
                if b.carga + atomo.tamano <= cupo and b.puede_recibir(atomo, config)
            ),
            None,
        )
        if destino is not None:
            destino.recibir(atomo)
            continue
        needing = max(cupo, atomo.tamano)
        aptas = sorted(
            c for c in disponibles
            if disponibles[c] > 0 and c >= needing and (tope_plazas is None or c <= tope_plazas)
        )
        if not aptas and tope_plazas is not None:
            aptas = sorted(c for c in disponibles if disponibles[c] > 0 and c >= needing)
            if aptas and not relajado:
                relajado = True
                avisos.append(
                    f"Grupo '{nombre_grupo}': no quedan habitaciones de {tope_plazas} plazas o menos; "
                    f"se usarán las de {aptas[0]} plazas para los estudiantes."
                )
        if not aptas:
            detalle = ", ".join(f"{c} lugares x{disponibles[c]}" for c in sorted(disponibles)) or "ninguna"
            raise ErrorValidacion(
                f"Grupo '{nombre_grupo}': se piden habitaciones de {cupo} estudiantes pero no queda ninguna "
                f"habitación de capacidad >= {needing} libre en esas fechas "
                f"(disponible: {detalle})."
            )
        cap_real = aptas[0]
        disponibles[cap_real] -= 1
        nuevo = EstadoHabitacion(min(cap_real, cupo))
        nuevo.recibir(atomo)
        bins.append(nuevo)

    bins.sort(key=lambda b: (-b.carga, -b.capacidad))
    virtuales = [HabitacionVirtual(capacidad=b.capacidad, atomos=list(b.atomos)) for b in bins]
    return virtuales, disponibles, []


def _distribuir_adultos(
    atomos: list[Atomo],
    grupo: Grupo,
    caps_trabajo: list[int],
    opciones: OpcionesMotor,
) -> ResultadoEmpaque:
    """Reparto libre (receptivo, familias, parejas, estudiantes sueltos).
    Es exactamente el comportamiento previo, sin cohortes ni topes."""
    config = grupo.config
    pool_est, resto = _separar_estudiantes(atomos, config)
    bins_estudiantes, caps_restantes, _ = _empaquetar_estudiantes(
        pool_est, Counter(caps_trabajo), config, grupo.nombre,
        cupo=config.estudiantes_por_habitacion,
    )

    if not resto:
        return ResultadoEmpaque(
            habitaciones=bins_estudiantes,
            pendientes=[],
            estrategia="CUPO-FIJO" if pool_est else "-",
            optimo=True,
            tiempo=0.0,
        )

    empaque = empaquetar(
        resto,
        list(caps_restantes.elements()),
        config,
        estrategia=opciones.estrategia,
        semilla=opciones.semilla,
        tiempo_max=opciones.tiempo_max_por_grupo,
    )
    return ResultadoEmpaque(
        habitaciones=bins_estudiantes + empaque.habitaciones,
        pendientes=list(empaque.pendientes),
        estrategia=empaque.estrategia,
        optimo=empaque.optimo,
        tiempo=empaque.tiempo,
    )


def _distribuir_estudiantil(
    atomos: list[Atomo],
    grupo: Grupo,
    caps_trabajo: list[int],
    avisos: list[str],
) -> ResultadoEmpaque:
    """Turismo estudiantil: los alumnos van a habitaciones del cupo, sin mezclarse con
    los acompanantes, y los adultos van de a cohortes (mismo rol y sexo) en la habitacion
    mas justa que entre."""
    config = grupo.config
    alumnos = [a for a in atomos if _es_alumno(a)]
    adultos = [a for a in atomos if not _es_alumno(a)]

    cupo = config.estudiantes_por_habitacion
    tope = config.max_plazas_estudiantes
    if cupo is not None and tope is not None and cupo > tope:
        avisos.append(
            f"Grupo '{grupo.nombre}': el cupo de {cupo} estudiantes por habitación supera el máximo "
            f"de {tope} plazas; los alumnos se alojan de a {tope}."
        )
        cupo = tope
    elif cupo is None and tope is not None:
        cupo = tope

    disponibles = Counter(caps_trabajo)
    bins: list[HabitacionVirtual] = []
    pendientes: list[Atomo] = []

    if alumnos:
        nuevos, disponibles, colgados = _empaquetar_estudiantes(
            alumnos, disponibles, config, grupo.nombre,
            cupo=cupo, tope_plazas=tope, avisos=avisos,
        )
        bins += nuevos
        pendientes += colgados

    for cohorte in _cohortes_adultas(adultos, config):
        nuevos, disponibles, colgados = _empaquetar_cohorte(cohorte, disponibles, config)
        bins += nuevos
        pendientes += colgados

    etiqueta = f"ESTUDIANTIL (alumnos de a {cupo})" if cupo else "ESTUDIANTIL"
    return ResultadoEmpaque(
        habitaciones=bins,
        pendientes=pendientes,
        estrategia=etiqueta,
        optimo=False,
        tiempo=0.0,
    )


def _asignar_fisicas(
    bins_virtuales: list[HabitacionVirtual],
    libres: list[Habitacion],
) -> list[AsignacionHabitacion]:
    orden_bins = sorted(bins_virtuales, key=lambda b: (-b.carga, -b.capacidad))
    disponibles = sorted(libres, key=lambda h: (h.capacidad, clave_orden_habitacion(h.numero)))
    usadas: set[str] = set()
    salida: list[AsignacionHabitacion] = []
    for bin_virtual in orden_bins:
        elegida = next(
            (h for h in disponibles if h.numero not in usadas and h.capacidad >= bin_virtual.carga),
            None,
        )
        if elegida is None:
            raise ErrorDistribucion(
                "Error interno: el empaquetado requiere más habitaciones físicas de las libres. "
                "Por favor informá este error."
            )
        usadas.add(elegida.numero)
        salida.append(AsignacionHabitacion(habitacion=elegida, atomos=list(bin_virtual.atomos)))
    salida.sort(key=lambda a: (clave_orden_habitacion(a.habitacion.numero)))
    return salida


def _distribuir_grupo(
    hotel: Hotel,
    grupo: Grupo,
    sesion: dict[str, list[ReservaSesion]],
    opciones: OpcionesMotor,
    solo_habitaciones: list[str] | None = None,
) -> DistribucionGrupo:
    ini, fin = grupo.fecha_ingreso, grupo.fecha_egreso
    libres = [h for h in hotel.habitaciones
              if _habitacion_libre(h, ini, fin, sesion, ignorar_grupo=grupo.nombre)]
    if solo_habitaciones is not None:
        fijas = {n.strip().lower() for n in solo_habitaciones}
        libres = [h for h in libres if h.numero.strip().lower() in fijas]

    mapa = opciones.capacidades_efectivas or {}
    caps_trabajo: list[int] = []
    for h in libres:
        cap = h.capacidad
        if cap in mapa:
            cap = mapa[cap]
        caps_trabajo.append(cap)

    dist = DistribucionGrupo(grupo=grupo)

    try:
        atomos, avisos = atomizar(grupo)
        dist.avisos.extend(avisos)
    except ErrorDistribucion as exc:
        dist.faltantes = list(grupo.pax)
        dist.avisos.append(f"No se pudo distribuir: {exc}")
        dist.deficit = Counter()
        return dist

    if not libres and atomos:
        caps_hotel = sorted({h.capacidad for h in hotel.habitaciones}, reverse=True)
        try:
            dist.deficit = calcular_deficit(atomos, caps_hotel, grupo.config)
        except ErrorDistribucion as exc:
            dist.deficit = Counter()
            dist.avisos.append(str(exc))
        dist.faltantes = [m for a in atomos for m in a.miembros]
        dist.avisos.insert(
            0,
            f"No hay ninguna habitación libre entre {ini} y {fin}. "
            "Revisá bloqueos, ventanas de disponibilidad u otros grupos.",
        )
        return dist

    resultado_final = ResultadoEmpaque(habitaciones=[], pendientes=[], estrategia="")
    avisos_modo: list[str] = []
    try:
        modo = _modo_efectivo(atomos, grupo.config)
        if modo == ModoHospedaje.ESTUDIANTIL.value:
            resultado_final = _distribuir_estudiantil(atomos, grupo, caps_trabajo, avisos_modo)
        else:
            resultado_final = _distribuir_adultos(atomos, grupo, caps_trabajo, opciones)
        dist.avisos.extend(avisos_modo)
    except ErrorValidacion:
        raise
    except ErrorDistribucion as exc:
        dist.avisos.append(str(exc))
        dist.faltantes = [m for a in atomos for m in a.miembros]
        caps_hotel = sorted({caps_trabajo[i] for i in range(len(libres))}) or sorted(
            {h.capacidad for h in hotel.habitaciones}
        )
        try:
            dist.deficit = calcular_deficit(atomos, caps_hotel, grupo.config)
        except ErrorDistribucion:
            dist.deficit = Counter()
        return dist

    if resultado_final.habitaciones:
        dist.asignaciones = _asignar_fisicas(resultado_final.habitaciones, libres)

    if resultado_final.pendientes:
        dist.faltantes = [m for a in resultado_final.pendientes for m in a.miembros]
        caps_distintas = sorted(set(caps_trabajo), reverse=True)
        try:
            dist.deficit = calcular_deficit(resultado_final.pendientes, caps_distintas, grupo.config)
        except ErrorDistribucion as exc:
            dist.deficit = Counter()
            dist.avisos.append(str(exc))
        dist.avisos.append(
            f"{len(dist.faltantes)} pasajero(s) sin habitación: "
            + ", ".join(p.nombre for p in dist.faltantes)
        )

    dist.estrategia_usada = resultado_final.estrategia
    dist.optimo = resultado_final.optimo
    dist.tiempo = resultado_final.tiempo
    return dist


def distribuir(
    hotel: Hotel,
    grupos: list[Grupo],
    opciones: OpcionesMotor | None = None,
    previo: ReporteGeneral | None = None,
    reoptimizar: bool = False,
) -> ReporteGeneral:
    """Distribuye grupos en las habitaciones del hotel.

    - previo: reporte anterior. Si un grupo está `confirmado`, conserva
      sus habitaciones actuales (congeladas) y los demás se acomodan en lo libre.
    - reoptimizar=True: ignora confirmaciones y reacomoda todo desde cero.
    """
    opciones = opciones or OpcionesMotor()

    errores = hotel.validar()
    for g in grupos:
        errores.extend(g.validar())
    errores.extend(opciones.validar(hotel))
    if errores:
        raise ErrorDistribucion(
            "No se puede ejecutar la distribución por estos motivos:\n- " + "\n- ".join(errores)
        )

    nombres_por_fecha: dict[tuple[date, date], set[str]] = {}
    con_fechas = [g for g in grupos if g.tiene_fechas]
    sin_fechas = [g for g in grupos if not g.tiene_fechas]
    for g in con_fechas:
        clave = (g.fecha_ingreso, g.fecha_egreso)
        for p in g.pax:
            nombres_por_fecha.setdefault(clave, set()).add(p.nombre.strip().lower())
    avisos_duplicados = []
    claves_ordenadas = sorted(nombres_por_fecha.keys())
    for i in range(len(claves_ordenadas)):
        ini_a, fin_a = claves_ordenadas[i]
        nombres_a = nombres_por_fecha[claves_ordenadas[i]]
        for j in range(i + 1, len(claves_ordenadas)):
            ini_b, fin_b = claves_ordenadas[j]
            if not (ini_a < fin_b and ini_b < fin_a):
                continue
            comunes = nombres_a & nombres_por_fecha[claves_ordenadas[j]]
            if comunes:
                avisos_duplicados.append(
                    f"Estas personas aparecen en dos grupos con fechas solapadas "
                    f"({ini_a}..{fin_a} y {ini_b}..{fin_b}): " + ", ".join(sorted(comunes))
                )

    habitaciones_fijas: dict[str, list[str]] = {}
    pre_confirmados: dict[str, bool] = {g.nombre: g.confirmado for g in grupos}
    if previo is not None and not reoptimizar:
        for d in previo.distribuciones:
            if pre_confirmados.get(d.grupo.nombre):
                habitaciones_fijas[d.grupo.nombre] = [
                    a.habitacion.numero for a in d.asignaciones
                ]

    sesion: dict[str, list[ReservaSesion]] = {}
    for grupo in grupos:
        nums = habitaciones_fijas.get(grupo.nombre)
        if not grupo.confirmado or not nums:
            continue
        if not grupo.tiene_fechas:
            continue
        for numero in nums:
            sesion.setdefault(numero, []).append(
                ReservaSesion(desde=grupo.fecha_ingreso, hasta=grupo.fecha_egreso, grupo=grupo.nombre)
            )

    con_fechas = [g for g in grupos if g.tiene_fechas]
    sin_fechas = [g for g in grupos if not g.tiene_fechas]

    ordenados = sorted(
        con_fechas,
        key=lambda g: (g.fecha_ingreso, g.nombre.lower()),
    )
    distribuciones: list[DistribucionGrupo] = []

    for grupo in ordenados:
        solo = habitaciones_fijas.get(grupo.nombre) if grupo.confirmado else None
        dist = _distribuir_grupo(hotel, grupo, sesion, opciones, solo_habitaciones=solo)
        distribuciones.append(dist)
        for asig in dist.asignaciones:
            lista = sesion.setdefault(asig.habitacion.numero, [])
            lista.append(
                ReservaSesion(desde=grupo.fecha_ingreso, hasta=grupo.fecha_egreso, grupo=grupo.nombre)
            )

    reporte = ReporteGeneral(hotel=hotel, distribuciones=distribuciones)
    reporte.avisos.extend(avisos_duplicados)
    if sin_fechas:
        reporte.avisos.append(
            "Grupo(s) sin fecha (cola de espera, no se distribuyen): "
            + "; ".join(g.nombre for g in sin_fechas)
            + "."
        )

    if distribuciones:
        fecha_min = min(d.grupo.fecha_ingreso for d in distribuciones)
        fecha_max = max(d.grupo.fecha_egreso for d in distribuciones)
        for noche in noches_entdesde_hasta(fecha_min, fecha_max):
            pax_noche = 0
            hab_noche = 0
            for d in distribuciones:
                if d.grupo.fecha_ingreso <= noche < d.grupo.fecha_egreso:
                    pax_noche += d.pax_alojados
                    hab_noche += d.habitaciones_usadas
            reporte.pax_por_noche[noche] = pax_noche
            reporte.habitaciones_por_noche[noche] = hab_noche

    violaciones = verificar_reporte(reporte)
    if violaciones:
        raise ErrorDistribucion(
            "Error interno del algoritmo (invariantes violadas):\n- " + "\n- ".join(violaciones)
        )

    return reporte


def verificar_reporte(reporte: ReporteGeneral) -> list[str]:
    violaciones: list[str] = []
    for dist in reporte.distribuciones:
        grupo = dist.grupo
        cfg = grupo.config
        etiqueta = f"[{grupo.nombre}]"

        esperados = {p.nombre.strip().lower(): p for p in grupo.pax}
        alojados: set[str] = set()
        for asig in dist.asignaciones:
            hab = asig.habitacion
            carga = asig.carga
            if carga > hab.capacidad:
                violaciones.append(
                    f"{etiqueta} habitación {hab.numero}: capacidad excedida ({carga}/{hab.capacidad})."
                )

            singles = sum(1 for a in asig.atomos if a.requiere_single)
            if singles >= 1 and len(asig.atomos) > 1:
                violaciones.append(
                    f"{etiqueta} habitación {hab.numero}: pasajero que exige single comparte habitación."
                )

            sexos = {s for a in asig.atomos for s in a.sexos_no_x}
            if cfg.separar_sexos and len(sexos) > 1 and len(asig.atomos) > 1:
                violaciones.append(
                    f"{etiqueta} habitación {hab.numero}: sexos mezclados sin vínculo familiar/pareja."
                )

            menores = sum(a.menores for a in asig.atomos)
            adultos = any(a.tiene_adulto for a in asig.atomos)
            if menores > 0 and not adultos and cfg.ninos_con_adulto_obligatorio:
                violaciones.append(f"{etiqueta} habitación {hab.numero}: hay menores sin adulto.")
            if cfg.max_ninos_por_habitacion is not None and menores > cfg.max_ninos_por_habitacion:
                violaciones.append(
                    f"{etiqueta} habitación {hab.numero}: {menores} menores superan el máximo "
                    f"de {cfg.max_ninos_por_habitacion}."
                )

            for nombre in asig.nombres_pax:
                clave = nombre.strip().lower()
                if clave in alojados:
                    violaciones.append(f"{etiqueta}: '{nombre}' aparece en más de una habitación.")
                alojados.add(clave)

        faltantes_claves = {p.nombre.strip().lower() for p in dist.faltantes}
        interseccion = alojados & faltantes_claves
        if interseccion:
            violaciones.append(f"{etiqueta}: pasajeros simultáneamente alojados y faltantes: {interseccion}.")
        faltan_desconocidos = faltantes_claves - set(esperados)
        if faltan_desconocidos:
            violaciones.append(f"{etiqueta}: faltantes que no pertenecen al grupo: {faltan_desconocidos}.")

        total_cuenta = len(alojados) + len(faltantes_claves - alojados)
        if total_cuenta != len(esperados):
            perdidos = set(esperados) - alojados - faltantes_claves
            if perdidos:
                violaciones.append(f"{etiqueta}: pasajeros desaparecidos del cálculo: {sorted(perdidos)}.")

    ocupadas: dict[str, list[tuple[date, date, str]]] = {}
    for dist in reporte.distribuciones:
        for asig in dist.asignaciones:
            numero = asig.habitacion.numero
            intervalo = (dist.grupo.fecha_ingreso, dist.grupo.fecha_egreso, dist.grupo.nombre)
            for prev_ini, prev_fin, prev_grupo in ocupadas.get(numero, []):
                if intervalo[0] < prev_fin and prev_ini < intervalo[1]:
                    violaciones.append(
                        f"Habitación {numero} doblemente asignada: '{prev_grupo}' y '{dist.grupo.nombre}' "
                        f"solapan entre {intervalo[0]} y {intervalo[1]}."
                    )
            ocupadas.setdefault(numero, []).append(intervalo)

    return violaciones


@dataclass
class SugerenciaMezcla:
    mezcla: Counter = field(default_factory=Counter)
    pendientes: list[Atomo] = field(default_factory=list)
    avisos: list[str] = field(default_factory=list)
    estrategia: str = ""
    pax_total: int = 0

    @property
    def completo(self) -> bool:
        return not self.pendientes

    @property
    def habitaciones(self) -> int:
        return sum(self.mezcla.values())

    @property
    def camas(self) -> int:
        return sum(cap * cantidad for cap, cantidad in self.mezcla.items())

    @property
    def pax_alojados(self) -> int:
        return self.pax_total - sum(a.tamano for a in self.pendientes)

    @property
    def desperdicio(self) -> int:
        return self.camas - self.pax_alojados

    def texto(self) -> str:
        partes = [
            f"{cantidad} x {cap} plazas" for cap, cantidad in sorted(self.mezcla.items(), reverse=True)
        ]
        return ", ".join(partes) or "ninguna"

    def mas_capacidad(self) -> int:
        return max(self.mezcla, default=0)


def distribucion_estandar(
    total_habitaciones: int,
    capacidad_max: int,
    ratio: float | None = None,
) -> list[int]:
    """Genera capacidades balanceadas para un inventario del que solo conocemos
    la cantidad total de apartamentos, la capacidad máxima y (opcional) el ratio
    camas/apartamento del hotel."""
    if total_habitaciones < 1:
        raise ErrorValidacion("La cantidad de apartamentos debe ser >= 1.")
    if capacidad_max < 1:
        raise ErrorValidacion("La capacidad máxima por apartamento debe ser >= 1.")
    base = max(1, min(int(ratio or capacidad_max), capacidad_max))
    if ratio is None:
        deseado = base * total_habitaciones
        if base < capacidad_max:
            deseado += round(total_habitaciones * 0.35)
    else:
        deseado = int(round(total_habitaciones * ratio))
    deseado = max(total_habitaciones, min(deseado, total_habitaciones * capacidad_max))
    caps = [base] * total_habitaciones
    camas = sum(caps)

    def subir() -> bool:
        menor_idx = min(range(total_habitaciones), key=lambda i: (caps[i], i))
        if caps[menor_idx] >= capacidad_max:
            return False
        caps[menor_idx] += 1
        return True

    while camas < deseado:
        if not subir():
            break
        camas += 1
    while camas > deseado:
        mayor_idx = max(range(total_habitaciones), key=lambda i: (caps[i], -i))
        if caps[mayor_idx] <= 1:
            break
        caps[mayor_idx] -= 1
        camas -= 1
    return caps


def generar_hotel_estimado(
    nombre: str,
    total_habitaciones: int,
    capacidad_max: int,
    ratio: float | None = None,
    desde: date | None = None,
    hasta: date | None = None,
) -> Hotel:
    """Construye un hotel ficticio con capacidades balanceadas cuando no se conoce
    la composición exacta del inventario (solo total de apartamentos y capacidad máxima)."""
    caps = distribucion_estandar(total_habitaciones, capacidad_max, ratio)
    hotel = Hotel(nombre=nombre)
    for i, cap in enumerate(caps, start=1):
        hotel.habitaciones.append(
            Habitacion(
                numero=f"{100 + i}",
                capacidad=cap,
                disponible_desde=desde,
                disponible_hasta=hasta,
            )
        )
    return hotel


def crear_grupo_generico(
    nombre: str,
    fecha_ingreso: date,
    fecha_egreso: date,
    total_pax: int,
    varones: int = 0,
    mujeres: int = 0,
    config: ConfigGrupo | None = None,
    prefijo: str = "Pax",
) -> Grupo:
    """Crea un grupo con pasajeros genéricos (sin nombres reales) para planificar
    cuando solo se conoce la cantidad de personas y las fechas."""
    if total_pax < 1:
        raise ErrorValidacion(f"Grupo '{nombre}': la cantidad de pasajeros debe ser >= 1.")
    if varones + mujeres > total_pax:
        raise ErrorValidacion(
            f"Grupo '{nombre}': la suma de varones ({varones}) y mujeres ({mujeres}) "
            f"supera el total ({total_pax})."
        )
    sexo_x = total_pax - varones - mujeres
    pax: list[Pax] = []
    n = 0
    for _ in range(varones):
        n += 1
        pax.append(Pax(nombre=f"{prefijo} {n:03d}", tipo=TipoPax.ADULTO, sexo=Sexo.M))
    for _ in range(mujeres):
        n += 1
        pax.append(Pax(nombre=f"{prefijo} {n:03d}", tipo=TipoPax.ADULTO, sexo=Sexo.F))
    for _ in range(sexo_x):
        n += 1
        pax.append(Pax(nombre=f"{prefijo} {n:03d}", tipo=TipoPax.ADULTO, sexo=Sexo.X))
    return Grupo(
        nombre=nombre,
        fecha_ingreso=fecha_ingreso,
        fecha_egreso=fecha_egreso,
        pax=pax,
        config=config or ConfigGrupo(separar_sexos=bool(varones and mujeres)),
    )


def sugerir_mezcla(
    atomos: list[Atomo],
    capacidades_disponibles: list[int],
    config: ConfigGrupo,
    total_habitaciones: int | None = None,
    estrategia: str = Estrategia.BFD.value,
) -> SugerenciaMezcla:
    """Calcula cuántas habitaciones de cada capacidad harían falta para alojar
    los átomos dados, optimizando la mezcla dentro de las capacidades conocidas.

    - Sin `total_habitaciones`: inventario ilimitado; devuelve la mezcla óptima
      (mínimas habitaciones y mínimo desperdicio).
    - Con `total_habitaciones`: respeta ese techo y reporta pendientes si no
      alcanza (déficit).
    """
    caps = sorted({c for c in capacidades_disponibles if c >= 1}, reverse=True)
    sugerencia = SugerenciaMezcla(pax_total=sum(a.tamano for a in atomos))
    if not atomos:
        return sugerencia
    if not caps:
        raise ErrorDistribucion("No hay capacidades de habitación definidas.")
    cap_maxima = caps[0]
    excedidos = [a for a in atomos if a.tamano > cap_maxima]
    validos = [a for a in atomos if a.tamano <= cap_maxima]
    if excedidos and not validos:
        sugerencia.pendientes = list(excedidos)
        nombres = "; ".join(f"{a.descripcion} ({a.tamano} personas)" for a in excedidos[:5])
        sugerencia.avisos.append(
            f"Ninguna habitación puede alojar a los grupos más grandes: {nombres}"
            f" (capacidad máxima: {cap_maxima})."
        )
        return sugerencia

    if total_habitaciones is not None:
        if total_habitaciones < 1:
            raise ErrorValidacion("La cantidad de apartamentos debe ser >= 1.")
        capacidades_trabajo = list(capacidades_disponibles)[:total_habitaciones]
        while len(capacidades_trabajo) < total_habitaciones:
            capacidades_trabajo.append(caps[0])
    else:
        veces = len(validos) + 1
        capacidades_trabajo = []
        for c in caps:
            capacidades_trabajo.extend([c] * veces)

    emp = empaquetar(
        validos,
        capacidades_trabajo,
        config,
        estrategia=estrategia,
        tiempo_max=3.0,
    )
    sugerencia.estrategia = emp.estrategia
    sugerencia.mezcla = Counter(h.capacidad for h in emp.habitaciones)
    sugerencia.pendientes = list(excedidos) + list(emp.pendientes)
    if emp.pendientes:
        sugerencia.avisos.append(
            f"{len(emp.pendientes)} pasajero(s) quedan sin habitación, "
            f"faltarían: " + _mix_texto_faltante(emp.pendientes, caps)
        )
    return sugerencia


def _mix_texto_faltante(pendientes: list[Atomo], caps: list[int]) -> str:
    from empaquetador import calcular_deficit

    try:
        deficit = calcular_deficit(pendientes, caps, ConfigGrupo())
    except ErrorDistribucion:
        return "capacidades insuficientes"
    return ", ".join(
        f"{cantidad} x {cap} plazas" for cap, cantidad in sorted(deficit.items(), reverse=True)
    )
