from __future__ import annotations

import re
from dataclasses import dataclass, field
from datetime import date, timedelta
from enum import Enum


class ErrorDistribucion(Exception):
    pass


class ErrorValidacion(ErrorDistribucion):
    pass


class TipoPax(str, Enum):
    ADULTO = "ADULTO"
    NINO = "NINO"
    BEBE = "BEBE"
    ESTUDIANTE = "ESTUDIANTE"
    GUIA = "GUIA"
    COORDINADOR = "COORDINADOR"
    CHOFER = "CHOFER"


ALIAS_TIPO = {
    "ADULTO": TipoPax.ADULTO, "ADULTA": TipoPax.ADULTO, "A": TipoPax.ADULTO,
    "NINO": TipoPax.NINO, "NIÑO": TipoPax.NINO, "NENA": TipoPax.NINO,
    "NIÑA": TipoPax.NINO, "MENOR": TipoPax.NINO, "CHICO": TipoPax.NINO, "N": TipoPax.NINO,
    "BEBE": TipoPax.BEBE, "BEBÉ": TipoPax.BEBE, "BB": TipoPax.BEBE, "B": TipoPax.BEBE,
    "ESTUDIANTE": TipoPax.ESTUDIANTE, "EST": TipoPax.ESTUDIANTE, "ESTUDIANTES": TipoPax.ESTUDIANTE, "E": TipoPax.ESTUDIANTE,
    "GUIA": TipoPax.GUIA, "GUÍA": TipoPax.GUIA, "G": TipoPax.GUIA,
    "COORDINADOR": TipoPax.COORDINADOR, "COORD": TipoPax.COORDINADOR, "C": TipoPax.COORDINADOR,
    "CHOFER": TipoPax.CHOFER, "CHOFER": TipoPax.CHOFER, "DRIVER": TipoPax.CHOFER, "D": TipoPax.CHOFER,
}


def tipo_desde_texto(texto: str) -> TipoPax:
    clave = texto.strip().upper()
    if clave in ALIAS_TIPO:
        return ALIAS_TIPO[clave]
    raise ErrorDistribucion(
        f"Tipo de pasajero desconocido: '{texto}'. Opciones válidas: "
        "ADULTO, NINO, BEBE, ESTUDIANTE, GUIA, COORDINADOR, CHOFER."
    )


class Sexo(str, Enum):
    M = "M"
    F = "F"
    X = "X"


ALIAS_SEXO = {"M": Sexo.M, "MASculino".upper(): Sexo.M, "HOMBRE": Sexo.M, "VARON": Sexo.M, "VARÓN": Sexo.M,
              "F": Sexo.F, "FEMENINO": Sexo.F, "MUJER": Sexo.F,
              "X": Sexo.X, "INDISTINTO": Sexo.X, "-": Sexo.X, "": Sexo.X, "S/D": Sexo.X}


def sexo_desde_texto(texto: str) -> Sexo:
    clave = texto.strip().upper()
    if clave in ALIAS_SEXO:
        return ALIAS_SEXO[clave]
    raise ErrorDistribucion(f"Sexo desconocido: '{texto}'. Usá M, F o X (indistinto).")


TIPOS_TRIPULACION = frozenset({TipoPax.GUIA, TipoPax.COORDINADOR, TipoPax.CHOFER})
TIPOS_MENORES = frozenset({TipoPax.NINO, TipoPax.BEBE})


class ModoHospedaje(str, Enum):
    """Como se reparte un grupo. El modo estudiantil separa alumnos de acompanantes
    y agrupa a los adultos por rol y sexo; el modo adultos deja el reparto libre
    (receptivo, familias, parejas)."""

    ESTUDIANTIL = "ESTUDIANTIL"
    ADULTOS = "ADULTOS"


@dataclass
class Pax:
    nombre: str
    tipo: TipoPax = TipoPax.ADULTO
    sexo: Sexo = Sexo.X
    edad: int | None = None
    pareja_de: str | None = None
    familia: str | None = None
    requiere_single: bool = False
    notas: str = ""

    def es_menor(self) -> bool:
        return self.tipo in TIPOS_MENORES

    def es_adulto(self) -> bool:
        return not self.es_menor()

    def descripcion(self) -> str:
        partes = [self.nombre, self.tipo.value]
        if self.sexo != Sexo.X:
            partes.append(self.sexo.value)
        if self.edad is not None:
            partes.append(f"{self.edad}a")
        if self.requiere_single:
            partes.append("[single]")
        if self.familia:
            partes.append(f"[familia {self.familia}]")
        if self.pareja_de:
            partes.append(f"[pareja de {self.pareja_de}]")
        if self.notas:
            partes.append(f"({self.notas})")
        return " ".join(partes)


@dataclass
class Bloqueo:
    desde: date
    hasta: date
    titulo: str = ""

    def solapa_con(self, ini: date, fin: date) -> bool:
        return ini < self.hasta and self.desde < fin


@dataclass
class Habitacion:
    numero: str
    capacidad: int
    piso: str = ""
    disponible_desde: date | None = None
    disponible_hasta: date | None = None
    bloqueos: list[Bloqueo] = field(default_factory=list)

    def validar(self) -> list[str]:
        errores = []
        if not self.numero or not self.numero.strip():
            errores.append("Habitación sin número.")
        if self.capacidad < 1:
            errores.append(f"Habitación {self.numero}: capacidad debe ser >= 1 (tiene {self.capacidad}).")
        if self.disponible_desde and self.disponible_hasta and self.disponible_desde >= self.disponible_hasta:
            errores.append(
                f"Habitación {self.numero}: ventana de disponibilidad inválida "
                f"({self.disponible_desde} a {self.disponible_hasta}); la fecha hasta debe ser posterior."
            )
        for b in self.bloqueos:
            if b.desde >= b.hasta:
                errores.append(f"Habitación {self.numero}: bloqueo inválido {b.desde} a {b.hasta}.")
        return errores

    def libre_en_rango(self, ini: date, fin: date) -> bool:
        if fin <= ini:
            return False
        if self.disponible_desde and ini < self.disponible_desde:
            return False
        if self.disponible_hasta and fin > self.disponible_hasta:
            return False
        return not any(b.solapa_con(ini, fin) for b in self.bloqueos)

    def agregar_bloqueo(self, desde: date, hasta: date, titulo: str = "") -> None:
        if desde >= hasta:
            raise ErrorDistribucion(f"Bloqueo inválido en habitación {self.numero}: {desde} a {hasta}.")
        for b in self.bloqueos:
            if b.solapa_con(desde, hasta):
                raise ErrorDistribucion(
                    f"Habitación {self.numero} ya está bloqueada de {b.desde} a {b.hasta}"
                    + (f" ({b.titulo})" if b.titulo else "")
                    + f"; no se puede agregar {desde} a {hasta}."
                )
        self.bloqueos.append(Bloqueo(desde=desde, hasta=hasta, titulo=titulo))
        self.bloqueos.sort(key=lambda x: x.desde)

    def quitar_bloqueo(self, desde: date, hasta: date, titulo: str = "") -> bool:
        for i, b in enumerate(self.bloqueos):
            if b.desde == desde and b.hasta == hasta and (not titulo or b.titulo == titulo):
                del self.bloqueos[i]
                return True
        return False


@dataclass
class ConfigGrupo:
    separar_sexos: bool = True
    ninos_con_adulto_obligatorio: bool = True
    guias_choferes_single: bool = False
    estudiantes_por_habitacion: int | None = None
    estudiantes_estricto: bool = False
    max_ninos_por_habitacion: int | None = None
    familias_exentas_sexo: bool = True
    bebes_no_cuentan_cupo: bool = False
    modo_hospedaje: str = ModoHospedaje.ESTUDIANTIL.value
    max_plazas_estudiantes: int | None = 4
    adultos_por_rol: bool = True

    @property
    def es_estudiantil(self) -> bool:
        return self.modo_hospedaje == ModoHospedaje.ESTUDIANTIL.value

    def validar(self) -> list[str]:
        errores = []
        if self.estudiantes_por_habitacion is not None and self.estudiantes_por_habitacion < 1:
            errores.append("estudiantes_por_habitacion debe ser >= 1 o vacío.")
        if self.max_ninos_por_habitacion is not None and self.max_ninos_por_habitacion < 0:
            errores.append("max_ninos_por_habitacion debe ser >= 0 o vacío.")
        if self.modo_hospedaje not in {m.value for m in ModoHospedaje}:
            errores.append(
                f"modo_hospedaje inválido '{self.modo_hospedaje}'. "
                f"Opciones: {', '.join(m.value for m in ModoHospedaje)}."
            )
        if self.max_plazas_estudiantes is not None and self.max_plazas_estudiantes < 1:
            errores.append("max_plazas_estudiantes debe ser >= 1 o vacío.")
        return errores


@dataclass
class Grupo:
    nombre: str
    fecha_ingreso: date | None = None
    fecha_egreso: date | None = None
    pax: list[Pax] = field(default_factory=list)
    config: ConfigGrupo = field(default_factory=ConfigGrupo)
    empresa: str = ""
    provisional: bool = False
    confirmado: bool = False

    def validar(self) -> list[str]:
        errores = []
        if not self.nombre or not self.nombre.strip():
            errores.append("Grupo sin nombre.")
        if (self.fecha_ingreso is None) != (self.fecha_egreso is None):
            errores.append(
                f"Grupo '{self.nombre}': hay que definir las dos fechas (ingreso y egreso) "
                "o ninguna (queda en 'sin fecha', cola de espera)."
            )
        if (self.fecha_ingreso is not None and self.fecha_egreso is not None
                and self.fecha_egreso <= self.fecha_ingreso):
            errores.append(
                f"Grupo '{self.nombre}': fecha de egreso ({self.fecha_egreso}) debe ser posterior "
                f"a la de ingreso ({self.fecha_ingreso})."
            )
        errores.extend(self.config.validar())
        if not self.pax:
            errores.append(f"Grupo '{self.nombre}' no tiene pasajeros cargados.")
        nombres_vistos = set()
        for p in self.pax:
            clave = p.nombre.strip().lower()
            if not clave:
                errores.append(f"Grupo '{self.nombre}': hay un pasajero sin nombre.")
            elif clave in nombres_vistos:
                errores.append(f"Grupo '{self.nombre}': nombre duplicado '{p.nombre}'.")
            else:
                nombres_vistos.add(clave)
        return errores

    @property
    def tiene_fechas(self) -> bool:
        return self.fecha_ingreso is not None and self.fecha_egreso is not None

    def noches(self) -> int:
        if not self.tiene_fechas:
            return 0
        return (self.fecha_egreso - self.fecha_ingreso).days

    def total_pax(self) -> int:
        return len(self.pax)


@dataclass
class LoteCamas:
    """Lote de camas del inventario independiente (no ligadas a apartamentos)."""
    titulo: str
    cantidad: int
    disponible_desde: date | None = None
    disponible_hasta: date | None = None

    def validar(self) -> list[str]:
        errores = []
        if self.cantidad < 1:
            errores.append(f"Lote '{self.titulo}': la cantidad de camas debe ser >= 1.")
        if (self.disponible_desde and self.disponible_hasta
                and self.disponible_desde >= self.disponible_hasta):
            errores.append(
                f"Lote '{self.titulo}': ventana inválida "
                f"({self.disponible_desde} a {self.disponible_hasta}); "
                "la fecha hasta debe ser posterior.")
        return errores

    def activo_en(self, dia: date) -> bool:
        if self.disponible_desde and dia < self.disponible_desde:
            return False
        if self.disponible_hasta and dia > self.disponible_hasta:
            return False
        return True


@dataclass
class Hotel:
    nombre: str = "Mi Hotel"
    habitaciones: list[Habitacion] = field(default_factory=list)
    lotes_camas: list[LoteCamas] = field(default_factory=list)
    provisional: dict | None = field(default=None)

    def validar(self) -> list[str]:
        errores = []
        numeros = set()
        for h in self.habitaciones:
            errores.extend(h.validar())
            clave = h.numero.strip().lower()
            if clave in numeros:
                errores.append(f"Número de habitación duplicado: '{h.numero}'.")
            numeros.add(clave)
        for l in self.lotes_camas:
            errores.extend(l.validar())
        return errores

    def habitacion(self, numero: str) -> Habitacion | None:
        clave = numero.strip().lower()
        for h in self.habitaciones:
            if h.numero.strip().lower() == clave:
                return h
        return None

    def total_camas(self) -> int:
        return sum(h.capacidad for h in self.habitaciones)

    def camas_inventario_en(self, dia: date) -> int:
        return sum(l.cantidad for l in self.lotes_camas if l.activo_en(dia))


def parsear_fecha(texto: str) -> date:
    limpio = texto.strip()
    if not limpio:
        raise ValueError("Fecha vacía. Usá el formato AAAA-MM-DD (ejemplo: 2026-03-15).")
    try:
        return date.fromisoformat(limpio)
    except ValueError:
        pass
    patrones = [
        (r"^(\d{1,2})/(\d{1,2})/(\d{4})$", (1, 2, 3)),
        (r"^(\d{1,2})-(\d{1,2})-(\d{4})$", (1, 2, 3)),
    ]
    for patron, orden in patrones:
        m = re.match(patron, limpio)
        if m:
            d_, m_, a_ = (int(m.group(i)) for i in orden)
            try:
                return date(a_, m_, d_)
            except ValueError:
                break
    raise ValueError(
        f"Fecha inválida: '{texto}'. Formatos aceptados: AAAA-MM-DD, DD/MM/AAAA (ejemplo: 2026-03-15)."
    )


def noches_entdesde_hasta(ini: date, fin: date):
    d = ini
    while d < fin:
        yield d
        d += timedelta(days=1)


@dataclass
class Atomo:
    miembros: tuple[Pax, ...]
    tamano: int
    sexos_no_x: frozenset
    tiene_menor: bool
    tiene_adulto: bool
    requiere_single: bool = False
    exento_sexo: bool = False

    @property
    def es_mixto(self) -> bool:
        return len(self.sexos_no_x) > 1

    @property
    def menores(self) -> int:
        return sum(1 for m in self.miembros if m.es_menor())

    @property
    def descripcion(self) -> str:
        return ", ".join(m.nombre for m in self.miembros)

    @property
    def clave_orden(self) -> tuple:
        return (
            -int(self.requiere_single),
            -self.tamano,
            -self.menores,
            tuple(sorted(m.nombre.lower() for m in self.miembros)),
        )


SEXO_MIXTO = "MIXTO"


class EstadoHabitacion:
    __slots__ = ("capacidad", "carga", "sexo", "contiene_single", "menores", "atomos")

    def __init__(self, capacidad: int):
        self.capacidad = capacidad
        self.carga = 0
        self.sexo = None
        self.contiene_single = False
        self.menores = 0
        self.atomos: list[Atomo] = []

    @property
    def vacia(self) -> bool:
        return not self.atomos

    @property
    def espacio_libre(self) -> int:
        return self.capacidad - self.carga

    def puede_recibir(self, atomo: Atomo, config: ConfigGrupo) -> bool:
        if self.carga + atomo.tamano > self.capacidad:
            return False
        if self.contiene_single:
            return False
        if atomo.requiere_single:
            return self.vacia
        if self.sexo == SEXO_MIXTO:
            return False
        if atomo.es_mixto:
            return self.vacia
        if config.separar_sexos and atomo.sexos_no_x and self.sexo is not None:
            if next(iter(atomo.sexos_no_x)) != self.sexo:
                return False
        if config.max_ninos_por_habitacion is not None:
            if self.menores + atomo.menores > config.max_ninos_por_habitacion:
                return False
        return True

    def recibir(self, atomo: Atomo) -> None:
        self.carga += atomo.tamano
        self.menores += atomo.menores
        if atomo.es_mixto:
            self.sexo = SEXO_MIXTO
        elif atomo.sexos_no_x and self.sexo is None:
            self.sexo = next(iter(atomo.sexos_no_x))
        self.contiene_single = self.contiene_single or atomo.requiere_single
        self.atomos.append(atomo)

    def quitar(self, atomo: Atomo) -> None:
        for indice, presente in enumerate(self.atomos):
            if presente is atomo:
                del self.atomos[indice]
                break
        else:
            raise ValueError(f"'{atomo.descripcion}' no está en esta habitación.")
        self.carga -= atomo.tamano
        self.menores -= atomo.menores
        if self.vacia:
            self.sexo = None
            self.contiene_single = False
        elif self.sexo == SEXO_MIXTO:
            restantes = {s for a in self.atomos for s in a.sexos_no_x}
            self.sexo = SEXO_MIXTO if len(restantes) > 1 else (next(iter(restantes)) if restantes else None)
        self.contiene_single = any(a.requiere_single for a in self.atomos)

    def firma(self) -> tuple:
        return (self.capacidad, self.carga, self.sexo, self.contiene_single, self.menores)
