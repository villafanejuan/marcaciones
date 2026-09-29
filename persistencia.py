from __future__ import annotations

import json
import os
from datetime import date

from modelos import (
    Bloqueo,
    ConfigGrupo,
    ErrorDistribucion,
    Grupo,
    Habitacion,
    Hotel,
    LoteCamas,
    ModoHospedaje,
    Pax,
    Sexo,
    TipoPax,
)

VERSION_FORMATO = 2


def _fecha_a_str(fecha: date | None) -> str | None:
    return fecha.isoformat() if fecha is not None else None


def _str_a_fecha(texto, contexto: str) -> date | None:
    if texto is None or texto == "":
        return None
    try:
        return date.fromisoformat(str(texto))
    except (ValueError, TypeError):
        raise ErrorDistribucion(
            f"{contexto}: fecha inválida '{texto}'. El archivo usa el formato AAAA-MM-DD."
        )


def _enum_de(valor, enum_cls, contexto: str):
    try:
        return enum_cls(str(valor))
    except ValueError:
        raise ErrorDistribucion(
            f"{contexto}: valor desconocido '{valor}' para {enum_cls.__name__}. "
            f"Valores válidos: {', '.join(e.value for e in enum_cls)}."
        )


def pax_a_dict(p: Pax) -> dict:
    return {
        "nombre": p.nombre,
        "tipo": p.tipo.value,
        "sexo": p.sexo.value,
        "edad": p.edad,
        "pareja_de": p.pareja_de,
        "familia": p.familia,
        "requiere_single": p.requiere_single,
        "notas": p.notas,
    }


def pax_de_dict(datos: dict, contexto: str) -> Pax:
    try:
        nombre = datos["nombre"]
    except (KeyError, TypeError):
        raise ErrorDistribucion(f"{contexto}: pasajero sin campo 'nombre'.")
    return Pax(
        nombre=str(nombre),
        tipo=_enum_de(datos.get("tipo", TipoPax.ADULTO.value), TipoPax, f"{contexto} ({nombre})"),
        sexo=_enum_de(datos.get("sexo", Sexo.X.value), Sexo, f"{contexto} ({nombre})"),
        edad=datos.get("edad"),
        pareja_de=datos.get("pareja_de"),
        familia=datos.get("familia"),
        requiere_single=bool(datos.get("requiere_single", False)),
        notas=str(datos.get("notas", "")),
    )


def config_a_dict(c: ConfigGrupo) -> dict:
    return {
        "separar_sexos": c.separar_sexos,
        "ninos_con_adulto_obligatorio": c.ninos_con_adulto_obligatorio,
        "guias_choferes_single": c.guias_choferes_single,
        "estudiantes_por_habitacion": c.estudiantes_por_habitacion,
        "estudiantes_estricto": c.estudiantes_estricto,
        "max_ninos_por_habitacion": c.max_ninos_por_habitacion,
        "familias_exentas_sexo": c.familias_exentas_sexo,
        "bebes_no_cuentan_cupo": c.bebes_no_cuentan_cupo,
        "modo_hospedaje": c.modo_hospedaje,
        "max_plazas_estudiantes": c.max_plazas_estudiantes,
        "adultos_por_rol": c.adultos_por_rol,
    }


def config_de_dict(datos: dict, contexto: str) -> ConfigGrupo:
    datos = datos or {}
    return ConfigGrupo(
        separar_sexos=bool(datos.get("separar_sexos", True)),
        ninos_con_adulto_obligatorio=bool(datos.get("ninos_con_adulto_obligatorio", True)),
        guias_choferes_single=bool(datos.get("guias_choferes_single", False)),
        estudiantes_por_habitacion=datos.get("estudiantes_por_habitacion"),
        estudiantes_estricto=bool(datos.get("estudiantes_estricto", False)),
        max_ninos_por_habitacion=datos.get("max_ninos_por_habitacion"),
        familias_exentas_sexo=bool(datos.get("familias_exentas_sexo", True)),
        bebes_no_cuentan_cupo=bool(datos.get("bebes_no_cuentan_cupo", False)),
        modo_hospedaje=_modo_desde(datos.get("modo_hospedaje"), contexto),
        max_plazas_estudiantes=datos.get("max_plazas_estudiantes", 4),
        adultos_por_rol=bool(datos.get("adultos_por_rol", True)),
    )


def _modo_desde(valor, contexto: str) -> str:
    """Lee el modo de hospedaje tolerando archivos viejos (que no lo tenian)."""
    validos = {m.value for m in ModoHospedaje}
    if valor is None or valor == "":
        return ModoHospedaje.ESTUDIANTIL.value
    texto = str(valor).strip().upper()
    if texto in validos:
        return texto
    raise ErrorDistribucion(
        f"{contexto}: modo de hospedaje desconocido '{valor}'. "
        f"Valores válidos: {', '.join(sorted(validos))}."
    )


def bloqueo_a_dict(b: Bloqueo) -> dict:
    return {"desde": _fecha_a_str(b.desde), "hasta": _fecha_a_str(b.hasta), "titulo": b.titulo}


def bloqueo_de_dict(datos: dict, contexto: str) -> Bloqueo:
    desde = _str_a_fecha(datos.get("desde"), contexto)
    hasta = _str_a_fecha(datos.get("hasta"), contexto)
    if desde is None or hasta is None:
        raise ErrorDistribucion(f"{contexto}: bloqueo sin fechas completas.")
    return Bloqueo(desde=desde, hasta=hasta, titulo=str(datos.get("titulo", "")))


def habitacion_a_dict(h: Habitacion) -> dict:
    return {
        "numero": h.numero,
        "capacidad": h.capacidad,
        "piso": h.piso,
        "disponible_desde": _fecha_a_str(h.disponible_desde),
        "disponible_hasta": _fecha_a_str(h.disponible_hasta),
        "bloqueos": [bloqueo_a_dict(b) for b in h.bloqueos],
    }


def habitacion_de_dict(datos: dict, indice: int) -> Habitacion:
    contexto = f"Habitación #{indice + 1}"
    numero = datos.get("numero")
    if not numero:
        raise ErrorDistribucion(f"{contexto}: falta el número de habitación.")
    capacidad = datos.get("capacidad")
    if not isinstance(capacidad, int) or capacidad < 1:
        raise ErrorDistribucion(
            f"{contexto} ('{numero}'): capacidad inválida '{capacidad}'. Debe ser un entero >= 1."
        )
    hab = Habitacion(
        numero=str(numero),
        capacidad=capacidad,
        piso=str(datos.get("piso", "")),
        disponible_desde=_str_a_fecha(datos.get("disponible_desde"), contexto),
        disponible_hasta=_str_a_fecha(datos.get("disponible_hasta"), contexto),
    )
    for i, b in enumerate(datos.get("bloqueos", [])):
        hab.bloqueos.append(bloqueo_de_dict(b, f"{contexto} ('{numero}') bloqueo #{i + 1}"))
    return hab


def lote_camas_a_dict(l: LoteCamas) -> dict:
    return {
        "titulo": l.titulo,
        "cantidad": l.cantidad,
        "disponible_desde": _fecha_a_str(l.disponible_desde),
        "disponible_hasta": _fecha_a_str(l.disponible_hasta),
    }


def lote_camas_de_dict(datos: dict, indice: int) -> LoteCamas:
    contexto = f"Lote de camas #{indice + 1}"
    cantidad = datos.get("cantidad")
    if not isinstance(cantidad, int) or cantidad < 1:
        raise ErrorDistribucion(
            f"{contexto}: cantidad inválida '{cantidad}'. Debe ser un entero >= 1."
        )
    return LoteCamas(
        titulo=str(datos.get("titulo", f"Lote {indice + 1}")),
        cantidad=cantidad,
        disponible_desde=_str_a_fecha(datos.get("disponible_desde"), contexto),
        disponible_hasta=_str_a_fecha(datos.get("disponible_hasta"), contexto),
    )


def hotel_a_dict(hotel: Hotel) -> dict:
    prov = None
    if hotel.provisional:
        prov = dict(hotel.provisional)
        prov["desde"] = _fecha_a_str(prov.get("desde"))
        prov["hasta"] = _fecha_a_str(prov.get("hasta"))
    return {
        "nombre": hotel.nombre,
        "habitaciones": [habitacion_a_dict(h) for h in hotel.habitaciones],
        "lotes_camas": [lote_camas_a_dict(l) for l in hotel.lotes_camas],
        "provisional": prov,
    }


def hotel_de_dict(datos: dict) -> Hotel:
    prov = datos.get("provisional")
    if isinstance(prov, dict):
        prov = dict(prov)
        prov["desde"] = _str_a_fecha(prov.get("desde"), "Cotización provisional")
        prov["hasta"] = _str_a_fecha(prov.get("hasta"), "Cotización provisional")
    else:
        prov = None
    return Hotel(
        nombre=str(datos.get("nombre", "Mi Hotel")),
        habitaciones=[habitacion_de_dict(d, i) for i, d in enumerate(datos.get("habitaciones", []))],
        lotes_camas=[lote_camas_de_dict(d, i)
                     for i, d in enumerate(datos.get("lotes_camas", []))],
        provisional=prov,
    )


def grupo_a_dict(g: Grupo) -> dict:
    return {
        "nombre": g.nombre,
        "fecha_ingreso": _fecha_a_str(g.fecha_ingreso),
        "fecha_egreso": _fecha_a_str(g.fecha_egreso),
        "config": config_a_dict(g.config),
        "pax": [pax_a_dict(p) for p in g.pax],
        "empresa": g.empresa,
        "provisional": g.provisional,
        "confirmado": g.confirmado,
    }


def grupo_de_dict(datos: dict, indice: int) -> Grupo:
    contexto = f"Grupo #{indice + 1}"
    nombre = datos.get("nombre")
    if not nombre:
        raise ErrorDistribucion(f"{contexto}: falta el nombre del grupo.")
    ingreso = _str_a_fecha(datos.get("fecha_ingreso"), f"Grupo '{nombre}'")
    egreso = _str_a_fecha(datos.get("fecha_egreso"), f"Grupo '{nombre}'")
    return Grupo(
        nombre=str(nombre),
        fecha_ingreso=ingreso,
        fecha_egreso=egreso,
        config=config_de_dict(datos.get("config"), f"Grupo '{nombre}'"),
        pax=[pax_de_dict(p, f"Grupo '{nombre}'") for p in datos.get("pax", [])],
        empresa=str(datos.get("empresa", "")),
        provisional=bool(datos.get("provisional", False)),
        confirmado=bool(datos.get("confirmado", False)),
    )


def _hotel_completo_a_dict(hotel: Hotel, grupos: list[Grupo]) -> dict:
    datos = hotel_a_dict(hotel)
    datos["grupos"] = [grupo_a_dict(g) for g in grupos]
    return datos


def _hotel_completo_de_dict(datos: dict, indice: int) -> tuple[Hotel, list[Grupo]]:
    hotel = hotel_de_dict(datos)
    grupos = [grupo_de_dict(g, i) for i, g in enumerate(datos.get("grupos", []))]
    return hotel, grupos


def guardar_multi(ruta: str, hoteles: list[Hotel],
                  grupos_por_hotel: list[list[Grupo]],
                  estado: dict | None = None) -> None:
    datos = {
        "version": VERSION_FORMATO,
        "hoteles": [_hotel_completo_a_dict(h, g)
                    for h, g in zip(hoteles, grupos_por_hotel)],
        "estado": estado or {},
    }
    try:
        with open(ruta, "w", encoding="utf-8") as archivo:
            json.dump(datos, archivo, ensure_ascii=False, indent=2)
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo guardar '{ruta}': {exc}")


def guardar_datos(ruta: str, hotel: Hotel, grupos: list[Grupo]) -> None:
    guardar_multi(ruta, [hotel], [grupos])


def _leer_datos(ruta: str) -> dict:
    if not os.path.exists(ruta):
        raise ErrorDistribucion(f"El archivo '{ruta}' no existe.")
    try:
        with open(ruta, "r", encoding="utf-8") as archivo:
            datos = json.load(archivo)
    except OSError as exc:
        raise ErrorDistribucion(f"No se pudo leer '{ruta}': {exc}")
    except json.JSONDecodeError as exc:
        raise ErrorDistribucion(f"El archivo '{ruta}' no es JSON válido: {exc}")
    if not isinstance(datos, dict):
        raise ErrorDistribucion(f"El archivo '{ruta}' tiene un formato no reconocido.")
    version = datos.get("version")
    if version not in (1, VERSION_FORMATO):
        raise ErrorDistribucion(
            f"El archivo '{ruta}' declara versión {version}; este programa "
            f"maneja la versión {VERSION_FORMATO}."
        )
    return datos


def cargar_multi(ruta: str) -> tuple[list[Hotel], list[list[Grupo]]]:
    hoteles, grupos_por_hotel, _estado = cargar_area(ruta)
    return hoteles, grupos_por_hotel


def cargar_area(ruta: str) -> tuple[list[Hotel], list[list[Grupo]], dict]:
    """Carga hoteles + grupos por hotel + estado de interfaz guardado."""
    datos = _leer_datos(ruta)
    estado = datos.get("estado", {})
    if not isinstance(estado, dict):
        estado = {}
    if "hoteles" in datos and datos["hoteles"]:
        hoteles = list(datos["hoteles"])
        devueltos = [_hotel_completo_de_dict(h, i) for i, h in enumerate(hoteles)]
        return [h for h, _ in devueltos], [g for _, g in devueltos], estado
    legacy = [
        ("hotel", "grupos"),
    ]
    for clave_h, clave_g in legacy:
        if clave_h in datos:
            hotel = hotel_de_dict(datos.get(clave_h, {}))
            grupos = [grupo_de_dict(g, i)
                      for i, g in enumerate(datos.get(clave_g, []))]
            return [hotel], [grupos], estado
    raise ErrorDistribucion(
        f"El archivo '{ruta}' no contiene datos de hoteles reconocibles.")


def cargar_datos(ruta: str) -> tuple[Hotel, list[Grupo]]:
    hoteles, grupos_por_hotel = cargar_multi(ruta)
    return hoteles[0], grupos_por_hotel[0]
