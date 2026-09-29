from __future__ import annotations

from modelos import (
    Atomo,
    ConfigGrupo,
    ErrorDistribucion,
    Grupo,
    Pax,
    TIPOS_TRIPULACION,
)


def _tamano(pax_list: list[Pax], config: ConfigGrupo) -> int:
    total = 0
    for p in pax_list:
        if p.tipo.value == "BEBE" and config.bebes_no_cuentan_cupo:
            continue
        total += 1
    return total


def atomizar(grupo: Grupo) -> tuple[list[Atomo], list[str]]:
    config = grupo.config
    avisos: list[str] = []

    nombres_vistos = {}
    for p in grupo.pax:
        clave = p.nombre.strip().lower()
        if clave in nombres_vistos:
            raise ErrorDistribucion(
                f"Grupo '{grupo.nombre}': nombre duplicado '{p.nombre}'. "
                "Cada pasajero necesita un nombre único dentro del grupo."
            )
        nombres_vistos[clave] = p

    for p in grupo.pax:
        if p.tipo.value == "NINO" and p.edad is not None and p.edad >= 18:
            avisos.append(f"'{p.nombre}' está marcado como NINO con {p.edad} años. Revisá si debería ser ADULTO.")
        if p.tipo.value == "BEBE" and p.edad is not None and p.edad > 3:
            avisos.append(f"'{p.nombre}' está marcado como BEBE con {p.edad} años. Revisá si debería ser NINO.")

    for p in grupo.pax:
        if p.pareja_de:
            clave_pareja = p.pareja_de.strip().lower()
            if clave_pareja == p.nombre.strip().lower():
                raise ErrorDistribucion(f"Grupo '{grupo.nombre}': '{p.nombre}' no puede ser pareja de sí mismo.")
            pareja = nombres_vistos.get(clave_pareja)
            if pareja is None:
                raise ErrorDistribucion(
                    f"Grupo '{grupo.nombre}': '{p.nombre}' declara pareja '{p.pareja_de}' pero ese nombre "
                    "no existe en la lista del grupo."
                )
            menor_por_edad = (p.edad is not None and p.edad < 18) or (
                pareja.edad is not None and pareja.edad < 18
            )
            if pareja.es_menor() or p.es_menor() or menor_por_edad:
                raise ErrorDistribucion(
                    f"Grupo '{grupo.nombre}': la pareja '{p.nombre} - {pareja.nombre}' incluye un menor. "
                    "Usá 'familia' para vincular menores con adultos."
                )

    familias: dict[str, list[Pax]] = {}
    individuales: list[Pax] = []
    usados_en_familia: set[str] = set()

    for p in grupo.pax:
        if p.familia:
            clave = p.familia.strip().upper()
            familias.setdefault(clave, []).append(p)
            usados_en_familia.add(p.nombre.strip().lower())

    atomos: list[Atomo] = []

    for id_familia, miembros in sorted(familias.items()):
        tiene_adulto = any(m.es_adulto() for m in miembros)
        tiene_menor = any(m.es_menor() for m in miembros)
        if tiene_menor and not tiene_adulto and config.ninos_con_adulto_obligatorio:
            nombres = ", ".join(m.nombre for m in miembros)
            raise ErrorDistribucion(
                f"Grupo '{grupo.nombre}': la familia '{id_familia}' tiene menores sin ningún adulto "
                f"({nombres}). Agregá un adulto con esa misma familia o desactivá la regla "
                "'niños con adulto obligatorio'."
            )
        if config.max_ninos_por_habitacion is not None:
            cant_menores = sum(1 for m in miembros if m.es_menor())
            if cant_menores > config.max_ninos_por_habitacion:
                raise ErrorDistribucion(
                    f"Grupo '{grupo.nombre}': la familia '{id_familia}' tiene {cant_menores} menores y el "
                    f"máximo por habitación es {config.max_ninos_por_habitacion}. Ajustá la familia o el máximo."
                )
        sexos = frozenset(m.sexo for m in miembros if m.sexo.value != "X")
        atomos.append(
            Atomo(
                miembros=tuple(miembros),
                tamano=_tamano(miembros, config),
                sexos_no_x=sexos,
                tiene_menor=tiene_menor,
                tiene_adulto=tiene_adulto,
                requiere_single=False,
                exento_sexo=config.familias_exentas_sexo,
            )
        )
        for m in miembros:
            if m.pareja_de:
                clave_pareja = m.pareja_de.strip().lower()
                if clave_pareja not in {x.nombre.strip().lower() for x in miembros}:
                    avisos.append(
                        f"Grupo '{grupo.nombre}': '{m.nombre}' declara pareja fuera de su familia; "
                        "el vínculo de pareja se ignora porque la familia ya los mantiene juntos."
                    )

    parejas_procesadas: set[str] = set()
    for p in grupo.pax:
        clave_p = p.nombre.strip().lower()
        if not p.pareja_de or clave_p in usados_en_familia or clave_p in parejas_procesadas:
            continue
        pareja = nombres_vistos[p.pareja_de.strip().lower()]
        clave_q = pareja.nombre.strip().lower()
        if clave_q in usados_en_familia or clave_q in parejas_procesadas:
            continue
        parejas_procesadas.add(clave_p)
        parejas_procesadas.add(clave_q)
        duo = [p, pareja]
        sexos = frozenset(m.sexo for m in duo if m.sexo.value != "X")
        atomos.append(
            Atomo(
                miembros=tuple(duo),
                tamano=_tamano(duo, config),
                sexos_no_x=sexos,
                tiene_menor=False,
                tiene_adulto=True,
                requiere_single=False,
                exento_sexo=True,
            )
        )

    for p in grupo.pax:
        clave_p = p.nombre.strip().lower()
        if clave_p in usados_en_familia or clave_p in parejas_procesadas:
            continue
        es_single = p.requiere_single or (config.guias_choferes_single and p.tipo in TIPOS_TRIPULACION)
        sexos = frozenset({p.sexo}) if p.sexo.value != "X" else frozenset()
        atomos.append(
            Atomo(
                miembros=(p,),
                tamano=_tamano([p], config),
                sexos_no_x=sexos,
                tiene_menor=p.es_menor(),
                tiene_adulto=p.es_adulto(),
                requiere_single=es_single,
                exento_sexo=False,
            )
        )

    if config.ninos_con_adulto_obligatorio:
        for a in atomos:
            if a.tiene_menor and not a.tiene_adulto:
                raise ErrorDistribucion(
                    f"Grupo '{grupo.nombre}': menores sin adulto responsable ({a.descripcion}). "
                    "Vinculá un adulto mediante 'familia' o desactivá la regla."
                )

    atomos.sort(key=lambda a: a.clave_orden)
    return atomos, avisos
