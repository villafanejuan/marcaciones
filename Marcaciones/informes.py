from __future__ import annotations


def formato_fecha(d) -> str:
    return d.strftime("%d/%m/%Y")


def texto_deficit(deficit: dict[int, int]) -> str:
    if not deficit:
        return ""
    partes = [f"{cantidad} x habitacion de {cap} lugares" for cap, cantidad in sorted(deficit.items(), reverse=True)]
    return "Faltarian: " + ", ".join(partes)


def barra(valor: int, maximo: int, ancho: int = 30) -> str:
    if maximo <= 0:
        return ""
    llenos = round(valor * ancho / maximo)
    return "#" * llenos + "." * (ancho - llenos)


def renderizar_reporte(reporte) -> str:
    lineas: list[str] = []
    hotel = reporte.hotel
    total_hab = len(hotel.habitaciones)
    total_camas = hotel.total_camas()

    lineas.append("=" * 74)
    lineas.append(f"HOTEL: {hotel.nombre}")
    lineas.append(
        f"Habitaciones: {total_hab} | Camas totales: {total_camas} "
        f"| Grupos procesados: {len(reporte.distribuciones)}"
    )
    lineas.append("=" * 74)

    for aviso in reporte.avisos:
        lineas.append(f"[AVISO GENERAL] {aviso}")

    for dist in reporte.distribuciones:
        grupo = dist.grupo
        camas_totales = sum(a.habitacion.capacidad for a in dist.asignaciones)
        camas_ocupadas = sum(a.carga for a in dist.asignaciones)
        estado = "COMPLETO" if dist.completo else "INCOMPLETO"
        if grupo.confirmado:
            estado += " · CONFIRMADO"
        elif grupo.provisional:
            estado += " · ESTIMADO"
        optimo = " | solucion optima" if dist.optimo else ""
        lineas.append("")
        lineas.append("-" * 74)
        lineas.append(
            f"GRUPO '{grupo.nombre}' [{estado}] {formato_fecha(grupo.fecha_ingreso)} -> "
            f"{formato_fecha(grupo.fecha_egreso)} ({grupo.noches()} noches)"
        )
        lineas.append(
            f"Pax: {dist.pax_alojados}/{grupo.total_pax()} alojados | "
            f"Habitaciones usadas: {dist.habitaciones_usadas} | "
            f"Camas: {camas_ocupadas}/{camas_totales}"
            f"{optimo} | Estrategia: {dist.estrategia_usada or '-'}"
        )
        lineas.append("-" * 74)

        for asig in dist.asignaciones:
            hab = asig.habitacion
            etiqueta_piso = f" (piso {hab.piso})" if hab.piso else ""
            lineas.append(
                f"  Hab {hab.numero}{etiqueta_piso} [{hab.capacidad} lugares] "
                f"ocupacion {asig.carga}/{hab.capacidad}:"
            )
            for atomo in asig.atomos:
                prefijo = "   * pareja/familia: " if len(atomo.miembros) > 1 else "     - "
                if atomo.requiere_single and len(atomo.miembros) == 1:
                    prefijo = "     * single obligatorio: "
                for p in atomo.miembros:
                    extra = f", {p.edad} anos" if p.edad is not None else ""
                    notas = f" ({p.notas})" if p.notas else ""
                    lineas.append(f"{prefijo}{p.nombre} [{p.tipo.value} {p.sexo.value}{extra}]{notas}")

        for aviso in dist.avisos:
            lineas.append(f"  AVISO: {aviso}")

        if dist.faltantes:
            lineas.append("  SIN HABITACION:")
            for p in dist.faltantes:
                lineas.append(f"    ! {p.descripcion()}")
            if dist.deficit:
                lineas.append(f"  {texto_deficit(dist.deficit)}")
            lineas.append(
                "  Sugerencia: ampliar inventario en esas fechas, mover fechas del grupo o "
                "relajar reglas (separacion de sexos, cupo fijo, etc.)."
            )

    if reporte.pax_por_noche:
        maximo_pax = max(reporte.pax_por_noche.values())
        maximo_hab = max(reporte.habitaciones_por_noche.values())
        lineas.append("")
        lineas.append("OCUPACION POR NOCHE:")
        fechas = sorted(reporte.pax_por_noche)
        mostrar_todas = len(fechas) <= 31
        if mostrar_todas:
            lista = fechas
        else:
            picos = sorted(reporte.pax_por_noche, key=lambda f: -reporte.pax_por_noche[f])[:10]
            lista = sorted(picos)
            lineas.append("  (rango largo: se muestran las 10 noches mas ocupadas)")
        for fecha in lista:
            pax = reporte.pax_por_noche[fecha]
            hab = reporte.habitaciones_por_noche[fecha]
            lineas.append(
                f"  {formato_fecha(fecha)} | {barra(pax, maximo_pax)} {pax:>3} pax | {hab:>3} hab"
            )
        lineas.append(f"  Pico: {maximo_pax} pax / {maximo_hab} habitaciones en una misma noche.")

    lineas.append("")
    lineas.append("=" * 74)
    return "\n".join(lineas)
