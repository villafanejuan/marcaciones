from __future__ import annotations

import math
import os
import re
import sys
import traceback
from datetime import date, timedelta

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

import tkinter as tk
from tkinter import filedialog, messagebox, ttk

import customtkinter as ctk

from exportador import exportar_csv, exportar_resumen_csv
from excel_exportador import (
    exportar_excel,
    exportar_mapa_grupos,
    exportar_rooming_excel,
    exportar_rooming_hoteles,
)
from informes import renderizar_reporte
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
    parsear_fecha,
)
from motor import (
    OpcionesMotor,
    clave_orden_habitacion,
    crear_grupo_generico,
    distribuir,
    generar_hotel_estimado,
    sugerir_mezcla,
)
from reglas import atomizar
from persistencia import (
    cargar_datos,
    cargar_multi,
    cargar_area,
    guardar_datos,
    guardar_multi,
    hotel_a_dict,
    hotel_de_dict,
)

COLOR_LIBRE = "#1e3a2f"
BORDE_LIBRE = "#2f5d47"
COLOR_BLOQUEADA = "#4a1d1d"
BORDE_BLOQUEO = "#8a3232"
FONDO_MAPA = "#161618"
TEXTO_CLARO = "#f5f5f5"
PALETA_GRUPOS = [
    "#c2410c",
    "#0e7490",
    "#a21caf",
    "#4d7c0f",
    "#b45309",
    "#be185d",
    "#4338ca",
    "#15803d",
    "#7c3aed",
    "#0369a1",
]
TAM_CARD_ANCHO = 104
TAM_CARD_ALTO = 66
GAP_X = 12
GAP_Y = 14


def _carpeta_base_datos() -> str:
    if getattr(sys, "frozen", False):
        return os.path.dirname(sys.executable)
    return os.path.dirname(os.path.abspath(__file__))


CARPETA_EXPORTACIONES = os.path.join(_carpeta_base_datos(), "exportaciones")
RUTA_ESPACIO_TRABAJO = os.path.join(CARPETA_EXPORTACIONES, "espacio_trabajo.json")


def formatear(d: date) -> str:
    return d.strftime("%d/%m/%Y")


def piso_de(numero: str) -> str:
    limpio = numero.strip()
    if limpio.isdigit():
        return str(int(limpio) // 100) if int(limpio) >= 100 else "0"
    return limpio[:1].upper()


def generar_numeros(existentes: set[str], cantidad: int) -> list[str]:
    usados = {n for n in existentes if n.strip()}
    numericos = [int(n) for n in usados if n.isdigit()]
    if numericos:
        mayor = max(numericos)
        proximo = (mayor // 100) * 100 + 101
        if proximo <= mayor:
            proximo += 100
    else:
        proximo = 101
    elegidos = []
    while len(elegidos) < cantidad:
        candidato = str(proximo)
        proximo += 1
        if candidato in usados:
            continue
        elegidos.append(candidato)
    return elegidos


def crear_pax_en_lote(cantidad: int, tipo: TipoPax, sexo: Sexo, prefijo: str = "Pax",
                      desde: int = 1) -> list[Pax]:
    return [
        Pax(nombre=f"{prefijo} {indice:02d}", tipo=tipo, sexo=sexo)
        for indice in range(desde, desde + cantidad)
    ]


def ruta_log() -> str:
    if getattr(sys, "frozen", False):
        base = os.path.dirname(sys.executable)
    else:
        base = sys.path[0]
    return os.path.join(base, "error_interfaz.log")


class DialogoHabitaciones(ctk.CTkToplevel):
    def __init__(self, master, numeros_existentes: set[str] | None = None,
                 capacidad_sugerida: int = 4, explicacion: str = ""):
        super().__init__(master)
        self.numeros_existentes = set(numeros_existentes or ())
        self.capacidad_sugerida = max(1, int(capacidad_sugerida))
        self.capacidad_sugerida_usada = False
        self.title("Agregar habitaciones")
        self.geometry("380x420")
        self.resizable(False, False)
        self.resultado: list[Habitacion] | None = None
        self.transient(master)
        self.grab_set()
        marco = ctk.CTkFrame(self)
        marco.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(marco, text="Número inicial (vacío = numeración automática)", anchor="w").pack(fill="x")
        self.entrada_numero = ctk.CTkEntry(marco, placeholder_text="ej: 501 o vacío")
        self.entrada_numero.pack(fill="x", pady=(2, 10))

        fila = ctk.CTkFrame(marco, fg_color="transparent")
        fila.pack(fill="x")
        col1 = ctk.CTkFrame(fila, fg_color="transparent")
        col1.pack(side="left", fill="x", expand=True, padx=(0, 6))
        col2 = ctk.CTkFrame(fila, fg_color="transparent")
        col2.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkLabel(col1, text="Cantidad", anchor="w").pack(fill="x")
        self.entrada_cantidad = ctk.CTkEntry(col1, placeholder_text="ej: 8")
        self.entrada_cantidad.pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(col2, text="Capacidad (vacío = valor sugerido)", anchor="w").pack(fill="x")
        self.entrada_capacidad = ctk.CTkEntry(col2, placeholder_text=f"ej: 4 o vacío (= {self.capacidad_sugerida})")
        self.entrada_capacidad.pack(fill="x", pady=(2, 10))
        if explicacion:
            ctk.CTkLabel(col2, text=explicacion, font=ctk.CTkFont(size=11),
                         text_color="#9ca3af", wraplength=170, justify="left").pack(fill="x")

        fila2 = ctk.CTkFrame(marco, fg_color="transparent")
        fila2.pack(fill="x")
        col3 = ctk.CTkFrame(fila2, fg_color="transparent")
        col3.pack(side="left", fill="x", expand=True, padx=(0, 6))
        col4 = ctk.CTkFrame(fila2, fg_color="transparent")
        col4.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkLabel(col3, text="Disponible desde (vacío = siempre)", anchor="w").pack(fill="x")
        self.entrada_desde = ctk.CTkEntry(col3, placeholder_text="AAAA-MM-DD")
        self.entrada_desde.pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(col4, text="Disponible hasta (vacío = siempre)", anchor="w").pack(fill="x")
        self.entrada_hasta = ctk.CTkEntry(col4, placeholder_text="AAAA-MM-DD")
        self.entrada_hasta.pack(fill="x", pady=(2, 10))

        self.var_ventana = ctk.BooleanVar(value=False)
        ctk.CTkSwitch(marco, text="Limitar ventana de disponibilidad",
                      variable=self.var_ventana, command=self._alternar_ventana).pack(anchor="w", pady=4)
        self.entrada_desde.configure(state="disabled")
        self.entrada_hasta.configure(state="disabled")

        botones = ctk.CTkFrame(marco, fg_color="transparent")
        botones.pack(side="bottom", fill="x", pady=(12, 0))
        ctk.CTkButton(botones, text="Cancelar", fg_color="gray30",
                      command=self.destroy).pack(side="left", expand=True, padx=4)
        ctk.CTkButton(botones, text="Agregar habitaciones",
                      command=self._confirmar).pack(side="left", expand=True, padx=4)

    def _alternar_ventana(self):
        estado = "normal" if self.var_ventana.get() else "disabled"
        self.entrada_desde.configure(state=estado)
        self.entrada_hasta.configure(state=estado)

    def _confirmar(self):
        try:
            texto_inicial = self.entrada_numero.get().strip()
            texto_cantidad = self.entrada_cantidad.get().strip()
            if not texto_cantidad.isdigit():
                raise ValueError("Ingresá cuántas habitaciones querés agregar (número entero).")
            cantidad = int(texto_cantidad)
            texto_capacidad = self.entrada_capacidad.get().strip()
            if texto_capacidad:
                if not texto_capacidad.isdigit():
                    raise ValueError("La capacidad debe ser un número entero "
                                     "(o dejá el campo vacío para el valor sugerido).")
                capacidad = int(texto_capacidad)
            else:
                capacidad = self.capacidad_sugerida
                self.capacidad_sugerida_usada = True
            if cantidad < 1:
                raise ValueError("La cantidad debe ser mayor a cero.")
            if capacidad < 1:
                raise ValueError("La capacidad debe ser mayor a cero.")
            disponible_desde = disponible_hasta = None
            if self.var_ventana.get():
                texto_desde = self.entrada_desde.get().strip() or str(date.today())
                texto_hasta = self.entrada_hasta.get().strip() or str(date.today().year + 1)
                disponible_desde = parsear_fecha(texto_desde)
                disponible_hasta = parsear_fecha(texto_hasta)
                if disponible_desde > disponible_hasta:
                    raise ValueError("La fecha 'desde' debe ser anterior a la 'hasta'.")
            if texto_inicial:
                if not texto_inicial.isdigit():
                    raise ValueError("El número inicial debe ser entero "
                                     "(o dejá el campo vacío para autonumerar).")
                inicial = int(texto_inicial)
                ancho = len(str(inicial))
                numeros = [str(inicial + i).zfill(ancho) for i in range(cantidad)]
            else:
                numeros = generar_numeros(self.numeros_existentes, cantidad)
            nuevas = [
                Habitacion(
                    numero=numero,
                    capacidad=capacidad,
                    disponible_desde=disponible_desde,
                    disponible_hasta=disponible_hasta,
                )
                for numero in numeros
            ]
        except ValueError as exc:
            messagebox.showerror("Datos inválidos", str(exc), parent=self)
            return
        self.resultado = nuevas
        self.destroy()


class DialogoBloqueo(ctk.CTkToplevel):
    def __init__(self, master, habitacion: Habitacion):
        super().__init__(master)
        self.habitacion = habitacion
        self.title(f"Bloquear habitación {habitacion.numero}")
        self.geometry("360x300")
        self.resizable(False, False)
        self.transient(master)
        self.grab_set()
        marco = ctk.CTkFrame(self)
        marco.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(
            marco,
            text=f"Bloqueo de mantenimiento / reserva manual para la habitación {habitacion.numero}.",
            wraplength=320,
            justify="left",
        ).pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(marco, text="Desde", anchor="w").pack(fill="x")
        self.entrada_desde = ctk.CTkEntry(marco, placeholder_text="AAAA-MM-DD")
        self.entrada_desde.pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(marco, text="Hasta (el día de egreso queda libre)", anchor="w").pack(fill="x")
        self.entrada_hasta = ctk.CTkEntry(marco, placeholder_text="AAAA-MM-DD")
        self.entrada_hasta.pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(marco, text="Motivo / título", anchor="w").pack(fill="x")
        self.entrada_titulo = ctk.CTkEntry(marco, placeholder_text="ej: Mantenimiento de aire")
        self.entrada_titulo.pack(fill="x", pady=(2, 10))
        botones = ctk.CTkFrame(marco, fg_color="transparent")
        botones.pack(side="bottom", fill="x", pady=(12, 0))
        ctk.CTkButton(botones, text="Cancelar", fg_color="gray30",
                      command=self.destroy).pack(side="left", expand=True, padx=4)
        ctk.CTkButton(botones, text="Bloquear", command=self._confirmar).pack(
            side="left", expand=True, padx=4
        )

    def _confirmar(self):
        try:
            desde = parsear_fecha(self.entrada_desde.get())
            hasta = parsear_fecha(self.entrada_hasta.get())
        except ValueError as exc:
            messagebox.showerror("Fecha inválida", str(exc), parent=self)
            return
        titulo = self.entrada_titulo.get().strip() or "Bloqueo manual"
        try:
            self.habitacion.agregar_bloqueo(desde, hasta, titulo=titulo)
        except ErrorDistribucion as exc:
            messagebox.showerror("No se pudo bloquear", str(exc), parent=self)
            return
        self.resultado = True
        self.destroy()


class DialogoPax(ctk.CTkToplevel):
    TIPOS = [t.value for t in TipoPax]

    def __init__(self, master, nombres_existentes: list[str], pax: Pax | None = None):
        super().__init__(master)
        self.title("Editar pasajero" if pax else "Nuevo pasajero")
        self.geometry("400x520")
        self.resizable(False, False)
        self.resultado: Pax | None = None
        self.transient(master)
        self.grab_set()
        marco = ctk.CTkFrame(self)
        marco.pack(fill="both", expand=True, padx=16, pady=16)

        ctk.CTkLabel(marco, text="Nombre y apellido", anchor="w").pack(fill="x")
        self.entrada_nombre = ctk.CTkEntry(marco)
        self.entrada_nombre.pack(fill="x", pady=(2, 10))

        ctk.CTkLabel(marco, text="Tipo de pasajero", anchor="w").pack(fill="x")
        self.menu_tipo = ctk.CTkOptionMenu(marco, values=self.TIPOS)
        self.menu_tipo.set(TipoPax.ADULTO.value if pax is None else pax.tipo.value)
        self.menu_tipo.pack(fill="x", pady=(2, 10))

        fila = ctk.CTkFrame(marco, fg_color="transparent")
        fila.pack(fill="x")
        izq = ctk.CTkFrame(fila, fg_color="transparent")
        izq.pack(side="left", fill="x", expand=True, padx=(0, 6))
        der = ctk.CTkFrame(fila, fg_color="transparent")
        der.pack(side="left", fill="x", expand=True, padx=(6, 0))
        ctk.CTkLabel(izq, text="Sexo", anchor="w").pack(fill="x")
        self.sexo_botones = ctk.CTkSegmentedButton(izq, values=["M", "F", "X"])
        self.sexo_botones.set(pax.sexo.value if pax else "X")
        self.sexo_botones.pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(der, text="Edad (vacío = sin dato)", anchor="w").pack(fill="x")
        self.entrada_edad = ctk.CTkEntry(der)
        if pax and pax.edad is not None:
            self.entrada_edad.insert(0, str(pax.edad))
        self.entrada_edad.pack(fill="x", pady=(2, 10))

        ctk.CTkLabel(marco, text="Familia (agrupa a los que comparten apellido de grupo)", anchor="w").pack(fill="x")
        self.entrada_familia = ctk.CTkEntry(marco, placeholder_text="ej: GARCIA (opcional)")
        if pax and pax.familia:
            self.entrada_familia.insert(0, pax.familia)
        self.entrada_familia.pack(fill="x", pady=(2, 10))

        ctk.CTkLabel(marco, text="Pareja de (comparten cama matrimonial)", anchor="w").pack(fill="x")
        valores_pareja = [""] + sorted(nombres_existentes)
        self.menu_pareja = ctk.CTkOptionMenu(marco, values=valores_pareja or [""])
        self.menu_pareja.set(pax.pareja_de if pax and pax.pareja_de else "")
        self.menu_pareja.pack(fill="x", pady=(2, 10))

        self.var_single = ctk.BooleanVar(value=bool(pax.requiere_single) if pax else False)
        ctk.CTkSwitch(marco, text="Necesita habitación individual (single obligatoria)",
                      variable=self.var_single).pack(anchor="w", pady=6)

        botones = ctk.CTkFrame(marco, fg_color="transparent")
        botones.pack(side="bottom", fill="x", pady=(12, 0))
        ctk.CTkButton(botones, text="Cancelar", fg_color="gray30",
                      command=self.destroy).pack(side="left", expand=True, padx=4)
        ctk.CTkButton(botones, text="Guardar pasajero", command=self._confirmar).pack(
            side="left", expand=True, padx=4
        )
        if pax:
            self.entrada_nombre.insert(0, pax.nombre)

    def _confirmar(self):
        nombre = self.entrada_nombre.get().strip()
        if not nombre:
            messagebox.showerror("Falta el nombre", "El pasajero necesita un nombre.", parent=self)
            return
        edad_cruda = self.entrada_edad.get().strip()
        edad = None
        if edad_cruda:
            try:
                edad = int(edad_cruda)
            except ValueError:
                messagebox.showerror("Edad inválida", "La edad debe ser un número entero.", parent=self)
                return
        pareja = self.menu_pareja.get().strip()
        pax = Pax(
            nombre=nombre,
            tipo=tipo_desde_valor(self.menu_tipo.get()),
            sexo=Sexo(self.sexo_botones.get()),
            edad=edad,
            familia=self.entrada_familia.get().strip() or None,
            pareja_de=pareja or None,
            requiere_single=self.var_single.get(),
        )
        self.resultado = pax
        self.destroy()


def tipo_desde_valor(valor: str) -> TipoPax:
    return TipoPax(valor)


class DialogoGrupo(ctk.CTkToplevel):
    COLUMNAS = ("nombre", "tipo", "sexo", "edad", "familia", "pareja", "single")
    TIPOS = [t.value for t in TipoPax]

    def __init__(self, master, grupo: Grupo | None = None):
        super().__init__(master)
        self.title("Editar grupo" if grupo else "Nuevo grupo")
        self.geometry("880x720")
        self.minsize(820, 640)
        self.resultado: Grupo | None = None
        self.pax_editable: list[Pax] = list(grupo.pax) if grupo else []
        self.transient(master)

        base_cfg = grupo.config if grupo else ConfigGrupo()
        self.vars_reglas = {
            "separar_sexos": ctk.BooleanVar(value=base_cfg.separar_sexos),
            "guias_choferes_single": ctk.BooleanVar(value=base_cfg.guias_choferes_single),
            "ninos_con_adulto_obligatorio": ctk.BooleanVar(value=base_cfg.ninos_con_adulto_obligatorio),
            "familias_exentas_sexo": ctk.BooleanVar(value=base_cfg.familias_exentas_sexo),
            "bebes_no_cuentan_cupo": ctk.BooleanVar(value=base_cfg.bebes_no_cuentan_cupo),
            "adultos_por_rol": ctk.BooleanVar(value=base_cfg.adultos_por_rol),
        }
        MODO_ESTUDIANTIL = ModoHospedaje.ESTUDIANTIL.value
        MODO_ADULTOS = ModoHospedaje.ADULTOS.value
        self.modo_hospedaje_inicial = (
            MODO_ADULTOS if base_cfg.modo_hospedaje == MODO_ADULTOS else MODO_ESTUDIANTIL
        )

        marco = ctk.CTkScrollableFrame(self, orientation="vertical")
        marco.pack(fill="both", expand=True, padx=14, pady=14)

        fila_sup = ctk.CTkFrame(marco, fg_color="transparent")
        fila_sup.pack(fill="x")
        col_nombre = ctk.CTkFrame(fila_sup, fg_color="transparent")
        col_nombre.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkLabel(col_nombre, text="Nombre del grupo", anchor="w").pack(fill="x")
        self.entrada_nombre = ctk.CTkEntry(col_nombre, placeholder_text="ej: Bachilleres 3° A")
        self.entrada_nombre.pack(fill="x", pady=(2, 10))
        col_empresa = ctk.CTkFrame(fila_sup, fg_color="transparent")
        col_empresa.pack(side="left", fill="x", expand=True, padx=(0, 10))
        ctk.CTkLabel(col_empresa, text="Empresa / escuela", anchor="w").pack(fill="x")
        self.entrada_empresa = ctk.CTkEntry(col_empresa, placeholder_text="ej: Colegio San Martín")
        self.entrada_empresa.pack(fill="x", pady=(2, 10))
        hoy = date.today()
        col_ini = ctk.CTkFrame(fila_sup, fg_color="transparent")
        col_ini.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(col_ini, text="Ingreso (vacío = sin fecha)", anchor="w").pack(fill="x")
        self.entrada_ingreso = ctk.CTkEntry(col_ini, width=130, placeholder_text="AAAA-MM-DD")
        self.entrada_ingreso.insert(0, str(hoy + timedelta(days=7)))
        self.entrada_ingreso.pack(pady=(2, 10))
        col_fin = ctk.CTkFrame(fila_sup, fg_color="transparent")
        col_fin.pack(side="left")
        ctk.CTkLabel(col_fin, text="Egreso (vacío = sin fecha)", anchor="w").pack(fill="x")
        self.entrada_egreso = ctk.CTkEntry(col_fin, width=130, placeholder_text="AAAA-MM-DD")
        self.entrada_egreso.insert(0, str(hoy + timedelta(days=11)))
        self.entrada_egreso.pack(pady=(2, 10))

        fila_modo = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        fila_modo.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(fila_modo, text="Modo del grupo",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 4))
        self.modo = ctk.CTkSegmentedButton(fila_modo, values=["GRUPAL", "INDIVIDUAL (provisional)"])
        self.modo.set("GRUPAL")
        self.modo.pack(anchor="w", padx=10, pady=(0, 8))
        ctk.CTkLabel(
            fila_modo,
            text="GRUPAL: datos completos, reglas estrictas.\n"
                 "INDIVIDUAL (provisional): cantidades parciales, se marca como «estimado». "
                 "Confirmás el grupo cuando ya lo enviaste al hotel.",
            text_color="#9ca3af", justify="left",
        ).pack(anchor="w", padx=10, pady=(0, 8))

        fila_tipo = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        fila_tipo.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(fila_tipo, text="Tipo de hospedaje",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))
        self.modo_hospedaje = ctk.CTkSegmentedButton(
            fila_tipo,
            values=[MODO_ESTUDIANTIL, MODO_ADULTOS],
            command=self._al_cambiar_modo_hospedaje,
        )
        self.modo_hospedaje.set(self.modo_hospedaje_inicial)
        self.modo_hospedaje.pack(anchor="w", padx=10, pady=(2, 6))
        self.etiqueta_modo_help = ctk.CTkLabel(fila_tipo, text="", text_color="#9ca3af", justify="left")
        self.etiqueta_modo_help.pack(anchor="w", padx=10, pady=(0, 6))
        fila_cupo = ctk.CTkFrame(fila_tipo, fg_color="transparent")
        fila_cupo.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(fila_cupo, text="Máx. plazas por hab. de estudiantes").pack(side="left")
        self.entrada_max_plazas = ctk.CTkEntry(fila_cupo, width=60, placeholder_text="4")
        if grupo and grupo.config.max_plazas_estudiantes is not None:
            self.entrada_max_plazas.insert(0, str(grupo.config.max_plazas_estudiantes))
        self.entrada_max_plazas.pack(side="left", padx=(8, 0))
        self.switch_adultos_por_rol = ctk.CTkSwitch(
            fila_cupo, text="Agrupar acompañantes por rol (choferes aparte de guías)",
            variable=self.vars_reglas["adultos_por_rol"])
        self.switch_adultos_por_rol.pack(side="left", padx=(18, 0))
        self.controles_estudiantiles = [self.entrada_max_plazas, self.switch_adultos_por_rol]

        reglas = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        reglas.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(reglas, text="Reglas de distribución para este grupo",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))
        textos_reglas = {
            "separar_sexos": "Separar sexos en habitaciones compartidas",
            "guias_choferes_single": "Forzar 1 persona por hab. a guías, choferes y coordinadores",
            "ninos_con_adulto_obligatorio": "Los menores solo pueden ir con un adulto responsable",
            "familias_exentas_sexo": "Las familias pueden mezclar sexos entre sí",
            "bebes_no_cuentan_cupo": "Los bebés no ocupan una plaza (van con los padres)",
        }
        for clave, texto in textos_reglas.items():
            ctk.CTkSwitch(reglas, text=texto, variable=self.vars_reglas[clave]).pack(
                anchor="w", padx=10, pady=2
            )

        fila_opciones = ctk.CTkFrame(reglas, fg_color="transparent")
        fila_opciones.pack(fill="x", padx=10, pady=(8, 10))
        c1 = ctk.CTkFrame(fila_opciones, fg_color="transparent")
        c1.pack(side="left", padx=(0, 14))
        ctk.CTkLabel(c1, text="Cupo fijo estudiantes por hab. (vacío = libre)", anchor="w").pack(fill="x")
        self.entrada_cupo = ctk.CTkEntry(c1, width=70, placeholder_text="ej: 4")
        if grupo and grupo.config.estudiantes_por_habitacion is not None:
            self.entrada_cupo.insert(0, str(grupo.config.estudiantes_por_habitacion))
        self.entrada_cupo.pack(pady=(2, 0))
        c2 = ctk.CTkFrame(fila_opciones, fg_color="transparent")
        c2.pack(side="left", padx=(0, 14))
        ctk.CTkLabel(c2, text="Máximo niños por habitación (vacío = sin tope)", anchor="w").pack(fill="x")
        self.entrada_max_ninos = ctk.CTkEntry(c2, width=70, placeholder_text="ej: 2")
        if grupo and grupo.config.max_ninos_por_habitacion is not None:
            self.entrada_max_ninos.insert(0, str(grupo.config.max_ninos_por_habitacion))
        self.entrada_max_ninos.pack(pady=(2, 0))
        c3 = ctk.CTkFrame(fila_opciones, fg_color="transparent")
        c3.pack(side="left")
        self.var_estricto = ctk.BooleanVar(value=bool(base_cfg.estudiantes_estricto))
        ctk.CTkSwitch(c3, text="Modo estricto\n(todas las hab. de estudiantes completas)",
                      variable=self.var_estricto).pack(anchor="w", pady=(14, 0))

        zona_pax = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        zona_pax.pack(fill="both", expand=True)
        encabezado = ctk.CTkFrame(zona_pax, fg_color="transparent")
        encabezado.pack(fill="x", padx=10, pady=(8, 2))
        self.etiqueta_conteo_pax = ctk.CTkLabel(
            encabezado, text=f"Pasajeros ({len(self.pax_editable)})",
            font=ctk.CTkFont(size=13, weight="bold"))
        self.etiqueta_conteo_pax.pack(side="left")
        botones_pax = ctk.CTkFrame(encabezado, fg_color="transparent")
        botones_pax.pack(side="right")
        ctk.CTkButton(botones_pax, text="+ Agregar", width=100, height=28,
                      command=self._agregar_pax).pack(side="left", padx=3)
        ctk.CTkButton(botones_pax, text="Editar", width=80, height=28, fg_color="gray35",
                      command=self._editar_pax).pack(side="left", padx=3)
        ctk.CTkButton(botones_pax, text="Quitar", width=80, height=28, fg_color="gray35",
                      command=self._quitar_pax).pack(side="left", padx=3)
        ctk.CTkButton(botones_pax, text="⚭ Vincular pareja", width=140, height=28,
                      fg_color="gray20", command=self._vincular_pareja).pack(side="left", padx=3)
        ctk.CTkButton(botones_pax, text="Desvincular", width=110, height=28,
                      fg_color="gray20", command=self._desvincular).pack(side="left", padx=3)
        ctk.CTkButton(botones_pax, text="Single SÍ/NO", width=120, height=28,
                      fg_color="gray20", command=self._alternar_single).pack(side="left", padx=3)

        carga_rapida = ctk.CTkFrame(zona_pax, fg_color="transparent")
        carga_rapida.pack(fill="x", padx=10, pady=(0, 4))
        ctk.CTkLabel(carga_rapida, text="Carga rápida:", text_color="#9ca3af").pack(
            side="left", padx=(0, 8))
        self.entrada_lote_cantidad = ctk.CTkEntry(carga_rapida, width=70,
                                                  placeholder_text="cant.")
        self.entrada_lote_cantidad.pack(side="left", padx=(0, 6))
        self.menu_lote_tipo = ctk.CTkOptionMenu(carga_rapida, values=self.TIPOS, width=150)
        self.menu_lote_tipo.set(TipoPax.ADULTO.value)
        self.menu_lote_tipo.pack(side="left", padx=(0, 6))
        self.sexo_lote = ctk.CTkSegmentedButton(carga_rapida, values=["M", "F", "X"],
                                                width=110)
        self.sexo_lote.set("X")
        self.sexo_lote.pack(side="left", padx=(0, 6))
        boton_lote = ctk.CTkButton(carga_rapida, text="＋ Agregar en lote", width=160,
                                   height=28, command=self._agregar_lote)
        boton_lote.pack(side="left")
        self.entrada_lote_cantidad.bind("<Return>", lambda _e: self._agregar_lote())

        contenedor_tabla = ctk.CTkFrame(zona_pax, fg_color="transparent")
        contenedor_tabla.pack(fill="both", expand=True, padx=10, pady=(4, 10))
        estilo = ttk.Style()
        estilo.theme_use("clam")
        estilo.configure("Tabla.Treeview", background="#1e1e21", fieldbackground="#1e1e21",
                         foreground=TEXTO_CLARO, rowheight=26, borderwidth=0)
        estilo.configure("Tabla.Treeview.Heading", background="#2b2b30",
                         foreground=TEXTO_CLARO, borderwidth=0)
        estilo.map("Tabla.Treeview", background=[("selected", "#2563eb")])
        self.tabla = ttk.Treeview(contenedor_tabla, columns=self.COLUMNAS, show="headings",
                                  style="Tabla.Treeview", selectmode="extended")
        anchos = {"nombre": 220, "tipo": 110, "sexo": 60, "edad": 60,
                  "familia": 110, "pareja": 150, "single": 90}
        titulos = {"nombre": "Nombre", "tipo": "Tipo", "sexo": "Sexo", "edad": "Edad",
                   "familia": "Familia", "pareja": "Pareja de", "single": "Individual"}
        for col in self.COLUMNAS:
            self.tabla.heading(col, text=titulos[col])
            self.tabla.column(col, width=anchos[col], anchor="w")
        scroll = ttk.Scrollbar(contenedor_tabla, orient="vertical", command=self.tabla.yview)
        self.tabla.configure(yscrollcommand=scroll.set)
        self.tabla.pack(side="left", fill="both", expand=True)
        scroll.pack(side="right", fill="y")
        self.tabla.bind("<Double-1>", lambda _e: self._editar_pax())

        botones_finales = ctk.CTkFrame(marco, fg_color="transparent")
        botones_finales.pack(fill="x", pady=(10, 0))
        ctk.CTkButton(botones_finales, text="Cancelar", fg_color="gray30",
                      height=36, command=self.destroy).pack(side="left", expand=True, padx=4)
        ctk.CTkButton(botones_finales, text="Guardar grupo", height=36,
                      command=self._guardar).pack(side="left", expand=True, padx=4)

        if grupo:
            self.entrada_nombre.insert(0, grupo.nombre)
            self.entrada_empresa.insert(0, grupo.empresa)
            self.entrada_ingreso.delete(0, "end")
            if grupo.fecha_ingreso:
                self.entrada_ingreso.insert(0, str(grupo.fecha_ingreso))
            self.entrada_egreso.delete(0, "end")
            if grupo.fecha_egreso:
                self.entrada_egreso.insert(0, str(grupo.fecha_egreso))
            self.modo.set("INDIVIDUAL (provisional)" if grupo.provisional else "GRUPAL")
        self._al_cambiar_modo_hospedaje()
        self._refrescar_tabla()

    def _set_modo_hospedaje(self, valor: str) -> None:
        """CTkSegmentedButton.set() no dispara `command`, asi que centralizamos el cambio
        para que la ayuda y los controles habilitados nunca queden desfasados."""
        self.modo_hospedaje.set(valor)
        self._al_cambiar_modo_hospedaje()

    def _al_cambiar_modo_hospedaje(self, _valor=None):
        estudiantil = self.modo_hospedaje.get() == ModoHospedaje.ESTUDIANTIL.value
        if estudiantil:
            texto = (
                "Turismo estudiantil: los alumnos van a habitaciones del cupo, nunca mezclados "
                "con los acompañantes, y los adultos se agrupan por rol y sexo en la habitación "
                "más justa que entre (2 apoderadas = un duplo, 3 = un triplo, 1 = un single).\n"
                "El máximo de plazas evita que los alumnos entren en habitaciones muy grandes."
            )
        else:
            texto = (
                "Receptivo: el reparto es libre, sin separar alumnos de acompañantes "
                "(familias, parejas, grupos de adultos). Los ajustes de arriba quedan apagados."
            )
        self.etiqueta_modo_help.configure(text=texto)
        for widget in self.controles_estudiantiles:
            widget.configure(state="normal" if estudiantil else "disabled")

    def _refrescar_tabla(self):
        self.etiqueta_conteo_pax.configure(text=f"Pasajeros ({len(self.pax_editable)})")
        for item in self.tabla.get_children():
            self.tabla.delete(item)
        for i, p in enumerate(sorted(self.pax_editable, key=lambda x: x.nombre.lower())):
            self.tabla.insert("", "end", iid=str(i), values=(
                p.nombre,
                p.tipo.value,
                p.sexo.value,
                "" if p.edad is None else p.edad,
                p.familia or "",
                p.pareja_de or "",
                "SÍ" if p.requiere_single else "",
            ))

    def _seleccionados(self) -> list[Pax]:
        elegidos = []
        for iid in self.tabla.selection():
            elegidos.append(self.pax_editable[int(iid)])
        return elegidos

    def _agregar_pax(self):
        dialogo = DialogoPax(self, [p.nombre for p in self.pax_editable])
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.pax_editable.append(dialogo.resultado)
            self._refrescar_tabla()

    def _agregar_lote(self):
        crudo = self.entrada_lote_cantidad.get().strip()
        try:
            cantidad = int(crudo)
        except ValueError:
            messagebox.showerror("Cantidad inválida",
                                 "Ingresá cuántos pasajeros querés cargar (número entero).",
                                 parent=self)
            return
        if cantidad < 1:
            messagebox.showerror("Cantidad inválida", "La cantidad debe ser al menos 1.",
                                 parent=self)
            return
        tipo = tipo_desde_valor(self.menu_lote_tipo.get())
        sexo = Sexo(self.sexo_lote.get())
        patron = re.compile(r"^Pax\s+(\d+)$", re.IGNORECASE)
        mayor = 0
        for p in self.pax_editable:
            coincide = patron.match(p.nombre.strip())
            if coincide:
                mayor = max(mayor, int(coincide.group(1)))
        self.pax_editable.extend(crear_pax_en_lote(cantidad, tipo, sexo, desde=mayor + 1))
        self._refrescar_tabla()
        self.etiqueta_conteo_pax.configure(text_color="#4ade80")
        self.after(1200, lambda: self.etiqueta_conteo_pax.configure(text_color="#f5f5f5"))

    def _editar_pax(self):
        seleccion = self.tabla.selection()
        if len(seleccion) != 1:
            messagebox.showinfo("Elegí uno", "Seleccioná exactamente un pasajero para editar.", parent=self)
            return
        indice = int(seleccion[0])
        dialogo = DialogoPax(self, [p.nombre for p in self.pax_editable], pax=self.pax_editable[indice])
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.pax_editable[indice] = dialogo.resultado
            self._refrescar_tabla()

    def _quitar_pax(self):
        seleccion = self.tabla.selection()
        if not seleccion:
            return
        for iid in sorted(seleccion, key=int, reverse=True):
            del self.pax_editable[int(iid)]
        self._refrescar_tabla()

    def _vincular_pareja(self):
        elegidos = self._seleccionados()
        if len(elegidos) != 2:
            messagebox.showinfo("Vincular pareja", "Seleccioná exactamente dos pasajeros.", parent=self)
            return
        a, b = elegidos
        a.pareja_de, b.pareja_de = b.nombre, a.nombre
        self._refrescar_tabla()

    def _desvincular(self):
        elegidos = self._seleccionados()
        for p in elegidos:
            if p.pareja_de:
                clave = p.pareja_de.strip().lower()
                p.pareja_de = None
                for otro in self.pax_editable:
                    if otro.nombre.strip().lower() == clave:
                        otro.pareja_de = None
        self._refrescar_tabla()

    def _alternar_single(self):
        elegidos = self._seleccionados()
        for p in elegidos:
            p.requiere_single = not p.requiere_single
        self._refrescar_tabla()

    def _guardar(self):
        nombre = self.entrada_nombre.get().strip()
        if not nombre:
            messagebox.showerror("Falta el nombre", "El grupo necesita un nombre.", parent=self)
            return
        if not self.pax_editable:
            messagebox.showerror("Grupo vacío", "Agregá al menos un pasajero.", parent=self)
            return
        texto_ingreso = self.entrada_ingreso.get().strip()
        texto_egreso = self.entrada_egreso.get().strip()
        if texto_ingreso and texto_egreso:
            try:
                ingreso = parsear_fecha(texto_ingreso)
                egreso = parsear_fecha(texto_egreso)
            except ValueError as exc:
                messagebox.showerror("Fecha inválida", str(exc), parent=self)
                return
            if ingreso >= egreso:
                messagebox.showerror("Fechas inconsistentes",
                                     "El egreso debe ser posterior al ingreso.", parent=self)
                return
        elif texto_ingreso or texto_egreso:
            messagebox.showerror("Fechas incompletas",
                                 "Completá las dos fechas (ingreso y egreso) o dejá ambas vacías "
                                 "para el grupo en cola de espera (sin fecha).",
                                 parent=self)
            return
        else:
            ingreso = None
            egreso = None
        cupo_crudo = self.entrada_cupo.get().strip()
        max_ninos_crudo = self.entrada_max_ninos.get().strip()
        max_plazas_crudo = self.entrada_max_plazas.get().strip()
        try:
            cupo = int(cupo_crudo) if cupo_crudo else None
            max_ninos = int(max_ninos_crudo) if max_ninos_crudo else None
            max_plazas = int(max_plazas_crudo) if max_plazas_crudo else None
        except ValueError:
            messagebox.showerror("Número inválido",
                                 "El cupo, el máximo de niños y el máximo de plazas "
                                 "deben ser números enteros.",
                                 parent=self)
            return
        if max_plazas is not None and max_plazas < 1:
            messagebox.showerror("Número inválido",
                                 "El máximo de plazas por habitación de estudiantes debe ser >= 1.",
                                 parent=self)
            return
        config = ConfigGrupo(
            separar_sexos=self.vars_reglas["separar_sexos"].get(),
            guias_choferes_single=self.vars_reglas["guias_choferes_single"].get(),
            ninos_con_adulto_obligatorio=self.vars_reglas["ninos_con_adulto_obligatorio"].get(),
            familias_exentas_sexo=self.vars_reglas["familias_exentas_sexo"].get(),
            bebes_no_cuentan_cupo=self.vars_reglas["bebes_no_cuentan_cupo"].get(),
            estudiantes_por_habitacion=cupo,
            estudiantes_estricto=self.var_estricto.get(),
            max_ninos_por_habitacion=max_ninos,
            modo_hospedaje=self.modo_hospedaje.get(),
            max_plazas_estudiantes=max_plazas,
            adultos_por_rol=self.vars_reglas["adultos_por_rol"].get(),
        )
        from reglas import atomizar
        grupo = Grupo(nombre=nombre, fecha_ingreso=ingreso, fecha_egreso=egreso,
                      pax=list(self.pax_editable), config=config,
                      empresa=self.entrada_empresa.get().strip(),
                      provisional=self.modo.get().startswith("INDIVIDUAL"))
        try:
            if grupo.tiene_fechas:
                atomizar(grupo)
        except ErrorDistribucion as exc:
            messagebox.showerror("Reglas incumplidas", str(exc), parent=self)
            return
        self.resultado = grupo
        self.destroy()


class DialogoEditarHabitacion(ctk.CTkToplevel):
    def __init__(self, master, habitacion: Habitacion):
        super().__init__(master)
        self.habitacion = habitacion
        self.resultado = False
        self.title(f"Editar habitación {habitacion.numero}")
        self.geometry("360x260")
        self.resizable(False, False)
        self.transient(master)
        try:
            self.grab_set()
        except tk.TclError:
            pass
        marco = ctk.CTkFrame(self)
        marco.pack(fill="both", expand=True, padx=16, pady=16)
        ctk.CTkLabel(marco, text=f"Habitación {habitacion.numero} · Piso {piso_de(habitacion.numero)}",
                     font=ctk.CTkFont(size=15, weight="bold"), anchor="w").pack(fill="x")
        ctk.CTkLabel(marco, text=f"Camas actuales: {habitacion.capacidad}",
                     text_color="#9ca3af", anchor="w").pack(fill="x", pady=(2, 10))
        ctk.CTkLabel(marco, text="Cantidad de camas (plazas)", anchor="w").pack(fill="x")
        self.entrada_capacidad = ctk.CTkEntry(marco, placeholder_text="ej: 3")
        self.entrada_capacidad.insert(0, str(habitacion.capacidad))
        self.entrada_capacidad.pack(fill="x", pady=(2, 6))
        ctk.CTkLabel(marco,
                     text="Sumá o quitá camas según cómo esté armada la unidad.\n"
                          "El mapa y la próxima simulación usan el nuevo valor.",
                     font=ctk.CTkFont(size=11), text_color="#9ca3af",
                     justify="left", wraplength=300).pack(fill="x", pady=(0, 10))
        botones = ctk.CTkFrame(marco, fg_color="transparent")
        botones.pack(side="bottom", fill="x", pady=(12, 0))
        ctk.CTkButton(botones, text="Cancelar", fg_color="gray30",
                      command=self.destroy).pack(side="left", expand=True, padx=4)
        ctk.CTkButton(botones, text="Guardar camas", command=self._confirmar).pack(
            side="left", expand=True, padx=4)
        self.entrada_capacidad.focus_set()

    def _confirmar(self):
        crudo = self.entrada_capacidad.get().strip()
        try:
            capacidad = int(crudo)
        except ValueError:
            messagebox.showerror("Valor invalido",
                                 "La cantidad de camas debe ser un numero entero.", parent=self)
            return
        if capacidad < 1:
            messagebox.showerror("Valor invalido", "Debe haber al menos 1 cama.", parent=self)
            return
        self.habitacion.capacidad = capacidad
        self.resultado = True
        self.destroy()


class DialogoCamasLote(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.resultado = False
        self.title("Cargar camas (inventario)")
        self.geometry("400x420")
        try:
            self.grab_set()
        except TclError:
            pass
        ctk.CTkLabel(self, text="Nuevo lote de camas del inventario",
                     font=ctk.CTkFont(size=15, weight="bold")).pack(pady=(14, 2))
        ctk.CTkLabel(self, text="Las camas se cargan aparte de los apartamentos:\n"
                                "sirven para saber cuántas personas podés alojar por día.",
                     font=ctk.CTkFont(size=11), text_color="#9ca3af",
                     justify="center").pack(pady=(0, 10))

        fila_titulo = ctk.CTkFrame(self, fg_color="transparent")
        fila_titulo.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(fila_titulo, text="Nombre:", width=110, anchor="w").pack(side="left")
        self.entrada_titulo = ctk.CTkEntry(fila_titulo,
                                           placeholder_text="ej: Dormis piso 3")
        self.entrada_titulo.pack(side="left", expand=True, fill="x")

        fila_cantidad = ctk.CTkFrame(self, fg_color="transparent")
        fila_cantidad.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(fila_cantidad, text="Camas:", width=110, anchor="w").pack(side="left")
        self.entrada_cantidad = ctk.CTkEntry(fila_cantidad, placeholder_text="ej: 40")
        self.entrada_cantidad.pack(side="left", expand=True, fill="x")

        fila_desde = ctk.CTkFrame(self, fg_color="transparent")
        fila_desde.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(fila_desde, text="Disponible desde:", width=110,
                     anchor="w").pack(side="left")
        self.entrada_desde = ctk.CTkEntry(fila_desde, placeholder_text="dd/mm/aaaa (vacío = siempre)")
        self.entrada_desde.pack(side="left", expand=True, fill="x")

        fila_hasta = ctk.CTkFrame(self, fg_color="transparent")
        fila_hasta.pack(fill="x", padx=20, pady=4)
        ctk.CTkLabel(fila_hasta, text="Disponible hasta:", width=110,
                     anchor="w").pack(side="left")
        self.entrada_hasta = ctk.CTkEntry(fila_hasta, placeholder_text="dd/mm/aaaa (vacío = indefinido)")
        self.entrada_hasta.pack(side="left", expand=True, fill="x")

        ctk.CTkButton(self, text="＋  Agregar lote de camas", height=36,
                      command=self._confirmar).pack(fill="x", padx=20, pady=(16, 16))
        self.entrada_titulo.focus_set()

    def _confirmar(self):
        titulo = self.entrada_titulo.get().strip() or "Camas"
        texto_cantidad = self.entrada_cantidad.get().strip()
        if not texto_cantidad.isdigit() or int(texto_cantidad) < 1:
            messagebox.showerror("Cantidad inválida",
                                 "Indicá cuántas camas tiene el lote.")
            return
        texto_desde = self.entrada_desde.get().strip()
        texto_hasta = self.entrada_hasta.get().strip()
        desde = hasta = None
        try:
            if texto_desde:
                desde = parsear_fecha(texto_desde)
            if texto_hasta:
                hasta = parsear_fecha(texto_hasta)
        except ValueError as error:
            messagebox.showerror("Fecha inválida", str(error))
            return
        if desde and hasta and desde >= hasta:
            messagebox.showerror("Fechas inválidas",
                                 "La fecha 'hasta' debe ser posterior a la fecha 'desde'.")
            return
        lote = LoteCamas(titulo=titulo, cantidad=int(texto_cantidad),
                         disponible_desde=desde, disponible_hasta=hasta)
        errores = lote.validar()
        if errores:
            messagebox.showerror("Datos inválidos", "\n".join(errores))
            return
        self.lote = lote
        self.resultado = True
        self.destroy()


class DialogoNombreHotel(ctk.CTkToplevel):
    def __init__(self, master, titulo: str, etiqueta: str, placeholder: str,
                 por_defecto: str | None = None):
        super().__init__(master)
        self.resultado = False
        self.nombre = ""
        self.title(titulo)
        self.geometry("380x180")
        try:
            self.grab_set()
        except TclError:
            pass
        ctk.CTkLabel(self, text=etiqueta, font=ctk.CTkFont(size=14),
                     anchor="w").pack(fill="x", padx=20, pady=(20, 6))
        self.entrada = ctk.CTkEntry(self, placeholder_text=placeholder)
        self.entrada.pack(fill="x", padx=20)
        if por_defecto:
            self.entrada.insert(0, por_defecto)
        ctk.CTkButton(self, text="✔  Aceptar", height=34,
                      command=self._confirmar).pack(fill="x", padx=20, pady=(18, 0))
        self.entrada.focus_set()
        self.entrada.bind("<Return>", lambda _e: self._confirmar())

    def _confirmar(self):
        nombre = self.entrada.get().strip()
        if not nombre:
            messagebox.showerror("Nombre vacío", "Poné un nombre.")
            return
        self.nombre = nombre
        self.resultado = True
        self.destroy()


class DialogoRack(ctk.CTkToplevel):
    def __init__(self, master, hotel: Hotel):
        super().__init__(master)
        self.hotel = hotel
        self.title("Rack / Cotizador rapido")
        self.geometry("600x640")
        self.transient(master)
        try:
            self.grab_set()
        except tk.TclError:
            pass
        marco = ctk.CTkScrollableFrame(self)
        marco.pack(fill="both", expand=True, padx=14, pady=14)

        ctk.CTkLabel(marco, text="Fechas de la estadia",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        fila_fechas = ctk.CTkFrame(marco, fg_color="transparent")
        fila_fechas.pack(fill="x", pady=(2, 10))
        col_ini = ctk.CTkFrame(fila_fechas, fg_color="transparent")
        col_ini.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(col_ini, text="Ingreso", anchor="w").pack(fill="x")
        self.entrada_ingreso = ctk.CTkEntry(col_ini, width=130, placeholder_text="AAAA-MM-DD")
        self.entrada_ingreso.insert(0, str(date.today() + timedelta(days=7)))
        self.entrada_ingreso.pack(pady=(2, 0))
        col_fin = ctk.CTkFrame(fila_fechas, fg_color="transparent")
        col_fin.pack(side="left")
        ctk.CTkLabel(col_fin, text="Egreso", anchor="w").pack(fill="x")
        self.entrada_egreso = ctk.CTkEntry(col_fin, width=130, placeholder_text="AAAA-MM-DD")
        self.entrada_egreso.insert(0, str(date.today() + timedelta(days=11)))
        self.entrada_egreso.pack(pady=(2, 0))

        ctk.CTkLabel(marco, text="Que te pasaron?",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w")
        self.selector_modo = ctk.CTkSegmentedButton(
            marco, values=["Apartamentos", "Pasajeros"], command=self._actualizar_etiqueta)
        self.selector_modo.set("Apartamentos")
        self.selector_modo.pack(fill="x", pady=(2, 8))

        self.etiqueta_cantidad = ctk.CTkLabel(marco, anchor="w",
                                              text="Cantidad de apartamentos que piden")
        self.etiqueta_cantidad.pack(fill="x")
        self.entrada_cantidad = ctk.CTkEntry(marco, placeholder_text="ej: 20")
        self.entrada_cantidad.pack(fill="x", pady=(2, 8))

        ctk.CTkLabel(marco, text="Multiplicador manual camas/apartamento "
                                 "(vacio = automatico del hotel)", anchor="w").pack(fill="x")
        self.entrada_multiplicador = ctk.CTkEntry(marco, placeholder_text="ej: 3.5 o vacio")
        self.entrada_multiplicador.pack(fill="x", pady=(2, 10))

        ctk.CTkButton(marco, text="Calcular aproximacion", height=36,
                      command=self._calcular).pack(fill="x", pady=(0, 10))

        self.etiqueta_marcado = ctk.CTkLabel(marco, text="", anchor="w", justify="left",
                                             font=ctk.CTkFont(size=12, weight="bold"))
        self.etiqueta_marcado.pack(fill="x", pady=(0, 4))
        fila_marcar = ctk.CTkFrame(marco, fg_color="transparent")
        fila_marcar.pack(fill="x", pady=(0, 10))
        self.boton_marcar = ctk.CTkButton(fila_marcar, text="📌  Marcar cotización en ambos mapas",
                                          height=36, command=self._marcar_en_mapas)
        self.boton_marcar.pack(side="left", expand=True, fill="x", padx=(0, 4))
        self.boton_desmarcar = ctk.CTkButton(fila_marcar, text="❌ Quitar", width=80,
                                             fg_color="gray25", hover_color="#7f1d1d",
                                             height=36, command=self._quitar_marcado)
        self.boton_desmarcar.pack(side="right")
        self._actualizar_marcado()

        self.texto_resultado = tk.Text(marco, height=15, bg="#1e1e21", fg=TEXTO_CLARO,
                                       font=("Consolas", 11), state="disabled",
                                       relief="flat", wrap="word")
        self.texto_resultado.pack(fill="both", expand=True)

    def _actualizar_marcado(self):
        prov = self.hotel.provisional
        if prov:
            self.etiqueta_marcado.configure(
                text=(f"📌 Marcado ACTIVO: {prov['titulo']} · "
                      f"{formatear(prov['desde'])} → {formatear(prov['hasta'])}\n"
                      f"≈{prov['aptos_noche']} apartamentos/noche · "
                      f"{prov['pax_noche']} pax/noche"),
                text_color="#fbbf24")
            self.boton_desmarcar.configure(state="normal")
            self.boton_marcar.configure(text="🔄  Remarcar en ambos mapas")
        else:
            self.etiqueta_marcado.configure(
                text="Sin cotización marcada: los mapas muestran solo lo simulado.",
                text_color="#9ca3af")
            self.boton_desmarcar.configure(state="disabled")
            self.boton_marcar.configure(text="📌  Marcar cotización en ambos mapas")

    def _quitar_marcado(self):
        self.hotel.provisional = None
        self._actualizar_marcado()
        self._mostrar(["Cotización provisoria quitada de los mapas."])

    def _marcar_en_mapas(self):
        validos = self._datos_validados()
        if validos is None:
            return
        if validos[0] == "null":
            messagebox.showwarning(
                "Sin apartamentos libres",
                "No hay apartamentos libres en esas fechas.\n"
                "Si solo trabajás con camas, cargá el lote y usá Pasajeros",
                parent=self)
            return
        _ini, _fin, ratio, _origen, _total_hab, _total_camas, _caps = validos
        ini, fin = _ini, _fin
        crudo = self.entrada_cantidad.get().strip()
        try:
            n = int(crudo)
        except ValueError:
            messagebox.showerror("Cantidad invalida", "Ingresá un numero entero.", parent=self)
            return
        if n < 1:
            messagebox.showerror("Cantidad invalida", "Debe ser al menos 1.", parent=self)
            return
        if self.selector_modo.get() == "Apartamentos":
            titulo = f"{n} apartamentos"
            pax_noche = round(n * ratio)
            aptos_noche = n
        else:
            titulo = f"{n} pax"
            pax_noche = n
            aptos_noche = math.ceil(n / ratio)
        self.hotel.provisional = {
            "titulo": titulo,
            "desde": ini,
            "hasta": fin,
            "pax_noche": pax_noche,
            "aptos_noche": aptos_noche,
        }
        self._actualizar_marcado()
        self.master._refrescar_todo()
        self._mostrar([
            "Cotización marcada en ambos mapas.",
            f"   {titulo} · {formatear(ini)} al {formatear(fin)}",
            f"   ≈{aptos_noche} apartamentos/noche · {pax_noche} pax/noche",
            "El mapa de UH y el mapa de camas lo muestran como provisorio",
            "(ámbar) sobre lo ya simulado.",
        ])

    def _actualizar_etiqueta(self, *_args):
        if self.selector_modo.get() == "Apartamentos":
            self.etiqueta_cantidad.configure(text="Cantidad de apartamentos que piden")
            self.entrada_cantidad.configure(placeholder_text="ej: 20")
        else:
            self.etiqueta_cantidad.configure(text="Cantidad de pasajeros que piden")
            self.entrada_cantidad.configure(placeholder_text="ej: 80")

    def _mostrar(self, lineas: list[str]):
        self.texto_resultado.configure(state="normal")
        self.texto_resultado.delete("1.0", "end")
        self.texto_resultado.insert("1.0", "\n".join(lineas))
        self.texto_resultado.configure(state="disabled")

    def _greedy_apartamentos(self, pax: int, capacidades_desc: list[int],
                             capacidades_asc: list[int]) -> tuple[int | None, int | None]:
        def contar(lista: list[int]) -> int | None:
            acumulado = 0
            for i, cap in enumerate(lista, start=1):
                acumulado += cap
                if acumulado >= pax:
                    return i
            return None
        return contar(capacidades_desc), contar(capacidades_asc)

    def _datos_validados(self):
        try:
            ini = parsear_fecha(self.entrada_ingreso.get())
            fin = parsear_fecha(self.entrada_egreso.get())
        except ValueError as exc:
            messagebox.showerror("Fecha invalida", str(exc), parent=self)
            return None
        if ini >= fin:
            messagebox.showerror("Fechas inconsistentes",
                                 "El egreso debe ser posterior al ingreso.", parent=self)
            return None
        crudo_mult = self.entrada_multiplicador.get().strip()
        mult_manual = None
        if crudo_mult:
            try:
                mult_manual = float(crudo_mult.replace(",", "."))
            except ValueError:
                messagebox.showerror("Multiplicador invalido",
                                     "Debe ser un numero, ej: 3.5", parent=self)
                return None
            if mult_manual <= 0:
                messagebox.showerror("Multiplicador invalido",
                                     "El multiplicador debe ser mayor a cero.", parent=self)
                return None
        libres = [h for h in self.hotel.habitaciones if h.libre_en_rango(ini, fin)]
        caps = sorted(h.capacidad for h in libres)
        total_hab = len(caps)
        total_camas = sum(caps)
        if total_hab == 0:
            return ("null", ini, fin, 0, 0, 0, [])
        ratio = mult_manual if mult_manual else total_camas / total_hab
        origen = "manual" if mult_manual else "automatico"
        return (ini, fin, ratio, origen, total_hab, total_camas, caps)

    def _calcular(self):
        validos = self._datos_validados()
        if validos is None:
            return
        ini, fin, ratio, origen, total_hab, total_camas, caps = validos
        if total_hab == 0:
            self._mostrar([
                f"DISPONIBLE DEL {formatear(ini)} AL {formatear(fin)}",
                "No hay apartamentos libres en esas fechas (todo bloqueado u ocupado).",
                "Para marcar una cotización en mapas sin apartamentos,",
                "completá el multiplicador manual.",
            ])
            return
        lineas = [
            f"DISPONIBLE DEL {formatear(ini)} AL {formatear(fin)}",
            f"   Apartamentos libres: {total_hab} de {len(self.hotel.habitaciones)}",
            f"   Camas libres: {total_camas}",
        ]
        lineas.append(f"   Multiplicador: {ratio:g} camas/apartamento ({origen})")
        crudo = self.entrada_cantidad.get().strip()
        try:
            n = int(crudo)
        except ValueError:
            messagebox.showerror("Cantidad invalida",
                                 "Ingresá un numero entero.", parent=self)
            return
        if n < 1:
            messagebox.showerror("Cantidad invalida", "Debe ser al menos 1.", parent=self)
            return
        lineas.append("")
        if self.selector_modo.get() == "Apartamentos":
            estimado = round(n * ratio)
            lineas.append(f"PIDEN {n} APARTAMENTOS:")
            lineas.append(f"   Pax estimados: ≈{estimado}")
            if n <= total_hab:
                minimo = sum(caps[:n])
                maximo = sum(caps[-n:])
                lineas.append(f"   Rango posible segun mezcla: entre {minimo}"
                              f" (las mas chicas) y {maximo} (las mas grandes)")
            else:
                lineas.append(f"   OJO: pedis mas que los {total_hab} libres esas fechas;")
                lineas.append("   la estimacion extrapola el promedio.")
        else:
            estimado = math.ceil(n / ratio)
            lineas.append(f"PIDEN {n} PAX:")
            lineas.append(f"   Apartamentos estimados: ≈{estimado}")
            if n > total_camas:
                lineas.append(f"   OJO: superan las {total_camas} camas libres esas fechas.")
            else:
                minimo, maximo = self._greedy_apartamentos(n, caps[::-1], caps)
                lineas.append(f"   Mejor caso: {minimo} apartamentos (usa los mas grandes)")
                lineas.append(f"   Peor caso: {maximo} apartamentos (usa los mas chicos)")
        lineas.append("")
        lineas.append("Es una aproximacion para cotizar; la distribucion exacta")
        lineas.append("sale cargando el grupo y presionando SIMULAR.")
        self._mostrar(lineas)


class DialogoPlanificador(ctk.CTkToplevel):
    def __init__(self, master):
        super().__init__(master)
        self.title("Planificador / Estimador de distribución")
        self.geometry("640x720")
        self.minsize(580, 640)
        self.transient(master)
        self.aplicado = False
        self.resultado_grupo = None
        self.resultado_hotel_estimado = None
        self.resultado_sugerencia = None
        try:
            self.grab_set()
        except tk.TclError:
            pass

        marco = ctk.CTkScrollableFrame(self, orientation="vertical")
        marco.pack(fill="both", expand=True, padx=14, pady=14)

        seccion_pax = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        seccion_pax.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(seccion_pax, text="DATOS DEL GRUPO",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))
        fila_nombre = ctk.CTkFrame(seccion_pax, fg_color="transparent")
        fila_nombre.pack(fill="x", padx=10, pady=(4, 6))
        ctk.CTkLabel(fila_nombre, text="Nombre:", width=100, anchor="w").pack(side="left")
        self.entrada_nombre = ctk.CTkEntry(fila_nombre, placeholder_text="ej: Colegio Norte")
        self.entrada_nombre.pack(side="left", expand=True, fill="x")

        fila_total = ctk.CTkFrame(seccion_pax, fg_color="transparent")
        fila_total.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(fila_total, text="Total pasajeros:", width=110, anchor="w").pack(side="left")
        self.entrada_total_pax = ctk.CTkEntry(fila_total, placeholder_text="ej: 45", width=80)
        self.entrada_total_pax.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(fila_total, text="(solo número; si sabés más, completá abajo)",
                     font=ctk.CTkFont(size=11), text_color="#6b7280").pack(side="left")

        fila_fechas = ctk.CTkFrame(seccion_pax, fg_color="transparent")
        fila_fechas.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(fila_fechas, text="Ingreso:", width=100, anchor="w").pack(side="left")
        self.entrada_ingreso = ctk.CTkEntry(fila_fechas, placeholder_text="AAAA-MM-DD", width=130)
        self.entrada_ingreso.insert(0, str(date.today() + timedelta(days=7)))
        self.entrada_ingreso.pack(side="left", padx=(0, 8))
        ctk.CTkLabel(fila_fechas, text="Egreso:", anchor="w").pack(side="left")
        self.entrada_egreso = ctk.CTkEntry(fila_fechas, placeholder_text="AAAA-MM-DD", width=130)
        self.entrada_egreso.insert(0, str(date.today() + timedelta(days=11)))
        self.entrada_egreso.pack(side="left")

        seccion_comp = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        seccion_comp.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(seccion_comp, text="COMPOSICIÓN (opcional)",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))
        fila_comp = ctk.CTkFrame(seccion_comp, fg_color="transparent")
        fila_comp.pack(fill="x", padx=10, pady=(0, 8))
        c1 = ctk.CTkFrame(fila_comp, fg_color="transparent")
        c1.pack(side="left", padx=(0, 10))
        ctk.CTkLabel(c1, text="Varones", anchor="w").pack(fill="x")
        self.entrada_varones = ctk.CTkEntry(c1, width=70, placeholder_text="0")
        self.entrada_varones.pack(fill="x")
        c2 = ctk.CTkFrame(fila_comp, fg_color="transparent")
        c2.pack(side="left")
        ctk.CTkLabel(c2, text="Mujeres", anchor="w").pack(fill="x")
        self.entrada_mujeres = ctk.CTkEntry(c2, width=70, placeholder_text="0")
        self.entrada_mujeres.pack(fill="x")
        ctk.CTkLabel(seccion_comp, text="Si dejás en 0, todos se tratan como indistintos (mejor llenado).",
                     font=ctk.CTkFont(size=11), text_color="#6b7280").pack(anchor="w", padx=10, pady=(0, 8))

        seccion_hotel = ctk.CTkFrame(marco, fg_color="#242427", corner_radius=8)
        seccion_hotel.pack(fill="x", pady=(0, 10))
        ctk.CTkLabel(seccion_hotel, text="HOTEL (si lo conocés)",
                     font=ctk.CTkFont(size=13, weight="bold")).pack(anchor="w", padx=10, pady=(8, 2))

        fila_cap = ctk.CTkFrame(seccion_hotel, fg_color="transparent")
        fila_cap.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(fila_cap, text="Capacidad máxima:", width=130, anchor="w").pack(side="left")
        self.entrada_cap_max = ctk.CTkEntry(fila_cap, width=60, placeholder_text="ej: 4")
        self.entrada_cap_max.pack(side="left")
        ctk.CTkLabel(fila_cap, text=" plazas por apartamento",
                     font=ctk.CTkFont(size=11), text_color="#9ca3af").pack(side="left")

        fila_hab = ctk.CTkFrame(seccion_hotel, fg_color="transparent")
        fila_hab.pack(fill="x", padx=10, pady=(0, 6))
        ctk.CTkLabel(fila_hab, text="Total apartamentos:", width=130, anchor="w").pack(side="left")
        self.entrada_total_hab = ctk.CTkEntry(fila_hab, width=60, placeholder_text="vacío = sin tope")
        self.entrada_total_hab.pack(side="left")

        fila_ratio = ctk.CTkFrame(seccion_hotel, fg_color="transparent")
        fila_ratio.pack(fill="x", padx=10, pady=(0, 8))
        ctk.CTkLabel(fila_ratio, text="Ratio camas/apart:", width=130, anchor="w").pack(side="left")
        self.entrada_ratio = ctk.CTkEntry(fila_ratio, width=60, placeholder_text="ej: 3.5")
        self.entrada_ratio.pack(side="left")
        ctk.CTkLabel(fila_ratio, text="(vacío = auto)",
                     font=ctk.CTkFont(size=11), text_color="#6b7280").pack(side="left")

        botones = ctk.CTkFrame(marco, fg_color="transparent")
        botones.pack(fill="x", pady=(4, 6))
        ctk.CTkButton(botones, text="📐  Calcular / Ver resultado", height=36,
                      command=self._calcular).pack(fill="x")

        self.texto_resultado = ctk.CTkTextbox(
            marco, font=ctk.CTkFont(family="Consolas", size=13),
            fg_color="#1a1a1d", wrap="word", height=180,
        )
        self.texto_resultado.pack(fill="both", expand=True)
        self.texto_resultado.insert("1.0", "Completá los datos y presioná «Calcular».")
        self.texto_resultado.configure(state="disabled")

        botones_finales = ctk.CTkFrame(marco, fg_color="transparent")
        botones_finales.pack(fill="x", pady=(8, 0))
        ctk.CTkButton(botones_finales, text="Cancelar", fg_color="gray30", height=34,
                      command=self.destroy).pack(side="left", expand=True, padx=4)
        self.boton_aplicar = ctk.CTkButton(
            botones_finales, text="✔ Aplicar al programa", height=34,
            state="disabled", command=self._aplicar,
        )
        self.boton_aplicar.pack(side="left", expand=True, padx=4)

    def _leer_int(self, entrada, nombre, obligatorio=True):
        texto = entrada.get().strip()
        if not texto:
            if obligatorio:
                raise ValueError(f"Falta {nombre}.")
            return None
        try:
            return int(texto)
        except ValueError:
            raise ValueError(f"{nombre} debe ser un número entero (recibiste: '{texto}').")

    def _leer_float(self, entrada, nombre):
        texto = entrada.get().strip()
        if not texto:
            return None
        try:
            return float(texto.replace(",", "."))
        except ValueError:
            raise ValueError(f"{nombre} debe ser un número (ej: 3.5).")

    def _calcular(self):
        try:
            total_pax = self._leer_int(self.entrada_total_pax, "Total pasajeros")
            varones = self._leer_int(self.entrada_varones, "Varones", obligatorio=False) or 0
            mujeres = self._leer_int(self.entrada_mujeres, "Mujeres", obligatorio=False) or 0
            fecha_ingreso = parsear_fecha(self.entrada_ingreso.get())
            fecha_egreso = parsear_fecha(self.entrada_egreso.get())
            cap_max = self._leer_int(self.entrada_cap_max, "Capacidad máxima")
            total_hab = self._leer_int(self.entrada_total_hab, "Total apartamentos", obligatorio=False)
            ratio = self._leer_float(self.entrada_ratio, "Ratio camas/apartamento")
        except ValueError as exc:
            messagebox.showerror("Datos inválidos", str(exc), parent=self)
            return
        if fecha_ingreso >= fecha_egreso:
            messagebox.showerror("Fechas inválidas", "El egreso debe ser posterior al ingreso.", parent=self)
            return

        try:
            grupo = crear_grupo_generico(
                nombre=self.entrada_nombre.get().strip() or "Grupo Estimado",
                fecha_ingreso=fecha_ingreso,
                fecha_egreso=fecha_egreso,
                total_pax=total_pax,
                varones=varones,
                mujeres=mujeres,
            )
            atomos, avisos_grupo = atomizar(grupo)

            if total_hab is not None and total_hab < 1:
                messagebox.showerror("Valor inválido",
                                     "El total de apartamentos debe ser >= 1.", parent=self)
                return

            hotel_estimado = generar_hotel_estimado(
                nombre="Hotel Estimado",
                total_habitaciones=total_hab or max(1, round(total_pax / cap_max)),
                capacidad_max=cap_max,
                ratio=ratio,
            )

            sug = sugerir_mezcla(
                atomos,
                [h.capacidad for h in hotel_estimado.habitaciones],
                config=grupo.config,
                total_habitaciones=total_hab,
            )

            self.resultado_sugerencia = sug
            self.resultado_grupo = grupo
            self.resultado_hotel_estimado = hotel_estimado

            lineas = [
                f"RESULTADO PARA '{grupo.nombre}'",
                f"Pax: {total_pax} | Fechas: {formatear(fecha_ingreso)} → {formatear(fecha_egreso)}",
                f"Hotel estimado: {len(hotel_estimado.habitaciones)} apartamentos "
                f"(cap. máx {cap_max}, ratio {ratio or 'auto'})",
                f"Camas totales hotel: {hotel_estimado.total_camas()}",
                "",
                "MEZCLA SUGERIDA (óptima):",
                f"  {sug.texto()}",
                f"  Total habitaciones usadas: {sug.habitaciones}",
                f"  Camas ocupadas: {sug.pax_alojados}/{sug.camas} "
                f"(desperdicio: {sug.desperdicio})",
                "",
                f"Estado: {'COMPLETO' if sug.completo else 'INCOMPLETO (faltan habitaciones)'}",
            ]
            if avisos_grupo:
                lineas.append("")
                for a in avisos_grupo:
                    lineas.append(f"  AVISO: {a}")
            if sug.avisos:
                lineas.append("")
                for a in sug.avisos:
                    lineas.append(f"  {a}")
            if sug.completo:
                lineas.append("")
                lineas.append("Se puede registrar: se crea el hotel estimado y el grupo genérico.")
            else:
                lineas.append("")
                lineas.append("No se puede registrar: no alcanzan las habitaciones del hotel estimado.")

            self.texto_resultado.configure(state="normal")
            self.texto_resultado.delete("1.0", "end")
            self.texto_resultado.insert("1.0", "\n".join(lineas))
            self.texto_resultado.configure(state="disabled")
            self.boton_aplicar.configure(state="normal" if sug.completo else "disabled")

        except ErrorDistribucion as exc:
            messagebox.showerror("Error de distribución", str(exc), parent=self)

    def _aplicar(self):
        if not (self.resultado_sugerencia and self.resultado_sugerencia.completo):
            return
        if not messagebox.askyesno(
            "Aplicar estimación",
            "Esto va a:\n"
            "  • Agregar las habitaciones del hotel estimado al hotel actual\n"
            "  • Agregar el grupo genérico a la lista de grupos\n\n"
            "¿Continuar?",
            parent=self,
        ):
            return
        master = self.master
        master.hotel.habitaciones.extend(self.resultado_hotel_estimado.habitaciones)
        master.grupos.append(self.resultado_grupo)
        self.aplicado = True
        self.destroy()


class App(ctk.CTk):
    def __init__(self):
        super().__init__()
        ctk.set_appearance_mode("dark")
        self.title("Distribuidor de Habitaciones para Grupos")
        self.geometry("1360x800")
        self.minsize(1120, 680)
        self._avisos_error = 0
        self.hoteles: list[Hotel] = [Hotel(nombre="Mi Hotel")]
        self.grupos_por_hotel: list[list[Grupo]] = [[]]
        self.reporte_por_hotel: list = [None]
        self.seleccion_por_hotel: list[str | None] = [None]
        self.indice_hotel: int = 0
        self.hotel: Hotel = self.hoteles[0]
        self.grupos: list[Grupo] = self.grupos_por_hotel[0]
        self.reporte = None
        self.ruta_actual: str | None = None
        self.numero_seleccionado: str | None = None
        self.indice_grupo_seleccionado: int | None = None
        self.tarjetas_grupos: list[ctk.CTkButton] = []

        self.grid_columnconfigure(1, weight=1)
        self.grid_rowconfigure(1, weight=1)
        self._construir_barra_superior()
        self._construir_lateral()
        self._construir_zona_principal()
        self._construir_barra_estado()
        self._job_autoguardado = None
        self._autoguardado_activo = "--selftest" not in sys.argv
        if "--selftest" not in sys.argv:
            self._auto_iniciar()
        self._refrescar_todo()

    def _guardar_estado_hotel(self):
        self.grupos_por_hotel[self.indice_hotel] = self.grupos
        self.reporte_por_hotel[self.indice_hotel] = self.reporte
        self.seleccion_por_hotel[self.indice_hotel] = self.numero_seleccionado

    def _restaurar_estado_hotel(self):
        self.hotel = self.hoteles[self.indice_hotel]
        self.grupos = self.grupos_por_hotel[self.indice_hotel]
        self.reporte = self.reporte_por_hotel[self.indice_hotel]
        self.numero_seleccionado = self.seleccion_por_hotel[self.indice_hotel]

    def _auto_iniciar(self):
        if not os.path.exists(RUTA_ESPACIO_TRABAJO):
            return
        try:
            hoteles, grupos_por_hotel, estado = cargar_area(RUTA_ESPACIO_TRABAJO)
        except ErrorDistribucion:
            return
        self.hoteles, self.grupos_por_hotel = hoteles, grupos_por_hotel
        self.indice_hotel = int(estado.get("indice_hotel", 0))
        self.indice_hotel = max(0, min(self.indice_hotel, len(self.hoteles) - 1))
        self.reporte_por_hotel = [None for _ in hoteles]
        self.seleccion_por_hotel = [None for _ in hoteles]
        self._restaurar_estado_hotel()
        fecha_d = estado.get("vista_desde", "")
        fecha_h = estado.get("vista_hasta", "")
        if fecha_d and fecha_h:
            try:
                parsear_fecha(fecha_d)
                parsear_fecha(fecha_h)
                self.entrada_vista_desde.delete(0, "end")
                self.entrada_vista_desde.insert(0, fecha_d)
                self.entrada_vista_hasta.delete(0, "end")
                self.entrada_vista_hasta.insert(0, fecha_h)
            except ValueError:
                pass
        self.selector_hoteles.configure(values=[h.nombre for h in self.hoteles])
        self.selector_hoteles.set(self.hotel.nombre)
        self.indice_grupo_seleccionado = None
        self._status(f"Espacio de trabajo cargado: {len(self.hoteles)} hotel(es), "
                     f"{sum(len(g) for g in self.grupos_por_hotel)} grupo(s).")

    def _programar_autoguardado(self):
        if not self._autoguardado_activo:
            return
        if self._job_autoguardado is not None:
            self.after_cancel(self._job_autoguardado)
        self._job_autoguardado = self.after(900, self._ejecutar_autoguardado)

    def _ejecutar_autoguardado(self):
        self._job_autoguardado = None
        try:
            self._guardar_estado_hotel()
            os.makedirs(CARPETA_EXPORTACIONES, exist_ok=True)
            estado = {
                "indice_hotel": self.indice_hotel,
                "vista_desde": self.entrada_vista_desde.get().strip(),
                "vista_hasta": self.entrada_vista_hasta.get().strip(),
            }
            guardar_multi(RUTA_ESPACIO_TRABAJO, self.hoteles, self.grupos_por_hotel,
                          estado=estado)
        except ErrorDistribucion:
            pass

    def _sync_botones_resultado(self):
        if self.reporte is not None:
            self.boton_reservar.configure(state="normal")
            self.boton_liberar.configure(state="normal")
        else:
            self._bloquear_acciones_resultado()

    def _cambiar_hotel(self, indice: int):
        if indice == self.indice_hotel or not (0 <= indice < len(self.hoteles)):
            return
        self._guardar_estado_hotel()
        self.indice_hotel = indice
        self._restaurar_estado_hotel()
        self.indice_grupo_seleccionado = None
        self._sync_botones_resultado()
        self._refrescar_todo()
        self._status(f"Hotel actual: {self.hotel.nombre}.")

    def _seleccionar_hotel_por_nombre(self, nombre: str):
        for i, h in enumerate(self.hoteles):
            if h.nombre == nombre:
                self._cambiar_hotel(i)
                return

    def _renombrar_hotel(self):
        nombre = self.etiqueta_hotel.get().strip() or self.hotel.nombre
        if nombre != self.hotel.nombre:
            duplicado = any(h is not self.hotel and h.nombre == nombre
                            for h in self.hoteles)
            if duplicado:
                messagebox.showwarning("Nombre duplicado",
                                       "Ya existe un hotel con ese nombre.")
                self.etiqueta_hotel.delete(0, "end")
                self.etiqueta_hotel.insert(0, self.hotel.nombre)
                return
            self.hotel.nombre = nombre
            self.selector_hoteles.configure(
                values=[h.nombre for h in self.hoteles])
            self.selector_hoteles.set(nombre)
            self._status(f"Hotel renombrado a «{nombre}».")

    def agregar_hotel(self):
        dialogo = DialogoNombreHotel(self, titulo="Nuevo hotel",
                                     etiqueta="Nombre del hotel:",
                                     placeholder="ej: Cabañas del Lago", por_defecto=None)
        self.wait_window(dialogo)
        if dialogo.resultado:
            self._crear_hotel(dialogo.nombre)

    def _crear_hotel(self, nombre: str):
        self._guardar_estado_hotel()
        self.hoteles.append(Hotel(nombre=nombre))
        self.grupos_por_hotel.append([])
        self.reporte_por_hotel.append(None)
        self.seleccion_por_hotel.append(None)
        self.indice_hotel = len(self.hoteles) - 1
        self._restaurar_estado_hotel()
        self.indice_grupo_seleccionado = None
        self.selector_hoteles.configure(values=[h.nombre for h in self.hoteles])
        self.selector_hoteles.set(nombre)
        self._bloquear_acciones_resultado()
        self._refrescar_todo()
        self._status(f"Hotel «{nombre}» creado. Agregá habitaciones o camas.")

    def eliminar_hotel(self):
        if len(self.hoteles) <= 1:
            messagebox.showwarning("No se puede",
                                   "Debe existir al menos un hotel.")
            return
        if not messagebox.askyesno(
                "Eliminar hotel",
                f"¿Eliminar «{self.hotel.nombre}» y sus {len(self.grupos)} grupo(s)? "
                "Los datos no se pueden recuperar."):
            return
        indice = self.indice_hotel
        self.hoteles.pop(indice)
        self.grupos_por_hotel.pop(indice)
        self.reporte_por_hotel.pop(indice)
        self.seleccion_por_hotel.pop(indice)
        self.indice_hotel = max(0, indice - 1)
        self._restaurar_estado_hotel()
        self.indice_grupo_seleccionado = None
        self.selector_hoteles.configure(values=[h.nombre for h in self.hoteles])
        self.selector_hoteles.set(self.hotel.nombre)
        self._sync_botones_resultado()
        self._refrescar_todo()
        self._status(f"Hotel eliminado. Actual: {self.hotel.nombre}.")

    def _construir_barra_superior(self):
        barra = ctk.CTkFrame(self, corner_radius=0, fg_color="#202023")
        barra.grid(row=0, column=0, columnspan=2, sticky="ew", padx=0, pady=0)
        ctk.CTkLabel(barra, text="🏨  Distribuidor de Habitaciones para Grupos",
                     font=ctk.CTkFont(size=19, weight="bold")).pack(side="left", padx=16, pady=10)
        self.etiqueta_hotel = ctk.CTkEntry(barra, width=220)
        self.etiqueta_hotel.insert(0, self.hotel.nombre)
        self.etiqueta_hotel.pack(side="left", padx=8)
        self.etiqueta_hotel.bind("<Return>", lambda _e: self._renombrar_hotel())
        self.etiqueta_hotel.bind("<FocusOut>", lambda _e: self._renombrar_hotel())

        self.selector_hoteles = ctk.CTkOptionMenu(
            barra, width=170, fg_color="gray25", button_color="#333338",
            values=[h.nombre for h in self.hoteles],
            command=self._seleccionar_hotel_por_nombre)
        self.selector_hoteles.set(self.hotel.nombre)
        self.selector_hoteles.pack(side="left", padx=4)
        ctk.CTkButton(barra, text="＋ Hotel", width=74, height=30,
                      command=self.agregar_hotel).pack(side="left", padx=2)
        ctk.CTkButton(barra, text="🗑", width=32, height=30, fg_color="gray25",
                      hover_color="#7f1d1d",
                      command=self.eliminar_hotel).pack(side="left", padx=2)
        self.etiqueta_stats = ctk.CTkLabel(barra, text="", font=ctk.CTkFont(size=13),
                                           text_color="#9ca3af")
        self.etiqueta_stats.pack(side="right", padx=16)

        fila_rango = ctk.CTkFrame(barra, fg_color="transparent")
        fila_rango.pack(side="right", padx=(0, 12))
        ctk.CTkLabel(fila_rango, text="Consumo del", font=ctk.CTkFont(size=13),
                     text_color="#9ca3af").pack(side="left")
        self.entrada_vista_desde = ctk.CTkEntry(fila_rango, width=100,
                                                placeholder_text="dd/mm/aaaa")
        self.entrada_vista_desde.pack(side="left", padx=4)
        ctk.CTkLabel(fila_rango, text="al", font=ctk.CTkFont(size=13),
                     text_color="#9ca3af").pack(side="left")
        self.entrada_vista_hasta = ctk.CTkEntry(fila_rango, width=100,
                                                placeholder_text="dd/mm/aaaa")
        self.entrada_vista_hasta.pack(side="left", padx=4)
        ctk.CTkButton(fila_rango, text="Aplicar", width=70, height=26,
                      command=self._aplicar_rango_vista).pack(side="left", padx=4)
        ctk.CTkButton(fila_rango, text="✕", width=30, height=26,
                      fg_color="gray25", hover_color="gray35",
                      command=self._limpiar_rango_vista).pack(side="left")

        self.etiqueta_consumo = ctk.CTkLabel(barra, text="", font=ctk.CTkFont(size=11),
                                             text_color="#9ca3af", anchor="w")
        self.etiqueta_consumo.pack(fill="x", padx=16, pady=(0, 4))

    def _construir_lateral(self):
        lateral = ctk.CTkScrollableFrame(self, width=296, corner_radius=0, fg_color="#232326")
        lateral.grid(row=1, column=0, sticky="nsw", padx=0, pady=0)
        ctk.CTkButton(lateral, text="🎬  Cargar demo completa", height=38,
                      command=self.cargar_demo).pack(fill="x", padx=10, pady=(10, 4))
        ctk.CTkButton(lateral, text="👥  Nuevo grupo…", height=34, fg_color="gray25",
                      hover_color="gray35", command=self.nuevo_grupo).pack(fill="x", padx=10, pady=4)

        seccion = ctk.CTkLabel(lateral, text="GRUPOS CARGADOS", font=ctk.CTkFont(
            size=12, weight="bold"), text_color="#9ca3af", anchor="w")
        seccion.pack(fill="x", padx=10, pady=(12, 2))
        self.contenedor_grupos = ctk.CTkFrame(lateral, fg_color="transparent")
        self.contenedor_grupos.pack(fill="x", padx=4)
        fila_edicion = ctk.CTkFrame(lateral, fg_color="transparent")
        fila_edicion.pack(fill="x", padx=10, pady=(4, 0))
        ctk.CTkButton(fila_edicion, text="✎ Editar", width=70, height=30,
                      fg_color="gray25", hover_color="gray35",
                      command=self.editar_grupo_seleccionado).pack(side="left",
                                                                   expand=True, fill="x", padx=(0, 3))
        ctk.CTkButton(fila_edicion, text="🗑 Eliminar", width=90, height=30,
                      fg_color="gray25", hover_color="gray35",
                      command=self.eliminar_grupo_seleccionado).pack(side="right",
                                                                     expand=True, fill="x", padx=(3, 0))
        fila_confirmar = ctk.CTkFrame(lateral, fg_color="transparent")
        fila_confirmar.pack(fill="x", padx=10, pady=(4, 0))
        self.boton_confirmar = ctk.CTkButton(fila_confirmar, text="✔ Confirmar grupo", width=110,
                                             height=30, fg_color="#14532d", hover_color="#166534",
                                             command=self.confirmar_grupo_seleccionado)
        self.boton_confirmar.pack(side="left", expand=True, fill="x", padx=(0, 3))
        self.boton_liberar_conf = ctk.CTkButton(fila_confirmar, text="↩ Quitar confirm.", width=110,
                                                height=30, fg_color="gray25", hover_color="gray35",
                                                command=self.liberar_confirmacion_grupo_seleccionado)
        self.boton_liberar_conf.pack(side="right", expand=True, fill="x", padx=(3, 0))

        separador1 = ctk.CTkFrame(lateral, height=2, fg_color="#333338")
        separador1.pack(fill="x", padx=10, pady=12)
        self.boton_simular = ctk.CTkButton(
            lateral, text="▶  SIMULAR DISTRIBUCIÓN", height=46,
            font=ctk.CTkFont(size=15, weight="bold"),
            command=self.simular_distribucion,
        )
        self.boton_simular.pack(fill="x", padx=10, pady=4)
        self.modo_simulacion = ctk.CTkSegmentedButton(
            lateral, values=["CONSERVAR", "DESDE CERO"],
            command=lambda _m: self._cambiar_modo_simulacion(),
        )
        self.modo_simulacion.set("CONSERVAR")
        self.modo_simulacion.pack(fill="x", padx=10, pady=4)
        ctk.CTkLabel(
            lateral,
            text="CONSERVAR: respeta los grupos confirmados y acomoda los nuevos en lo libre.\n"
                 "DESDE CERO: reoptimiza todo (ignora confirmaciones).",
            text_color="#9ca3af", justify="left", font=ctk.CTkFont(size=11),
        ).pack(fill="x", padx=10, pady=(0, 4))
        fila_estrategia = ctk.CTkFrame(lateral, fg_color="transparent")
        fila_estrategia.pack(fill="x", padx=10, pady=4)
        ctk.CTkLabel(fila_estrategia, text="Estrategia:", text_color="#9ca3af").pack(side="left")
        self.menu_estrategia = ctk.CTkOptionMenu(
            fila_estrategia,
            values=["AUTO", "EXACTA", "BFD", "FFD", "MULTISTART", "TODAS"],
            width=130,
        )
        self.menu_estrategia.set("AUTO")
        self.menu_estrategia.pack(side="right")
        fila_tope = ctk.CTkFrame(lateral, fg_color="transparent")
        fila_tope.pack(fill="x", padx=10, pady=(2, 4))
        ctk.CTkLabel(fila_tope, text="Máx plazas por habitación:",
                     text_color="#9ca3af").pack(side="left")
        self.entrada_tope_plazas = ctk.CTkEntry(fila_tope, width=70,
                                                placeholder_text="sin tope")
        self.entrada_tope_plazas.pack(side="right")

        separador2 = ctk.CTkFrame(lateral, height=2, fg_color="#333338")
        separador2.pack(fill="x", padx=10, pady=12)
        self.boton_reservar = ctk.CTkButton(
            lateral, text="✅  RESERVAR este resultado", height=38,
            state="disabled", command=self.reservar_resultado,
        )
        self.boton_reservar.pack(fill="x", padx=10, pady=4)
        self.boton_liberar = ctk.CTkButton(
            lateral, text="↩  Liberar reservas del resultado", height=32,
            fg_color="gray25", hover_color="gray35",
            state="disabled", command=self.liberar_resultado,
        )
        self.boton_liberar.pack(fill="x", padx=10, pady=4)

        separador3 = ctk.CTkFrame(lateral, height=2, fg_color="#333338")
        separador3.pack(fill="x", padx=10, pady=12)
        ctk.CTkButton(lateral, text="🛏  Agregar habitaciones…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.agregar_habitaciones).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="🛏  Cargar camas…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.cargar_camas_masivo).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="🧮  Rack / Cotizador…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.abrir_rack).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="📐  Planificar / Estimar…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.abrir_planificador).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="📤  Exportar Excel…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.exportar_csv_ui).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="📋  Exportar Rooming…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.exportar_rooming_ui).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="💾  Guardar JSON…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.guardar_json).pack(fill="x", padx=10, pady=4)
        ctk.CTkButton(lateral, text="📂  Abrir JSON…", height=34,
                      fg_color="gray25", hover_color="gray35",
                      command=self.abrir_json).pack(fill="x", padx=10, pady=4)

    def _construir_zona_principal(self):
        self.tabs = ctk.CTkTabview(self, anchor="nw", corner_radius=8)
        self.tabs.grid(row=1, column=1, sticky="nsew", padx=(6, 10), pady=6)
        self.tabs.add("🗺  Mapa del hotel")
        self.tabs.add("🛏  Mapa de camas")
        self.tabs.add("🗓  Matriz grupos × noches")
        self.tabs.add("📊  Reporte")
        self._construir_pestaña_mapa(self.tabs.tab("🗺  Mapa del hotel"))
        self._construir_pestaña_camas(self.tabs.tab("🛏  Mapa de camas"))
        self._construir_pestaña_matriz(self.tabs.tab("🗓  Matriz grupos × noches"))
        self._construir_pestaña_reporte(self.tabs.tab("📊  Reporte"))

    def _construir_pestaña_mapa(self, padre):
        padre.grid_rowconfigure(1, weight=1)
        padre.grid_columnconfigure(0, weight=1)
        self.leyenda = ctk.CTkFrame(padre, fg_color="transparent")
        self.leyenda.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 0))

        zona = ctk.CTkFrame(padre, fg_color="transparent")
        zona.grid(row=1, column=0, sticky="nsew", pady=4)
        zona.grid_rowconfigure(0, weight=1)
        zona.grid_columnconfigure(0, weight=1)

        self.canvas = tk.Canvas(zona, bg=FONDO_MAPA, highlightthickness=0)
        scroll_y = ttk.Scrollbar(zona, orient="vertical", command=self.canvas.yview)
        self.canvas.configure(yscrollcommand=scroll_y.set)
        self.canvas.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        self.canvas.bind("<MouseWheel>", self._rueda)
        self.canvas.bind("<Button-1>", self._click_mapa)
        self.canvas.bind("<Button-3>", self._menu_contextual_mapa)
        self.canvas.bind("<Configure>", lambda _e: self._redibujar_mapa())

        panel = ctk.CTkFrame(padre, width=330, corner_radius=10, fg_color="#202023")
        panel.grid(row=1, column=1, sticky="nsw", padx=(6, 0), pady=4)
        panel.grid_propagate(False)
        panel.grid_rowconfigure(2, weight=1)
        panel.grid_columnconfigure(0, weight=1)
        self.etiqueta_detalle = ctk.CTkLabel(panel, text="Habitación",
                                             font=ctk.CTkFont(size=17, weight="bold"), anchor="w")
        self.etiqueta_detalle.grid(row=0, column=0, sticky="ew", padx=12, pady=(12, 2))
        self.etiqueta_subdetalle = ctk.CTkLabel(panel, text="Hacé clic en una habitación del mapa",
                                                font=ctk.CTkFont(size=12), text_color="#9ca3af",
                                                anchor="w", justify="left")
        self.etiqueta_subdetalle.grid(row=1, column=0, sticky="ew", padx=12)
        self.texto_detalle = ctk.CTkTextbox(panel, font=ctk.CTkFont(size=13), wrap="word",
                                            fg_color="#1a1a1d")
        self.texto_detalle.grid(row=2, column=0, sticky="nsew", padx=12, pady=8)
        self.texto_detalle.configure(state="disabled")
        self.boton_editar_hab = ctk.CTkButton(panel, text="✎  Editar camas…", height=30,
                                              fg_color="gray25", hover_color="gray35",
                                              state="disabled",
                                              command=self._editar_seleccionada)
        self.boton_editar_hab.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 6))
        ctk.CTkLabel(panel, text="BLOQUEOS DE ESTA HABITACIÓN", font=ctk.CTkFont(
            size=11, weight="bold"), text_color="#9ca3af", anchor="w").grid(
            row=4, column=0, sticky="ew", padx=12)
        self.contenedor_bloqueos = ctk.CTkScrollableFrame(panel, height=140, fg_color="#1a1a1d")
        self.contenedor_bloqueos.grid(row=5, column=0, sticky="ew", padx=12, pady=(4, 12))

    def _construir_pestaña_reporte(self, padre):
        padre.grid_rowconfigure(1, weight=1)
        padre.grid_columnconfigure(0, weight=1)
        tarjetas = ctk.CTkFrame(padre, fg_color="transparent")
        tarjetas.grid(row=0, column=0, sticky="ew", padx=6, pady=6)
        self.tarjetas_resumen = {}
        definiciones = [("alojados", "Pasajeros alojados"), ("habitaciones", "Habitaciones usadas"),
                        ("camas", "Camas ocupadas"), ("pico", "Pico por noche"), ("faltantes", "Sin habitación")]
        for clave, titulo in definiciones:
            tarjeta = ctk.CTkFrame(tarjetas, corner_radius=10, fg_color="#242427")
            tarjeta.pack(side="left", expand=True, fill="both", padx=5, pady=2)
            valor = ctk.CTkLabel(tarjeta, text="—", font=ctk.CTkFont(size=22, weight="bold"))
            valor.pack(padx=18, pady=(10, 0))
            ctk.CTkLabel(tarjeta, text=titulo, font=ctk.CTkFont(size=11),
                         text_color="#9ca3af").pack(padx=18, pady=(0, 10))
            self.tarjetas_resumen[clave] = valor
        self.texto_reporte = ctk.CTkTextbox(padre, font=ctk.CTkFont(family="Consolas", size=13),
                                            wrap="none", fg_color="#141416")
        self.texto_reporte.grid(row=1, column=0, sticky="nsew", padx=6, pady=(0, 6))
        self.texto_reporte.insert("1.0",
                                  "Todavía no ejecutaste ninguna simulación.\n"
                                  "Cargá la demo o creá grupos y presioná «Simular distribución».")
        self.texto_reporte.configure(state="disabled")

    def _construir_barra_estado(self):
        self.barra_estado = ctk.CTkLabel(self, text="Bienvenida/o. Cargá la demo o armá tu hotel.",
                                         anchor="w", font=ctk.CTkFont(size=12),
                                         fg_color="#202023", corner_radius=0, text_color="#9ca3af")
        self.barra_estado.grid(row=2, column=0, columnspan=2, sticky="ew")

    def report_callback_exception(self, exc, val, tb):
        detalle = "".join(traceback.format_exception(exc, val, tb))
        try:
            with open(ruta_log(), "a", encoding="utf-8") as archivo:
                archivo.write(detalle + "\n")
        except OSError:
            pass
        self._avisos_error += 1
        if self._avisos_error <= 3:
            messagebox.showerror("Error interno",
                                 "Ocurrió un error en la operación.\n"
                                 "El detalle quedó guardado en error_interfaz.log\n\n"
                                 f"{val}")

    def _status(self, mensaje: str, ok: bool = True):
        color = "#4ade80" if ok else "#f87171"
        self.barra_estado.configure(text=mensaje, text_color=color)

    def _rueda(self, evento):
        evento.widget.yview_scroll(-int(evento.delta / 120), "units")

    def _refrescar_todo(self):
        self.etiqueta_hotel.delete(0, "end")
        self.etiqueta_hotel.insert(0, self.hotel.nombre)
        total_pax = sum(g.total_pax() for g in self.grupos)
        self.etiqueta_stats.configure(
            text=f"{len(self.hotel.habitaciones)} habitaciones · {total_pax} pasajeros "
                 f"en {len(self.grupos)} grupos"
        )
        self._refrescar_tarjetas_grupos()
        self._refrescar_leyenda()
        self._redibujar_mapa()
        self._redibujar_mapa_camas()
        self._redibujar_matriz()
        self._actualizar_consumo()
        self._actualizar_detalle()
        self._programar_autoguardado()

    def _refrescar_tarjetas_grupos(self):
        for tarjeta in self.contenedor_grupos.winfo_children():
            tarjeta.destroy()
        self.tarjetas_grupos.clear()
        if not self.grupos:
            vacio = ctk.CTkLabel(self.contenedor_grupos,
                                 text="(sin grupos — usá «Nuevo grupo»\no cargá la demo)",
                                 text_color="#6b7280", justify="left")
            vacio.pack(anchor="w", padx=8, pady=6)
            return
        ordenadas = sorted(range(len(self.grupos)), key=lambda i: (
            bool(self.grupos[i].tiene_fechas)),
            )
        con_fechas = [i for i in ordenadas if self.grupos[i].tiene_fechas]
        sin_fechas = [i for i in ordenadas if not self.grupos[i].tiene_fechas]
        con_fechas.sort(key=lambda i: (self.grupos[i].fecha_ingreso, self.grupos[i].nombre.lower()))
        sin_fechas.sort(key=lambda i: self.grupos[i].nombre.lower())
        for posicion, indice in enumerate(con_fechas + sin_fechas, start=1):
            g = self.grupos[indice]
            empresa = g.empresa or g.nombre
            marcadores = []
            if g.provisional:
                marcadores.append("ESTIMADO")
            if g.confirmado:
                marcadores.append("✔ CONFIRMADO")
            if not g.tiene_fechas:
                marcadores.append("SIN FECHA")
            etiquetas = " · ".join(marcadores)
            if g.tiene_fechas:
                noches = g.noches()
                fechas = f"{formatear(g.fecha_ingreso)} → {formatear(g.fecha_egreso)} · {noches} noche(s)"
            else:
                fechas = "sin fechas definidas (cola de espera)"
            texto = (f"[{posicion}] {g.nombre}\n"
                     f"{fechas} · {g.total_pax()} pax\n"
                     f"{empresa}"
                     + (f"\n{etiquetas}" if etiquetas else ""))
            color_borde = PALETA_GRUPOS[(indice) % len(PALETA_GRUPOS)]
            if not g.tiene_fechas:
                color_borde = "#6b7280"
            tarjeta = ctk.CTkButton(
                self.contenedor_grupos, text=texto, anchor="w", height=74,
                font=ctk.CTkFont(size=12), corner_radius=8,
                fg_color="#2c2c31" if indice != self.indice_grupo_seleccionado else "#3b3b44",
                hover_color="#37373d",
                border_width=2,
                border_color=color_borde,
                command=lambda idx=indice: self._elegir_grupo(idx),
            )
            tarjeta.pack(fill="x", padx=6, pady=3)
            self.tarjetas_grupos.append(tarjeta)

    def _elegir_grupo(self, indice: int):
        self.indice_grupo_seleccionado = indice
        self._refrescar_tarjetas_grupos()
        g = self.grupos[indice]
        extras = []
        cfg = g.config
        if cfg.estudiantes_por_habitacion:
            extra_estudiantes = f"cupo estudiantes {cfg.estudiantes_por_habitacion}/hab"
            if cfg.estudiantes_estricto:
                extra_estudiantes += " (estricto)"
            extras.append(extra_estudiantes)
        if cfg.max_ninos_por_habitacion is not None:
            extras.append(f"máx {cfg.max_ninos_por_habitacion} niños/hab")
        resumen_reglas = ", ".join(extras) if extras else "reglas estándar"
        self._status(f"Grupo '{g.nombre}' seleccionado — {resumen_reglas}. "
                     f"Doble uso: botones Editar/Eliminar operan sobre esta selección.", ok=True)

    def confirmar_grupo_seleccionado(self):
        if self.indice_grupo_seleccionado is None:
            self._status("Primero seleccioná un grupo en la lista.", ok=False)
            return
        g = self.grupos[self.indice_grupo_seleccionado]
        if not g.tiene_fechas:
            self._status("No podés confirmar un grupo sin fechas. Definile el ingreso/egreso primero.",
                         ok=False)
            return
        tem = self.reporte
        g.confirmado = True
        g.provisional = False
        if tem is not None:
            reom = distribuir(self.hotel, list(self.grupos), OpcionesMotor(
                estrategia=self.menu_estrategia.get()), previo=tem,
                reoptimizar=self.modo_simulacion.get() != "CONSERVAR")
            self.reporte = reom
        else:
            self.reporte = None
            self._bloquear_acciones_resultado()
        self._refrescar_tarjetas_grupos()
        self._programar_autoguardado()
        self._status(f"Grupo '{g.nombre}' confirmado: sus habitaciones quedarán congeladas "
                     f"en las próximas simulaciones (Conservar).", ok=True)

    def liberar_confirmacion_grupo_seleccionado(self):
        if self.indice_grupo_seleccionado is None:
            self._status("Primero seleccioná un grupo en la lista.", ok=False)
            return
        g = self.grupos[self.indice_grupo_seleccionado]
        g.confirmado = False
        self.reporte = None
        self._bloquear_acciones_resultado()
        self._refrescar_tarjetas_grupos()
        self._programar_autoguardado()
        self._status(f"Grupo '{g.nombre}' desconfirmado: volverá a distribuirse libremente.", ok=True)

    def _cambiar_modo_simulacion(self):
        self.reporte = None
        self._bloquear_acciones_resultado()
        self._status("Modo de simulación cambiado. Presioná «SIMULAR DISTRIBUCIÓN» para recalcular.",
                     ok=True)

    def editar_grupo_seleccionado(self):
        if self.indice_grupo_seleccionado is None:
            self._status("Primero seleccioná un grupo en la lista.", ok=False)
            return
        self.editar_grupo(self.indice_grupo_seleccionado)

    def eliminar_grupo_seleccionado(self):
        if self.indice_grupo_seleccionado is None:
            self._status("Primero seleccioná un grupo en la lista.", ok=False)
            return
        self.eliminar_grupo(self.indice_grupo_seleccionado)

    def editar_grupo(self, indice: int):
        dialogo = DialogoGrupo(self, grupo=self.grupos[indice])
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.grupos[indice] = dialogo.resultado
            self.reporte = None
            self._bloquear_acciones_resultado()
            self._refrescar_todo()
            self._status(f"Grupo '{dialogo.resultado.nombre}' actualizado.")

    def eliminar_grupo(self, indice: int):
        g = self.grupos[indice]
        if not messagebox.askyesno("Eliminar grupo",
                                   f"¿Eliminar el grupo '{g.nombre}'?"):
            return
        del self.grupos[indice]
        self.indice_grupo_seleccionado = None
        self.reporte = None
        self._bloquear_acciones_resultado()
        self._refrescar_todo()
        self._status(f"Grupo '{g.nombre}' eliminado.")

    def nuevo_grupo(self):
        dialogo = DialogoGrupo(self)
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.grupos.append(dialogo.resultado)
            self.reporte = None
            self._bloquear_acciones_resultado()
            self._refrescar_todo()
            self._status(f"Grupo '{dialogo.resultado.nombre}' agregado.")

    def cargar_demo(self):
        from demo import construir_datos_demo
        hotel_demo, grupos_demo = construir_datos_demo()
        self._guardar_estado_hotel()
        self.hoteles.append(hotel_demo)
        self.grupos_por_hotel.append(grupos_demo)
        self.reporte_por_hotel.append(None)
        self.seleccion_por_hotel.append(None)
        self.indice_hotel = len(self.hoteles) - 1
        self._restaurar_estado_hotel()
        self.ruta_actual = None
        self.indice_grupo_seleccionado = None
        self.selector_hoteles.configure(values=[h.nombre for h in self.hoteles])
        self.selector_hoteles.set(hotel_demo.nombre)
        self._bloquear_acciones_resultado()
        self._refrescar_todo()
        self._status("Demo cargada: nuevo hotel «Cabanas del Sol» con 3 grupos. "
                     "Presioná «Simular».")

    def _capacidad_sugerida(self) -> int:
        tope = self._tope_plazas()
        if tope:
            return tope
        conteo: dict[int, int] = {}
        for h in self.hotel.habitaciones:
            conteo[h.capacidad] = conteo.get(h.capacidad, 0) + 1
        if not conteo:
            return 4
        mas_comun = max(conteo.values())
        return min(cap for cap, n in conteo.items() if n == mas_comun)

    def abrir_rack(self):
        DialogoRack(self, self.hotel)

    def abrir_planificador(self):
        dialogo = DialogoPlanificador(self)
        self.wait_window(dialogo)
        if not dialogo.aplicado:
            return
        self._refrescar_todo()
        self._status("Planificación estimada cargada del planificador.", ok=True)

    def agregar_habitaciones(self):
        sugerida = self._capacidad_sugerida()
        tope = self._tope_plazas()
        if tope:
            explicacion = f"Si dejás vacío se usan {sugerida} plazas (tu tope global)."
        elif self.hotel.habitaciones:
            explicacion = (f"Si dejás vacío se usan {sugerida} plazas "
                           "(la capacidad más común de tu hotel).")
        else:
            explicacion = f"Si dejás vacío se usan {sugerida} plazas (hotel vacío)."
        dialogo = DialogoHabitaciones(self, {h.numero for h in self.hotel.habitaciones},
                                      capacidad_sugerida=sugerida, explicacion=explicacion)
        self.wait_window(dialogo)
        if not dialogo.resultado:
            return
        existentes = {h.numero for h in self.hotel.habitaciones}
        repetidas = [h.numero for h in dialogo.resultado if h.numero in existentes]
        if repetidas:
            messagebox.showerror("Números duplicados",
                                 "Ya existen habitaciones con estos números: " + ", ".join(repetidas))
            return
        self.hotel.habitaciones.extend(dialogo.resultado)
        self._refrescar_todo()
        creadas = dialogo.resultado
        rango = (creadas[0].numero if len(creadas) == 1
                 else f"{creadas[0].numero} a {creadas[-1].numero}")
        mensaje = (f"Agregadas {len(creadas)} habitaciones ({rango}) "
                   f"con capacidad {creadas[0].capacidad}")
        if dialogo.capacidad_sugerida_usada:
            mensaje += (f" (valor sugerido: campo vacío — ajustá «Máx plazas por "
                        f"habitación» si querés otro)")
        self._status(mensaje + ".")

    def _tope_plazas(self) -> int | None:
        crudo = self.entrada_tope_plazas.get().strip()
        if not crudo:
            return None
        try:
            valor = int(crudo)
        except ValueError:
            return None
        return valor if valor >= 1 else None

    def simular_distribucion(self):
        if not self.hotel.habitaciones and not self.hotel.lotes_camas:
            messagebox.showwarning(
                "Hotel vacío",
                "Para simular necesitás habitaciones o camas:\n"
                "• «🛏 Agregar habitaciones…» para apartamentos\n"
                "• «🛏 Cargar camas…» para el inventario de camas\n"
                "• o cargá la demo completa.")
            return
        if not self.grupos:
            messagebox.showwarning("Sin grupos", "Necesitás al menos un grupo para distribuir.")
            return
        if not self.hotel.habitaciones:
            self._simular_solo_camas()
            return
        crudo_tope = self.entrada_tope_plazas.get().strip()
        tope = None
        if crudo_tope:
            try:
                tope = int(crudo_tope)
            except ValueError:
                messagebox.showerror("Tope inválido",
                                     "El máximo de plazas por habitación debe ser un número entero.")
                return
            if tope < 1:
                messagebox.showerror("Tope inválido", "El tope de plazas debe ser al menos 1.")
                return
        capacidades = {}
        if tope:
            capacidades = {h.capacidad: tope for h in self.hotel.habitaciones
                           if h.capacidad > tope}
        opciones = OpcionesMotor(estrategia=self.menu_estrategia.get(),
                                 capacidades_efectivas=capacidades)
        reoptimizar = self.modo_simulacion.get() != "CONSERVAR"
        try:
            reporte = distribuir(self.hotel, list(self.grupos), opciones,
                                 previo=self.reporte, reoptimizar=reoptimizar)
        except ErrorDistribucion as exc:
            messagebox.showerror("No se pudo distribuir", str(exc))
            return
        self.reporte = reporte
        self.numero_seleccionado = None
        self._refrescar_todo()
        self._mostrar_reporte()
        total_alojados = sum(d.pax_alojados for d in reporte.distribuciones)
        total_pax = sum(x.grupo.total_pax() for x in reporte.distribuciones)
        hab_usadas = sum(d.habitaciones_usadas for d in reporte.distribuciones)
        if reporte.todo_alojado:
            self._status(f"Distribución COMPLETA: {total_alojados}/{total_pax} pasajeros en "
                         f"{hab_usadas} habitaciones. Revisá el mapa; si te conforma, RESERVÁ.", ok=True)
            self.tabs.set("🗺  Mapa del hotel")
        else:
            faltantes = total_pax - total_alojados
            self._status(f"Atención: {faltantes} pasajero(s) quedaron SIN habitación "
                         f"(hotel lleno en esas fechas). Mirá el reporte y el déficit sugerido.", ok=False)
            self.tabs.set("📊  Reporte")
        self.boton_reservar.configure(state="normal")
        self.boton_liberar.configure(state="normal")

    def _bloquear_acciones_resultado(self):
        self.boton_reservar.configure(state="disabled")
        self.boton_liberar.configure(state="disabled")

    def _mostrar_reporte(self):
        if self.reporte is None:
            return
        r = self.reporte
        alojados = sum(d.pax_alojados for d in r.distribuciones)
        total = sum(x.grupo.total_pax() for x in r.distribuciones)
        hab_usadas = sum(d.habitaciones_usadas for d in r.distribuciones)
        camas = sum(a.carga for d in r.distribuciones for a in d.asignaciones)
        pico = max(r.pax_por_noche.values()) if r.pax_por_noche else 0
        self.tarjetas_resumen["alojados"].configure(text=f"{alojados}/{total}")
        self.tarjetas_resumen["habitaciones"].configure(text=str(hab_usadas))
        self.tarjetas_resumen["camas"].configure(text=str(camas))
        self.tarjetas_resumen["pico"].configure(text=str(pico))
        self.tarjetas_resumen["faltantes"].configure(
            text=str(total - alojados),
            text_color="#ef4444" if total > alojados else "#4ade80",
        )
        self.texto_reporte.configure(state="normal")
        self.texto_reporte.delete("1.0", "end")
        self.texto_reporte.insert("1.0", renderizar_reporte(r))
        self.texto_reporte.configure(state="disabled")

    def reservar_resultado(self):
        if self.reporte is None:
            return
        exitosas = 0
        fallos = []
        for dist in self.reporte.distribuciones:
            for asig in dist.asignaciones:
                try:
                    asig.habitacion.agregar_bloqueo(
                        dist.grupo.fecha_ingreso, dist.grupo.fecha_egreso,
                        titulo=dist.grupo.nombre,
                    )
                    exitosas += 1
                except ErrorDistribucion as exc:
                    fallos.append(str(exc))
        self._refrescar_todo()
        mensaje = f"Reservadas {exitosas} habitaciones."
        if fallos:
            mensaje += f" Avisos: {'; '.join(fallos[:3])}"
        self._status(mensaje + " No olvides Guardar JSON para persistirlas.", ok=not fallos)
        if fallos:
            messagebox.showwarning("Reservas con avisos", "\n".join(fallos[:8]))

    def liberar_resultado(self):
        if self.reporte is None:
            return
        liberadas = 0
        for dist in self.reporte.distribuciones:
            for asig in dist.asignaciones:
                if asig.habitacion.quitar_bloqueo(
                    dist.grupo.fecha_ingreso, dist.grupo.fecha_egreso,
                    titulo=dist.grupo.nombre,
                ):
                    liberadas += 1
        self._refrescar_todo()
        self._status(f"Liberadas {liberadas} habitaciones.")

    def exportar_csv_ui(self):
        if self.reporte is None:
            messagebox.showinfo("Sin datos", "Primero simulá una distribución para poder exportarla.")
            return
        ruta = filedialog.asksaveasfilename(
            defaultextension=".xlsx",
            initialfile="distribucion.xlsx",
            filetypes=[("Excel con mapa", "*.xlsx"), ("CSV (dos archivos)", "*.csv")],
        )
        if not ruta:
            return
        if ruta.lower().endswith(".csv"):
            base = ruta[:-4]
            filas_asig = exportar_csv(self.reporte, f"{base}_asignaciones.csv")
            filas_resumen = exportar_resumen_csv(self.reporte, f"{base}_resumen.csv")
            self._status(f"CSV exportados: {filas_asig} filas de asignaciones y "
                         f"{filas_resumen} grupos en el resumen.")
            return
        if not ruta.lower().endswith(".xlsx"):
            ruta += ".xlsx"
        try:
            filas = exportar_excel(self.reporte, ruta)
        except ErrorDistribucion as exc:
            messagebox.showerror("No se pudo exportar", str(exc))
            return
        self._status(f"Excel exportado: mapa del hotel + resumen + {filas} pasajeros.")

    def exportar_rooming_ui(self):
        hoteles_con_grupos = [(h, gs, self._reporte_de(h))
                              for h, gs in zip(self.hoteles, self.grupos_por_hotel)
                              if gs]
        if not hoteles_con_grupos:
            messagebox.showinfo("Sin datos", "No hay hoteles con grupos para exportar.")
            return
        if self.reporte is None and any(h.habitaciones for h, _, _ in hoteles_con_grupos):
            messagebox.showinfo("Sin simular",
                                "Algunos hoteles aún no se simularon: se calculará su "
                                "distribución automáticamente para el Rooming.")
        ruta_carpeta = filedialog.askdirectory(
            initialdir=CARPETA_EXPORTACIONES,
            title="Carpeta donde guardar el Rooming (uno por hotel)",
        )
        if not ruta_carpeta:
            return
        reportes = [self._reporte_de(h) for h in self.hoteles]
        try:
            rutas = exportar_rooming_hoteles(
                ruta_carpeta, self.hoteles, self.grupos_por_hotel,
                reportes=reportes, destino="")
        except ErrorDistribucion as exc:
            messagebox.showerror("No se pudo exportar", str(exc))
            return
        if not rutas:
            messagebox.showwarning("Sin exportar", "No se generó ningún archivo (sin grupos).")
            return
        nombres = "\n".join(os.path.basename(r) for r in rutas)
        self._status(f"Rooming exportado por hotel: {len(rutas)} archivo(s) en "
                     f"{os.path.basename(ruta_carpeta)} (incluye límites por noche).")
        messagebox.showinfo("Rooming exportado",
                            f"Se guardó un Excel por cada hotel en:\n{ruta_carpeta}\n\n{nombres}")

    def _reporte_de(self, hotel: Hotel):
        for i, h in enumerate(self.hoteles):
            if h is hotel:
                return self.reporte_por_hotel[i] if i < len(self.reporte_por_hotel) else None
        return None

    def guardar_json(self):
        ruta = filedialog.asksaveasfilename(
            defaultextension=".json",
            initialfile=self.ruta_actual or "datos_hotel.json",
            filetypes=[("JSON", "*.json")],
        )
        if not ruta:
            return
        try:
            guardar_multi(ruta, self.hoteles, self.grupos_por_hotel)
        except ErrorDistribucion as exc:
            messagebox.showerror("No se pudo guardar", str(exc))
            return
        self.ruta_actual = ruta
        self._status(f"Datos guardados en {os.path.basename(ruta)} "
                     f"({len(self.hoteles)} hoteles).")

    def abrir_json(self):
        ruta = filedialog.askopenfilename(filetypes=[("JSON", "*.json")])
        if not ruta:
            return
        try:
            hoteles, grupos_por_hotel = cargar_multi(ruta)
        except ErrorDistribucion as exc:
            messagebox.showerror("Archivo inválido", str(exc))
            return
        self.hoteles, self.grupos_por_hotel = hoteles, grupos_por_hotel
        self.reporte_por_hotel = [None for _ in hoteles]
        self.seleccion_por_hotel = [None for _ in hoteles]
        self.indice_hotel = 0
        self._restaurar_estado_hotel()
        self.ruta_actual = ruta
        self.selector_hoteles.configure(values=[h.nombre for h in self.hoteles])
        self.selector_hoteles.set(self.hotel.nombre)
        self.indice_grupo_seleccionado = None
        self._bloquear_acciones_resultado()
        self._refrescar_todo()
        self._status(f"Cargado {os.path.basename(ruta)}: {len(self.hoteles)} hotel(es). "
                     f"Presioná «Simular» cuando quieras.")

    def _refrescar_leyenda(self):
        for hijo in self.leyenda.winfo_children():
            hijo.destroy()

        def chip(color_hex, texto):
            marco = ctk.CTkFrame(self.leyenda, fg_color="transparent")
            marco.pack(side="left", padx=8, pady=4)
            lienzo = tk.Canvas(marco, width=16, height=16, bg=FONDO_MAPA, highlightthickness=1,
                               highlightbackground=color_hex)
            lienzo.create_rectangle(2, 2, 15, 15, fill=color_hex, outline="")
            lienzo.pack(side="left", padx=(0, 5))
            ctk.CTkLabel(marco, text=texto, font=ctk.CTkFont(size=12)).pack(side="left")

        chip(COLOR_LIBRE, "Libre")
        chip(COLOR_BLOQUEADA, "Bloqueada / mantenimiento")
        for indice, g in enumerate(self.grupos):
            chip(PALETA_GRUPOS[indice % len(PALETA_GRUPOS)], g.nombre)

    def _rango_vista(self):
        crudo_d = self.entrada_vista_desde.get().strip()
        crudo_h = self.entrada_vista_hasta.get().strip()
        if not crudo_d or not crudo_h:
            return None
        try:
            desde = parsear_fecha(crudo_d)
            hasta = parsear_fecha(crudo_h)
        except ValueError:
            return None
        if hasta < desde:
            desde, hasta = hasta, desde
        return (desde, hasta)

    @staticmethod
    def _solapa(desde, hasta, rango):
        if rango is None:
            return True
        return desde < rango[1] and hasta > rango[0]

    def _asignaciones_visibles(self, habitacion):
        resultado = []
        if self.reporte is None:
            return resultado
        rango = self._rango_vista()
        for dist in self.reporte.distribuciones:
            grupo = dist.grupo
            if not self._solapa(grupo.fecha_ingreso, grupo.fecha_egreso, rango):
                continue
            for asig in dist.asignaciones:
                if asig.habitacion.numero == habitacion.numero:
                    resultado.append((dist, asig))
        return resultado

    def _bloqueada_visible(self, habitacion) -> bool:
        rango = self._rango_vista()
        for b in habitacion.bloqueos:
            if self._solapa(b.desde, b.hasta, rango):
                return True
        return False

    def _estado_de(self, habitacion: Habitacion):
        visibles = self._asignaciones_visibles(habitacion)
        if visibles:
            dist, asig = visibles[0]
            indice = next((i for i, g in enumerate(self.grupos)
                           if g.nombre == dist.grupo.nombre), 0)
            return ("ocupada", indice, dist, asig)
        if self._bloqueada_visible(habitacion):
            return ("bloqueada", None, None, None)
        return ("libre", None, None, None)

    def _consumo_rango(self):
        habitaciones = set()
        camas = 0
        rango = self._rango_vista()
        if rango and self.reporte is not None:
            for dist in self.reporte.distribuciones:
                grupo = dist.grupo
                if not self._solapa(grupo.fecha_ingreso, grupo.fecha_egreso, rango):
                    continue
                for asig in dist.asignaciones:
                    habitaciones.add(asig.habitacion.numero)
                    camas += asig.carga
        return len(habitaciones), camas

    def _actualizar_consumo(self):
        base = ""
        prov = self.hotel.provisional
        if prov:
            titulo = prov["titulo"]
            convierto = (f"≈{prov['aptos_noche']} aptos/noche"
                         if titulo.endswith(" pax")
                         else f"≈{prov['pax_noche']} pax/noche")
            base += (f" 📌 Cotización {titulo}: {convierto} · "
                     f"{formatear(prov['desde'])} → {formatear(prov['hasta'])}")
        rango = self._rango_vista()
        if rango is None:
            self.etiqueta_consumo.configure(
                text="Sin rango de fechas: los mapas muestran la última simulación completa."
                     + base)
            return
        uh, camas = self._consumo_rango()
        prov_rango = 0
        if prov and prov["desde"]:
            solape = max(0, (min(rango[1], prov["hasta"])
                             - max(rango[0], prov["desde"])).days)
            prov_rango = solape * int(prov["aptos_noche"])
        self.etiqueta_consumo.configure(
            text=f"Consumo {formatear(rango[0])} → {formatear(rango[1])}: "
                 f"{uh} UH y {camas} camas ocupadas"
                 + (f" + {prov_rango} aptos de cotización" if prov_rango else "") + base)

    def _aplicar_rango_vista(self):
        crudo_d = self.entrada_vista_desde.get().strip()
        crudo_h = self.entrada_vista_hasta.get().strip()
        if bool(crudo_d) != bool(crudo_h):
            messagebox.showwarning("Rango incompleto",
                                   "Completá ambas fechas (o dejá las dos vacías).")
            return
        if crudo_d:
            try:
                parsear_fecha(crudo_d)
                parsear_fecha(crudo_h)
            except ValueError:
                messagebox.showerror("Fecha inválida",
                                     "Usá el formato dd/mm/aaaa en ambas fechas.")
                return
        self._refrescar_todo()

    def _limpiar_rango_vista(self):
        self.entrada_vista_desde.delete(0, "end")
        self.entrada_vista_hasta.delete(0, "end")
        self._refrescar_todo()

    def _pax_provisional_en(self, dia: date) -> int:
        prov = self.hotel.provisional
        if prov and prov["desde"] and prov["hasta"]:
            desde = prov["desde"]
            hasta = prov["hasta"]
            if desde <= dia < hasta:
                return int(prov["pax_noche"])
        return 0

    def _banner_provisional(self, canvas, ancho, y):
        prov = self.hotel.provisional
        if not prov:
            return y
        titulo = prov["titulo"]
        convierto = (f"≈{prov['pax_noche']} pax/noche"
                     if not titulo.endswith(" pax")
                     else f"≈{prov['aptos_noche']} aptos/noche")
        canvas.create_rectangle(16, y, ancho - 16, y + 30,
                                fill="#3a2f16", outline="#fbbf24", width=1)
        canvas.create_text(
            24, y + 15, anchor="w",
            text=(f"📌 COTIZACIÓN {titulo} · {convierto} · "
                  f"{formatear(prov['desde'])} → {formatear(prov['hasta'])} · "
                  f"PROVISORIO (no está reservado)"),
            fill="#fbbf24", font=("Segoe UI", 9, "bold"))
        return y + 34

    def _cuartos_para_provisional(self) -> list[Habitacion]:
        prov = self.hotel.provisional
        if not prov:
            return []
        desde, hasta = prov["desde"], prov["hasta"]
        ocupadas = set()
        if self.reporte:
            for dist in self.reporte.distribuciones:
                entrada, salida = dist.grupo.fecha_ingreso, dist.grupo.fecha_egreso
                for asig in dist.asignaciones:
                    if entrada < hasta and desde < salida:
                        ocupadas.add(asig.habitacion.numero)
        libres = [h for h in self.hotel.habitaciones
                  if h.libre_en_rango(desde, hasta) and h.numero not in ocupadas]
        if not libres:
            return []
        objetivo = -(-prov["pax_noche"] // prov["aptos_noche"]) if prov["aptos_noche"] else 1
        ordenados = sorted(libres, key=lambda h: (abs(h.capacidad - objetivo), h.capacidad))
        cantidad = min(prov["aptos_noche"], len(ordenados))
        return ordenados[:cantidad]

    def _redibujar_mapa(self):
        canvas = getattr(self, "canvas", None)
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete("all")
        ancho = max(canvas.winfo_width(), 600)
        por_fila = max(1, (ancho - 24) // (TAM_CARD_ANCHO + GAP_X))
        pisos: dict[str, list[Habitacion]] = {}
        for h in self.hotel.habitaciones:
            pisos.setdefault(piso_de(h.numero), []).append(h)
        claves = sorted(pisos.keys())
        y = 14
        y = self._banner_provisional(canvas, ancho, y)
        for piso in claves:
            rooms = sorted(pisos[piso], key=lambda h: clave_orden_habitacion(h.numero))
            canvas.create_text(16, y + 8, anchor="w", text=f"PISO {piso}",
                               fill="#9ca3af", font=("Segoe UI", 11, "bold"))
            y += 26
            for posicion, h in enumerate(rooms):
                columna = posicion % por_fila
                if posicion and columna == 0:
                    y += TAM_CARD_ALTO + GAP_Y
                x = 16 + columna * (TAM_CARD_ANCHO + GAP_X)
                estado, indice, dist, asig = self._estado_de(h)
                tope_vista = self._tope_plazas()
                efectiva = min(h.capacidad, tope_vista) if tope_vista else h.capacidad
                texto_capacidad = (f"cap {h.capacidad}→{efectiva}"
                                   if efectiva < h.capacidad else f"cap {h.capacidad}")
                if estado == "ocupada":
                    relleno = PALETA_GRUPOS[indice % len(PALETA_GRUPOS)]
                    borde = relleno
                    linea2 = dist.grupo.nombre[:13]
                    linea3 = f"{len(asig.nombres_pax)} pax · {texto_capacidad}"
                elif estado == "bloqueada":
                    relleno, borde = COLOR_BLOQUEADA, BORDE_BLOQUEO
                    linea2 = "🔧 " + h.bloqueos[0].titulo[:11]
                    linea3 = f"{formatear(h.bloqueos[0].desde)}"
                else:
                    relleno, borde = COLOR_LIBRE, BORDE_LIBRE
                    linea2 = "Libre"
                    linea3 = texto_capacidad + (
                        "" if h.disponible_desde is None else f"  ⏳{formatear(h.disponible_desde)}")
                tags = ("room", f"room:{h.numero}")
                canvas.create_rectangle(x, y, x + TAM_CARD_ANCHO, y + TAM_CARD_ALTO,
                                        fill=relleno, outline=borde, width=2,
                                        tags=tags)
                canvas.create_text(x + TAM_CARD_ANCHO / 2, y + 15, text=h.numero,
                                   fill=TEXTO_CLARO, font=("Segoe UI", 13, "bold"), tags=tags)
                canvas.create_text(x + TAM_CARD_ANCHO / 2, y + 34, text=linea2,
                                   fill=TEXTO_CLARO, font=("Segoe UI", 10), tags=tags)
                canvas.create_text(x + TAM_CARD_ANCHO / 2, y + 50, text=linea3,
                                   fill="#d4d4d8", font=("Segoe UI", 9), tags=tags)
                if efectiva <= 4:
                    for cama in range(efectiva):
                        cx = x + 12 + cama * 22
                        canvas.create_rectangle(cx, y + TAM_CARD_ALTO - 9,
                                                cx + 16, y + TAM_CARD_ALTO - 3,
                                                fill="#26262b", outline="")
            y += TAM_CARD_ALTO + GAP_Y
        if self.numero_seleccionado:
            coords = canvas.coords(f"room:{self.numero_seleccionado}")
            if coords:
                canvas.create_rectangle(coords[0] - 3, coords[1] - 3, coords[2] + 3, coords[3] + 3,
                                        outline="#ffffff", width=3)
        for h in self._cuartos_para_provisional():
            coords = canvas.coords(f"room:{h.numero}")
            if not coords:
                continue
            x0, y0, x1, y1 = coords
            canvas.create_rectangle(x0, y0, x1, y1, outline="#fbbf24", width=3,
                                    dash=(5, 3), tags=(f"room:{h.numero}", f"prov:{h.numero}"))
            canvas.create_oval(x0 + 6, y0 + 6, x0 + 26, y0 + 26, fill="#fbbf24", outline="",
                               tags=(f"room:{h.numero}", f"prov:{h.numero}"))
            canvas.create_text(x0 + 16, y0 + 16, text="CO",
                               fill="#151515", font=("Segoe UI", 9, "bold"),
                               tags=(f"room:{h.numero}", f"prov:{h.numero}"))
        alto_total = y + 30
        canvas.configure(scrollregion=(0, 0, ancho, max(alto_total, canvas.winfo_height())))

    def _construir_pestaña_camas(self, padre):
        padre.grid_rowconfigure(1, weight=1)
        padre.grid_columnconfigure(0, weight=1)
        leyenda = ctk.CTkFrame(padre, fg_color="transparent")
        leyenda.grid(row=0, column=0, columnspan=2, sticky="ew", padx=6, pady=(4, 0))
        ctk.CTkLabel(leyenda,
                     text="Inventario de camas INDEPENDIENTE de los apartamentos. "
                          "La tabla muestra cuántas personas podés alojar cada día "
                          "(filtrá el período arriba). Clic derecho sobre un lote = quitar.",
                     font=ctk.CTkFont(size=11), text_color="#9ca3af",
                     anchor="w", justify="left").pack(anchor="w")

        zona = ctk.CTkFrame(padre, fg_color="transparent")
        zona.grid(row=1, column=0, sticky="nsew", pady=4)
        zona.grid_rowconfigure(0, weight=1)
        zona.grid_columnconfigure(0, weight=1)

        self.canvas_camas = tk.Canvas(zona, bg=FONDO_MAPA, highlightthickness=0)
        scroll_y = ttk.Scrollbar(zona, orient="vertical", command=self.canvas_camas.yview)
        self.canvas_camas.configure(yscrollcommand=scroll_y.set)
        self.canvas_camas.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        self.canvas_camas.bind("<MouseWheel>", self._rueda)
        self.canvas_camas.bind("<Button-3>", self._menu_contextual_camas)
        self.canvas_camas.bind("<Configure>",
                               lambda _e: self._redibujar_mapa_camas())

        nota = ctk.CTkLabel(padre, corner_radius=10, fg_color="#202023",
                            text="Las camas NO se asignan a\napartamentos: son stock propio.\n\n"
                                 "Usá «🛏 Cargar camas…» para\nsumar lotes con su ventana\nde fechas.",
                            font=ctk.CTkFont(size=12), text_color="#9ca3af", justify="left")
        nota.grid(row=1, column=1, sticky="ns", padx=(6, 0), pady=4)

    def _construir_pestaña_matriz(self, padre):
        padre.grid_rowconfigure(1, weight=1)
        padre.grid_columnconfigure(0, weight=1)
        cabecera = ctk.CTkFrame(padre, fg_color="transparent")
        cabecera.grid(row=0, column=0, sticky="ew", padx=6, pady=(4, 0))
        self.matriz_modo = ctk.StringVar(value="Pessoas")
        selector = ctk.CTkSegmentedButton(cabecera, values=["Pessoas", "Apartamentos"],
                                          variable=self.matriz_modo,
                                          command=lambda _m: self._redibujar_matriz())
        selector.pack(side="left")
        ctk.CTkLabel(cabecera, text="El mapa de grupos por noche: misma planilla para "
                                    "personas que para apartamentos, solo cambian los números.",
                     font=ctk.CTkFont(size=11), text_color="#9ca3af").pack(
            side="left", padx=10)
        ctk.CTkButton(cabecera, text="📤 Exportar Mapa de Pessoas…", height=30,
                      fg_color="gray25", hover_color="gray35",
                      command=lambda: self._exportar_mapa_grupos_ui("Pessoas")).pack(
            side="right", padx=(4, 0))
        ctk.CTkButton(cabecera, text="📤 Exportar Mapa de Apartamentos…", height=30,
                      fg_color="gray25", hover_color="gray35",
                      command=lambda: self._exportar_mapa_grupos_ui("Apartamentos")).pack(
            side="right")

        zona = ctk.CTkFrame(padre, fg_color="transparent")
        zona.grid(row=1, column=0, sticky="nsew", pady=4)
        zona.grid_rowconfigure(0, weight=1)
        zona.grid_columnconfigure(0, weight=1)
        self.canvas_matriz = tk.Canvas(zona, bg=FONDO_MAPA, highlightthickness=0)
        scroll_y = ttk.Scrollbar(zona, orient="vertical", command=self.canvas_matriz.yview)
        scroll_x = ttk.Scrollbar(zona, orient="horizontal", command=self.canvas_matriz.xview)
        self.canvas_matriz.configure(yscrollcommand=scroll_y.set, xscrollcommand=scroll_x.set)
        self.canvas_matriz.grid(row=0, column=0, sticky="nsew")
        scroll_y.grid(row=0, column=1, sticky="ns")
        scroll_x.grid(row=1, column=0, sticky="ew")
        self.canvas_matriz.bind("<MouseWheel>", self._rueda_matriz)
        self.canvas_matriz.bind("<Shift-MouseWheel>", self._rueda_matriz_x)
        self.canvas_matriz.bind("<Configure>", lambda _e: self._redibujar_matriz())

    def _rueda_matriz(self, evento):
        self.canvas_matriz.yview_scroll(int(-evento.delta / 120), "units")

    def _rueda_matriz_x(self, evento):
        self.canvas_matriz.xview_scroll(int(-evento.delta / 120), "units")

    def _rango_matriz(self):
        rango = self._rango_vista()
        if rango is not None:
            return rango
        if self.reporte is not None and self.reporte.pax_por_noche:
            fechas = sorted(self.reporte.pax_por_noche)
            return fechas[0], fechas[-1]
        if self.grupos:
            con_fechas = [g for g in self.grupos if g.tiene_fechas]
            if con_fechas:
                inicios = [g.fecha_ingreso for g in con_fechas]
                egresos = [g.fecha_egreso for g in con_fechas]
                return min(inicios), max(egresos) - timedelta(days=1)
        return None

    def _celdas_grupo(self, g):
        camas = g.total_pax()
        aptos = None
        if self.reporte is not None:
            for dist in self.reporte.distribuciones:
                if dist.grupo is g:
                    aptos = dist.habitaciones_usadas
                    break
        return camas, aptos

    def _redibujar_matriz(self):
        canvas = getattr(self, "canvas_matriz", None)
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete("all")
        ancho = max(canvas.winfo_width(), 640)
        y = 10
        canvas.create_text(16, y + 6, anchor="w",
                           text=f"{self.hotel.nombre} · mapa de "
                                f"{self.matriz_modo.get().lower()}",
                           fill="#9ca3af", font=("Segoe UI", 11, "bold"))
        y += 22
        rango = self._rango_matriz()
        modo = self.matriz_modo.get()
        if rango is None:
            canvas.create_text(16, y + 6, anchor="w",
                               text="Sin grupos ni rango definidos. Cargá grupos y simulá.",
                               fill="#71717a", font=("Segoe UI", 10))
            canvas.configure(scrollregion=(0, 0, ancho, y + 40))
            return
        desde, hasta = rango
        dias = [desde + timedelta(days=i) for i in range((hasta - desde).days + 1)]
        if len(dias) > 92:
            dias = dias[:92]
        col_emp, col_pax, col_in, col_out = 16, 250, 300, 360
        col_dia0 = 420
        ancho_dia = 46
        col_total = col_dia0 + len(dias) * ancho_dia + 10
        fila_alto = 20
        if modo == "Pessoas":
            limites = [self.hotel.camas_inventario_en(d) if self.hotel.lotes_camas
                       else self.hotel.total_camas() for d in dias]
        else:
            limites = [len(self.hotel.habitaciones)] * len(dias)
        ocupados = [0] * len(dias)

        def encabezado(txt, x, ancho):
            canvas.create_rectangle(x, y - 2, x + ancho, y + 16, fill="#242427", outline="")
            canvas.create_text(x + 6, y + 7, anchor="w", text=txt,
                               fill="#9ca3af", font=("Segoe UI", 9, "bold"))

        canvas.create_rectangle(16, y - 2, col_total + 44, y + 16, fill="#242427", outline="")
        encabezado("GRUPO / EMPRESA", col_emp, 234)
        encabezado("PAX", col_pax, 50)
        encabezado("IN", col_in, 60)
        encabezado("OUT", col_out, 60)
        for i, d in enumerate(dias):
            encabezado(f"{d.day:02d}" + ("" if i % 2 else ""), col_dia0 + i * ancho_dia, ancho_dia)
        encabezado("TOTAL", col_total, 44)
        y += 18

        canvas.create_text(16, y + 7, anchor="w",
                           text="Cada noches = columna de abajo; fuera de la estadía queda vacía.",
                           fill="#71717a", font=("Segoe UI", 8))
        y += 14
        for g in self.grupos:
            if g.total_pax() < 1:
                continue
            if not g.tiene_fechas:
                continue
            camas, aptos = self._celdas_grupo(g)
            canvas.create_text(col_emp, y + 8, anchor="w", text=g.nombre[:40],
                               fill="#d4d4d8", font=("Segoe UI", 9))
            canvas.create_text(col_pax, y + 8, anchor="w", text=str(camas),
                               fill="#d4d4d8", font=("Segoe UI", 9, "bold"))
            canvas.create_text(col_in, y + 8, anchor="w", text=formatear(g.fecha_ingreso),
                               fill="#a78bfa", font=("Segoe UI", 9))
            canvas.create_text(col_out, y + 8, anchor="w", text=formatear(g.fecha_egreso),
                               fill="#a78bfa", font=("Segoe UI", 9))
            total = 0
            for i, d in enumerate(dias):
                if not (g.fecha_ingreso <= d < g.fecha_egreso):
                    continue
                if modo == "Pessoas":
                    valor = camas
                else:
                    if aptos is None:
                        continue
                    valor = aptos
                cell_x = col_dia0 + i * ancho_dia
                canvas.create_rectangle(cell_x, y, cell_x + ancho_dia - 4, y + 16,
                                        fill="#1c1c1f", outline="")
                canvas.create_text(cell_x + 6, y + 8, anchor="w", text=str(valor),
                                   fill="#e2e8f0", font=("Segoe UI", 9))
                ocupados[i] += valor
                total += valor
            if total:
                canvas.create_text(col_total, y + 8, anchor="w", text=str(total),
                                   fill="#fbbf24", font=("Segoe UI", 9, "bold"))
            y += fila_alto

        y += 8
        for nombre, valores in (("LÍMITE DIÁRIO", limites),
                                ("OCUPADOS", ocupados),
                                ("DISPONIBLES", [l - o for l, o in zip(limites, ocupados)])):
            canvas.create_rectangle(16, y, ancho - 16, y + 18, fill="#242427", outline="")
            canvas.create_text(col_emp, y + 9, anchor="w", text=nombre,
                               fill="#9ca3af", font=("Segoe UI", 9, "bold"))
            for i, v in enumerate(valores):
                cell_x = col_dia0 + i * ancho_dia
                color = "#f87171" if (nombre == "DISPONIBLES" and v < 0) else "#4ade80"
                if nombre == "LÍMITE DIÁRIO":
                    color = "#d4d4d8"
                canvas.create_text(cell_x + 6, y + 9, anchor="w", text=str(v),
                                   fill=color, font=("Segoe UI", 9, "bold"))
            y += 20
        y += 6
        canvas.create_text(16, y, anchor="nw",
                           text="DISPONIBLES rojo = noche por encima del límite (sobreventa). "
                                "Exportá «Mapa de Pessoas» o «Mapa de Apartamentos» para la "
                                "planilla Excel con el formato que usás hoy.",
                           fill="#71717a", font=("Segoe UI", 9))
        y += 24
        alto = y
        canvas.configure(scrollregion=(0, 0, max(col_total + 44, ancho), max(alto, canvas.winfo_height())))

    def _exportar_mapa_grupos_ui(self, modo: str):
        rango = self._rango_matriz()
        if rango is None:
            messagebox.showwarning("Sin datos", "Cargá grupos y simulá antes de exportar.",
                                   parent=self)
            return
        desde, hasta = rango
        prefijo = "mapa-pessoas" if modo == "Pessoas" else "mapa-apartamentos"
        ruta = filedialog.asksaveasfilename(
            parent=self, defaultextension=".xlsx",
            initialfile=f"{prefijo}_{self.hotel.nombre}_{desde.strftime('%d%m%Y')}-"
                        f"{hasta.strftime('%d%m%Y')}.xlsx",
            filetypes=[("Excel", "*.xlsx")])
        if not ruta:
            return
        pendientes = exportar_mapa_grupos(ruta, self.hotel, self.grupos, self.reporte,
                                          desde, hasta, modo)
        if pendientes:
            self._status(f"Excel guardado, pero {pendientes} grupo(s) sin simular: "
                         "los apartamentos saldrán al simular.")
        else:
            self._status(f"Excel '{os.path.basename(ruta)}' guardado.")

    def _camas_consumidas_en(self, dia: date) -> int:
        if self.reporte is not None:
            total = 0
            for dist in self.reporte.distribuciones:
                grupo = dist.grupo
                if grupo.fecha_ingreso <= dia < grupo.fecha_egreso:
                    total += sum(a.carga for a in dist.asignaciones)
            return total
        return sum(g.total_pax() for g in self.grupos
                   if g.tiene_fechas and g.fecha_ingreso <= dia < g.fecha_egreso)

    def _simular_solo_camas(self):
        """Distribución sin apartamentos: valida el inventario de camas por día."""
        self.reporte = None
        self.numero_seleccionado = None
        dias, _hay_mas = self._dias_tabla_camas()
        lineas = ["MODO SOLO CAMAS — no hay apartamentos cargados.",
                  "Se compara la demanda de tus grupos contra el inventario de camas por día.",
                  ""]
        deficit_dias = 0
        max_demanda = 0
        max_stock = 0
        for dia in dias:
            stock = self.hotel.camas_inventario_en(dia)
            demanda = sum(g.total_pax() for g in self.grupos
                          if g.tiene_fechas and g.fecha_ingreso <= dia < g.fecha_egreso)
            max_stock = max(max_stock, stock)
            max_demanda = max(max_demanda, demanda)
            marca = ""
            if demanda > stock:
                deficit_dias += 1
                marca = f"  ← FALTAN {demanda - stock} camas"
                lineas.append(f"{formatear(dia)}: demanda {demanda} pax / "
                              f"{stock} camas{marca}")
            elif demanda:
                lineas.append(f"{formatear(dia)}: demanda {demanda} pax / "
                              f"{stock} camas ({max(stock - demanda, 0)} libres)")
        lineas.append("")
        if deficit_dias:
            lineas.append(f"⚠ {deficit_dias} día(s) con déficit de camas para la demanda cargada.")
        else:
            lineas.append("✅ El inventario de camas cubre la demanda de todos los grupos.")
        self.texto_reporte.configure(state="normal")
        self.texto_reporte.delete("1.0", "end")
        self.texto_reporte.insert("1.0", "\n".join(lineas))
        self.texto_reporte.configure(state="disabled")
        self.tarjetas_resumen["alojados"].configure(text="—")
        self.tarjetas_resumen["habitaciones"].configure(text="0")
        self.tarjetas_resumen["camas"].configure(text=str(max_stock))
        self.tarjetas_resumen["pico"].configure(text=str(max_demanda))
        self.tarjetas_resumen["faltantes"].configure(
            text=str(deficit_dias),
            text_color="#ef4444" if deficit_dias else "#4ade80")
        self._refrescar_todo()
        self._bloquear_acciones_resultado()
        self.tabs.set("📊  Reporte")
        if deficit_dias:
            self._status(f"Solo camas: {deficit_dias} día(s) sin stock suficiente "
                         f"(máx stock {max_stock}, pico demanda {max_demanda}).", ok=False)
        else:
            self._status(f"Solo camas OK: inventario cubre la demanda "
                         f"(pico {max_demanda} pax sobre {max_stock} camas).", ok=True)

    def _dias_tabla_camas(self):
        rango = self._rango_vista()
        if rango:
            inicio, fin = rango
        elif self.reporte is not None and self.reporte.distribuciones:
            inicios = [d.grupo.fecha_ingreso for d in self.reporte.distribuciones]
            fines = [d.grupo.fecha_egreso for d in self.reporte.distribuciones]
            inicio, fin = min(inicios), max(fines)
            if fin <= inicio:
                fin = inicio + timedelta(days=1)
        else:
            inicio = date.today()
            fin = inicio + timedelta(days=14)
        maximo = 62
        dias = []
        actual = inicio
        while actual < fin and len(dias) < maximo:
            dias.append(actual)
            actual += timedelta(days=1)
        return dias, (fin > actual if len(dias) == maximo and actual < fin else False)

    def _redibujar_mapa_camas(self):
        canvas = getattr(self, "canvas_camas", None)
        if canvas is None or not canvas.winfo_exists():
            return
        canvas.delete("all")
        ancho = max(canvas.winfo_width(), 600)
        por_fila = max(1, (ancho - 24) // (TAM_CARD_ANCHO + GAP_X))
        hoy = date.today()
        y = 14
        y = self._banner_provisional(canvas, ancho, y)
        canvas.create_text(16, y + 8, anchor="w",
                           text="INVENTARIO DE CAMAS (independiente de los apartamentos)",
                           fill="#9ca3af", font=("Segoe UI", 11, "bold"))
        y += 26
        lotes = self.hotel.lotes_camas
        if not lotes:
            canvas.create_text(24, y + 10, anchor="w",
                               text="(sin camas cargadas — botón «🛏 Cargar camas…»)",
                               fill="#6b7280", font=("Segoe UI", 11))
            y += 40
        for posicion, lote in enumerate(lotes):
            columna = posicion % por_fila
            if posicion and columna == 0:
                y += TAM_CARD_ALTO + GAP_Y
            x = 16 + columna * (TAM_CARD_ANCHO + GAP_X)
            activo_hoy = lote.activo_en(hoy)
            futuro = lote.disponible_desde and lote.disponible_desde > hoy
            pasado = lote.disponible_hasta and lote.disponible_hasta < hoy
            if activo_hoy:
                relleno, borde = COLOR_LIBRE, BORDE_LIBRE
                estado_txt = "activas hoy"
            elif pasado:
                relleno, borde = COLOR_BLOQUEADA, BORDE_BLOQUEO
                estado_txt = "finalizó"
            elif futuro:
                relleno, borde = "#1f2937", "#374151"
                estado_txt = f"desde {formatear(lote.disponible_desde)}"
            else:
                relleno, borde = "#1f2937", "#374151"
                estado_txt = "sin vigencia"
            tags = ("lote", f"lote:{posicion}")
            canvas.create_rectangle(x, y, x + TAM_CARD_ANCHO, y + TAM_CARD_ALTO,
                                    fill=relleno, outline=borde, width=2, tags=tags)
            canvas.create_text(x + 8, y + 14, anchor="w", text=f"🛏 {lote.titulo[:13]}",
                               fill=TEXTO_CLARO, font=("Segoe UI", 11, "bold"), tags=tags)
            canvas.create_text(x + TAM_CARD_ANCHO / 2, y + 34, text=f"{lote.cantidad} camas",
                               fill="#d4d4d8", font=("Segoe UI", 11), tags=tags)
            ventana = "siempre"
            if lote.disponible_desde or lote.disponible_hasta:
                desde = formatear(lote.disponible_desde) if lote.disponible_desde else "inicio"
                hasta = formatear(lote.disponible_hasta) if lote.disponible_hasta else "indefinido"
                ventana = f"{desde} → {hasta}"
            canvas.create_text(x + TAM_CARD_ANCHO / 2, y + 52, text=ventana[:20],
                               fill="#a1a1aa", font=("Segoe UI", 8), tags=tags)
            canvas.create_text(x + 8, y + 2, anchor="nw", text=estado_txt,
                               fill="#67e8f9" if activo_hoy else "#94a3b8",
                               font=("Segoe UI", 7), tags=tags)
        y += TAM_CARD_ALTO + 30
        canvas.create_text(16, y + 8, anchor="w", text="PERSONAS QUE PUEDO METER POR DÍA",
                           fill="#9ca3af", font=("Segoe UI", 11, "bold"))
        y += 28
        dias, hay_mas = self._dias_tabla_camas()
        col_fecha, col_total, col_cons, col_prov, col_libres = 24, 118, 190, 268, 352
        canvas.create_text(col_fecha, y, anchor="w", text="DÍA", fill="#71717a",
                           font=("Segoe UI", 9, "bold"))
        canvas.create_text(col_total, y, anchor="w", text="CAMAS", fill="#71717a",
                           font=("Segoe UI", 9, "bold"))
        canvas.create_text(col_cons, y, anchor="w", text="CONSUMIDAS", fill="#71717a",
                           font=("Segoe UI", 9, "bold"))
        canvas.create_text(col_prov, y, anchor="w", text="+ COTIZ.", fill="#fbbf24",
                           font=("Segoe UI", 9, "bold"))
        canvas.create_text(col_libres, y, anchor="w", text="LIBRES", fill="#71717a",
                           font=("Segoe UI", 9, "bold"))
        y += 18
        for indice_dia, dia in enumerate(dias):
            total = self.hotel.camas_inventario_en(dia)
            consumidas = self._camas_consumidas_en(dia)
            prov = self._pax_provisional_en(dia)
            libres = max(total - consumidas - prov, 0)
            excedido = (consumidas + prov) > total
            if indice_dia % 2 == 0:
                canvas.create_rectangle(16, y - 2, ancho - 16, y + 18,
                                        fill="#1c1c1f", outline="", tags=f"dia:{dia.isoformat()}")
            if prov:
                canvas.create_rectangle(16, y - 2, ancho - 16, y + 18,
                                        fill="#3a2f16", outline="", tags=f"dia:{dia.isoformat()}")
            color_dia = "#d4d4d8" if dia.weekday() < 5 else "#a78bfa"
            canvas.create_text(col_fecha, y + 7, anchor="w",
                               text=f"{formatear(dia)} ({'LMXJVSD'[dia.weekday()]})",
                               fill=color_dia, font=("Segoe UI", 10),
                               tags=f"dia:{dia.isoformat()}")
            canvas.create_text(col_total, y + 7, anchor="w", text=str(total),
                               fill="#d4d4d8", font=("Segoe UI", 10),
                               tags=f"dia:{dia.isoformat()}")
            canvas.create_text(col_cons, y + 7, anchor="w", text=str(consumidas),
                               fill="#f87171" if excedido else "#d4d4d8",
                               font=("Segoe UI", 10), tags=f"dia:{dia.isoformat()}")
            canvas.create_text(col_prov, y + 7, anchor="w", text=str(prov) if prov else "—",
                               fill="#fbbf24", font=("Segoe UI", 10, "bold"),
                               tags=f"dia:{dia.isoformat()}")
            canvas.create_text(col_libres, y + 7, anchor="w", text=str(libres),
                               fill="#f87171" if excedido else "#4ade80",
                               font=("Segoe UI", 10, "bold"), tags=f"dia:{dia.isoformat()}")
            y += 21
        if hay_mas:
            canvas.create_text(24, y + 6, anchor="w",
                               text="+ más días: poné el rango en «Consumo del … al …» arriba",
                               fill="#6b7280", font=("Segoe UI", 9))
            y += 22
        canvas.configure(scrollregion=(0, 0, ancho, max(y + 30, canvas.winfo_height())))

    def _habitacion_por_numero(self, numero: str) -> Habitacion | None:
        for h in self.hotel.habitaciones:
            if h.numero == numero:
                return h
        return None

    def _numero_en(self, evento) -> str | None:
        lienzo = evento.widget
        items = lienzo.find_withtag("current")
        for item in items:
            etiquetas = lienzo.gettags(item)
            for etiqueta in etiquetas:
                if etiqueta.startswith("room:"):
                    return etiqueta[5:]
        return None

    def _redibujar_ambos(self):
        self._redibujar_mapa()
        self._redibujar_mapa_camas()

    def _click_mapa(self, evento):
        numero = self._numero_en(evento)
        if not numero:
            return
        self.numero_seleccionado = numero
        self._redibujar_ambos()
        self._actualizar_detalle()

    def _menu_contextual_mapa(self, evento):
        numero = self._numero_en(evento)
        if not numero:
            return
        self.numero_seleccionado = numero
        self._redibujar_ambos()
        self._actualizar_detalle()
        habitacion = self._habitacion_por_numero(numero)
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"🔧 Bloquear {numero}…", command=lambda: self._bloquear(habitacion))
        menu.add_command(label=f"🛏 Editar camas de {numero}…",
                         command=lambda: self._editar_habitacion(habitacion))
        if habitacion.bloqueos:
            menu.add_command(label="🧹 Quitar último bloqueo",
                             command=lambda: self._quitar_ultimo_bloqueo(habitacion))
        menu.add_separator()
        menu.add_command(label=f"🗑 Eliminar habitación {numero}",
                         command=lambda: self._eliminar_habitacion(habitacion))
        menu.tk_popup(evento.x_root, evento.y_root)

    def _bloquear(self, habitacion: Habitacion):
        dialogo = DialogoBloqueo(self, habitacion)
        self.wait_window(dialogo)
        if getattr(dialogo, "resultado", None):
            self._refrescar_todo()
            self._status(f"Habitación {habitacion.numero} bloqueada.")

    def _editar_seleccionada(self):
        if not self.numero_seleccionado:
            return
        h = self._habitacion_por_numero(self.numero_seleccionado)
        if h:
            self._editar_habitacion(h)

    def _editar_habitacion(self, habitacion: Habitacion):
        dialogo = DialogoEditarHabitacion(self, habitacion)
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.reporte = None
            self._bloquear_acciones_resultado()
            self._refrescar_todo()
            self._status(f"Habitación {habitacion.numero} ahora tiene "
                         f"{habitacion.capacidad} camas. Simulá de nuevo para recalcular.")

    def cargar_camas_masivo(self):
        dialogo = DialogoCamasLote(self)
        self.wait_window(dialogo)
        if dialogo.resultado:
            self.hotel.lotes_camas.append(dialogo.lote)
            self._refrescar_todo()
            self.tabs.set("🛏  Mapa de camas")
            self._status(f"✅ Lote '{dialogo.lote.titulo}' agregado: "
                         f"{dialogo.lote.cantidad} camas.")

    def _menu_contextual_camas(self, evento):
        lienzo = evento.widget
        indice = None
        for item in lienzo.find_withtag("current"):
            for etiqueta in lienzo.gettags(item):
                if etiqueta.startswith("lote:"):
                    indice = int(etiqueta[5:])
                    break
            if indice is not None:
                break
        if indice is None or indice >= len(self.hotel.lotes_camas):
            return
        lote = self.hotel.lotes_camas[indice]
        menu = tk.Menu(self, tearoff=0)
        menu.add_command(label=f"🗑 Quitar '{lote.titulo}' ({lote.cantidad} camas)",
                         command=lambda: self._quitar_lote_camas(indice))
        menu.tk_popup(evento.x_root, evento.y_root)

    def _quitar_lote_camas(self, indice: int):
        lote = self.hotel.lotes_camas[indice]
        if not messagebox.askyesno("Quitar lote de camas",
                                   f"¿Quitar el lote '{lote.titulo}' ({lote.cantidad} camas)?"):
            return
        self.hotel.lotes_camas.pop(indice)
        self._refrescar_todo()
        self._status(f"Lote '{lote.titulo}' quitado del inventario de camas.")

    def _quitar_ultimo_bloqueo(self, habitacion: Habitacion):
        if not habitacion.bloqueos:
            return
        ultimo = habitacion.bloqueos[-1]
        if habitacion.quitar_bloqueo(ultimo.desde, ultimo.hasta, titulo=ultimo.titulo):
            self._refrescar_todo()
            self._status(f"Bloqueo '{ultimo.titulo}' quitado de la habitación {habitacion.numero}.")

    def _eliminar_habitacion(self, habitacion: Habitacion):
        if not messagebox.askyesno("Eliminar habitación",
                                   f"¿Eliminar la habitación {habitacion.numero}?"):
            return
        self.hotel.habitaciones.remove(habitacion)
        if self.numero_seleccionado == habitacion.numero:
            self.numero_seleccionado = None
        self._refrescar_todo()
        self._status(f"Habitación {habitacion.numero} eliminada.")

    def _actualizar_detalle(self):
        if not self.numero_seleccionado:
            self.etiqueta_detalle.configure(text="Habitación")
            self.etiqueta_subdetalle.configure(text="Hacé clic en una habitación del mapa")
            self.texto_detalle.configure(state="normal")
            self.texto_detalle.delete("1.0", "end")
            self.texto_detalle.configure(state="disabled")
            if hasattr(self, "boton_editar_hab"):
                self.boton_editar_hab.configure(state="disabled")
            for hijo in self.contenedor_bloqueos.winfo_children():
                hijo.destroy()
            return
        h = self._habitacion_por_numero(self.numero_seleccionado)
        if h is None:
            self.numero_seleccionado = None
            return
        self.etiqueta_detalle.configure(
            text=f"Habitación {h.numero} · Piso {piso_de(h.numero)}"
        )
        ventana = "siempre disponible"
        if h.disponible_desde or h.disponible_hasta:
            desde = formatear(h.disponible_desde) if h.disponible_desde else "inicio"
            hasta = formatear(h.disponible_hasta) if h.disponible_hasta else "indefinido"
            ventana = f"disponible del {desde} al {hasta}"
        tope_detalle = self._tope_plazas()
        efectiva_detalle = (min(h.capacidad, tope_detalle) if tope_detalle else h.capacidad)
        texto_capacidad = f"Capacidad: {h.capacidad} plazas"
        if efectiva_detalle < h.capacidad:
            texto_capacidad += f" — con el tope activo se usan {efectiva_detalle}"
        lineas = [texto_capacidad, f"Ventana: {ventana}", ""]
        estado, _indice, dist, asig = self._estado_de(h)
        if dist is not None:
            lineas.append("── ASIGNADA EN LA ÚLTIMA SIMULACIÓN ──")
            lineas.append(f"Grupo: {dist.grupo.nombre}")
            lineas.append(f"Fechas: {formatear(dist.grupo.fecha_ingreso)} → "
                          f"{formatear(dist.grupo.fecha_egreso)}")
            lineas.append(f"Huéspedes ({len(asig.nombres_pax)}):")
            lineas += [f"   • {n}" for n in asig.nombres_pax]
            lineas.append("")
        if h.bloqueos:
            lineas.append("── HISTORIAL DE BLOQUEOS ──")
            lineas += [f"🔧 {b.titulo}: {formatear(b.desde)} → {formatear(b.hasta)}"
                       for b in h.bloqueos]
        else:
            lineas.append("(sin bloqueos)")
        self.texto_detalle.configure(state="normal")
        self.texto_detalle.delete("1.0", "end")
        self.texto_detalle.insert("1.0", "\n".join(lineas))
        self.texto_detalle.configure(state="disabled")
        self.boton_editar_hab.configure(state="normal")
        for hijo in self.contenedor_bloqueos.winfo_children():
            hijo.destroy()
        if not h.bloqueos:
            ctk.CTkLabel(self.contenedor_bloqueos, text="(ninguno)", text_color="#6b7280").pack(
                anchor="w", padx=8, pady=4)
        for b in h.bloqueos:
            fila = ctk.CTkFrame(self.contenedor_bloqueos, fg_color="transparent")
            fila.pack(fill="x", padx=6, pady=2)
            ctk.CTkLabel(fila, text=f"{b.titulo}\n{formatear(b.desde)} → {formatear(b.hasta)}",
                         font=ctk.CTkFont(size=11), justify="left").pack(side="left", anchor="w")
            ctk.CTkButton(fila, text="✕", width=30, height=26, fg_color="#7f1d1d",
                          hover_color="#991b1b",
                          command=lambda bl=b: self._quitar_bloqueo_directo(h, bl)).pack(
                side="right", padx=2)

    def _quitar_bloqueo_directo(self, h: Habitacion, b: Bloqueo):
        if h.quitar_bloqueo(b.desde, b.hasta, titulo=b.titulo):
            self._refrescar_todo()
            self._status(f"Bloqueo '{b.titulo}' quitado de {h.numero}.")


def ejecutar_selftest() -> int:
    app = App()
    app.withdraw()
    app.update_idletasks()
    app.cargar_demo()
    assert len(app.grupos) == 4, "demo deberia tener 4 grupos"
    assert len(app.hotel.habitaciones) == 43
    app.simular_distribucion()
    assert app.reporte is not None and app.reporte.todo_alojado, "demo debia quedar completa"
    app.numero_seleccionado = "401"
    app._actualizar_detalle()
    import tempfile
    carpeta = tempfile.mkdtemp(prefix="selftest_gui_")
    ruta_json = os.path.join(carpeta, "datos.json")
    guardar_datos(ruta_json, app.hotel, app.grupos)
    hotel2, grupos2 = cargar_datos(ruta_json)
    assert len(grupos2) == 4 and len(hotel2.habitaciones) == 43
    ruta_csv = os.path.join(carpeta, "salida.csv")
    filas = exportar_csv(app.reporte, ruta_csv)
    assert filas == 49, f"csv esperaba 49 filas, salieron {filas}"
    assert generar_numeros(set(), 3) == ["101", "102", "103"]
    assert generar_numeros({"101", "102", "205"}, 2) == ["301", "302"]
    assert generar_numeros({"A1", "9"}, 2) == ["101", "102"]
    lote = crear_pax_en_lote(40, TipoPax.ADULTO, Sexo.F)
    assert len(lote) == 40 and all(p.nombre.startswith("Pax ") for p in lote)
    grupo_lote = Grupo(nombre="Lote rapido",
                       fecha_ingreso=date.today() + timedelta(days=40),
                       fecha_egreso=date.today() + timedelta(days=42),
                       pax=lote[:10], config=ConfigGrupo())
    app.grupos.append(grupo_lote)
    app.simular_distribucion()
    assert app.reporte.todo_alojado, "el lote rapido debia alojarse completo"
    app.hotel = Hotel(nombre="Tope Test", habitaciones=[
        Habitacion(numero=f"{i}00", capacidad=10) for i in range(1, 4)])
    app.grupos = [Grupo(nombre="TopeGrupo",
                        fecha_ingreso=date.today() + timedelta(days=60),
                        fecha_egreso=date.today() + timedelta(days=62),
                        pax=crear_pax_en_lote(9, TipoPax.ADULTO, Sexo.X),
                        config=ConfigGrupo())]
    app.reporte = None
    app.numero_seleccionado = None
    app.entrada_tope_plazas.delete(0, "end")
    app.entrada_tope_plazas.insert(0, "4")
    app.simular_distribucion()
    assert app.reporte.todo_alojado, "grupo de 9 debia alojarse con tope 4"
    usadas = sum(d.habitaciones_usadas for d in app.reporte.distribuciones)
    assert usadas == 3, f"tope 4 debia repartir en 3 habitaciones, uso {usadas}"
    app.entrada_tope_plazas.delete(0, "end")
    assert app._capacidad_sugerida() == 10, "la sugerencia debia adaptar a las hab. de 10"
    ruta_xlsx = os.path.join(carpeta, "distribucion.xlsx")
    filas_excel = exportar_excel(app.reporte, ruta_xlsx)
    assert filas_excel == 9, f"excel esperaba 9 pasajeros, salieron {filas_excel}"
    from openpyxl import load_workbook
    libro = load_workbook(ruta_xlsx)
    assert set(libro.sheetnames) == {"Mapa del hotel", "Resumen", "Pasajeros"}
    mapa = libro["Mapa del hotel"]
    assert mapa.max_row > 8 and any(
        celda.value and str(celda.value).startswith("PISO")
        for fila_celdas in mapa.iter_rows() for celda in fila_celdas)
    rack = DialogoRack(app, app.hotel)
    rack.entrada_ingreso.delete(0, "end")
    rack.entrada_ingreso.insert(0, str(date.today() + timedelta(days=90)))
    rack.entrada_egreso.delete(0, "end")
    rack.entrada_egreso.insert(0, str(date.today() + timedelta(days=92)))
    rack.entrada_cantidad.insert(0, "20")
    rack._calcular()
    texto_aptos = rack.texto_resultado.get("1.0", "end")
    assert "Pax estimados" in texto_aptos and "Multiplicador" in texto_aptos
    rack.selector_modo.set("Pasajeros")
    rack._actualizar_etiqueta()
    rack.entrada_cantidad.delete(0, "end")
    rack.entrada_cantidad.insert(0, "25")
    rack._calcular()
    texto_pax = rack.texto_resultado.get("1.0", "end")
    assert "Apartamentos estimados" in texto_pax and "Mejor caso" in texto_pax
    assert "≈3" in texto_pax, f"esperaba 3 aptos para 25 pax en hab de 10: {texto_pax}"
    rack.entrada_multiplicador.delete(0, "end")
    rack.entrada_multiplicador.insert(0, "3.5")
    rack.selector_modo.set("Apartamentos")
    rack._actualizar_etiqueta()
    rack.entrada_cantidad.delete(0, "end")
    rack.entrada_cantidad.insert(0, "4")
    rack._marcar_en_mapas()
    prov4 = app.hotel.provisional
    assert prov4 and prov4["aptos_noche"] == 4 and prov4["pax_noche"] == 14, prov4
    rack.selector_modo.set("Pasajeros")
    rack._actualizar_etiqueta()
    rack.entrada_cantidad.delete(0, "end")
    rack.entrada_cantidad.insert(0, "25")
    rack._marcar_en_mapas()
    prov25 = app.hotel.provisional
    assert prov25 and prov25["aptos_noche"] == 8 and prov25["pax_noche"] == 25, prov25
    texto_marcado = rack.etiqueta_marcado.cget("text")
    assert "25 pax" in texto_marcado and "8 apartamentos/noche" in texto_marcado
    rack.destroy()
    hab_100 = app._habitacion_por_numero("100")
    cap_antes = hab_100.capacidad
    wait_original = app.wait_window

    def espera_y_edita(dialogo):
        if isinstance(dialogo, DialogoEditarHabitacion):
            dialogo.entrada_capacidad.delete(0, "end")
            dialogo.entrada_capacidad.insert(0, str(cap_antes + 2))
            dialogo._confirmar()
        else:
            wait_original(dialogo)

    app.wait_window = espera_y_edita
    app.numero_seleccionado = "100"
    app._actualizar_detalle()
    app.boton_editar_hab.invoke()
    app.wait_window = wait_original
    assert hab_100.capacidad == cap_antes + 2 and app.reporte is None

    from unittest import mock as _mock
    _errores_e = []

    _dlg_error = DialogoHabitaciones(app, set())
    _dlg_error.entrada_cantidad.insert(0, "")
    with _mock.patch("tkinter.messagebox.showerror",
                     side_effect=lambda t, m, **k: _errores_e.append(m)):
        _dlg_error._confirmar()
    _dlg_error.destroy()
    assert _dlg_error.resultado is None, "cantidad vacia no debe agregar"
    assert _errores_e and "cuántas habitaciones" in _errores_e[-1], _errores_e

    _dlg_error = DialogoHabitaciones(app, set())
    _dlg_error.entrada_cantidad.insert(0, "3")
    _dlg_error.entrada_capacidad.insert(0, "4,5")
    with _mock.patch("tkinter.messagebox.showerror",
                     side_effect=lambda t, m, **k: _errores_e.append(m)):
        _dlg_error._confirmar()
    _dlg_error.destroy()
    assert _dlg_error.resultado is None and "La capacidad" in _errores_e[-1], _errores_e[-1]

    _dlg_error = DialogoHabitaciones(app, set())
    _dlg_error.entrada_cantidad.insert(0, "3")
    _dlg_error.entrada_numero.insert(0, "A12")
    with _mock.patch("tkinter.messagebox.showerror",
                     side_effect=lambda t, m, **k: _errores_e.append(m)):
        _dlg_error._confirmar()
    _dlg_error.destroy()
    assert _dlg_error.resultado is None and "número inicial" in _errores_e[-1], _errores_e[-1]

    _dlg_ok = DialogoHabitaciones(app, {"101", "102"})
    _dlg_ok.entrada_cantidad.insert(0, "2")
    _dlg_ok._confirmar()
    assert _dlg_ok.resultado is not None
    assert [h.numero for h in _dlg_ok.resultado] == ["201", "202"], \
    [h.numero for h in _dlg_ok.resultado]
    _dlg_ok.destroy()

    app.hotel.lotes_camas.clear()
    dlg_lote = DialogoCamasLote(app)
    dlg_lote.entrada_titulo.insert(0, "Temporada alta")
    dlg_lote.entrada_cantidad.insert(0, "50")
    dlg_lote.entrada_desde.insert(0, formatear(date.today() + timedelta(days=59)))
    dlg_lote.entrada_hasta.insert(0, formatear(date.today() + timedelta(days=63)))
    dlg_lote._confirmar()
    assert dlg_lote.resultado
    app.hotel.lotes_camas.append(dlg_lote.lote)
    assert len(app.hotel.lotes_camas) == 1
    app.hotel.lotes_camas.append(LoteCamas(titulo="Fijas", cantidad=20))
    dia_pico = date.today() + timedelta(days=61)
    assert app.hotel.camas_inventario_en(dia_pico) == 70
    assert app.hotel.camas_inventario_en(date.today()) == 20
    app.simular_distribucion()
    assert app._camas_consumidas_en(dia_pico) == 9
    datos_hotel = hotel_a_dict(app.hotel)
    hotel2 = hotel_de_dict(datos_hotel)
    assert len(hotel2.lotes_camas) == 2 and hotel2.lotes_camas[0].cantidad == 50
    assert hotel2.lotes_camas[0].disponible_hasta == date.today() + timedelta(days=63)
    app.entrada_vista_desde.insert(0, formatear(date.today() + timedelta(days=59)))
    app.entrada_vista_hasta.insert(0, formatear(date.today() + timedelta(days=63)))
    app._aplicar_rango_vista()
    texto_consumo = app.etiqueta_consumo.cget("text")
    assert "1 UH" in texto_consumo and "9 camas" in texto_consumo, texto_consumo
    assert len(app.canvas_camas.find_all()) > 0, "mapa de camas vacío"

    _60 = date.today() + timedelta(days=60)
    app.hotel.provisional = {"titulo": "25 pax",
                            "desde": date.today() + timedelta(days=59),
                            "hasta": date.today() + timedelta(days=62),
                            "pax_noche": 25, "aptos_noche": 8}
    app._refrescar_todo()
    assert app._pax_provisional_en(dia_pico) == 25
    assert app._pax_provisional_en(date.today() + timedelta(days=63)) == 0
    textos_camas = [app.canvas_camas.itemcget(i, "text")
                    for i in app.canvas_camas.find_all() if app.canvas_camas.type(i) == "text"]
    assert "36" in textos_camas, f"camas libres con provisorio deberia ser 36: {textos_camas}"
    app._actualizar_consumo()
    texto_consumo = app.etiqueta_consumo.cget("text")
    assert "24 aptos de cotización" in texto_consumo, texto_consumo
    assert "8 aptos/noche" in texto_consumo and "Cotización 25 pax" in texto_consumo, texto_consumo
    datos_hotel = hotel_a_dict(app.hotel)
    hotel2 = hotel_de_dict(datos_hotel)
    assert hotel2.provisional and hotel2.provisional["pax_noche"] == 25, hotel2.provisional
    cuartos_prov = app._cuartos_para_provisional()
    assert 0 < len(cuartos_prov) <= 8, f"marca hasta {len(cuartos_prov)} aptos"
    assert all(h.libre_en_rango(app.hotel.provisional["desde"], app.hotel.provisional["hasta"])
               for h in cuartos_prov)
    marcados = set()
    for i in app.canvas.find_all():
        for tag in app.canvas.gettags(i):
            if tag.startswith("prov:"):
                marcados.add(tag[5:])
    assert marcados == {h.numero for h in cuartos_prov}, (marcados, [h.numero for h in cuartos_prov])
    app.hotel.provisional = None
    app._refrescar_todo()
    assert app._pax_provisional_en(dia_pico) == 0
    assert not [i for i in app.canvas.find_all()
                if any(t.startswith("prov:") for t in app.canvas.gettags(i))], \
        "no deberian quedar marcados provisorios tras quitar la cotizacion"

    assert app._rango_matriz() is not None
    desde_m, hasta_m = app._rango_matriz()
    dias_matriz = (hasta_m - desde_m).days + 1
    assert dias_matriz == 5, f"rango matriz esperaba 5 dias, hay {dias_matriz}"
    app._redibujar_matriz()
    textos_m = [app.canvas_matriz.itemcget(i, "text")
                for i in app.canvas_matriz.find_all()
                if app.canvas_matriz.type(i) == "text"]
    assert "LÍMITE DIÁRIO" in textos_m and "70" in textos_m and "61" in textos_m, textos_m
    app.matriz_modo.set("Apartamentos")
    app._redibujar_matriz()
    textos_ma = [app.canvas_matriz.itemcget(i, "text")
                 for i in app.canvas_matriz.find_all()
                 if app.canvas_matriz.type(i) == "text"]
    assert "3" in textos_ma and "2" in textos_ma, textos_ma
    app.matriz_modo.set("Pessoas")
    ruta_mapa_p = os.path.join(carpeta, "mapa-pessoas.xlsx")
    pend_p = exportar_mapa_grupos(ruta_mapa_p, app.hotel, app.grupos, app.reporte,
                                  desde_m, hasta_m, "Pessoas")
    assert pend_p == 0
    libro_mp = load_workbook(ruta_mapa_p)
    hoja_mp = libro_mp["Mapa"]
    assert hoja_mp.cell(5, 1).value == "Empresa" and hoja_mp.cell(5, 2).value == "PAX"
    filas_mp = {hoja_mp.cell(r, 1).value: r for r in range(7, hoja_mp.max_row + 1)}
    assert "TopeGrupo" in filas_mp
    assert filas_mp["TopeGrupo"] and hoja_mp.cell(filas_mp["TopeGrupo"], 2).value == 9
    vals_p = [hoja_mp.cell(filas_mp["TopeGrupo"], c).value for c in range(5, 5 + dias_matriz)]
    assert vals_p.count(9) == 2, vals_p
    vals_lim = [hoja_mp.cell(filas_mp["Limite Diário"], c).value
                for c in range(5, 5 + dias_matriz)]
    assert set(vals_lim) == {70}, vals_lim
    vals_disp = [hoja_mp.cell(filas_mp["Disponíveis"], c).value
                 for c in range(5, 5 + dias_matriz)]
    assert 61 in vals_disp and 70 in vals_disp, vals_disp
    ruta_mapa_a = os.path.join(carpeta, "mapa-apartamentos.xlsx")
    pend_a = exportar_mapa_grupos(ruta_mapa_a, app.hotel, app.grupos, app.reporte,
                                  desde_m, hasta_m, "Apartamentos")
    assert pend_a == 0
    libro_ma = load_workbook(ruta_mapa_a)
    hoja_ma = libro_ma["Mapa"]
    filas_ma = {hoja_ma.cell(r, 1).value: r for r in range(7, hoja_ma.max_row + 1)}
    vals_a = [hoja_ma.cell(filas_ma["TopeGrupo"], c).value for c in range(5, 5 + dias_matriz)]
    assert 1 in vals_a, vals_a

    ruta_rooming = os.path.join(carpeta, "rooming.xlsx")
    filas_rooming = exportar_rooming_excel(ruta_rooming, app.reporte, destino="Camboriú")
    assert filas_rooming > 0
    libro_rm = load_workbook(ruta_rooming)
    assert libro_rm.sheetnames == ["Resumen por Grupo", "Detalle por Categoría",
                                   "Cambios", "Límites por noche"]
    hoja_rm = libro_rm["Resumen por Grupo"]
    filas_rm = [hoja_rm.cell(r, 1).value for r in range(2, hoja_rm.max_row + 1)]
    assert all(r for r in filas_rm), filas_rm
    assert all(hoja_rm.cell(r, 7).value > 0 for r in range(2, hoja_rm.max_row + 1)), filas_rm
    detalle_rm = libro_rm["Detalle por Categoría"]
    categorias_rm = {detalle_rm.cell(r, 5).value for r in range(2, detalle_rm.max_row + 1)}
    assert categorias_rm, categorias_rm

    rooms_backup = list(app.hotel.habitaciones)
    app.hotel.habitaciones.clear()
    app.reporte = None
    app.simular_distribucion()
    texto_solo_camas = app.texto_reporte.get("1.0", "end")
    assert "MODO SOLO CAMAS" in texto_solo_camas
    assert int(app.tarjetas_resumen["camas"].cget("text")) >= 70
    assert app._camas_consumidas_en(dia_pico) == 9
    app.hotel.habitaciones.extend(rooms_backup)
    app.simular_distribucion()
    assert app.reporte is not None

    n_hoteles = len(app.hoteles)
    app._crear_hotel("Hotel Sol")
    assert len(app.hoteles) == n_hoteles + 1
    assert app.hotel.nombre == "Hotel Sol" and app.grupos == []
    app.grupos.append(Grupo(nombre="G-Sol",
                            fecha_ingreso=date.today() + timedelta(days=5),
                            fecha_egreso=date.today() + timedelta(days=8),
                            config=ConfigGrupo(), pax=[]))
    app._cambiar_hotel(0)
    assert app.hotel.nombre == "Mi Hotel" and app.grupos == []
    app._cambiar_hotel(n_hoteles)
    assert app.grupos[0].nombre == "G-Sol"
    import tempfile
    ruta_tmp = os.path.join(tempfile.gettempdir(), "multi_hoteles_test.json")
    guardar_multi(ruta_tmp, app.hoteles, app.grupos_por_hotel)
    hoteles2, grupos2 = cargar_multi(ruta_tmp)
    os.remove(ruta_tmp)
    assert len(hoteles2) == n_hoteles + 1
    assert [g.nombre for g in grupos2[n_hoteles]] == ["G-Sol"]
    assert hoteles2[n_hoteles].nombre == "Hotel Sol"

    # --- Modo de hospedaje: el dialogo tiene que exponer y guardar los 3 ajustes nuevos ---
    from modelos import ModoHospedaje
    # hotel propio de 8 habitaciones de 4 plazas, para no depender del orden de los tests
    app._crear_hotel("Hotel Modo")
    app.hotel.habitaciones = generar_hotel_estimado("Hotel Modo", 8, 4).habitaciones
    app._refrescar_todo()
    assert len(app.hotel.habitaciones) == 8 and app.hotel.habitaciones[0].capacidad == 4

    dlg_modo = DialogoGrupo(app)
    dlg_modo.entrada_nombre.insert(0, "Grupo Escol")
    dlg_modo.entrada_max_plazas.insert(0, "3")
    dlg_modo._set_modo_hospedaje(ModoHospedaje.ADULTOS.value)
    assert dlg_modo.switch_adultos_por_rol.cget("state") == "disabled", \
        "en modo adultos el switch de rol debe quedar apagado"
    dlg_modo._set_modo_hospedaje(ModoHospedaje.ESTUDIANTIL.value)
    assert dlg_modo.switch_adultos_por_rol.cget("state") == "normal", \
        "en modo estudiantil el switch de rol debe quedar habilitado"
    dlg_modo.entrada_cupo.insert(0, "3")
    for _tipo in (TipoPax.ADULTO, TipoPax.CHOFER, TipoPax.ESTUDIANTE):
        dlg_modo.entrada_lote_cantidad.delete(0, "end")
        dlg_modo.entrada_lote_cantidad.insert(0, "6")
        dlg_modo.menu_lote_tipo.set(_tipo.value)
        dlg_modo.sexo_lote.set("M")
        dlg_modo._agregar_lote()
    dlg_modo._guardar()
    cfg_modo = dlg_modo.resultado.config
    assert cfg_modo.modo_hospedaje == ModoHospedaje.ESTUDIANTIL.value
    assert cfg_modo.max_plazas_estudiantes == 3
    assert cfg_modo.adultos_por_rol is True
    assert cfg_modo.estudiantes_por_habitacion == 3

    # los choferes se agrupan entre si (2 habitaciones, no 6) y van aparte de los alumnos
    d = distribuir(app.hotel, [dlg_modo.resultado]).distribuciones[0]
    casas_chofer = [a for a in d.asignaciones
                    if all(m.tipo == TipoPax.CHOFER for at in a.atomos for m in at.miembros)]
    capacidad_max = max(h.capacidad for h in app.hotel.habitaciones)
    esperadas = -(-6 // capacidad_max)
    assert len(casas_chofer) == esperadas, \
        f"los 6 choferes deberian ocupar {esperadas} hab (grupo), no {len(casas_chofer)}"
    for a in d.asignaciones:
        tipos = {m.tipo for at in a.atomos for m in at.miembros}
        assert not (TipoPax.ESTUDIANTE in tipos and len(tipos) > 1), \
            f"se mezclaron alumnos con acompanantes en {a.habitacion.numero}"
    assert "ESTUDIANTIL" in d.estrategia_usada, d.estrategia_usada

    print(f"SELFTEST GUI OK — demo+lote {49 + len(lote[:10])}/"
          f"{49 + len(lote[:10])} pax, tope 9pax en 3hab, excel ok (mapa+resumen+pax), "
          f"sugerida adaptativa=10, rack ok (x3.5 marca en ambos mapas, camas libres 36), "
          f"inventario camas (lotes+por dia+json) ok, matriz grupos×noches "
          f"(pessoas 70/61 + apartamentos ok + export xlsx ok), "
          f"rooming xlsx ok, "
          f"modo solo camas ok, multi-hotel (crear+cambiar+json) ok, "
          f"modo estudiantil ok (alumnos separados + choferes agrupados + cap 3)")
    app.destroy()
    return 0


def main() -> int:
    if "--selftest" in sys.argv:
        return ejecutar_selftest()
    aplicacion = App()
    aplicacion.mainloop()
    return 0


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        detalle = traceback.format_exc()
        try:
            with open(ruta_log(), "w", encoding="utf-8") as archivo:
                archivo.write(detalle)
        except OSError:
            pass
        if "--selftest" in sys.argv:
            try:
                sys.stderr.write(detalle)
            except (AttributeError, OSError):
                pass
            sys.exit(1)
        raiz = tk.Tk()
        raiz.withdraw()
        messagebox.showerror("Error inesperado",
                             "Ocurrió un error y la aplicación se cerró.\n\n"
                             f"{detalle[-900:]}")
