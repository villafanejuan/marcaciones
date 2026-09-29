# -*- coding: utf-8 -*-
"""Genera las capturas de pantalla usadas en el manual (docs/imagenes/).

Ejecutar con:  python docs/capturar_manual.py
Requiere: PIL (pip install pillow) y la app funcionando.
"""
from __future__ import annotations

import ctypes
import os
import sys
import time

sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

from PIL import ImageGrab

from interfaz import (  # noqa: E402
    App,
    DialogoGrupo,
    DialogoHabitaciones,
    DialogoPlanificador,
    DialogoRack,
)

CARPETA = os.path.join(os.path.dirname(os.path.abspath(__file__)), "imagenes")
os.makedirs(CARPETA, exist_ok=True)


def _pixeles(ventana):
    ventana.update_idletasks()
    ventana.update()
    x = ventana.winfo_rootx()
    y = ventana.winfo_rooty()
    ancho = ventana.winfo_width()
    alto = ventana.winfo_height()
    return x, y, ancho, alto


def capturar_ventana(ventana, nombre: str) -> str | None:
    try:
        ventana.deiconify()
        ventana.lift()
        ventana.focus_force()
        ventana.attributes("-topmost", True)
    except Exception:
        pass
    for _ in range(20):
        ventana.update_idletasks()
        ventana.update()
        time.sleep(0.05)
    x, y, ancho, alto = _pixeles(ventana)
    if ancho < 2 or alto < 2:
        print(f"  ! ventana {nombre} sin dimensiones ({ancho}x{alto})")
        return None
    imagen = ImageGrab.grab(bbox=(x, y, x + ancho, y + alto))
    # pequeño recorte del borde para evitar la barra de la ventana del sistema
    imagen = imagen.crop((2, 2, ancho - 2, alto - 2))
    ruta = os.path.join(CARPETA, nombre)
    imagen.save(ruta)
    print(f"  OK {nombre} ({ancho}x{alto})")
    return ruta


def principal():
    app = App()
    app.geometry("1360x800+80+60")
    app.cargar_demo()
    app.simular_distribucion()
    app.numero_seleccionado = "401"
    app._actualizar_detalle()
    capturar_ventana(app, "01_pantalla_principal.png")
    app.destroy()

    dg = DialogoGrupo(None)
    dg.geometry("880x720+120+50")
    capturar_ventana(dg, "02_dialogo_nuevo_grupo.png")
    dg.destroy()

    dh = DialogoHabitaciones(None, capacidad_sugerida=4)
    dh.geometry("380x480+200+150")
    capturar_ventana(dh, "03_dialogo_habitaciones.png")
    dh.destroy()

    app2 = App()
    app2.withdraw()
    app2.update_idletasks()
    app2.cargar_demo()
    dr = DialogoRack(app2, app2.hotel)
    dr.geometry("640x720+120+50")
    dr.entrada_cantidad.delete(0, "end")
    dr.entrada_cantidad.insert(0, "20")
    dr._calcular()
    capturar_ventana(dr, "04_dialogo_rack.png")
    dr.destroy()
    app2.destroy()

    dp = DialogoPlanificador(None)
    dp.geometry("720x760+120+40")
    dp.entrada_nombre.delete(0, "end")
    dp.entrada_nombre.insert(0, "Colegio Norte")
    dp.entrada_total_pax.delete(0, "end")
    dp.entrada_total_pax.insert(0, "45")
    capturar_ventana(dp, "05_dialogo_planificador.png")
    dp.destroy()

    print("Listo. Imagenes en", CARPETA)


if __name__ == "__main__":
    principal()