# -*- coding: utf-8 -*-
"""Genera un ejemplo comparativo del Rooming antes/después de las modificaciones.

Genera tmp_ejemplo/rooming_anterior.xlsx (estado viejo, sin ESTIMADO/CONFIRMADO)
y tmp_ejemplo/rooming_nuevo.xlsx (con las marcas nuevas). Después captura la hoja
'Resumen por Grupo' de ambos en PNG para revisión visual.
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

import win32gui  # noqa: E402
from PIL import ImageGrab  # noqa: E402

from demo import construir_datos_demo  # noqa: E402
from excel_exportador import exportar_rooming_excel  # noqa: E402
from motor import OpcionesMotor, distribuir  # noqa: E402

BASE = os.path.join(os.path.dirname(os.path.abspath(__file__)), "tmp_ejemplo")
os.makedirs(BASE, exist_ok=True)
ANTERIOR = os.path.join(BASE, "rooming_anterior.xlsx")
NUEVO = os.path.join(BASE, "rooming_nuevo.xlsx")


def _hwnd_excel(excel):
    # Prioridad: la ventana de la INSTANCIA que abrimos (excel.Hwnd disponible
    # en Office 2010+). Evita capturar otras ventanas de Excel abiertas por el usuario.
    try:
        hwnd = excel.Hwnd
        if hwnd:
            return hwnd
    except Exception:
        pass
    # Fallback: buscar la última ventana de clase XLMAIN (menos fiable).
    encontradas = []
    def _cb(hwnd, extra):
        if win32gui.GetClassName(hwnd) == "XLMAIN":
            encontradas.append(hwnd)
        return True
    win32gui.EnumWindows(_cb, None)
    return encontradas[-1] if encontradas else None


def _traer_al_frente(hwnd):
    try:
        win32gui.ShowWindow(hwnd, 9)
        win32gui.SetForegroundWindow(hwnd)
        win32gui.BringWindowToTop(hwnd)
    except Exception:
        pass


def _título(hwnd):
    try:
        return win32gui.GetWindowText(hwnd)
    except Exception:
        return ""


def capturar(excel, libro, hoja: str, zoom: int, nombre: str) -> None:
    hoja_activa = libro.Worksheets(hoja)
    hoja_activa.Activate()
    hoja_activa.Select()
    excel.ActiveWindow.Zoom = zoom
    excel.ActiveWindow.DisplayGridlines = True
    excel.ScreenUpdating = True
    excel.Visible = True
    try:
        excel.ActiveWindow.ScrollRow = 1
        excel.ActiveWindow.ScrollColumn = 1
    except Exception:
        pass
    for _ in range(40):
        time.sleep(0.02)
    hwnd = _hwnd_excel(excel)
    if hwnd is None:
        hwnd = excel.Hwnd if hasattr(excel, "Hwnd") else None
    if hwnd:
        _traer_al_frente(hwnd)
        time.sleep(0.8)
    for _ in range(40):
        time.sleep(0.02)
    rect = win32gui.GetWindowRect(hwnd) if hwnd else None
    if rect:
        x1, y1, x2, y2 = rect
        ancho, alto = x2 - x1, y2 - y1
        imagen = ImageGrab.grab(bbox=(x1, y1, x2, y2))
        imagen = imagen.crop((2, 8, ancho - 2, alto - 2))
    else:
        imagen = ImageGrab.grab()
    ruta = os.path.join(BASE, nombre)
    imagen.save(ruta)
    print(f"  OK {nombre} ({imagen.size}) titulo='{_título(hwnd) if hwnd else ''}'")


def principal():
    hotel, grupos = construir_datos_demo()

    # ── Escenario ANTERIOR: nada confirmado/provisional (comportamiento previo) ──
    hotel_a, grupos_a = construir_datos_demo()
    for g in grupos_a:
        g.provisional = False
        g.confirmado = False

    # ── Escenario NUEVO: 1 estimado, 1 confirmado, 1 sin fecha ──
    hotel_n, grupos_n = construir_datos_demo()

    # Primera corrida SIN congelar: se distribuye igual que siempre
    reporte_inicial = distribuir(hotel_n, list(grupos_n), OpcionesMotor())

    # Confirmar 'Primaria 5° B' y volver a simular en modo CONSERVAR
    for g in grupos_n:
        if g.nombre == "Primaria 5° B":
            g.confirmado = True
            g.provisional = False
    reporte_nuevo = distribuir(hotel_n, list(grupos_n), OpcionesMotor(),
                               previo=reporte_inicial, reoptimizar=False)

    # 'Bachilleres 3° A' queda como ESTIMADO (provisional)
    for g in grupos_n:
        if g.nombre == "Bachilleres 3° A":
            g.provisional = True
    reporte_nuevo = distribuir(hotel_n, list(grupos_n), OpcionesMotor(),
                               previo=reporte_nuevo, reoptimizar=False)

    exp_anterior = distribuir(hotel_a, list(grupos_a), OpcionesMotor())
    exportar_rooming_excel(ANTERIOR, exp_anterior, destino="Camboriú")
    exportar_rooming_excel(NUEVO, reporte_nuevo, destino="Camboriú")

    print("Archivos generados:")
    for r in (ANTERIOR, NUEVO):
        print("  ", r, os.path.getsize(r), "bytes")

    import win32com.client
    excel = win32com.client.DispatchEx("Excel.Application")
    excel.Visible = True
    excel.DisplayAlerts = False
    try:
        libro_a = excel.Workbooks.Open(ANTERIOR)
        excel.ActiveWindow.WindowState = 2  # maximizar para capturar bien
        capturar(excel, libro_a, "Resumen por Grupo", 110, "08_rooming_anterior.png")
        libro_a.Close(False)
        libro_n = excel.Workbooks.Open(NUEVO)
        excel.ActiveWindow.WindowState = 2
        capturar(excel, libro_n, "Resumen por Grupo", 110, "09_rooming_nuevo.png")
        capturar(excel, libro_n, "Límites por noche", 120, "10_rooming_limites.png")
        libro_n.Close(False)
    finally:
        excel.Quit()
    print("Listo. Imagenes en", BASE)


if __name__ == "__main__":
    principal()