# -*- coding: utf-8 -*-
"""Renderiza las hojas PROGRAMACION de las planillas generadas a imagenes PNG."""
import os
import openpyxl
from PIL import Image, ImageDraw, ImageFont

BASE = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(BASE, "img")
os.makedirs(OUT, exist_ok=True)

FONT = r"C:/Windows/Fonts/arial.ttf"
FONT_B = r"C:/Windows/Fonts/arialbd.ttf"
SCALE = 2


def font(size, bold=False):
    p = FONT_B if bold else FONT
    return ImageFont.truetype(p, size * SCALE)


def wrap(draw, text, f, maxw):
    lines = []
    for par in text.split("\n"):
        cur = ""
        for w in par.split(" "):
            t = (cur + " " + w).strip()
            if draw.textlength(t, font=f) > maxw and cur:
                lines.append(cur)
                cur = w
            else:
                cur = t
        lines.append(cur)
    return lines


def render_sheet(path, sheet_name, out_png, col_start=2, n_cols=8, max_rows=5,
                 width_col=190, height_col=110, header_h=60, title=None):
    wb = openpyxl.load_workbook(path)
    ws = wb[sheet_name]
    cols = list(range(col_start, col_start + n_cols))

    rows = list(range(1, 1 + max_rows))
    H = header_h + len(rows) * height_col
    W = len(cols) * width_col
    img = Image.new("RGB", (W * SCALE, H * SCALE), "white")
    d = ImageDraw.Draw(img)

    y = 0
    if title:
        d.text((8 * SCALE, 8 * SCALE), title, fill="black", font=font(20, True))
        y += 44 * SCALE
    d.rectangle([0, y, W * SCALE, (y + header_h) * SCALE], fill="#DDEBF7")
    for i, c in enumerate(cols):
        v = ws.cell(row=2, column=c).value or ""
        d.text(((i * width_col + 6) * SCALE, (y + 10) * SCALE), str(v),
               fill="black", font=font(14, True))
    y += header_h

    for j, r in enumerate(rows):
        ry = y + j * height_col
        for i, c in enumerate(cols):
            cx = i * width_col
            v = ws.cell(row=r, column=c).value
            d.rectangle([cx * SCALE, ry * SCALE, (cx + width_col) * SCALE,
                         (ry + height_col) * SCALE], outline="#B0B0B0", width=1)
            if v is None:
                continue
            f = font(11)
            lines = wrap(d, str(v), f, (width_col - 14) * SCALE)
            ty = ry * SCALE + 6 * SCALE
            for ln in lines[:11]:
                d.text(((cx + 7) * SCALE, ty), ln, font=f, fill="black")
                ty += int(f.size * 1.15)
    img = img.resize((W, H), Image.LANCZOS)
    img.save(out_png)
    return out_png


SAN = os.path.join(BASE, "..", "salidas", "SALIDA_WORD_OPERACIONES SAN ANTONIO DE PAD.xlsx")
REC = os.path.join(BASE, "..", "salidas", "SALIDA_SISTEMA_RECEPTIVO.xlsx")
INKA = os.path.join(BASE, "..", "salidas", "SALIDA_OPERATIVO_2026.xlsx")

print(render_sheet(SAN, "PROGRAMACION 09-10", os.path.join(OUT, "word.jpg"),
                   title="Formato Word (datos completos: horas, lugares, vuelo)"))
print(render_sheet(REC, "PROGRAMACION 17-09", os.path.join(OUT, "receptivo.jpg"),
                   title="Formato Receptivo (sin horas -> ____H)"))
print(render_sheet(INKA, "PROGRAMACION", os.path.join(OUT, "inka.jpg"),
                   title="Formato INKA (sin horas -> ____H)"))