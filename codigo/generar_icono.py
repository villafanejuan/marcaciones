# -*- coding: utf-8 -*-
"""Genera el icono del ejecutable (planilla + check de aprobacion)."""
import os
from PIL import Image, ImageDraw

RAIZ = os.path.dirname(os.path.abspath(__file__))
OUT = os.path.join(RAIZ, "..", "icono.ico")
T = 512
S = T * 4  # supersample

img = Image.new("RGBA", (S, S), (0, 0, 0, 0))
d = ImageDraw.Draw(img)


def rr(x0, y0, x1, y1, rad, fill):
    d.rounded_rectangle([x0, y0, x1, y1], radius=rad, fill=fill)


# fondo redondeado con gradiente azul vertical
GRAD = [(31, 78, 121), (38, 92, 138), (46, 109, 164)]
for y in range(S):
    t = y / S
    i0 = int(t * (len(GRAD) - 1))
    i1 = min(i0 + 1, len(GRAD) - 1)
    f = t * (len(GRAD) - 1) - i0
    c = tuple(int(GRAD[i0][k] + (GRAD[i1][k] - GRAD[i0][k]) * f) for k in range(3))
    rr(0, y, S, y + 1, 0, c + (255,))
rr(int(S * 0.02), int(S * 0.02), int(S * 0.98), int(S * 0.98), int(S * 0.16), None)  # clip via mask later

# mascara redondeada
mask = Image.new("L", (S, S), 0)
dm = ImageDraw.Draw(mask)
dm.rounded_rectangle([0, 0, S, S], radius=int(S * 0.16), fill=255)
img.putalpha(mask)

# hoja de planilla (blanca)
x0, y0, x1, y1 = int(S * 0.20), int(S * 0.14), int(S * 0.80), int(S * 0.86)
rr(x0, y0, x1, y1, int(S * 0.035), (255, 255, 255, 255))

# barra superior azul (titulo)
top_y = int(S * 0.16)
d.rounded_rectangle([x0, top_y, x1, top_y + int(S * 0.09)],
                    radius=int(S * 0.035), fill=(31, 78, 121, 255))
d.rectangle([x0, top_y + int(S * 0.06), x1, top_y + int(S * 0.09)],
            fill=(31, 78, 121, 255))

# renglones / grillas
gris = (90, 102, 116, 255)
lbl = (60, 72, 86, 255)
lin_x = x0 + int(S * 0.19)
d.line([lin_x, top_y + int(S * 0.13), lin_x, y1 - int(S * 0.05)], fill=gris, width=3)
d.line([x0 + int(S * 0.06), top_y + int(S * 0.25), x1 - int(S * 0.05), top_y + int(S * 0.25)], fill=gris, width=3)
d.line([x0 + int(S * 0.06), top_y + int(S * 0.36), x1 - int(S * 0.05), top_y + int(S * 0.36)], fill=gris, width=3)
d.line([x0 + int(S * 0.06), top_y + int(S * 0.47), x1 - int(S * 0.05), top_y + int(S * 0.47)], fill=gris, width=3)

# mark-check verde (circulo + tilde)
cx, cy, rad = int(S * 0.775), int(S * 0.185), int(S * 0.062)
rr(cx - rad, cy - rad, cx + rad, cy + rad, rad, (46, 160, 67, 255))
d.line([cx - int(S * 0.032), cy - int(S * 0.002), cx - int(S * 0.008), cy + int(S * 0.024)],
       fill=(255, 255, 255, 255), width=int(S * 0.015), joint="curve")
d.line([cx - int(S * 0.008), cy + int(S * 0.024), cx + int(S * 0.036), cy - int(S * 0.026)],
       fill=(255, 255, 255, 255), width=int(S * 0.015), joint="curve")

# baja resolucion
img512 = img.reduce(4)  # 512 -> 128
img512.save(OUT, sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)])
print("icono guardado:", OUT)