#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
generar_icono.py
-----------------
Genera el ícono de la aplicación (scripts/icono.ico) con la estética de la
marca: fondo negro (#121212) y un "$" en verde lima (#B0FF2A).

Se corre una sola vez (o cada vez que se quiera regenerar el ícono):

    python scripts\\generar_icono.py

Requiere Pillow (ya se usa en otras partes del proyecto).
"""

from pathlib import Path

from PIL import Image, ImageDraw, ImageFont

SCRIPT_DIR = Path(__file__).resolve().parent
DESTINO = SCRIPT_DIR / "icono.ico"

COLOR_FONDO = "#121212"
COLOR_ACENTO = "#B0FF2A"

TAMANO_BASE = 256
RADIO_ESQUINAS = 48
TAMANOS_ICO = [(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)]

FUENTES_CANDIDATAS = [
    r"C:\Windows\Fonts\arialbd.ttf",
    r"C:\Windows\Fonts\segoeuib.ttf",
    "arialbd.ttf",
    "arial.ttf",
]


def _cargar_fuente(tamano):
    for ruta in FUENTES_CANDIDATAS:
        try:
            return ImageFont.truetype(ruta, tamano)
        except OSError:
            continue
    return ImageFont.load_default(size=tamano)


def generar():
    imagen = Image.new("RGBA", (TAMANO_BASE, TAMANO_BASE), (0, 0, 0, 0))
    dibujo = ImageDraw.Draw(imagen)
    dibujo.rounded_rectangle(
        [(0, 0), (TAMANO_BASE - 1, TAMANO_BASE - 1)],
        radius=RADIO_ESQUINAS,
        fill=COLOR_FONDO,
    )

    fuente = _cargar_fuente(int(TAMANO_BASE * 0.62))
    texto = "$"
    caja = dibujo.textbbox((0, 0), texto, font=fuente)
    ancho_texto = caja[2] - caja[0]
    alto_texto = caja[3] - caja[1]
    posicion = (
        (TAMANO_BASE - ancho_texto) / 2 - caja[0],
        (TAMANO_BASE - alto_texto) / 2 - caja[1],
    )
    dibujo.text(posicion, texto, font=fuente, fill=COLOR_ACENTO)

    imagen.save(DESTINO, format="ICO", sizes=TAMANOS_ICO)
    print(f"Ícono generado en: {DESTINO}")


if __name__ == "__main__":
    generar()
