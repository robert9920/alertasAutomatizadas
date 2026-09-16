"""Genera recursos/app.ico (curva S sobre un sobre de correo).

    python recursos/crear_icono.py
"""
from __future__ import annotations

from pathlib import Path

from PIL import Image, ImageDraw

AZUL = (37, 99, 235, 255)
BLANCO = (255, 255, 255, 255)
VERDE = (78, 167, 46, 255)

LADO = 512


def construir() -> Image.Image:
    imagen = Image.new("RGBA", (LADO, LADO), (0, 0, 0, 0))
    lienzo = ImageDraw.Draw(imagen)

    # Fondo redondeado plano
    lienzo.rounded_rectangle([0, 0, LADO - 1, LADO - 1], radius=112, fill=AZUL)

    # Sobre
    izq, arr, der, aba = 96, 150, LADO - 96, LADO - 150
    lienzo.rounded_rectangle([izq, arr, der, aba], radius=26, fill=BLANCO)
    lienzo.line([(izq + 10, arr + 14), (LADO // 2, arr + 132), (der - 10, arr + 14)],
                fill=AZUL, width=28, joint="curve")

    # Curva S sobre el sobre
    puntos = [
        (izq + 34, aba - 34), (izq + 96, aba - 48), (izq + 142, aba - 96),
        (LADO // 2, arr + 118), (der - 130, arr + 84), (der - 70, arr + 72),
        (der - 34, arr + 66),
    ]
    lienzo.line(puntos, fill=VERDE, width=26, joint="curve")
    for x, y in puntos[::3]:
        lienzo.ellipse([x - 18, y - 18, x + 18, y + 18], fill=VERDE)
    return imagen


def main() -> None:
    destino = Path(__file__).resolve().parent / "app.ico"
    imagen = construir()
    imagen.save(
        destino, format="ICO",
        sizes=[(16, 16), (24, 24), (32, 32), (48, 48), (64, 64), (128, 128), (256, 256)],
    )
    imagen.resize((256, 256), Image.LANCZOS).save(destino.with_suffix(".png"))
    print(f"Icono generado: {destino}")


if __name__ == "__main__":
    main()
