"""Registro de los correos enviados (logs/historial.csv)."""
from __future__ import annotations

import csv
import datetime as dt
import logging
from pathlib import Path

from .config import carpeta_logs

CABECERA = [
    "Fecha y hora", "Código Proyecto", "Proyecto", "Asunto", "Para", "CC", "CCO",
    "Semana", "N° entregables", "Resultado", "Detalle",
]


def ruta() -> Path:
    return carpeta_logs() / "historial.csv"


def registrar(codigo: str, proyecto: str, asunto: str, para: list[str], cc: list[str],
              cco: list[str], semana: str, entregables: int, resultado: str,
              detalle: str = "") -> None:
    """Anota el envio. Nunca interrumpe: se llama despues de enviar el correo,
    asi que un fallo al escribir el CSV no debe convertirse en un error visible."""
    destino = ruta()
    nuevo = not destino.exists()
    try:
        with destino.open("a", newline="", encoding="utf-8-sig") as archivo:
            escritor = csv.writer(archivo, delimiter=";")
            if nuevo:
                escritor.writerow(CABECERA)
            escritor.writerow([
                dt.datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                codigo, proyecto, asunto,
                "; ".join(para), "; ".join(cc), "; ".join(cco),
                semana, entregables, resultado, detalle.replace("\n", " | ")[:500],
            ])
    except OSError as exc:
        logging.getLogger("appalertas").warning(
            "No se pudo escribir el historial en %s: %s", destino, exc
        )


def ultimos(cantidad: int = 50) -> list[dict[str, str]]:
    destino = ruta()
    if not destino.is_file():
        return []
    try:
        with destino.open("r", newline="", encoding="utf-8-sig") as archivo:
            filas = list(csv.DictReader(archivo, delimiter=";"))
    except OSError:
        return []
    return filas[-cantidad:][::-1]
