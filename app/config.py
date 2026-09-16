"""Rutas de la aplicacion y preferencias locales (config.json)."""
from __future__ import annotations

import json
import logging
import os
import sys
from logging.handlers import RotatingFileHandler
from pathlib import Path

from . import APP_NOMBRE
from .constantes import NOMBRE_BD

NOMBRE_CONFIG = "config.json"
_CARPETA_DATOS: Path | None = None


def carpeta_app() -> Path:
    """Carpeta del .exe (o del proyecto cuando se ejecuta desde el codigo)."""
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent.parent


def _es_escribible(carpeta: Path) -> bool:
    try:
        carpeta.mkdir(parents=True, exist_ok=True)
        prueba = carpeta / ".escritura"
        prueba.write_text("", encoding="utf-8")
        prueba.unlink()
        return True
    except OSError:
        return False


def carpeta_datos() -> Path:
    """Donde la app escribe sus logs y preferencias.

    Normalmente la carpeta del .exe, para que siga siendo portable. Si esa
    carpeta es de solo lectura (instalada en Archivos de programa, o abierta
    desde una carpeta de red ajena) se usa el perfil del usuario: sin esto la
    aplicacion no llegaria ni a abrirse.
    """
    global _CARPETA_DATOS
    if _CARPETA_DATOS is not None:
        return _CARPETA_DATOS

    candidata = carpeta_app()
    if not _es_escribible(candidata):
        respaldo = os.environ.get("LOCALAPPDATA") or os.environ.get("APPDATA")
        alternativa = Path(respaldo) / APP_NOMBRE if respaldo else Path.home() / APP_NOMBRE
        if _es_escribible(alternativa):
            candidata = alternativa
    _CARPETA_DATOS = candidata
    return candidata


def ruta_recurso(nombre: str) -> Path | None:
    """Busca un archivo de recursos tanto en el codigo como dentro del .exe."""
    candidatas = []
    empaquetado = getattr(sys, "_MEIPASS", None)
    if empaquetado:
        candidatas.append(Path(empaquetado) / "recursos" / nombre)
    candidatas.append(Path(__file__).resolve().parent.parent / "recursos" / nombre)
    candidatas.append(carpeta_app() / "recursos" / nombre)
    for candidata in candidatas:
        if candidata.is_file():
            return candidata
    return None


def carpeta_logs() -> Path:
    destino = carpeta_datos() / "logs"
    try:
        destino.mkdir(parents=True, exist_ok=True)
    except OSError:
        pass
    return destino


def ruta_config() -> Path:
    return carpeta_datos() / NOMBRE_CONFIG


def leer_config() -> dict:
    ruta = ruta_config()
    if not ruta.is_file():
        return {}
    try:
        return json.loads(ruta.read_text(encoding="utf-8"))
    except (OSError, json.JSONDecodeError):
        return {}


def guardar_config(datos: dict) -> None:
    try:
        ruta_config().write_text(
            json.dumps(datos, indent=2, ensure_ascii=False), encoding="utf-8"
        )
    except OSError:
        logging.getLogger("appalertas").warning("No se pudo guardar config.json")


def ruta_bd() -> Path | None:
    """Ubicacion del Excel de base de datos, si se puede determinar."""
    guardada = leer_config().get("ruta_bd")
    if guardada and Path(guardada).is_file():
        return Path(guardada)
    candidata = carpeta_app() / NOMBRE_BD
    if candidata.is_file():
        return candidata
    return None


def fijar_ruta_bd(ruta: Path) -> None:
    datos = leer_config()
    datos["ruta_bd"] = str(Path(ruta).resolve())
    guardar_config(datos)


def configurar_log() -> logging.Logger:
    logger = logging.getLogger("appalertas")
    if logger.handlers:
        return logger
    logger.setLevel(logging.INFO)
    try:
        manejador = RotatingFileHandler(
            carpeta_logs() / "app.log", maxBytes=1_000_000, backupCount=3, encoding="utf-8"
        )
        manejador.setFormatter(
            logging.Formatter("%(asctime)s  %(levelname)-7s  %(message)s", "%Y-%m-%d %H:%M:%S")
        )
    except OSError:
        # Sin sitio donde escribir, la aplicacion debe abrirse igual.
        manejador = logging.NullHandler()
    logger.addHandler(manejador)
    return logger
