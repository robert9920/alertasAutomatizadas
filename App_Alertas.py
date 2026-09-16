"""Lanzador de App Alertas (punto de entrada del ejecutable).

Se mantiene minimo a proposito: prepara la salida estandar, importa la
aplicacion y deja por escrito cualquier fallo de arranque, incluidos los que
ocurren al importar los modulos. Empaquetado con --windowed no hay consola, asi
que sin esto un error de importacion solo mostraria un cuadro de dialogo vacio.
"""
from __future__ import annotations

import io
import sys
import traceback
from pathlib import Path


class _Silencio(io.TextIOBase):
    """Sustituto de stdout/stderr cuando no hay consola."""

    def write(self, texto: str) -> int:
        return len(texto)

    def flush(self) -> None:
        return None

    def isatty(self) -> bool:
        return False


if sys.stdout is None:
    sys.stdout = _Silencio()
if sys.stderr is None:
    sys.stderr = _Silencio()

sys.path.insert(0, str(Path(__file__).resolve().parent))


def _carpeta_app() -> Path:
    if getattr(sys, "frozen", False):
        return Path(sys.executable).resolve().parent
    return Path(__file__).resolve().parent


def _registrar_fallo(detalle: str) -> Path | None:
    try:
        carpeta = _carpeta_app() / "logs"
        carpeta.mkdir(parents=True, exist_ok=True)
        destino = carpeta / "arranque_error.log"
        destino.write_text(detalle, encoding="utf-8")
        return destino
    except Exception:                                           # noqa: BLE001
        return None


def _avisar(detalle: str, archivo: Path | None) -> None:
    mensaje = detalle[-2000:]
    if archivo:
        mensaje = f"Detalle completo en:\n{archivo}\n\n{mensaje}"
    try:
        from PySide6.QtWidgets import QApplication, QMessageBox

        if QApplication.instance() is None:
            QApplication(sys.argv)
        QMessageBox.critical(None, "App Alertas no pudo iniciarse", mensaje)
    except Exception:                                           # noqa: BLE001
        try:
            import ctypes

            ctypes.windll.user32.MessageBoxW(
                None, mensaje, "App Alertas no pudo iniciarse", 0x10
            )
        except Exception:                                       # noqa: BLE001
            pass


def main() -> int:
    try:
        from app.main import main as arrancar
    except BaseException:                                       # noqa: BLE001
        detalle = traceback.format_exc()
        _avisar(detalle, _registrar_fallo(detalle))
        return 1

    try:
        return arrancar()
    except SystemExit as salida:
        return int(salida.code or 0)
    except BaseException:                                       # noqa: BLE001
        detalle = traceback.format_exc()
        _avisar(detalle, _registrar_fallo(detalle))
        return 1


if __name__ == "__main__":
    raise SystemExit(main())
