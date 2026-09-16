"""Punto de entrada de App Alertas."""
from __future__ import annotations

import io
import sys
from pathlib import Path


class _Silencio(io.TextIOBase):
    """Sustituto de stdout/stderr cuando no hay consola."""

    def write(self, texto: str) -> int:                         # noqa: D102
        return len(texto)

    def flush(self) -> None:                                    # noqa: D102
        return None

    def isatty(self) -> bool:                                   # noqa: D102
        return False


# Empaquetada con --windowed, PyInstaller deja stdout y stderr en None y
# cualquier print() o aviso de una libreria aborta el arranque.
if sys.stdout is None:
    sys.stdout = _Silencio()
if sys.stderr is None:
    sys.stderr = _Silencio()

# Permite ejecutar "python app/main.py" ademas de "python -m app.main".
if __package__ in (None, ""):
    sys.path.insert(0, str(Path(__file__).resolve().parent.parent))

import traceback                                                # noqa: E402
import warnings                                                 # noqa: E402

import matplotlib                                               # noqa: E402

matplotlib.use("Agg")
warnings.filterwarnings("ignore", module="openpyxl")

from PySide6.QtCore import Qt                                   # noqa: E402
from PySide6.QtGui import QIcon                                 # noqa: E402
from PySide6.QtWidgets import QApplication, QMessageBox         # noqa: E402

from app import APP_NOMBRE, __version__                         # noqa: E402
from app.config import configurar_log, ruta_recurso             # noqa: E402
from app.ui.main_window import VentanaPrincipal                 # noqa: E402


def _manejar_excepcion(tipo, valor, rastro) -> None:
    detalle = "".join(traceback.format_exception(tipo, valor, rastro))
    configurar_log().error("Excepción no controlada:\n%s", detalle)
    QMessageBox.critical(
        None, "Error inesperado",
        f"{valor}\n\nEl detalle quedó registrado en logs/app.log.",
    )


def main() -> int:
    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.PassThrough
    )
    aplicacion = QApplication(sys.argv)
    aplicacion.setApplicationName(APP_NOMBRE)
    aplicacion.setApplicationVersion(__version__)
    aplicacion.setOrganizationName("Lara Consulting")

    icono = ruta_recurso("app.ico")
    if icono is not None:
        aplicacion.setWindowIcon(QIcon(str(icono)))

    sys.excepthook = _manejar_excepcion
    configurar_log().info("%s %s iniciada", APP_NOMBRE, __version__)

    ventana = VentanaPrincipal()
    ventana.show()
    return aplicacion.exec()


def _arrancar() -> int:
    """Arranque protegido: cualquier fallo queda por escrito y a la vista."""
    try:
        return main()
    except Exception:                                           # noqa: BLE001
        detalle = traceback.format_exc()
        try:
            configurar_log().error("Fallo al arrancar:\n%s", detalle)
        except Exception:                                       # noqa: BLE001
            pass
        try:
            from PySide6.QtWidgets import QApplication as _App
            from PySide6.QtWidgets import QMessageBox as _Msg

            if _App.instance() is None:
                _App(sys.argv)
            _Msg.critical(None, "App Alertas no pudo iniciarse", detalle[-1500:])
        except Exception:                                       # noqa: BLE001
            pass
        return 1


if __name__ == "__main__":
    raise SystemExit(_arrancar())
