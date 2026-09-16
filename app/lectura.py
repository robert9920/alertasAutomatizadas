"""Abre el Excel de control de un proyecto una sola vez y extrae todo."""
from __future__ import annotations

from dataclasses import dataclass, field

from . import ev_reader, let_reader
from .bd import Proyecto
from .excel_compat import abrir as abrir_libro
from .excel_utils import (
    ErrorExcel,
    buscar_hoja,
    mapa_combinadas,
    normalizar,
    texto_limpio,
    valor,
)


@dataclass
class DatosProyecto:
    proyecto: Proyecto
    let: let_reader.DatosLET
    ev: ev_reader.DatosEV
    nombre: str = ""
    cliente: str = ""
    avisos: list[str] = field(default_factory=list)

    def diagnostico(self) -> list[tuple[str, str, str]]:
        return self.let.diagnostico() + self.ev.diagnostico()


def _texto_caratula(wb, *etiquetas: str) -> str:
    """Busca una etiqueta en la hoja CARATULA y devuelve el texto a su derecha."""
    try:
        ws = wb[buscar_hoja(wb, "CARATULA")]
    except ErrorExcel:
        return ""
    combinadas = mapa_combinadas(ws)
    objetivos = [normalizar(e) for e in etiquetas]
    tope = min(80, ws.max_row or 80)
    for fila in range(1, tope + 1):
        for col in range(1, 20):
            texto = normalizar(valor(ws, fila, col, combinadas)).rstrip(" :")
            if not texto or texto not in objetivos:
                continue
            for desplazamiento in range(1, 6):
                candidato = texto_limpio(valor(ws, fila, col + desplazamiento, combinadas))
                if candidato and normalizar(candidato) not in objetivos:
                    return candidato
    return ""


def cargar(proyecto: Proyecto) -> DatosProyecto:
    ruta = proyecto.ruta_excel
    if not ruta.is_file():
        raise ErrorExcel(
            f"No se encuentra el Excel de entregables del proyecto {proyecto.codigo}:\n{ruta}"
        )
    wb = abrir_libro(ruta, data_only=True)
    try:
        datos_let = let_reader.leer(wb, proyecto)
        datos_ev = ev_reader.leer(wb, proyecto)
        nombre = proyecto.nombre or _texto_caratula(wb, "Nombre de Proyecto", "PROYECTO")
        cliente = proyecto.cliente or _texto_caratula(wb, "Cliente")
    finally:
        wb.close()

    resultado = DatosProyecto(
        proyecto=proyecto,
        let=datos_let,
        ev=datos_ev,
        nombre=nombre or proyecto.codigo,
        cliente=cliente,
    )
    resultado.avisos = list(datos_let.avisos) + list(datos_ev.avisos)
    return resultado
