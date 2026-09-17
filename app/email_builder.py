"""Construccion del correo: asunto, cuerpo HTML y alternativa en texto plano.

El HTML se escribe al estilo "Outlook": tablas con atributos y estilos en linea,
sin flex ni grid. Es lo que mejor renderiza Outlook y, a la vez, lo unico que el
editor de texto enriquecido de Qt sabe editar sin romper el formato.

La imagen de la Curva S se referencia siempre como `cid:curva_s`, tanto en el
editor (donde se registra como recurso del documento) como en el correo enviado.
"""
from __future__ import annotations

import datetime as dt
import html
import re
from string import Formatter

from .constantes import COLUMNAS_CORREO

CID_GRAFICO = "curva_s"
CID_FIRMA = "firma"
FUENTE = "Calibri, Arial, sans-serif"

# Ancho con el que se MUESTRA la Curva S en el correo, que no tiene nada que ver
# con la resolucion del PNG. "100%" hace que ocupe lo mismo que la tabla de
# entregables, que tambien va al 100%.
ANCHO_IMAGEN = "100%"

_FORMATOS_FECHA = [
    ("yyyy", "%Y"), ("yy", "%y"),
    ("mmmm", "%B"), ("mmm", "%b"), ("mm", "%m"),
    ("dddd", "%A"), ("ddd", "%a"), ("dd", "%d"),
]


class _FormateadorTolerante(Formatter):
    """Deja intactos los comodines que no existan en el contexto."""

    def get_value(self, key, args, kwargs):
        if isinstance(key, str):
            return kwargs.get(key, "{" + key + "}")
        return super().get_value(key, args, kwargs)

    def format_field(self, value, format_spec):
        try:
            return super().format_field(value, format_spec)
        except (ValueError, TypeError):
            return str(value)


_FORMATEADOR = _FormateadorTolerante()


def rellenar(plantilla: str, contexto: dict) -> str:
    if not plantilla:
        return ""
    try:
        return _FORMATEADOR.vformat(plantilla, (), contexto)
    except (IndexError, KeyError, ValueError):
        return plantilla


def formato_fecha(patron: str) -> str:
    """Convierte 'dd/mm/yyyy' al formato de strftime."""
    resultado = patron or "dd/mm/yyyy"
    if "%" in resultado:
        return resultado
    for origen, destino in _FORMATOS_FECHA:
        resultado = re.sub(origen, destino, resultado, flags=re.IGNORECASE)
    return resultado


def contexto(bd, datos, entregables: list) -> dict[str, str]:
    """Valores disponibles para los comodines de la plantilla."""
    decimales = bd.cfg_int("Decimales avance", 0)
    kpis = datos.ev.kpis(decimales)
    hoy = dt.date.today().strftime(formato_fecha(bd.txt("Formato fecha asunto")))
    return {
        "fecha_hoy": hoy,
        "codigo_proyecto": datos.proyecto.codigo,
        "nombre_proyecto": datos.nombre,
        "cliente": datos.cliente,
        "n_entregables": str(len(entregables)),
        **kpis,
    }


def construir_asunto(bd, datos, entregables: list) -> str:
    return rellenar(bd.txt("Asunto"), contexto(bd, datos, entregables)).strip()


# --------------------------------------------------------------------------- #
def _parrafo(texto: str) -> str:
    if not texto or not texto.strip():
        return ""
    cuerpo = texto.replace("\n", "<br>")
    return f'<p style="margin:0 0 10px 0;">{cuerpo}</p>'


def _tabla(entregables: list, color_encabezado: str, color_estatus: str) -> str:
    anchos = ["25%", "12%", "21%", "9%", "11%", "8%", "14%"]
    filas = [
        '<table border="1" cellspacing="0" cellpadding="5" width="100%" '
        'style="border-collapse:collapse; border:1px solid #000000; '
        f'font-family:{FUENTE}; font-size:9.5pt;">'
    ]

    filas.append("<tr>")
    for i, (clave, titulo) in enumerate(COLUMNAS_CORREO):
        fondo = color_estatus if clave == "estatus" else color_encabezado
        filas.append(
            f'<td width="{anchos[i]}" style="background-color:{fondo}; '
            'border:1px solid #000000; text-align:center; vertical-align:middle; '
            'padding:6px;"><b>' + html.escape(titulo) + "</b></td>"
        )
    filas.append("</tr>")

    for entregable in entregables:
        filas.append("<tr>")
        for i, (clave, _) in enumerate(COLUMNAS_CORREO):
            alineacion = "left" if i == 0 else "center"
            valor = html.escape(str(entregable.valor(clave) or ""))
            filas.append(
                f'<td style="border:1px solid #000000; text-align:{alineacion}; '
                'vertical-align:middle; padding:5px;">' + valor + "</td>"
            )
        filas.append("</tr>")

    filas.append("</table>")
    return "".join(filas)


def construir_html(bd, datos, entregables: list, incluir_grafico: bool = True,
                   incluir_firma: bool = False) -> str:
    ctx = contexto(bd, datos, entregables)
    color = bd.cfg("Color encabezado tabla", "#F8827F")
    color_estatus = bd.cfg("Color encabezado estatus", "#F2F2F2")

    partes = [
        f'<html><body style="font-family:{FUENTE}; font-size:11pt; color:#000000;">',
        _parrafo(rellenar(bd.txt("Saludo"), ctx)),
        _parrafo(rellenar(bd.txt("Párrafo intro"), ctx)),
        _parrafo(rellenar(bd.txt("Párrafo Curva S"), ctx)),
    ]

    # Los cuatro indicadores ya salen dentro del grafico, asi que por defecto no
    # se repiten como lista. Quien los prefiera tambien en texto lo activa en Config.
    if bd.cfg_bool("Mostrar indicadores en el texto", False):
        partes += [
            '<p style="margin:14px 0 6px 0;"><b><i><u>'
            + html.escape(rellenar(bd.txt("Título Avance"), ctx))
            + "</u></i></b></p>",
            '<ul style="margin:0 0 12px 0;">',
            f'<li><b>Avance Planificado:</b> {ctx["avance_planificado"]}</li>',
            f'<li><b>Avance Real:</b> {ctx["avance_real"]}</li>',
            f'<li><b>Desviación:</b> {ctx["desviacion"]}</li>',
            f'<li><b>SPI:</b> {ctx["spi"]}</li>',
            "</ul>",
        ]

    if incluir_grafico:
        ancho = bd.cfg("Ancho imagen en el correo") or ANCHO_IMAGEN
        partes.append(
            f'<p style="margin:0 0 14px 0;"><img src="cid:{CID_GRAFICO}" '
            f'width="{ancho}" alt="Curva S del proyecto"></p>'
        )

    partes += [
        _parrafo(rellenar(bd.txt("Párrafo entregables 1"), ctx)),
        _parrafo(rellenar(bd.txt("Párrafo entregables 2"), ctx)),
        _tabla(entregables, color, color_estatus) if entregables else _parrafo(
            "<i>No hay entregables pendientes con el estatus seleccionado.</i>"
        ),
        '<p style="margin:14px 0 0 0;">'
        + html.escape(rellenar(bd.txt("Cierre"), ctx)) + "</p>",
    ]

    if incluir_firma:
        ancho_firma = bd.cfg_int("Ancho firma px", 330)
        partes.append(
            f'<p style="margin:10px 0 0 0;"><img src="cid:{CID_FIRMA}" '
            f'width="{ancho_firma}" alt="Firma"></p>'
        )

    firma = rellenar(bd.txt("Firma"), ctx)
    if firma.strip():
        partes.append(_parrafo(firma))

    partes.append("</body></html>")
    return "".join(p for p in partes if p)


# --------------------------------------------------------------------------- #
_RE_BLOQUE = re.compile(r"</(p|tr|div|h[1-6]|li|table|ul|ol)>", re.IGNORECASE)
_RE_CELDA = re.compile(r"</t[dh]>", re.IGNORECASE)
_RE_SALTO = re.compile(r"<br\s*/?>", re.IGNORECASE)
_RE_ETIQUETA = re.compile(r"<[^>]+>")


def a_texto_plano(html_correo: str) -> str:
    """Version legible sin HTML, para el multipart/alternative."""
    texto = _RE_SALTO.sub("\n", html_correo)
    texto = _RE_CELDA.sub("\t", texto)
    texto = _RE_BLOQUE.sub("\n", texto)
    texto = _RE_ETIQUETA.sub("", texto)
    texto = html.unescape(texto)
    texto = re.sub(r"[ \t]+\n", "\n", texto)
    texto = re.sub(r"\n{3,}", "\n\n", texto)
    return texto.strip()


# --------------------------------------------------------------------------- #
# Ancho de la Curva S: el correo y la vista previa necesitan valores distintos
#
# Outlook renderiza `width="100%"` sin problema, pero el motor de texto de Qt no
# entiende porcentajes en imagenes (devuelve un ancho de -2 px y colapsa el
# documento). Por eso la vista previa recibe pixeles y se restaura el valor
# configurado justo antes de enviar.
# --------------------------------------------------------------------------- #
_RE_IMG = re.compile(r"<img\b[^>]*>", re.IGNORECASE)
_RE_ATRIBUTO_WIDTH = re.compile(r'\swidth\s*=\s*"([^"]*)"', re.IGNORECASE)
_RE_ATRIBUTO_HEIGHT = re.compile(r'\sheight\s*=\s*"[^"]*"', re.IGNORECASE)
_RE_ATRIBUTO_STYLE = re.compile(r'\sstyle\s*=\s*"([^"]*)"', re.IGNORECASE)
_RE_DIMENSION_EN_STYLE = re.compile(
    r"(?:max-|min-)?(?:width|height)\s*:[^;]*;?", re.IGNORECASE
)


def _limpiar_estilo(coincidencia: re.Match) -> str:
    """Quita del `style` cualquier ancho o alto: manda el atributo width."""
    cuerpo = _RE_DIMENSION_EN_STYLE.sub("", coincidencia.group(1))
    cuerpo = re.sub(r"\s*;\s*", "; ", cuerpo).strip().strip(";").strip()
    return f' style="{cuerpo}"' if cuerpo else ""


def _fijar_ancho(etiqueta: str, ancho: str) -> str:
    """Deja la etiqueta <img> con ese ancho y sin alto, para no deformarla.

    Qt reescribe la etiqueta al guardar el documento (reordena los atributos, la
    cierra con /> y anade un `height` fijo); si ese alto sobreviviera, Outlook
    dibujaria la imagen al 100% de ancho pero con la altura antigua.
    """
    nueva = _RE_ATRIBUTO_WIDTH.sub("", etiqueta)
    nueva = _RE_ATRIBUTO_HEIGHT.sub("", nueva)
    nueva = _RE_ATRIBUTO_STYLE.sub(_limpiar_estilo, nueva)
    return re.sub(r"^<img\b", f'<img width="{ancho}"', nueva, count=1, flags=re.IGNORECASE)


def _ancho_actual(etiqueta: str) -> str:
    encontrado = _RE_ATRIBUTO_WIDTH.search(etiqueta)
    return encontrado.group(1).strip() if encontrado else ""


def _reescribir_grafico(html: str, funcion) -> str:
    """Aplica `funcion` solo a la etiqueta <img> de la Curva S."""
    def reemplazo(coincidencia: re.Match) -> str:
        etiqueta = coincidencia.group(0)
        if f"cid:{CID_GRAFICO}" not in etiqueta:
            return etiqueta
        return funcion(etiqueta)

    return _RE_IMG.sub(reemplazo, html)


def ancho_para_vista(html: str, ancho_px: int) -> str:
    """Porcentaje -> pixeles, porque Qt no sabe renderizar porcentajes.

    Si el ancho configurado ya esta en pixeles no hay nada que convertir.
    """
    def convertir(etiqueta: str) -> str:
        if not _ancho_actual(etiqueta).endswith("%"):
            return etiqueta
        return _fijar_ancho(etiqueta, str(max(int(ancho_px), 200)))

    return _reescribir_grafico(html, convertir)


def ancho_para_correo(html: str, ancho: str = ANCHO_IMAGEN) -> str:
    """Devuelve a la Curva S el ancho configurado, antes de enviar o guardar."""
    ancho = (ancho or ANCHO_IMAGEN).strip() or ANCHO_IMAGEN
    return _reescribir_grafico(html, lambda etiqueta: _fijar_ancho(etiqueta, ancho))


# --------------------------------------------------------------------------- #
# Limpieza del HTML que devuelve el editor de Qt
# --------------------------------------------------------------------------- #
_RE_DOCTYPE = re.compile(r"<!DOCTYPE[^>]*>", re.IGNORECASE)
_RE_META_QT = re.compile(r'<meta\s+name="qrichtext"[^>]*/?>', re.IGNORECASE)
_RE_PROP_QT = re.compile(r"\s*-qt-[a-z-]+\s*:[^;\"']*;?")
_RE_ESTILO = re.compile(r'style="\s*([^"]*?)\s*"')
_RE_ESTILO_VACIO = re.compile(r'\sstyle="\s*"')


def limpiar_html_editor(html_qt: str) -> str:
    """Quita el ruido propio de Qt para aligerar el correo."""
    texto = _RE_DOCTYPE.sub("", html_qt)
    texto = _RE_META_QT.sub("", texto)
    texto = _RE_PROP_QT.sub("", texto)

    def compactar(coincidencia):
        cuerpo = re.sub(r"\s+", " ", coincidencia.group(1))
        cuerpo = re.sub(r"\s*;\s*", "; ", cuerpo).strip().strip(";")
        return f'style="{cuerpo}"' if cuerpo else ""

    texto = _RE_ESTILO.sub(compactar, texto)
    texto = _RE_ESTILO_VACIO.sub("", texto)
    return texto.strip()
