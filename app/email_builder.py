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

# Aire entre el ultimo parrafo y la tabla de entregables. Se usa un parrafo
# espaciador y no un margen mayor porque Outlook (motor de Word) colapsa los
# margenes que preceden a una tabla con borde; el espaciador si lo respeta.
ESPACIO_ANTES_TABLA = '<p style="margin:0; font-size:6pt; line-height:8px;">&nbsp;</p>'

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


def _tabla(entregables: list, color_encabezado: str, color_estatus: str,
           texto_encabezado: str = "#000000", texto_estatus: str = "#000000") -> str:
    anchos = ["25%", "12%", "21%", "9%", "11%", "8%", "14%"]
    filas = [
        '<table border="1" cellspacing="0" cellpadding="5" width="100%" '
        'style="border-collapse:collapse; border:1px solid #000000; '
        f'font-family:{FUENTE}; font-size:9.5pt;">'
    ]

    filas.append("<tr>")
    for i, (clave, titulo) in enumerate(COLUMNAS_CORREO):
        es_estatus = clave == "estatus"
        fondo = color_estatus if es_estatus else color_encabezado
        letra = texto_estatus if es_estatus else texto_encabezado
        filas.append(
            f'<td width="{anchos[i]}" style="background-color:{fondo}; '
            f'color:{letra}; '
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
    texto_encabezado = bd.cfg("Color texto encabezado tabla", "#000000")
    texto_estatus = bd.cfg("Color texto encabezado estatus", "#000000")

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
        ancho = normalizar_ancho(bd.cfg("Ancho imagen en el correo"))
        partes.append(
            f'<p style="margin:0 0 14px 0;"><img src="cid:{CID_GRAFICO}" '
            f'width="{ancho}" alt="Curva S del proyecto"></p>'
        )

    partes += [
        _parrafo(rellenar(bd.txt("Párrafo entregables 1"), ctx)),
        _parrafo(rellenar(bd.txt("Párrafo entregables 2"), ctx)),
        ESPACIO_ANTES_TABLA,
        _tabla(entregables, color, color_estatus, texto_encabezado, texto_estatus)
        if entregables else _parrafo(
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
ANCHO_MIN_PCT, ANCHO_MAX_PCT = 10, 100
ANCHO_MIN_PX, ANCHO_MAX_PX = 200, 2400
# Por debajo de esto el numero es la fraccion que guarda Excel, no un ancho:
# al teclear «90%» la celda almacena 0,9.
FRACCION_MAXIMA = 1.5

_RE_NUMERO = re.compile(r"\d+(?:[.,]\d+)?")
_RE_MILES = re.compile(r"\d{1,3}[.,]\d{3}")


def _numero(bruto: str) -> float | None:
    """Primer numero del texto, tolerando «1.200», «80,5», «1200px» y «90 %»."""
    limpio = bruto.replace(" ", "").replace(" ", "")
    # El punto o la coma solo son separador de miles si detras van tres cifras.
    if _RE_MILES.fullmatch(limpio.replace("%", "").replace("px", "")):
        limpio = limpio.replace(".", "").replace(",", "")
    encontrado = _RE_NUMERO.search(limpio.replace(",", "."))
    if not encontrado:
        return None
    try:
        return float(encontrado.group(0))
    except ValueError:
        return None


def normalizar_ancho(texto: str | None) -> str:
    """Convierte lo que haya en Config a algo que el correo entienda.

    Devuelve un porcentaje ("90%") o un numero de pixeles ("1400"). Ante la duda
    manda el porcentaje: hasta 100 se entiende como porcentaje y por encima de
    100, como pixeles.

    El caso raro que hay que atender es el de Excel: al escribir «90%» en la
    celda no guarda ese texto, guarda el numero 0,9 con formato de porcentaje.
    Por eso un numero que no llegue a 1,5 se interpreta como fraccion.
    """
    bruto = str(texto or "").strip()
    if not bruto:
        return ANCHO_IMAGEN

    numero = _numero(bruto)
    if numero is None:
        return ANCHO_IMAGEN

    if "%" not in bruto and numero > ANCHO_MAX_PCT:
        return str(int(round(min(max(numero, ANCHO_MIN_PX), ANCHO_MAX_PX))))

    if "%" not in bruto and numero <= FRACCION_MAXIMA:
        numero *= 100.0
    return f"{int(round(min(max(numero, ANCHO_MIN_PCT), ANCHO_MAX_PCT)))}%"


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

    El porcentaje se aplica sobre el ancho util del editor, igual que hara el
    cliente de correo con el suyo. Si el ancho configurado ya esta en pixeles no
    hay nada que convertir.
    """
    def convertir(etiqueta: str) -> str:
        actual = _ancho_actual(etiqueta)
        if not actual.endswith("%"):
            return etiqueta
        fraccion = float(normalizar_ancho(actual).rstrip("%")) / 100.0
        return _fijar_ancho(etiqueta, str(max(int(ancho_px * fraccion), 200)))

    return _reescribir_grafico(html, convertir)


def ancho_para_correo(html: str, ancho: str = ANCHO_IMAGEN) -> str:
    """Devuelve a la Curva S el ancho configurado, antes de enviar o guardar."""
    normalizado = normalizar_ancho(ancho)
    return _reescribir_grafico(html, lambda etiqueta: _fijar_ancho(etiqueta, normalizado))


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
