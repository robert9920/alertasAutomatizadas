"""Curva S del proyecto: replica del grafico del Excel de control.

Un unico eje X de semanas, eje Y izquierdo para las lineas acumuladas y eje Y
derecho para las barras semanales, con dos bandas inferiores (semanas y meses).
Devuelve la imagen en PNG lista para incrustar en el correo.

Tamanos y colores salen de la hoja Config; los tamanos que se dejan vacios se
ajustan solos al numero de semanas.
"""
from __future__ import annotations

import io
from dataclasses import dataclass

import matplotlib

matplotlib.use("Agg")

from matplotlib.figure import Figure          # noqa: E402
from matplotlib.lines import Line2D           # noqa: E402
from matplotlib.patches import FancyBboxPatch, Patch, Rectangle  # noqa: E402

from .constantes import (                      # noqa: E402
    COLOR_ALERTA,
    COLOR_BANDA_EJE_X,
    COLOR_BARRA_PREVISTO,
    COLOR_CORPORATIVO,
    COLOR_FONDO_INDICADOR,
    COLOR_LINEA_PREVISTO,
    COLOR_LINEA_REAL,
    COLOR_LINEA_TENDENCIA,
    COLOR_MARCO_GRAFICO,
    COLOR_OK,
    COLOR_TEXTO_EJE_X,
)

GRIS_REJILLA = "#D9D9D9"
# Los porcentajes de los ejes Y y la leyenda no siguen al color del eje X:
# si no, oscurecer las bandas los dejaria invisibles.
GRIS_TEXTO = "#595959"
BORDE_BANDA = "#BFBFBF"
NEGRO_ETIQUETA = "#000000"

FUENTES = ["Segoe UI", "Calibri", "DejaVu Sans", "sans-serif"]

# Espacio que se reserva a la derecha del area de trazado, en pixeles.
# Valores medidos sobre la figura real a 8 pt, con algo de holgura.
ANCHO_LEYENDA_PX = 150      # "% Tendencia Acum" mide 141 px
ANCHO_ETIQUETAS_PX = 34     # las etiquetas del eje derecho ocupan 27 px
AIRE_PX = 16                # separacion entre esas etiquetas y la leyenda
MARGEN_DERECHO_PX = 14

IZQUIERDA = 0.055

# Reparto vertical de la decoracion, en pixeles. Al calcularlo asi, agrandar la
# imagen o la letra reparte el espacio sin descuadrar nada.
MARGEN_FIGURA_PX = 10
HUECO_TITULO_PX = 14
HUECO_INDICADORES_PX = 16
MARGEN_INFERIOR_PX = 10
RADIO_ESQUINA_PX = 9
SEPARACION_TARJETAS_PX = 12
ALTO_ACENTO_PX = 4

# Umbral del SPI a partir del cual el indicador se considera en objetivo.
SPI_OBJETIVO = 0.95

# Dos cifras que difieran menos que esto se consideran la misma y se dibuja una.
EPSILON_IGUAL = 0.05
# Por debajo de esta fraccion del eje izquierdo no cabe una etiqueta "abajo":
# se saldria del area de trazado y pisaria la banda de semanas.
FRACCION_PISO = 0.07


@dataclass
class OpcionesGrafico:
    # -- tamano de la imagen y ejes ---------------------------------------- #
    ancho_px: int = 1400
    alto_px: int = 520
    dpi: int = 110
    izq_min: float = 0.0
    izq_max: float = 120.0
    der_min: float = 0.0
    der_max: float = 60.0
    etiquetas: bool = True

    # -- tamanos de letra (None = se ajusta al numero de semanas) ---------- #
    tam_etiquetas_lineas: float | None = None
    tam_etiquetas_barras: float | None = None
    tam_eje_izq: float = 8.0
    tam_eje_der: float = 8.0
    tam_semanas: float | None = None
    tam_meses: float | None = None
    tam_leyenda: float = 8.0
    tam_titulo: float = 16.0
    tam_subtitulo: float = 10.0
    tam_kpi_titulo: float = 9.0
    tam_kpi_valor: float = 20.0

    # -- cabecera e indicadores -------------------------------------------- #
    subtitulo: str = "CURVA S DE AVANCE DEL PROYECTO"
    color_fondo_titulo: str = COLOR_CORPORATIVO
    mostrar_indicadores: bool = True
    color_fondo_indicador: str = COLOR_FONDO_INDICADOR
    color_marco: str = COLOR_MARCO_GRAFICO
    decimales: int = 0

    # -- colores de las series --------------------------------------------- #
    color_linea_previsto: str = COLOR_LINEA_PREVISTO
    color_linea_real: str = COLOR_LINEA_REAL
    color_linea_tendencia: str = COLOR_LINEA_TENDENCIA
    color_barra_previsto: str = COLOR_BARRA_PREVISTO
    color_barra_real: str = COLOR_LINEA_REAL
    color_barra_tendencia: str = COLOR_LINEA_TENDENCIA

    # -- colores de las etiquetas (None = heredado) ------------------------ #
    color_etq_linea_previsto: str | None = None
    color_etq_linea_real: str | None = None
    color_etq_linea_tendencia: str | None = None
    color_etq_barra_previsto: str | None = None
    color_etq_barra_real: str | None = None
    color_etq_barra_tendencia: str | None = None

    # -- colores de los valores de los indicadores (None = automatico) ------ #
    color_kpi_planificado: str | None = None
    color_kpi_real: str | None = None
    color_kpi_desviacion: str | None = None
    color_kpi_spi: str | None = None

    # -- eje X y titulo ----------------------------------------------------- #
    color_banda: str = COLOR_BANDA_EJE_X
    color_texto_banda: str = COLOR_TEXTO_EJE_X
    color_titulo: str = "#FFFFFF"

    @classmethod
    def desde_bd(cls, bd) -> "OpcionesGrafico":
        def color(clave: str, defecto: str) -> str:
            return bd.cfg_opcional(clave) or defecto

        return cls(
            ancho_px=bd.cfg_int("Ancho gráfico px", 1400),
            alto_px=bd.cfg_int("Alto gráfico px", 700),
            dpi=bd.cfg_int("DPI", 110),
            izq_min=bd.cfg_float("Eje Y izq mín", 0.0),
            izq_max=bd.cfg_float("Eje Y izq máx", 120.0),
            der_min=bd.cfg_float("Eje Y der mín", 0.0),
            der_max=bd.cfg_float("Eje Y der máx", 60.0),
            etiquetas=bd.cfg_bool("Mostrar etiquetas de datos", True),

            tam_etiquetas_lineas=bd.cfg_float_opcional("Tamaño etiquetas líneas"),
            tam_etiquetas_barras=bd.cfg_float_opcional("Tamaño etiquetas barras"),
            tam_eje_izq=bd.cfg_float("Tamaño eje Y izquierdo", 8.0),
            tam_eje_der=bd.cfg_float("Tamaño eje Y derecho", 8.0),
            tam_semanas=bd.cfg_float_opcional("Tamaño eje X semanas"),
            tam_meses=bd.cfg_float_opcional("Tamaño eje X meses"),
            tam_leyenda=bd.cfg_float("Tamaño leyenda", 8.0),
            tam_titulo=bd.cfg_float("Tamaño título gráfico", 16.0),
            tam_subtitulo=bd.cfg_float("Tamaño subtítulo gráfico", 10.0),
            tam_kpi_titulo=bd.cfg_float("Tamaño título indicadores", 9.0),
            tam_kpi_valor=bd.cfg_float("Tamaño valor indicadores", 20.0),

            subtitulo=bd.cfg("Subtítulo gráfico"),
            color_fondo_titulo=color("Color fondo título", COLOR_CORPORATIVO),
            mostrar_indicadores=bd.cfg_bool("Mostrar indicadores en gráfico", True),
            color_fondo_indicador=color("Color fondo indicadores", COLOR_FONDO_INDICADOR),
            color_marco=color("Color marco gráfico", COLOR_MARCO_GRAFICO),
            decimales=bd.cfg_int("Decimales avance", 0),

            color_linea_previsto=color("Color línea Previsto Acum", COLOR_LINEA_PREVISTO),
            color_linea_real=color("Color línea Real Acum", COLOR_LINEA_REAL),
            color_linea_tendencia=color("Color línea Tendencia Acum", COLOR_LINEA_TENDENCIA),
            color_barra_previsto=color("Color barra Previsto", COLOR_BARRA_PREVISTO),
            color_barra_real=color("Color barra Real", COLOR_LINEA_REAL),
            color_barra_tendencia=color("Color barra Tendencia", COLOR_LINEA_TENDENCIA),

            color_etq_linea_previsto=bd.cfg_opcional("Color etiqueta Previsto Acum"),
            color_etq_linea_real=bd.cfg_opcional("Color etiqueta Real Acum"),
            color_etq_linea_tendencia=bd.cfg_opcional("Color etiqueta Tendencia Acum"),
            color_etq_barra_previsto=bd.cfg_opcional("Color etiqueta barra Previsto"),
            color_etq_barra_real=bd.cfg_opcional("Color etiqueta barra Real"),
            color_etq_barra_tendencia=bd.cfg_opcional("Color etiqueta barra Tendencia"),

            color_kpi_planificado=bd.cfg_opcional("Color valor Avance Planificado"),
            color_kpi_real=bd.cfg_opcional("Color valor Avance Real"),
            color_kpi_desviacion=bd.cfg_opcional("Color valor Desviación"),
            color_kpi_spi=bd.cfg_opcional("Color valor SPI"),

            color_banda=color("Color bandas eje X", COLOR_BANDA_EJE_X),
            color_texto_banda=color("Color texto eje X", COLOR_TEXTO_EJE_X),
            color_titulo=color("Color título gráfico", "#FFFFFF"),
        )

    # -- alturas de la decoracion, en pixeles ------------------------------ #
    def alto_cabecera(self, con_titulo: bool) -> float:
        if not con_titulo:
            return 0.0
        alto = self.tam_titulo * 2.0 + 18
        if self.subtitulo.strip():
            alto += self.tam_subtitulo * 1.9
        return alto

    def alto_indicadores(self) -> float:
        if not self.mostrar_indicadores:
            return 0.0
        return self.tam_kpi_titulo * 1.8 + self.tam_kpi_valor * 1.9 + 34

    # -- tamanos resueltos segun el numero de semanas ---------------------- #
    def letra_etiquetas_lineas(self, n: int) -> float:
        return self.tam_etiquetas_lineas or (7.0 if n <= 26 else 6.0)

    def letra_etiquetas_barras(self, n: int) -> float:
        return self.tam_etiquetas_barras or (5.5 if n <= 30 else 4.8)

    def letra_semanas(self, n: int) -> float:
        return self.tam_semanas or (7.0 if n <= 30 else 5.8)

    def letra_meses(self, n: int) -> float:
        return self.tam_meses or (7.5 if n <= 30 else 6.5)


# --------------------------------------------------------------------------- #
# Posicion de las etiquetas
# --------------------------------------------------------------------------- #
def posiciones_etiquetas(
    previsto: list[float], real: list[float], tendencia: list[float]
) -> tuple[list[str | None], list[str | None], list[str | None]]:
    """Decide, semana a semana, de que lado va la etiqueta de cada linea.

    Devuelve "arriba", "abajo" o None (no dibujarla) para previsto, real y
    tendencia. La referencia de cada semana es el % Real Acum y, cuando este ya
    no tiene datos, el % Tendencia Acum: el previsto siempre queda al lado
    contrario de su referencia. Si dos cifras coinciden se dibuja una sola, y la
    primera etiqueta de la tendencia se omite porque repite el ultimo valor real.
    """
    n = len(previsto)
    pos_p: list[str | None] = [None] * n
    pos_r: list[str | None] = [None] * n
    pos_t: list[str | None] = [None] * n

    for i in range(n):
        p, r, t = previsto[i], real[i], tendencia[i]
        hay_p, hay_r, hay_t = p == p, r == r, t == t       # nan != nan

        if hay_r:
            pos_r[i] = "arriba" if (not hay_p or r > p) else "abajo"
        if hay_t:
            pos_t[i] = "arriba" if (not hay_p or t > p) else "abajo"

        if hay_p:
            # una cifra practicamente igual a la del previsto se omite
            if hay_r and abs(r - p) < EPSILON_IGUAL:
                pos_r[i] = None
            if hay_t and abs(t - p) < EPSILON_IGUAL:
                pos_t[i] = None

            lado_referencia = pos_r[i] if hay_r else (pos_t[i] if hay_t else None)
            pos_p[i] = "abajo" if lado_referencia == "arriba" else "arriba"

    primera_tendencia = next((i for i, v in enumerate(tendencia) if v == v), None)
    if primera_tendencia is not None:
        pos_t[primera_tendencia] = None

    return pos_p, pos_r, pos_t


# --------------------------------------------------------------------------- #
def _a_porcentaje(serie: list[float | None], n: int) -> list[float]:
    salida: list[float] = []
    for i in range(n):
        v = serie[i] if i < len(serie) else None
        salida.append(float("nan") if v is None else v * 100.0)
    return salida


def _banda(fig, gs, indice: int, limites: tuple[float, float]):
    ax = fig.add_subplot(gs[indice, 0])
    ax.set_xlim(*limites)
    ax.set_ylim(0, 1)
    ax.set_xticks([])
    ax.set_yticks([])
    for lado in ax.spines.values():
        lado.set_visible(False)
    return ax


def _cajas(ax, bloques: list[tuple[str, float, float]], tam_letra: float,
           color_fondo: str, color_texto: str) -> None:
    for etiqueta, ini, fin in bloques:
        ax.add_patch(
            Rectangle(
                (ini, 0.05), fin - ini, 0.9,
                facecolor=color_fondo, edgecolor=BORDE_BANDA, linewidth=0.6,
            )
        )
        ax.text(
            (ini + fin) / 2, 0.5, etiqueta,
            ha="center", va="center", fontsize=tam_letra, color=color_texto,
        )


def _caja(lienzo, x, y, ancho, alto, relleno, borde=None, radio=RADIO_ESQUINA_PX,
          grosor=1.0, zorder=1):
    """Rectangulo de esquinas redondeadas en coordenadas de pixel."""
    caja = FancyBboxPatch(
        (x + radio, y + radio), ancho - 2 * radio, alto - 2 * radio,
        boxstyle=f"round,pad={radio}",
        facecolor=relleno, edgecolor=borde or "none",
        linewidth=grosor if borde else 0, zorder=zorder,
    )
    lienzo.add_patch(caja)
    return caja


def _dibujar_cabecera(lienzo, opciones: OpcionesGrafico, titulo: str,
                      arriba_px: float, alto_px: float) -> None:
    """Banda superior con el nombre del proyecto y el subtitulo."""
    izquierda = MARGEN_FIGURA_PX
    ancho = opciones.ancho_px - 2 * MARGEN_FIGURA_PX
    _caja(lienzo, izquierda, arriba_px - alto_px, ancho, alto_px,
          relleno=opciones.color_fondo_titulo, zorder=2)

    x_texto = izquierda + 22
    subtitulo = opciones.subtitulo.strip()
    if subtitulo:
        lienzo.text(
            x_texto, arriba_px - alto_px * 0.38, titulo,
            ha="left", va="center", fontsize=opciones.tam_titulo,
            color=opciones.color_titulo, fontweight="bold", zorder=3,
        )
        lienzo.text(
            x_texto, arriba_px - alto_px * 0.74, subtitulo,
            ha="left", va="center", fontsize=opciones.tam_subtitulo,
            color=opciones.color_titulo, alpha=0.92, zorder=3,
        )
    else:
        lienzo.text(
            x_texto, arriba_px - alto_px / 2, titulo,
            ha="left", va="center", fontsize=opciones.tam_titulo,
            color=opciones.color_titulo, fontweight="bold", zorder=3,
        )


def _dibujar_indicadores(lienzo, opciones: OpcionesGrafico,
                         tarjetas: list[tuple[str, str, str, str]],
                         abajo_px: float, alto_px: float) -> None:
    """Fila inferior de indicadores, a todo el ancho del grafico."""
    izquierda = MARGEN_FIGURA_PX
    total = opciones.ancho_px - 2 * MARGEN_FIGURA_PX
    cantidad = len(tarjetas)
    ancho = (total - SEPARACION_TARJETAS_PX * (cantidad - 1)) / cantidad

    for i, (rotulo, valor, pie, color) in enumerate(tarjetas):
        x = izquierda + i * (ancho + SEPARACION_TARJETAS_PX)
        _caja(lienzo, x, abajo_px, ancho, alto_px,
              relleno=opciones.color_fondo_indicador,
              borde=opciones.color_marco, radio=RADIO_ESQUINA_PX, zorder=2)
        # franja de color inferior, para identificar el indicador de un vistazo
        lienzo.add_patch(Rectangle(
            (x + RADIO_ESQUINA_PX, abajo_px + 2),
            ancho - 2 * RADIO_ESQUINA_PX, ALTO_ACENTO_PX,
            facecolor=color, edgecolor="none", zorder=3,
        ))

        centro = x + ancho / 2
        lienzo.text(centro, abajo_px + alto_px - opciones.tam_kpi_titulo * 1.5,
                    rotulo, ha="center", va="center",
                    fontsize=opciones.tam_kpi_titulo, color=GRIS_TEXTO,
                    fontweight="bold", zorder=4)
        lienzo.text(centro, abajo_px + alto_px * 0.44, valor,
                    ha="center", va="center", fontsize=opciones.tam_kpi_valor,
                    color=color, fontweight="bold", zorder=4)
        lienzo.text(centro, abajo_px + ALTO_ACENTO_PX + 12, pie,
                    ha="center", va="center",
                    fontsize=max(opciones.tam_kpi_titulo - 1.5, 5.5),
                    color=GRIS_TEXTO, alpha=0.85, zorder=4)


def _tarjetas_indicadores(datos_ev, opciones: OpcionesGrafico
                          ) -> list[tuple[str, str, str, str]]:
    """(rotulo, valor, pie, color) de los cuatro indicadores del proyecto.

    El color se calcula como siempre -las dos primeras tarjetas heredan el de su
    linea de la Curva S y las otras dos son verdes o rojas segun el resultado- y
    solo se sustituye si en Config hay un color escrito para esa tarjeta.
    """
    kpis = datos_ev.kpis(opciones.decimales)
    semana = datos_ev.semana_corte
    pie_semana = f"(Semana {semana[1:]})" if semana.upper().startswith("S") else f"({semana})"

    desviacion = kpis["desviacion"]
    color_desviacion = COLOR_ALERTA if desviacion.startswith("-") else COLOR_OK
    try:
        color_spi = COLOR_OK if float(kpis["spi"]) >= SPI_OBJETIVO else COLOR_ALERTA
    except ValueError:
        color_spi = GRIS_TEXTO

    return [
        ("AVANCE PLANIFICADO", kpis["avance_planificado"], pie_semana,
         opciones.color_kpi_planificado or opciones.color_linea_previsto),
        ("AVANCE REAL", kpis["avance_real"], pie_semana,
         opciones.color_kpi_real or opciones.color_linea_real),
        ("DESVIACIÓN", desviacion, "(Real vs. Planificado)",
         opciones.color_kpi_desviacion or color_desviacion),
        ("SPI", kpis["spi"], "(Índice de programación)",
         opciones.color_kpi_spi or color_spi),
    ]


def generar(datos_ev, opciones: OpcionesGrafico | None = None,
            titulo: str = "") -> bytes:
    """Devuelve el PNG de la Curva S."""
    opciones = opciones or OpcionesGrafico()
    n = len(datos_ev.semanas)
    if n == 0:
        raise ValueError("La hoja EV no tiene semanas para graficar.")

    previsto = _a_porcentaje(datos_ev.series.get("previsto", []), n)
    real = _a_porcentaje(datos_ev.series.get("real", []), n)
    tendencia = _a_porcentaje(datos_ev.series.get("tendencia", []), n)
    previsto_ac = _a_porcentaje(datos_ev.series.get("previsto_acum", []), n)
    real_ac = _a_porcentaje(datos_ev.series.get("real_acum", []), n)
    tendencia_ac = _a_porcentaje(datos_ev.series.get("tendencia_acum", []), n)

    fig = Figure(
        figsize=(opciones.ancho_px / opciones.dpi, opciones.alto_px / opciones.dpi),
        dpi=opciones.dpi,
        facecolor="white",
    )
    matplotlib.rcParams["font.family"] = FUENTES

    # Capa de fondo en pixeles: el marco, la cabecera y las tarjetas se dibujan
    # aqui. Queda por debajo del area de trazado porque `ax` no pinta su fondo y
    # es `ax2` quien tapa esta capa solo dentro de la curva.
    lienzo = fig.add_axes((0, 0, 1, 1), zorder=-1)
    lienzo.set_xlim(0, opciones.ancho_px)
    lienzo.set_ylim(0, opciones.alto_px)
    lienzo.axis("off")
    _caja(lienzo, 1, 1, opciones.ancho_px - 2, opciones.alto_px - 2,
          relleno="white", borde=opciones.color_marco, radio=RADIO_ESQUINA_PX + 3,
          grosor=1.4, zorder=0)

    # El hueco de la derecha y el del titulo se calculan en pixeles y no como
    # fraccion fija: asi la leyenda nunca pisa las etiquetas del eje derecho y el
    # area de trazado no se descuadra al cambiar el tamano de la imagen.
    reserva = ANCHO_ETIQUETAS_PX + AIRE_PX + ANCHO_LEYENDA_PX + MARGEN_DERECHO_PX
    derecha = min(max(1 - reserva / opciones.ancho_px, 0.55), 0.90)
    ancla_leyenda = 1 - (ANCHO_LEYENDA_PX + MARGEN_DERECHO_PX) / opciones.ancho_px

    # Reparto vertical, tambien en pixeles:
    #   margen · cabecera · hueco · [trazado + bandas] · hueco · indicadores · margen
    con_titulo = bool(titulo.strip())
    alto_cabecera = opciones.alto_cabecera(con_titulo)
    alto_kpis = opciones.alto_indicadores()

    ocupado_arriba = MARGEN_FIGURA_PX + alto_cabecera + (HUECO_TITULO_PX if con_titulo else 0)
    ocupado_abajo = MARGEN_INFERIOR_PX + alto_kpis + (
        HUECO_INDICADORES_PX if opciones.mostrar_indicadores else 0
    )
    arriba = min(max(1 - ocupado_arriba / opciones.alto_px, 0.45), 0.985)
    abajo = min(max(ocupado_abajo / opciones.alto_px, 0.02), 0.45)

    alto_banda = 0.055 if n <= 40 else 0.05
    gs = fig.add_gridspec(
        3, 1, height_ratios=[1, alto_banda, alto_banda], hspace=0.0,
        left=IZQUIERDA, right=derecha, top=arriba, bottom=abajo,
    )
    ax = fig.add_subplot(gs[0, 0])
    ax2 = ax.twinx()

    if con_titulo:
        _dibujar_cabecera(
            lienzo, opciones, titulo.strip(),
            arriba_px=opciones.alto_px - MARGEN_FIGURA_PX, alto_px=alto_cabecera,
        )
    if opciones.mostrar_indicadores:
        _dibujar_indicadores(
            lienzo, opciones, _tarjetas_indicadores(datos_ev, opciones),
            abajo_px=MARGEN_INFERIOR_PX, alto_px=alto_kpis,
        )

    limites = (-0.6, n - 0.4)
    ax.set_xlim(*limites)
    ax.set_ylim(opciones.izq_min, opciones.izq_max)
    ax2.set_ylim(opciones.der_min, opciones.der_max)

    # --- barras semanales (eje derecho) --------------------------------- #
    indices = list(range(n))
    ancho = 0.26 if n <= 30 else 0.28
    separacion = ancho + 0.02          # deja aire entre barras contiguas
    barras = [
        ([x - separacion for x in indices], previsto, opciones.color_barra_previsto,
         opciones.color_etq_barra_previsto or NEGRO_ETIQUETA, "% Previsto"),
        (indices, real, opciones.color_barra_real,
         opciones.color_etq_barra_real or NEGRO_ETIQUETA, "% Real"),
        ([x + separacion for x in indices], tendencia, opciones.color_barra_tendencia,
         opciones.color_etq_barra_tendencia or NEGRO_ETIQUETA, "% Tendencia"),
    ]
    for posiciones, valores, color, _, _ in barras:
        ax2.bar(posiciones, valores, width=ancho, color=color, zorder=1,
                edgecolor="white", linewidth=0.5)

    # --- lineas acumuladas (eje izquierdo) ------------------------------ #
    ax.set_zorder(ax2.get_zorder() + 1)
    ax.patch.set_visible(False)

    tam_marcador = 6 if n <= 30 else 5
    ax.plot(indices, previsto_ac, color=opciones.color_linea_previsto, linewidth=1.6,
            marker="o", markersize=tam_marcador, markerfacecolor="white",
            markeredgecolor=opciones.color_linea_previsto, markeredgewidth=1.2, zorder=6)
    ax.plot(indices, real_ac, color=opciones.color_linea_real, linewidth=2.0,
            marker="o", markersize=tam_marcador, markerfacecolor=opciones.color_linea_real,
            markeredgecolor=opciones.color_linea_real, zorder=7)
    ax.plot(indices, tendencia_ac, color=opciones.color_linea_tendencia, linewidth=1.6,
            linestyle=(0, (4, 3)), marker="D", markersize=tam_marcador - 1,
            markerfacecolor="white", markeredgecolor=opciones.color_linea_tendencia,
            markeredgewidth=1.2, zorder=5)

    # --- ejes ------------------------------------------------------------ #
    ax.set_axisbelow(True)
    ax.grid(axis="y", color=GRIS_REJILLA, linewidth=0.7, zorder=0)
    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=opciones.tam_eje_izq, colors=GRIS_TEXTO, length=0)
    ax2.tick_params(axis="y", labelsize=opciones.tam_eje_der, colors=GRIS_TEXTO, length=0)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax2.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    for lado in ("top", "right", "bottom"):
        ax.spines[lado].set_visible(False)
        ax2.spines[lado].set_visible(False)
    ax.spines["left"].set_color(GRIS_REJILLA)
    ax2.spines["left"].set_visible(False)

    # --- etiquetas de datos --------------------------------------------- #
    if opciones.etiquetas:
        tam = opciones.letra_etiquetas_lineas(n)
        piso = opciones.izq_min + (opciones.izq_max - opciones.izq_min) * FRACCION_PISO
        pos_p, pos_r, pos_t = posiciones_etiquetas(previsto_ac, real_ac, tendencia_ac)

        series = (
            (previsto_ac, pos_p,
             opciones.color_etq_linea_previsto or opciones.color_linea_previsto),
            (real_ac, pos_r,
             opciones.color_etq_linea_real or opciones.color_linea_real),
            (tendencia_ac, pos_t,
             opciones.color_etq_linea_tendencia or opciones.color_linea_tendencia),
        )
        for valores, posicion, color in series:
            for i, v in enumerate(valores):
                if posicion[i] is None or v != v:
                    continue
                # abajo del todo no cabe: se saldria del area de trazado
                arriba_final = posicion[i] == "arriba" or v < piso
                ax.annotate(
                    f"{v:.1f}%", (i, v), textcoords="offset points",
                    xytext=(0, 11 if arriba_final else -13), ha="center",
                    fontsize=tam, color=color, fontweight="bold", zorder=8,
                )

        tam_barra = opciones.letra_etiquetas_barras(n)
        for posiciones, valores, _, color_etq, _ in barras:
            for x, v in zip(posiciones, valores):
                if v != v or v < 1.5:            # cifras ilegibles en barras minimas
                    continue
                ax2.annotate(
                    f"{v:.1f}%", (x, v), textcoords="offset points",
                    xytext=(0, 2), ha="center", va="bottom", rotation=90,
                    fontsize=tam_barra, color=color_etq, zorder=3,
                )

    # --- bandas de semanas y meses --------------------------------------- #
    ax_semanas = _banda(fig, gs, 1, limites)
    _cajas(
        ax_semanas,
        [(etiqueta, i - 0.5, i + 0.5) for i, etiqueta in enumerate(datos_ev.semanas)],
        tam_letra=opciones.letra_semanas(n),
        color_fondo=opciones.color_banda,
        color_texto=opciones.color_texto_banda,
    )

    ax_meses = _banda(fig, gs, 2, limites)
    bloques = datos_ev.meses or [("", -0.5, n - 0.5)]
    _cajas(
        ax_meses,
        [(etiqueta, ini - 0.5, fin + 0.5) for etiqueta, ini, fin in bloques],
        tam_letra=opciones.letra_meses(n),
        color_fondo=opciones.color_banda,
        color_texto=opciones.color_texto_banda,
    )

    # --- leyenda ---------------------------------------------------------- #
    elementos = [
        Patch(facecolor=opciones.color_barra_previsto, label="% Previsto"),
        Patch(facecolor=opciones.color_barra_real, label="% Real"),
        Patch(facecolor=opciones.color_barra_tendencia, label="% Tendencia"),
        Line2D([], [], color=opciones.color_linea_previsto, linewidth=1.6, marker="o",
               markersize=6, markerfacecolor="white", label="% Previsto Acum"),
        Line2D([], [], color=opciones.color_linea_real, linewidth=2.0, marker="o",
               markersize=6, label="% Real Acum"),
        Line2D([], [], color=opciones.color_linea_tendencia, linewidth=1.6,
               linestyle=(0, (4, 3)), marker="D", markersize=5,
               markerfacecolor="white", label="% Tendencia Acum"),
    ]
    # Se ancla el borde izquierdo de la leyenda: crezca lo que crezca el texto,
    # nunca invade la zona de las etiquetas del eje derecho.
    fig.legend(
        handles=elementos, loc="center left",
        bbox_to_anchor=(ancla_leyenda, (arriba + abajo) / 2),
        frameon=False, fontsize=opciones.tam_leyenda,
        labelcolor=GRIS_TEXTO, handlelength=1.8,
    )

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", facecolor="white", edgecolor="none")
    return buffer.getvalue()
