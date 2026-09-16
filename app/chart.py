"""Curva S del proyecto: replica del grafico del Excel de control.

Un unico eje X de semanas, eje Y izquierdo para las lineas acumuladas y eje Y
derecho para las barras semanales, con dos bandas inferiores (semanas y meses).
Devuelve la imagen en PNG lista para incrustar en el correo.
"""
from __future__ import annotations

import io
from dataclasses import dataclass

import matplotlib

matplotlib.use("Agg")

from matplotlib.figure import Figure          # noqa: E402
from matplotlib.lines import Line2D           # noqa: E402
from matplotlib.patches import Patch, Rectangle  # noqa: E402

# Colores tomados del grafico original
NEGRO = "#1A1A1A"
VERDE = "#4EA72E"
AZUL = "#2F5BE8"
GRIS_BARRA = "#BFBFBF"
GRIS_REJILLA = "#D9D9D9"
GRIS_TEXTO = "#595959"
GRIS_BANDA = "#F2F2F2"
BORDE_BANDA = "#BFBFBF"

FUENTES = ["Segoe UI", "Calibri", "DejaVu Sans", "sans-serif"]

# Espacio que se reserva a la derecha del area de trazado, en pixeles.
# Valores medidos sobre la figura real a 8 pt, con algo de holgura.
ANCHO_LEYENDA_PX = 150      # "% Tendencia Acum" mide 141 px
ANCHO_ETIQUETAS_PX = 34     # las etiquetas del eje derecho ocupan 27 px
AIRE_PX = 16                # separacion entre esas etiquetas y la leyenda
MARGEN_DERECHO_PX = 14


@dataclass
class OpcionesGrafico:
    ancho_px: int = 1400
    alto_px: int = 520
    dpi: int = 110
    izq_min: float = 0.0
    izq_max: float = 120.0
    der_min: float = 0.0
    der_max: float = 60.0
    etiquetas: bool = True

    @classmethod
    def desde_bd(cls, bd) -> "OpcionesGrafico":
        return cls(
            ancho_px=bd.cfg_int("Ancho gráfico px", 1400),
            alto_px=bd.cfg_int("Alto gráfico px", 520),
            dpi=bd.cfg_int("DPI", 110),
            izq_min=bd.cfg_float("Eje Y izq mín", 0.0),
            izq_max=bd.cfg_float("Eje Y izq máx", 120.0),
            der_min=bd.cfg_float("Eje Y der mín", 0.0),
            der_max=bd.cfg_float("Eje Y der máx", 60.0),
            etiquetas=bd.cfg_bool("Mostrar etiquetas de datos", True),
        )


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
           color_texto: str, negrita: bool = False) -> None:
    for etiqueta, ini, fin in bloques:
        ax.add_patch(
            Rectangle(
                (ini, 0.05), fin - ini, 0.9,
                facecolor=GRIS_BANDA, edgecolor=BORDE_BANDA, linewidth=0.6,
            )
        )
        ax.text(
            (ini + fin) / 2, 0.5, etiqueta,
            ha="center", va="center", fontsize=tam_letra, color=color_texto,
            fontweight="bold" if negrita else "normal",
        )


def generar(datos_ev, opciones: OpcionesGrafico | None = None) -> bytes:
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
    fig.patch.set_edgecolor("#808080")
    fig.patch.set_linewidth(1.0)
    matplotlib.rcParams["font.family"] = FUENTES

    # El hueco de la derecha se calcula en pixeles y no como fraccion fija: asi
    # la leyenda nunca pisa las etiquetas del eje derecho, sea cual sea el
    # "Ancho grafico px" configurado. Hay que hacerlo antes de crear la rejilla
    # porque gs.update() no reposiciona los ejes ya creados.
    reserva = ANCHO_ETIQUETAS_PX + AIRE_PX + ANCHO_LEYENDA_PX + MARGEN_DERECHO_PX
    derecha = min(max(1 - reserva / opciones.ancho_px, 0.55), 0.90)
    ancla_leyenda = 1 - (ANCHO_LEYENDA_PX + MARGEN_DERECHO_PX) / opciones.ancho_px

    alto_banda = 0.055 if n <= 40 else 0.05
    gs = fig.add_gridspec(
        3, 1, height_ratios=[1, alto_banda, alto_banda], hspace=0.0,
        left=0.055, right=derecha, top=0.965, bottom=0.04,
    )
    ax = fig.add_subplot(gs[0, 0])
    ax2 = ax.twinx()

    limites = (-0.6, n - 0.4)
    ax.set_xlim(*limites)
    ax.set_ylim(opciones.izq_min, opciones.izq_max)
    ax2.set_ylim(opciones.der_min, opciones.der_max)

    # --- barras semanales (eje derecho) --------------------------------- #
    indices = list(range(n))
    ancho = 0.26 if n <= 30 else 0.28
    separacion = ancho + 0.02          # deja aire entre barras contiguas
    barras = [
        ([x - separacion for x in indices], previsto, GRIS_BARRA, "% Previsto"),
        (indices, real, VERDE, "% Real"),
        ([x + separacion for x in indices], tendencia, AZUL, "% Tendencia"),
    ]
    for posiciones, valores, color, _ in barras:
        ax2.bar(posiciones, valores, width=ancho, color=color, zorder=1,
                edgecolor="white", linewidth=0.5)

    # --- lineas acumuladas (eje izquierdo) ------------------------------ #
    ax.set_zorder(ax2.get_zorder() + 1)
    ax.patch.set_visible(False)

    tam_marcador = 6 if n <= 30 else 5
    ax.plot(indices, previsto_ac, color=NEGRO, linewidth=1.6, marker="o",
            markersize=tam_marcador, markerfacecolor="white",
            markeredgecolor=NEGRO, markeredgewidth=1.2, zorder=6)
    ax.plot(indices, real_ac, color=VERDE, linewidth=2.0, marker="o",
            markersize=tam_marcador, markerfacecolor=VERDE,
            markeredgecolor=VERDE, zorder=7)
    ax.plot(indices, tendencia_ac, color=AZUL, linewidth=1.6, linestyle=(0, (4, 3)),
            marker="D", markersize=tam_marcador - 1, markerfacecolor="white",
            markeredgecolor=AZUL, markeredgewidth=1.2, zorder=5)

    # --- ejes ------------------------------------------------------------ #
    ax.set_axisbelow(True)
    ax.grid(axis="y", color=GRIS_REJILLA, linewidth=0.7, zorder=0)
    ax.set_xticks([])
    ax.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTO, length=0)
    ax2.tick_params(axis="y", labelsize=8, colors=GRIS_TEXTO, length=0)
    ax.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    ax2.yaxis.set_major_formatter(lambda v, _: f"{v:.0f}%")
    for lado in ("top", "right", "bottom"):
        ax.spines[lado].set_visible(False)
        ax2.spines[lado].set_visible(False)
    ax.spines["left"].set_color(GRIS_REJILLA)
    ax2.spines["left"].set_visible(False)

    # --- etiquetas de datos --------------------------------------------- #
    if opciones.etiquetas:
        tam = 7.0 if n <= 26 else 6.0
        piso = opciones.izq_min + (opciones.izq_max - opciones.izq_min) * 0.07

        def etiquetar(valores, color, arriba: bool, omitir_si_igual=None):
            for i, v in enumerate(valores):
                if v != v:                                   # nan
                    continue
                if omitir_si_igual is not None:
                    otro = omitir_si_igual[i]
                    if otro == otro and abs(otro - v) < 0.6:  # duplicaria la cifra
                        continue
                # nunca por debajo de la linea cuando no hay espacio libre
                hacia_arriba = arriba or v < piso
                ax.annotate(
                    f"{v:.1f}%", (i, v), textcoords="offset points",
                    xytext=(0, 11 if hacia_arriba else -13), ha="center",
                    fontsize=tam, color=color, fontweight="bold", zorder=8,
                )

        etiquetar(previsto_ac, NEGRO, arriba=True)
        etiquetar(tendencia_ac, AZUL, arriba=True, omitir_si_igual=previsto_ac)
        etiquetar(real_ac, VERDE, arriba=False, omitir_si_igual=previsto_ac)

        tam_barra = 5.5 if n <= 30 else 4.8
        for posiciones, valores, color, _ in barras:
            for x, v in zip(posiciones, valores):
                if v != v or v < 1.5:            # cifras ilegibles en barras minimas
                    continue
                ax2.annotate(
                    f"{v:.1f}%", (x, v), textcoords="offset points",
                    xytext=(0, 2), ha="center", va="bottom", rotation=90,
                    fontsize=tam_barra, color=GRIS_TEXTO, zorder=3,
                )

    # --- bandas de semanas y meses --------------------------------------- #
    ax_semanas = _banda(fig, gs, 1, limites)
    _cajas(
        ax_semanas,
        [(etiqueta, i - 0.5, i + 0.5) for i, etiqueta in enumerate(datos_ev.semanas)],
        tam_letra=7.0 if n <= 30 else 5.8,
        color_texto=GRIS_TEXTO,
    )

    ax_meses = _banda(fig, gs, 2, limites)
    bloques = datos_ev.meses or [("", -0.5, n - 0.5)]
    _cajas(
        ax_meses,
        [(etiqueta, ini - 0.5, fin + 0.5) for etiqueta, ini, fin in bloques],
        tam_letra=7.5 if n <= 30 else 6.5,
        color_texto=GRIS_TEXTO,
    )

    # --- leyenda ---------------------------------------------------------- #
    elementos = [
        Patch(facecolor=GRIS_BARRA, label="% Previsto"),
        Patch(facecolor=VERDE, label="% Real"),
        Patch(facecolor=AZUL, label="% Tendencia"),
        Line2D([], [], color=NEGRO, linewidth=1.6, marker="o", markersize=6,
               markerfacecolor="white", label="% Previsto Acum"),
        Line2D([], [], color=VERDE, linewidth=2.0, marker="o", markersize=6,
               label="% Real Acum"),
        Line2D([], [], color=AZUL, linewidth=1.6, linestyle=(0, (4, 3)), marker="D",
               markersize=5, markerfacecolor="white", label="% Tendencia Acum"),
    ]
    # Se ancla el borde izquierdo de la leyenda: crezca lo que crezca el texto,
    # nunca invade la zona de las etiquetas del eje derecho.
    fig.legend(
        handles=elementos, loc="center left", bbox_to_anchor=(ancla_leyenda, 0.55),
        frameon=False, fontsize=8, labelcolor=GRIS_TEXTO, handlelength=1.8,
    )

    buffer = io.BytesIO()
    fig.savefig(buffer, format="png", facecolor="white", edgecolor="#808080")
    return buffer.getvalue()
