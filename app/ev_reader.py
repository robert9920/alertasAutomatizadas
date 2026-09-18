"""Lectura de la hoja EV: series de avance, semana de corte, KPIs y SPI."""
from __future__ import annotations

import datetime as dt
import re
from dataclasses import dataclass, field

from .constantes import (
    CAMPOS_EV,
    CELDA_FECHA_INICIO,
    CELDA_SPI,
    COL_INICIO_SEMANAS,
    ETIQUETA_FECHA_INICIO,
)
from .excel_utils import (
    a_celda,
    a_float,
    a_indice_columna,
    a_numero_fila,
    buscar_etiqueta,
    buscar_hoja,
    es_etiqueta_semana,
    letra_columna,
    mapa_combinadas,
    normalizar,
    primer_numero_a_la_derecha,
    texto_limpio,
    valor,
)

MAX_FILA_ETIQUETAS = 90
MAX_COL_ETIQUETAS = 6
MAX_COL_SEMANAS = 200
SERIES = ("previsto", "previsto_acum", "real", "real_acum", "tendencia", "tendencia_acum")


@dataclass
class DatosEV:
    hoja: str = ""
    filas: dict[str, int] = field(default_factory=dict)
    origen: dict[str, str] = field(default_factory=dict)
    col_inicio: int = 0
    col_fin: int = 0
    semanas: list[str] = field(default_factory=list)
    meses: list[tuple[str, int, int]] = field(default_factory=list)
    series: dict[str, list[float | None]] = field(default_factory=dict)
    idx_corte: int = -1
    spi: float | None = None
    celda_spi: str = ""
    fecha_inicio: dt.date | None = None
    fecha_fin: dt.date | None = None
    avisos: list[str] = field(default_factory=list)

    # -- derivados -------------------------------------------------------- #
    @property
    def semana_corte(self) -> str:
        if 0 <= self.idx_corte < len(self.semanas):
            return self.semanas[self.idx_corte]
        return ""

    @property
    def numero_semana(self) -> int:
        m = re.search(r"\d+", self.semana_corte)
        if m:
            return int(m.group(0))
        return self.idx_corte + 1 if self.idx_corte >= 0 else 0

    def _en_corte(self, serie: str) -> float | None:
        datos = self.series.get(serie) or []
        if 0 <= self.idx_corte < len(datos):
            return datos[self.idx_corte]
        return None

    @property
    def avance_planificado(self) -> float | None:
        return self._en_corte("previsto_acum")

    @property
    def avance_real(self) -> float | None:
        return self._en_corte("real_acum")

    def kpis(self, decimales: int = 0) -> dict[str, str]:
        """Textos ya formateados. La desviacion se calcula sobre los valores
        redondeados para que los tres numeros del correo sean consistentes."""
        def pct(v):
            return None if v is None else round(v * 100, decimales)

        plan, real = pct(self.avance_planificado), pct(self.avance_real)
        desv = None if (plan is None or real is None) else round(real - plan, decimales)

        def fmt(v, signo=False):
            if v is None:
                return "-"
            if decimales <= 0:
                texto = f"{int(round(v)):d}"
            else:
                texto = f"{v:.{decimales}f}"
            if signo and v > 0:
                texto = "+" + texto
            return texto + "%"

        return {
            "avance_planificado": fmt(plan),
            "avance_real": fmt(real),
            "desviacion": fmt(desv, signo=True),
            "spi": "-" if self.spi is None else f"{self.spi:.2f}",
            "semana": str(self.numero_semana),
        }

    def sin_datos(self) -> bool:
        return not any(
            v is not None for serie in SERIES for v in (self.series.get(serie) or [])
        )

    def diagnostico(self) -> list[tuple[str, str, str]]:
        filas = [
            ("Hoja", self.hoja, self.origen.get("hoja", "")),
            ("Columna inicio semanas",
             letra_columna(self.col_inicio) if self.col_inicio else "-",
             self.origen.get("col_inicio", "")),
        ]
        for clave, titulo, _ in CAMPOS_EV:
            fila = self.filas.get(clave)
            filas.append((titulo, str(fila) if fila else "-", self.origen.get(clave, "")))
        filas.append(("SPI", self.celda_spi or "-", self.origen.get("spi", "")))
        return filas


# --------------------------------------------------------------------------- #
def _resolver_filas(ws, mapeo: dict, datos: DatosEV, combinadas: dict) -> None:
    for clave, etiqueta, respaldo in CAMPOS_EV:
        fila = a_numero_fila(mapeo.get(clave))
        if fila:
            datos.filas[clave] = fila
            datos.origen[clave] = "BD"
            continue
        encontrado = buscar_etiqueta(
            ws, etiqueta,
            range(1, min(MAX_FILA_ETIQUETAS, ws.max_row or MAX_FILA_ETIQUETAS) + 1),
            range(1, MAX_COL_ETIQUETAS + 1),
            combinadas,
        )
        if encontrado:
            datos.filas[clave] = encontrado[0]
            datos.origen[clave] = "busqueda por texto"
            continue
        datos.filas[clave] = respaldo
        datos.origen[clave] = "respaldo"
        datos.avisos.append(
            f"No se encontro la etiqueta '{etiqueta}' en {datos.hoja}; "
            f"se usa la fila de respaldo {respaldo}."
        )


def _resolver_columna_inicio(ws, mapeo: dict, datos: DatosEV, combinadas: dict) -> int:
    col = a_indice_columna(mapeo.get("col_inicio"))
    if col:
        datos.origen["col_inicio"] = "BD"
        return col
    fila_semana = datos.filas.get("semana")
    if fila_semana:
        for candidata in range(1, MAX_COL_SEMANAS + 1):
            if es_etiqueta_semana(valor(ws, fila_semana, candidata, combinadas)):
                datos.origen["col_inicio"] = "busqueda por texto"
                return candidata
    datos.origen["col_inicio"] = "respaldo"
    return a_indice_columna(COL_INICIO_SEMANAS) or 3


def _leer_semanas(ws, datos: DatosEV, combinadas: dict) -> tuple[list[str], int]:
    fila_semana = datos.filas.get("semana", 0)
    etiquetas: list[str] = []
    col_fin = datos.col_inicio - 1
    vacias = 0
    for col in range(datos.col_inicio, MAX_COL_SEMANAS + 1):
        bruto = valor(ws, fila_semana, col, combinadas)
        texto = texto_limpio(bruto)
        if texto:
            etiquetas.append(texto if not texto.isdigit() else f"S{texto}")
            col_fin = col
            vacias = 0
        else:
            vacias += 1
            if vacias >= 3:
                break
            etiquetas.append("")
    while etiquetas and not etiquetas[-1]:
        etiquetas.pop()
    return etiquetas, col_fin


def _leer_serie(ws, fila: int, col_ini: int, cantidad: int, combinadas: dict) -> list[float | None]:
    serie: list[float | None] = []
    for i in range(cantidad):
        serie.append(a_float(valor(ws, fila, col_ini + i, combinadas)) if fila else None)
    return serie


def _escalar(series: dict[str, list[float | None]]) -> None:
    """Si el Excel guarda 58.6 en vez de 0.586, lo pasa a fraccion."""
    maximo = 0.0
    for valores in series.values():
        for v in valores:
            if v is not None:
                maximo = max(maximo, abs(v))
    if maximo > 1.5:
        for clave, valores in series.items():
            series[clave] = [None if v is None else v / 100.0 for v in valores]


def _leer_meses(ws, datos: DatosEV, combinadas: dict) -> list[tuple[str, int, int]]:
    fila_mes = datos.filas.get("mes")
    if not fila_mes:
        return []
    bloques: list[tuple[str, int, int]] = []
    actual: str | None = None
    ini = 0
    for i in range(len(datos.semanas)):
        etiqueta = texto_limpio(valor(ws, fila_mes, datos.col_inicio + i, combinadas))
        if etiqueta and etiqueta != actual:
            if actual:
                bloques.append((actual, ini, i - 1))
            actual, ini = etiqueta, i
        elif not etiqueta and actual is None:
            continue
    if actual:
        bloques.append((actual, ini, len(datos.semanas) - 1))
    return bloques


def _resolver_spi(ws, mapeo: dict, datos: DatosEV, combinadas: dict) -> None:
    referencia = a_celda(mapeo.get("spi"))
    if referencia:
        datos.spi = a_float(valor(ws, referencia[0], referencia[1], combinadas))
        datos.celda_spi = f"{letra_columna(referencia[1])}{referencia[0]}"
        datos.origen["spi"] = "BD"
        if datos.spi is not None:
            return

    tope = min(MAX_FILA_ETIQUETAS + 30, (ws.max_row or 120))
    candidatos: list[tuple[int, int]] = []
    for fila in range(1, tope + 1):
        for col in range(1, 70):
            if normalizar(valor(ws, fila, col, combinadas)).rstrip(" :") == "spi":
                candidatos.append((fila, col))
    for fila, col in candidatos:
        v = primer_numero_a_la_derecha(ws, fila, col, 6, combinadas)
        if v is not None:
            datos.spi = v
            datos.celda_spi = f"cerca de {letra_columna(col)}{fila}"
            datos.origen["spi"] = "busqueda por texto"
            return

    referencia = a_celda(CELDA_SPI)
    if referencia:
        datos.spi = a_float(valor(ws, referencia[0], referencia[1], combinadas))
        datos.celda_spi = CELDA_SPI
        datos.origen["spi"] = "respaldo"
    if datos.spi is None:
        datos.avisos.append("No se pudo determinar el SPI en la hoja EV.")


def _a_fecha(bruto) -> dt.date | None:
    if isinstance(bruto, dt.datetime):
        return bruto.date()
    if isinstance(bruto, dt.date):
        return bruto
    return None


def _resolver_fechas(ws, datos: DatosEV, combinadas: dict) -> None:
    """Fecha de inicio (al lado de su etiqueta) y fecha de la ultima semana.

    La de inicio esta a la derecha de «Fecha de corte:», en una celda que suele
    ir combinada; `valor()` ya devuelve el contenido de la celda maestra. La de
    fin sale de la fila FECHA, en la ultima columna que se grafica, que es la que
    `leer()` deja en `col_fin` tras recortar las semanas sin previsto.
    """
    encontrado = buscar_etiqueta(
        ws, ETIQUETA_FECHA_INICIO,
        range(1, min(MAX_FILA_ETIQUETAS, ws.max_row or MAX_FILA_ETIQUETAS) + 1),
        range(1, MAX_COL_ETIQUETAS + 1),
        combinadas,
    )
    if encontrado:
        fila, col = encontrado
        for salto in range(1, 4):           # la etiqueta puede ocupar dos celdas
            datos.fecha_inicio = _a_fecha(valor(ws, fila, col + salto, combinadas))
            if datos.fecha_inicio:
                break
    if datos.fecha_inicio is None:
        referencia = a_celda(CELDA_FECHA_INICIO)
        if referencia:
            datos.fecha_inicio = _a_fecha(
                valor(ws, referencia[0], referencia[1], combinadas)
            )

    fila_fecha = datos.filas.get("fecha")
    if fila_fecha and datos.col_fin >= datos.col_inicio:
        datos.fecha_fin = _a_fecha(valor(ws, fila_fecha, datos.col_fin, combinadas))


# --------------------------------------------------------------------------- #
def leer(wb, proyecto) -> DatosEV:
    """Extrae las series de avance de la hoja EV del libro ya abierto."""
    datos = DatosEV()
    nombre_hoja = buscar_hoja(wb, "EV", proyecto.hoja_ev)
    datos.hoja = nombre_hoja
    datos.origen["hoja"] = "BD" if proyecto.hoja_ev else "busqueda por nombre"
    ws = wb[nombre_hoja]
    combinadas = mapa_combinadas(ws)

    mapeo = proyecto.mapeo_ev or {}
    _resolver_filas(ws, mapeo, datos, combinadas)
    datos.col_inicio = _resolver_columna_inicio(ws, mapeo, datos, combinadas)

    semanas, col_fin = _leer_semanas(ws, datos, combinadas)
    datos.semanas = semanas
    datos.col_fin = col_fin

    cantidad = len(semanas)
    for clave in SERIES:
        datos.series[clave] = _leer_serie(
            ws, datos.filas.get(clave, 0), datos.col_inicio, cantidad, combinadas
        )
    _escalar(datos.series)

    # Recorta las semanas sin ningun dato planificado al final del horizonte.
    referencia = datos.series.get("previsto_acum") or []
    alternativa = datos.series.get("tendencia_acum") or []
    ultimo = -1
    for i in range(cantidad):
        tiene = (i < len(referencia) and referencia[i] is not None) or (
            i < len(alternativa) and alternativa[i] is not None
        )
        if tiene:
            ultimo = i
    if ultimo >= 0 and ultimo < cantidad - 1:
        datos.semanas = datos.semanas[: ultimo + 1]
        for clave in SERIES:
            datos.series[clave] = datos.series[clave][: ultimo + 1]
        datos.col_fin = datos.col_inicio + ultimo

    datos.meses = _leer_meses(ws, datos, combinadas)
    datos.idx_corte = _semana_de_corte(datos)
    _resolver_spi(ws, mapeo, datos, combinadas)
    _resolver_fechas(ws, datos, combinadas)

    if datos.sin_datos():
        datos.avisos.append(
            "La hoja EV no devolvio valores calculados. Abre el archivo en Excel, "
            "guardalo y vuelve a actualizar."
        )
    elif datos.idx_corte < 0:
        datos.avisos.append(
            "No hay ninguna semana con % Previsto y % Real a la vez; "
            "no se pudo determinar la semana de corte."
        )
    return datos


def _semana_de_corte(datos: DatosEV) -> int:
    """Ultima semana con valor en previsto Y real (acumulados de preferencia)."""
    for plan, real in (("previsto_acum", "real_acum"), ("previsto", "real")):
        serie_plan = datos.series.get(plan) or []
        serie_real = datos.series.get(real) or []
        indice = -1
        for i in range(min(len(serie_plan), len(serie_real))):
            if serie_plan[i] is not None and serie_real[i] is not None:
                indice = i
        if indice >= 0:
            return indice
    return -1
