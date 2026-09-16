"""Lectura de la hoja LET: listado de entregables y estatus disponibles."""
from __future__ import annotations

from collections import Counter
from dataclasses import dataclass, field

from .constantes import (
    CAMPOS_LET,
    FILA_ENCABEZADO_LET,
    FILA_INICIO_DATOS_LET,
)
from .excel_utils import (
    a_fecha_texto,
    a_indice_columna,
    a_numero_fila,
    buscar_en_fila,
    buscar_hoja,
    coincide,
    letra_columna,
    mapa_combinadas,
    normalizar,
    rango_tabla,
    texto_limpio,
    valor,
)

MAX_COL_BUSQUEDA = 200


@dataclass
class Entregable:
    fila: int
    nombre: str = ""
    disciplina: str = ""
    codigo_cliente: str = ""
    revision: str = ""
    fecha_envio: str = ""
    estatus: str = ""

    def valor(self, clave: str) -> str:
        return getattr(self, clave, "")


@dataclass
class DatosLET:
    hoja: str = ""
    fila_encabezado: int = 0
    fila_inicio: int = 0
    fila_fin: int = 0
    columnas: dict[str, int] = field(default_factory=dict)
    origen: dict[str, str] = field(default_factory=dict)
    entregables: list[Entregable] = field(default_factory=list)
    conteo_estatus: Counter = field(default_factory=Counter)
    avisos: list[str] = field(default_factory=list)

    @property
    def estatus_disponibles(self) -> list[str]:
        return [v for v, _ in self.conteo_estatus.most_common()]

    def diagnostico(self) -> list[tuple[str, str, str]]:
        """(campo, referencia usada, origen) para el panel de diagnostico."""
        filas = [
            ("Hoja", self.hoja, self.origen.get("hoja", "")),
            ("Fila encabezado", str(self.fila_encabezado), self.origen.get("fila_encabezado", "")),
            ("Fila inicio datos", str(self.fila_inicio), self.origen.get("fila_inicio", "")),
        ]
        for clave, titulo, _ in CAMPOS_LET:
            col = self.columnas.get(clave)
            filas.append(
                (titulo, letra_columna(col) if col else "-", self.origen.get(clave, ""))
            )
        return filas

    def filtrar(self, estatus: list[str]) -> list[Entregable]:
        objetivos = {normalizar(e) for e in estatus}
        if not objetivos:
            return []
        return [e for e in self.entregables if normalizar(e.estatus) in objetivos]

    def estatus_sugeridos(self, deseado: str) -> list[str]:
        """Estatus de la hoja que corresponden al filtro pedido (tolerante)."""
        if not deseado:
            return []
        exactos = [v for v in self.estatus_disponibles if normalizar(v) == normalizar(deseado)]
        if exactos:
            return exactos
        return [v for v in self.estatus_disponibles if coincide(v, deseado, umbral=0.88)]


# --------------------------------------------------------------------------- #
def _resolver_fila_encabezado(ws, mapeo: dict, datos: DatosLET) -> int:
    fila = a_numero_fila(mapeo.get("fila_encabezado"))
    if fila:
        datos.origen["fila_encabezado"] = "BD"
        return fila

    rango = rango_tabla(ws)
    if rango:
        datos.origen["fila_encabezado"] = "tabla de Excel"
        return rango[0]

    combinadas = mapa_combinadas(ws)
    mejor_fila, mejor_puntaje = 0, 0
    tope = min(25, ws.max_row or 25)
    for candidata in range(1, tope + 1):
        puntaje = 0
        for _, titulo, _ in CAMPOS_LET:
            if buscar_en_fila(ws, candidata, titulo, MAX_COL_BUSQUEDA, combinadas):
                puntaje += 1
        if puntaje > mejor_puntaje:
            mejor_fila, mejor_puntaje = candidata, puntaje
    if mejor_puntaje >= 3:
        datos.origen["fila_encabezado"] = "busqueda por texto"
        return mejor_fila

    datos.origen["fila_encabezado"] = "respaldo"
    return FILA_ENCABEZADO_LET


def _resolver_columnas(ws, mapeo: dict, fila_enc: int, datos: DatosLET) -> None:
    combinadas = mapa_combinadas(ws)
    for clave, titulo, respaldo in CAMPOS_LET:
        col = a_indice_columna(mapeo.get(clave))
        if col:
            datos.columnas[clave] = col
            datos.origen[clave] = "BD"
            continue
        col = buscar_en_fila(ws, fila_enc, titulo, MAX_COL_BUSQUEDA, combinadas)
        if col:
            datos.columnas[clave] = col
            datos.origen[clave] = "busqueda por texto"
            continue
        datos.columnas[clave] = a_indice_columna(respaldo)
        datos.origen[clave] = "respaldo"
        datos.avisos.append(
            f"No se encontro la columna '{titulo}' en la hoja {datos.hoja}; "
            f"se usa la posicion de respaldo {respaldo}."
        )


def _resolver_rango_datos(ws, mapeo: dict, fila_enc: int, datos: DatosLET) -> tuple[int, int]:
    fila_ini = a_numero_fila(mapeo.get("fila_inicio"))
    if fila_ini:
        datos.origen["fila_inicio"] = "BD"
    elif datos.origen.get("fila_encabezado") == "respaldo":
        fila_ini = FILA_INICIO_DATOS_LET
        datos.origen["fila_inicio"] = "respaldo"
    else:
        fila_ini = fila_enc + 1
        datos.origen["fila_inicio"] = "encabezado + 1"

    rango = rango_tabla(ws)
    if rango and rango[0] == fila_enc:
        return fila_ini, rango[2]

    col_ref = datos.columnas.get("nombre") or datos.columnas.get("codigo_cliente") or 1
    fila_fin = fila_ini - 1
    vacias = 0
    for fila in range(fila_ini, (ws.max_row or fila_ini) + 1):
        if texto_limpio(ws.cell(row=fila, column=col_ref).value):
            fila_fin = fila
            vacias = 0
        else:
            vacias += 1
            if vacias >= 40:
                break
    return fila_ini, max(fila_fin, fila_ini - 1)


# --------------------------------------------------------------------------- #
def leer(wb, proyecto) -> DatosLET:
    """Extrae los entregables de la hoja LET del libro ya abierto."""
    datos = DatosLET()
    nombre_hoja = buscar_hoja(wb, "LET", proyecto.hoja_let)
    datos.hoja = nombre_hoja
    datos.origen["hoja"] = "BD" if proyecto.hoja_let else "busqueda por nombre"
    ws = wb[nombre_hoja]

    mapeo = proyecto.mapeo_let or {}
    fila_enc = _resolver_fila_encabezado(ws, mapeo, datos)
    datos.fila_encabezado = fila_enc
    _resolver_columnas(ws, mapeo, fila_enc, datos)
    fila_ini, fila_fin = _resolver_rango_datos(ws, mapeo, fila_enc, datos)
    datos.fila_inicio, datos.fila_fin = fila_ini, fila_fin

    combinadas = mapa_combinadas(ws)
    col = datos.columnas
    for fila in range(fila_ini, fila_fin + 1):
        nombre = texto_limpio(valor(ws, fila, col.get("nombre", 0), combinadas))
        codigo = texto_limpio(valor(ws, fila, col.get("codigo_cliente", 0), combinadas))
        estatus = texto_limpio(valor(ws, fila, col.get("estatus", 0), combinadas))
        if not nombre and not codigo and not estatus:
            continue
        entregable = Entregable(
            fila=fila,
            nombre=nombre,
            disciplina=texto_limpio(valor(ws, fila, col.get("disciplina", 0), combinadas)),
            codigo_cliente=codigo,
            revision=texto_limpio(valor(ws, fila, col.get("revision", 0), combinadas)),
            fecha_envio=a_fecha_texto(valor(ws, fila, col.get("fecha_envio", 0), combinadas)),
            estatus=estatus,
        )
        datos.entregables.append(entregable)
        if estatus:
            datos.conteo_estatus[estatus] += 1

    if not datos.entregables:
        datos.avisos.append(
            f"La hoja {nombre_hoja} no devolvio filas entre {fila_ini} y {fila_fin}."
        )
    return datos
