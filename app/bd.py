"""Lectura y validacion del Excel de base de datos (BD_Reportes.xlsx)."""
from __future__ import annotations

from dataclasses import dataclass, field
from pathlib import Path

from .constantes import (
    CAMPOS_EV,
    CAMPOS_LET,
    CONFIG_DEFECTO,
    ESTADO_FILTRO_DEFECTO,
    PLANTILLA_DEFECTO,
    SMTP_DEFECTO,
    TIPOS_DESTINATARIO,
)
from .excel_compat import abrir as abrir_libro
from .excel_utils import ErrorExcel, a_float, buscar_hoja, normalizar, texto_limpio

VERDADEROS = {"si", "s", "true", "verdadero", "1", "x", "yes", "y"}


def es_verdadero(valor, por_defecto: bool = True) -> bool:
    if valor is None or str(valor).strip() == "":
        return por_defecto
    if isinstance(valor, bool):
        return valor
    return normalizar(valor) in VERDADEROS


# --------------------------------------------------------------------------- #
# Estructuras
# --------------------------------------------------------------------------- #
@dataclass
class Proyecto:
    codigo: str
    nombre: str = ""
    cliente: str = ""
    ruta: str = ""
    archivo: str = ""
    hoja_let: str = ""
    hoja_ev: str = ""
    estado_filtro: str = ""
    activo: bool = True
    mapeo_let: dict[str, str] = field(default_factory=dict)
    mapeo_ev: dict[str, str] = field(default_factory=dict)

    @property
    def ruta_excel(self) -> Path:
        nombre = (self.archivo or "").strip()
        if not nombre:
            return Path(self.ruta or "")
        if not nombre.lower().endswith((".xlsx", ".xlsm", ".xls")):
            nombre += ".xlsx"
        return Path(self.ruta or "") / nombre

    @property
    def etiqueta(self) -> str:
        return f"{self.codigo} - {self.nombre}" if self.nombre else self.codigo

    def existe_excel(self) -> bool:
        try:
            return self.ruta_excel.is_file()
        except OSError:
            return False


@dataclass
class Destinatario:
    codigo: str
    nombre: str = ""
    correo: str = ""
    tipo: str = "Para"
    activo: bool = True


@dataclass
class ConfigSMTP:
    servidor: str = ""
    puerto: int = 587
    seguridad: str = "STARTTLS"
    usuario: str = ""
    contrasena: str = ""
    remitente: str = ""
    nombre_remitente: str = ""
    responder_a: str = ""
    cco_fijo: str = ""
    timeout: int = 30

    @property
    def desde(self) -> str:
        return (self.remitente or self.usuario or "").strip()

    def problemas(self) -> list[str]:
        faltas = []
        if not self.servidor.strip():
            faltas.append("Servidor")
        if not self.puerto:
            faltas.append("Puerto")
        if not self.desde:
            faltas.append("Remitente o Usuario")
        return faltas

    @property
    def configurado(self) -> bool:
        return not self.problemas()


@dataclass
class BaseDatos:
    ruta: Path
    proyectos: list[Proyecto] = field(default_factory=list)
    destinatarios: list[Destinatario] = field(default_factory=list)
    smtp: ConfigSMTP = field(default_factory=ConfigSMTP)
    plantilla: dict[str, str] = field(default_factory=dict)
    config: dict[str, str] = field(default_factory=dict)
    avisos: list[str] = field(default_factory=list)

    def destinatarios_de(self, codigo: str) -> list[Destinatario]:
        objetivo = normalizar(codigo)
        return [d for d in self.destinatarios if normalizar(d.codigo) == objetivo]

    def proyecto(self, codigo: str) -> Proyecto | None:
        objetivo = normalizar(codigo)
        for p in self.proyectos:
            if normalizar(p.codigo) == objetivo:
                return p
        return None

    # -- accesos a Config con conversion ---------------------------------- #
    def cfg(self, clave: str, defecto: str = "") -> str:
        valor = self.config.get(normalizar(clave))
        if valor is None or str(valor).strip() == "":
            return CONFIG_DEFECTO.get(clave, defecto)
        return str(valor).strip()

    def cfg_float(self, clave: str, defecto: float) -> float:
        v = a_float(self.cfg(clave))
        return defecto if v is None else v

    def cfg_int(self, clave: str, defecto: int) -> int:
        v = a_float(self.cfg(clave))
        return defecto if v is None else int(v)

    def cfg_bool(self, clave: str, defecto: bool = True) -> bool:
        return es_verdadero(self.cfg(clave), defecto)

    def cfg_opcional(self, clave: str) -> str | None:
        """Texto del parámetro, o None si está vacío (significa «automático»)."""
        valor = self.cfg(clave).strip()
        return valor or None

    def cfg_float_opcional(self, clave: str) -> float | None:
        return a_float(self.cfg_opcional(clave))

    def txt(self, clave: str) -> str:
        valor = self.plantilla.get(normalizar(clave))
        if valor is None or str(valor).strip() == "":
            return PLANTILLA_DEFECTO.get(clave, "")
        return str(valor)


# --------------------------------------------------------------------------- #
# Lectura de hojas
# --------------------------------------------------------------------------- #
def _fila_encabezado(ws, maximo: int = 15) -> int:
    for fila in range(1, min(maximo, ws.max_row or 1) + 1):
        llenas = sum(
            1 for celda in ws[fila] if celda.value is not None and str(celda.value).strip()
        )
        if llenas >= 2:
            return fila
    return 1


def _leer_tabla(ws) -> list[dict[str, object]]:
    """Filas de una hoja tabular, con claves normalizadas por encabezado."""
    if not ws.max_row:
        return []
    fila_enc = _fila_encabezado(ws)
    encabezados: dict[int, str] = {}
    for celda in ws[fila_enc]:
        texto = normalizar(celda.value)
        if texto:
            encabezados[celda.column] = texto
    if not encabezados:
        return []

    filas: list[dict[str, object]] = []
    vacias = 0
    for fila in range(fila_enc + 1, (ws.max_row or fila_enc) + 1):
        registro: dict[str, object] = {}
        hay_datos = False
        for col, clave in encabezados.items():
            valor = ws.cell(row=fila, column=col).value
            registro[clave] = valor
            if valor is not None and str(valor).strip():
                hay_datos = True
        if hay_datos:
            filas.append(registro)
            vacias = 0
        else:
            vacias += 1
            if vacias >= 30:
                break
    return filas


def _leer_parametros(ws) -> dict[str, str]:
    """Hoja de dos columnas Parametro / Valor."""
    datos: dict[str, str] = {}
    if not ws.max_row:
        return datos
    for fila in range(1, (ws.max_row or 1) + 1):
        clave = normalizar(ws.cell(row=fila, column=1).value)
        if not clave or clave == "parametro":
            continue
        valor = ws.cell(row=fila, column=2).value
        datos[clave] = "" if valor is None else texto_limpio(valor)
    return datos


def _campo(registro: dict, *alias: str):
    for nombre in alias:
        clave = normalizar(nombre)
        if clave in registro:
            return registro[clave]
    for nombre in alias:                      # segunda pasada: parcial
        clave = normalizar(nombre)
        for existente, valor in registro.items():
            if clave and (clave in existente or existente in clave):
                return valor
    return None


# --------------------------------------------------------------------------- #
# Carga completa
# --------------------------------------------------------------------------- #
def cargar(ruta: str | Path) -> BaseDatos:
    ruta = Path(ruta)
    if not ruta.is_file():
        raise ErrorExcel(f"No se encontro el Excel de base de datos:\n{ruta}")

    wb = abrir_libro(ruta, data_only=True)
    bd = BaseDatos(ruta=ruta)
    try:
        _cargar_proyectos(wb, bd)
        _cargar_destinatarios(wb, bd)
        _cargar_mapeos(wb, bd)
        _cargar_smtp(wb, bd)
        _cargar_plantilla(wb, bd)
        _cargar_config(wb, bd)
    finally:
        wb.close()

    _validar(bd)
    return bd


def _cargar_proyectos(wb, bd: BaseDatos) -> None:
    ws = wb[buscar_hoja(wb, "Proyectos")]
    for registro in _leer_tabla(ws):
        codigo = texto_limpio(_campo(registro, "Código Proyecto", "Codigo Proyecto"))
        if not codigo:
            continue
        bd.proyectos.append(
            Proyecto(
                codigo=codigo,
                nombre=texto_limpio(_campo(registro, "Nombre Proyecto")),
                cliente=texto_limpio(_campo(registro, "Cliente")),
                ruta=texto_limpio(_campo(registro, "Ruta L")),
                archivo=texto_limpio(_campo(registro, "Nombre de Excel")),
                hoja_let=texto_limpio(_campo(registro, "Hoja LET")),
                hoja_ev=texto_limpio(_campo(registro, "Hoja EV")),
                estado_filtro=texto_limpio(_campo(registro, "Estado Filtro")),
                activo=es_verdadero(_campo(registro, "Activo"), True),
            )
        )


def _cargar_destinatarios(wb, bd: BaseDatos) -> None:
    try:
        ws = wb[buscar_hoja(wb, "Destinatarios")]
    except ErrorExcel:
        bd.avisos.append("No existe la hoja 'Destinatarios' en la base de datos.")
        return
    for registro in _leer_tabla(ws):
        correo = texto_limpio(_campo(registro, "Correo"))
        codigo = texto_limpio(_campo(registro, "Código Proyecto", "Codigo Proyecto"))
        if not correo or not codigo:
            continue
        tipo_n = normalizar(_campo(registro, "Tipo Destinatario", "Tipo"))
        if tipo_n in ("cc", "copia", "con copia"):
            tipo = "CC"
        elif tipo_n in ("cco", "bcc", "copia oculta"):
            tipo = "CCO"
        else:
            tipo = "Para"
        bd.destinatarios.append(
            Destinatario(
                codigo=codigo,
                nombre=texto_limpio(_campo(registro, "Nombre")),
                correo=correo,
                tipo=tipo,
                activo=es_verdadero(_campo(registro, "Activo"), True),
            )
        )


def _cargar_mapeos(wb, bd: BaseDatos) -> None:
    mapas_let: dict[str, dict[str, str]] = {}
    mapas_ev: dict[str, dict[str, str]] = {}

    try:
        ws = wb[buscar_hoja(wb, "MapeoLET")]
        etiquetas = [("Fila Encabezado", "fila_encabezado"),
                     ("Fila Inicio Datos", "fila_inicio")]
        etiquetas += [(titulo, clave) for clave, titulo, _ in CAMPOS_LET]
        for registro in _leer_tabla(ws):
            codigo = normalizar(_campo(registro, "Código Proyecto", "Codigo Proyecto"))
            if not codigo:
                continue
            destino: dict[str, str] = {}
            for titulo, clave in etiquetas:
                valor = texto_limpio(_campo(registro, titulo))
                if valor:
                    destino[clave] = valor
            mapas_let[codigo] = destino
    except ErrorExcel:
        pass

    try:
        ws = wb[buscar_hoja(wb, "MapeoEV")]
        etiquetas = [("Columna Inicio Semanas", "col_inicio"), ("Celda SPI", "spi")]
        etiquetas += [(titulo, clave) for clave, titulo, _ in CAMPOS_EV]
        for registro in _leer_tabla(ws):
            codigo = normalizar(_campo(registro, "Código Proyecto", "Codigo Proyecto"))
            if not codigo:
                continue
            destino = {}
            for titulo, clave in etiquetas:
                valor = texto_limpio(_campo(registro, titulo))
                if valor:
                    destino[clave] = valor
            mapas_ev[codigo] = destino
    except ErrorExcel:
        pass

    for proyecto in bd.proyectos:
        clave = normalizar(proyecto.codigo)
        proyecto.mapeo_let = mapas_let.get(clave, {})
        proyecto.mapeo_ev = mapas_ev.get(clave, {})


def _cargar_smtp(wb, bd: BaseDatos) -> None:
    try:
        datos = _leer_parametros(wb[buscar_hoja(wb, "SMTP")])
    except ErrorExcel:
        bd.avisos.append("No existe la hoja 'SMTP' en la base de datos.")
        datos = {}

    def leer(clave: str) -> str:
        valor = datos.get(normalizar(clave), "")
        return valor if str(valor).strip() else SMTP_DEFECTO.get(clave, "")

    puerto = a_float(leer("Puerto")) or 587
    timeout = a_float(leer("Timeout (s)")) or 30
    seguridad_n = normalizar(leer("Seguridad"))
    if seguridad_n in ("ssl", "ssl/tls", "tls implicito", "implicito"):
        seguridad = "SSL"
    elif seguridad_n in ("ninguna", "ninguno", "sin cifrado", "none"):
        seguridad = "Ninguna"
    else:
        seguridad = "STARTTLS"

    bd.smtp = ConfigSMTP(
        servidor=leer("Servidor").strip(),
        puerto=int(puerto),
        seguridad=seguridad,
        usuario=leer("Usuario").strip(),
        contrasena=str(datos.get(normalizar("Contraseña"), "")),
        remitente=leer("Remitente").strip(),
        nombre_remitente=leer("Nombre Remitente").strip(),
        responder_a=leer("Responder a").strip(),
        cco_fijo=leer("CCO fijo").strip(),
        timeout=int(timeout),
    )


def _cargar_plantilla(wb, bd: BaseDatos) -> None:
    try:
        bd.plantilla = _leer_parametros(wb[buscar_hoja(wb, "Plantilla")])
    except ErrorExcel:
        bd.avisos.append("No existe la hoja 'Plantilla'; se usaran los textos por defecto.")
        bd.plantilla = {}


def _cargar_config(wb, bd: BaseDatos) -> None:
    try:
        bd.config = _leer_parametros(wb[buscar_hoja(wb, "Config")])
    except ErrorExcel:
        bd.config = {}


def _validar(bd: BaseDatos) -> None:
    if not bd.proyectos:
        bd.avisos.append("La hoja 'Proyectos' no tiene ningun proyecto registrado.")

    vistos: set[str] = set()
    for proyecto in bd.proyectos:
        clave = normalizar(proyecto.codigo)
        if clave in vistos:
            bd.avisos.append(f"Codigo de proyecto duplicado: {proyecto.codigo}")
        vistos.add(clave)
        if not proyecto.estado_filtro:
            proyecto.estado_filtro = ESTADO_FILTRO_DEFECTO
        if proyecto.activo and not proyecto.existe_excel():
            bd.avisos.append(
                f"{proyecto.codigo}: no se encuentra el archivo {proyecto.ruta_excel}"
            )

    codigos = {normalizar(p.codigo) for p in bd.proyectos}
    huerfanos = {d.codigo for d in bd.destinatarios if normalizar(d.codigo) not in codigos}
    for codigo in sorted(huerfanos):
        bd.avisos.append(
            f"Hay destinatarios con el codigo '{codigo}', que no existe en 'Proyectos'."
        )

    for destinatario in bd.destinatarios:
        if "@" not in destinatario.correo:
            bd.avisos.append(f"Correo invalido: '{destinatario.correo}'")
        if destinatario.tipo not in TIPOS_DESTINATARIO:
            destinatario.tipo = "Para"
