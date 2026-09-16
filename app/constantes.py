"""Nombres de campos, posiciones de respaldo y valores por defecto.

Las posiciones de respaldo provienen del archivo modelo `Referencia/LE.xlsx`
y solo se usan cuando el mapeo de la BD esta vacio y la busqueda por texto falla.
"""
from __future__ import annotations

NOMBRE_BD = "BD_Reportes.xlsx"

# --------------------------------------------------------------------------- #
# Hoja LET  ->  (clave interna, encabezado a buscar, columna de respaldo)
# --------------------------------------------------------------------------- #
CAMPOS_LET: list[tuple[str, str, str]] = [
    ("nombre",         "NOMBRE DEL ENTREGABLE",        "M"),
    ("disciplina",     "DISCIPLINA",                   "N"),
    ("codigo_cliente", "CÓDIGO DE ENTREGABLE CLIENTE", "W"),
    ("revision",       "REVISIÓN ACTUAL",              "BS"),
    ("fecha_envio",    "FECHA ÚLTIMO ENVÍO A CLIENTE", "BU"),
    ("estatus",        "ESTATUS DEL ENTREGABLE LC",    "CC"),
]
FILA_ENCABEZADO_LET = 8          # respaldo
FILA_INICIO_DATOS_LET = 9        # respaldo

# --------------------------------------------------------------------------- #
# Hoja EV  ->  (clave interna, etiqueta a buscar, fila de respaldo)
# --------------------------------------------------------------------------- #
CAMPOS_EV: list[tuple[str, str, int]] = [
    ("mes",             "MES",               10),
    ("semana",          "SEMANA",            11),
    ("previsto",        "% Previsto",        12),
    ("previsto_acum",   "% Previsto Acum",   13),
    ("real",            "% Real",            15),
    ("real_acum",       "% Real Acum",       16),
    ("tendencia",       "% Tendencia",       18),
    ("tendencia_acum",  "% Tendencia Acum",  19),
]
COL_INICIO_SEMANAS = "C"         # respaldo
CELDA_SPI = "H44"                # respaldo

ESTADO_FILTRO_DEFECTO = "En revisión del cliente"
TIPOS_DESTINATARIO = ("Para", "CC", "CCO")

# Encabezados de la tabla del correo, en orden
COLUMNAS_CORREO: list[tuple[str, str]] = [
    ("nombre",         "NOMBRE DEL ENTREGABLE"),
    ("disciplina",     "DISCIPLINA"),
    ("codigo_cliente", "CÓDIGO DE ENTREGABLE CLIENTE"),
    ("revision",       "REVISIÓN ACTUAL"),
    ("fecha_envio",    "FECHA ÚLTIMO ENVÍO A CLIENTE"),
    ("dias_espera",    "DÍAS DE ESPERA"),
    ("estatus",        "ESTATUS DEL ENTREGABLE LC"),
]

# --------------------------------------------------------------------------- #
# Hoja Plantilla
# --------------------------------------------------------------------------- #
PLANTILLA_DEFECTO: dict[str, str] = {
    "Asunto": "Gestión: Reporte de proyecto {fecha_hoy} | {codigo_proyecto} - {nombre_proyecto}",
    "Formato fecha asunto": "dd/mm/yyyy",
    "Saludo": "Estimados ingenieros, buenas noches",
    "Párrafo intro": (
        "En línea con el seguimiento y control del proyecto, compartimos el "
        "<b>Reporte Semanal de Avance correspondiente a la Semana {semana}</b>, "
        "el cual presenta el estado actualizado del proyecto y los principales "
        "aspectos que requieren seguimiento."
    ),
    "Párrafo Curva S": (
        "Como parte del reporte, se presenta la <b>Curva S del proyecto</b>, donde se "
        "muestra la comparación entre el avance planificado y el avance real, "
        "permitiendo identificar la tendencia actual del proyecto y las principales "
        "desviaciones respecto a la línea base."
    ),
    "Título Avance": "Avance del Proyecto",
    "Párrafo entregables 1": (
        "Asimismo, se incluye el <b>estado de los entregables pendientes de "
        "conformidad</b>, indicando aquellos que se encuentran aún en revisión y "
        "requieren respuesta para continuar con el cierre de los respectivos entregables."
    ),
    "Párrafo entregables 2": (
        "Agradeceremos su apoyo con la revisión y emisión de los comentarios o "
        "conformidades pendientes, a fin de mantener actualizado el control "
        "documentario y evitar impactos en la programación de las actividades subsecuentes."
    ),
    "Cierre": "Muchas gracias",
    "Firma": "",
}

AYUDA_PLANTILLA: dict[str, str] = {
    "Asunto": "Admite {fecha_hoy} {codigo_proyecto} {nombre_proyecto} {cliente} {semana}",
    "Formato fecha asunto": "dd/mm/yyyy, dd-mm-yyyy, yyyy-mm-dd ...",
    "Saludo": "Primera línea del correo.",
    "Párrafo intro": "Admite {semana}. Se permite HTML simple (<b>, <i>, <u>).",
    "Párrafo Curva S": "Texto que antecede al gráfico de la Curva S.",
    "Título Avance": "Encabezado de la lista de indicadores.",
    "Párrafo entregables 1": "Texto que antecede a la tabla de entregables.",
    "Párrafo entregables 2": "Segundo párrafo antes de la tabla.",
    "Cierre": "Despedida al final del correo.",
    "Firma": "Opcional. Admite HTML.",
}

# --------------------------------------------------------------------------- #
# Hoja Config
# --------------------------------------------------------------------------- #
COLOR_LINEA_PREVISTO = "#1A1A1A"
COLOR_LINEA_REAL = "#4EA72E"
COLOR_LINEA_TENDENCIA = "#2F5BE8"
COLOR_BARRA_PREVISTO = "#BFBFBF"
COLOR_BANDA_EJE_X = "#F2F2F2"
COLOR_TEXTO_EJE_X = "#595959"

CONFIG_DEFECTO: dict[str, str] = {
    # -- gráfico: tamaño y ejes ------------------------------------------- #
    "Ancho gráfico px": "1400",
    "Alto gráfico px": "520",
    "DPI": "110",
    "Eje Y izq mín": "0",
    "Eje Y izq máx": "120",
    "Eje Y der mín": "0",
    "Eje Y der máx": "60",
    "Semanas a mostrar": "Todas",
    "Mostrar etiquetas de datos": "Sí",
    # -- gráfico: tamaños de letra (vacío = automático) -------------------- #
    "Tamaño etiquetas líneas": "",
    "Tamaño etiquetas barras": "",
    "Tamaño eje Y izquierdo": "8",
    "Tamaño eje Y derecho": "8",
    "Tamaño eje X semanas": "",
    "Tamaño eje X meses": "",
    "Tamaño leyenda": "8",
    "Tamaño título gráfico": "11",
    # -- gráfico: colores de las series ------------------------------------ #
    "Color línea Previsto Acum": COLOR_LINEA_PREVISTO,
    "Color línea Real Acum": COLOR_LINEA_REAL,
    "Color línea Tendencia Acum": COLOR_LINEA_TENDENCIA,
    "Color barra Previsto": COLOR_BARRA_PREVISTO,
    "Color barra Real": COLOR_LINEA_REAL,
    "Color barra Tendencia": COLOR_LINEA_TENDENCIA,
    # -- gráfico: colores de las etiquetas (vacío = heredado) -------------- #
    "Color etiqueta Previsto Acum": "",
    "Color etiqueta Real Acum": "",
    "Color etiqueta Tendencia Acum": "",
    "Color etiqueta barra Previsto": "",
    "Color etiqueta barra Real": "",
    "Color etiqueta barra Tendencia": "",
    # -- gráfico: eje X y título ------------------------------------------- #
    "Color bandas eje X": COLOR_BANDA_EJE_X,
    "Color texto eje X": COLOR_TEXTO_EJE_X,
    "Color título gráfico": COLOR_LINEA_PREVISTO,
    # -- correo: tabla de entregables -------------------------------------- #
    "Color encabezado tabla": "#F8827F",
    "Color encabezado estatus": "#F2F2F2",
    "Decimales avance": "0",
    # -- correo: firma ------------------------------------------------------ #
    "Ruta firma": "",
    "Ancho firma px": "330",
    # -- aplicación --------------------------------------------------------- #
    "Tema": "Claro",
}

# Títulos de sección que `crear_bd.py` escribe antes de la clave indicada.
SECCIONES_CONFIG: dict[str, str] = {
    "Ancho gráfico px": "GRÁFICO · tamaño de la imagen y ejes",
    "Tamaño etiquetas líneas": "GRÁFICO · tamaños de letra  (vacío = se ajusta solo)",
    "Color línea Previsto Acum": "GRÁFICO · colores de las series",
    "Color etiqueta Previsto Acum": "GRÁFICO · colores de las etiquetas  (vacío = heredado)",
    "Color bandas eje X": "GRÁFICO · eje X y título",
    "Color encabezado tabla": "CORREO · tabla de entregables",
    "Ruta firma": "CORREO · firma",
    "Tema": "APLICACIÓN",
}

_AYUDA_VACIO_AUTO = "Vacío: se ajusta solo al número de semanas. Con un número, manda ese."

AYUDA_CONFIG: dict[str, str] = {
    "Ancho gráfico px": "Ancho de la imagen de la Curva S.",
    "Alto gráfico px": "Alto de la imagen de la Curva S.",
    "DPI": "Resolución de la imagen (110 es nítido y liviano).",
    "Eje Y izq mín": "Mínimo del eje de las líneas acumuladas, en %.",
    "Eje Y izq máx": "Máximo del eje de las líneas acumuladas, en %.",
    "Eje Y der mín": "Mínimo del eje de las barras semanales, en %.",
    "Eje Y der máx": "Máximo del eje de las barras semanales, en %.",
    "Semanas a mostrar": "Todas  |  Hasta semana de corte",
    "Mostrar etiquetas de datos": "Sí / No. Muestra el % sobre cada punto y barra.",
    "Tamaño etiquetas líneas": _AYUDA_VACIO_AUTO,
    "Tamaño etiquetas barras": _AYUDA_VACIO_AUTO,
    "Tamaño eje Y izquierdo": "Tamaño de los porcentajes del eje izquierdo.",
    "Tamaño eje Y derecho": "Tamaño de los porcentajes del eje derecho.",
    "Tamaño eje X semanas": _AYUDA_VACIO_AUTO,
    "Tamaño eje X meses": _AYUDA_VACIO_AUTO,
    "Tamaño leyenda": "Tamaño del texto de la leyenda.",
    "Tamaño título gráfico": "Tamaño del nombre del proyecto sobre el gráfico.",
    "Color línea Previsto Acum": "Color de la línea, en formato #RRGGBB.",
    "Color línea Real Acum": "Color de la línea, en formato #RRGGBB.",
    "Color línea Tendencia Acum": "Color de la línea, en formato #RRGGBB.",
    "Color barra Previsto": "Color de las barras semanales previstas.",
    "Color barra Real": "Color de las barras semanales reales.",
    "Color barra Tendencia": "Color de las barras semanales de tendencia.",
    "Color etiqueta Previsto Acum": "Vacío: usa el color de su propia línea.",
    "Color etiqueta Real Acum": "Vacío: usa el color de su propia línea.",
    "Color etiqueta Tendencia Acum": "Vacío: usa el color de su propia línea.",
    "Color etiqueta barra Previsto": "Vacío: negro.",
    "Color etiqueta barra Real": "Vacío: negro.",
    "Color etiqueta barra Tendencia": "Vacío: negro.",
    "Color bandas eje X": "Fondo de los recuadros de semanas y meses.",
    "Color texto eje X": "Color de «S1», «S2», «Mes 1»… Ajústalo si oscureces las bandas.",
    "Color título gráfico": "Color del nombre del proyecto sobre el gráfico.",
    "Color encabezado tabla": "Fondo del encabezado de la tabla, salvo la última columna.",
    "Color encabezado estatus": "Fondo del encabezado de «ESTATUS DEL ENTREGABLE LC».",
    "Decimales avance": "Decimales de Planificado, Real y Desviación (0 = enteros).",
    "Ruta firma": (
        "Vacío: busca firma.png (o .jpg) en la carpeta del programa. "
        "Si la tienes en otro sitio, escribe aquí la ruta completa del archivo."
    ),
    "Ancho firma px": "Ancho con el que se inserta la firma en el correo.",
    "Tema": "Claro  |  Oscuro",
}

# --------------------------------------------------------------------------- #
# Hoja SMTP
# --------------------------------------------------------------------------- #
SMTP_DEFECTO: dict[str, str] = {
    "Servidor": "smtp.office365.com",
    "Puerto": "587",
    "Seguridad": "STARTTLS",
    "Usuario": "",
    "Contraseña": "",
    "Remitente": "",
    "Nombre Remitente": "",
    "Responder a": "",
    "CCO fijo": "",
    "Timeout (s)": "30",
}

AYUDA_SMTP: dict[str, str] = {
    "Servidor": "Host SMTP. Ej: smtp.office365.com, smtp.gmail.com",
    "Puerto": "587 para STARTTLS, 465 para SSL, 25 sin cifrado.",
    "Seguridad": "STARTTLS | SSL | Ninguna",
    "Usuario": "Cuenta con la que se autentica el envío.",
    "Contraseña": "Contraseña o App Password. Se guarda en texto plano.",
    "Remitente": "Dirección que aparece en De. Si se deja vacío se usa Usuario.",
    "Nombre Remitente": "Nombre visible del remitente. Opcional.",
    "Responder a": "Reply-To. Opcional.",
    "CCO fijo": "Correo que siempre irá en copia oculta. Opcional.",
    "Timeout (s)": "Segundos de espera antes de abortar la conexión.",
}
