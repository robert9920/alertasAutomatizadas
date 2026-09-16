"""Pruebas de humo de App Alertas.

Comprueba la lectura del Excel modelo, la tolerancia a cambios de estructura
(nombres de hoja, columnas y filas desplazadas), el mapeo explicito de la BD,
el filtro por estatus y el armado del correo.

    python pruebas.py
"""
from __future__ import annotations

import sys
import tempfile
import warnings
from pathlib import Path

warnings.filterwarnings("ignore")
sys.path.insert(0, str(Path(__file__).resolve().parent))

from openpyxl import Workbook                                   # noqa: E402

from app import bd as modulo_bd                                 # noqa: E402
from app import chart, email_builder, lectura, sender           # noqa: E402
from app.constantes import NOMBRE_BD                            # noqa: E402
from app.excel_compat import abrir as abrir_libro               # noqa: E402
from app.excel_utils import coincide, normalizar                # noqa: E402

BASE = Path(__file__).resolve().parent
REFERENCIA = BASE.parent / "Referencia" / "LE.xlsx"

_resultados: list[tuple[bool, str, str]] = []


def comprobar(condicion: bool, titulo: str, detalle: str = "") -> None:
    _resultados.append((bool(condicion), titulo, detalle))
    marca = "OK  " if condicion else "FALLA"
    print(f"  [{marca}] {titulo}" + (f"  -> {detalle}" if detalle else ""))


def proyecto(ruta: Path, **extra) -> modulo_bd.Proyecto:
    return modulo_bd.Proyecto(
        codigo="PRUEBA", nombre="Proyecto de prueba",
        ruta=str(ruta.parent), archivo=ruta.stem,
        estado_filtro="En revisión del cliente", **extra
    )


# --------------------------------------------------------------------------- #
def prueba_normalizacion() -> None:
    print("\n1. Normalización de texto")
    comprobar(normalizar("  CÓDIGO  DEL ENTREGABLE ") == "codigo del entregable",
              "quita tildes, mayúsculas y espacios duros")
    comprobar(coincide("En revisión de cliente", "En revisión del cliente"),
              "tolera 'de' frente a 'del'")
    comprobar(coincide("ESTATUS DEL ENTREGABLE LC", "estatus del entregable lc"),
              "ignora mayúsculas")
    comprobar(not coincide("Aprobado", "En revisión del cliente"),
              "no confunde estatus distintos")


def prueba_referencia() -> lectura.DatosProyecto | None:
    print("\n2. Lectura del archivo modelo (Referencia/LE.xlsx)")
    if not REFERENCIA.is_file():
        comprobar(False, "existe Referencia/LE.xlsx", str(REFERENCIA))
        return None

    datos = lectura.cargar(proyecto(REFERENCIA))
    ev, let = datos.ev, datos.let
    comprobar(let.hoja == "LET" and ev.hoja == "EV", "detecta las hojas LET y EV",
              f"{let.hoja} / {ev.hoja}")
    comprobar(ev.semana_corte == "S15", "semana de corte", ev.semana_corte)
    kpis = ev.kpis(0)
    comprobar(kpis["avance_planificado"] == "60%", "avance planificado",
              kpis["avance_planificado"])
    comprobar(kpis["avance_real"] == "44%", "avance real", kpis["avance_real"])
    comprobar(kpis["desviacion"] == "-16%", "desviación", kpis["desviacion"])
    comprobar(kpis["spi"] == "0.74", "SPI", kpis["spi"])
    comprobar(len(ev.semanas) == 24, "semanas graficadas", str(len(ev.semanas)))
    comprobar(len(ev.meses) == 6 and ev.meses[0][0] == "Mes 1", "bandas de mes",
              str([m[0] for m in ev.meses]))
    comprobar(len(let.entregables) == 405, "filas leídas en LET",
              str(len(let.entregables)))
    sugeridos = let.estatus_sugeridos("En revisión del cliente")
    comprobar(len(let.filtrar(sugeridos)) == 17, "entregables en revisión del cliente",
              str(len(let.filtrar(sugeridos))))
    comprobar(len(let.filtrar(["Cerrado"])) == 25, "filtro alternativo 'Cerrado'",
              str(len(let.filtrar(["Cerrado"]))))
    return datos


def prueba_hojas_renombradas(carpeta: Path) -> None:
    print("\n3. Hojas renombradas con el nombre del cliente")
    if not REFERENCIA.is_file():
        comprobar(False, "no se puede probar sin el archivo modelo")
        return
    wb = abrir_libro(REFERENCIA, data_only=True)
    wb["EV"].title = "EV - RAURA"
    wb["LET"].title = "LET_RAURA"
    destino = carpeta / "LE_renombrado.xlsx"
    wb.save(destino)
    wb.close()

    datos = lectura.cargar(proyecto(destino))
    comprobar(datos.ev.hoja == "EV - RAURA", "encuentra 'EV - RAURA'", datos.ev.hoja)
    comprobar(datos.let.hoja == "LET_RAURA", "encuentra 'LET_RAURA'", datos.let.hoja)
    comprobar(datos.ev.semana_corte == "S15", "sigue calculando la semana de corte",
              datos.ev.semana_corte)


def _libro_desplazado(destino: Path) -> None:
    """Mismo contenido que el modelo, pero en otras filas y columnas."""
    wb = Workbook()
    wb.remove(wb.active)

    # --- hoja EV con etiquetas en la columna B y semanas desde la columna E --- #
    ev = wb.create_sheet("EV DEL CLIENTE")
    ev["B17"] = "MES"
    ev["B18"] = "SEMANA"
    etiquetas = ["% Previsto", "% Previsto Acum", "% Real", "% Real Acum",
                 "% Tendencia", "% Tendencia Acum"]
    for i, etiqueta in enumerate(etiquetas):
        ev.cell(row=19 + i, column=2, value=etiqueta)

    previsto = [0.10, 0.20, 0.15, 0.25, 0.30]
    real = [0.10, 0.18, 0.12, None, None]
    tendencia = [None, None, 0.12, 0.30, 0.34]
    acumulado = 0.0
    acumulado_real = 0.0
    acumulado_tend = 0.0
    for i in range(5):
        columna = 5 + i
        ev.cell(row=17, column=columna, value=f"Mes {1 if i < 3 else 2}")
        ev.cell(row=18, column=columna, value=f"S{i + 1}")
        ev.cell(row=19, column=columna, value=previsto[i])
        acumulado += previsto[i]
        ev.cell(row=20, column=columna, value=acumulado)
        if real[i] is not None:
            ev.cell(row=21, column=columna, value=real[i])
            acumulado_real += real[i]
            ev.cell(row=22, column=columna, value=acumulado_real)
        if tendencia[i] is not None:
            ev.cell(row=23, column=columna, value=tendencia[i])
            acumulado_tend += tendencia[i]
            ev.cell(row=24, column=columna, value=acumulado_tend)
    ev["C60"] = "SPI"
    ev["E60"] = 0.88

    # --- hoja LET con encabezados en la fila 5 y columnas movidas ----------- #
    let = wb.create_sheet("LET DEL CLIENTE")
    posiciones = {
        "NOMBRE DEL ENTREGABLE": 4,
        "DISCIPLINA": 9,
        "CÓDIGO DE ENTREGABLE CLIENTE": 15,
        "REVISIÓN ACTUAL": 30,
        "FECHA ÚLTIMO ENVÍO A CLIENTE": 31,
        "ESTATUS DEL ENTREGABLE LC": 44,
    }
    let["A5"] = "ITEM"
    for titulo, columna in posiciones.items():
        let.cell(row=5, column=columna, value=titulo)
    filas = [
        ("Memoria de cálculo", "CIVIL", "ABC-001", "Rev B", "01/09/2026",
         "En revision de cliente"),
        ("Plano de planta", "MECÁNICA", "ABC-002", "Rev C", "02/09/2026",
         "En revision de cliente"),
        ("Especificación técnica", "ELÉCTRICA", "ABC-003", "Rev 0", "03/09/2026",
         "Aprobado"),
    ]
    for i, valores in enumerate(filas):
        fila = 6 + i
        let.cell(row=fila, column=1, value=i + 1)
        for (titulo, columna), valor in zip(posiciones.items(), valores):
            let.cell(row=fila, column=columna, value=valor)

    wb.save(destino)


def prueba_estructura_desplazada(carpeta: Path) -> Path:
    print("\n4. Estructura en otras filas y columnas (búsqueda por texto)")
    destino = carpeta / "Estructura_movida.xlsx"
    _libro_desplazado(destino)

    datos = lectura.cargar(proyecto(destino))
    ev, let = datos.ev, datos.let
    comprobar(let.hoja == "LET DEL CLIENTE" and ev.hoja == "EV DEL CLIENTE",
              "encuentra las hojas por prefijo", f"{let.hoja} / {ev.hoja}")
    comprobar(let.fila_encabezado == 5, "fila de encabezado de LET",
              str(let.fila_encabezado))
    comprobar(let.columnas["codigo_cliente"] == 15, "columna del código de cliente",
              str(let.columnas["codigo_cliente"]))
    comprobar(let.columnas["estatus"] == 44, "columna del estatus",
              str(let.columnas["estatus"]))
    comprobar(all(v == "busqueda por texto" for k, v in let.origen.items()
                  if k in let.columnas),
              "todas las columnas se hallaron por texto")
    comprobar(ev.filas["previsto_acum"] == 20, "fila de % Previsto Acum",
              str(ev.filas["previsto_acum"]))
    comprobar(ev.col_inicio == 5, "columna de la primera semana", str(ev.col_inicio))
    comprobar(ev.semana_corte == "S3", "semana de corte con datos parciales",
              ev.semana_corte)
    comprobar(ev.spi == 0.88, "SPI hallado por su etiqueta", str(ev.spi))
    sugeridos = let.estatus_sugeridos("En revisión del cliente")
    comprobar(len(let.filtrar(sugeridos)) == 2,
              "filtra 'En revision de cliente' pese a la variante sin tilde",
              str(len(let.filtrar(sugeridos))))
    return destino


def prueba_mapeo_explicito(ruta: Path) -> None:
    print("\n5. Mapeo explícito indicado en la base de datos")
    p = proyecto(
        ruta,
        hoja_let="LET DEL CLIENTE", hoja_ev="EV DEL CLIENTE",
        mapeo_let={"fila_encabezado": "5", "fila_inicio": "6", "nombre": "d",
                   "disciplina": "I", "codigo_cliente": "o", "revision": "AD",
                   "fecha_envio": "AE", "estatus": "AR"},
        mapeo_ev={"col_inicio": "e", "spi": "E60", "mes": "17", "semana": "18",
                  "previsto": "19", "previsto_acum": "20", "real": "21",
                  "real_acum": "22", "tendencia": "23", "tendencia_acum": "24"},
    )
    datos = lectura.cargar(p)
    comprobar(datos.let.origen["estatus"] == "BD",
              "la columna del estatus viene de la BD", datos.let.origen["estatus"])
    comprobar(datos.ev.origen["previsto_acum"] == "BD",
              "la fila de % Previsto Acum viene de la BD")
    comprobar(datos.ev.origen["spi"] == "BD" and datos.ev.spi == 0.88,
              "el SPI se lee de la celda indicada", str(datos.ev.spi))
    comprobar(datos.let.columnas["nombre"] == 4,
              "acepta la letra de columna en minúscula ('d')",
              str(datos.let.columnas["nombre"]))
    comprobar(len(datos.let.entregables) == 3, "lee las 3 filas de datos",
              str(len(datos.let.entregables)))


def prueba_correo(datos) -> None:
    print("\n6. Armado del correo")
    if datos is None:
        comprobar(False, "no se puede probar sin datos del archivo modelo")
        return
    base = modulo_bd.cargar(BASE / NOMBRE_BD)
    sugeridos = datos.let.estatus_sugeridos("En revisión del cliente")
    entregables = datos.let.filtrar(sugeridos)

    asunto = email_builder.construir_asunto(base, datos, entregables)
    comprobar(asunto.startswith("Gestión: Reporte de proyecto"), "asunto generado", asunto[:60])
    comprobar("{" not in asunto, "no quedan comodines sin reemplazar")

    html = email_builder.construir_html(base, datos, entregables)
    comprobar(html.count("<tr>") == len(entregables) + 1,
              "la tabla trae encabezado + una fila por entregable",
              f"{html.count('<tr>')} filas")
    comprobar("cid:curva_s" in html, "la Curva S se referencia como imagen incrustada")
    comprobar("Semana 15" in html, "el texto menciona la semana correcta")
    comprobar("<b>Avance Planificado:</b> 60%" in html, "bullet de avance planificado")

    texto = email_builder.a_texto_plano(html)
    comprobar("Estimados ingenieros" in texto and "<" not in texto,
              "alternativa en texto plano sin etiquetas")

    png = chart.generar(datos.ev, chart.OpcionesGrafico.desde_bd(base))
    envio = sender.Envio(
        asunto=asunto, html=html, texto=texto, imagen=png,
        para=["destino@ejemplo.com"], cc=["copia@ejemplo.com"],
    )
    mensaje = sender.construir_mensaje(base.smtp, envio)
    crudo = mensaje.as_string()
    comprobar(mensaje.is_multipart(), "el mensaje MIME es multiparte")
    comprobar("curva_s" in crudo and "image/png" in crudo,
              "la imagen viaja incrustada en el mensaje")
    comprobar(mensaje["Cc"] == "copia@ejemplo.com", "cabecera CC")
    comprobar(envio.destinos == ["destino@ejemplo.com", "copia@ejemplo.com"],
              "lista de destinatarios del sobre SMTP")


# --------------------------------------------------------------------------- #
def main() -> int:
    print("=" * 74)
    print("Pruebas de App Alertas")
    print("=" * 74)

    prueba_normalizacion()
    datos = prueba_referencia()
    with tempfile.TemporaryDirectory(prefix="pruebas_alertas_") as temporal:
        carpeta = Path(temporal)
        prueba_hojas_renombradas(carpeta)
        ruta_movida = prueba_estructura_desplazada(carpeta)
        prueba_mapeo_explicito(ruta_movida)
        prueba_correo(datos)

    fallos = [t for ok, t, _ in _resultados if not ok]
    print("\n" + "=" * 74)
    print(f"{len(_resultados) - len(fallos)} de {len(_resultados)} comprobaciones correctas")
    if fallos:
        print("Fallaron:")
        for titulo in fallos:
            print("  -", titulo)
    print("=" * 74)
    return 1 if fallos else 0


if __name__ == "__main__":
    raise SystemExit(main())
