"""Panel de datos: indicadores, filtro de estatus y tabla de entregables."""
from __future__ import annotations

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QCheckBox,
    QDialog,
    QFrame,
    QGridLayout,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QPushButton,
    QScrollArea,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constantes import COLUMNAS_CORREO


class TarjetaKpi(QFrame):
    def __init__(self, titulo: str, parent=None):
        super().__init__(parent)
        self.setObjectName("Tarjeta")
        disposicion = QVBoxLayout(self)
        disposicion.setContentsMargins(12, 9, 12, 9)
        disposicion.setSpacing(2)
        etiqueta = QLabel(titulo)
        etiqueta.setObjectName("KpiTitulo")
        self.valor = QLabel("-")
        self.valor.setObjectName("Kpi")
        disposicion.addWidget(etiqueta)
        disposicion.addWidget(self.valor)

    def establecer(self, texto: str, color: str | None = None) -> None:
        self.valor.setText(texto)
        self.valor.setStyleSheet(f"color: {color};" if color else "")


class DialogoDiagnostico(QDialog):
    def __init__(self, filas: list[tuple[str, str, str]], avisos: list[str], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Cómo se leyó el Excel del proyecto")
        self.resize(640, 560)
        disposicion = QVBoxLayout(self)

        explicacion = QLabel(
            "Origen de cada referencia: <b>BD</b> = indicada en el Excel de base de datos, "
            "<b>búsqueda por texto</b> = encontrada por el nombre del campo, "
            "<b>respaldo</b> = posición del archivo modelo."
        )
        explicacion.setWordWrap(True)
        explicacion.setObjectName("Subtitulo")
        disposicion.addWidget(explicacion)

        tabla = QTableWidget(len(filas), 3)
        tabla.setHorizontalHeaderLabels(["Campo", "Referencia", "Origen"])
        tabla.verticalHeader().setVisible(False)
        tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        tabla.setAlternatingRowColors(True)
        for i, (campo, referencia, origen) in enumerate(filas):
            tabla.setItem(i, 0, QTableWidgetItem(campo))
            tabla.setItem(i, 1, QTableWidgetItem(referencia))
            tabla.setItem(i, 2, QTableWidgetItem(origen))
        tabla.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeMode.Stretch)
        tabla.resizeColumnToContents(1)
        tabla.resizeColumnToContents(2)
        disposicion.addWidget(tabla)

        if avisos:
            texto = QLabel("Avisos:\n• " + "\n• ".join(avisos))
            texto.setWordWrap(True)
            texto.setStyleSheet("color:#B45309;")
            disposicion.addWidget(texto)

        cerrar = QPushButton("Cerrar")
        cerrar.clicked.connect(self.accept)
        fila = QHBoxLayout()
        fila.addStretch(1)
        fila.addWidget(cerrar)
        disposicion.addLayout(fila)


class PanelDatos(QWidget):
    """Indicadores, seleccion de estatus y entregables a incluir."""

    filtro_cambiado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self.datos = None
        self._casillas: list[QCheckBox] = []
        self._bloqueado = False

        raiz = QVBoxLayout(self)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.setSpacing(12)

        # -- indicadores --------------------------------------------------- #
        fila_kpis = QHBoxLayout()
        fila_kpis.setSpacing(10)
        self.kpi_semana = TarjetaKpi("Semana de corte")
        self.kpi_plan = TarjetaKpi("Avance planificado")
        self.kpi_real = TarjetaKpi("Avance real")
        self.kpi_desv = TarjetaKpi("Desviación")
        self.kpi_spi = TarjetaKpi("SPI")
        for tarjeta in (self.kpi_semana, self.kpi_plan, self.kpi_real,
                        self.kpi_desv, self.kpi_spi):
            fila_kpis.addWidget(tarjeta)
        raiz.addLayout(fila_kpis)

        # -- filtro de estatus --------------------------------------------- #
        grupo = QGroupBox("Estatus del entregable a incluir")
        interior = QVBoxLayout(grupo)
        self.contenedor_estatus = QWidget()
        self.rejilla_estatus = QGridLayout(self.contenedor_estatus)
        self.rejilla_estatus.setContentsMargins(0, 0, 0, 0)
        self.rejilla_estatus.setSpacing(6)
        area = QScrollArea()
        area.setWidgetResizable(True)
        area.setFrameShape(QFrame.Shape.NoFrame)
        area.setWidget(self.contenedor_estatus)
        area.setMaximumHeight(110)
        interior.addWidget(area)
        raiz.addWidget(grupo)

        # -- tabla de entregables ------------------------------------------ #
        barra = QHBoxLayout()
        self.resumen = QLabel("Sin proyecto seleccionado")
        self.resumen.setObjectName("Subtitulo")
        barra.addWidget(self.resumen)
        barra.addStretch(1)
        self.btn_todos = QPushButton("Marcar todo")
        self.btn_todos.setObjectName("Plano")
        self.btn_todos.clicked.connect(lambda: self._marcar_filas(True))
        self.btn_ninguno = QPushButton("Desmarcar todo")
        self.btn_ninguno.setObjectName("Plano")
        self.btn_ninguno.clicked.connect(lambda: self._marcar_filas(False))
        self.btn_diagnostico = QPushButton("Ver cómo se leyó el Excel")
        self.btn_diagnostico.setObjectName("Plano")
        self.btn_diagnostico.clicked.connect(self._mostrar_diagnostico)
        for boton in (self.btn_todos, self.btn_ninguno, self.btn_diagnostico):
            barra.addWidget(boton)
        raiz.addLayout(barra)

        self.tabla = QTableWidget(0, len(COLUMNAS_CORREO) + 1)
        self.tabla.setHorizontalHeaderLabels(
            ["Incluir"] + [titulo for _, titulo in COLUMNAS_CORREO]
        )
        self.tabla.verticalHeader().setVisible(False)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabla.setTextElideMode(Qt.TextElideMode.ElideRight)
        cabecera = self.tabla.horizontalHeader()
        cabecera.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        # anchos fijos pero ajustables: si no, el codigo del cliente (muy largo)
        # se come todo el espacio y deja ilegible el nombre del entregable
        for indice, ancho in enumerate((330, 140, 290, 95, 120, 170), start=1):
            cabecera.setSectionResizeMode(indice, QHeaderView.ResizeMode.Interactive)
            self.tabla.setColumnWidth(indice, ancho)
        self.tabla.itemChanged.connect(self._fila_cambiada)
        raiz.addWidget(self.tabla, 1)

    # -- carga ------------------------------------------------------------- #
    def cargar(self, datos, estatus_iniciales: list[str], decimales: int = 0) -> None:
        self.datos = datos
        self._actualizar_kpis(datos, decimales)
        self._construir_estatus(datos, estatus_iniciales)
        self.refrescar_tabla()

    def _actualizar_kpis(self, datos, decimales: int) -> None:
        kpis = datos.ev.kpis(decimales)
        semana = datos.ev.semana_corte or "-"
        self.kpi_semana.establecer(semana)
        self.kpi_plan.establecer(kpis["avance_planificado"])
        self.kpi_real.establecer(kpis["avance_real"])

        desviacion = kpis["desviacion"]
        color = None
        if desviacion not in ("-", ""):
            color = "#C0392B" if desviacion.startswith("-") else "#107C41"
        self.kpi_desv.establecer(desviacion, color)

        spi = kpis["spi"]
        color_spi = None
        try:
            color_spi = "#C0392B" if float(spi) < 0.95 else "#107C41"
        except ValueError:
            pass
        self.kpi_spi.establecer(spi, color_spi)

    def _construir_estatus(self, datos, seleccionados: list[str]) -> None:
        for casilla in self._casillas:
            casilla.setParent(None)
        self._casillas.clear()

        objetivo = {v.strip().lower() for v in seleccionados}
        valores = datos.let.conteo_estatus.most_common()
        columnas = 3
        for i, (valor, cantidad) in enumerate(valores):
            casilla = QCheckBox(f"{valor}  ({cantidad})")
            casilla.setProperty("valor", valor)
            casilla.setChecked(valor.strip().lower() in objetivo)
            casilla.stateChanged.connect(self._estatus_cambiado)
            self.rejilla_estatus.addWidget(casilla, i // columnas, i % columnas)
            self._casillas.append(casilla)

        if not valores:
            etiqueta = QLabel("La hoja LET no devolvió ningún estatus.")
            etiqueta.setObjectName("Subtitulo")
            self.rejilla_estatus.addWidget(etiqueta, 0, 0)

    # -- tabla ------------------------------------------------------------- #
    def refrescar_tabla(self) -> None:
        if self.datos is None:
            return
        entregables = self.datos.let.filtrar(self.estatus_seleccionados())
        self._bloqueado = True
        try:
            self.tabla.setRowCount(len(entregables))
            for fila, entregable in enumerate(entregables):
                marca = QTableWidgetItem()
                marca.setFlags(
                    Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled
                )
                marca.setCheckState(Qt.CheckState.Checked)
                marca.setData(Qt.ItemDataRole.UserRole, entregable)
                self.tabla.setItem(fila, 0, marca)
                for col, (clave, _) in enumerate(COLUMNAS_CORREO, start=1):
                    self.tabla.setItem(
                        fila, col, QTableWidgetItem(str(entregable.valor(clave) or ""))
                    )
        finally:
            self._bloqueado = False
        self._actualizar_resumen()

    def _actualizar_resumen(self) -> None:
        total = self.tabla.rowCount()
        incluidos = len(self.entregables_seleccionados())
        if self.datos is None:
            self.resumen.setText("Sin proyecto seleccionado")
            return
        leidos = len(self.datos.let.entregables)
        self.resumen.setText(
            f"{incluidos} de {total} entregables filtrados se incluirán en el correo "
            f"(la hoja {self.datos.let.hoja} tiene {leidos} filas)."
        )

    def _marcar_filas(self, marcar: bool) -> None:
        self._bloqueado = True
        try:
            estado = Qt.CheckState.Checked if marcar else Qt.CheckState.Unchecked
            for fila in range(self.tabla.rowCount()):
                item = self.tabla.item(fila, 0)
                if item is not None:
                    item.setCheckState(estado)
        finally:
            self._bloqueado = False
        self._actualizar_resumen()
        self.filtro_cambiado.emit()

    def _fila_cambiada(self, item) -> None:
        if self._bloqueado or item.column() != 0:
            return
        self._actualizar_resumen()
        self.filtro_cambiado.emit()

    def _estatus_cambiado(self) -> None:
        if self._bloqueado:
            return
        self.refrescar_tabla()
        self.filtro_cambiado.emit()

    def _mostrar_diagnostico(self) -> None:
        if self.datos is None:
            return
        DialogoDiagnostico(self.datos.diagnostico(), self.datos.avisos, self).exec()

    # -- consultas --------------------------------------------------------- #
    def estatus_seleccionados(self) -> list[str]:
        return [c.property("valor") for c in self._casillas if c.isChecked()]

    def entregables_seleccionados(self) -> list:
        seleccion = []
        for fila in range(self.tabla.rowCount()):
            item = self.tabla.item(fila, 0)
            if item is not None and item.checkState() == Qt.CheckState.Checked:
                seleccion.append(item.data(Qt.ItemDataRole.UserRole))
        return seleccion

    def limpiar(self) -> None:
        self.datos = None
        self.tabla.setRowCount(0)
        for casilla in self._casillas:
            casilla.setParent(None)
        self._casillas.clear()
        for tarjeta in (self.kpi_semana, self.kpi_plan, self.kpi_real,
                        self.kpi_desv, self.kpi_spi):
            tarjeta.establecer("-")
        self.resumen.setText("Sin proyecto seleccionado")
