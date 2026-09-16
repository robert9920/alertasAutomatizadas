"""Panel derecho: asunto, destinatarios, adjuntos y boton de envio."""
from __future__ import annotations

from pathlib import Path

from PySide6.QtCore import Qt, Signal
from PySide6.QtWidgets import (
    QAbstractItemView,
    QComboBox,
    QFileDialog,
    QGroupBox,
    QHBoxLayout,
    QHeaderView,
    QInputDialog,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QPushButton,
    QTableWidget,
    QTableWidgetItem,
    QVBoxLayout,
    QWidget,
)

from ..constantes import TIPOS_DESTINATARIO


class PanelEnvio(QWidget):
    enviar_solicitado = Signal()
    regenerar_solicitado = Signal()
    guardar_borrador_solicitado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        raiz = QVBoxLayout(self)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.setSpacing(10)

        # -- asunto -------------------------------------------------------- #
        grupo_asunto = QGroupBox("Asunto")
        interior = QVBoxLayout(grupo_asunto)
        self.asunto = QLineEdit()
        self.asunto.setPlaceholderText("Asunto del correo")
        interior.addWidget(self.asunto)
        raiz.addWidget(grupo_asunto)

        # -- destinatarios ------------------------------------------------- #
        grupo_destinos = QGroupBox("Destinatarios")
        interior = QVBoxLayout(grupo_destinos)
        self.tabla = QTableWidget(0, 3)
        self.tabla.setHorizontalHeaderLabels(["Enviar", "Correo", "Tipo"])
        self.tabla.verticalHeader().setVisible(False)
        self.tabla.setAlternatingRowColors(True)
        self.tabla.setSelectionBehavior(QAbstractItemView.SelectionBehavior.SelectRows)
        self.tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
        cabecera = self.tabla.horizontalHeader()
        cabecera.setSectionResizeMode(0, QHeaderView.ResizeMode.ResizeToContents)
        cabecera.setSectionResizeMode(1, QHeaderView.ResizeMode.Stretch)
        cabecera.setSectionResizeMode(2, QHeaderView.ResizeMode.Fixed)
        self.tabla.setColumnWidth(2, 86)
        # Las filas llevan un desplegable dentro: con el alto por defecto (30 px)
        # el control se recorta por abajo.
        self.tabla.verticalHeader().setDefaultSectionSize(34)
        self.tabla.setMinimumHeight(150)
        interior.addWidget(self.tabla)

        fila = QHBoxLayout()
        self.btn_agregar = QPushButton("Añadir correo")
        self.btn_agregar.setObjectName("Plano")
        self.btn_agregar.clicked.connect(self._agregar)
        self.btn_quitar = QPushButton("Quitar")
        self.btn_quitar.setObjectName("Plano")
        self.btn_quitar.clicked.connect(self._quitar)
        fila.addWidget(self.btn_agregar)
        fila.addWidget(self.btn_quitar)
        fila.addStretch(1)
        interior.addLayout(fila)
        raiz.addWidget(grupo_destinos)

        # -- adjuntos ------------------------------------------------------ #
        grupo_adjuntos = QGroupBox("Adjuntos")
        interior = QVBoxLayout(grupo_adjuntos)
        self.lista_adjuntos = QListWidget()
        self.lista_adjuntos.setMaximumHeight(90)
        interior.addWidget(self.lista_adjuntos)
        fila = QHBoxLayout()
        boton = QPushButton("Añadir archivo")
        boton.setObjectName("Plano")
        boton.clicked.connect(self._agregar_adjunto)
        fila.addWidget(boton)
        boton = QPushButton("Quitar")
        boton.setObjectName("Plano")
        boton.clicked.connect(self._quitar_adjunto)
        fila.addWidget(boton)
        fila.addStretch(1)
        interior.addLayout(fila)
        raiz.addWidget(grupo_adjuntos)

        raiz.addStretch(1)

        # -- acciones ------------------------------------------------------ #
        self.aviso = QLabel("")
        self.aviso.setWordWrap(True)
        self.aviso.setObjectName("Subtitulo")
        raiz.addWidget(self.aviso)

        fila = QHBoxLayout()
        self.btn_regenerar = QPushButton("Regenerar correo")
        self.btn_regenerar.clicked.connect(self.regenerar_solicitado.emit)
        self.btn_borrador = QPushButton("Guardar .eml")
        self.btn_borrador.clicked.connect(self.guardar_borrador_solicitado.emit)
        fila.addWidget(self.btn_regenerar)
        fila.addWidget(self.btn_borrador)
        raiz.addLayout(fila)

        self.btn_enviar = QPushButton("ENVIAR CORREO")
        self.btn_enviar.setObjectName("Primario")
        self.btn_enviar.setMinimumHeight(42)
        self.btn_enviar.clicked.connect(self.enviar_solicitado.emit)
        raiz.addWidget(self.btn_enviar)

    # -- destinatarios ------------------------------------------------------ #
    def cargar_destinatarios(self, destinatarios: list) -> None:
        self.tabla.setRowCount(0)
        for destinatario in destinatarios:
            self._insertar(destinatario.correo, destinatario.tipo, destinatario.activo,
                           destinatario.nombre)

    def _insertar(self, correo: str, tipo: str = "Para", activo: bool = True,
                  nombre: str = "") -> None:
        fila = self.tabla.rowCount()
        self.tabla.insertRow(fila)

        marca = QTableWidgetItem()
        marca.setFlags(Qt.ItemFlag.ItemIsUserCheckable | Qt.ItemFlag.ItemIsEnabled)
        marca.setCheckState(Qt.CheckState.Checked if activo else Qt.CheckState.Unchecked)
        self.tabla.setItem(fila, 0, marca)

        item = QTableWidgetItem(correo)
        if nombre:
            item.setToolTip(nombre)
        self.tabla.setItem(fila, 1, item)

        combo = QComboBox()
        combo.addItems(TIPOS_DESTINATARIO)
        combo.setCurrentText(tipo if tipo in TIPOS_DESTINATARIO else "Para")
        self.tabla.setCellWidget(fila, 2, combo)

    def _agregar(self) -> None:
        correo, aceptado = QInputDialog.getText(
            self, "Añadir destinatario", "Correo electrónico:"
        )
        if aceptado and correo.strip():
            self._insertar(correo.strip())

    def _quitar(self) -> None:
        filas = sorted({i.row() for i in self.tabla.selectedIndexes()}, reverse=True)
        for fila in filas:
            self.tabla.removeRow(fila)

    def destinatarios(self) -> tuple[list[str], list[str], list[str]]:
        para: list[str] = []
        cc: list[str] = []
        cco: list[str] = []
        for fila in range(self.tabla.rowCount()):
            marca = self.tabla.item(fila, 0)
            if marca is None or marca.checkState() != Qt.CheckState.Checked:
                continue
            correo_item = self.tabla.item(fila, 1)
            correo = correo_item.text().strip() if correo_item else ""
            if not correo:
                continue
            combo = self.tabla.cellWidget(fila, 2)
            tipo = combo.currentText() if combo else "Para"
            {"Para": para, "CC": cc, "CCO": cco}.get(tipo, para).append(correo)
        return para, cc, cco

    # -- adjuntos ----------------------------------------------------------- #
    def _agregar_adjunto(self) -> None:
        rutas, _ = QFileDialog.getOpenFileNames(self, "Seleccionar adjuntos")
        for ruta in rutas:
            item = QListWidgetItem(Path(ruta).name)
            item.setData(Qt.ItemDataRole.UserRole, ruta)
            item.setToolTip(ruta)
            self.lista_adjuntos.addItem(item)

    def _quitar_adjunto(self) -> None:
        for item in self.lista_adjuntos.selectedItems():
            self.lista_adjuntos.takeItem(self.lista_adjuntos.row(item))

    def adjuntos(self) -> list[Path]:
        return [
            Path(self.lista_adjuntos.item(i).data(Qt.ItemDataRole.UserRole))
            for i in range(self.lista_adjuntos.count())
        ]

    # -- varios -------------------------------------------------------------- #
    def limpiar(self) -> None:
        self.asunto.clear()
        self.tabla.setRowCount(0)
        self.lista_adjuntos.clear()
        self.aviso.clear()

    def establecer_aviso(self, texto: str, color: str = "") -> None:
        self.aviso.setText(texto)
        self.aviso.setStyleSheet(f"color: {color};" if color else "")

    def habilitar_envio(self, habilitado: bool) -> None:
        self.btn_enviar.setEnabled(habilitado)
        self.btn_regenerar.setEnabled(habilitado)
        self.btn_borrador.setEnabled(habilitado)
