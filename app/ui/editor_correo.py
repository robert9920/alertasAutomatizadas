"""Editor del correo: vista WYSIWYG mas pestana de codigo HTML."""
from __future__ import annotations

from PySide6.QtCore import QByteArray, QUrl, Signal
from PySide6.QtGui import QFont, QImage, QTextCharFormat, QTextDocument, QTextListFormat
from PySide6.QtWidgets import (
    QColorDialog,
    QComboBox,
    QPlainTextEdit,
    QTabWidget,
    QTextEdit,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from ..email_builder import CID_GRAFICO, limpiar_html_editor

TAMANOS = ["8", "9", "10", "11", "12", "14", "16", "18"]


class EditorCorreo(QWidget):
    """Mantiene sincronizadas la vista visual y la vista HTML."""

    modificado = Signal()

    def __init__(self, parent=None):
        super().__init__(parent)
        self._cargando = False
        self._editado = False
        self._imagen: QImage | None = None

        self._fuente = 0          # 0 = vista visual, 1 = codigo HTML

        self.visual = QTextEdit()
        self.visual.setObjectName("Papel")
        self.visual.setAcceptRichText(True)
        self.visual.setAutoFormatting(QTextEdit.AutoFormattingFlag.AutoNone)
        self.visual.textChanged.connect(lambda: self._marcar_editado(0))

        self.codigo = QPlainTextEdit()
        self.codigo.setFont(QFont("Consolas", 10))
        self.codigo.setLineWrapMode(QPlainTextEdit.LineWrapMode.NoWrap)
        self.codigo.textChanged.connect(lambda: self._marcar_editado(1))

        self.barra = self._crear_barra()

        visual_contenedor = QWidget()
        disposicion = QVBoxLayout(visual_contenedor)
        disposicion.setContentsMargins(0, 0, 0, 0)
        disposicion.setSpacing(0)
        disposicion.addWidget(self.barra)
        disposicion.addWidget(self.visual)

        self.pestanas = QTabWidget()
        self.pestanas.addTab(visual_contenedor, "Correo")
        self.pestanas.addTab(self.codigo, "HTML")
        self.pestanas.currentChanged.connect(self._cambiar_pestana)

        raiz = QVBoxLayout(self)
        raiz.setContentsMargins(0, 0, 0, 0)
        raiz.addWidget(self.pestanas)

    # -- construccion ----------------------------------------------------- #
    def _crear_barra(self) -> QToolBar:
        barra = QToolBar()
        barra.setMovable(False)

        def accion(texto, ayuda, funcion, atajo=None, alternable=False):
            act = barra.addAction(texto, funcion)
            act.setToolTip(ayuda)
            if atajo:
                act.setShortcut(atajo)
            act.setCheckable(alternable)
            return act

        self.act_negrita = accion("N", "Negrita (Ctrl+B)", self._negrita, "Ctrl+B", True)
        self.act_cursiva = accion("K", "Cursiva (Ctrl+I)", self._cursiva, "Ctrl+I", True)
        self.act_subrayado = accion("S", "Subrayado (Ctrl+U)", self._subrayado, "Ctrl+U", True)
        fuente = self.act_negrita.font()
        fuente.setBold(True)
        self.act_negrita.setFont(fuente)
        barra.addSeparator()

        self.combo_tamano = QComboBox()
        self.combo_tamano.addItems(TAMANOS)
        self.combo_tamano.setCurrentText("11")
        self.combo_tamano.setFixedWidth(62)
        self.combo_tamano.setToolTip("Tamaño de letra")
        self.combo_tamano.currentTextChanged.connect(self._tamano)
        barra.addWidget(self.combo_tamano)

        accion("Color", "Color del texto", self._color)
        barra.addSeparator()
        accion("• Lista", "Viñetas", self._vinetas)
        accion("Quitar formato", "Deja el texto sin formato", self._quitar_formato)
        barra.addSeparator()
        accion("Eliminar fila", "Elimina la fila de la tabla donde está el cursor",
               self._eliminar_fila)
        barra.addSeparator()
        accion("Deshacer", "Deshacer (Ctrl+Z)", self.visual.undo, "Ctrl+Z")
        accion("Rehacer", "Rehacer (Ctrl+Y)", self.visual.redo, "Ctrl+Y")
        barra.addSeparator()
        accion("A-", "Reducir la vista", self.visual.zoomOut)
        accion("A+", "Ampliar la vista", self.visual.zoomIn)

        self.visual.currentCharFormatChanged.connect(self._sincronizar_barra)
        return barra

    # -- formato ---------------------------------------------------------- #
    def _aplicar(self, formato: QTextCharFormat) -> None:
        cursor = self.visual.textCursor()
        cursor.mergeCharFormat(formato)
        self.visual.mergeCurrentCharFormat(formato)
        self.visual.setFocus()

    def _negrita(self) -> None:
        formato = QTextCharFormat()
        peso = QFont.Weight.Bold if self.act_negrita.isChecked() else QFont.Weight.Normal
        formato.setFontWeight(peso)
        self._aplicar(formato)

    def _cursiva(self) -> None:
        formato = QTextCharFormat()
        formato.setFontItalic(self.act_cursiva.isChecked())
        self._aplicar(formato)

    def _subrayado(self) -> None:
        formato = QTextCharFormat()
        formato.setFontUnderline(self.act_subrayado.isChecked())
        self._aplicar(formato)

    def _tamano(self, texto: str) -> None:
        try:
            tamano = float(texto)
        except ValueError:
            return
        formato = QTextCharFormat()
        formato.setFontPointSize(tamano)
        self._aplicar(formato)

    def _color(self) -> None:
        color = QColorDialog.getColor(self.visual.textColor(), self, "Color del texto")
        if color.isValid():
            formato = QTextCharFormat()
            formato.setForeground(color)
            self._aplicar(formato)

    def _vinetas(self) -> None:
        cursor = self.visual.textCursor()
        cursor.createList(QTextListFormat.Style.ListDisc)
        self.visual.setFocus()

    def _quitar_formato(self) -> None:
        cursor = self.visual.textCursor()
        if not cursor.hasSelection():
            return
        texto = cursor.selectedText().replace(chr(0x2029), chr(10))
        cursor.insertText(texto, QTextCharFormat())
        self.visual.setFocus()

    def _eliminar_fila(self) -> None:
        cursor = self.visual.textCursor()
        tabla = cursor.currentTable()
        if tabla is None:
            return
        celda = tabla.cellAt(cursor)
        if celda.isValid() and tabla.rows() > 1:
            tabla.removeRows(celda.row(), 1)
        self.visual.setFocus()

    def _sincronizar_barra(self, formato: QTextCharFormat) -> None:
        self.act_negrita.setChecked(formato.fontWeight() >= QFont.Weight.Bold)
        self.act_cursiva.setChecked(formato.fontItalic())
        self.act_subrayado.setChecked(formato.fontUnderline())
        tamano = formato.fontPointSize()
        if tamano > 0:
            self.combo_tamano.blockSignals(True)
            self.combo_tamano.setCurrentText(f"{tamano:g}")
            self.combo_tamano.blockSignals(False)

    # -- contenido -------------------------------------------------------- #
    def _registrar_imagen(self) -> None:
        if self._imagen is None:
            return
        self.visual.document().addResource(
            QTextDocument.ResourceType.ImageResource,
            QUrl(f"cid:{CID_GRAFICO}"),
            self._imagen,
        )

    def establecer_html(self, html: str, imagen: bytes | None = None) -> None:
        """Carga el correo generado y descarta la marca de edicion manual."""
        if imagen is not None:
            qimagen = QImage()
            if qimagen.loadFromData(QByteArray(imagen), "PNG"):
                self._imagen = qimagen
        self._cargando = True
        try:
            self._registrar_imagen()
            self.visual.setHtml(html)
            self.codigo.setPlainText(html)
        finally:
            self._cargando = False
        self._editado = False
        self._fuente = 0
        self.pestanas.setCurrentIndex(0)

    def agregar_pestana(self, widget: QWidget, titulo: str) -> int:
        """Permite a la ventana principal colgar aqui otras vistas."""
        return self.pestanas.addTab(widget, titulo)

    def html(self) -> str:
        if self._fuente == 1:
            return self.codigo.toPlainText()
        return limpiar_html_editor(self.visual.toHtml())

    @property
    def editado(self) -> bool:
        return self._editado

    def _marcar_editado(self, fuente: int) -> None:
        if self._cargando:
            return
        self._fuente = fuente
        self._editado = True
        self.modificado.emit()

    def _cambiar_pestana(self, indice: int) -> None:
        """Sincroniza las dos vistas segun cual se editó por última vez."""
        if self._cargando or indice > 1:
            return
        self._cargando = True
        try:
            if indice == 1 and self._fuente == 0:
                self.codigo.setPlainText(limpiar_html_editor(self.visual.toHtml()))
                self._fuente = 1
            elif indice == 0 and self._fuente == 1:
                self._registrar_imagen()
                self.visual.setHtml(self.codigo.toPlainText())
                self._fuente = 0
        finally:
            self._cargando = False

    def vaciar(self, mensaje: str = "") -> None:
        self._cargando = True
        try:
            self.visual.setHtml(
                f'<p style="color:#777; font-style:italic;">{mensaje}</p>' if mensaje else ""
            )
            self.codigo.setPlainText("")
        finally:
            self._cargando = False
        self._editado = False
        self._fuente = 0
