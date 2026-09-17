"""Ventana principal de App Alertas."""
from __future__ import annotations

import os
from pathlib import Path

from PySide6.QtCore import Qt, QThread, Signal
from PySide6.QtGui import QAction
from PySide6.QtWidgets import (
    QAbstractItemView,
    QApplication,
    QDialog,
    QFileDialog,
    QHBoxLayout,
    QHeaderView,
    QLabel,
    QLineEdit,
    QListWidget,
    QListWidgetItem,
    QMainWindow,
    QMessageBox,
    QPushButton,
    QSizePolicy,
    QSplitter,
    QTableWidget,
    QTableWidgetItem,
    QToolBar,
    QVBoxLayout,
    QWidget,
)

from .. import bd as modulo_bd
from .. import chart, config, email_builder, historial, lectura, sender
from ..constantes import NOMBRE_BD
from .editor_correo import EditorCorreo
from .panel_datos import PanelDatos
from .panel_envio import PanelEnvio
from .theme import qss


class Tarea(QThread):
    """Ejecuta una funcion en segundo plano para no congelar la interfaz."""

    listo = Signal(object)
    fallo = Signal(str)

    def __init__(self, funcion, parent=None):
        super().__init__(parent)
        self._funcion = funcion

    def run(self) -> None:                                      # noqa: D102
        try:
            self.listo.emit(self._funcion())
        except Exception as exc:                                # noqa: BLE001
            self.fallo.emit(str(exc))


class DialogoHistorial(QDialog):
    def __init__(self, filas: list[dict], parent=None):
        super().__init__(parent)
        self.setWindowTitle("Historial de envíos")
        self.resize(900, 460)
        disposicion = QVBoxLayout(self)
        if not filas:
            disposicion.addWidget(QLabel("Todavía no se ha enviado ningún correo."))
        else:
            columnas = list(filas[0].keys())
            tabla = QTableWidget(len(filas), len(columnas))
            tabla.setHorizontalHeaderLabels(columnas)
            tabla.verticalHeader().setVisible(False)
            tabla.setEditTriggers(QAbstractItemView.EditTrigger.NoEditTriggers)
            tabla.setAlternatingRowColors(True)
            for f, fila in enumerate(filas):
                for c, columna in enumerate(columnas):
                    tabla.setItem(f, c, QTableWidgetItem(str(fila.get(columna, ""))))
            tabla.horizontalHeader().setSectionResizeMode(
                QHeaderView.ResizeMode.ResizeToContents
            )
            disposicion.addWidget(tabla)
        cerrar = QPushButton("Cerrar")
        cerrar.clicked.connect(self.accept)
        fila_botones = QHBoxLayout()
        fila_botones.addStretch(1)
        fila_botones.addWidget(cerrar)
        disposicion.addLayout(fila_botones)


class VentanaPrincipal(QMainWindow):
    def __init__(self):
        super().__init__()
        self.log = config.configurar_log()
        self.bd: modulo_bd.BaseDatos | None = None
        self.datos = None
        self.png: bytes | None = None
        self.firma: bytes | None = None
        self.firma_subtipo: str = "png"
        self.oscuro = False
        self._tareas: list[Tarea] = []

        self.setWindowTitle("App Alertas — Reporte de avance y entregables")
        self.resize(1560, 920)
        self._construir()
        self.aplicar_tema(False)
        self.cargar_bd(inicial=True)

    # ------------------------------------------------------------------ #
    # Construccion
    # ------------------------------------------------------------------ #
    def _construir(self) -> None:
        barra = QToolBar()
        barra.setMovable(False)
        self.addToolBar(barra)

        self.accion_actualizar = QAction("⟳  Actualizar", self)
        self.accion_actualizar.setShortcut("F5")
        self.accion_actualizar.setToolTip(
            "Vuelve a leer el Excel de base de datos y el Excel del proyecto (F5)"
        )
        self.accion_actualizar.triggered.connect(lambda: self.cargar_bd())
        barra.addAction(self.accion_actualizar)
        barra.addSeparator()

        self.etiqueta_bd = QLabel("Base de datos: -")
        self.etiqueta_bd.setObjectName("Subtitulo")
        barra.addWidget(self.etiqueta_bd)

        espaciador = QWidget()
        espaciador.setSizePolicy(
            QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Preferred
        )
        barra.addWidget(espaciador)

        for texto, ayuda, funcion in (
            ("Abrir BD", "Abre el Excel de base de datos en Excel", self._abrir_bd),
            ("Cambiar BD", "Elige otro Excel de base de datos", self._elegir_bd),
            ("Probar SMTP", "Comprueba la conexión con el servidor de correo",
             self._probar_smtp),
            ("Historial", "Correos enviados desde esta aplicación", self._ver_historial),
        ):
            accion = QAction(texto, self)
            accion.setToolTip(ayuda)
            accion.triggered.connect(funcion)
            barra.addAction(accion)

        self.accion_tema = QAction("◑  Tema", self)
        self.accion_tema.setToolTip("Alterna entre tema claro y oscuro")
        self.accion_tema.triggered.connect(lambda: self.aplicar_tema(not self.oscuro))
        barra.addAction(self.accion_tema)

        # -- panel izquierdo ------------------------------------------------ #
        izquierda = QWidget()
        disposicion = QVBoxLayout(izquierda)
        disposicion.setContentsMargins(10, 10, 4, 10)
        titulo = QLabel("Proyectos")
        titulo.setObjectName("Titulo")
        disposicion.addWidget(titulo)
        self.buscador = QLineEdit()
        self.buscador.setPlaceholderText("Buscar por código o nombre…")
        self.buscador.textChanged.connect(self._filtrar_proyectos)
        disposicion.addWidget(self.buscador)
        self.lista = QListWidget()
        self.lista.currentItemChanged.connect(self._proyecto_elegido)
        disposicion.addWidget(self.lista, 1)
        self.etiqueta_avisos = QLabel("")
        self.etiqueta_avisos.setWordWrap(True)
        self.etiqueta_avisos.setObjectName("Subtitulo")
        disposicion.addWidget(self.etiqueta_avisos)

        # -- centro --------------------------------------------------------- #
        self.editor = EditorCorreo()
        self.panel_datos = PanelDatos()
        contenedor_datos = QWidget()
        interior = QVBoxLayout(contenedor_datos)
        interior.setContentsMargins(12, 12, 12, 12)
        interior.addWidget(self.panel_datos)
        self.editor.agregar_pestana(contenedor_datos, "Datos y filtros")
        self.panel_datos.filtro_cambiado.connect(self._filtro_cambiado)

        centro = QWidget()
        disposicion = QVBoxLayout(centro)
        disposicion.setContentsMargins(4, 10, 4, 10)
        disposicion.addWidget(self.editor)

        # -- panel derecho --------------------------------------------------- #
        self.panel_envio = PanelEnvio()
        self.panel_envio.enviar_solicitado.connect(self._enviar)
        self.panel_envio.regenerar_solicitado.connect(
            lambda: self._regenerar(preguntar=True, ir_a_correo=True))
        self.panel_envio.guardar_borrador_solicitado.connect(self._guardar_borrador)
        derecha = QWidget()
        disposicion = QVBoxLayout(derecha)
        disposicion.setContentsMargins(4, 10, 10, 10)
        disposicion.addWidget(self.panel_envio)

        divisor = QSplitter(Qt.Orientation.Horizontal)
        divisor.addWidget(izquierda)
        divisor.addWidget(centro)
        divisor.addWidget(derecha)
        divisor.setStretchFactor(1, 1)
        divisor.setSizes([290, 880, 380])
        self.setCentralWidget(divisor)

        self.estado = self.statusBar()
        self.estado_smtp = QLabel("")
        self.estado.addPermanentWidget(self.estado_smtp)
        self.estado.showMessage("Listo")

    def aplicar_tema(self, oscuro: bool) -> None:
        self.oscuro = oscuro
        QApplication.instance().setStyleSheet(qss(oscuro))

    # ------------------------------------------------------------------ #
    # Utilidades
    # ------------------------------------------------------------------ #
    def _ocupado(self, ocupado: bool) -> None:
        self.accion_actualizar.setEnabled(not ocupado)
        self.lista.setEnabled(not ocupado)
        self.panel_envio.habilitar_envio(not ocupado)

    def _lanzar(self, funcion, al_terminar, mensaje: str) -> None:
        self.estado.showMessage(mensaje)
        self._ocupado(True)

        tarea = Tarea(funcion, self)
        self._tareas.append(tarea)

        def terminado(resultado):
            self._ocupado(False)
            al_terminar(resultado)

        def fallido(mensaje_error):
            self._ocupado(False)
            self.estado.showMessage("Se produjo un error")
            self.log.error("Error en tarea: %s", mensaje_error)
            QMessageBox.critical(self, "Error", mensaje_error)

        def limpiar():
            if tarea in self._tareas:
                self._tareas.remove(tarea)

        tarea.listo.connect(terminado)
        tarea.fallo.connect(fallido)
        tarea.finished.connect(limpiar)
        tarea.start()

    def _proyecto_actual(self) -> modulo_bd.Proyecto | None:
        item = self.lista.currentItem()
        return item.data(Qt.ItemDataRole.UserRole) if item else None

    # ------------------------------------------------------------------ #
    # Carga de la base de datos
    # ------------------------------------------------------------------ #
    def cargar_bd(self, inicial: bool = False) -> None:
        ruta = config.ruta_bd()
        if ruta is None:
            if inicial:
                self._pedir_bd()
            else:
                QMessageBox.warning(
                    self, "Base de datos",
                    f"No se encuentra {NOMBRE_BD}. Usa «Cambiar BD» para indicar su ruta.",
                )
            return

        codigo_previo = getattr(self._proyecto_actual(), "codigo", None)
        self._lanzar(
            lambda: modulo_bd.cargar(ruta),
            lambda resultado: self._bd_cargada(resultado, codigo_previo),
            f"Leyendo {Path(ruta).name}…",
        )

    def _pedir_bd(self) -> None:
        respuesta = QMessageBox.question(
            self, "Base de datos no encontrada",
            f"No se encontró {NOMBRE_BD} junto a la aplicación.\n\n"
            "¿Quieres buscarlo en otra carpeta?\n"
            "Si respondes «No», genera uno nuevo ejecutando crear_bd.py.",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
        )
        if respuesta == QMessageBox.StandardButton.Yes:
            self._elegir_bd()

    def _elegir_bd(self) -> None:
        ruta, _ = QFileDialog.getOpenFileName(
            self, "Selecciona el Excel de base de datos", "", "Excel (*.xlsx *.xlsm)"
        )
        if ruta:
            config.fijar_ruta_bd(Path(ruta))
            self.cargar_bd()

    def _abrir_bd(self) -> None:
        ruta = config.ruta_bd()
        if ruta:
            self._abrir_archivo(ruta)

    def _abrir_archivo(self, ruta: Path) -> None:
        try:
            os.startfile(str(ruta))                             # noqa: S606
        except OSError as exc:
            QMessageBox.warning(self, "No se pudo abrir", f"{ruta}\n\n{exc}")

    def _cargar_firma(self) -> None:
        """Imagen de firma para el pie del correo; se relee en cada Actualizar."""
        self.firma, self.firma_subtipo = None, "png"
        ruta = config.ruta_firma(self.bd.cfg("Ruta firma") if self.bd else "")
        if ruta is None:
            return
        try:
            self.firma = ruta.read_bytes()
        except OSError as exc:
            self.log.warning("No se pudo leer la firma %s: %s", ruta, exc)
            return
        self.firma_subtipo = "jpeg" if ruta.suffix.lower() in (".jpg", ".jpeg") else "png"

    def _bd_cargada(self, base: modulo_bd.BaseDatos, codigo_previo: str | None) -> None:
        self.bd = base
        self._cargar_firma()
        self.etiqueta_bd.setText(f"Base de datos: {base.ruta}")
        self._pintar_proyectos(codigo_previo)

        if base.smtp.configurado:
            self.estado_smtp.setText(
                f"SMTP: {base.smtp.servidor}:{base.smtp.puerto} ({base.smtp.seguridad})"
            )
            self.estado_smtp.setStyleSheet("")
        else:
            self.estado_smtp.setText("SMTP sin configurar")
            self.estado_smtp.setStyleSheet("color:#C0392B;")

        if base.avisos:
            self.etiqueta_avisos.setText("⚠ " + "\n⚠ ".join(base.avisos[:6]))
            self.etiqueta_avisos.setStyleSheet("color:#B45309;")
        else:
            self.etiqueta_avisos.clear()

        self.estado.showMessage(f"{len(base.proyectos)} proyectos en la base de datos")

    def _pintar_proyectos(self, codigo_previo: str | None) -> None:
        self.lista.blockSignals(True)
        self.lista.clear()
        seleccionar = None
        for proyecto in self.bd.proyectos:
            texto = proyecto.codigo
            if proyecto.nombre:
                texto += f"\n{proyecto.nombre}"
            if not proyecto.activo:
                texto += "  (inactivo)"
            item = QListWidgetItem(texto)
            item.setData(Qt.ItemDataRole.UserRole, proyecto)
            item.setToolTip(str(proyecto.ruta_excel))
            if not proyecto.activo:
                item.setForeground(Qt.GlobalColor.gray)
            self.lista.addItem(item)
            if codigo_previo and proyecto.codigo == codigo_previo:
                seleccionar = item
        self.lista.blockSignals(False)

        self._filtrar_proyectos(self.buscador.text())
        if seleccionar is not None:
            self.lista.setCurrentItem(seleccionar)
        elif self.lista.count():
            self.lista.setCurrentRow(0)
        else:
            self._limpiar_vista()

    def _filtrar_proyectos(self, texto: str) -> None:
        objetivo = texto.strip().lower()
        for i in range(self.lista.count()):
            item = self.lista.item(i)
            item.setHidden(bool(objetivo) and objetivo not in item.text().lower())

    def _limpiar_vista(self) -> None:
        self.datos = None
        self.png = None
        self.panel_datos.limpiar()
        self.panel_envio.limpiar()
        self.editor.vaciar("Selecciona un proyecto para generar el correo.")

    # ------------------------------------------------------------------ #
    # Carga de un proyecto
    # ------------------------------------------------------------------ #
    def _proyecto_elegido(self, actual, _anterior) -> None:
        if actual is None or self.bd is None:
            return
        proyecto = actual.data(Qt.ItemDataRole.UserRole)
        opciones = chart.OpcionesGrafico.desde_bd(self.bd)

        def trabajo():
            datos = lectura.cargar(proyecto)
            imagen = None
            if datos.ev.semanas:
                try:
                    imagen = chart.generar(datos.ev, opciones, titulo=datos.nombre)
                except Exception as exc:                        # noqa: BLE001
                    datos.avisos.append(f"No se pudo dibujar la Curva S: {exc}")
            return datos, imagen

        self._lanzar(
            trabajo, self._proyecto_cargado,
            f"Leyendo {proyecto.ruta_excel.name}…",
        )

    def _proyecto_cargado(self, resultado) -> None:
        datos, imagen = resultado
        self.datos = datos
        self.png = imagen

        sugeridos = datos.let.estatus_sugeridos(datos.proyecto.estado_filtro)
        decimales = self.bd.cfg_int("Decimales avance", 0)
        self.panel_datos.cargar(datos, sugeridos, decimales)
        self.panel_envio.cargar_destinatarios(
            self.bd.destinatarios_de(datos.proyecto.codigo)
        )

        if not sugeridos:
            self.panel_envio.establecer_aviso(
                f"No se encontró el estatus «{datos.proyecto.estado_filtro}» en la hoja "
                f"{datos.let.hoja}. Elige uno en la pestaña «Datos y filtros».",
                "#B45309",
            )
        elif datos.avisos:
            self.panel_envio.establecer_aviso(
                "Hay avisos de lectura. Revísalos en «Ver cómo se leyó el Excel».",
                "#B45309",
            )
        else:
            self.panel_envio.establecer_aviso("")

        self._regenerar(preguntar=False, ir_a_correo=True)
        self.estado.showMessage(
            f"{datos.proyecto.codigo} · semana {datos.ev.semana_corte or '-'} · "
            f"{len(datos.let.entregables)} filas leídas en {datos.let.hoja}"
        )

    # ------------------------------------------------------------------ #
    # Generacion del correo
    # ------------------------------------------------------------------ #
    def _filtro_cambiado(self) -> None:
        if self.datos is None:
            return
        self._regenerar(preguntar=True)

    def _regenerar(self, preguntar: bool = True, ir_a_correo: bool = False) -> None:
        """Rehace el correo. Solo cambia de pestana si la accion lo justifica:
        al abrir un proyecto o al pulsar «Regenerar», nunca al tocar un filtro."""
        if self.datos is None or self.bd is None:
            return
        if preguntar and self.editor.editado:
            respuesta = QMessageBox.question(
                self, "Regenerar correo",
                "Has editado el correo manualmente.\n\n"
                "¿Quieres volver a generarlo y descartar esos cambios?",
                QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
                QMessageBox.StandardButton.No,
            )
            if respuesta != QMessageBox.StandardButton.Yes:
                return

        entregables = self.panel_datos.entregables_seleccionados()
        html = email_builder.construir_html(
            self.bd, self.datos, entregables,
            incluir_grafico=self.png is not None,
            incluir_firma=self.firma is not None,
        )
        # Qt no sabe renderizar porcentajes en imagenes: la vista previa recibe
        # el equivalente en pixeles y se restaura el valor al enviar.
        self.editor.establecer_html(
            email_builder.ancho_para_vista(html, self.editor.ancho_util()),
            {
                email_builder.CID_GRAFICO: self.png,
                email_builder.CID_FIRMA: self.firma,
            },
        )
        self.panel_envio.asunto.setText(
            email_builder.construir_asunto(self.bd, self.datos, entregables)
        )
        self.panel_envio.asunto.setCursorPosition(0)
        if ir_a_correo:
            self.editor.mostrar_correo()

    # ------------------------------------------------------------------ #
    # Envio
    # ------------------------------------------------------------------ #
    def _armar_envio(self) -> sender.Envio | None:
        if self.datos is None or self.bd is None:
            return None
        para, cc, cco = self.panel_envio.destinatarios()
        if self.bd.smtp.cco_fijo and self.bd.smtp.cco_fijo not in cco:
            cco.append(self.bd.smtp.cco_fijo)
        html = email_builder.ancho_para_correo(
            self.editor.html(),
            self.bd.cfg("Ancho imagen en el correo"),
        )
        return sender.Envio(
            asunto=self.panel_envio.asunto.text().strip(),
            html=html,
            texto=email_builder.a_texto_plano(html),
            imagen=self.png if "cid:" + email_builder.CID_GRAFICO in html else None,
            firma=self.firma if "cid:" + email_builder.CID_FIRMA in html else None,
            firma_subtipo=self.firma_subtipo,
            para=para, cc=cc, cco=cco,
            adjuntos=self.panel_envio.adjuntos(),
        )

    def _enviar(self) -> None:
        envio = self._armar_envio()
        if envio is None:
            QMessageBox.information(self, "Enviar", "Primero selecciona un proyecto.")
            return

        problemas = self.bd.smtp.problemas()
        if problemas:
            QMessageBox.warning(
                self, "SMTP incompleto",
                "Faltan estos datos en la hoja SMTP de la base de datos:\n\n• "
                + "\n• ".join(problemas),
            )
            return
        problemas = envio.problemas()
        if problemas:
            QMessageBox.warning(self, "Revisa el correo", "\n".join(problemas))
            return

        detalle = (
            f"Proyecto: {self.datos.proyecto.codigo}\n"
            f"Asunto: {envio.asunto}\n\n"
            f"Para ({len(envio.para)}): {', '.join(envio.para)}\n"
            f"CC ({len(envio.cc)}): {', '.join(envio.cc) or '—'}\n"
            f"CCO ({len(envio.cco)}): {', '.join(envio.cco) or '—'}\n\n"
            f"Entregables en la tabla: "
            f"{len(self.panel_datos.entregables_seleccionados())}"
        )
        respuesta = QMessageBox.question(
            self, "Confirmar envío", detalle + "\n\n¿Enviar ahora?",
            QMessageBox.StandardButton.Yes | QMessageBox.StandardButton.No,
            QMessageBox.StandardButton.No,
        )
        if respuesta != QMessageBox.StandardButton.Yes:
            return

        cfg = self.bd.smtp
        datos = self.datos
        entregables = len(self.panel_datos.entregables_seleccionados())

        def trabajo():
            sender.enviar(cfg, envio)
            return True

        def terminado(_resultado):
            historial.registrar(
                datos.proyecto.codigo, datos.nombre, envio.asunto,
                envio.para, envio.cc, envio.cco,
                datos.ev.semana_corte, entregables, "Enviado",
            )
            self.log.info("Correo enviado: %s -> %s", datos.proyecto.codigo, envio.para)
            self.estado.showMessage("Correo enviado correctamente")
            QMessageBox.information(
                self, "Enviado",
                f"El correo de {datos.proyecto.codigo} se envió correctamente.",
            )

        self._lanzar(trabajo, terminado, "Enviando correo…")

    def _guardar_borrador(self) -> None:
        envio = self._armar_envio()
        if envio is None or self.bd is None:
            return
        sugerido = f"{self.datos.proyecto.codigo} - reporte.eml"
        ruta, _ = QFileDialog.getSaveFileName(
            self, "Guardar borrador", sugerido, "Correo (*.eml)"
        )
        if not ruta:
            return
        try:
            sender.guardar_eml(self.bd.smtp, envio, Path(ruta))
        except Exception as exc:                                # noqa: BLE001
            QMessageBox.critical(self, "Error", f"No se pudo guardar el borrador:\n{exc}")
            return
        self.estado.showMessage(f"Borrador guardado en {ruta}")

    def _probar_smtp(self) -> None:
        if self.bd is None:
            return
        cfg = self.bd.smtp
        self._lanzar(
            lambda: sender.probar_conexion(cfg),
            lambda mensaje: QMessageBox.information(self, "SMTP", mensaje),
            "Probando la conexión SMTP…",
        )

    def _ver_historial(self) -> None:
        DialogoHistorial(historial.ultimos(200), self).exec()

    # ------------------------------------------------------------------ #
    def closeEvent(self, evento) -> None:                       # noqa: N802
        for tarea in list(self._tareas):
            tarea.wait(5000)
        super().closeEvent(evento)
