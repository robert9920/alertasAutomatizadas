"""Paleta y hoja de estilos (QSS) de la aplicacion."""
from __future__ import annotations

CLARO = {
    "fondo": "#F4F6F9",
    "superficie": "#FFFFFF",
    "superficie_alt": "#EEF1F6",
    "borde": "#DDE2EA",
    "texto": "#1B2430",
    "texto_suave": "#5B6674",
    "acento": "#2563EB",
    "acento_hover": "#1D4FD7",
    "acento_suave": "#E3EDFF",
    "exito": "#107C41",
    "aviso": "#B45309",
    "peligro": "#C0392B",
    "seleccion": "#DCE9FF",
}

OSCURO = {
    "fondo": "#14171D",
    "superficie": "#1C2028",
    "superficie_alt": "#242935",
    "borde": "#2E3542",
    "texto": "#E7EAEF",
    "texto_suave": "#9AA4B2",
    "acento": "#4C8DFF",
    "acento_hover": "#6BA1FF",
    "acento_suave": "#1E2A44",
    "exito": "#3FB950",
    "aviso": "#D29922",
    "peligro": "#F85149",
    "seleccion": "#24334D",
}


def paleta(oscuro: bool) -> dict[str, str]:
    return OSCURO if oscuro else CLARO


def qss(oscuro: bool) -> str:
    c = paleta(oscuro)
    return f"""
* {{
    font-family: "Segoe UI", "Calibri", sans-serif;
    font-size: 13px;
}}
QWidget {{
    background-color: {c['fondo']};
    color: {c['texto']};
}}
QFrame#Tarjeta, QGroupBox {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    border-radius: 10px;
}}
QGroupBox {{
    margin-top: 14px;
    padding: 14px 10px 10px 10px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: {c['texto_suave']};
}}
QLabel#Titulo {{
    font-size: 17px;
    font-weight: 700;
}}
QLabel#Subtitulo {{
    color: {c['texto_suave']};
    font-size: 12px;
}}
QLabel#Kpi {{
    font-size: 19px;
    font-weight: 700;
}}
QLabel#KpiTitulo {{
    color: {c['texto_suave']};
    font-size: 11px;
    text-transform: uppercase;
}}

QPushButton {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    border-radius: 7px;
    padding: 7px 14px;
    color: {c['texto']};
}}
QPushButton:hover {{ background-color: {c['superficie_alt']}; }}
QPushButton:pressed {{ background-color: {c['acento_suave']}; }}
QPushButton:disabled {{ color: {c['texto_suave']}; background-color: {c['superficie_alt']}; }}
QPushButton#Primario {{
    background-color: {c['acento']};
    border: 1px solid {c['acento']};
    color: #FFFFFF;
    font-weight: 600;
    padding: 9px 18px;
}}
QPushButton#Primario:hover {{ background-color: {c['acento_hover']}; }}
QPushButton#Primario:disabled {{ background-color: {c['borde']}; border-color: {c['borde']}; color: {c['texto_suave']}; }}
QPushButton#Plano {{
    background-color: transparent;
    border: none;
    padding: 6px 10px;
    color: {c['acento']};
}}
QPushButton#Plano:hover {{ background-color: {c['acento_suave']}; border-radius: 6px; }}
QToolButton {{
    background-color: transparent;
    border: 1px solid transparent;
    border-radius: 6px;
    padding: 4px 8px;
}}
QToolButton:hover {{ background-color: {c['superficie_alt']}; border-color: {c['borde']}; }}
QToolButton:checked {{ background-color: {c['acento_suave']}; border-color: {c['acento']}; }}

QLineEdit, QPlainTextEdit, QTextEdit, QComboBox, QSpinBox {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    border-radius: 7px;
    padding: 6px 8px;
    selection-background-color: {c['seleccion']};
    selection-color: {c['texto']};
}}
QLineEdit:focus, QPlainTextEdit:focus, QComboBox:focus {{ border-color: {c['acento']}; }}
/* El cuerpo del correo se ve siempre sobre papel blanco, como en Outlook */
QTextEdit#Papel {{
    background-color: #FFFFFF;
    color: #111111;
    border: 1px solid {c['borde']};
    border-radius: 0px;
    padding: 14px 18px;
    selection-background-color: #CCE0FF;
    selection-color: #111111;
}}
/* Dentro de una tabla el relleno de formulario deja el control sin sitio y lo recorta */
QTableWidget QComboBox {{
    padding: 1px 6px;
    border-radius: 5px;
}}
QComboBox::drop-down {{ border: none; width: 18px; }}
QComboBox QAbstractItemView {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    selection-background-color: {c['acento_suave']};
    outline: none;
}}

QListWidget, QTableWidget, QTreeWidget {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    border-radius: 9px;
    outline: none;
    alternate-background-color: {c['superficie_alt']};
}}
QListWidget::item {{
    padding: 9px 10px;
    border-bottom: 1px solid {c['borde']};
}}
QListWidget::item:selected {{
    background-color: {c['acento_suave']};
    color: {c['texto']};
    border-left: 3px solid {c['acento']};
}}
QListWidget::item:hover {{ background-color: {c['superficie_alt']}; }}
QTableWidget::item {{ padding: 4px 6px; }}
QTableWidget::item:selected {{ background-color: {c['acento_suave']}; color: {c['texto']}; }}
QHeaderView::section {{
    background-color: {c['superficie_alt']};
    color: {c['texto_suave']};
    border: none;
    border-bottom: 1px solid {c['borde']};
    border-right: 1px solid {c['borde']};
    padding: 7px 6px;
    font-weight: 600;
}}
QTableCornerButton::section {{ background-color: {c['superficie_alt']}; border: none; }}

QTabWidget::pane {{
    border: 1px solid {c['borde']};
    border-radius: 9px;
    background-color: {c['superficie']};
    top: -1px;
}}
QTabBar::tab {{
    background: transparent;
    color: {c['texto_suave']};
    padding: 8px 16px;
    margin-right: 4px;
    border: 1px solid transparent;
    border-top-left-radius: 8px;
    border-top-right-radius: 8px;
}}
QTabBar::tab:selected {{
    background: {c['superficie']};
    color: {c['texto']};
    border-color: {c['borde']};
    border-bottom-color: {c['superficie']};
    font-weight: 600;
}}
QTabBar::tab:hover:!selected {{ color: {c['texto']}; }}

QCheckBox {{ spacing: 7px; }}
QCheckBox::indicator {{
    width: 16px; height: 16px;
    border: 1px solid {c['borde']};
    border-radius: 4px;
    background-color: {c['superficie']};
}}
QCheckBox::indicator:checked {{
    background-color: {c['acento']};
    border-color: {c['acento']};
    image: none;
}}
QCheckBox::indicator:disabled {{ background-color: {c['superficie_alt']}; }}

QScrollBar:vertical {{ background: transparent; width: 11px; margin: 2px; }}
QScrollBar::handle:vertical {{
    background: {c['borde']}; border-radius: 5px; min-height: 28px;
}}
QScrollBar::handle:vertical:hover {{ background: {c['texto_suave']}; }}
QScrollBar:horizontal {{ background: transparent; height: 11px; margin: 2px; }}
QScrollBar::handle:horizontal {{
    background: {c['borde']}; border-radius: 5px; min-width: 28px;
}}
QScrollBar::add-line, QScrollBar::sub-line {{ height: 0; width: 0; }}
QScrollBar::add-page, QScrollBar::sub-page {{ background: transparent; }}

QSplitter::handle {{ background-color: transparent; }}
QSplitter::handle:horizontal {{ width: 8px; }}

QStatusBar {{ background-color: {c['superficie']}; border-top: 1px solid {c['borde']}; }}
QStatusBar::item {{ border: none; }}

QToolBar {{
    background-color: {c['superficie']};
    border-bottom: 1px solid {c['borde']};
    spacing: 6px;
    padding: 7px 10px;
}}
QMenu {{
    background-color: {c['superficie']};
    border: 1px solid {c['borde']};
    border-radius: 8px;
    padding: 4px;
}}
QMenu::item {{ padding: 6px 22px 6px 14px; border-radius: 5px; }}
QMenu::item:selected {{ background-color: {c['acento_suave']}; }}

QProgressBar {{
    border: none; background-color: {c['superficie_alt']};
    border-radius: 3px; height: 6px; text-align: center;
}}
QProgressBar::chunk {{ background-color: {c['acento']}; border-radius: 3px; }}
"""
