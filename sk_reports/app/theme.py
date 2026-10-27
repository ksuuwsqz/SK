"""
Спокойная официальная цветовая гамма и общие стили.
"""

# Основные цвета
COLOR_BG = "#F4F6F9"          # фон окна
COLOR_SIDEBAR = "#1F3A5F"     # тёмно-синий (официальный)
COLOR_SIDEBAR_HOVER = "#2B4C76"
COLOR_SIDEBAR_ACTIVE = "#3A6395"
COLOR_CARD = "#FFFFFF"
COLOR_TEXT = "#1E1E1E"
COLOR_MUTED = "#6B7280"
COLOR_ACCENT = "#2F6FBF"
COLOR_OK = "#1E7A3C"
COLOR_WARN = "#B7791F"
COLOR_ERR = "#B3261E"
COLOR = "#1E1E1E"

# Цвета для подсветки ошибок
ROW_ERROR = "#FBE9E7"
ROW_WARN = "#FFF6E0"
ROW_OK = "#E9F6EC"

APP_STYLESHEET = f"""
QWidget {{
    background: {COLOR_BG};
    color: {COLOR_TEXT};
    font-family: 'Segoe UI', 'Arial', sans-serif;
    font-size: 10.5pt;
}}
QFrame#Sidebar {{
    background-color: {COLOR_SIDEBAR};
    border: none;
}}
QPushButton#NavButton {{
    color: #E8EEF7;
    background: transparent;
    text-align: left;
    padding: 12px 18px;
    border: none;
    border-left: 4px solid transparent;
    font-size: 11pt;
}}
QPushButton#NavButton:hover {{
    background: {COLOR_SIDEBAR_HOVER};
}}
QPushButton#NavButton:checked {{
    background: {COLOR_SIDEBAR_ACTIVE};
    border-left: 4px solid #FFFFFF;
    font-weight: 600;
}}
QLabel#PageTitle {{
    font-size: 18pt;
    font-weight: 700;
    color: {COLOR_SIDEBAR};
}}
QLabel#CardTitle {{
    font-size: 10pt;
    color: {COLOR_MUTED};
}}
QLabel#CardValue {{
    font-size: 22pt;
    font-weight: 700;
    color: {COLOR_SIDEBAR};
}}
QFrame#Card {{
    background: {COLOR_CARD};
    border: 1px solid #E1E5EC;
    border-radius: 10px;
}}
QPushButton {{
    background: {COLOR_ACCENT};
    color: white;
    border: none;
    padding: 8px 16px;
    border-radius: 6px;
}}
QPushButton:hover {{ background: #3A7FD0; }}
QPushButton:disabled {{ background: #B9C3D2; }}
QPushButton#Secondary {{
    background: #E7ECF3; color: {COLOR_SIDEBAR};
}}
QPushButton#Secondary:hover {{ background: #D7DFE9; }}
QTableWidget {{
    background: white;
    border: 1px solid #E1E5EC;
    gridline-color: #EEF1F5;
    selection-background-color: #D6E4F5;
}}
QHeaderView::section {{
    background: #EEF1F5;
    padding: 6px;
    border: none;
    font-weight: 600;
    color: {COLOR_SIDEBAR};
}}
QComboBox, QLineEdit, QDateEdit, QTextEdit, QPlainTextEdit {{
    background: white;
    border: 1px solid #CBD3DF;
    border-radius: 6px;
    padding: 5px;
}}
QTextEdit, QPlainTextEdit {{ padding: 8px; }}
QComboBox:focus, QLineEdit:focus, QTextEdit:focus {{ border: 1px solid {COLOR_ACCENT}; }}
QGroupBox {{
    background: white;
    border: 1px solid #E1E5EC;
    border-radius: 10px;
    margin-top: 12px;
    padding: 12px;
    font-weight: 600;
}}
QGroupBox::title {{
    subcontrol-origin: margin;
    left: 12px;
    padding: 0 6px;
    color: {COLOR_SIDEBAR};
}}
QStatusBar {{ background: #E9EEF5; color: {COLOR_SIDEBAR}; }}
"""