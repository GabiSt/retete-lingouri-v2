"""Stil vizual: culori si foi de stil Qt (QSS) folosite in toata interfata."""

CULOARE_BLEUMARIN = "#1f3b73"
CULOARE_EROARE = "#c0392b"
CULOARE_SUCCES = "#1f8a55"
CULOARE_AVERTISMENT = "#c98a1d"
CULOARE_GRI_TEXT = "#6b7280"
CULOARE_FUNDAL_SECTIUNE = "#f4f6f8"
CULOARE_BORDURA = "#d8dee3"

STIL_BUTON_PRINCIPAL = f"""
    QPushButton {{
        background-color: {CULOARE_BLEUMARIN};
        color: white;
        font-weight: 600;
        padding: 7px 14px;
        border-radius: 6px;
        border: none;
    }}
    QPushButton:hover {{ background-color: #16294f; }}
    QPushButton:disabled {{ background-color: #9aa7b8; }}
"""

STIL_BUTON_SECUNDAR = f"""
    QPushButton {{
        background-color: white;
        color: {CULOARE_BLEUMARIN};
        font-weight: 600;
        padding: 7px 14px;
        border-radius: 6px;
        border: 1px solid {CULOARE_BLEUMARIN};
    }}
    QPushButton:hover {{ background-color: #eef2f8; }}
"""

STIL_BUTON_PERICOL = f"""
    QPushButton {{
        background: transparent;
        color: {CULOARE_EROARE};
        border: none;
        font-weight: 700;
        padding: 2px 8px;
    }}
    QPushButton:hover {{ background-color: #fbeceb; border-radius: 4px; }}
"""

STIL_CAMP = f"""
    QLineEdit, QComboBox {{
        border: 1px solid {CULOARE_BORDURA};
        border-radius: 5px;
        padding: 5px 7px;
        background: white;
    }}
    QLineEdit:focus, QComboBox:focus {{ border: 1px solid {CULOARE_BLEUMARIN}; }}
"""

TAG_STYLES = {
    "good": f"color:{CULOARE_SUCCES}; background-color:#e5f5ec; border-radius:5px; padding:2px 8px; font-weight:600;",
    "warn": f"color:{CULOARE_AVERTISMENT}; background-color:#faf1de; border-radius:5px; padding:2px 8px; font-weight:600;",
    "bad": f"color:{CULOARE_EROARE}; background-color:#fbeceb; border-radius:5px; padding:2px 8px; font-weight:600;",
}


def _clear_layout(layout):
    """Goleste un QLayout, stergand toate widget-urile copil."""
    while layout.count():
        item = layout.takeAt(0)
        widget = item.widget()
        if widget is not None:
            widget.deleteLater()
