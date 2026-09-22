# -*- coding: utf-8 -*-
"""Formulierveld-widgets: bouwt een echte QWidget per PyMuPDF form-widget.

Poort van _create_form_overlays uit NVict_Reader.py (regel 4320-4431), maar
met QGraphicsProxyWidget in plaats van canvas.create_window - zelfde idee
(een echte widget ingebed op de canvas/scene), andere UI-toolkit.

Widget-type-constanten zijn de stabiele PyMuPDF/PDF-waarden (fitz.PDF_WIDGET_TYPE_*),
hier als losse ints zodat deze module geen fitz-import nodig heeft.
"""

FIELD_TYPE_CHECKBOX = 2
FIELD_TYPE_COMBOBOX = 3
FIELD_TYPE_LISTBOX = 4
FIELD_TYPE_RADIOBUTTON = 5
FIELD_TYPE_TEXT = 7

MULTILINE_FLAG = 4096

from PySide6.QtWidgets import QButtonGroup, QCheckBox, QComboBox, QLineEdit, QPlainTextEdit, QRadioButton


def is_checked(widget) -> bool:
    """Bepaal of een checkbox/radiobutton-widget momenteel aan staat.

    Radiogroepen delen hun field_value (de naam van het geselecteerde kind);
    een specifieke radio is 'aan' als field_value gelijk is aan ZIJN eigen
    on-state-naam (widget.on_state()).
    """
    try:
        on_state = widget.on_state()
    except Exception:
        on_state = "Yes"
    value = widget.field_value
    if value is None:
        return False
    return str(value) == str(on_state)


def create_field_widget(widget, on_value_changed, radio_groups: dict):
    """Bouw de juiste QWidget voor dit PDF-formulierveld.

    `on_value_changed(xref, value)` wordt direct bij elke wijziging
    aangeroepen (geen aparte 'sla waarden op'-stap nodig zoals in tkinter -
    Qt-signals maken dat overbodig).
    """
    field_type = widget.field_type
    xref = widget.xref

    if field_type == FIELD_TYPE_TEXT and (widget.field_flags or 0) & MULTILINE_FLAG:
        w = QPlainTextEdit()
        w.setPlainText(widget.field_value or "")
        w.textChanged.connect(lambda: on_value_changed(xref, w.toPlainText()))
        return w

    if field_type == FIELD_TYPE_TEXT:
        w = QLineEdit()
        w.setText(widget.field_value or "")
        w.textChanged.connect(lambda text: on_value_changed(xref, text))
        return w

    if field_type == FIELD_TYPE_CHECKBOX:
        w = QCheckBox()
        w.setChecked(is_checked(widget))
        w.toggled.connect(lambda checked: on_value_changed(xref, checked))
        return w

    if field_type in (FIELD_TYPE_COMBOBOX, FIELD_TYPE_LISTBOX):
        w = QComboBox()
        choices = widget.choice_values or []
        w.addItems(choices)
        current = widget.field_value or ""
        if current in choices:
            w.setCurrentText(current)
        w.currentTextChanged.connect(lambda text: on_value_changed(xref, text))
        return w

    if field_type == FIELD_TYPE_RADIOBUTTON:
        w = QRadioButton()
        w.setChecked(is_checked(widget))
        group_name = widget.field_name
        group = radio_groups.get(group_name)
        if group is None:
            group = QButtonGroup()
            group.setExclusive(True)
            radio_groups[group_name] = group
        group.addButton(w)
        w.toggled.connect(lambda checked: on_value_changed(xref, checked))
        return w

    # Onbekend veldtype: fallback tekstveld met gele achtergrond, zoals tkinter.
    w = QLineEdit()
    w.setText(str(widget.field_value or ""))
    w.setStyleSheet("background-color: #ffffcc;")
    w.textChanged.connect(lambda text: on_value_changed(xref, text))
    return w
