# -*- coding: utf-8 -*-
"""Handtekening plaatsen: tekenen met de muis of een afbeelding uploaden.

Volledig nieuwe feature (geen tkinter-equivalent) - puur visuele annotatie,
geen cryptografische ondertekening. De actieve tab op het moment van "OK"
bepaalt welke bron gebruikt wordt (tekenen of upload).
"""

from PySide6.QtCore import Qt
from PySide6.QtGui import QColor, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

WIDTH_PRESETS = {"Klein": 100.0, "Middel": 160.0, "Groot": 220.0}
PEN_WIDTH = 3


class SignaturePadWidget(QWidget):
    """Vrijhand-tekenvlak: sleep met de muis om een handtekening te zetten."""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setMinimumSize(400, 150)
        self.setCursor(Qt.CursorShape.CrossCursor)
        self._strokes = []
        self._current_stroke = None

    def clear(self):
        self._strokes = []
        self._current_stroke = None
        self.update()

    def has_content(self):
        return any(len(stroke) > 1 for stroke in self._strokes)

    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._current_stroke = [event.position()]
            self._strokes.append(self._current_stroke)

    def mouseMoveEvent(self, event):
        if self._current_stroke is not None:
            self._current_stroke.append(event.position())
            self.update()

    def mouseReleaseEvent(self, event):
        self._current_stroke = None

    def _pen(self):
        return QPen(QColor("black"), PEN_WIDTH, Qt.PenStyle.SolidLine, Qt.PenCapStyle.RoundCap, Qt.PenJoinStyle.RoundJoin)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("white"))
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(self._pen())
        for stroke in self._strokes:
            for i in range(1, len(stroke)):
                painter.drawLine(stroke[i - 1], stroke[i])

    def to_image(self) -> QImage:
        """Geef de tekening terug, bijgesneden op de getekende inhoud (transparante achtergrond)."""
        full = QImage(self.size(), QImage.Format.Format_ARGB32)
        full.fill(Qt.GlobalColor.transparent)
        painter = QPainter(full)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, True)
        painter.setPen(self._pen())
        for stroke in self._strokes:
            for i in range(1, len(stroke)):
                painter.drawLine(stroke[i - 1], stroke[i])
        painter.end()

        if not self.has_content():
            return full

        margin = 10
        xs = [p.x() for stroke in self._strokes for p in stroke]
        ys = [p.y() for stroke in self._strokes for p in stroke]
        x0 = max(0, int(min(xs)) - margin)
        y0 = max(0, int(min(ys)) - margin)
        x1 = min(self.width(), int(max(xs)) + margin)
        y1 = min(self.height(), int(max(ys)) + margin)
        return full.copy(x0, y0, x1 - x0, y1 - y0)


class SignatureDialog(QDialog):
    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Handtekening plaatsen")
        self.setMinimumWidth(460)
        self._uploaded_image = None

        layout = QVBoxLayout(self)

        self.tabs = QTabWidget(self)
        layout.addWidget(self.tabs)

        draw_tab = QWidget()
        draw_layout = QVBoxLayout(draw_tab)
        self.pad = SignaturePadWidget()
        draw_layout.addWidget(self.pad)
        clear_btn = QPushButton("Wissen", draw_tab)
        clear_btn.clicked.connect(self.pad.clear)
        draw_layout.addWidget(clear_btn)
        self.tabs.addTab(draw_tab, "Tekenen")

        upload_tab = QWidget()
        upload_layout = QVBoxLayout(upload_tab)
        choose_btn = QPushButton("Kies afbeelding...", upload_tab)
        choose_btn.clicked.connect(self._choose_image)
        upload_layout.addWidget(choose_btn)
        self.preview_label = QLabel("Geen afbeelding gekozen", upload_tab)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(150)
        upload_layout.addWidget(self.preview_label)
        self.tabs.addTab(upload_tab, "Afbeelding")

        size_row = QHBoxLayout()
        size_row.addWidget(QLabel("Formaat:", self))
        self.size_combo = QComboBox(self)
        self.size_combo.addItems(list(WIDTH_PRESETS.keys()))
        self.size_combo.setCurrentText("Middel")
        size_row.addWidget(self.size_combo)
        layout.addLayout(size_row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _choose_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, "Afbeelding kiezen", "", "Afbeeldingen (*.png *.jpg *.jpeg *.bmp)"
        )
        if not path:
            return
        image = QImage(path)
        if image.isNull():
            QMessageBox.critical(self, "Ongeldige afbeelding", "Kan deze afbeelding niet openen.")
            return
        self._uploaded_image = image
        self.preview_label.setPixmap(
            QPixmap.fromImage(image).scaled(
                200, 150, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
        )

    def _on_accept(self):
        if self.get_image() is None:
            QMessageBox.warning(self, "Geen handtekening", "Teken een handtekening of kies een afbeelding.")
            return
        self.accept()

    def get_image(self):
        """Geef de te plaatsen QImage terug, of None. De actieve tab bepaalt de bron."""
        if self.tabs.currentIndex() == 0:
            return self.pad.to_image() if self.pad.has_content() else None
        return self._uploaded_image

    def get_target_width_pt(self) -> float:
        return WIDTH_PRESETS.get(self.size_combo.currentText(), 160.0)
