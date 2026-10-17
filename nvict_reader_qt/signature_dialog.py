# -*- coding: utf-8 -*-
"""Handtekening plaatsen: tekenen met de muis of een afbeelding uploaden.

Volledig nieuwe feature (geen tkinter-equivalent) - puur visuele annotatie,
geen cryptografische ondertekening. De actieve tab op het moment van "OK"
bepaalt welke bron gebruikt wordt (tekenen of upload).
"""

from PySide6.QtCore import QSize, Qt
from PySide6.QtGui import QColor, QIcon, QImage, QPainter, QPen, QPixmap
from PySide6.QtWidgets import (
    QCheckBox,
    QComboBox,
    QDialog,
    QDialogButtonBox,
    QFileDialog,
    QHBoxLayout,
    QLabel,
    QListView,
    QListWidget,
    QListWidgetItem,
    QMessageBox,
    QPushButton,
    QTabWidget,
    QVBoxLayout,
    QWidget,
)

from . import settings, signature_store, theme
from .i18n import tr

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
        self.setWindowTitle(tr("Handtekening plaatsen"))
        self.setMinimumWidth(460)
        self._uploaded_image = None

        layout = QVBoxLayout(self)

        self.tabs = QTabWidget(self)
        layout.addWidget(self.tabs)

        # Bewaarde handtekeningen: één klik en klaar.
        saved_tab = QWidget()
        saved_layout = QVBoxLayout(saved_tab)
        self.saved_list = QListWidget(saved_tab)
        self.saved_list.setViewMode(QListView.ViewMode.IconMode)
        self.saved_list.setIconSize(QSize(180, 70))
        self.saved_list.setResizeMode(QListView.ResizeMode.Adjust)
        self.saved_list.setMovement(QListView.Movement.Static)
        self.saved_list.setSpacing(6)
        self.saved_list.setMinimumHeight(150)
        self.saved_list.setUniformItemSizes(True)
        accent = theme.resolve_theme(settings.get_theme_mode())["ACCENT_COLOR"]
        self.saved_list.setStyleSheet(
            "QListWidget::item { border: 2px solid transparent; border-radius: 4px; }"
            f"QListWidget::item:selected {{ background: transparent; border: 2px solid {accent}; }}"
        )
        self.saved_list.itemDoubleClicked.connect(lambda _item: self._on_accept())
        saved_layout.addWidget(self.saved_list)
        delete_btn = QPushButton(tr("Geselecteerde verwijderen"), saved_tab)
        delete_btn.clicked.connect(self._delete_saved)
        saved_layout.addWidget(delete_btn)
        self._saved_tab = saved_tab

        draw_tab = QWidget()
        draw_layout = QVBoxLayout(draw_tab)
        self.pad = SignaturePadWidget()
        draw_layout.addWidget(self.pad)
        clear_btn = QPushButton(tr("Wissen"), draw_tab)
        clear_btn.clicked.connect(self.pad.clear)
        draw_layout.addWidget(clear_btn)
        self.tabs.addTab(draw_tab, tr("Tekenen"))

        upload_tab = QWidget()
        upload_layout = QVBoxLayout(upload_tab)
        choose_btn = QPushButton(tr("Kies afbeelding..."), upload_tab)
        choose_btn.clicked.connect(self._choose_image)
        upload_layout.addWidget(choose_btn)
        self.preview_label = QLabel(tr("Geen afbeelding gekozen"), upload_tab)
        self.preview_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.preview_label.setMinimumHeight(150)
        upload_layout.addWidget(self.preview_label)
        self.tabs.addTab(upload_tab, tr("Afbeelding"))

        self.remember_check = QCheckBox(tr("Handtekening bewaren voor de volgende keer"), self)
        self.remember_check.setChecked(True)
        layout.addWidget(self.remember_check)
        self.tabs.currentChanged.connect(self._on_tab_changed)
        self._load_saved()

        size_row = QHBoxLayout()
        size_row.addWidget(QLabel(tr("Formaat:"), self))
        self.size_combo = QComboBox(self)
        # Label vertaald, de Nederlandse sleutel als itemData (zie get_target_width_pt).
        for key in WIDTH_PRESETS:
            self.size_combo.addItem(tr(key), key)
        self.size_combo.setCurrentIndex(self.size_combo.findData("Middel"))
        size_row.addWidget(self.size_combo)
        layout.addLayout(size_row)

        buttons = QDialogButtonBox(QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel, self)
        buttons.accepted.connect(self._on_accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _load_saved(self):
        """Tabblad "Opgeslagen" vooraan tonen zodra er iets bewaard is."""
        self.saved_list.clear()
        for path, image in signature_store.list_saved():
            preview = QPixmap(180, 70)
            preview.fill(QColor("white"))
            painter = QPainter(preview)
            scaled = QPixmap.fromImage(image).scaled(
                170, 64, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
            painter.drawPixmap((180 - scaled.width()) // 2, (70 - scaled.height()) // 2, scaled)
            painter.end()
            icon = QIcon(preview)
            icon.addPixmap(preview, QIcon.Mode.Selected)  # niet blauw inkleuren; het kader toont de selectie
            item = QListWidgetItem(icon, "")
            item.setSizeHint(QSize(190, 80))  # geen lege tekstregel onder het plaatje
            item.setData(Qt.ItemDataRole.UserRole, path)
            self.saved_list.addItem(item)
        index = self.tabs.indexOf(self._saved_tab)
        if self.saved_list.count():
            if index < 0:
                self.tabs.insertTab(0, self._saved_tab, tr("Opgeslagen"))
            self.tabs.setCurrentIndex(0)
            self.saved_list.setCurrentRow(0)
        elif index >= 0:
            self.tabs.removeTab(index)
        self._on_tab_changed()

    def _delete_saved(self):
        item = self.saved_list.currentItem()
        if item is None:
            return
        signature_store.delete(item.data(Qt.ItemDataRole.UserRole))
        self._load_saved()

    def _on_tab_changed(self, _index=None):
        # Bewaren slaat alleen op een nieuwe handtekening.
        self.remember_check.setVisible(self.tabs.currentWidget() is not self._saved_tab)

    def _choose_image(self):
        path, _ = QFileDialog.getOpenFileName(
            self, tr("Afbeelding kiezen"), "", tr("Afbeeldingen (*.png *.jpg *.jpeg *.bmp)")
        )
        if not path:
            return
        image = QImage(path)
        if image.isNull():
            QMessageBox.critical(self, tr("Ongeldige afbeelding"), tr("Kan deze afbeelding niet openen."))
            return
        self._uploaded_image = image
        self.preview_label.setPixmap(
            QPixmap.fromImage(image).scaled(
                200, 150, Qt.AspectRatioMode.KeepAspectRatio, Qt.TransformationMode.SmoothTransformation
            )
        )

    def _on_accept(self):
        image = self.get_image()
        if image is None:
            QMessageBox.warning(self, tr("Geen handtekening"), tr("Teken een handtekening of kies een afbeelding."))
            return
        if self.tabs.currentWidget() is not self._saved_tab and self.remember_check.isChecked():
            try:
                signature_store.save(image)
            except OSError:
                pass  # bewaren is een gemak; plaatsen moet gewoon doorgaan
        self.accept()

    def get_image(self):
        """Geef de te plaatsen QImage terug, of None. De actieve tab bepaalt de bron."""
        current = self.tabs.currentWidget()
        if current is self._saved_tab:
            item = self.saved_list.currentItem()
            if item is None:
                return None
            image = QImage(item.data(Qt.ItemDataRole.UserRole))
            return None if image.isNull() else image
        if current is self.pad.parentWidget():
            return self.pad.to_image() if self.pad.has_content() else None
        return self._uploaded_image

    def get_target_width_pt(self) -> float:
        return WIDTH_PRESETS.get(self.size_combo.currentData(), 160.0)
