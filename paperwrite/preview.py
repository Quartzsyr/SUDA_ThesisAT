from __future__ import annotations

import fitz
from PyQt6.QtCore import Qt, pyqtSignal
from PyQt6.QtGui import QImage, QPixmap, QWheelEvent
from PyQt6.QtWidgets import QLabel, QVBoxLayout, QWidget
from qfluentwidgets import ScrollArea


class PdfPreviewWidget(ScrollArea):
    zoom_changed = pyqtSignal(float)

    def __init__(self, parent=None) -> None:
        super().__init__(parent)
        self.zoom = 1.0
        self.page_count = 0
        self._pdf_path: str | None = None
        self._fit_mode = "width"
        if hasattr(self, "setSmoothMode"):
            try:
                self.setSmoothMode(None)
            except TypeError:
                pass

        self._container = QWidget(self)
        self._layout = QVBoxLayout(self._container)
        self._layout.setAlignment(Qt.AlignmentFlag.AlignTop | Qt.AlignmentFlag.AlignHCenter)
        self._layout.setContentsMargins(18, 18, 18, 18)
        self._layout.setSpacing(18)
        self.setWidget(self._container)
        self.setWidgetResizable(True)

        self._placeholder = QLabel("预览将在这里显示", self._container)
        self._placeholder.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self._placeholder.setMinimumHeight(320)
        self._placeholder.setStyleSheet(
            "color: #6B7280; background: #FFFFFF; border: 1px dashed rgba(148,163,184,0.35);"
            "border-radius: 18px; padding: 28px; font-size: 13px;"
        )
        self._layout.addWidget(self._placeholder)

    def _clear_pages(self) -> None:
        while self._layout.count():
            item = self._layout.takeAt(0)
            widget = item.widget()
            if widget is not None:
                widget.deleteLater()

    def load_pdf(self, pdf_path: str) -> None:
        self._pdf_path = pdf_path
        doc = fitz.open(pdf_path)
        try:
            self.page_count = len(doc)
        finally:
            doc.close()
        self.render_pages()

    def set_zoom(self, zoom: float) -> None:
        self.zoom = max(0.45, min(zoom, 3.0))
        self._fit_mode = "manual"
        self.render_pages()
        self.zoom_changed.emit(self.zoom)

    def fit_width(self) -> None:
        self._fit_mode = "width"
        if hasattr(self, "setSmoothMode"):
            try:
                self.setSmoothMode(None)
            except TypeError:
                pass
        self.render_pages()
        self.zoom_changed.emit(self.zoom)

    def scroll_to_top(self) -> None:
        self.verticalScrollBar().setValue(0)

    def wheelEvent(self, event: QWheelEvent) -> None:
        if event.modifiers() & Qt.KeyboardModifier.ControlModifier:
            delta = event.angleDelta().y()
            self.set_zoom(self.zoom + (0.08 if delta > 0 else -0.08))
            event.accept()
            return
        delta = event.angleDelta().y()
        if delta:
            bar = self.verticalScrollBar()
            step = max(60, abs(delta))
            bar.setValue(bar.value() - step if delta > 0 else bar.value() + step)
            event.accept()
            return
        super().wheelEvent(event)

    def resizeEvent(self, event) -> None:
        super().resizeEvent(event)
        if self._pdf_path and self._fit_mode == "width":
            self.render_pages()
            self.zoom_changed.emit(self.zoom)

    def _compute_zoom(self, page) -> float:
        if self._fit_mode != "width":
            return self.zoom
        available_w = max(self.viewport().width() - 64, 280)
        self.zoom = max(0.45, min(available_w / page.rect.width, 2.2))
        return self.zoom

    def render_pages(self) -> None:
        if not self._pdf_path:
            return
        self._clear_pages()
        doc = fitz.open(self._pdf_path)
        try:
            for index, page in enumerate(doc):
                zoom = self._compute_zoom(page) if index == 0 else self.zoom
                pix = page.get_pixmap(matrix=fitz.Matrix(zoom, zoom), alpha=False)
                image = QImage(pix.samples, pix.width, pix.height, pix.stride, QImage.Format.Format_RGB888).copy()
                label = QLabel(self._container)
                label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                label.setStyleSheet(
                    "background: #FFFFFF; border-radius: 18px; border: 1px solid rgba(148,163,184,0.14); padding: 12px;"
                )
                label.setPixmap(QPixmap.fromImage(image))
                self._layout.addWidget(label, 0, Qt.AlignmentFlag.AlignHCenter)
        finally:
            doc.close()

