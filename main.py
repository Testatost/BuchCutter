import os
import re
import sys
import math
import ctypes
import tempfile
import time
from dataclasses import dataclass
from typing import List, Optional, Tuple

import fitz  # PyMuPDF
from PIL import Image, ImageDraw, ImageOps, ImageEnhance, ImageFilter, ImageQt

from PySide6.QtCore import Qt, QRectF, QPointF, Signal, QObject, QRunnable, QThreadPool
from PySide6.QtGui import QAction, QColor, QPainter, QPen, QPixmap, QPalette, QIcon, QKeySequence
from PySide6.QtWidgets import (
    QApplication, QMainWindow, QWidget, QFileDialog, QMessageBox, QVBoxLayout, QHBoxLayout,
    QPushButton, QLabel, QSplitter, QTableWidget, QTableWidgetItem, QHeaderView,
    QAbstractItemView, QCheckBox, QProgressBar, QToolButton, QMenu, QSizePolicy,
    QFrame
)

# =========================================================
# Utilities
# =========================================================

SUPPORTED_EXTENSIONS = {".png", ".jpg", ".jpeg", ".bmp", ".tif", ".tiff", ".pdf"}

def resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", os.path.abspath("."))
    return os.path.join(base_path, relative_path)

def polygon_area(poly: List[Tuple[float, float]]) -> float:
    if not poly or len(poly) < 3:
        return 0.0
    area = 0.0
    for i in range(len(poly)):
        x1, y1 = poly[i]
        x2, y2 = poly[(i + 1) % len(poly)]
        area += x1 * y2 - x2 * y1
    return abs(area) * 0.5

def clip_polygon_halfplane(polygon, a, b, c):
    if not polygon:
        return []

    def inside(pt):
        return a * pt[0] + b * pt[1] + c >= -1e-9

    def intersect(p1, p2):
        x1, y1 = p1
        x2, y2 = p2
        v1 = a * x1 + b * y1 + c
        v2 = a * x2 + b * y2 + c
        denom = v1 - v2
        if abs(denom) < 1e-12:
            return p2
        t = v1 / (v1 - v2)
        return (x1 + t * (x2 - x1), y1 + t * (y2 - y1))

    output = []
    S = polygon[-1]
    for E in polygon:
        if inside(E):
            if inside(S):
                output.append(E)
            else:
                output.append(intersect(S, E))
                output.append(E)
        else:
            if inside(S):
                output.append(intersect(S, E))
        S = E
    return output

def pil_to_qpixmap(img: Image.Image) -> QPixmap:
    if img.mode not in ("RGB", "RGBA"):
        img = img.convert("RGBA")
    qimage = ImageQt.ImageQt(img)
    return QPixmap.fromImage(qimage)

def render_pdf_page_to_pil(pdf_path: str, page_index: int, dpi: int = 300) -> Image.Image:
    doc = fitz.open(pdf_path)
    try:
        page = doc.load_page(page_index)
        zoom = dpi / 72.0
        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat, alpha=False)
        return Image.frombytes("RGB", [pix.width, pix.height], pix.samples)
    finally:
        doc.close()

PREVIEW_PDF_DPI = 144
EXPORT_PDF_DPI = 300

def load_image_or_pdf_page(path: str, page_index: Optional[int] = None, pdf_dpi: int = EXPORT_PDF_DPI) -> Image.Image:
    if page_index is None:
        return Image.open(path).convert("RGB")
    return render_pdf_page_to_pil(path, page_index, dpi=pdf_dpi)

def normalize_dropped_paths(raw_list: List[str]) -> List[str]:
    cleaned = []
    for path in raw_list:
        p = path.strip().strip('"').strip("'")
        if p:
            cleaned.append(p)
    return cleaned

# =========================================================
# Data models
# =========================================================

@dataclass
class ItemState:
    source_path: str
    page_index: Optional[int] = None
    display_name: str = ""
    selected: bool = False
    crop_enabled: bool = False
    split_enabled: bool = False
    color_mode: str = "RGB"  # "RGB" oder "GRAY"
    contrast_enabled: bool = False
    rotation_angle: float = 0.0

    def unique_key(self) -> str:
        if self.page_index is None:
            return self.source_path
        return f"{self.source_path}::page::{self.page_index}"

@dataclass
class Separator:
    cx: float
    cy: float
    angle: float = 0.0  # 0 = vertikal

    HANDLE_R = 8
    ROT_R = 12
    ROT_OFFSET = 30

    def direction_vector(self) -> Tuple[float, float]:
        # angle = 0 => vertical line
        return math.sin(self.angle), -math.cos(self.angle)

    def clipped_endpoints(self, w: float, h: float) -> Optional[Tuple[float, float, float, float]]:
        if w <= 1 or h <= 1:
            return None

        vx, vy = self.direction_vector()
        eps = 1e-9
        candidates = []

        if abs(vx) > eps:
            for x in (0.0, float(w)):
                t = (x - self.cx) / vx
                y = self.cy + t * vy
                if -1e-6 <= y <= h + 1e-6:
                    candidates.append((t, x, max(0.0, min(float(h), y))))

        if abs(vy) > eps:
            for y in (0.0, float(h)):
                t = (y - self.cy) / vy
                x = self.cx + t * vx
                if -1e-6 <= x <= w + 1e-6:
                    candidates.append((t, max(0.0, min(float(w), x)), y))

        if len(candidates) < 2:
            return None

        unique = []
        for t, x, y in candidates:
            found = False
            for _, ux, uy in unique:
                if abs(x - ux) < 1e-4 and abs(y - uy) < 1e-4:
                    found = True
                    break
            if not found:
                unique.append((t, x, y))

        if len(unique) < 2:
            return None

        unique.sort(key=lambda item: item[0])
        _, x1, y1 = unique[0]
        _, x2, y2 = unique[-1]
        return x1, y1, x2, y2

    def top_handle(self, w: float, h: float):
        pts = self.clipped_endpoints(w, h)
        if pts is None:
            return self.cx, self.cy
        x1, y1, x2, y2 = pts
        if y1 < y2 or (abs(y1 - y2) < 1e-6 and x1 <= x2):
            return x1, y1
        return x2, y2

    def bottom_handle(self, w: float, h: float):
        pts = self.clipped_endpoints(w, h)
        if pts is None:
            return self.cx, self.cy
        x1, y1, x2, y2 = pts
        if y1 < y2 or (abs(y1 - y2) < 1e-6 and x1 <= x2):
            return x2, y2
        return x1, y1

    def distance_to_line(self, px: float, py: float, w: float, h: float) -> float:
        pts = self.clipped_endpoints(w, h)
        if pts is None:
            return 1e9
        x1, y1, x2, y2 = pts
        vx = x2 - x1
        vy = y2 - y1
        wx = px - x1
        wy = py - y1
        denom = math.hypot(vx, vy)
        if denom == 0:
            return math.hypot(px - x1, py - y1)
        return abs(vx * wy - vy * wx) / denom

    def set_from_points(self, p1: Tuple[float, float], p2: Tuple[float, float]):
        x1, y1 = p1
        x2, y2 = p2
        self.cx = (x1 + x2) / 2.0
        self.cy = (y1 + y2) / 2.0
        dx = x2 - x1
        dy = y2 - y1
        if abs(dx) < 1e-9 and abs(dy) < 1e-9:
            return
        self.angle = math.atan2(dx, -dy)

    def move_by(self, dx: float, dy: float, w: float, h: float):
        self.cx = max(0.0, min(float(w), self.cx + dx))
        self.cy = max(0.0, min(float(h), self.cy + dy))

    def rotation_handle_pos(self):
        px = math.cos(self.angle)
        py = math.sin(self.angle)
        return self.cx + px * self.ROT_OFFSET, self.cy + py * self.ROT_OFFSET

    def angle_deg(self):
        return math.degrees(self.angle) % 360.0


# =========================================================
# Worker
# =========================================================

class WorkerSignals(QObject):
    progress = Signal(int)
    finished = Signal(str)
    error = Signal(str)


class BatchWorker(QRunnable):
    def __init__(self, app_ref):
        super().__init__()
        self.app_ref = app_ref
        self.signals = WorkerSignals()

    def run(self):
        try:
            total = len(self.app_ref.items)
            processed = 0
            for item in self.app_ref.items:
                if self.app_ref.stop_requested:
                    self.signals.finished.emit("Abgebrochen.")
                    return

                if not item.crop_enabled and not item.split_enabled:
                    processed += 1
                    self.signals.progress.emit(int(processed / total * 100))
                    continue

                self.app_ref.process_item(item)
                processed += 1
                self.signals.progress.emit(int(processed / total * 100))

            self.signals.finished.emit("Alle Einträge wurden verarbeitet.")
        except Exception as e:
            self.signals.error.emit(str(e))


class PersistentCheckMenu(QMenu):
    def mouseReleaseEvent(self, event):
        action = self.activeAction()
        if action and action.isEnabled() and action.isCheckable():
            action.trigger()
            event.accept()
            return
        super().mouseReleaseEvent(event)


# =========================================================
# Editor Canvas
# =========================================================

class EditorCanvas(QWidget):
    changed = Signal()
    filesDropped = Signal(list)

    def __init__(self):
        super().__init__()
        self.setMouseTracking(True)
        self.setAcceptDrops(True)
        self.setMinimumSize(600, 450)
        self.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Expanding)

        self.base_image: Optional[Image.Image] = None
        self.view_image: Optional[Image.Image] = None
        self.view_pixmap: Optional[QPixmap] = None

        self.zoom = 1.0
        self.fit_scale = 1.0

        self.show_crop = False
        self.show_separator = False

        self.crop_rect: Optional[QRectF] = None
        self.separator: Optional[Separator] = None

        self.drag_mode = None
        self.drag_start = QPointF()
        self.rect_before = None
        self.sep_offset = QPointF()

        self.rotation_mode = False
        self.show_grid = False

        # final gespeicherter Winkel
        self.rotation_angle = 0.0

        # nur für flüssige Live-Vorschau während des Draggens
        self.preview_rotation_angle = 0.0
        self.is_preview_rotating = False

        self.rotation_start_angle = 0.0
        self.rotation_start_mouse_angle = 0.0

    # -------------------------
    # Drag & Drop
    # -------------------------

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                paths.append(url.toLocalFile())
        self.filesDropped.emit(paths)
        event.acceptProposedAction()

    # -------------------------
    # Image management
    # -------------------------

    def set_image(self, img: Optional[Image.Image], reset_zoom: bool = True):
        self.base_image = img
        if reset_zoom:
            self.zoom = 1.0
        self._update_view_image()
        if self.view_image and self.show_crop and self.crop_rect is None:
            self.create_default_crop()
        self._ensure_separator_inside()
        self.update()
        self.changed.emit()

    def _update_view_image(self):
        if self.base_image is None:
            self.view_image = None
            self.view_pixmap = None
            return

        cw = max(10, self.width())
        ch = max(10, self.height())
        iw, ih = self.base_image.size
        self.fit_scale = min(cw / iw, ch / ih)
        scale = self.fit_scale * self.zoom
        nw = max(1, int(iw * scale))
        nh = max(1, int(ih * scale))
        self.view_image = self.base_image.resize((nw, nh), Image.LANCZOS)
        self.view_pixmap = pil_to_qpixmap(self.view_image)

        if self.crop_rect is not None:
            self.crop_rect = self.crop_rect.intersected(QRectF(0, 0, nw, nh))

    def create_default_crop(self):
        if not self.view_image:
            return
        w, h = self.view_image.size
        m = 0.05
        self.crop_rect = QRectF(w * m, h * m, w * (1 - 2 * m), h * (1 - 2 * m))
        self.changed.emit()

    def _ensure_separator_inside(self):
        if self.view_image is None or self.separator is None:
            return
        w, h = self.view_image.size
        self.separator.cx = max(0.0, min(float(w), self.separator.cx))
        self.separator.cy = max(0.0, min(float(h), self.separator.cy))

    def get_crop_orig(self) -> Optional[Tuple[int, int, int, int]]:
        if self.crop_rect is None or self.base_image is None or self.view_image is None:
            return None

        bw, bh = self.base_image.size
        vw, vh = self.view_image.size
        sx = bw / vw
        sy = bh / vh

        x1 = max(0, min(self.crop_rect.left(), vw - 2))
        y1 = max(0, min(self.crop_rect.top(), vh - 2))
        x2 = max(x1 + 2, min(self.crop_rect.right(), vw))
        y2 = max(y1 + 2, min(self.crop_rect.bottom(), vh))

        return (
            int(round(x1 * sx)),
            int(round(y1 * sy)),
            int(round(x2 * sx)),
            int(round(y2 * sy)),
        )

    def set_crop_from_orig(self, crop_orig: Optional[Tuple[int, int, int, int]]):
        if crop_orig is None or self.base_image is None or self.view_image is None:
            self.crop_rect = None
            self.update()
            return

        bw, bh = self.base_image.size
        vw, vh = self.view_image.size
        sx = vw / bw
        sy = vh / bh
        x1, y1, x2, y2 = crop_orig
        self.crop_rect = QRectF(x1 * sx, y1 * sy, (x2 - x1) * sx, (y2 - y1) * sy)
        self.update()

    # -------------------------
    # Border helpers for separator endpoints
    # -------------------------

    def _project_to_border(self, x: float, y: float) -> Tuple[float, float]:
        if self.view_image is None:
            return x, y

        w, h = self.view_image.size
        candidates = [
            (0.0, max(0.0, min(float(h), y))),
            (float(w), max(0.0, min(float(h), y))),
            (max(0.0, min(float(w), x)), 0.0),
            (max(0.0, min(float(w), x)), float(h)),
        ]

        best = candidates[0]
        best_d = (x - best[0]) ** 2 + (y - best[1]) ** 2
        for cx, cy in candidates[1:]:
            d = (x - cx) ** 2 + (y - cy) ** 2
            if d < best_d:
                best = (cx, cy)
                best_d = d
        return best

    def _mouse_angle_from_center(self, p: QPointF) -> float:
        if self.view_image is None:
            return 0.0
        w, h = self.view_image.size
        cx = w / 2.0
        cy = h / 2.0
        return math.degrees(math.atan2(p.y() - cy, p.x() - cx))

    # -------------------------
    # Painting
    # -------------------------

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor("#171a1f"))

        if self.view_pixmap is None:
            painter.setPen(QColor("#888"))
            painter.drawText(self.rect(), Qt.AlignCenter, "Bilder oder PDFs hier hineinziehen oder laden")
            return

        angle = self.preview_rotation_angle if self.is_preview_rotating else 0.0

        if abs(angle) > 0.01:
            painter.save()

            w = self.view_pixmap.width()
            h = self.view_pixmap.height()
            cx = w / 2.0
            cy = h / 2.0

            painter.translate(cx, cy)
            painter.rotate(angle)
            painter.translate(-cx, -cy)
            painter.drawPixmap(0, 0, self.view_pixmap)

            painter.restore()
        else:
            painter.drawPixmap(0, 0, self.view_pixmap)

        if self.show_grid:
            self._paint_grid(painter)

        if self.show_crop and self.crop_rect is not None:
            self._paint_crop(painter)

        if self.show_separator and self.separator is not None:
            self._paint_separator(painter)

    def _paint_crop(self, painter: QPainter):
        rect = self.crop_rect
        if rect is None:
            return

        painter.setPen(QPen(QColor("#ff4d4d"), 2))
        painter.setBrush(Qt.NoBrush)
        painter.drawRect(rect)

        handle_size = 10
        painter.setPen(QPen(QColor("black"), 1))

        corners = [
            rect.topLeft(), rect.topRight(),
            rect.bottomRight(), rect.bottomLeft()
        ]
        mids = [
            QPointF(rect.center().x(), rect.top()),
            QPointF(rect.right(), rect.center().y()),
            QPointF(rect.center().x(), rect.bottom()),
            QPointF(rect.left(), rect.center().y()),
        ]

        painter.setBrush(QColor("#ff4d4d"))
        for p in corners:
            painter.drawRect(QRectF(p.x() - handle_size / 2, p.y() - handle_size / 2, handle_size, handle_size))

        painter.setBrush(QColor("#ffb347"))
        for p in mids:
            painter.drawRect(QRectF(p.x() - handle_size / 2, p.y() - handle_size / 2, handle_size, handle_size))

    def _paint_separator(self, painter: QPainter):
        if self.view_image is None or self.separator is None:
            return

        w, h = self.view_image.size
        pts = self.separator.clipped_endpoints(w, h)
        if pts is None:
            return

        x1, y1, x2, y2 = pts
        painter.setPen(QPen(QColor("#58d68d"), 3))
        painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))

        painter.setPen(QPen(QColor("black"), 1))
        painter.setBrush(QColor("#ffc107"))

        p_top = self.separator.top_handle(w, h)
        p_bottom = self.separator.bottom_handle(w, h)

        for hx, hy in (p_top, p_bottom):
            painter.drawEllipse(QPointF(hx, hy), self.separator.HANDLE_R, self.separator.HANDLE_R)

        rx, ry = self.separator.rotation_handle_pos()
        painter.setBrush(QColor("#ffffff"))
        painter.setPen(QPen(QColor("#555"), 1))
        painter.drawEllipse(QPointF(rx, ry), self.separator.ROT_R, self.separator.ROT_R)
        painter.setPen(QColor("#222"))
        painter.drawText(QRectF(rx - 12, ry - 12, 24, 24), Qt.AlignCenter, "↻")

        btn_rect = QRectF(rx - 24, ry + 18, 48, 22)
        painter.setBrush(QColor("#ffffff"))
        painter.setPen(QPen(QColor("#111"), 1))
        painter.drawRoundedRect(btn_rect, 4, 4)
        painter.drawText(btn_rect, Qt.AlignCenter, "+90°")

        reset_rect = QRectF(rx - 24, ry + 46, 48, 22)
        painter.drawRoundedRect(reset_rect, 4, 4)
        painter.drawText(reset_rect, Qt.AlignCenter, "Reset")

        angle_box = QRectF(rx + 18, ry - 14, 68, 28)
        painter.setBrush(QColor("#ffffff"))
        painter.drawRoundedRect(angle_box, 4, 4)
        painter.setPen(QColor("#111"))
        painter.drawText(angle_box, Qt.AlignCenter, f"{self.separator.angle_deg():.1f}°")

    def _paint_grid(self, painter: QPainter):
        canvas_w = self.width()
        canvas_h = self.height()

        if canvas_w <= 1 or canvas_h <= 1:
            return

        painter.save()
        painter.setRenderHint(QPainter.Antialiasing, False)

        pen = QPen(QColor(0, 0, 0, 110), 1)
        pen.setCosmetic(True)
        painter.setPen(pen)

        cells = 14  # kleineres Raster als 5x5

        for i in range(1, cells):
            x = int(round(canvas_w * i / cells))
            painter.drawLine(x, 0, x, canvas_h)

        for i in range(1, cells):
            y = int(round(canvas_h * i / cells))
            painter.drawLine(0, y, canvas_w, y)

        painter.restore()

    # -------------------------
    # Hit tests
    # -------------------------

    def _point_in_crop(self, p: QPointF) -> bool:
        return self.crop_rect is not None and self.crop_rect.contains(p)

    def _crop_edge_at(self, p: QPointF):
        if self.crop_rect is None:
            return None
        r = self.crop_rect
        s = 8
        edges = []
        if abs(p.x() - r.left()) <= s and r.top() - s <= p.y() <= r.bottom() + s:
            edges.append("left")
        if abs(p.x() - r.right()) <= s and r.top() - s <= p.y() <= r.bottom() + s:
            edges.append("right")
        if abs(p.y() - r.top()) <= s and r.left() - s <= p.x() <= r.right() + s:
            edges.append("top")
        if abs(p.y() - r.bottom()) <= s and r.left() - s <= p.x() <= r.right() + s:
            edges.append("bottom")
        return "-".join(edges) if edges else None

    def _separator_hit(self, p: QPointF):
        if self.separator is None or self.view_image is None:
            return None

        w, h = self.view_image.size
        rx, ry = self.separator.rotation_handle_pos()
        if (p.x() - rx) ** 2 + (p.y() - ry) ** 2 <= (self.separator.ROT_R + 5) ** 2:
            return "rotate"

        btn_rect = QRectF(rx - 24, ry + 18, 48, 22)
        if btn_rect.contains(p):
            return "rotate90"

        reset_rect = QRectF(rx - 24, ry + 46, 48, 22)
        if reset_rect.contains(p):
            return "reset"

        tx, ty = self.separator.top_handle(w, h)
        bx, by = self.separator.bottom_handle(w, h)

        if (p.x() - tx) ** 2 + (p.y() - ty) ** 2 <= (self.separator.HANDLE_R + 4) ** 2:
            return "top"

        if (p.x() - bx) ** 2 + (p.y() - by) ** 2 <= (self.separator.HANDLE_R + 4) ** 2:
            return "bottom"

        if self.separator.distance_to_line(p.x(), p.y(), w, h) < 8:
            return "line"

        return None

    # -------------------------
    # Mouse
    # -------------------------

    def mousePressEvent(self, event):
        if self.view_image is None:
            return

        p = event.position()

        if self.rotation_mode:
            self.drag_mode = "img_rotate"
            self.rotation_start_angle = self.rotation_angle
            self.rotation_start_mouse_angle = self._mouse_angle_from_center(p)
            self.preview_rotation_angle = 0.0
            self.is_preview_rotating = True
            self.setCursor(Qt.ClosedHandCursor)
            return

        if self.show_separator and self.separator is not None:
            hit = self._separator_hit(p)
            if hit is not None:
                if hit == "rotate90":
                    self.separator.angle = (self.separator.angle + math.pi / 2) % (2 * math.pi)
                    self.update()
                    self.changed.emit()
                    return

                if hit == "reset":
                    self.separator.angle = 0.0
                    self.update()
                    self.changed.emit()
                    return

                if hit == "top":
                    self.drag_mode = "sep_top"
                elif hit == "bottom":
                    self.drag_mode = "sep_bottom"
                elif hit == "line":
                    self.drag_mode = "sep_line"
                    self.sep_offset = QPointF(self.separator.cx - p.x(), self.separator.cy - p.y())
                elif hit == "rotate":
                    self.drag_mode = "sep_rotate"

                self.drag_start = p
                self.update()
                return

        if self.show_crop:
            edge = self._crop_edge_at(p)
            if self.crop_rect is not None and edge:
                self.drag_mode = f"crop_resize:{edge}"
                self.drag_start = p
                self.rect_before = QRectF(self.crop_rect)
                return

            if self._point_in_crop(p):
                self.drag_mode = "crop_move"
                self.drag_start = p
                self.rect_before = QRectF(self.crop_rect)
                return

            self.drag_mode = "crop_new"
            self.drag_start = p
            self.crop_rect = QRectF(p, p)
            self.update()
            self.changed.emit()

    def mouseMoveEvent(self, event):
        p = event.position()

        if self.drag_mode == "img_rotate":
            current_mouse_angle = self._mouse_angle_from_center(p)
            delta = current_mouse_angle - self.rotation_start_mouse_angle
            new_angle = self.rotation_start_angle + delta

            if event.modifiers() & Qt.ControlModifier:
                step = 1.0
                new_angle = round(new_angle / step) * step

            self.preview_rotation_angle = new_angle - self.rotation_angle
            self.update()
            return

        if self.drag_mode == "sep_top" and self.separator and self.view_image is not None:
            w, h = self.view_image.size
            fixed = self.separator.bottom_handle(w, h)
            dragged = self._project_to_border(p.x(), p.y())
            if abs(dragged[0] - fixed[0]) > 1e-6 or abs(dragged[1] - fixed[1]) > 1e-6:
                self.separator.set_from_points(dragged, fixed)
                self.update()
                self.changed.emit()
            return

        if self.drag_mode == "sep_bottom" and self.separator and self.view_image is not None:
            w, h = self.view_image.size
            fixed = self.separator.top_handle(w, h)
            dragged = self._project_to_border(p.x(), p.y())
            if abs(dragged[0] - fixed[0]) > 1e-6 or abs(dragged[1] - fixed[1]) > 1e-6:
                self.separator.set_from_points(fixed, dragged)
                self.update()
                self.changed.emit()
            return

        if self.drag_mode == "sep_line" and self.separator and self.view_image is not None:
            w, h = self.view_image.size
            new_x = p.x() + self.sep_offset.x()
            new_y = p.y() + self.sep_offset.y()
            dx = new_x - self.separator.cx
            dy = new_y - self.separator.cy
            self.separator.move_by(dx, dy, w, h)
            self.update()
            self.changed.emit()
            return

        if self.drag_mode == "sep_rotate" and self.separator:
            dx = p.x() - self.separator.cx
            dy = p.y() - self.separator.cy
            if abs(dx) > 1e-6 or abs(dy) > 1e-6:
                raw = math.atan2(dy, dx) - math.pi / 2
                if event.modifiers() & Qt.ControlModifier:
                    step = math.radians(5)
                    raw = round(raw / step) * step
                self.separator.angle = raw
                self.update()
                self.changed.emit()
            return

        if self.drag_mode == "crop_move" and self.crop_rect and self.rect_before:
            delta = p - self.drag_start
            r = QRectF(self.rect_before)
            r.translate(delta)
            self.crop_rect = self._clamp_rect(r)
            self.update()
            self.changed.emit()
            return

        if self.drag_mode and self.drag_mode.startswith("crop_resize:") and self.rect_before:
            edge = self.drag_mode.split(":", 1)[1]
            r = QRectF(self.rect_before)
            if "left" in edge:
                r.setLeft(min(p.x(), r.right() - 5))
            if "right" in edge:
                r.setRight(max(p.x(), r.left() + 5))
            if "top" in edge:
                r.setTop(min(p.y(), r.bottom() - 5))
            if "bottom" in edge:
                r.setBottom(max(p.y(), r.top() + 5))
            self.crop_rect = self._clamp_rect(r)
            self.update()
            self.changed.emit()
            return

        if self.drag_mode == "crop_new":
            x1 = min(self.drag_start.x(), p.x())
            y1 = min(self.drag_start.y(), p.y())
            x2 = max(self.drag_start.x(), p.x())
            y2 = max(self.drag_start.y(), p.y())
            self.crop_rect = self._clamp_rect(QRectF(x1, y1, x2 - x1, y2 - y1))
            self.update()
            self.changed.emit()
            return

        self._update_cursor(p)

    def mouseReleaseEvent(self, event):
        if self.drag_mode == "img_rotate":
            self.rotation_angle = self.rotation_angle + self.preview_rotation_angle
            self.preview_rotation_angle = 0.0
            self.is_preview_rotating = False

        self.drag_mode = None
        self.rect_before = None
        self.sep_offset = QPointF()
        self._update_cursor(event.position())
        self.update()
        self.changed.emit()

    def wheelEvent(self, event):
        if self.base_image is None:
            return

        factor = 1.1 if event.angleDelta().y() > 0 else 0.9
        old_crop = self.get_crop_orig()
        self.zoom = max(0.2, min(6.0, self.zoom * factor))
        self._update_view_image()
        self.set_crop_from_orig(old_crop)
        self._ensure_separator_inside()
        self.update()
        self.changed.emit()

    def resizeEvent(self, event):
        old_crop = self.get_crop_orig()
        self._update_view_image()
        self.set_crop_from_orig(old_crop)
        self._ensure_separator_inside()
        self.update()
        super().resizeEvent(event)

    def _clamp_rect(self, rect: QRectF) -> QRectF:
        if self.view_image is None:
            return rect
        w, h = self.view_image.size
        x1 = max(0, min(rect.left(), w - 5))
        y1 = max(0, min(rect.top(), h - 5))
        x2 = max(x1 + 5, min(rect.right(), w))
        y2 = max(y1 + 5, min(rect.bottom(), h))
        return QRectF(x1, y1, x2 - x1, y2 - y1)

    def _update_cursor(self, p: QPointF):
        if self.rotation_mode:
            self.setCursor(Qt.OpenHandCursor)
            return

        if self.show_separator and self.separator is not None:
            hit = self._separator_hit(p)
            if hit in ("rotate", "top", "bottom", "line"):
                self.setCursor(Qt.SizeAllCursor)
                return
            if hit in ("rotate90", "reset"):
                self.setCursor(Qt.PointingHandCursor)
                return

        if self.show_crop:
            edge = self._crop_edge_at(p)
            if edge:
                if edge in ("left", "right"):
                    self.setCursor(Qt.SizeHorCursor)
                elif edge in ("top", "bottom"):
                    self.setCursor(Qt.SizeVerCursor)
                else:
                    self.setCursor(Qt.SizeFDiagCursor)
                return
            if self._point_in_crop(p):
                self.setCursor(Qt.SizeAllCursor)
                return

        self.setCursor(Qt.CrossCursor)


# =========================================================
# Main Window
# =========================================================

class MainWindow(QMainWindow):
    def __init__(self):
        super().__init__()
        self.setWindowTitle("BuchCutter")
        self.resize(1500, 900)
        self.setAcceptDrops(True)

        icon_path = resource_path("icon.ico")
        if os.path.exists(icon_path):
            self.setWindowIcon(QIcon(icon_path))

        self.items: List[ItemState] = []
        self.current_index = -1
        self.output_folder = ""
        self.stop_requested = False
        self.threadpool = QThreadPool.globalInstance()

        self.current_crop_orig = None

        self.save_formats = {
            "JPEG": True,
            "PNG": False,
            "TIFF": False,
            "BMP": False,
            "PDF": False,
        }

        self.is_dark_mode = False

        self._build_ui()
        self.apply_light_theme()
        self.theme_switch.setChecked(False)
        self.theme_icon.setText("☀️")

    def get_effective_crop_area(self, item: ItemState, img: Image.Image):
        if item.crop_enabled:
            if self.current_crop_orig is None:
                raise RuntimeError("Crop ist für diesen Eintrag aktiviert, aber es wurde kein Crop-Bereich gesetzt.")
            return self.current_crop_orig
        return (0, 0, img.size[0], img.size[1])

    def get_preview_image_for_item(self, item: ItemState) -> Image.Image:
        img = load_image_or_pdf_page(item.source_path, item.page_index, pdf_dpi=PREVIEW_PDF_DPI)
        img = self.apply_item_image_options(img, item)
        return img

    # -------------------------
    # Drag & Drop
    # -------------------------

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            return
        event.ignore()

    def dropEvent(self, event):
        if not event.mimeData().hasUrls():
            event.ignore()
            return
        paths = []
        for url in event.mimeData().urls():
            if url.isLocalFile():
                paths.append(url.toLocalFile())
        self.add_paths(paths)
        event.acceptProposedAction()

    # -------------------------
    # UI
    # -------------------------

    def _build_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        root = QVBoxLayout(central)
        root.setContentsMargins(8, 8, 8, 8)
        root.setSpacing(8)

        toolbar = QFrame()
        toolbar.setObjectName("ToolbarFrame")
        tl = QHBoxLayout(toolbar)
        tl.setContentsMargins(10, 10, 10, 10)
        tl.setSpacing(8)

        btn_load = QPushButton("Bilder / PDFs laden")
        btn_load.clicked.connect(self.load_files)
        tl.addWidget(btn_load)

        btn_out = QPushButton("Speicherordner")
        btn_out.clicked.connect(self.select_output_folder)
        tl.addWidget(btn_out)

        self.btn_rotate_mode = QPushButton("Rotation: AUS")
        self.btn_rotate_mode.setCheckable(True)
        self.btn_rotate_mode.toggled.connect(self.toggle_rotation_mode)
        tl.addWidget(self.btn_rotate_mode)

        self.btn_grid = QPushButton("Raster")
        self.btn_grid.setCheckable(True)
        self.btn_grid.toggled.connect(self.toggle_grid)
        tl.addWidget(self.btn_grid)

        btn_rot_left = QPushButton("↺ 90°")
        btn_rot_left.clicked.connect(lambda: self.rotate_current_by(-90))
        tl.addWidget(btn_rot_left)

        btn_rot_right = QPushButton("↻ 90°")
        btn_rot_right.clicked.connect(lambda: self.rotate_current_by(90))
        tl.addWidget(btn_rot_right)

        btn_rot_reset = QPushButton("Rotation Reset")
        btn_rot_reset.clicked.connect(self.reset_current_rotation)
        tl.addWidget(btn_rot_reset)

        self.chk_show_crop = QCheckBox("Crop-Bereich")
        self.chk_show_crop.stateChanged.connect(self.toggle_crop)
        tl.addWidget(self.chk_show_crop)

        self.chk_show_sep = QCheckBox("Trennbalken")
        self.chk_show_sep.stateChanged.connect(self.toggle_separator)
        tl.addWidget(self.chk_show_sep)

        self.chk_smart_sep = QCheckBox("Smart Split")
        self.chk_smart_sep.setChecked(True)
        tl.addWidget(self.chk_smart_sep)

        self.format_button = QToolButton()
        self.format_button.setText("Formate")
        self.format_button.setPopupMode(QToolButton.InstantPopup)
        fmt_menu = PersistentCheckMenu(self)
        for name in self.save_formats.keys():
            act = QAction(name, self, checkable=True)
            act.setChecked(self.save_formats[name])
            act.toggled.connect(lambda checked, n=name: self._set_format(n, checked))
            fmt_menu.addAction(act)
        self.format_button.setMenu(fmt_menu)
        tl.addWidget(self.format_button)

        theme_wrap = QFrame()
        theme_wrap.setObjectName("ThemeWrap")
        theme_layout = QHBoxLayout(theme_wrap)
        theme_layout.setContentsMargins(6, 0, 6, 0)
        theme_layout.setSpacing(6)

        self.theme_icon = QLabel("☀️")
        self.theme_icon.setObjectName("ThemeIcon")
        self.theme_icon.setAlignment(Qt.AlignCenter)
        self.theme_icon.setToolTip("Hell-/Dunkelmodus umschalten")
        theme_layout.addWidget(self.theme_icon)

        self.theme_switch = QCheckBox()
        self.theme_switch.setProperty("themeSwitch", True)
        self.theme_switch.setChecked(False)
        self.theme_switch.setCursor(Qt.PointingHandCursor)
        self.theme_switch.setToolTip("Hell-/Dunkelmodus umschalten")
        self.theme_switch.stateChanged.connect(self.toggle_theme)
        theme_layout.addWidget(self.theme_switch)

        tl.addWidget(theme_wrap)

        tl.addStretch(1)

        btn_current = QPushButton("Einmal bearbeiten")
        btn_current.clicked.connect(self.process_current)
        tl.addWidget(btn_current)

        btn_all = QPushButton("Alle bearbeiten")
        btn_all.clicked.connect(self.process_all)
        tl.addWidget(btn_all)

        btn_stop = QPushButton("Stopp")
        btn_stop.clicked.connect(self.stop_processing)
        tl.addWidget(btn_stop)

        root.addWidget(toolbar)

        splitter = QSplitter()
        root.addWidget(splitter, 1)
        splitter.setChildrenCollapsible(False)

        left = QFrame()
        left.setObjectName("SidePanel")
        left.setMinimumWidth(340)
        left.setMaximumWidth(520)
        ll = QVBoxLayout(left)
        ll.setContentsMargins(8, 8, 8, 8)
        ll.setSpacing(8)

        ll.addWidget(QLabel("Wartebereich"))

        mode_row = QHBoxLayout()

        btn_rgb = QPushButton("Farbig (RGB)")
        btn_rgb.clicked.connect(self.apply_rgb_to_selected)
        mode_row.addWidget(btn_rgb)

        btn_gray = QPushButton("Grau (S/W)")
        btn_gray.clicked.connect(self.apply_gray_to_selected)
        mode_row.addWidget(btn_gray)

        btn_contrast = QPushButton("Kontrast")
        btn_contrast.clicked.connect(self.toggle_contrast_for_selected)
        mode_row.addWidget(btn_contrast)

        ll.addLayout(mode_row)

        self.table = QTableWidget(0, 5)
        self.table.setHorizontalHeaderLabels(["#", "☑", "Dateien", "Crop", "Trennen"])
        self.table.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.table.horizontalHeader().setSectionResizeMode(3, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().setSectionResizeMode(4, QHeaderView.ResizeToContents)
        self.table.horizontalHeader().sectionClicked.connect(self.on_header_clicked)
        self.table.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.table.setSelectionMode(QAbstractItemView.ExtendedSelection)
        self.table.itemSelectionChanged.connect(self.on_table_selection_changed)
        self.table.itemChanged.connect(self.on_table_item_changed)

        self.table.setContextMenuPolicy(Qt.CustomContextMenu)
        self.table.customContextMenuRequested.connect(self.show_table_context_menu)
        self.table.verticalHeader().setVisible(False)
        self.table.setColumnWidth(0, 34)
        self.table.setColumnWidth(1, 34)
        self.table.setColumnWidth(3, 60)
        self.table.setColumnWidth(4, 75)
        ll.addWidget(self.table)

        btn_clear = QPushButton("Alle löschen")
        btn_clear.clicked.connect(self.clear_all)
        ll.addWidget(btn_clear)

        splitter.addWidget(left)

        self.canvas = EditorCanvas()
        self.canvas.changed.connect(self.sync_from_canvas)
        self.canvas.filesDropped.connect(self.add_paths)
        splitter.addWidget(self.canvas)

        splitter.setCollapsible(0, False)
        splitter.setCollapsible(1, False)

        splitter.setSizes([380, 1120])

        self.progress = QProgressBar()
        self.progress.setValue(0)
        root.addWidget(self.progress)

        self.act_paste = QAction("Einfügen", self)
        self.act_paste.setShortcut(QKeySequence("Ctrl+V"))
        self.act_paste.setShortcutContext(Qt.ApplicationShortcut)
        self.act_paste.triggered.connect(self.paste_from_clipboard)
        self.addAction(self.act_paste)

        self.act_delete = QAction("Löschen", self)
        self.act_delete.setShortcut(QKeySequence(Qt.Key_Delete))
        self.act_delete.setShortcutContext(Qt.ApplicationShortcut)
        self.act_delete.triggered.connect(self.delete_marked_or_selected_items)
        self.addAction(self.act_delete)

    def apply_dark_theme(self):
        app = QApplication.instance()
        app.setStyle("Fusion")

        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(23, 26, 31))
        palette.setColor(QPalette.WindowText, Qt.white)
        palette.setColor(QPalette.Base, QColor(30, 34, 40))
        palette.setColor(QPalette.AlternateBase, QColor(40, 44, 52))
        palette.setColor(QPalette.ToolTipBase, Qt.white)
        palette.setColor(QPalette.ToolTipText, Qt.white)
        palette.setColor(QPalette.Text, Qt.white)
        palette.setColor(QPalette.Button, QColor(38, 43, 51))
        palette.setColor(QPalette.ButtonText, Qt.white)
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(88, 166, 255))
        palette.setColor(QPalette.HighlightedText, Qt.black)
        app.setPalette(palette)

        self.setStyleSheet("""
        QMainWindow, QWidget {
            font-size: 13px;
            color: white;
        }
        #ToolbarFrame, #SidePanel {
            background: #1e232a;
            border: 1px solid #2f3945;
            border-radius: 12px;
        }
        QPushButton, QToolButton {
            background: #2d3642;
            color: white;
            border: 1px solid #44505f;
            border-radius: 10px;
            padding: 8px 12px;
        }
        QPushButton:hover, QToolButton:hover {
            background: #364152;
        }
        QPushButton:pressed, QToolButton:pressed {
            background: #26303c;
        }
        QTableWidget {
            background: #20252c;
            color: white;
            gridline-color: #33404f;
            border-radius: 8px;
        }
        QHeaderView::section {
            background: #2c3440;
            color: white;
            padding: 6px;
            border: none;
            border-bottom: 1px solid #43505f;
        }
        QProgressBar {
            border: 1px solid #43505f;
            border-radius: 8px;
            background: #20252c;
            text-align: center;
            min-height: 22px;
            color: white;
        }
        QProgressBar::chunk {
            background: #42b883;
            border-radius: 7px;
        }

        QCheckBox {
            spacing: 8px;
            color: white;
        }

        QCheckBox::indicator {
            width: 46px;
            height: 24px;
            border-radius: 12px;
            background: #4a5563;
            border: 1px solid #5c6878;
        }

        QCheckBox::indicator:checked {
            background: #42b883;
            border: 1px solid #42b883;
        }

        /* normal checkboxes im UI nicht als Switch darstellen */
        QTableWidget QCheckBox::indicator,
        #ToolbarFrame QCheckBox[text="Crop-Bereich"]::indicator,
        #ToolbarFrame QCheckBox[text="Trennbalken"]::indicator,
        #ToolbarFrame QCheckBox[text="Smart Split"]::indicator {
            width: 14px;
            height: 14px;
            border-radius: 3px;
            background: #20252c;
            border: 1px solid #5c6878;
        }

        QTableWidget QCheckBox::indicator:checked,
        #ToolbarFrame QCheckBox[text="Crop-Bereich"]::indicator:checked,
        #ToolbarFrame QCheckBox[text="Trennbalken"]::indicator:checked,
        #ToolbarFrame QCheckBox[text="Smart Split"]::indicator:checked {
            background: #42b883;
            border: 1px solid #42b883;
        }

        #ThemeWrap {
            background: transparent;
            border: none;
        }

        #ThemeIcon {
            font-size: 18px;
            color: #ffd54a;
            padding-right: 2px;
        }

        /* nur der Theme-Switch als Slider */
        QCheckBox[themeSwitch="true"]::indicator {
            width: 46px;
            height: 24px;
            border-radius: 12px;
            background: #4a5563;
            border: 1px solid #5c6878;
        }

        QCheckBox[themeSwitch="true"]::indicator:checked {
            background: #42b883;
            border: 1px solid #42b883;
        }
        """)

        if hasattr(self, "theme_switch"):
            self.theme_switch.setProperty("themeSwitch", True)
            self.theme_switch.style().unpolish(self.theme_switch)
            self.theme_switch.style().polish(self.theme_switch)

    def apply_light_theme(self):
        app = QApplication.instance()
        app.setStyle("Fusion")

        palette = QPalette()
        palette.setColor(QPalette.Window, QColor(245, 247, 250))
        palette.setColor(QPalette.WindowText, QColor(30, 35, 40))
        palette.setColor(QPalette.Base, QColor(255, 255, 255))
        palette.setColor(QPalette.AlternateBase, QColor(245, 247, 250))
        palette.setColor(QPalette.ToolTipBase, QColor(255, 255, 255))
        palette.setColor(QPalette.ToolTipText, QColor(30, 35, 40))
        palette.setColor(QPalette.Text, QColor(30, 35, 40))
        palette.setColor(QPalette.Button, QColor(235, 239, 244))
        palette.setColor(QPalette.ButtonText, QColor(30, 35, 40))
        palette.setColor(QPalette.BrightText, Qt.red)
        palette.setColor(QPalette.Highlight, QColor(88, 166, 255))
        palette.setColor(QPalette.HighlightedText, Qt.white)
        app.setPalette(palette)

        self.setStyleSheet("""
        QMainWindow, QWidget {
            font-size: 13px;
            color: #1e2328;
        }
        #ToolbarFrame, #SidePanel {
            background: #f4f6f8;
            border: 1px solid #d5dbe3;
            border-radius: 12px;
        }
        QPushButton, QToolButton {
            background: #e8edf3;
            color: #1e2328;
            border: 1px solid #c8d0da;
            border-radius: 10px;
            padding: 8px 12px;
        }
        QPushButton:hover, QToolButton:hover {
            background: #dde5ee;
        }
        QPushButton:pressed, QToolButton:pressed {
            background: #d3dce6;
        }
        QTableWidget {
            background: white;
            color: #1e2328;
            gridline-color: #d8dee6;
            border-radius: 8px;
        }
        QHeaderView::section {
            background: #e9eef4;
            color: #1e2328;
            padding: 6px;
            border: none;
            border-bottom: 1px solid #cfd7e0;
        }
        QProgressBar {
            border: 1px solid #cfd7e0;
            border-radius: 8px;
            background: #eef2f6;
            text-align: center;
            min-height: 22px;
            color: #1e2328;
        }
        QProgressBar::chunk {
            background: #42b883;
            border-radius: 7px;
        }

        QCheckBox {
            spacing: 8px;
            color: #1e2328;
        }

        QTableWidget QCheckBox::indicator,
        #ToolbarFrame QCheckBox[text="Crop-Bereich"]::indicator,
        #ToolbarFrame QCheckBox[text="Trennbalken"]::indicator,
        #ToolbarFrame QCheckBox[text="Smart Split"]::indicator {
            width: 14px;
            height: 14px;
            border-radius: 3px;
            background: white;
            border: 1px solid #aab4c0;
        }

        QTableWidget QCheckBox::indicator:checked,
        #ToolbarFrame QCheckBox[text="Crop-Bereich"]::indicator:checked,
        #ToolbarFrame QCheckBox[text="Trennbalken"]::indicator:checked,
        #ToolbarFrame QCheckBox[text="Smart Split"]::indicator:checked {
            background: #42b883;
            border: 1px solid #42b883;
        }

        #ThemeWrap {
            background: transparent;
            border: none;
        }

        #ThemeIcon {
            font-size: 18px;
            color: #d89b00;
            padding-right: 2px;
        }

        QCheckBox[themeSwitch="true"]::indicator {
            width: 46px;
            height: 24px;
            border-radius: 12px;
            background: #cfd7e0;
            border: 1px solid #b8c2cd;
        }

        QCheckBox[themeSwitch="true"]::indicator:checked {
            background: #42b883;
            border: 1px solid #42b883;
        }
        """)

        if hasattr(self, "theme_switch"):
            self.theme_switch.setProperty("themeSwitch", True)
            self.theme_switch.style().unpolish(self.theme_switch)
            self.theme_switch.style().polish(self.theme_switch)

    def toggle_theme(self):
        self.is_dark_mode = self.theme_switch.isChecked()
        if self.is_dark_mode:
            self.apply_dark_theme()
            self.theme_switch.setToolTip("Dunkelmodus aktiv")
            self.theme_icon.setText("🌙")
        else:
            self.apply_light_theme()
            self.theme_switch.setToolTip("Hellmodus aktiv")
            self.theme_icon.setText("☀️")

    # -------------------------
    # File loading
    # -------------------------

    def load_files(self):
        paths, _ = QFileDialog.getOpenFileNames(
            self,
            "Dateien laden",
            "",
            "Bilder/PDF (*.png *.jpg *.jpeg *.bmp *.tif *.tiff *.pdf)"
        )
        if not paths:
            return
        self.add_paths(paths)

    def add_paths(self, paths: List[str]):
        if not paths:
            return

        added = 0
        existing_keys = {item.unique_key() for item in self.items}

        for path in normalize_dropped_paths(paths):
            if not os.path.isfile(path):
                continue

            ext = os.path.splitext(path)[1].lower()
            if ext not in SUPPORTED_EXTENSIONS:
                continue

            if ext == ".pdf":
                try:
                    doc = fitz.open(path)
                    try:
                        for page_index in range(doc.page_count):
                            item = ItemState(
                                source_path=path,
                                page_index=page_index,
                                display_name=f"{os.path.basename(path)} [Seite {page_index + 1}]",
                            )
                            if item.unique_key() not in existing_keys:
                                self.items.append(item)
                                existing_keys.add(item.unique_key())
                                added += 1
                    finally:
                        doc.close()
                except Exception as e:
                    QMessageBox.warning(self, "PDF-Fehler", f"PDF konnte nicht geladen werden:\n{path}\n\n{e}")
            else:
                item = ItemState(
                    source_path=path,
                    page_index=None,
                    display_name=os.path.basename(path),
                )
                if item.unique_key() not in existing_keys:
                    self.items.append(item)
                    existing_keys.add(item.unique_key())
                    added += 1

        self.refresh_table()

        if added and self.current_index < 0:
            self.current_index = 0
            self.select_current_row()
            self.load_current_item()

    def select_output_folder(self):
        folder = QFileDialog.getExistingDirectory(self, "Zielordner auswählen")
        if folder:
            self.output_folder = folder

    def clear_all(self):
        self.items.clear()
        self.current_index = -1
        self.current_crop_orig = None
        self.refresh_table()
        self.canvas.set_image(None)
        self.canvas.crop_rect = None
        self.canvas.separator = None
        self.canvas.update()
        self.progress.setValue(0)

    def paste_from_clipboard(self):
        clipboard = QApplication.clipboard()
        mime = clipboard.mimeData()

        paths = []

        # Fall 1: kopierte Dateien
        if mime and mime.hasUrls():
            for url in mime.urls():
                if url.isLocalFile():
                    path = url.toLocalFile()
                    ext = os.path.splitext(path)[1].lower()
                    if ext in SUPPORTED_EXTENSIONS and os.path.isfile(path):
                        paths.append(path)

        # Fall 2: Text mit Dateipfaden
        elif mime and mime.hasText():
            raw_text = mime.text().strip()
            if raw_text:
                for line in raw_text.splitlines():
                    path = line.strip().strip('"').strip("'")
                    ext = os.path.splitext(path)[1].lower()
                    if os.path.isfile(path) and ext in SUPPORTED_EXTENSIONS:
                        paths.append(path)

        # Fall 3: Bild direkt aus Zwischenablage
        elif mime and mime.hasImage():
            image = clipboard.image()
            if not image.isNull():
                temp_dir = os.path.join(tempfile.gettempdir(), "BuchCutter_clipboard")
                os.makedirs(temp_dir, exist_ok=True)

                filename = f"clipboard_{time.strftime('%Y%m%d_%H%M%S')}_{os.getpid()}.png"
                out_path = os.path.join(temp_dir, filename)

                if image.save(out_path, "PNG"):
                    paths.append(out_path)

        if not paths:
            QMessageBox.information(
                self,
                "Zwischenablage",
                "Es wurden keine unterstützten Bilder oder PDFs in der Zwischenablage gefunden."
            )
            return

        self.add_paths(paths)

    def get_rows_for_deletion(self):
        # 1) Vorrang: Checkboxen in Spalte "☑"
        checked_rows = [i for i, item in enumerate(self.items) if item.selected]
        if checked_rows:
            return sorted(set(checked_rows))

        # 2) Sonst: markierte Tabellenzeilen
        selection_model = self.table.selectionModel()
        if selection_model is not None:
            selected_rows = [idx.row() for idx in selection_model.selectedRows()]
            if selected_rows:
                return sorted(set(selected_rows))

        # 3) Fallback: aktueller Eintrag
        if 0 <= self.current_index < len(self.items):
            return [self.current_index]

        return []

    def delete_marked_or_selected_items(self):
        rows = self.get_rows_for_deletion()
        if not rows:
            return

        new_index = min(rows[0], len(self.items) - len(rows) - 1)
        if new_index < 0:
            new_index = -1

        for row in sorted(rows, reverse=True):
            if 0 <= row < len(self.items):
                del self.items[row]

        if not self.items:
            self.current_index = -1
            self.current_crop_orig = None
            self.refresh_table()
            self.canvas.set_image(None)
            self.canvas.crop_rect = None
            self.canvas.separator = None
            self.canvas.update()
            return

        self.current_index = min(max(new_index, 0), len(self.items) - 1)
        self.current_crop_orig = None
        self.refresh_table()
        self.select_current_row()
        self.load_current_item()

    def show_table_context_menu(self, pos):
        item = self.table.itemAt(pos)

        # Nur wenn aktuell nichts selektiert ist, selektieren wir die Rechtsklick-Zeile
        if item is not None and not self.table.selectionModel().selectedRows():
            self.table.selectRow(item.row())

        menu = QMenu(self)
        delete_action = menu.addAction("Löschen")
        chosen = menu.exec(self.table.viewport().mapToGlobal(pos))

        if chosen == delete_action:
            self.delete_marked_or_selected_items()
    # -------------------------
    # Table
    # -------------------------
    def on_header_clicked(self, column: int):
        if column not in (1, 3, 4):
            return

        if column == 1:
            all_checked = all(item.selected for item in self.items) if self.items else False
            new_value = not all_checked
            self.table.blockSignals(True)
            for row, item in enumerate(self.items):
                item.selected = new_value
                cell = self.table.item(row, 1)
                if cell:
                    cell.setCheckState(Qt.Checked if new_value else Qt.Unchecked)
            self.table.blockSignals(False)

        elif column == 3:
            all_checked = all(item.crop_enabled for item in self.items) if self.items else False
            new_value = not all_checked
            self.table.blockSignals(True)
            for row, item in enumerate(self.items):
                item.crop_enabled = new_value
                cell = self.table.item(row, 3)
                if cell:
                    cell.setCheckState(Qt.Checked if new_value else Qt.Unchecked)
            self.table.blockSignals(False)

        elif column == 4:
            all_checked = all(item.split_enabled for item in self.items) if self.items else False
            new_value = not all_checked
            self.table.blockSignals(True)
            for row, item in enumerate(self.items):
                item.split_enabled = new_value
                cell = self.table.item(row, 4)
                if cell:
                    cell.setCheckState(Qt.Checked if new_value else Qt.Unchecked)
            self.table.blockSignals(False)

    def refresh_table(self):
        self.table.blockSignals(True)
        self.table.setRowCount(0)

        for row, item in enumerate(self.items):
            self.table.insertRow(row)

            nr_item = QTableWidgetItem(str(row + 1))
            nr_item.setFlags(Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            nr_item.setTextAlignment(Qt.AlignCenter)
            self.table.setItem(row, 0, nr_item)

            sel_item = QTableWidgetItem()
            sel_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            sel_item.setCheckState(Qt.Checked if item.selected else Qt.Unchecked)
            self.table.setItem(row, 1, sel_item)

            name_item = QTableWidgetItem(item.display_name)
            name_item.setFlags(name_item.flags() & ~Qt.ItemIsEditable)
            tooltip = f"Modus: {'Farbig' if item.color_mode == 'RGB' else 'Grau'}"
            if item.contrast_enabled:
                tooltip += " | Kontrast: an"
            else:
                tooltip += " | Kontrast: aus"
            tooltip += f" | Rotation: {item.rotation_angle:.1f}°"
            name_item.setToolTip(tooltip)
            self.table.setItem(row, 2, name_item)

            crop_item = QTableWidgetItem()
            crop_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            crop_item.setCheckState(Qt.Checked if item.crop_enabled else Qt.Unchecked)
            self.table.setItem(row, 3, crop_item)

            split_item = QTableWidgetItem()
            split_item.setFlags(Qt.ItemIsUserCheckable | Qt.ItemIsEnabled | Qt.ItemIsSelectable)
            split_item.setCheckState(Qt.Checked if item.split_enabled else Qt.Unchecked)
            self.table.setItem(row, 4, split_item)

        self.table.blockSignals(False)

    def set_all_checks(self, mode: str, value: bool):
        state = Qt.Checked if value else Qt.Unchecked

        self.table.blockSignals(True)
        for row, item in enumerate(self.items):
            if mode == "select":
                item.selected = value
                table_item = self.table.item(row, 1)
                if table_item:
                    table_item.setCheckState(state)
            elif mode == "crop":
                item.crop_enabled = value
                table_item = self.table.item(row, 3)
                if table_item:
                    table_item.setCheckState(state)
            elif mode == "split":
                item.split_enabled = value
                table_item = self.table.item(row, 4)
                if table_item:
                    table_item.setCheckState(state)
        self.table.blockSignals(False)

    def apply_rgb_to_selected(self):
        changed = False
        for item in self.items:
            if item.selected:
                item.color_mode = "RGB"
                changed = True
        if changed:
            self.refresh_table()
            self.load_current_item()

    def apply_gray_to_selected(self):
        changed = False
        for item in self.items:
            if item.selected:
                item.color_mode = "GRAY"
                changed = True
        if changed:
            self.refresh_table()
            self.load_current_item()

    def toggle_contrast_for_selected(self):
        selected_items = [item for item in self.items if item.selected]
        if not selected_items:
            return

        all_on = all(item.contrast_enabled for item in selected_items)
        new_value = not all_on

        for item in selected_items:
            item.contrast_enabled = new_value

        self.refresh_table()
        self.load_current_item()

    def on_table_item_changed(self, table_item):
        row = table_item.row()
        if not (0 <= row < len(self.items)):
            return

        if table_item.column() == 1:
            self.items[row].selected = table_item.checkState() == Qt.Checked
        elif table_item.column() == 3:
            self.items[row].crop_enabled = table_item.checkState() == Qt.Checked
        elif table_item.column() == 4:
            self.items[row].split_enabled = table_item.checkState() == Qt.Checked

    def on_table_selection_changed(self):
        rows = self.table.selectionModel().selectedRows()
        if not rows:
            return

        # Für die Vorschau immer die erste selektierte Zeile verwenden
        row_indices = sorted(idx.row() for idx in rows)
        self.current_index = row_indices[0]
        self.load_current_item()

    def select_current_row(self):
        if 0 <= self.current_index < len(self.items):
            self.table.selectRow(self.current_index)

    # -------------------------
    # Navigation
    # -------------------------

    def load_current_item(self):
        if not (0 <= self.current_index < len(self.items)):
            return

        item = self.items[self.current_index]
        try:
            self.canvas.rotation_angle = item.rotation_angle
            img = self.get_preview_image_for_item(item)

            reset_zoom = self.canvas.base_image is None
            self.canvas.set_image(img, reset_zoom=reset_zoom)

            self.setWindowTitle(f"BuchCutter - {item.display_name} | Rotation: {item.rotation_angle:.1f}°")
        except Exception as e:
            QMessageBox.critical(self, "Ladefehler", f"Datei konnte nicht geladen werden:\n\n{e}")

    # -------------------------
    # Canvas sync / toggles
    # -------------------------

    def sync_from_canvas(self):
        if not (0 <= self.current_index < len(self.items)):
            self.current_crop_orig = self.canvas.get_crop_orig()
            return

        item = self.items[self.current_index]
        new_angle = self.canvas.rotation_angle

        if abs(item.rotation_angle - new_angle) > 0.01:
            item.rotation_angle = new_angle
            self.current_crop_orig = None
            self.canvas.crop_rect = None
            self.load_current_item()
            return

        self.current_crop_orig = self.canvas.get_crop_orig()

    def toggle_crop(self):
        self.canvas.show_crop = self.chk_show_crop.isChecked()
        if self.canvas.show_crop and self.canvas.crop_rect is None and self.canvas.view_image is not None:
            self.canvas.create_default_crop()
        self.canvas.update()

    def toggle_separator(self):
        self.canvas.show_separator = self.chk_show_sep.isChecked()
        if self.canvas.show_separator and self.canvas.separator is None and self.canvas.view_image is not None:
            w, h = self.canvas.view_image.size
            self.canvas.separator = Separator(cx=w / 2, cy=h / 2, angle=0.0)
        if not self.canvas.show_separator:
            self.canvas.separator = None
        self.canvas.update()

    def toggle_rotation_mode(self, checked):
        self.canvas.rotation_mode = checked
        self.btn_rotate_mode.setText("Rotation: AN" if checked else "Rotation: AUS")
        self.canvas.update()

    def toggle_grid(self, checked):
        self.canvas.show_grid = checked
        self.canvas.update()

    def rotate_current_by(self, delta_degrees):
        if not (0 <= self.current_index < len(self.items)):
            return

        item = self.items[self.current_index]
        item.rotation_angle = (item.rotation_angle + delta_degrees) % 360.0
        self.canvas.rotation_angle = item.rotation_angle
        self.current_crop_orig = None
        self.load_current_item()

    def reset_current_rotation(self):
        if not (0 <= self.current_index < len(self.items)):
            return

        item = self.items[self.current_index]
        item.rotation_angle = 0.0
        self.canvas.rotation_angle = 0.0
        self.current_crop_orig = None
        self.load_current_item()

    def _set_format(self, name: str, checked: bool):
        self.save_formats[name] = checked

    # -------------------------
    # Processing
    # -------------------------

    def process_current(self):
        if not (0 <= self.current_index < len(self.items)):
            return

        item = self.items[self.current_index]

        if item.crop_enabled and self.current_crop_orig is None:
            QMessageBox.warning(self, "Fehler", "Bitte zuerst einen Crop-Bereich festlegen.")
            return

        if not item.crop_enabled and not item.split_enabled:
            QMessageBox.warning(self, "Fehler", "Für diesen Eintrag ist weder Crop noch Trennen aktiviert.")
            return

        try:
            saved = self.process_item(item)
            QMessageBox.information(self, "Fertig", f"{len(saved)} Datei(en) gespeichert.")
        except Exception as e:
            QMessageBox.critical(self, "Fehler", str(e))

    def process_all(self):
        if not self.items:
            return

        has_work = any(item.crop_enabled or item.split_enabled for item in self.items)
        if not has_work:
            QMessageBox.warning(self, "Fehler", "Es ist nichts zur Verarbeitung aktiviert.")
            return

        if any(item.crop_enabled for item in self.items) and self.current_crop_orig is None:
            QMessageBox.warning(
                self,
                "Fehler",
                "Mindestens ein Eintrag hat Crop aktiviert, aber es ist kein Crop-Bereich gesetzt."
            )
            return

        self.stop_requested = False
        self.progress.setValue(0)

        worker = BatchWorker(self)
        worker.signals.progress.connect(self.progress.setValue)
        worker.signals.error.connect(lambda msg: QMessageBox.critical(self, "Fehler", msg))
        worker.signals.finished.connect(self._batch_finished)
        self.threadpool.start(worker)

    def _batch_finished(self, message: str):
        self.progress.setValue(0)
        QMessageBox.information(self, "Status", message)

    def stop_processing(self):
        self.stop_requested = True

    def process_item(self, item: ItemState) -> List[str]:
        img = load_image_or_pdf_page(item.source_path, item.page_index, pdf_dpi=EXPORT_PDF_DPI)
        img = self.apply_item_image_options(img, item)

        crop_area = self.get_effective_crop_area(item, img)
        line_segments = self.get_separator_lines_for_processing(img, crop_area)
        segments = self.compute_segments_for_crop(crop_area, line_segments) if item.split_enabled else []

        return self.save_outputs(item, img, crop_area, segments)

    def get_separator_lines_for_processing(self, img: Image.Image, crop_area):
        if not self.chk_show_sep.isChecked() or self.canvas.separator is None:
            return []

        if self.canvas.view_image is None:
            return []

        vw, vh = self.canvas.view_image.size
        bw, bh = img.size
        sx = bw / vw
        sy = bh / vh

        pts = self.canvas.separator.clipped_endpoints(vw, vh)
        if pts is None:
            return []

        x1d, y1d, x2d, y2d = pts
        line_orig = (x1d * sx, y1d * sy, x2d * sx, y2d * sy)

        if self.chk_smart_sep.isChecked():
            line_orig = self.smart_adjust_split_line(img, crop_area, line_orig)

        return [line_orig]

    # -------------------------
    # Geometry for saving
    # -------------------------

    def smart_adjust_split_line(self, img: Image.Image, crop_area, line_orig):
        ox1, oy1, ox2, oy2 = crop_area
        crop = img.crop((ox1, oy1, ox2, oy2)).convert("L")
        w, h = crop.size
        if w < 20 or h < 20:
            return line_orig

        x1, y1, x2, y2 = line_orig
        px = crop.load()

        def expected_x(global_y):
            if abs(y2 - y1) < 1e-6:
                return (x1 + x2) * 0.5
            t = (global_y - y1) / (y2 - y1)
            return x1 + t * (x2 - x1)

        band = max(2, min(6, w // 120))
        search_radius = max(20, min(120, w // 8))
        y_step = max(6, h // 80)

        samples = []

        for local_y in range(6, h - 6, y_step):
            global_y = oy1 + local_y
            ex = int(round(expected_x(global_y) - ox1))

            xmin = max(6, ex - search_radius)
            xmax = min(w - 7, ex + search_radius)
            if xmin >= xmax:
                continue

            best_x = None
            best_score = None

            for x in range(xmin, xmax + 1):
                center_vals = []
                left_vals = []
                right_vals = []

                for yy in range(local_y - 2, local_y + 3):
                    for xx in range(x - band, x + band + 1):
                        center_vals.append(px[xx, yy])

                    for xx in range(max(0, x - 14), max(0, x - 4)):
                        left_vals.append(px[xx, yy])

                    for xx in range(min(w - 1, x + 4), min(w, x + 15)):
                        right_vals.append(px[xx, yy])

                if not center_vals or not left_vals or not right_vals:
                    continue

                center_mean = sum(center_vals) / len(center_vals)
                left_mean = sum(left_vals) / len(left_vals)
                right_mean = sum(right_vals) / len(right_vals)

                contrast = ((left_mean + right_mean) * 0.5) - center_mean
                distance_penalty = abs(x - ex) * 0.15

                score = center_mean - contrast * 1.8 + distance_penalty
                if best_score is None or score < best_score:
                    best_score = score
                    best_x = x

            if best_x is not None:
                samples.append((local_y, best_x))

        if len(samples) < 2:
            return line_orig

        smoothed = []
        for i in range(len(samples)):
            xs = []
            for j in range(max(0, i - 2), min(len(samples), i + 3)):
                xs.append(samples[j][1])
            smoothed.append((samples[i][0], sum(xs) / len(xs)))

        n = len(smoothed)
        sum_y = sum(y for y, _ in smoothed)
        sum_x = sum(x for _, x in smoothed)
        sum_yy = sum(y * y for y, _ in smoothed)
        sum_yx = sum(y * x for y, x in smoothed)

        denom = n * sum_yy - sum_y * sum_y
        if abs(denom) < 1e-9:
            return line_orig

        m = (n * sum_yx - sum_y * sum_x) / denom
        b = (sum_x - m * sum_y) / n

        x_top_local = b
        x_bottom_local = m * (h - 1) + b

        x_top = max(0, min(img.size[0], ox1 + x_top_local))
        x_bottom = max(0, min(img.size[0], ox1 + x_bottom_local))

        return (x_top, oy1, x_bottom, oy2)

    def compute_segments_for_crop(self, crop_area, line_segments_orig):
        ox1, oy1, ox2, oy2 = crop_area
        rect_poly = [(ox1, oy1), (ox2, oy1), (ox2, oy2), (ox1, oy2)]

        if not line_segments_orig:
            return [rect_poly]

        entries = []
        for x1, y1, x2, y2 in line_segments_orig:
            vx = x2 - x1
            vy = y2 - y1
            nx = -vy
            ny = vx
            norm = math.hypot(nx, ny)
            if norm < 1e-12:
                continue

            nx /= norm
            ny /= norm
            c = -(nx * x1 + ny * y1)
            d = -c
            entries.append((d, nx, ny, c))

        entries.sort(key=lambda e: e[0])
        if not entries:
            return [rect_poly]

        segments = []
        N = len(entries)

        for i in range(N + 1):
            poly = rect_poly[:]

            if i == 0:
                a, b, c = entries[0][1], entries[0][2], entries[0][3]
                poly = clip_polygon_halfplane(poly, -a, -b, -c)
            elif i == N:
                a, b, c = entries[-1][1], entries[-1][2], entries[-1][3]
                poly = clip_polygon_halfplane(poly, a, b, c)
            else:
                a1, b1, c1 = entries[i - 1][1], entries[i - 1][2], entries[i - 1][3]
                a2, b2, c2 = entries[i][1], entries[i][2], entries[i][3]
                poly = clip_polygon_halfplane(poly, a1, b1, c1)
                poly = clip_polygon_halfplane(poly, -a2, -b2, -c2)

            if polygon_area(poly) > 1.0:
                segments.append(poly)

        return segments

    def apply_item_image_options(self, img: Image.Image, item: ItemState) -> Image.Image:
        out = img

        if item.color_mode == "GRAY":
            out = ImageOps.grayscale(out).convert("RGB")

        if item.contrast_enabled:
            if out.mode not in ("RGB", "RGBA"):
                out = out.convert("RGB")
            out = ImageOps.autocontrast(out, cutoff=1)
            out = ImageEnhance.Contrast(out).enhance(2.2)
            out = ImageEnhance.Sharpness(out).enhance(1.4)

        if abs(item.rotation_angle) > 0.01:
            if out.mode not in ("RGB", "RGBA"):
                out = out.convert("RGB")
            out = out.rotate(
                -item.rotation_angle,
                expand=True,
                resample=Image.BICUBIC,
                fillcolor="white"
            )

        return out

    # -------------------------
    # Saving
    # -------------------------

    def save_outputs(self, item: ItemState, img: Image.Image, crop_area, segments_polygons):
        if not item.crop_enabled and not item.split_enabled:
            return []

        formats = [name for name, enabled in self.save_formats.items() if enabled]
        if not formats:
            formats = ["JPEG"]

        ext_map = {
            "JPEG": "jpg",
            "PNG": "png",
            "TIFF": "tiff",
            "BMP": "bmp",
            "PDF": "pdf",
        }

        ox1, oy1, ox2, oy2 = crop_area
        crop = img.crop((ox1, oy1, ox2, oy2))

        root_outdir = self.output_folder or os.path.dirname(item.source_path)
        os.makedirs(root_outdir, exist_ok=True)

        base_name = os.path.splitext(os.path.basename(item.source_path))[0]
        if item.page_index is not None:
            base_name += f"_seite_{item.page_index + 1}"

        saved_files = []
        separators_active = item.split_enabled and bool(segments_polygons)

        if item.crop_enabled and not separators_active:
            crop_dir = os.path.join(root_outdir, "Crop-Ordner")
            os.makedirs(crop_dir, exist_ok=True)

            n = 1
            while True:
                stem = f"{base_name}_crop_edit_{n}"
                exists_any = False
                for fmt in formats:
                    fmt_dir = os.path.join(crop_dir, fmt)
                    ext = ext_map[fmt]
                    if os.path.exists(os.path.join(fmt_dir, f"{stem}.{ext}")):
                        exists_any = True
                        break
                if not exists_any:
                    break
                n += 1

            for fmt in formats:
                ext = ext_map[fmt]
                fmt_dir = os.path.join(crop_dir, fmt)
                os.makedirs(fmt_dir, exist_ok=True)
                outpath = os.path.join(fmt_dir, f"{base_name}_crop_edit_{n}.{ext}")

                self._save_pil(crop, outpath, fmt)

                saved_files.append(outpath)

        if separators_active:
            split_dir = os.path.join(root_outdir, "Trenn-Ordner")
            os.makedirs(split_dir, exist_ok=True)

            current_index = self._next_global_split_index(split_dir)

            ordered_polys = sorted(
                segments_polygons,
                key=lambda poly: sum(x for x, _ in poly) / len(poly)
            )

            for poly in ordered_polys:
                if not poly or polygon_area(poly) < 1.0:
                    continue

                local = [(x - ox1, y - oy1) for (x, y) in poly]

                full_w, full_h = crop.size
                full_rgba = Image.new("RGBA", (full_w, full_h), (0, 0, 0, 0))
                mask = Image.new("L", (full_w, full_h), 0)
                draw = ImageDraw.Draw(mask)
                draw.polygon(local, fill=255)
                full_rgba.paste(crop.convert("RGBA"), (0, 0), mask)

                min_x = max(0, int(math.floor(min(x for x, _ in local))))
                min_y = max(0, int(math.floor(min(y for _, y in local))))
                max_x = min(full_w, int(math.ceil(max(x for x, _ in local))))
                max_y = min(full_h, int(math.ceil(max(y for _, y in local))))

                if max_x - min_x < 2 or max_y - min_y < 2:
                    continue

                segment_img = full_rgba.crop((min_x, min_y, max_x, max_y))

                for fmt in formats:
                    ext = ext_map[fmt]
                    fmt_dir = os.path.join(split_dir, fmt)
                    os.makedirs(fmt_dir, exist_ok=True)
                    outpath = os.path.join(fmt_dir, f"{base_name}_teil_{current_index}.{ext}")

                    self._save_pil(segment_img, outpath, fmt)

                    saved_files.append(outpath)

                current_index += 1

        return saved_files

    def _save_pil(self, img: Image.Image, outpath: str, fmt: str):
        dpi = (300, 300)

        if fmt == "PDF":
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "RGBA":
                bg.paste(img, (0, 0), img.split()[-1])
            else:
                bg.paste(img.convert("RGB"))
            bg.save(outpath, format="PDF", resolution=300.0)

        elif fmt == "PNG":
            if img.mode != "RGBA":
                img = img.convert("RGBA")
            img.save(outpath, format="PNG", dpi=dpi)

        elif fmt in ("JPEG", "BMP", "TIFF"):
            bg = Image.new("RGB", img.size, (255, 255, 255))
            if img.mode == "RGBA":
                bg.paste(img, (0, 0), img.split()[-1])
            else:
                bg.paste(img.convert("RGB"))

            if fmt == "JPEG":
                bg.save(outpath, format=fmt, quality=95, dpi=dpi, subsampling=0)
            elif fmt == "TIFF":
                bg.save(outpath, format=fmt, dpi=dpi)
            elif fmt == "BMP":
                bg.save(outpath, format=fmt)
            else:
                bg.save(outpath, format=fmt, dpi=dpi)

        else:
            img.save(outpath, dpi=dpi)

    def _next_global_split_index(self, folder: str) -> int:
        max_n = 0
        pattern = re.compile(r"_teil_(\d+)(?:\.[^.]+)?$", re.IGNORECASE)
        if not os.path.isdir(folder):
            return 1

        for sub in os.listdir(folder):
            sub_path = os.path.join(folder, sub)
            if os.path.isdir(sub_path):
                files = os.listdir(sub_path)
            else:
                files = [sub]

            for fn in files:
                m = pattern.search(fn)
                if m:
                    try:
                        max_n = max(max_n, int(m.group(1)))
                    except ValueError:
                        pass

        return max_n + 1


# =========================================================
# Main
# =========================================================

def main():
    if sys.platform.startswith("win"):
        try:
            ctypes.windll.shell32.SetCurrentProcessExplicitAppUserModelID("buchcutter.app")
        except Exception:
            pass

    QApplication.setHighDpiScaleFactorRoundingPolicy(
        Qt.HighDpiScaleFactorRoundingPolicy.Round
    )

    app = QApplication(sys.argv)

    icon_path = resource_path("icon.ico")
    if os.path.exists(icon_path):
        app.setWindowIcon(QIcon(icon_path))

    win = MainWindow()

    if os.path.exists(icon_path):
        win.setWindowIcon(QIcon(icon_path))

    win.showMaximized()
    sys.exit(app.exec())


if __name__ == "__main__":
    main()