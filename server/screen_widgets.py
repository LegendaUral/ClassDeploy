"""
Виджеты экранов учеников:
  ScreenTile     - плитка с миниатюрой экрана
  ScreenGrid     - сетка всех плиток
  FullScreenView - полноэкранный просмотр + управление мышью/клавиатурой
"""
from __future__ import annotations
import base64
import math
import time
from typing import Callable, Dict, Optional

from PyQt6.QtCore import Qt, QEvent, QPoint, pyqtSignal, QTimer, QRect
from PyQt6.QtGui import (
    QPixmap, QKeyEvent, QCursor, QColor, QPainter, QPen, QFont, QBrush,
)
from PyQt6.QtWidgets import (
    QWidget, QLabel, QVBoxLayout, QGridLayout,
    QScrollArea, QHBoxLayout, QPushButton, QDialog, QFrame,
    QSizePolicy,
)


# ════════════════════════════════════════════════════════════════
#  ScreenTile — одна плитка в сетке
# ════════════════════════════════════════════════════════════════

class ScreenTile(QFrame):
    double_clicked = pyqtSignal(str)

    def __init__(self, hostname: str):
        super().__init__()
        self.hostname  = hostname
        self.last_w    = 1920
        self.last_h    = 1080
        self._fps_cnt  = 0
        self._fps_ts   = time.monotonic()
        self._fps      = 0.0

        self.setFrameStyle(QFrame.Shape.Box)
        self.setStyleSheet("QFrame { background:#1a1a1a; border:1px solid #333; }")
        self.setMinimumSize(220, 165)
        self.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(2, 2, 2, 2)
        lay.setSpacing(2)

        # Заголовок
        self.lbl_name = QLabel(hostname)
        self.lbl_name.setStyleSheet(
            "color:#ddd; background:#2a2a2a; padding:2px 5px;"
            "font-size:10px; font-weight:bold;"
        )
        self.lbl_name.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.lbl_name)

        # Картинка
        self.image = QLabel("нет сигнала")
        self.image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image.setStyleSheet("background:#000; color:#444; font-size:10px;")
        self.image.setScaledContents(False)
        self.image.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        lay.addWidget(self.image, stretch=1)

        # Статус-строка
        self.lbl_status = QLabel("ожидание")
        self.lbl_status.setStyleSheet("color:#555; font-size:9px; padding:1px 4px;")
        self.lbl_status.setAlignment(Qt.AlignmentFlag.AlignCenter)
        lay.addWidget(self.lbl_status)

    def set_frame(self, b64_jpeg: str, w: int, h: int):
        self.last_w, self.last_h = w, h
        # FPS счётчик
        self._fps_cnt += 1
        now = time.monotonic()
        elapsed = now - self._fps_ts
        if elapsed >= 2.0:
            self._fps    = self._fps_cnt / elapsed
            self._fps_cnt = 0
            self._fps_ts  = now
            self.lbl_status.setText(f"{w}×{h}  {self._fps:.1f} fps")

        try:
            data = base64.b64decode(b64_jpeg)
            pix  = QPixmap()
            if pix.loadFromData(data, "JPEG"):
                scaled = pix.scaled(
                    self.image.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.FastTransformation,
                )
                self.image.setPixmap(scaled)
        except Exception:
            pass

    def clear_frame(self):
        self.image.setPixmap(QPixmap())
        self.image.setText("нет сигнала")
        self.lbl_status.setText("нет сигнала")
        self._fps_cnt = 0

    def set_online(self, online: bool):
        color = "#1a7f37" if online else "#555555"
        bg = "rgba(26,127,55,48)" if online else "rgba(85,85,85,48)"
        self.lbl_name.setStyleSheet(
            f"color:#ddd; background:{bg}; border-bottom:2px solid {color};"
            "padding:2px 5px; font-size:10px; font-weight:bold;"
        )

    def mouseDoubleClickEvent(self, e):
        self.double_clicked.emit(self.hostname)

    def resizeEvent(self, e):
        # При изменении размера плитки — перемасштабируем последний кадр
        super().resizeEvent(e)
        pix = self.image.pixmap()
        if pix and not pix.isNull():
            self.image.setPixmap(pix.scaled(
                self.image.size(),
                Qt.AspectRatioMode.KeepAspectRatio,
                Qt.TransformationMode.FastTransformation,
            ))


# ════════════════════════════════════════════════════════════════
#  ScreenGrid — сетка плиток
# ════════════════════════════════════════════════════════════════

class ScreenGrid(QWidget):
    tile_double_clicked = pyqtSignal(str)

    def __init__(self):
        super().__init__()
        self.tiles: Dict[str, ScreenTile] = {}

        root = QVBoxLayout(self)
        root.setContentsMargins(0, 0, 0, 0)

        self.scroll = QScrollArea()
        self.scroll.setWidgetResizable(True)
        self._container = QWidget()
        self.grid = QGridLayout(self._container)
        self.grid.setSpacing(4)
        self.scroll.setWidget(self._container)
        root.addWidget(self.scroll)

    def set_hosts(self, hosts: list[str]):
        current = set(self.tiles.keys())
        target  = set(hosts)

        for h in current - target:
            t = self.tiles.pop(h)
            t.setParent(None)
            t.deleteLater()

        for h in target - current:
            t = ScreenTile(h)
            t.double_clicked.connect(self.tile_double_clicked.emit)
            self.tiles[h] = t

        self._relayout()

    def _relayout(self):
        while self.grid.count():
            item = self.grid.takeAt(0)
            if item.widget():
                item.widget().hide()
                item.widget().setParent(None)

        hosts = sorted(self.tiles.keys())
        if not hosts:
            return
        cols = max(1, math.ceil(math.sqrt(len(hosts))))
        for i, h in enumerate(hosts):
            tile = self.tiles[h]
            self.grid.addWidget(tile, i // cols, i % cols)
            tile.show()

    def update_frame(self, hostname: str, b64: str, w: int, h: int):
        tile = self.tiles.get(hostname)
        if tile:
            tile.set_frame(b64, w, h)

    def clear_host(self, hostname: str):
        tile = self.tiles.get(hostname)
        if tile:
            tile.clear_frame()

    def set_online(self, hostname: str, online: bool):
        tile = self.tiles.get(hostname)
        if tile:
            tile.set_online(online)


# ════════════════════════════════════════════════════════════════
#  FullScreenView — полноэкранный просмотр + управление
# ════════════════════════════════════════════════════════════════

class _ScreenCanvas(QLabel):
    """QLabel с кастомным paint — рисует перекрестие и оверлей при управлении."""

    def __init__(self):
        super().__init__()
        self.control_active = False
        self._cursor_x      = -1
        self._cursor_y      = -1

    def update_cursor(self, x: int, y: int):
        self._cursor_x = x
        self._cursor_y = y
        self.update()

    def paintEvent(self, e):
        super().paintEvent(e)
        if not self.control_active:
            return
        p = QPainter(self)

        # Полупрозрачная рамка «управление активно»
        p.setPen(QPen(QColor(30, 200, 100, 180), 3))
        p.drawRect(1, 1, self.width() - 2, self.height() - 2)

        # Текст-статус
        p.fillRect(QRect(0, 0, self.width(), 26), QColor(0, 0, 0, 140))
        p.setPen(QColor(30, 220, 100))
        p.setFont(QFont("Segoe UI", 11, QFont.Weight.Bold))
        p.drawText(
            QRect(0, 0, self.width(), 26),
            Qt.AlignmentFlag.AlignCenter,
            "⌨🖱  УПРАВЛЕНИЕ АКТИВНО  -  Esc или Ctrl+Shift+Z для выхода",
        )

        # Перекрестие в позиции курсора
        if self._cursor_x >= 0 and self._cursor_y >= 0:
            cx, cy = self._cursor_x, self._cursor_y
            pen = QPen(QColor(30, 220, 100, 220), 1)
            p.setPen(pen)
            p.drawLine(cx - 12, cy, cx + 12, cy)
            p.drawLine(cx, cy - 12, cx, cy + 12)
            p.drawEllipse(cx - 5, cy - 5, 10, 10)

        p.end()


class FullScreenView(QDialog):
    """
    Полноэкранный просмотр ПК ученика.
    Управление в окне трансляции оставлено только для мыши.
    Клавиатура удалённо из трансляции не отправляется.
    """

    def __init__(
        self,
        hostname: str,
        send_input: Callable[[str, dict], None],
        parent=None,
    ):
        super().__init__(parent)
        self.hostname         = hostname
        self.send_input       = send_input
        self.last_w           = 1920
        self.last_h           = 1080
        self.control_enabled  = False
        self._pressed_mouse_buttons: set[str] = set()
        self._last_mouse_pos: Optional[tuple[int, int]] = None

        self.setWindowTitle(f"Экран: {hostname}")
        self.resize(1280, 820)

        lay = QVBoxLayout(self)
        lay.setContentsMargins(0, 0, 0, 0)
        lay.setSpacing(0)

        # ── Панель управления ──────────────────────────────────
        bar = QWidget()
        bar.setStyleSheet("background:#1e1e1e;")
        bar.setFixedHeight(40)
        bar_l = QHBoxLayout(bar)
        bar_l.setContentsMargins(8, 4, 8, 4)

        lbl_host = QLabel(f"<b style='color:#ddd'>{hostname}</b>")
        bar_l.addWidget(lbl_host)

        bar_l.addStretch()

        self.lbl_res = QLabel("")
        self.lbl_res.setStyleSheet("color:#666; font-size:10px;")
        bar_l.addWidget(self.lbl_res)

        self.btn_ctrl = QPushButton("🖱 Включить мышь")
        self.btn_ctrl.setCheckable(True)
        self.btn_ctrl.setFixedHeight(28)
        self.btn_ctrl.setStyleSheet("""
            QPushButton {
                background:#333; color:#ccc; border:1px solid #555;
                border-radius:4px; padding:0 10px; font-size:10px;
            }
            QPushButton:checked {
                background:#1a7f37; color:white; border:1px solid #2a9f47;
                font-weight:bold;
            }
        """)
        self.btn_ctrl.toggled.connect(self._on_ctrl_toggle)
        bar_l.addWidget(self.btn_ctrl)

        self.btn_release = QPushButton("⎋ Выйти из управления")
        self.btn_release.setFixedHeight(28)
        self.btn_release.setEnabled(False)
        self.btn_release.clicked.connect(lambda: self.btn_ctrl.setChecked(False))
        bar_l.addWidget(self.btn_release)

        self.btn_close = QPushButton("✕")
        self.btn_close.setFixedSize(28, 28)
        self.btn_close.setStyleSheet(
            "QPushButton { background:#c0392b; color:white; border:none; border-radius:4px; }"
            "QPushButton:hover { background:#e74c3c; }"
        )
        self.btn_close.clicked.connect(self.accept)
        bar_l.addWidget(self.btn_close)

        lay.addWidget(bar)

        # ── Экран ─────────────────────────────────────────────
        self.image = _ScreenCanvas()
        self.image.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.image.setStyleSheet("background:#0a0a0a;")
        self.image.setMouseTracking(True)
        self.image.setSizePolicy(QSizePolicy.Policy.Expanding, QSizePolicy.Policy.Expanding)
        lay.addWidget(self.image, stretch=1)

        self.installEventFilter(self)
        self.image.installEventFilter(self)
        self.setFocusPolicy(Qt.FocusPolicy.StrongFocus)
        self.image.setFocusPolicy(Qt.FocusPolicy.StrongFocus)

        # ── FPS-таймер ────────────────────────────────────────
        self._fps_cnt  = 0
        self._fps_ts   = time.monotonic()
        self._fps_timer = QTimer(self)
        self._fps_timer.timeout.connect(self._update_fps)
        self._fps_timer.start(1000)

    # ── Управление ────────────────────────────────────────────

    def _on_ctrl_toggle(self, on: bool):
        self.control_enabled    = on
        self.image.control_active = on

        if on:
            self.btn_ctrl.setText("🖱 Мышь активна")
            self.btn_release.setEnabled(True)
            self.setCursor(QCursor(Qt.CursorShape.BlankCursor))
            self.image.setCursor(QCursor(Qt.CursorShape.BlankCursor))
            self.activateWindow()
            self.raise_()
            try:
                self.image.grabMouse(QCursor(Qt.CursorShape.BlankCursor))
            except Exception:
                pass
            self.image.setFocus()
        else:
            self.btn_ctrl.setText("🖱 Включить мышь")
            self.btn_release.setEnabled(False)
            self.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
            self.image.setCursor(QCursor(Qt.CursorShape.ArrowCursor))
            self._release_all_mouse_buttons()
            try:
                self.image.releaseMouse()
            except Exception:
                pass

        self.image.update()

    # ── Кадры ─────────────────────────────────────────────────

    def set_frame(self, b64_jpeg: str, w: int, h: int):
        self.last_w, self.last_h = w, h
        self._fps_cnt += 1
        try:
            data = base64.b64decode(b64_jpeg)
            pix  = QPixmap()
            if pix.loadFromData(data, "JPEG"):
                scaled = pix.scaled(
                    self.image.size(),
                    Qt.AspectRatioMode.KeepAspectRatio,
                    Qt.TransformationMode.SmoothTransformation,
                )
                self.image.setPixmap(scaled)
        except Exception:
            pass

    def _update_fps(self):
        now     = time.monotonic()
        elapsed = now - self._fps_ts
        if elapsed > 0:
            fps = self._fps_cnt / elapsed
            self.lbl_res.setText(f"{self.last_w}×{self.last_h}  {fps:.1f} fps")
        self._fps_cnt = 0
        self._fps_ts  = now

    # ── Координаты ────────────────────────────────────────────

    def _to_screen(self, x: int, y: int) -> tuple[int, int]:
        pix = self.image.pixmap()
        if not pix or pix.isNull():
            return 0, 0
        pw, ph = pix.width(), pix.height()
        lw, lh = self.image.width(), self.image.height()
        ox = (lw - pw) // 2
        oy = (lh - ph) // 2
        ix = max(0, min(pw - 1, x - ox))
        iy = max(0, min(ph - 1, y - oy))
        if pw <= 0 or ph <= 0:
            return 0, 0
        return int(ix * self.last_w / pw), int(iy * self.last_h / ph)

    # ── EventFilter: мышь ─────────────────────────────────────

    def eventFilter(self, obj, ev):
        t = ev.type()

        if obj is self.image and self.control_enabled:
            if t == QEvent.Type.MouseMove:
                p = ev.position().toPoint()
                sx, sy = self._to_screen(p.x(), p.y())
                self.image.update_cursor(p.x(), p.y())
                self._inp({"kind": "mouse_move", "x": sx, "y": sy,
                           "screen_w": self.last_w, "screen_h": self.last_h})
                return True

            if t == QEvent.Type.MouseButtonPress:
                p = ev.position().toPoint()
                sx, sy = self._to_screen(p.x(), p.y())
                button = _qt_btn(ev.button())
                self._pressed_mouse_buttons.add(button)
                self.activateWindow()
                self.raise_()
                self.image.setFocus()
                self._inp({"kind": "mouse_button", "x": sx, "y": sy,
                           "button": button, "down": True,
                           "screen_w": self.last_w, "screen_h": self.last_h})
                return True

            if t == QEvent.Type.MouseButtonDblClick:
                p = ev.position().toPoint()
                sx, sy = self._to_screen(p.x(), p.y())
                self._inp({"kind": "mouse_click", "x": sx, "y": sy,
                           "button": _qt_btn(ev.button()), "double": True,
                           "screen_w": self.last_w, "screen_h": self.last_h})
                return True

            if t == QEvent.Type.MouseButtonRelease:
                p = ev.position().toPoint()
                sx, sy = self._to_screen(p.x(), p.y())
                button = _qt_btn(ev.button())
                self._pressed_mouse_buttons.discard(button)
                self._inp({"kind": "mouse_button", "x": sx, "y": sy,
                           "button": button, "down": False,
                           "screen_w": self.last_w, "screen_h": self.last_h})
                return True

            if t == QEvent.Type.Wheel:
                p = ev.position().toPoint()
                sx, sy = self._to_screen(p.x(), p.y())
                self._inp({"kind": "scroll", "x": sx, "y": sy,
                           "delta": ev.angleDelta().y(),
                           "screen_w": self.last_w, "screen_h": self.last_h})
                return True

        if self.control_enabled and t == QEvent.Type.KeyPress:
            key = ev.key()
            mods = ev.modifiers()
            if key == Qt.Key.Key_Escape or (
                key == Qt.Key.Key_Z
                and mods & Qt.KeyboardModifier.ControlModifier
                and mods & Qt.KeyboardModifier.ShiftModifier
            ):
                self.btn_ctrl.setChecked(False)
                ev.accept()
                return True

        return super().eventFilter(obj, ev)

    def keyPressEvent(self, ev):
        if self.control_enabled:
            key = ev.key()
            mods = ev.modifiers()
            if key == Qt.Key.Key_Escape or (
                key == Qt.Key.Key_Z
                and mods & Qt.KeyboardModifier.ControlModifier
                and mods & Qt.KeyboardModifier.ShiftModifier
            ):
                self.btn_ctrl.setChecked(False)
                ev.accept()
                return
        super().keyPressEvent(ev)

    def closeEvent(self, ev):
        self._release_all_mouse_buttons()
        try:
            self.image.releaseMouse()
        except Exception:
            pass
        self._fps_timer.stop()
        super().closeEvent(ev)

    def _release_all_mouse_buttons(self):
        for button in list(self._pressed_mouse_buttons):
            self._inp({
                "kind": "mouse_button",
                "x": 0,
                "y": 0,
                "button": button,
                "down": False,
                "screen_w": self.last_w,
                "screen_h": self.last_h,
            })
        self._pressed_mouse_buttons.clear()

    def _inp(self, payload: dict):
        try:
            self.send_input(self.hostname, payload)
        except Exception:
            pass


# ── Таблицы конвертации ──────────────────────────────────────────

def _qt_btn(btn) -> str:
    return {
        Qt.MouseButton.LeftButton:   "left",
        Qt.MouseButton.RightButton:  "right",
        Qt.MouseButton.MiddleButton: "middle",
    }.get(btn, "left")

