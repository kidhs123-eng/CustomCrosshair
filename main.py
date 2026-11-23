import ctypes
from ctypes import wintypes
import json
import math
import os
import sys
import traceback
import psutil
import win32api
import win32gui
import win32process

from PyQt6.QtCore import Qt, QTimer, QPointF, QRectF
from PyQt6.QtGui import (
    QAction, QBrush, QColor, QIcon, QImage, QPainter, QPen, QPixmap
)
from PyQt6.QtWidgets import (
    QApplication, QCheckBox, QColorDialog, QComboBox, QDialog, QDoubleSpinBox,
    QFileDialog, QFormLayout, QGroupBox, QHBoxLayout, QLabel, QMainWindow,
    QMenu, QMessageBox, QPushButton, QSlider, QSpinBox, QSystemTrayIcon,
    QVBoxLayout, QWidget
)

if getattr(sys, "frozen", False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))
CONFIG_FILE = os.path.join(BASE_DIR, "crosshair_config.json")
DRAWN_IMAGE_FILE = os.path.join(BASE_DIR, "custom_drawn_crosshair.png")
PLACEHOLDER_TEXT = "Выберите приложение..."

DEFAULT_CONFIG = {
    "config_version": 3,
    "target_exe": "",
    "enabled": True,
    "always_on": False,
    "style": "classic",
    "color": "#00FF66",
    "opacity": 255,
    "outline": True,
    "outline_thickness": 1,
    "length": 10,
    "thickness": 2,
    "gap": 4,
    "dot": True,
    "dot_size": 2,
    "circle_radius": 14,
    "offset_x": 0,
    "offset_y": 0,
    "image_path": "",
    "drawn_path": DRAWN_IMAGE_FILE,
    "image_scale": 150,
    "draw_scale": 1.0,
    "editor_scale": 1.0
}


# --- Структуры Windows CCD API для определения родного и текущего разрешения монитора ---
_QDC_ONLY_ACTIVE_PATHS = 0x00000002
_DISPLAYCONFIG_MODE_INFO_TYPE_SOURCE = 1
_DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_PREFERRED_MODE = 3


class _LUID(ctypes.Structure):
    _fields_ = [("LowPart", wintypes.DWORD), ("HighPart", wintypes.LONG)]


class _DISPLAYCONFIG_RATIONAL(ctypes.Structure):
    _fields_ = [("Numerator", ctypes.c_uint32), ("Denominator", ctypes.c_uint32)]


class _DISPLAYCONFIG_2DREGION(ctypes.Structure):
    _fields_ = [("cx", ctypes.c_uint32), ("cy", ctypes.c_uint32)]


class _DISPLAYCONFIG_VIDEO_SIGNAL_INFO(ctypes.Structure):
    _fields_ = [
        ("pixelRate", ctypes.c_uint64),
        ("hSyncFreq", _DISPLAYCONFIG_RATIONAL),
        ("vSyncFreq", _DISPLAYCONFIG_RATIONAL),
        ("activeSize", _DISPLAYCONFIG_2DREGION),
        ("totalSize", _DISPLAYCONFIG_2DREGION),
        ("videoStandard", ctypes.c_uint32),
        ("scanLineOrdering", ctypes.c_uint32),
    ]


class _DISPLAYCONFIG_TARGET_MODE(ctypes.Structure):
    _fields_ = [("targetVideoSignalInfo", _DISPLAYCONFIG_VIDEO_SIGNAL_INFO)]


class _DISPLAYCONFIG_SOURCE_MODE(ctypes.Structure):
    _fields_ = [
        ("width", ctypes.c_uint32),
        ("height", ctypes.c_uint32),
        ("pixelFormat", ctypes.c_uint32),
        ("position", wintypes.POINT),
    ]


class _DISPLAYCONFIG_MODE_INFO_UNION(ctypes.Union):
    _fields_ = [
        ("targetMode", _DISPLAYCONFIG_TARGET_MODE),
        ("sourceMode", _DISPLAYCONFIG_SOURCE_MODE),
    ]


class _DISPLAYCONFIG_MODE_INFO(ctypes.Structure):
    _fields_ = [
        ("infoType", ctypes.c_uint32),
        ("id", ctypes.c_uint32),
        ("adapterId", _LUID),
        ("u", _DISPLAYCONFIG_MODE_INFO_UNION),
    ]


class _DISPLAYCONFIG_PATH_SOURCE_INFO(ctypes.Structure):
    _fields_ = [
        ("adapterId", _LUID),
        ("id", ctypes.c_uint32),
        ("modeInfoIdx", ctypes.c_uint32),
        ("statusFlags", ctypes.c_uint32),
    ]


class _DISPLAYCONFIG_PATH_TARGET_INFO(ctypes.Structure):
    _fields_ = [
        ("adapterId", _LUID),
        ("id", ctypes.c_uint32),
        ("modeInfoIdx", ctypes.c_uint32),
        ("outputTechnology", ctypes.c_uint32),
        ("rotation", ctypes.c_uint32),
        ("scaling", ctypes.c_uint32),
        ("refreshRate", _DISPLAYCONFIG_RATIONAL),
        ("scanLineOrdering", ctypes.c_uint32),
        ("targetAvailable", wintypes.BOOL),
        ("statusFlags", ctypes.c_uint32),
    ]


class _DISPLAYCONFIG_PATH_INFO(ctypes.Structure):
    _fields_ = [
        ("sourceInfo", _DISPLAYCONFIG_PATH_SOURCE_INFO),
        ("targetInfo", _DISPLAYCONFIG_PATH_TARGET_INFO),
        ("flags", ctypes.c_uint32),
    ]


class _DISPLAYCONFIG_DEVICE_INFO_HEADER(ctypes.Structure):
    _fields_ = [
        ("type", ctypes.c_uint32),
        ("size", ctypes.c_uint32),
        ("adapterId", _LUID),
        ("id", ctypes.c_uint32),
    ]


class _DISPLAYCONFIG_TARGET_PREFERRED_MODE(ctypes.Structure):
    _fields_ = [
        ("header", _DISPLAYCONFIG_DEVICE_INFO_HEADER),
        ("width", ctypes.c_uint32),
        ("height", ctypes.c_uint32),
        ("targetMode", _DISPLAYCONFIG_TARGET_MODE),
    ]


_user32 = ctypes.WinDLL("user32")
_user32.GetDisplayConfigBufferSizes.argtypes = [
    ctypes.c_uint32,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.POINTER(ctypes.c_uint32),
]
_user32.GetDisplayConfigBufferSizes.restype = ctypes.c_long

_user32.QueryDisplayConfig.argtypes = [
    ctypes.c_uint32,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.c_void_p,
    ctypes.POINTER(ctypes.c_uint32),
    ctypes.c_void_p,
    ctypes.c_void_p,
]
_user32.QueryDisplayConfig.restype = ctypes.c_long

_user32.DisplayConfigGetDeviceInfo.argtypes = [ctypes.c_void_p]
_user32.DisplayConfigGetDeviceInfo.restype = ctypes.c_long


_SCALING_CENTERED = 2                # картинка по центру без растяжения
_SCALING_ASPECT = 4                  # растяжение с сохранением пропорций (чёрные полосы)


def get_primary_display_resolutions() -> tuple[int, int, int, int, int] | None:
    """Возвращает (cur_w, cur_h, native_w, native_h, scaling) основного монитора."""
    try:
        num_paths = ctypes.c_uint32()
        num_modes = ctypes.c_uint32()
        if _user32.GetDisplayConfigBufferSizes(
            _QDC_ONLY_ACTIVE_PATHS, ctypes.byref(num_paths), ctypes.byref(num_modes)
        ) != 0:
            return None

        paths = (_DISPLAYCONFIG_PATH_INFO * num_paths.value)()
        modes = (_DISPLAYCONFIG_MODE_INFO * num_modes.value)()
        if _user32.QueryDisplayConfig(
            _QDC_ONLY_ACTIVE_PATHS,
            ctypes.byref(num_paths),
            ctypes.byref(paths),
            ctypes.byref(num_modes),
            ctypes.byref(modes),
            None,
        ) != 0:
            return None

        for i in range(num_paths.value):
            path = paths[i]
            src_idx = path.sourceInfo.modeInfoIdx
            if src_idx >= num_modes.value:
                continue
            mode = modes[src_idx]
            if mode.infoType != _DISPLAYCONFIG_MODE_INFO_TYPE_SOURCE:
                continue
            src = mode.u.sourceMode
            if src.position.x != 0 or src.position.y != 0:
                continue

            cur_w, cur_h = int(src.width), int(src.height)
            pref = _DISPLAYCONFIG_TARGET_PREFERRED_MODE()
            pref.header.type = _DISPLAYCONFIG_DEVICE_INFO_GET_TARGET_PREFERRED_MODE
            pref.header.size = ctypes.sizeof(pref)
            pref.header.adapterId = path.targetInfo.adapterId
            pref.header.id = path.targetInfo.id

            if _user32.DisplayConfigGetDeviceInfo(ctypes.byref(pref)) == 0 and pref.width > 0 and pref.height > 0:
                native_w, native_h = int(pref.width), int(pref.height)
                if path.targetInfo.rotation in (2, 4):
                    native_w, native_h = native_h, native_w
                if cur_w > 0 and cur_h > 0:
                    return (cur_w, cur_h, native_w, native_h, int(path.targetInfo.scaling))
    except Exception:
        pass
    return None


def exception_hook(exctype, value, tb):
    err_msg = "".join(traceback.format_exception(exctype, value, tb))
    print(err_msg, file=sys.stderr)
    app = QApplication.instance()
    if app:
        QMessageBox.critical(None, "Критическая ошибка", err_msg)


sys.excepthook = exception_hook


def get_running_apps() -> list[tuple[str, str]]:
    """Возвращает список (Заголовок окна, имя.exe) для всех видимых окон."""
    apps = {}

    def enum_handler(hwnd, _):
        if win32gui.IsWindowVisible(hwnd):
            title = win32gui.GetWindowText(hwnd)
            if title:
                try:
                    _, pid = win32process.GetWindowThreadProcessId(hwnd)
                    exe = psutil.Process(pid).name()
                    if exe.lower() not in ("explorer.exe", "textinputhost.exe", "applicationframehost.exe"):
                        apps[exe] = title
                except Exception:
                    pass

    win32gui.EnumWindows(enum_handler, None)
    return sorted([(title, exe) for exe, title in apps.items()], key=lambda x: x[1].lower())


_fg_cache: list = [None, ""]   # [(hwnd, pid), exe]


def get_active_exe() -> str:
    """Возвращает имя .exe активного окна (с кэшем: процесс открывается только при смене окна)."""
    try:
        hwnd = win32gui.GetForegroundWindow()
        if not hwnd:
            return ""
        _, pid = win32process.GetWindowThreadProcessId(hwnd)
        key = (hwnd, pid)
        if _fg_cache[0] != key:
            _fg_cache[1] = psutil.Process(pid).name()
            _fg_cache[0] = key
        return _fg_cache[1]
    except Exception:
        _fg_cache[0] = None
        return ""


# --- Хук смены активного окна: реакция мгновенная, без постоянного опроса ---
_EVENT_SYSTEM_FOREGROUND = 0x0003
_WINEVENT_OUTOFCONTEXT = 0x0000
_WINEVENT_SKIPOWNPROCESS = 0x0002
_WINEVENTPROC = ctypes.WINFUNCTYPE(
    None, wintypes.HANDLE, wintypes.DWORD, wintypes.HWND,
    wintypes.LONG, wintypes.LONG, wintypes.DWORD, wintypes.DWORD
)
_user32_hook = ctypes.WinDLL("user32")
_user32_hook.SetWinEventHook.argtypes = [
    wintypes.DWORD, wintypes.DWORD, wintypes.HMODULE, _WINEVENTPROC,
    wintypes.DWORD, wintypes.DWORD, wintypes.DWORD
]
_user32_hook.SetWinEventHook.restype = wintypes.HANDLE
_user32_hook.UnhookWinEvent.argtypes = [wintypes.HANDLE]
_user32_hook.UnhookWinEvent.restype = wintypes.BOOL

_HWND_TOPMOST = -1
_SWP_FLAGS = 0x0001 | 0x0002 | 0x0010   # NOSIZE | NOMOVE | NOACTIVATE

SS = 8                      # подсетка: 1 клетка = 8x8 точек (точность 1/8 пикселя)
GRID = 64                   # размер холста в клетках
HIRES = GRID * SS


def hires_path(path: str) -> str:
    root, ext = os.path.splitext(path)
    return root + "_hires" + (ext or ".png")


def blank_hires() -> QImage:
    img = QImage(HIRES, HIRES, QImage.Format.Format_ARGB32_Premultiplied)
    img.fill(Qt.GlobalColor.transparent)
    return img


def load_hires(path: str) -> QImage:
    """Загружает рисунок с подсеткой. Если hires-файла нет, увеличивает обычный 64x64."""
    hp = hires_path(path)
    img = None
    if os.path.exists(hp) and os.path.exists(path) and os.path.getmtime(hp) >= os.path.getmtime(path) - 2:
        im = QImage(hp)
        if not im.isNull():
            img = im.scaled(HIRES, HIRES, Qt.AspectRatioMode.IgnoreAspectRatio,
                            Qt.TransformationMode.FastTransformation)
    if img is None and os.path.exists(path):
        im = QImage(path)
        if not im.isNull():
            img = im.scaled(HIRES, HIRES, Qt.AspectRatioMode.IgnoreAspectRatio,
                            Qt.TransformationMode.FastTransformation)
    if img is None:
        return blank_hires()
    return img.convertToFormat(QImage.Format.Format_ARGB32_Premultiplied)


def to_final(hires: QImage) -> QImage:
    """Усреднение подсетки до 64x64: половина пикселя -> частичная прозрачность клетки."""
    pm = hires.convertToFormat(QImage.Format.Format_ARGB32_Premultiplied)
    small = pm.scaled(GRID, GRID, Qt.AspectRatioMode.IgnoreAspectRatio,
                      Qt.TransformationMode.SmoothTransformation)
    return small.convertToFormat(QImage.Format.Format_ARGB32)

class PixelCanvas(QWidget):
    """Холст 64x64 клетки (внутри подсетка 8x): рисование, зум колесом, точное масштабирование."""

    VIEW = 448
    MIN_ZOOM = 7.0
    MAX_ZOOM = 40.0

    def __init__(self, image: QImage, color: QColor):
        super().__init__()
        self.grid_size = GRID
        self.zoom = self.MIN_ZOOM
        self.off_x = 0.0
        self.off_y = 0.0
        self.setFixedSize(self.VIEW, self.VIEW)
        self.setToolTip("Колесо — приближение, средняя кнопка — перемещение")

        self.image = image            # hires 512x512
        self.brush_color = color
        self.brush_size = 1
        self.eraser_mode = False
        self.symmetry = False
        self.drawing = False
        self.panning = False
        self._pan_last = None
        self._last_cell = None

        self.scale_mode = False
        self.scale = 1.0
        self.source = None
        self.bbox = None              # (x0, y0, x1, y1) в клетках, дробные
        self.drag_corner = None
        self.on_scale = None

    # ---------- вид ----------
    def clamp_view(self):
        total = self.grid_size * self.zoom
        self.off_x = min(0.0, max(self.VIEW - total, self.off_x))
        self.off_y = min(0.0, max(self.VIEW - total, self.off_y))

    def mid(self) -> tuple[float, float]:
        half = (self.grid_size // 2) * self.zoom
        return self.off_x + half, self.off_y + half

    def wheelEvent(self, event):
        dy = event.angleDelta().y()
        if dy == 0:
            return
        p = event.position()
        factor = 1.15 if dy > 0 else 1 / 1.15
        new_zoom = max(self.MIN_ZOOM, min(self.MAX_ZOOM, self.zoom * factor))
        if abs(new_zoom - self.zoom) < 1e-9:
            return
        gx = (p.x() - self.off_x) / self.zoom
        gy = (p.y() - self.off_y) / self.zoom
        self.zoom = new_zoom
        self.off_x = p.x() - gx * new_zoom
        self.off_y = p.y() - gy * new_zoom
        self.clamp_view()
        self.update()
        event.accept()

    # ---------- рисование ----------
    @staticmethod
    def _line_cells(a, b):
        """Клетки на отрезке a->b (Брезенхэм), чтобы быстрый штрих не рвался."""
        x0, y0 = a
        x1, y1 = b
        dx, dy = abs(x1 - x0), -abs(y1 - y0)
        sx = 1 if x0 < x1 else -1
        sy = 1 if y0 < y1 else -1
        err = dx + dy
        cells = []
        while True:
            cells.append((x0, y0))
            if x0 == x1 and y0 == y1:
                return cells
            e2 = 2 * err
            if e2 >= dy:
                err += dy
                x0 += sx
            if e2 <= dx:
                err += dx
                y0 += sy

    def paint_at(self, pos, erase: bool):
        gx = int(math.floor((pos.x() - self.off_x) / self.zoom))
        gy = int(math.floor((pos.y() - self.off_y) / self.zoom))
        cur = (gx, gy)
        if cur == self._last_cell:
            return
        cells = self._line_cells(self._last_cell, cur) if self._last_cell else [cur]
        self._last_cell = cur

        color = QColor(0, 0, 0, 0) if erase else self.brush_color
        n = self.brush_size
        offset = n // 2
        last = self.grid_size - 1

        p = QPainter(self.image)       # QPainter сам обрезает всё за пределами холста
        p.setCompositionMode(QPainter.CompositionMode.CompositionMode_Source)
        for cx_, cy_ in cells:
            points = [(cx_, cy_)]
            if self.symmetry:
                points.extend([(last - cx_, cy_), (cx_, last - cy_), (last - cx_, last - cy_)])
            for px, py in points:
                p.fillRect((px - offset) * SS, (py - offset) * SS, n * SS, n * SS, color)
        p.end()
        self.update()

    def clear(self):
        self.image.fill(Qt.GlobalColor.transparent)
        if self.source is not None:
            self.source.fill(Qt.GlobalColor.transparent)
        self.compute_bbox()
        self.update()

    # ---------- масштабирование ----------
    def transformed(self, src: QImage, s: float) -> QImage:
        """src, увеличенный/уменьшенный на s вокруг центра холста. Без сглаживания: края точные."""
        n = src.width()
        out = QImage(n, n, QImage.Format.Format_ARGB32_Premultiplied)
        out.fill(Qt.GlobalColor.transparent)
        p = QPainter(out)
        p.translate(n / 2.0, n / 2.0)
        p.scale(s, s)
        p.drawImage(QRectF(-n / 2.0, -n / 2.0, n, n), src)
        p.end()
        return out

    def compute_bbox(self):
        img = self.source if self.source is not None else self.image
        w = img.width()
        ptr = img.constBits()
        ptr.setsize(img.sizeInBytes())
        alpha = bytes(ptr)[3::4]            # альфа-канал (ARGB32 little-endian)
        stride = img.bytesPerLine() // 4
        y0 = y1 = None
        x0, x1 = w, 0
        for y in range(img.height()):
            row = alpha[y * stride: y * stride + w]
            stripped_l = row.lstrip(b"\x00")
            if not stripped_l:
                continue
            first = w - len(stripped_l)
            last = len(row.rstrip(b"\x00"))
            x0 = min(x0, first)
            x1 = max(x1, last)
            if y0 is None:
                y0 = y
            y1 = y + 1
        if y0 is None:
            self.bbox = None
        else:
            self.bbox = (x0 / SS, y0 / SS, x1 / SS, y1 / SS)

    def begin_scale(self):
        self.source = self.image.copy()
        self.scale = 1.0
        self.scale_mode = True
        self.compute_bbox()
        self.update()

    def apply_scale(self):
        self.source = None
        self.scale = 1.0
        self.scale_mode = False
        self.compute_bbox()
        self.update()

    def corners(self):
        if not self.bbox:
            return []
        c = self.grid_size / 2.0
        x0, y0, x1, y1 = self.bbox
        return [(x0 - c, y0 - c), (x1 - c, y0 - c), (x0 - c, y1 - c), (x1 - c, y1 - c)]

    def set_scale(self, s: float):
        self.scale = max(0.05, min(20.0, s))
        if self.source is not None:
            self.image = self.transformed(self.source, self.scale)
        self.update()
        if self.on_scale:
            self.on_scale(self.scale)

    # ---------- мышь ----------
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.panning = True
            self._pan_last = event.position()
            return
        if self.scale_mode:
            if event.button() == Qt.MouseButton.LeftButton:
                mx0, my0 = self.mid()
                p = event.position()
                for cx_, cy_ in self.corners():
                    if abs(cx_) < 1e-9 and abs(cy_) < 1e-9:
                        continue
                    hx = mx0 + cx_ * self.scale * self.zoom
                    hy = my0 + cy_ * self.scale * self.zoom
                    if abs(p.x() - hx) < 12 and abs(p.y() - hy) < 12:
                        self.drag_corner = (cx_, cy_)
                        break
            return
        if event.button() in (Qt.MouseButton.LeftButton, Qt.MouseButton.RightButton):
            self.drawing = True
            self._last_cell = None
            erase = self.eraser_mode or (event.button() == Qt.MouseButton.RightButton)
            self.paint_at(event.position(), erase)

    def mouseMoveEvent(self, event):
        if self.panning:
            p = event.position()
            self.off_x += p.x() - self._pan_last.x()
            self.off_y += p.y() - self._pan_last.y()
            self._pan_last = p
            self.clamp_view()
            self.update()
            return
        if self.scale_mode:
            if self.drag_corner:
                mx0, my0 = self.mid()
                mx = (event.position().x() - mx0) / self.zoom
                my = (event.position().y() - my0) / self.zoom
                cx_, cy_ = self.drag_corner
                self.set_scale((mx * cx_ + my * cy_) / (cx_ * cx_ + cy_ * cy_))
            return
        if self.drawing:
            erase = self.eraser_mode or bool(event.buttons() & Qt.MouseButton.RightButton)
            self.paint_at(event.position(), erase)

    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.MiddleButton:
            self.panning = False
            return
        self.drawing = False
        self._last_cell = None
        self.drag_corner = None

    # ---------- отрисовка ----------
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), QColor(42, 45, 52))

        total = self.grid_size * self.zoom
        # nearest-neighbor: края фигуры стоят точно, хоть на половине клетки
        painter.drawImage(QRectF(self.off_x, self.off_y, total, total), self.image)

        x_lo, x_hi = max(0.0, self.off_x), min(float(self.width()), self.off_x + total)
        y_lo, y_hi = max(0.0, self.off_y), min(float(self.height()), self.off_y + total)

        # подсетка по половинкам клеток при сильном приближении
        if self.zoom >= 20:
            painter.setPen(QPen(QColor(255, 255, 255, 22), 1, Qt.PenStyle.DotLine))
            for i in range(self.grid_size):
                x = self.off_x + (i + 0.5) * self.zoom
                y = self.off_y + (i + 0.5) * self.zoom
                if x_lo <= x <= x_hi:
                    painter.drawLine(QPointF(x, y_lo), QPointF(x, y_hi))
                if y_lo <= y <= y_hi:
                    painter.drawLine(QPointF(x_lo, y), QPointF(x_hi, y))

        # основная сетка: каждая клетка + яркие линии каждые 8 клеток
        for i in range(self.grid_size + 1):
            major = (i % 8 == 0)
            painter.setPen(QPen(QColor(255, 255, 255, 120 if major else 55), 1))
            x = self.off_x + i * self.zoom
            y = self.off_y + i * self.zoom
            if x_lo <= x <= x_hi:
                painter.drawLine(QPointF(x, y_lo), QPointF(x, y_hi))
            if y_lo <= y <= y_hi:
                painter.drawLine(QPointF(x_lo, y), QPointF(x_hi, y))

        # центр (между 4 центральными клетками)
        mx0, my0 = self.mid()
        painter.setPen(QPen(QColor(255, 70, 70, 190), 1))
        painter.drawLine(QPointF(mx0, y_lo), QPointF(mx0, y_hi))
        painter.drawLine(QPointF(x_lo, my0), QPointF(x_hi, my0))

        if self.scale_mode and self.bbox:
            pts = [(mx0 + x * self.scale * self.zoom, my0 + y * self.scale * self.zoom)
                   for x, y in self.corners()]
            xs = [p[0] for p in pts]
            ys = [p[1] for p in pts]
            painter.setPen(QPen(QColor(138, 180, 248), 1, Qt.PenStyle.DashLine))
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawRect(QRectF(min(xs), min(ys), max(xs) - min(xs), max(ys) - min(ys)))
            painter.setPen(QPen(QColor(255, 255, 255), 1))
            painter.setBrush(QBrush(QColor(138, 180, 248)))
            for x, y in pts:
                painter.drawRect(QRectF(x - 5, y - 5, 10, 10))
class CrosshairDrawDialog(QDialog):
    """Модальное окно рисования собственного прицела."""

    def __init__(self, parent, save_path: str, initial_color: str, initial_scale: float = 1.0, editor_scale: float = 1.0):
        super().__init__(parent)
        self.save_path = save_path
        self.current_color = QColor(initial_color)
        self.scale = 1.0
        self.total_scale = max(0.01, float(editor_scale))
        # масштаб запекается в рисунок, снаружи всегда 1.0
        self.final_image = None       # итоговый 64x64 для оверлея

        self.setWindowTitle("Редактор прицела (64x64 px)")
        self.setWindowFlags(
            Qt.WindowType.Dialog
            | Qt.WindowType.WindowCloseButtonHint
            | Qt.WindowType.WindowStaysOnTopHint
        )
        self.setModal(True)

        layout = QVBoxLayout(self)

        tools_row = QHBoxLayout()

        self.btn_brush = QPushButton("Кисть (ЛКМ)")
        self.btn_brush.setCheckable(True)
        self.btn_brush.setChecked(True)
        self.btn_brush.clicked.connect(lambda: self.set_tool(False))

        self.btn_eraser = QPushButton("Ластик (ПКМ)")
        self.btn_eraser.setCheckable(True)
        self.btn_eraser.clicked.connect(lambda: self.set_tool(True))

        self.btn_color = QPushButton("Цвет")
        self.update_color_btn()
        self.btn_color.clicked.connect(self.pick_color)

        btn_clear = QPushButton("Очистить")
        btn_clear.clicked.connect(self.clear_canvas)

        tools_row.addWidget(self.btn_brush)
        tools_row.addWidget(self.btn_eraser)
        tools_row.addWidget(self.btn_color)
        tools_row.addWidget(btn_clear)
        layout.addLayout(tools_row)

        self.canvas = PixelCanvas(load_hires(save_path), self.current_color)
        self.canvas.on_scale = self.on_canvas_scale

        # старый сохранённый draw_scale один раз запекаем в рисунок
        if abs(initial_scale - 1.0) > 1e-3:
            self.canvas.image = self.canvas.transformed(self.canvas.image, initial_scale)

        opts_row = QHBoxLayout()
        opts_row.addWidget(QLabel("Размер кисти:"))
        self.spin_size = QSpinBox()
        self.spin_size.setRange(1, 6)
        self.spin_size.setValue(1)
        self.spin_size.valueChanged.connect(lambda v: setattr(self.canvas, "brush_size", v))
        opts_row.addWidget(self.spin_size)

        self.cb_sym = QCheckBox("Симметрия по X и Y")
        self.cb_sym.toggled.connect(lambda v: setattr(self.canvas, "symmetry", v))
        opts_row.addWidget(self.cb_sym)
        opts_row.addStretch()
        layout.addLayout(opts_row)

        scale_row = QHBoxLayout()
        self.cb_scale = QCheckBox("Режим масштаба (тяните за угол)")
        self.cb_scale.toggled.connect(self.toggle_scale_mode)
        self.spin_scale = QDoubleSpinBox()
        self.spin_scale.setRange(0.05, 20.0)
        self.spin_scale.setDecimals(2)
        self.spin_scale.setSingleStep(0.05)
        self.spin_scale.setPrefix("× ")
        self.spin_scale.setValue(1.0)
        self.spin_scale.setEnabled(False)
        self.spin_scale.valueChanged.connect(self.canvas.set_scale)
        self.lbl_size = QLabel("")
        scale_row.addWidget(self.cb_scale)
        scale_row.addWidget(self.spin_scale)
        scale_row.addWidget(self.lbl_size)
        scale_row.addStretch()
        layout.addLayout(scale_row)

        layout.addWidget(self.canvas, alignment=Qt.AlignmentFlag.AlignCenter)

        bottom_row = QHBoxLayout()
        btn_cancel = QPushButton("Отмена")
        btn_cancel.clicked.connect(self.reject)

        btn_save = QPushButton("Применить и сохранить")
        btn_save.setStyleSheet("background-color: #2e6f40; font-weight: bold;")
        btn_save.clicked.connect(self.save_and_close)

        bottom_row.addWidget(btn_cancel)
        bottom_row.addWidget(btn_save)
        layout.addLayout(bottom_row)

        self.update_size_label()

    def set_tool(self, eraser: bool):
        self.canvas.eraser_mode = eraser
        self.btn_brush.setChecked(not eraser)
        self.btn_eraser.setChecked(eraser)

    def pick_color(self):
        c = QColorDialog.getColor(self.current_color, self, "Цвет кисти")
        if c.isValid():
            self.current_color = c
            self.canvas.brush_color = c
            self.update_color_btn()

    def update_color_btn(self):
        c = self.current_color.name()
        self.btn_color.setStyleSheet(f"background-color: {c}; color: #000; font-weight: bold;")

    def clear_canvas(self):
        self.canvas.clear()
        self.update_size_label()

    def toggle_scale_mode(self, on: bool):
        self.spin_scale.blockSignals(True)
        self.spin_scale.setValue(1.0)
        self.spin_scale.blockSignals(False)
        self.spin_scale.setEnabled(on)
        if on:
            self.canvas.begin_scale()
        else:
            self.canvas.apply_scale()
        self.update_size_label()

    def update_size_label(self):
        if not self.canvas.scale_mode:
            self.canvas.compute_bbox()
        s = self.canvas.scale
        if self.canvas.bbox:
            x0, y0, x1, y1 = self.canvas.bbox
            self.lbl_size.setText(f"{(x1 - x0) * s:.2f} × {(y1 - y0) * s:.2f} px")
        else:
            self.lbl_size.setText("")

    def on_canvas_scale(self, s: float):
        self.spin_scale.blockSignals(True)
        self.spin_scale.setValue(s)
        self.spin_scale.blockSignals(False)
        self.update_size_label()

    def save_and_close(self):
        if self.canvas.scale_mode:
            self.cb_scale.blockSignals(True)
            self.cb_scale.setChecked(False)
            self.cb_scale.blockSignals(False)
            self.canvas.apply_scale()
        self.final_image = to_final(self.canvas.image)
        self.final_image.save(self.save_path, "PNG")                         # для оверлея
        self.canvas.image.save(hires_path(self.save_path), "PNG")            # точная версия для редактора
        self.scale = 1.0
        self.accept()

class CrosshairRenderer:
    """Общий рендерер прицела. Масштаб под растянутое разрешение всегда автоматический."""
    _cached_pixmap: QPixmap | None = None
    _cached_path: str = ""
    _cached_mtime: float = 0.0
    _display_scale: tuple[float, float] = (1.0, 1.0)
    _display_scale_initialized: bool = False

    @classmethod
    def refresh_display_scale(cls) -> tuple[float, float]:
        res = get_primary_display_resolutions()
        if res is not None:
            cur_w, cur_h, native_w, native_h, scaling = res
            sx = max(0.25, min(4.0, cur_w / native_w))
            sy = max(0.25, min(4.0, cur_h / native_h))
            if scaling == _SCALING_CENTERED:
                sx = sy = 1.0                       # монитор не растягивает картинку
            elif scaling == _SCALING_ASPECT:
                sx = sy = max(sx, sy)               # растяжение равномерное, по меньшей стороне
            cls._display_scale = (sx, sy)
            cls._display_scale_initialized = True
        return cls._display_scale

    @classmethod
    def get_effective_scale(cls, cfg: dict) -> tuple[float, float]:
        if not cls._display_scale_initialized:
            cls.refresh_display_scale()
        return cls._display_scale

    @classmethod
    def set_drawn_image(cls, qimage: QImage, path: str):
        cls._cached_pixmap = QPixmap.fromImage(qimage)
        cls._cached_path = path
        cls._cached_mtime = os.path.getmtime(path) if os.path.exists(path) else 0.0

    @classmethod
    def get_pixmap(cls, path: str) -> QPixmap | None:
        if not path:
            return None
        if not os.path.isabs(path):
            path = os.path.join(BASE_DIR, path)
        if not os.path.exists(path):
            return None
        try:
            mtime = os.path.getmtime(path)
            if cls._cached_pixmap is None or cls._cached_path != path or cls._cached_mtime != mtime:
                img = QImage(path)
                if img.isNull():
                    return None
                cls._cached_pixmap = QPixmap.fromImage(img)
                cls._cached_path = path
                cls._cached_mtime = mtime
            return cls._cached_pixmap
        except Exception:
            return None

    @classmethod
    def invalidate_cache(cls):
        cls._cached_pixmap = None
        cls._cached_path = ""
        cls._cached_mtime = 0.0

    @staticmethod
    def _snap(x: float, y: float, w: float, h: float) -> tuple[int, int, int, int]:
        """Прямоугольник в дробных координатах -> целые пиксели (каждая грань округляется отдельно).
        При центре на границе пикселей (чётный экран) чётная толщина ложится ровно по центру."""
        x0, y0 = math.floor(x + 0.5), math.floor(y + 0.5)
        x1, y1 = math.floor(x + w + 0.5), math.floor(y + h + 0.5)
        return x0, y0, max(1, x1 - x0), max(1, y1 - y0)

    @classmethod
    def draw(cls, painter: QPainter, cx: float, cy: float, cfg: dict, is_preview: bool = False):
        """cx, cy — точный геометрический центр (может быть на границе пикселей)."""
        painter.save()

        sx, sy = cls.get_effective_scale(cfg)
        is_scaled = abs(sx - 1.0) > 1e-3 or abs(sy - 1.0) > 1e-3

        style = cfg.get("style", "classic")
        painter.setRenderHint(QPainter.RenderHint.Antialiasing, style in ("circle_dot", "x_cross"))
        painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, style == "image" or is_scaled)

        base_color = QColor(cfg.get("color", "#00FF66"))
        base_color.setAlpha(cfg.get("opacity", 255))

        outline_color = QColor(0, 0, 0, min(255, cfg.get("opacity", 255)))
        has_outline = cfg.get("outline", True)
        out_t = cfg.get("outline_thickness", 1)
        out_x = max(1, round(out_t * sx)) if out_t > 0 else 0
        out_y = max(1, round(out_t * sy)) if out_t > 0 else 0

        length = cfg.get("length", 10)
        thick = cfg.get("thickness", 2)
        gap = cfg.get("gap", 4)

        if style in ("image", "draw"):
            path = cfg.get("drawn_path", DRAWN_IMAGE_FILE) if style == "draw" else cfg.get("image_path", "")
            pixmap = cls.get_pixmap(path)
            if pixmap and not pixmap.isNull():
                if style == "draw":
                    scale = max(0.05, float(cfg.get("draw_scale", 1.0)))
                    painter.setRenderHint(QPainter.RenderHint.SmoothPixmapTransform, True)
                else:
                    scale = max(10, cfg.get("image_scale", 100)) / 100.0
                w = pixmap.width() * scale * sx
                h = pixmap.height() * scale * sy
                painter.setOpacity(cfg.get("opacity", 255) / 255.0)
                painter.drawPixmap(
                    QRectF(cx - w / 2.0, cy - h / 2.0, w, h),
                    pixmap,
                    QRectF(0.0, 0.0, float(pixmap.width()), float(pixmap.height())),
                )
                painter.setOpacity(1.0)
            painter.restore()
            return

        if style in ("classic", "t_shape"):
            len_x = max(1, round(length * sx))
            len_y = max(1, round(length * sy))
            thick_x = max(1, round(thick * sx))
            thick_y = max(1, round(thick * sy))
            gap_x = round(gap * sx)
            gap_y = round(gap * sy)

            frects = [
                (cx - thick_x / 2, cy + gap_y, thick_x, len_y),
                (cx - gap_x - len_x, cy - thick_y / 2, len_x, thick_y),
                (cx + gap_x, cy - thick_y / 2, len_x, thick_y),
            ]
            if style == "classic":
                frects.append((cx - thick_x / 2, cy - gap_y - len_y, thick_x, len_y))
            rects = [cls._snap(*r) for r in frects]

            painter.setPen(Qt.PenStyle.NoPen)
            if has_outline:
                painter.setBrush(QBrush(outline_color))
                for rx, ry, rw, rh in rects:
                    painter.drawRect(rx - out_x, ry - out_y, rw + out_x * 2, rh + out_y * 2)

            painter.setBrush(QBrush(base_color))
            for rx, ry, rw, rh in rects:
                painter.drawRect(rx, ry, rw, rh)

        elif style == "circle_dot":
            painter.save()
            painter.translate(cx, cy)
            painter.scale(sx, sy)
            painter.translate(-cx, -cy)

            r = cfg.get("circle_radius", 14)
            if has_outline:
                pen_out = QPen(outline_color, thick + out_t * 2)
                painter.setPen(pen_out)
                painter.setBrush(Qt.BrushStyle.NoBrush)
                painter.drawEllipse(QPointF(cx, cy), r, r)

            pen_main = QPen(base_color, thick)
            painter.setPen(pen_main)
            painter.setBrush(Qt.BrushStyle.NoBrush)
            painter.drawEllipse(QPointF(cx, cy), r, r)
            painter.restore()

        elif style == "x_cross":
            painter.save()
            painter.translate(cx, cy)
            painter.scale(sx, sy)
            painter.translate(-cx, -cy)

            diag_gap = gap / math.sqrt(2)
            diag_len = length / math.sqrt(2)
            angles = [45, 135, 225, 315]

            for ang in angles:
                rad = math.radians(ang)
                x1 = cx + math.cos(rad) * diag_gap
                y1 = cy + math.sin(rad) * diag_gap
                x2 = cx + math.cos(rad) * (diag_gap + diag_len)
                y2 = cy + math.sin(rad) * (diag_gap + diag_len)

                if has_outline:
                    pen_out = QPen(outline_color, thick + out_t * 2, Qt.PenStyle.SolidLine, Qt.PenCapStyle.FlatCap)
                    painter.setPen(pen_out)
                    painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))

                pen_main = QPen(base_color, thick, Qt.PenStyle.SolidLine, Qt.PenCapStyle.FlatCap)
                painter.setPen(pen_main)
                painter.drawLine(QPointF(x1, y1), QPointF(x2, y2))
            painter.restore()

        if style == "dot" or cfg.get("dot", True):
            d = cfg.get("dot_size", 2)
            dot_w = max(1, round(d * sx))
            dot_h = max(1, round(d * sy))
            dx, dy, dw, dh = cls._snap(cx - dot_w / 2, cy - dot_h / 2, dot_w, dot_h)
            painter.setPen(Qt.PenStyle.NoPen)
            if has_outline:
                painter.setBrush(QBrush(outline_color))
                painter.drawRect(dx - out_x, dy - out_y, dw + out_x * 2, dh + out_y * 2)
            painter.setBrush(QBrush(base_color))
            painter.drawRect(dx, dy, dw, dh)

        painter.restore()


class CrosshairOverlay(QWidget):
    """Прозрачное кликабельное насквозь окно поверх игры."""

    def __init__(self, config: dict):
        super().__init__()
        self.cfg = config

        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint
            | Qt.WindowType.WindowStaysOnTopHint
            | Qt.WindowType.Tool
            | Qt.WindowType.WindowTransparentForInput
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)
        self.setAttribute(Qt.WidgetAttribute.WA_ShowWithoutActivating)

        # запас под большие масштабы рисунка; точный центр хранится отдельно (дробный)
        self.overlay_size = 640
        self._center = (320.0, 320.0)
        self._native_pos = (0, 0)
        self._gc = (0.0, 0.0)
        self._last_screen_geom = None
        self._last_scale = None
        self._cache: QPixmap | None = None
        self.update_position()

        primary = QApplication.primaryScreen()
        if primary:
            primary.geometryChanged.connect(self.on_screen_changed)

        # Смена активного окна приходит событием от Windows — без опроса каждые 150 мс
        self._hook_proc = _WINEVENTPROC(self._on_foreground)   # держим ссылку, иначе GC
        self._hook = _user32_hook.SetWinEventHook(
            _EVENT_SYSTEM_FOREGROUND, _EVENT_SYSTEM_FOREGROUND, None, self._hook_proc,
            0, 0, _WINEVENT_OUTOFCONTEXT | _WINEVENT_SKIPOWNPROCESS
        )

        # Медленный таймер: разрешение экрана, возврат TOPMOST и запасной вариант без хука
        self.timer = QTimer(self)
        self.timer.timeout.connect(self.slow_tick)
        self.timer.start(1000 if self._hook else 200)

    def _on_foreground(self, *_args):
        QTimer.singleShot(0, self.check_target)

    def unhook(self):
        if self._hook:
            _user32_hook.UnhookWinEvent(self._hook)
            self._hook = None

    def on_screen_changed(self, _geom=None):
        CrosshairRenderer.refresh_display_scale()
        self.update_position()

    def refresh(self):
        """Сбросить кэш картинки прицела и перерисовать."""
        self._cache = None
        self.update()

    # Оверлей работает в ФИЗИЧЕСКИХ пикселях экрана, независимо от масштаба Windows (125%, 150%...):
    # толщина 2 = ровно 2 пикселя, центр = точный центр монитора.

    @staticmethod
    def _monitor_rect() -> tuple[int, int, int, int]:
        """Физический прямоугольник основного монитора (x, y, w, h)."""
        try:
            hmon = win32api.MonitorFromPoint((0, 0), 1)       # MONITOR_DEFAULTTOPRIMARY
            l, t, r, b = win32api.GetMonitorInfo(hmon)["Monitor"]
            return l, t, r - l, b - t
        except Exception:
            s = QApplication.primaryScreen()
            g, d = s.geometry(), s.devicePixelRatio()
            return round(g.x() * d), round(g.y() * d), round(g.width() * d), round(g.height() * d)

    def update_position(self):
        mx, my, mw, mh = self._monitor_rect()
        sx, sy = CrosshairRenderer.get_effective_scale(self.cfg)
        self._last_screen_geom = (mx, my, mw, mh, self.devicePixelRatioF())
        self._last_scale = (sx, sy)
        # точный центр: при чётном разрешении это граница между пикселями (1280.0 на 2560)
        gcx = mx + mw / 2 + round(self.cfg.get("offset_x", 0) * sx)
        gcy = my + mh / 2 + round(self.cfg.get("offset_y", 0) * sy)
        half = self.overlay_size / 2
        self._gc = (gcx, gcy)
        # Qt хранит позицию в логических единицах и сам округляет её до физических пикселей
        # (при 150% окно нельзя поставить в любой пиксель). Поэтому окно ставим примерно,
        # а прицел внутри него сдвигаем так, чтобы он попал ровно в центр экрана.
        dpr = self.devicePixelRatioF()
        lx = math.floor((gcx - half) / dpr + 0.5)
        ly = math.floor((gcy - half) / dpr + 0.5)
        side = math.ceil(self.overlay_size / dpr) + 1
        self.setGeometry(lx, ly, side, side)
        self._set_origin(math.floor(lx * dpr + 0.5), math.floor(ly * dpr + 0.5))

    def _set_origin(self, nx: int, ny: int):
        """Физический левый-верхний угол окна -> центр прицела внутри окна."""
        self._native_pos = (nx, ny)
        self._center = (self._gc[0] - nx, self._gc[1] - ny)
        self.refresh()

    def _sync_actual_pos(self):
        """Сверка с реальной позицией окна: если Qt округлил иначе — перерисовать со сдвигом."""
        try:
            l, t, _r, _b = win32gui.GetWindowRect(int(self.winId()))
        except Exception:
            return
        if (l, t) != self._native_pos:
            self._set_origin(l, t)

    def _raise_topmost(self):
        try:
            win32gui.SetWindowPos(int(self.winId()), _HWND_TOPMOST, 0, 0, 0, 0, _SWP_FLAGS)
        except Exception:
            pass

    def slow_tick(self):
        CrosshairRenderer.refresh_display_scale()
        cur = (*self._monitor_rect(), self.devicePixelRatioF())
        if cur != self._last_screen_geom or CrosshairRenderer._display_scale != self._last_scale:
            self.update_position()
        self.check_target()
        # игры в borderless иногда перекрывают оверлей — тихо возвращаем его наверх
        if self.isVisible():
            self._sync_actual_pos()
            self._raise_topmost()

    def _should_show(self) -> bool:
        if not self.cfg.get("enabled", True):
            return False
        if self.cfg.get("always_on", False):
            return True
        target = self.cfg.get("target_exe", "").strip().lower()
        return bool(target) and get_active_exe().lower() == target

    def check_target(self):
        show = self._should_show()
        if show and not self.isVisible():
            self.show()
            self.raise_()
            QTimer.singleShot(0, self._sync_actual_pos)
        elif not show and self.isVisible():
            self.hide()

    def _render(self) -> QPixmap:
        # рисуем 1:1 в физических пикселях, потом помечаем DPR — Qt выведет картинку без масштабирования
        pm = QPixmap(self.overlay_size, self.overlay_size)
        pm.fill(Qt.GlobalColor.transparent)
        p = QPainter(pm)
        CrosshairRenderer.draw(p, self._center[0], self._center[1], self.cfg, is_preview=False)
        p.end()
        pm.setDevicePixelRatio(self.devicePixelRatioF())
        return pm

    def paintEvent(self, event):
        if self._cache is None or abs(self._cache.devicePixelRatio() - self.devicePixelRatioF()) > 1e-6:
            self._cache = self._render()
        painter = QPainter(self)
        painter.drawPixmap(0, 0, self._cache)


class PreviewWidget(QWidget):
    """Виджет предпросмотра прицела в окне настроек."""

    def __init__(self, config: dict):
        super().__init__()
        self.cfg = config
        self.setMinimumSize(180, 180)
        # шахматный фон — одна текстура 30x30, заливается кистью
        tile = 15
        tex = QPixmap(tile * 2, tile * 2)
        tex.fill(QColor(35, 38, 44))
        p = QPainter(tex)
        p.fillRect(0, 0, tile, tile, QColor(45, 48, 55))
        p.fillRect(tile, tile, tile, tile, QColor(45, 48, 55))
        p.end()
        self._bg = QBrush(tex)

    def paintEvent(self, event):
        painter = QPainter(self)
        painter.fillRect(self.rect(), self._bg)

        # в физических пикселях, как и оверлей: превью совпадает с реальным прицелом при любом масштабе Windows
        dpr = self.devicePixelRatioF()
        painter.scale(1 / dpr, 1 / dpr)
        cx = math.floor(self.width() * dpr / 2)
        cy = math.floor(self.height() * dpr / 2)
        CrosshairRenderer.draw(painter, cx, cy, self.cfg, is_preview=True)


class SettingsWindow(QMainWindow):
    """Главное окно настроек приложения."""

    def __init__(self):
        super().__init__()
        self.cfg = self.load_config()
        self.sliders: dict[str, tuple[QSlider, QLabel]] = {}
        self.overlay = CrosshairOverlay(self.cfg)

        # запись конфига откладывается: слайдер не пишет файл на каждый шаг
        self._save_timer = QTimer(self)
        self._save_timer.setSingleShot(True)
        self._save_timer.setInterval(400)
        self._save_timer.timeout.connect(self.save_config)

        self.setWindowTitle("Custom Crosshair Overlay")
        self.setFixedSize(485, 700)
        self.apply_dark_theme()

        self.init_ui()
        self.init_tray()
        self.refresh_process_list()

        primary = QApplication.primaryScreen()
        if primary:
            primary.geometryChanged.connect(lambda _g: self.preview.update())

    def load_config(self) -> dict:
        if os.path.exists(CONFIG_FILE):
            try:
                with open(CONFIG_FILE, "r", encoding="utf-8") as f:
                    data = json.load(f)
                    if isinstance(data, dict):
                        merged = {**DEFAULT_CONFIG, **data}
                        # битые/чужие типы (например "length": "10") роняли окно при старте
                        for k, v in DEFAULT_CONFIG.items():
                            cur = merged.get(k)
                            if isinstance(v, bool) or isinstance(cur, bool):
                                ok = type(cur) is type(v)
                            elif isinstance(v, (int, float)):
                                ok = isinstance(cur, (int, float))
                                if ok and isinstance(v, int):
                                    merged[k] = int(cur)
                            else:
                                ok = isinstance(cur, type(v))
                            if not ok:
                                merged[k] = v
                        merged["drawn_path"] = DRAWN_IMAGE_FILE
                        merged.pop("stretch_x", None)  # параметр удалён, всегда авто
                        merged["config_version"] = 3
                        return merged
            except Exception:
                pass
        return DEFAULT_CONFIG.copy()

    def save_config(self):
        try:
            # атомарно: при сбое/выключении ПК во время записи старый конфиг не портится
            tmp = CONFIG_FILE + ".tmp"
            with open(tmp, "w", encoding="utf-8") as f:
                json.dump(self.cfg, f, indent=4, ensure_ascii=False)
            os.replace(tmp, CONFIG_FILE)
        except Exception as e:
            print("Ошибка сохранения конфига:", e)

    def apply_dark_theme(self):
        self.setStyleSheet("""
            QMainWindow, QWidget, QDialog { background-color: #1e2024; color: #e4e6eb; font-family: 'Segoe UI', sans-serif; font-size: 13px; }
            QGroupBox { border: 1px solid #343840; border-radius: 6px; margin-top: 10px; padding-top: 12px; font-weight: bold; color: #8ab4f8; }
            QGroupBox::title { subcontrol-origin: margin; left: 10px; padding: 0 4px; }
            QPushButton { background-color: #2d3139; border: 1px solid #3e4451; border-radius: 4px; padding: 6px 12px; color: #fff; }
            QPushButton:hover { background-color: #3a3f4b; }
            QPushButton:checked { background-color: #3b5998; border-color: #8ab4f8; }
            QComboBox, QSpinBox, QDoubleSpinBox { background-color: #282c34; border: 1px solid #3e4451; border-radius: 4px; padding: 5px 8px; color: #fff; }
            QComboBox QAbstractItemView { background-color: #282c34; color: #fff; selection-background-color: #3a3f4b; }
            QSlider::groove:horizontal { height: 6px; background: #2d3139; border-radius: 3px; }
            QSlider::handle:horizontal { background: #8ab4f8; width: 14px; margin: -4px 0; border-radius: 7px; }
        """)

    def init_ui(self):
        central = QWidget()
        self.setCentralWidget(central)
        main_layout = QVBoxLayout(central)

        app_group = QGroupBox("1. Привязка к приложению")
        app_layout = QVBoxLayout(app_group)

        proc_row = QHBoxLayout()
        self.proc_combo = QComboBox()
        self.proc_combo.currentIndexChanged.connect(self.on_process_changed)

        btn_refresh = QPushButton("Обновить")
        btn_refresh.setFixedWidth(85)
        btn_refresh.clicked.connect(self.refresh_process_list)

        proc_row.addWidget(self.proc_combo)
        proc_row.addWidget(btn_refresh)
        app_layout.addLayout(proc_row)

        self.cb_always_on = QCheckBox("Показывать поверх ВСЕХ окон (для теста и настройки)")
        self.cb_always_on.setChecked(self.cfg["always_on"])
        self.cb_always_on.toggled.connect(lambda v: self.update_cfg("always_on", v))
        app_layout.addWidget(self.cb_always_on)

        self.cb_enabled = QCheckBox("Включить прицел")
        self.cb_enabled.setChecked(self.cfg["enabled"])
        self.cb_enabled.toggled.connect(self.on_enabled_toggled)
        app_layout.addWidget(self.cb_enabled)

        main_layout.addWidget(app_group)

        top_row = QHBoxLayout()

        self.preview = PreviewWidget(self.cfg)
        top_row.addWidget(self.preview)

        style_group = QGroupBox("2. Форма и цвет")
        style_form = QFormLayout(style_group)

        self.style_combo = QComboBox()
        self.styles_map = {
            "classic": "Классический (+)",
            "t_shape": "Т-образный (T)",
            "dot": "Только точка (•)",
            "circle_dot": "Круг + точка (◎)",
            "x_cross": "Косой крест (✕)",
            "image": "Своя картинка (PNG)",
            "draw": "Нарисовать свой (редактор)"
        }
        for k, v in self.styles_map.items():
            self.style_combo.addItem(v, k)
        idx = list(self.styles_map.keys()).index(self.cfg["style"]) if self.cfg["style"] in self.styles_map else 0
        self.style_combo.setCurrentIndex(idx)
        self.style_combo.activated.connect(self.on_style_activated)
        style_form.addRow("Тип:", self.style_combo)

        self.btn_color = QPushButton()
        self.update_color_btn()
        self.btn_color.clicked.connect(self.pick_color)
        style_form.addRow("Цвет:", self.btn_color)

        self.btn_image = QPushButton("Выбрать PNG...")
        self.btn_image.clicked.connect(self.pick_image)
        style_form.addRow("Файл:", self.btn_image)

        self.btn_open_draw = QPushButton("Открыть редактор...")
        self.btn_open_draw.clicked.connect(self.open_draw_editor)
        style_form.addRow("Рисунок:", self.btn_open_draw)

        self.cb_dot = QCheckBox("Центральная точка")
        self.cb_dot.setChecked(self.cfg["dot"])
        self.cb_dot.toggled.connect(lambda v: self.update_cfg("dot", v))
        style_form.addRow(self.cb_dot)

        self.cb_outline = QCheckBox("Черная обводка")
        self.cb_outline.setChecked(self.cfg["outline"])
        self.cb_outline.toggled.connect(lambda v: self.update_cfg("outline", v))
        style_form.addRow(self.cb_outline)

        top_row.addWidget(style_group)
        main_layout.addLayout(top_row)

        geom_group = QGroupBox("3. Размеры и позиционирование")
        geom_form = QFormLayout(geom_group)

        self.add_slider(geom_form, "Длина линий:", "length", 2, 50)
        self.add_slider(geom_form, "Толщина:", "thickness", 1, 15)
        self.add_slider(geom_form, "Зазор (отступ):", "gap", 0, 40)
        self.add_slider(geom_form, "Размер точки:", "dot_size", 1, 15)
        self.add_slider(geom_form, "Радиус круга:", "circle_radius", 4, 60)
        self.add_slider(geom_form, "Масштаб PNG (%):", "image_scale", 10, 300)
        self.add_slider(geom_form, "Прозрачность:", "opacity", 30, 255)

        offset_row = QHBoxLayout()
        self.spin_x = QSpinBox()
        self.spin_x.setRange(-500, 500)
        self.spin_x.setValue(self.cfg["offset_x"])
        self.spin_x.valueChanged.connect(lambda v: self.update_offset("offset_x", v))

        self.spin_y = QSpinBox()
        self.spin_y.setRange(-500, 500)
        self.spin_y.setValue(self.cfg["offset_y"])
        self.spin_y.valueChanged.connect(lambda v: self.update_offset("offset_y", v))

        offset_row.addWidget(QLabel("X:"))
        offset_row.addWidget(self.spin_x)
        offset_row.addWidget(QLabel("Y:"))
        offset_row.addWidget(self.spin_y)
        geom_form.addRow("Смещение (px):", offset_row)

        main_layout.addWidget(geom_group)

        bottom_row = QHBoxLayout()
        btn_reset = QPushButton("Сбросить")
        btn_reset.clicked.connect(self.reset_defaults)

        btn_hide_tray = QPushButton("Свернуть в трей")
        btn_hide_tray.clicked.connect(self.hide)

        btn_exit = QPushButton("Закрыть приложение")
        btn_exit.setStyleSheet("""
            QPushButton { background-color: #8b2635; border: 1px solid #a83244; color: #fff; font-weight: bold; }
            QPushButton:hover { background-color: #a83244; }
        """)
        btn_exit.clicked.connect(self.quit_app)

        bottom_row.addWidget(btn_reset)
        bottom_row.addWidget(btn_hide_tray)
        bottom_row.addWidget(btn_exit)
        main_layout.addLayout(bottom_row)

    def quit_app(self):
        """Полное завершение работы приложения."""
        if self._save_timer.isActive():
            self._save_timer.stop()
            self.save_config()
        self.overlay.unhook()
        self.overlay.close()
        self.tray.hide()
        QApplication.instance().quit()

    def add_slider(self, layout: QFormLayout, label: str, key: str, min_v: int, max_v: int):
        row = QHBoxLayout()
        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(min_v, max_v)
        slider.setValue(self.cfg[key])

        val_lbl = QLabel(str(self.cfg[key]))
        val_lbl.setFixedWidth(28)

        def on_change(val):
            val_lbl.setText(str(val))
            self.update_cfg(key, val)

        slider.valueChanged.connect(on_change)
        self.sliders[key] = (slider, val_lbl)
        row.addWidget(slider)
        row.addWidget(val_lbl)
        layout.addRow(label, row)

    def refresh_process_list(self):
        saved_exe = self.cfg["target_exe"]
        self.proc_combo.blockSignals(True)
        self.proc_combo.clear()

        self.proc_combo.addItem(PLACEHOLDER_TEXT, "")

        apps = get_running_apps()
        selected_index = 0

        for i, (title, exe) in enumerate(apps, start=1):
            short_title = (title[:32] + "...") if len(title) > 32 else title
            self.proc_combo.addItem(f"{exe}  —  ({short_title})", exe)
            if saved_exe and exe.lower() == saved_exe.lower():
                selected_index = i

        if saved_exe and selected_index == 0:
            self.proc_combo.addItem(f"{saved_exe}  —  (не запущено)", saved_exe)
            selected_index = self.proc_combo.count() - 1

        self.proc_combo.setCurrentIndex(selected_index)
        self.proc_combo.blockSignals(False)

    def on_process_changed(self, index: int):
        exe_name = self.proc_combo.itemData(index) or ""
        self.update_cfg("target_exe", exe_name)
        self.overlay.check_target()

    def on_enabled_toggled(self, checked: bool):
        self.update_cfg("enabled", checked)
        self.overlay.check_target()

    def on_style_activated(self, index: int):
        style_key = self.style_combo.itemData(index)
        self.update_cfg("style", style_key)
        if style_key == "draw" and not os.path.exists(self.cfg["drawn_path"]):
            QTimer.singleShot(50, self.open_draw_editor)

    def open_draw_editor(self):
        dlg = CrosshairDrawDialog(
            self, self.cfg["drawn_path"], self.cfg["color"], float(self.cfg.get("draw_scale", 1.0))
        )
        if dlg.exec() == QDialog.DialogCode.Accepted:
            CrosshairRenderer.set_drawn_image(dlg.final_image, self.cfg["drawn_path"])
            self.update_cfg("draw_scale", dlg.scale)
            self.update_cfg("style", "draw")
            idx = list(self.styles_map.keys()).index("draw")
            self.style_combo.blockSignals(True)
            self.style_combo.setCurrentIndex(idx)
            self.style_combo.blockSignals(False)
            self.preview.update()
            self.overlay.refresh()

    def pick_color(self):
        color = QColorDialog.getColor(QColor(self.cfg["color"]), self, "Выбор цвета прицела")
        if color.isValid():
            self.update_cfg("color", color.name())
            self.update_color_btn()

    def update_color_btn(self):
        c = self.cfg["color"]
        self.btn_color.setText(c.upper())
        self.btn_color.setStyleSheet(f"background-color: {c}; color: #000; font-weight: bold;")

    def pick_image(self):
        path, _ = QFileDialog.getOpenFileName(self, "Выберите картинку прицела", "", "Images (*.png *.svg *.ico)")
        if path:
            CrosshairRenderer.invalidate_cache()
            self.update_cfg("image_path", path)
            self.update_cfg("style", "image")
            idx = list(self.styles_map.keys()).index("image")
            self.style_combo.setCurrentIndex(idx)
            self.preview.update()
            self.overlay.refresh()   # тот же путь, но файл мог измениться

    def update_cfg(self, key: str, value):
        if self.cfg.get(key) == value:
            return
        self.cfg[key] = value
        self._save_timer.start()
        self.preview.update()
        self.overlay.refresh()

    def update_offset(self, key: str, value: int):
        self.cfg[key] = value
        self._save_timer.start()
        self.overlay.update_position()

    def reset_defaults(self):
        """Сброс без перезапуска процесса (os.execl ломался на путях с пробелами
        и оставлял «мёртвую» иконку в трее)."""
        self.cfg.clear()
        self.cfg.update(DEFAULT_CONFIG)

        widgets = [self.cb_always_on, self.cb_enabled, self.cb_dot, self.cb_outline,
                   self.style_combo, self.spin_x, self.spin_y]
        widgets += [s for s, _ in self.sliders.values()]
        for w in widgets:
            w.blockSignals(True)
        self.cb_always_on.setChecked(self.cfg["always_on"])
        self.cb_enabled.setChecked(self.cfg["enabled"])
        self.cb_dot.setChecked(self.cfg["dot"])
        self.cb_outline.setChecked(self.cfg["outline"])
        self.style_combo.setCurrentIndex(list(self.styles_map.keys()).index(self.cfg["style"]))
        self.spin_x.setValue(self.cfg["offset_x"])
        self.spin_y.setValue(self.cfg["offset_y"])
        for key, (slider, lbl) in self.sliders.items():
            slider.setValue(self.cfg[key])
            lbl.setText(str(self.cfg[key]))
        for w in widgets:
            w.blockSignals(False)

        self.update_color_btn()
        self.refresh_process_list()
        CrosshairRenderer.invalidate_cache()
        self._save_timer.stop()
        self.save_config()
        self.preview.update()
        self.overlay.update_position()
        self.overlay.check_target()

    def init_tray(self):
        pix = QPixmap(32, 32)
        pix.fill(Qt.GlobalColor.transparent)
        p = QPainter(pix)
        p.setPen(QPen(QColor("#00FF66"), 3))
        p.drawEllipse(6, 6, 20, 20)
        p.drawLine(16, 2, 16, 30)
        p.drawLine(2, 16, 30, 16)
        p.end()

        self.tray = QSystemTrayIcon(QIcon(pix), self)
        menu = QMenu()

        act_show = QAction("Открыть настройки", self)
        act_show.triggered.connect(self.showNormal)

        act_toggle = QAction("Вкл / Выкл прицел", self)
        act_toggle.triggered.connect(self.toggle_overlay)

        act_quit = QAction("Выход", self)
        act_quit.triggered.connect(self.quit_app)

        menu.addAction(act_show)
        menu.addAction(act_toggle)
        menu.addSeparator()
        menu.addAction(act_quit)

        self.tray.setContextMenu(menu)
        self.tray.activated.connect(lambda r: self.showNormal() if r == QSystemTrayIcon.ActivationReason.DoubleClick else None)
        self.tray.show()

    def toggle_overlay(self):
        self.cb_enabled.setChecked(not self.cb_enabled.isChecked())

    def closeEvent(self, event):
        event.ignore()
        self.hide()


if __name__ == "__main__":
    app = QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    win = SettingsWindow()
    win.show()
    sys.exit(app.exec())