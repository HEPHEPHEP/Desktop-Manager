"""
Desktop Folder Widget für Windows - Version 3.0 (PyQt)
======================================================
Tiefe Desktop-Integration:
- Widget wird in Desktop-Ebene (WorkerW) eingebettet
- Verschobene Verknüpfungen werden auf Desktop versteckt (Hidden-Attribut)
- Im Explorer bleiben sie sichtbar
- Kacheln rasten auf Desktop-Icon-Grid ein
- Icon-Ansicht wie auf dem Desktop

Autor: Claude
"""

import os
import json
import subprocess
import sys
from pathlib import Path
import ctypes
from ctypes import wintypes, windll
import atexit

from PyQt5 import QtCore, QtGui, QtWidgets
from PIL import Image
import win32gui
import win32ui
import win32con
import win32api
import pythoncom
from win32com.shell import shell, shellcon
from win32com.client import Dispatch

HAS_PIL = True
HAS_WIN32 = True
HAS_SHELL = True

# Windows DPI Awareness
try:
    windll.shcore.SetProcessDpiAwareness(1)
except Exception:
    pass

# ==========================================================================
# Windows API Definitionen
# ==========================================================================

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Typen für 64-Bit Kompatibilität
if ctypes.sizeof(ctypes.c_void_p) == 8:
    HWND = ctypes.c_uint64
    LONG_PTR = ctypes.c_int64
else:
    HWND = ctypes.c_uint32
    LONG_PTR = ctypes.c_int32

user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowW.restype = HWND

user32.FindWindowExW.argtypes = [HWND, HWND, ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowExW.restype = HWND

user32.SetParent.argtypes = [HWND, HWND]
user32.SetParent.restype = HWND

user32.GetParent.argtypes = [HWND]
user32.GetParent.restype = HWND

user32.GetWindowLongPtrW = user32.GetWindowLongPtrW if hasattr(user32, "GetWindowLongPtrW") else user32.GetWindowLongW
user32.GetWindowLongPtrW.argtypes = [HWND, ctypes.c_int]
user32.GetWindowLongPtrW.restype = LONG_PTR

user32.SetWindowLongPtrW = user32.SetWindowLongPtrW if hasattr(user32, "SetWindowLongPtrW") else user32.SetWindowLongW
user32.SetWindowLongPtrW.argtypes = [HWND, ctypes.c_int, LONG_PTR]
user32.SetWindowLongPtrW.restype = LONG_PTR

user32.SetWindowPos.argtypes = [HWND, HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
user32.SetWindowPos.restype = ctypes.c_bool

user32.SendMessageTimeoutW.argtypes = [HWND, ctypes.c_uint, ctypes.c_ulonglong, ctypes.c_longlong, ctypes.c_uint, ctypes.c_uint, ctypes.POINTER(ctypes.c_ulong)]
user32.SendMessageTimeoutW.restype = ctypes.c_long

kernel32.GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
kernel32.GetFileAttributesW.restype = ctypes.c_uint32

kernel32.SetFileAttributesW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
kernel32.SetFileAttributesW.restype = ctypes.c_bool

user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
user32.GetWindowThreadProcessId.restype = ctypes.c_uint32

class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]

user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(RECT)]
user32.GetWindowRect.restype = ctypes.c_bool

try:
    user32.SetWindowRgn.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_bool]
    user32.SetWindowRgn.restype = ctypes.c_int
except Exception:
    pass

WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, HWND, ctypes.c_void_p)
user32.EnumWindows.argtypes = [WNDENUMPROC, ctypes.c_void_p]
user32.EnumWindows.restype = ctypes.c_bool

GWL_EXSTYLE = -20
GWL_STYLE = -16
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000
WS_CHILD = 0x40000000
WS_POPUP = 0x80000000

SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
HWND_BOTTOM = HWND(1)

FILE_ATTRIBUTE_HIDDEN = 0x02

DESKTOP_GRID_X = 120
DESKTOP_GRID_Y = 120
DESKTOP_MARGIN_X = 20
DESKTOP_MARGIN_Y = 20

# ==========================================================================
# Windows Acrylic/Blur API für Glaseffekt
# ==========================================================================

class ACCENT_POLICY(ctypes.Structure):
    _fields_ = [
        ("AccentState", ctypes.c_int),
        ("AccentFlags", ctypes.c_int),
        ("GradientColor", ctypes.c_uint),
        ("AnimationId", ctypes.c_int),
    ]

class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ("Attribute", ctypes.c_int),
        ("Data", ctypes.POINTER(ACCENT_POLICY)),
        ("SizeOfData", ctypes.c_size_t),
    ]

ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
ACCENT_ENABLE_BLURBEHIND = 3
WCA_ACCENT_POLICY = 19


def enable_acrylic_blur(hwnd, gradient_color=0x80000000):
    try:
        accent = ACCENT_POLICY()
        accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2
        accent.GradientColor = gradient_color
        accent.AnimationId = 0

        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attribute = WCA_ACCENT_POLICY
        data.Data = ctypes.pointer(accent)
        data.SizeOfData = ctypes.sizeof(accent)

        result = ctypes.windll.user32.SetWindowCompositionAttribute(int(hwnd), ctypes.byref(data))
        return bool(result)
    except Exception as exc:
        print(f"    Acrylic Blur fehlgeschlagen: {exc}")
        try:
            accent = ACCENT_POLICY()
            accent.AccentState = ACCENT_ENABLE_BLURBEHIND
            accent.AccentFlags = 2
            accent.GradientColor = gradient_color
            accent.AnimationId = 0

            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attribute = WCA_ACCENT_POLICY
            data.Data = ctypes.pointer(accent)
            data.SizeOfData = ctypes.sizeof(accent)

            result = ctypes.windll.user32.SetWindowCompositionAttribute(int(hwnd), ctypes.byref(data))
            return bool(result)
        except Exception as exc2:
            print(f"    Blur-Fallback fehlgeschlagen: {exc2}")
            return False


def set_rounded_region(hwnd, width, height, radius=16):
    try:
        gdi32 = ctypes.windll.gdi32
        rgn = gdi32.CreateRoundRectRgn(0, 0, width + 1, height + 1, radius, radius)
        if rgn:
            user32.SetWindowRgn(int(hwnd), rgn, True)
            return True
    except Exception as exc:
        print(f"    SetWindowRgn fehlgeschlagen: {exc}")
    return False


class WindowsDesktopAPI:
    """Windows API für Desktop-Integration"""

    _workerw = None
    _progman = None

    @classmethod
    def find_desktop_window(cls):
        if cls._workerw:
            return cls._workerw

        try:
            cls._progman = user32.FindWindowW("Progman", None)
            if not cls._progman:
                return None

            result = ctypes.c_ulong()
            user32.SendMessageTimeoutW(
                HWND(cls._progman),
                0x052C,
                0,
                0,
                0x0000,
                1000,
                ctypes.byref(result),
            )

            workerw_list = []

            def enum_callback(hwnd, lparam):
                shell_view = user32.FindWindowExW(hwnd, HWND(0), "SHELLDLL_DefView", None)
                if shell_view:
                    next_worker = user32.FindWindowExW(HWND(0), hwnd, "WorkerW", None)
                    if next_worker:
                        workerw_list.append(next_worker)
                return True

            enum_func = WNDENUMPROC(enum_callback)
            user32.EnumWindows(enum_func, None)

            cls._workerw = workerw_list[0] if workerw_list else cls._progman
            return cls._workerw
        except Exception as exc:
            print(f"Fehler beim Finden des Desktop-Fensters: {exc}")
            return None

    @staticmethod
    def set_parent_to_desktop(hwnd):
        desktop_hwnd = WindowsDesktopAPI.find_desktop_window()
        if desktop_hwnd:
            try:
                hwnd = HWND(hwnd) if not isinstance(hwnd, HWND) else hwnd
                desktop_hwnd = HWND(desktop_hwnd) if not isinstance(desktop_hwnd, HWND) else desktop_hwnd

                user32.SetParent(hwnd, desktop_hwnd)

                style = user32.GetWindowLongPtrW(hwnd, GWL_STYLE)
                style = (style & ~WS_POPUP) | WS_CHILD
                user32.SetWindowLongPtrW(hwnd, GWL_STYLE, LONG_PTR(style))

                ex_style = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
                ex_style = ex_style | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, LONG_PTR(ex_style))

                return True
            except Exception as exc:
                print(f"Fehler beim Setzen des Parents: {exc}")
        return False

    @staticmethod
    def set_file_hidden(filepath, hidden=True):
        try:
            attrs = kernel32.GetFileAttributesW(filepath)
            if attrs == 0xFFFFFFFF:
                return False

            if hidden:
                new_attrs = attrs | FILE_ATTRIBUTE_HIDDEN
            else:
                new_attrs = attrs & ~FILE_ATTRIBUTE_HIDDEN

            result = kernel32.SetFileAttributesW(filepath, new_attrs)
            return bool(result)
        except Exception as exc:
            print(f"    Fehler in set_file_hidden: {exc}")
            return False

    @staticmethod
    def refresh_desktop():
        try:
            SHCNE_ASSOCCHANGED = 0x08000000
            SHCNF_IDLIST = 0x0000
            SHCNE_UPDATEDIR = 0x00001000
            SHCNF_PATH = 0x0005

            ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)

            desktop_path = WindowsDesktopAPI.get_desktop_path()
            if desktop_path:
                ctypes.windll.shell32.SHChangeNotify(
                    SHCNE_UPDATEDIR,
                    SHCNF_PATH,
                    desktop_path.encode("utf-16-le") + b"\x00\x00",
                    None,
                )
        except Exception as exc:
            print(f"    Warnung bei refresh_desktop: {exc}")

    @staticmethod
    def get_desktop_path():
        try:
            if HAS_SHELL:
                return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
        except Exception:
            pass
        return str(Path.home() / "Desktop")

    @staticmethod
    def snap_to_grid(x, y):
        grid_x = round((x - DESKTOP_MARGIN_X) / DESKTOP_GRID_X) * DESKTOP_GRID_X + DESKTOP_MARGIN_X
        grid_y = round((y - DESKTOP_MARGIN_Y) / DESKTOP_GRID_Y) * DESKTOP_GRID_Y + DESKTOP_MARGIN_Y
        return max(DESKTOP_MARGIN_X, grid_x), max(DESKTOP_MARGIN_Y, grid_y)

    @staticmethod
    def set_window_bottom(hwnd):
        try:
            hwnd = HWND(hwnd) if not isinstance(hwnd, HWND) else hwnd
            user32.SetWindowPos(
                hwnd, HWND_BOTTOM, 0, 0, 0, 0, SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            )
        except Exception:
            pass


class IconExtractor:
    """Extrahiert echte Windows-Icons aus Dateien"""

    ICON_CACHE = {}

    @staticmethod
    def get_icon(filepath, size=48):
        cache_key = f"{filepath}_{size}"
        if cache_key in IconExtractor.ICON_CACHE:
            return IconExtractor.ICON_CACHE[cache_key]

        img = IconExtractor.extract_windows_icon(filepath, size)
        if not img:
            img = IconExtractor.get_default_icon(filepath, size)

        if img:
            IconExtractor.ICON_CACHE[cache_key] = img
        return img

    @staticmethod
    def extract_windows_icon(filepath, size=48):
        if not HAS_WIN32 or not HAS_PIL:
            return None
        try:
            class SHFILEINFOW(ctypes.Structure):
                _fields_ = [
                    ("hIcon", ctypes.c_void_p),
                    ("iIcon", ctypes.c_int),
                    ("dwAttributes", ctypes.c_uint),
                    ("szDisplayName", ctypes.c_wchar * 260),
                    ("szTypeName", ctypes.c_wchar * 80),
                ]

            info = SHFILEINFOW()
            SHGFI_ICON = 0x100
            SHGFI_LARGEICON = 0x0

            result = ctypes.windll.shell32.SHGetFileInfoW(
                filepath, 0, ctypes.byref(info), ctypes.sizeof(info), SHGFI_ICON | SHGFI_LARGEICON
            )

            if not result or not info.hIcon:
                return None

            hicon = info.hIcon

            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)

            hdc_screen = win32gui.GetDC(0)
            hdc = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem = hdc.CreateCompatibleDC()

            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
            hdc_mem.SelectObject(hbmp)

            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0)
            hdc_mem.DrawIcon((0, 0), hicon)

            bmpstr = hbmp.GetBitmapBits(True)
            img_black = Image.frombuffer("RGB", (ico_x, ico_y), bmpstr, "raw", "BGRX", 0, 1)

            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0x00FFFFFF)
            hdc_mem.DrawIcon((0, 0), hicon)
            bmpstr2 = hbmp.GetBitmapBits(True)
            img_white = Image.frombuffer("RGB", (ico_x, ico_y), bmpstr2, "raw", "BGRX", 0, 1)

            import numpy as np

            black_arr = np.array(img_black, dtype=np.float32)
            white_arr = np.array(img_white, dtype=np.float32)

            diff = white_arr - black_arr
            alpha = 255.0 - np.mean(diff, axis=2)
            alpha = np.clip(alpha, 0, 255).astype(np.uint8)

            rgba = np.zeros((ico_y, ico_x, 4), dtype=np.uint8)
            mask = alpha > 0
            for c in range(3):
                rgba[:, :, c] = np.where(
                    mask,
                    np.clip(black_arr[:, :, c] * 255.0 / np.maximum(alpha, 1), 0, 255),
                    0,
                ).astype(np.uint8)
            rgba[:, :, 3] = alpha

            img = Image.fromarray(rgba, "RGBA")

            if ico_x != size or ico_y != size:
                resample = Image.Resampling.LANCZOS if hasattr(Image, "Resampling") else Image.LANCZOS
                img = img.resize((size, size), resample)

            win32gui.DestroyIcon(hicon)
            hdc_mem.DeleteDC()
            win32gui.ReleaseDC(0, hdc_screen)

            return img
        except Exception:
            return None

    @staticmethod
    def get_default_icon(filepath, size=48):
        if not HAS_PIL:
            return None
        try:
            from PIL import ImageDraw, ImageFont, ImageFilter
        except Exception:
            return None

        scale = 2
        s = size * scale

        img = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)

        ext = Path(filepath).suffix.lower() if filepath else ""
        name = Path(filepath).stem if filepath else ""

        type_colors = {
            ".exe": (0, 120, 212),
            ".msi": (0, 120, 212),
            ".lnk": (0, 120, 212),
            ".bat": (255, 165, 0),
            ".cmd": (255, 165, 0),
            ".py": (55, 118, 171),
            ".txt": (107, 107, 107),
            ".pdf": (220, 30, 30),
        }

        cr, cg, cb = type_colors.get(ext, (0, 120, 212))

        margin = 4 * scale
        radius = 8 * scale

        shadow = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        s_draw = ImageDraw.Draw(shadow)
        s_draw.rounded_rectangle(
            [margin + 3 * scale, margin + 3 * scale, s - margin + 1 * scale, s - margin + 1 * scale],
            radius=radius,
            fill=(0, 0, 0, 70),
        )
        shadow = shadow.filter(ImageFilter.GaussianBlur(radius=3 * scale))
        img = Image.alpha_composite(img, shadow)

        main = Image.new("RGBA", (s, s), (0, 0, 0, 0))
        m_draw = ImageDraw.Draw(main)
        m_draw.rounded_rectangle(
            [margin, margin, s - margin, s - margin],
            radius=radius,
            fill=(cr, cg, cb, 240),
        )
        img = Image.alpha_composite(img, main)

        letter = name[0].upper() if name else "?"

        font = None
        font_size = s // 2
        try:
            font = ImageFont.truetype("segoeui.ttf", font_size)
        except Exception:
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except Exception:
                font = ImageFont.load_default()

        if font:
            txt_layer = Image.new("RGBA", (s, s), (0, 0, 0, 0))
            txt_draw = ImageDraw.Draw(txt_layer)
            bbox = txt_draw.textbbox((0, 0), letter, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]

            x = (s - text_width) // 2
            y = (s - text_height) // 2 - 2 * scale
            txt_draw.text((x, y), letter, fill=(255, 255, 255, 230), font=font)
            img = Image.alpha_composite(img, txt_layer)

        if hasattr(Image, "Resampling"):
            img = img.resize((size, size), Image.Resampling.LANCZOS)
        else:
            img = img.resize((size, size), Image.LANCZOS)

        return img


class FolderTile(QtWidgets.QWidget):
    """Eine einzelne Ordner-Kachel auf dem Desktop (PyQt)."""

    def __init__(self, manager, tile_id, config):
        super().__init__()
        self.manager = manager
        self.tile_id = tile_id
        self.config = config
        self.is_expanded = False
        self.drag_data = {"start": None, "moving": False}

        self.collapsed_tile_w = self.config.get("collapsed_tile_w", 150)
        self.collapsed_tile_h = self.config.get("collapsed_tile_h", 150)
        self.expanded_tile_w = self.config.get("expanded_tile_w", 320)
        self.expanded_tile_h = self.config.get("expanded_tile_h", 320)
        self.collapsed_icon_w = self.config.get("collapsed_icon_w", 48)
        self.collapsed_icon_h = self.config.get("collapsed_icon_h", 48)
        self.expanded_icon_w = self.config.get("expanded_icon_w", 48)
        self.expanded_icon_h = self.config.get("expanded_icon_h", 48)

        self.setWindowFlags(
            QtCore.Qt.Tool | QtCore.Qt.FramelessWindowHint | QtCore.Qt.WindowStaysOnBottomHint
        )
        self.setAttribute(QtCore.Qt.WA_TranslucentBackground, True)
        self.setAcceptDrops(True)

        self._build_ui()
        self._apply_geometry()
        self._embed_into_desktop()

    def _build_ui(self):
        self.setStyleSheet("background-color: #0d0d1a; color: #d0d0e0;")
        self.main_layout = QtWidgets.QVBoxLayout(self)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(6)

        header_layout = QtWidgets.QHBoxLayout()
        header_layout.setContentsMargins(0, 0, 0, 0)
        self.title_label = QtWidgets.QLabel(self.config.get("name", "Ordner"))
        self.title_label.setAlignment(QtCore.Qt.AlignCenter)
        header_layout.addWidget(self.title_label, 1)
        self.toggle_button = QtWidgets.QToolButton()
        self.toggle_button.setText("⤢")
        self.toggle_button.clicked.connect(self.toggle_expand)
        self.toggle_button.setCursor(QtCore.Qt.PointingHandCursor)
        header_layout.addWidget(self.toggle_button)
        self.main_layout.addLayout(header_layout)

        self.collapsed_widget = QtWidgets.QWidget()
        self.collapsed_layout = QtWidgets.QGridLayout(self.collapsed_widget)
        self.collapsed_layout.setContentsMargins(0, 0, 0, 0)
        self.collapsed_layout.setSpacing(6)

        self.expanded_widget = QtWidgets.QWidget()
        self.expanded_layout = QtWidgets.QVBoxLayout(self.expanded_widget)
        self.expanded_layout.setContentsMargins(0, 0, 0, 0)
        self.expanded_layout.setSpacing(6)
        self.scroll_area = QtWidgets.QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setFrameShape(QtWidgets.QFrame.NoFrame)
        self.expanded_grid_holder = QtWidgets.QWidget()
        self.expanded_grid = QtWidgets.QGridLayout(self.expanded_grid_holder)
        self.expanded_grid.setContentsMargins(0, 0, 0, 0)
        self.expanded_grid.setSpacing(8)
        self.scroll_area.setWidget(self.expanded_grid_holder)
        self.expanded_layout.addWidget(self.scroll_area)

        self.main_layout.addWidget(self.collapsed_widget)
        self.main_layout.addWidget(self.expanded_widget)

        self.expanded_widget.hide()
        self.refresh_views()

    def _apply_geometry(self):
        x = self.config.get("pos_x", DESKTOP_MARGIN_X + int(self.tile_id) * self.collapsed_tile_w)
        y = self.config.get("pos_y", DESKTOP_MARGIN_Y)
        self.setGeometry(x, y, self.collapsed_tile_w, self.collapsed_tile_h)
        self._apply_rounding(self.width(), self.height())

    def _apply_rounding(self, width, height):
        if self.winId():
            set_rounded_region(int(self.winId()), width, height, radius=14)

    def _embed_into_desktop(self):
        try:
            hwnd = int(self.winId())
            WindowsDesktopAPI.set_parent_to_desktop(hwnd)
            WindowsDesktopAPI.set_window_bottom(hwnd)
            enable_acrylic_blur(hwnd, 0xA0101010)
        except Exception as exc:
            print(f"Fehler bei Desktop-Einbettung: {exc}")

    def refresh_views(self):
        self.title_label.setText(self.config.get("name", "Ordner"))
        self._populate_collapsed()
        self._populate_expanded()

    def _populate_collapsed(self):
        while self.collapsed_layout.count():
            item = self.collapsed_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        shortcuts = self.config.get("shortcuts", [])
        for idx, shortcut in enumerate(shortcuts[:4]):
            row = idx // 2
            col = idx % 2
            button = self._create_shortcut_button(shortcut, idx, is_expanded=False)
            self.collapsed_layout.addWidget(button, row, col)

        if not shortcuts:
            label = QtWidgets.QLabel("Leer\nDateien hierher ziehen")
            label.setAlignment(QtCore.Qt.AlignCenter)
            label.setStyleSheet("color: #555570;")
            self.collapsed_layout.addWidget(label, 0, 0, 1, 2)

    def _populate_expanded(self):
        while self.expanded_grid.count():
            item = self.expanded_grid.takeAt(0)
            if item.widget():
                item.widget().deleteLater()

        shortcuts = self.config.get("shortcuts", [])
        if not shortcuts:
            label = QtWidgets.QLabel("Leer\nDateien vom Desktop hierher ziehen")
            label.setAlignment(QtCore.Qt.AlignCenter)
            label.setStyleSheet("color: #555570;")
            self.expanded_grid.addWidget(label, 0, 0)
            return

        columns = 3
        for idx, shortcut in enumerate(shortcuts):
            row = idx // columns
            col = idx % columns
            button = self._create_shortcut_button(shortcut, idx, is_expanded=True)
            self.expanded_grid.addWidget(button, row, col)

    def _create_shortcut_button(self, shortcut, index, is_expanded):
        button = QtWidgets.QToolButton()
        button.setToolButtonStyle(QtCore.Qt.ToolButtonTextUnderIcon)
        button.setText(self._short_name(shortcut["name"], 10))
        button.setIcon(self._get_qt_icon(shortcut["path"], is_expanded))
        icon_size = self.expanded_icon_w if is_expanded else self.collapsed_icon_w
        button.setIconSize(QtCore.QSize(icon_size, icon_size))
        button.setCursor(QtCore.Qt.PointingHandCursor)
        button.clicked.connect(lambda checked=False, path=shortcut["path"]: self.launch_shortcut(path))
        button.setContextMenuPolicy(QtCore.Qt.CustomContextMenu)
        button.customContextMenuRequested.connect(
            lambda pos, idx=index, path=shortcut["path"]: self.show_item_context_menu(button, pos, idx, path)
        )
        return button

    def _get_qt_icon(self, filepath, is_expanded):
        size = self.expanded_icon_w if is_expanded else self.collapsed_icon_w
        pil_img = IconExtractor.get_icon(filepath, max(16, size))
        if pil_img:
            if hasattr(Image, "Resampling"):
                pil_img = pil_img.resize((size, size), Image.Resampling.LANCZOS)
            else:
                pil_img = pil_img.resize((size, size), Image.LANCZOS)
            from PIL.ImageQt import ImageQt

            qimage = ImageQt(pil_img)
            return QtGui.QIcon(QtGui.QPixmap.fromImage(qimage))

        pixmap = QtGui.QPixmap(size, size)
        pixmap.fill(QtGui.QColor("#0078D4"))
        return QtGui.QIcon(pixmap)

    def _short_name(self, name, limit):
        if len(name) > limit:
            return name[: limit - 1] + "…"
        return name

    def toggle_expand(self):
        if self.is_expanded:
            self.collapse()
        else:
            self.expand()

    def expand(self):
        if self.is_expanded:
            return
        self.is_expanded = True
        self.expanded_widget.show()
        self.collapsed_widget.hide()
        self.toggle_button.setText("⤡")
        self.resize(self.expanded_tile_w, self.expanded_tile_h)
        self._apply_rounding(self.width(), self.height())
        self.refresh_views()

    def collapse(self):
        if not self.is_expanded:
            return
        self.is_expanded = False
        self.expanded_widget.hide()
        self.collapsed_widget.show()
        self.toggle_button.setText("⤢")
        self.resize(self.collapsed_tile_w, self.collapsed_tile_h)
        self._apply_rounding(self.width(), self.height())
        self.refresh_views()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()

    def dropEvent(self, event):
        urls = [url.toLocalFile() for url in event.mimeData().urls()]
        self.on_drop_files(urls)

    def on_drop_files(self, filepaths):
        desktop_path = WindowsDesktopAPI.get_desktop_path()
        added = 0
        for filepath in filepaths:
            if not filepath or not os.path.exists(filepath):
                continue
            if any(s["path"] == filepath for s in self.config.get("shortcuts", [])):
                continue
            if "shortcuts" not in self.config:
                self.config["shortcuts"] = []
            self.config["shortcuts"].append({"name": Path(filepath).stem, "path": filepath})
            if Path(filepath).parent == Path(desktop_path):
                WindowsDesktopAPI.set_file_hidden(filepath, True)
            added += 1

        if added:
            WindowsDesktopAPI.refresh_desktop()
            self.manager.save_config()
            self.refresh_views()

    def mousePressEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            self.drag_data["start"] = event.globalPos()
            self.drag_data["moving"] = False
        super().mousePressEvent(event)

    def mouseMoveEvent(self, event):
        if self.drag_data["start"] is None:
            return
        diff = event.globalPos() - self.drag_data["start"]
        if diff.manhattanLength() > 5:
            self.drag_data["moving"] = True
            new_pos = self.pos() + diff
            self.move(new_pos)
            self.drag_data["start"] = event.globalPos()
        super().mouseMoveEvent(event)

    def mouseReleaseEvent(self, event):
        if event.button() == QtCore.Qt.LeftButton:
            if self.drag_data["moving"]:
                snap_x, snap_y = WindowsDesktopAPI.snap_to_grid(self.x(), self.y())
                self.move(snap_x, snap_y)
                self.config["pos_x"] = snap_x
                self.config["pos_y"] = snap_y
                self.manager.save_config()
            self.drag_data["start"] = None
            self.drag_data["moving"] = False
        super().mouseReleaseEvent(event)

    def show_item_context_menu(self, widget, pos, index, path):
        menu = QtWidgets.QMenu(self)
        menu.addAction("▶️ Öffnen", lambda: self.launch_shortcut(path))
        menu.addAction("📤 Auf Desktop wiederherstellen", lambda: self.restore_to_desktop(index))
        menu.addSeparator()
        menu.addAction("🗑️ Aus Ordner entfernen", lambda: self.remove_shortcut(index))
        menu.exec_(widget.mapToGlobal(pos))

    def contextMenuEvent(self, event):
        menu = QtWidgets.QMenu(self)
        menu.addAction("➕ Neue Kachel", self.manager.create_new_tile)
        menu.addAction("✏️ Umbenennen", self.rename)
        menu.addSeparator()
        menu.addAction("🗑️ Kachel löschen", self.delete_tile)
        menu.addAction("❌ Beenden", self.manager.quit)
        menu.exec_(event.globalPos())

    def rename(self):
        new_name, ok = QtWidgets.QInputDialog.getText(
            self, "Umbenennen", "Neuer Name:", text=self.config.get("name", "Ordner")
        )
        if ok and new_name.strip():
            self.config["name"] = new_name.strip()
            self.manager.save_config()
            self.refresh_views()

    def delete_tile(self):
        if (
            QtWidgets.QMessageBox.question(
                self, "Löschen", "Kachel löschen?\n\nVerknüpfungen werden auf dem Desktop wiederhergestellt."
            )
            == QtWidgets.QMessageBox.Yes
        ):
            self.restore_all_to_desktop()
            self.manager.delete_tile(self.tile_id)

    def restore_all_to_desktop(self):
        for shortcut in list(self.config.get("shortcuts", [])):
            filepath = shortcut.get("path", "")
            if filepath and os.path.exists(filepath):
                WindowsDesktopAPI.set_file_hidden(filepath, False)
        self.config["shortcuts"] = []
        WindowsDesktopAPI.refresh_desktop()
        self.manager.save_config()

    def restore_to_desktop(self, index):
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            filepath = shortcuts[index]["path"]
            WindowsDesktopAPI.set_file_hidden(filepath, False)
            WindowsDesktopAPI.refresh_desktop()
            del self.config["shortcuts"][index]
            self.manager.save_config()
            self.refresh_views()

    def remove_shortcut(self, index):
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            del self.config["shortcuts"][index]
            self.manager.save_config()
            self.refresh_views()

    def launch_shortcut(self, path):
        try:
            if path.lower().endswith(".lnk"):
                os.startfile(path)
            else:
                subprocess.Popen(path, shell=True)
        except Exception as exc:
            QtWidgets.QMessageBox.critical(self, "Fehler", f"Programm konnte nicht gestartet werden:\n{exc}")


class DesktopFolderManager:
    """Verwaltet alle Ordner-Kacheln (PyQt)."""

    CONFIG_FILE = Path.home() / ".desktop_folder_widget_v3.json"

    def __init__(self):
        self.app = QtWidgets.QApplication.instance() or QtWidgets.QApplication(sys.argv)
        self.tiles = {}
        self.config = self.load_config()

        if not self.config.get("tiles"):
            self.config["tiles"] = {
                "0": {
                    "name": "Apps",
                    "shortcuts": [],
                    "pos_x": DESKTOP_MARGIN_X,
                    "pos_y": DESKTOP_MARGIN_Y,
                }
            }
            self.save_config()

        for tile_id, tile_config in self.config["tiles"].items():
            tile = FolderTile(self, tile_id, tile_config)
            tile.show()
            self.tiles[tile_id] = tile

        self.check_dependencies()
        self.app.aboutToQuit.connect(self.quit)

    def check_dependencies(self):
        missing = []
        if not HAS_WIN32:
            missing.append("pywin32 + Pillow")
        if not HAS_SHELL:
            missing.append("pywin32 (Shell)")
        if missing:
            print(f"\n⚠️ Fehlende Abhängigkeiten: {', '.join(missing)}")
            print("   pip install pywin32 Pillow\n")

    def load_config(self):
        if self.CONFIG_FILE.exists():
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except Exception:
                pass
        return {"tiles": {}}

    def save_config(self):
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as exc:
            print(f"Speicherfehler: {exc}")

    def create_new_tile(self):
        existing = [int(i) for i in self.config["tiles"].keys()]
        new_id = str(max(existing) + 1 if existing else 0)

        last = list(self.tiles.values())[-1] if self.tiles else None
        if last:
            x = last.x() + DESKTOP_GRID_X
            y = last.y()
        else:
            x, y = DESKTOP_MARGIN_X, DESKTOP_MARGIN_Y
        x, y = WindowsDesktopAPI.snap_to_grid(x, y)

        tile_config = {
            "name": f"Ordner {new_id}",
            "shortcuts": [],
            "pos_x": x,
            "pos_y": y,
        }
        self.config["tiles"][new_id] = tile_config
        tile = FolderTile(self, new_id, tile_config)
        tile.show()
        self.tiles[new_id] = tile
        self.save_config()

    def delete_tile(self, tile_id):
        if tile_id in self.tiles:
            self.tiles[tile_id].close()
            del self.tiles[tile_id]
        if tile_id in self.config["tiles"]:
            del self.config["tiles"][tile_id]
        self.save_config()
        if not self.tiles:
            self.quit()

    def quit(self):
        print("\n" + "=" * 50)
        print("  Widget wird beendet...")
        print("=" * 50)

        restored_count = 0
        for tile_config in self.config.get("tiles", {}).values():
            for shortcut in tile_config.get("shortcuts", []):
                filepath = shortcut.get("path", "")
                name = shortcut.get("name", "Unbekannt")
                if filepath and os.path.exists(filepath):
                    if WindowsDesktopAPI.set_file_hidden(filepath, False):
                        restored_count += 1
                        print(f"  ✓ Wiederhergestellt: {name}")
                    else:
                        print(f"  ✗ Fehler bei: {name}")
                else:
                    print(f"  ? Datei nicht gefunden: {name}")

        WindowsDesktopAPI.refresh_desktop()
        print(f"\n{restored_count} Icons wiederhergestellt.")
        print("=" * 50)

        self.save_config()
        for tile in list(self.tiles.values()):
            tile.close()
        if QtWidgets.QApplication.instance():
            QtWidgets.QApplication.instance().quit()

    def run(self):
        return self.app.exec_()


_app_instance = None
_cleanup_done = False


def cleanup_on_exit():
    global _app_instance, _cleanup_done
    if _cleanup_done:
        return
    _cleanup_done = True

    if _app_instance and hasattr(_app_instance, "config"):
        print("\n[Cleanup] Stelle Desktop-Icons wieder her...")
        try:
            count = 0
            for tile_config in _app_instance.config.get("tiles", {}).values():
                for shortcut in tile_config.get("shortcuts", []):
                    filepath = shortcut.get("path", "")
                    if filepath and os.path.exists(filepath):
                        WindowsDesktopAPI.set_file_hidden(filepath, False)
                        count += 1
            WindowsDesktopAPI.refresh_desktop()
            print(f"[Cleanup] {count} Icons wiederhergestellt.")
        except Exception as exc:
            print(f"[Cleanup] Fehler: {exc}")


def main():
    global _app_instance

    print("=" * 55)
    print("  Desktop Folder Widget v3.0 (PyQt)")
    print("=" * 55)
    print()
    print("Bedienung:")
    print("  • Linksklick + ziehen → Kachel verschieben")
    print("  • Toggle-Button       → Kachel oeffnen/schließen")
    print("  • Rechtsklick         → Kontextmenue")
    print("  • Dateien hinziehen   → In Kachel verschieben")
    print()
    print("WICHTIG: Beim Beenden werden alle Icons wiederhergestellt!")
    print("-" * 55)
    print()

    atexit.register(cleanup_on_exit)

    try:
        _app_instance = DesktopFolderManager()
        _app_instance.run()
    except KeyboardInterrupt:
        print("\n[Beendet durch Benutzer]")
    except Exception as exc:
        print(f"\n[Fehler] {exc}")
    finally:
        cleanup_on_exit()


if __name__ == "__main__":
    main()
