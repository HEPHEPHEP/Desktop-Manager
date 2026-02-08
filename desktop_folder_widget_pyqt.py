"""
Desktop Folder Widget für Windows - PyQt6 Version mit Frosted Glass
====================================================================
Tiefe Desktop-Integration:
- Widget wird in Desktop-Ebene (WorkerW) eingebettet
- Verschobene Verknüpfungen werden auf Desktop versteckt (Hidden-Attribut)
- Im Explorer bleiben sie sichtbar
- Kacheln rasten auf Desktop-Icon-Grid ein
- Echter Frosted Glass / Acrylic Blur Effekt

Autor: Claude
Framework: PyQt6 (statt tkinter für bessere Blur-Unterstützung)
"""

import sys
import os
import json
import subprocess
from pathlib import Path
import ctypes
from ctypes import wintypes
import atexit

from PyQt6.QtWidgets import (
    QApplication, QWidget, QLabel, QFrame, QVBoxLayout, QHBoxLayout,
    QGridLayout, QScrollArea, QMenu, QInputDialog, QMessageBox,
    QGraphicsDropShadowEffect, QSizePolicy, QLineEdit, QPushButton,
    QFileIconProvider, QSpinBox, QDialog, QDialogButtonBox, QGroupBox,
    QSlider
)
from PyQt6.QtCore import (
    Qt, QPoint, QSize, QPropertyAnimation, QEasingCurve, QTimer,
    QRect, pyqtSignal, QMimeData, QUrl, QEvent, QFileInfo
)
from PyQt6.QtGui import (
    QPixmap, QImage, QPainter, QColor, QFont, QIcon, QCursor,
    QPainterPath, QBrush, QPen, QLinearGradient, QRadialGradient,
    QDrag, QRegion
)

# Für Icon-Extraktion
try:
    from PIL import Image, ImageDraw, ImageFilter, ImageFont
    import win32gui
    import win32ui
    import win32con
    import win32api
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False
    print("⚠️ pywin32 oder Pillow nicht installiert. Icons werden als Fallback gerendert.")

# Für Shell-Operationen
try:
    import pythoncom
    from win32com.shell import shell, shellcon
    HAS_SHELL = True
except ImportError:
    HAS_SHELL = False

# Für Blur-Effekt
try:
    from BlurWindow.blurWindow import GlobalBlur
    HAS_BLURWINDOW = True
except ImportError:
    HAS_BLURWINDOW = False
    print("⚠️ BlurWindow nicht installiert. pip install BlurWindow")

# DPI Awareness wird von Qt6 automatisch gesetzt
# (DPI_AWARENESS_CONTEXT_PER_MONITOR_AWARE_V2)

# ============================================================================
# Windows API Definitionen
# ============================================================================

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32
dwmapi = ctypes.windll.dwmapi

# Typen für 64-Bit Kompatibilität
if ctypes.sizeof(ctypes.c_void_p) == 8:
    HWND = ctypes.c_uint64
    LONG_PTR = ctypes.c_int64
else:
    HWND = ctypes.c_uint32
    LONG_PTR = ctypes.c_int32

# Funktions-Signaturen
user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowW.restype = HWND

user32.FindWindowExW.argtypes = [HWND, HWND, ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowExW.restype = HWND

user32.SetParent.argtypes = [HWND, HWND]
user32.SetParent.restype = HWND

user32.GetParent.argtypes = [HWND]
user32.GetParent.restype = HWND

user32.SetWindowLongPtrW = getattr(user32, 'SetWindowLongPtrW', user32.SetWindowLongW)
user32.GetWindowLongPtrW = getattr(user32, 'GetWindowLongPtrW', user32.GetWindowLongW)

user32.SetWindowPos.argtypes = [HWND, HWND, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_int, ctypes.c_uint]
user32.SetWindowPos.restype = ctypes.c_bool

kernel32.GetFileAttributesW.argtypes = [ctypes.c_wchar_p]
kernel32.GetFileAttributesW.restype = ctypes.c_uint32

kernel32.SetFileAttributesW.argtypes = [ctypes.c_wchar_p, ctypes.c_uint32]
kernel32.SetFileAttributesW.restype = ctypes.c_bool

# EnumWindows callback
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, HWND, ctypes.c_void_p)
user32.EnumWindows.argtypes = [WNDENUMPROC, ctypes.c_void_p]

# Window Styles
GWL_EXSTYLE = -20
GWL_STYLE = -16
WS_EX_TOOLWINDOW = 0x00000080
WS_EX_NOACTIVATE = 0x08000000
WS_EX_LAYERED = 0x00080000
WS_CHILD = 0x40000000
WS_POPUP = 0x80000000

# SetWindowPos Flags
SWP_NOMOVE = 0x0002
SWP_NOSIZE = 0x0001
SWP_NOACTIVATE = 0x0010
SWP_SHOWWINDOW = 0x0040
HWND_BOTTOM = 1

# File Attributes
FILE_ATTRIBUTE_HIDDEN = 0x02

# Desktop Grid
DESKTOP_GRID_X = 75
DESKTOP_GRID_Y = 75
DESKTOP_MARGIN_X = 0
DESKTOP_MARGIN_Y = 0

# ============================================================================
# Windows DWM API für Acrylic/Mica Effekt
# ============================================================================

class ACCENT_POLICY(ctypes.Structure):
    _fields_ = [
        ('AccentState', ctypes.c_int),
        ('AccentFlags', ctypes.c_int),
        ('GradientColor', ctypes.c_uint),
        ('AnimationId', ctypes.c_int),
    ]

class WINDOWCOMPOSITIONATTRIBDATA(ctypes.Structure):
    _fields_ = [
        ('Attribute', ctypes.c_int),
        ('Data', ctypes.POINTER(ACCENT_POLICY)),
        ('SizeOfData', ctypes.c_size_t),
    ]

class MARGINS(ctypes.Structure):
    _fields_ = [
        ('cxLeftWidth', ctypes.c_int),
        ('cxRightWidth', ctypes.c_int),
        ('cyTopHeight', ctypes.c_int),
        ('cyBottomHeight', ctypes.c_int),
    ]

# DWM Attribute für Windows 11
DWMWA_USE_IMMERSIVE_DARK_MODE = 20
DWMWA_MICA_EFFECT = 1029
DWMWA_SYSTEMBACKDROP_TYPE = 38

# Backdrop Types für Windows 11
DWMSBT_MAINWINDOW = 2  # Mica
DWMSBT_TRANSIENTWINDOW = 3  # Acrylic
DWMSBT_TABBEDWINDOW = 4  # Tabbed

# Accent States
ACCENT_ENABLE_BLURBEHIND = 3
ACCENT_ENABLE_ACRYLICBLURBEHIND = 4
WCA_ACCENT_POLICY = 19


def enable_blur_behind(hwnd, enable=True):
    """Aktiviert Windows Blur Behind für das Fenster"""
    try:
        # Margins auf -1 setzen für Vollbildblur
        margins = MARGINS(-1, -1, -1, -1)
        dwmapi.DwmExtendFrameIntoClientArea(hwnd, ctypes.byref(margins))
        return True
    except Exception as e:
        print(f"DwmExtendFrameIntoClientArea fehlgeschlagen: {e}")
        return False


def enable_acrylic_blur(hwnd, gradient_color=0x80000000):
    """
    Aktiviert Acrylic Blur (Glaseffekt) für ein Fenster.
    gradient_color: AABBGGRR Format
    """
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

        result = user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data))
        return bool(result)
    except Exception as e:
        print(f"Acrylic Blur fehlgeschlagen: {e}")
        # Fallback: Standard-Blur
        try:
            accent = ACCENT_POLICY()
            accent.AccentState = ACCENT_ENABLE_BLURBEHIND
            accent.AccentFlags = 2
            accent.GradientColor = gradient_color
            
            data = WINDOWCOMPOSITIONATTRIBDATA()
            data.Attribute = WCA_ACCENT_POLICY
            data.Data = ctypes.pointer(accent)
            data.SizeOfData = ctypes.sizeof(accent)
            
            return bool(user32.SetWindowCompositionAttribute(hwnd, ctypes.byref(data)))
        except:
            return False


def enable_mica_effect(hwnd):
    """Aktiviert Windows 11 Mica Effekt"""
    try:
        # Dark Mode aktivieren
        value = ctypes.c_int(1)
        dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_USE_IMMERSIVE_DARK_MODE,
            ctypes.byref(value), ctypes.sizeof(value)
        )
        
        # Mica Backdrop
        backdrop = ctypes.c_int(DWMSBT_MAINWINDOW)
        result = dwmapi.DwmSetWindowAttribute(
            hwnd, DWMWA_SYSTEMBACKDROP_TYPE,
            ctypes.byref(backdrop), ctypes.sizeof(backdrop)
        )
        return result == 0
    except Exception as e:
        print(f"Mica Effekt fehlgeschlagen: {e}")
        return False


# ============================================================================
# Windows Desktop API
# ============================================================================

class WindowsDesktopAPI:
    """Windows API für Desktop-Integration"""
    
    _workerw = None
    _progman = None
    
    @classmethod
    def find_desktop_window(cls):
        """Findet das Desktop-Fenster (WorkerW hinter den Icons)"""
        if cls._workerw:
            return cls._workerw
        
        try:
            cls._progman = user32.FindWindowW("Progman", None)
            if not cls._progman:
                return None
            
            result = ctypes.c_ulong()
            user32.SendMessageTimeoutW(
                HWND(cls._progman), 0x052C, 0, 0, 0x0000, 1000, ctypes.byref(result)
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
            
        except Exception as e:
            print(f"Fehler beim Finden des Desktop-Fensters: {e}")
            return None
    
    @staticmethod
    def set_file_hidden(filepath, hidden=True):
        """Setzt oder entfernt das Hidden-Attribut einer Datei"""
        try:
            attrs = kernel32.GetFileAttributesW(filepath)
            if attrs == 0xFFFFFFFF:
                return False
            
            if hidden:
                new_attrs = attrs | FILE_ATTRIBUTE_HIDDEN
            else:
                new_attrs = attrs & ~FILE_ATTRIBUTE_HIDDEN
            
            return bool(kernel32.SetFileAttributesW(filepath, new_attrs))
        except:
            return False
    
    @staticmethod
    def refresh_desktop():
        """Aktualisiert die Desktop-Ansicht"""
        try:
            SHCNE_ASSOCCHANGED = 0x08000000
            SHCNF_IDLIST = 0x0000
            ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
        except:
            pass
    
    @staticmethod
    def get_desktop_path():
        """Gibt den Desktop-Pfad zurück"""
        try:
            if HAS_SHELL:
                return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
        except:
            pass
        return str(Path.home() / "Desktop")
    
    @staticmethod
    def snap_to_grid(x, y):
        """Rastet Koordinaten auf dem Desktop-Grid ein"""
        grid_x = round((x - DESKTOP_MARGIN_X) / DESKTOP_GRID_X) * DESKTOP_GRID_X + DESKTOP_MARGIN_X
        grid_y = round((y - DESKTOP_MARGIN_Y) / DESKTOP_GRID_Y) * DESKTOP_GRID_Y + DESKTOP_MARGIN_Y
        return max(DESKTOP_MARGIN_X, grid_x), max(DESKTOP_MARGIN_Y, grid_y)


# ============================================================================
# Icon Extraction
# ============================================================================

class IconExtractor:
    """Extrahiert Datei-Icons mit Transparenz"""

    _cache = {}
    _provider = None

    @classmethod
    def _get_provider(cls):
        if cls._provider is None:
            cls._provider = QFileIconProvider()
        return cls._provider

    @classmethod
    def get_icon(cls, filepath, size=48):
        """Holt Icon für eine Datei als QPixmap"""
        cache_key = (filepath, size)
        if cache_key in cls._cache:
            return cls._cache[cache_key]

        # Primär: Qt QFileIconProvider (nutzt dieselbe Windows Shell API wie der Desktop)
        pixmap = cls._extract_icon_qt(filepath, size)

        # Fallback: Win32 API Extraktion
        if pixmap is None or pixmap.isNull():
            pil_img = cls._extract_icon(filepath, size)
            pixmap = pil_to_qpixmap(pil_img)

        if pixmap and not pixmap.isNull():
            cls._cache[cache_key] = pixmap
        return pixmap

    @classmethod
    def _extract_icon_qt(cls, filepath, size):
        """Extrahiert Icon über Qt's QFileIconProvider (gleiche Icons wie Desktop)"""
        try:
            if not filepath or not os.path.exists(filepath):
                return None
            provider = cls._get_provider()
            file_info = QFileInfo(filepath)
            icon = provider.icon(file_info)
            if icon.isNull():
                return None
            pixmap = icon.pixmap(QSize(size, size))
            return pixmap if not pixmap.isNull() else None
        except Exception:
            return None
    
    @classmethod
    def _extract_icon(cls, filepath, size):
        """Extrahiert Icon aus Datei"""
        if not HAS_WIN32 or not os.path.exists(filepath):
            return cls._get_default_icon(filepath, size)
        
        try:
            # SHFILEINFOW Struktur
            class SHFILEINFOW(ctypes.Structure):
                _fields_ = [
                    ('hIcon', ctypes.c_void_p),
                    ('iIcon', ctypes.c_int),
                    ('dwAttributes', ctypes.c_uint),
                    ('szDisplayName', ctypes.c_wchar * 260),
                    ('szTypeName', ctypes.c_wchar * 80),
                ]
            
            info = SHFILEINFOW()
            SHGFI_ICON = 0x000000100
            SHGFI_LARGEICON = 0x000000000
            
            result = ctypes.windll.shell32.SHGetFileInfoW(
                filepath, 0, ctypes.byref(info), ctypes.sizeof(info),
                SHGFI_ICON | SHGFI_LARGEICON
            )
            
            if not result or not info.hIcon:
                return cls._get_default_icon(filepath, size)
            
            hicon = info.hIcon
            
            # Icon Größe
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
            
            # DC erstellen
            hdc_screen = win32gui.GetDC(0)
            hdc = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem = hdc.CreateCompatibleDC()
            
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
            hdc_mem.SelectObject(hbmp)
            
            # Icon auf schwarzem Hintergrund zeichnen
            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0x00000000)
            hdc_mem.DrawIcon((0, 0), hicon)
            
            bmpstr = hbmp.GetBitmapBits(True)
            img_black = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr, 'raw', 'BGRX', 0, 1)
            
            # Icon auf weißem Hintergrund für Alpha-Berechnung
            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0x00FFFFFF)
            hdc_mem.DrawIcon((0, 0), hicon)
            bmpstr2 = hbmp.GetBitmapBits(True)
            img_white = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr2, 'raw', 'BGRX', 0, 1)
            
            # Alpha aus Differenz berechnen
            try:
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
                        0
                    ).astype(np.uint8)
                rgba[:, :, 3] = alpha
                
                img = Image.fromarray(rgba, 'RGBA')
            except ImportError:
                img = img_black.convert('RGBA')
            
            # Skalieren
            if ico_x != size or ico_y != size:
                img = img.resize((size, size), Image.Resampling.LANCZOS)
            
            # Aufräumen
            win32gui.DestroyIcon(hicon)
            hdc_mem.DeleteDC()
            win32gui.ReleaseDC(0, hdc_screen)
            
            return img
            
        except Exception as e:
            return cls._get_default_icon(filepath, size)
    
    @classmethod
    def _get_default_icon(cls, filepath, size):
        """Erstellt ein Fallback-Icon"""
        try:
            img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
            draw = ImageDraw.Draw(img)
            
            ext = Path(filepath).suffix.lower() if filepath else ""
            name = Path(filepath).stem if filepath else ""
            
            # Farben nach Dateityp
            colors = {
                '.exe': (0, 120, 212), '.lnk': (0, 120, 212),
                '.bat': (255, 165, 0), '.py': (55, 118, 171),
                '.txt': (107, 107, 107), '.pdf': (220, 30, 30),
            }
            color = colors.get(ext, (0, 120, 212))
            
            # Abgerundetes Rechteck
            margin = 4
            draw.rounded_rectangle(
                [margin, margin, size - margin, size - margin],
                radius=8, fill=(*color, 230)
            )
            
            # Buchstabe
            letter = name[0].upper() if name else "?"
            try:
                font = ImageFont.truetype("segoeui.ttf", size // 2)
            except:
                font = ImageFont.load_default()
            
            bbox = draw.textbbox((0, 0), letter, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            x = (size - text_width) // 2
            y = (size - text_height) // 2 - 2
            
            draw.text((x, y), letter, fill=(255, 255, 255, 230), font=font)
            
            return img
        except:
            return None


def pil_to_qpixmap(pil_image):
    """Konvertiert PIL Image zu QPixmap"""
    if pil_image is None:
        return QPixmap()
    
    if pil_image.mode != 'RGBA':
        pil_image = pil_image.convert('RGBA')
    
    data = pil_image.tobytes('raw', 'RGBA')
    qimage = QImage(data, pil_image.width, pil_image.height, QImage.Format.Format_RGBA8888)
    return QPixmap.fromImage(qimage)


# ============================================================================
# Frosted Glass Widget
# ============================================================================

class FrostedGlassWidget(QWidget):
    """Basis-Widget mit Frosted Glass Effekt"""

    def __init__(self, parent=None):
        super().__init__(parent)
        self.setWindowFlags(
            Qt.WindowType.FramelessWindowHint |
            Qt.WindowType.Tool |
            Qt.WindowType.WindowStaysOnTopHint
        )
        self.setAttribute(Qt.WidgetAttribute.WA_TranslucentBackground)

        self._blur_enabled = False
        self._corner_radius = 20

    def showEvent(self, event):
        """Aktiviert Blur-Effekt wenn Fenster angezeigt wird"""
        super().showEvent(event)
        self._enable_blur()

    def _enable_blur(self):
        """Aktiviert den Windows Blur-Effekt hinter dem Fenster"""
        if self._blur_enabled:
            return

        hwnd = int(self.winId())

        # BlurWindow mit Standard-Blur (nicht Acrylic - das erzeugt opakes Grau)
        if HAS_BLURWINDOW:
            try:
                GlobalBlur(hwnd, Dark=False, Acrylic=False, QWidget=self)
                print("✓ BlurWindow Blur aktiviert")
                self._blur_enabled = True
                return
            except Exception as e:
                print(f"BlurWindow fehlgeschlagen: {e}")

        # Fallback: DWM Blur Behind (kein Acrylic, kein Mica)
        if enable_blur_behind(hwnd):
            print("✓ Blur Behind aktiviert")
            self._blur_enabled = True
            return

        print("⚠ Kein Blur-Effekt verfügbar - reine Transparenz aktiv")

    def paintEvent(self, event):
        """Zeichnet die halbtransparente Glaskachel"""
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)

        # Halbtransparenter weißer Hintergrund
        painter.setBrush(QBrush(QColor(255, 255, 255, 40)))
        # Leichter weißer Rand als Lichtkante
        painter.setPen(QPen(QColor(255, 255, 255, 80), 1))
        # Abgerundetes Rechteck
        painter.drawRoundedRect(self.rect(), self._corner_radius, self._corner_radius)


# ============================================================================
# Icon Widget für Shortcuts
# ============================================================================

class IconWidget(QWidget):
    """Ein einzelnes Icon mit Label"""
    
    clicked = pyqtSignal()
    doubleClicked = pyqtSignal()
    rightClicked = pyqtSignal(QPoint)
    dragStarted = pyqtSignal(int)
    
    def __init__(self, shortcut, index, icon_w=36, icon_h=36, parent=None):
        super().__init__(parent)
        self.shortcut = shortcut
        self.index = index
        self.icon_w = icon_w
        self.icon_h = icon_h
        self._is_hovered = False
        self._is_pressed = False
        self._drag_start_pos = None

        self.setFixedSize(icon_w + 24, icon_h + 32)
        self.setCursor(Qt.CursorShape.PointingHandCursor)
        self.setAcceptDrops(True)

        # Icon laden
        self._pixmap = self._load_icon()

    def _load_icon(self):
        """Lädt das Icon für den Shortcut"""
        return IconExtractor.get_icon(self.shortcut["path"], max(self.icon_w, self.icon_h))
    
    def paintEvent(self, event):
        painter = QPainter(self)
        painter.setRenderHint(QPainter.RenderHint.Antialiasing)
        
        # Hover/Press Hintergrund (Frosted Glass)
        if self._is_pressed:
            painter.fillRect(self.rect(), QColor(255, 255, 255, 55))
        elif self._is_hovered:
            painter.fillRect(self.rect(), QColor(255, 255, 255, 35))
        
        # Icon zentriert
        icon_x = (self.width() - self.icon_w) // 2
        icon_y = 4

        if self._pixmap and not self._pixmap.isNull():
            painter.drawPixmap(icon_x, icon_y, self.icon_w, self.icon_h, self._pixmap)

        # Name
        name = self.shortcut.get("name", "")
        if len(name) > 10:
            name = name[:9] + "…"

        painter.setPen(QColor(224, 224, 224))
        painter.setFont(QFont("Segoe UI", 8))

        text_rect = QRect(0, icon_y + self.icon_h + 2, self.width(), 20)
        painter.drawText(text_rect, Qt.AlignmentFlag.AlignHCenter | Qt.AlignmentFlag.AlignTop, name)
    
    def enterEvent(self, event):
        self._is_hovered = True
        self.update()
    
    def leaveEvent(self, event):
        self._is_hovered = False
        self._is_pressed = False
        self.update()
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._is_pressed = True
            self._drag_start_pos = event.pos()
            self.update()
        elif event.button() == Qt.MouseButton.RightButton:
            self.rightClicked.emit(event.globalPosition().toPoint())
    
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._is_pressed = False
            self.update()
            if self.rect().contains(event.pos()):
                self.clicked.emit()
    
    def mouseDoubleClickEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self.doubleClicked.emit()
    
    def mouseMoveEvent(self, event):
        if self._drag_start_pos and (event.pos() - self._drag_start_pos).manhattanLength() > 10:
            self.dragStarted.emit(self.index)
            self._drag_start_pos = None


# ============================================================================
# Settings Dialog
# ============================================================================

class TileSettingsDialog(QDialog):
    """Einstellungs-Dialog mit Slidern und Live-Vorschau"""

    STYLE = """
        QDialog {
            background-color: #2d2d3d;
            color: #e0e0e0;
        }
        QGroupBox {
            font-family: 'Segoe UI';
            font-weight: bold;
            border: 1px solid rgba(255, 255, 255, 30);
            border-radius: 8px;
            margin-top: 12px;
            padding-top: 16px;
            color: #ffffff;
        }
        QGroupBox::title {
            subcontrol-origin: margin;
            left: 10px;
            padding: 0 4px;
        }
        QSlider::groove:horizontal {
            height: 6px;
            background: rgba(255, 255, 255, 20);
            border-radius: 3px;
        }
        QSlider::handle:horizontal {
            background: #0078d4;
            width: 16px;
            height: 16px;
            margin: -5px 0;
            border-radius: 8px;
        }
        QSlider::handle:horizontal:hover {
            background: #1a8ae8;
        }
        QSlider::sub-page:horizontal {
            background: rgba(0, 120, 212, 120);
            border-radius: 3px;
        }
        QLabel {
            color: #c0c0c0;
        }
        QLabel[class="value"] {
            color: #ffffff;
            font-family: 'Segoe UI';
            font-weight: bold;
            min-width: 45px;
        }
        QPushButton {
            background: rgba(255, 255, 255, 15);
            border: 1px solid rgba(255, 255, 255, 30);
            border-radius: 4px;
            color: #e0e0e0;
            padding: 4px 12px;
        }
        QPushButton:checked {
            background: rgba(0, 120, 212, 80);
            border-color: rgba(0, 120, 212, 150);
        }
        QPushButton:hover {
            background: rgba(255, 255, 255, 25);
        }
    """

    def __init__(self, tile, parent=None):
        super().__init__(parent)
        self.setWindowTitle("Kachel-Einstellungen")
        self.setMinimumWidth(460)
        self.setStyleSheet(self.STYLE)
        self._tile = tile
        self._config = tile.config
        self._original = self._snapshot()
        self._setup_ui()

    def _snapshot(self):
        """Speichert den aktuellen Zustand für Restore bei Cancel"""
        keys = [
            "collapsed_tile_w", "collapsed_tile_h",
            "expanded_tile_w", "expanded_tile_h",
            "collapsed_icon_w", "collapsed_icon_h",
            "expanded_icon_w", "expanded_icon_h",
            "collapsed_spacing_h", "collapsed_spacing_v",
            "expanded_spacing_h", "expanded_spacing_v",
        ]
        return {k: self._config.get(k, self._defaults()[k]) for k in keys}

    @staticmethod
    def _defaults():
        return {
            "collapsed_tile_w": 120, "collapsed_tile_h": 120,
            "expanded_tile_w": 260, "expanded_tile_h": 320,
            "collapsed_icon_w": 40, "collapsed_icon_h": 40,
            "expanded_icon_w": 36, "expanded_icon_h": 36,
            "collapsed_spacing_h": 4, "collapsed_spacing_v": 4,
            "expanded_spacing_h": 8, "expanded_spacing_v": 8,
        }

    def _create_slider_row(self, label, value, min_v, max_v):
        """Erstellt Label + Slider + Wert-Anzeige"""
        row = QWidget()
        layout = QHBoxLayout(row)
        layout.setContentsMargins(0, 0, 0, 0)

        lbl = QLabel(label)
        lbl.setFixedWidth(50)

        slider = QSlider(Qt.Orientation.Horizontal)
        slider.setRange(min_v, max_v)
        slider.setValue(value)

        val_label = QLabel(f"{value} px")
        val_label.setProperty("class", "value")
        val_label.setFixedWidth(50)

        slider.valueChanged.connect(lambda v: val_label.setText(f"{v} px"))

        layout.addWidget(lbl)
        layout.addWidget(slider, stretch=1)
        layout.addWidget(val_label)

        return row, slider, val_label

    def _create_linked_slider_pair(self, label_w, val_w, label_h, val_h, min_v, max_v):
        """Erstellt zwei Slider-Zeilen mit Link-Toggle"""
        container = QWidget()
        vlayout = QVBoxLayout(container)
        vlayout.setContentsMargins(0, 0, 0, 0)
        vlayout.setSpacing(4)

        row_w, slider_w, _ = self._create_slider_row(label_w, val_w, min_v, max_v)

        # Link-Toggle zwischen den Zeilen
        link_row = QWidget()
        link_layout = QHBoxLayout(link_row)
        link_layout.setContentsMargins(50, 0, 50, 0)
        link_btn = QPushButton("🔗 verknüpft")
        link_btn.setCheckable(True)
        link_btn.setChecked(val_w == val_h)
        link_btn.setFixedHeight(22)
        link_btn.clicked.connect(lambda checked: link_btn.setText(
            "🔗 verknüpft" if checked else "🔓 getrennt"))
        if val_w != val_h:
            link_btn.setText("🔓 getrennt")
        link_layout.addWidget(link_btn)

        row_h, slider_h, _ = self._create_slider_row(label_h, val_h, min_v, max_v)

        def on_w(v):
            if link_btn.isChecked():
                slider_h.blockSignals(True)
                slider_h.setValue(v)
                slider_h.blockSignals(False)
            self._apply_live()

        def on_h(v):
            if link_btn.isChecked():
                slider_w.blockSignals(True)
                slider_w.setValue(v)
                slider_w.blockSignals(False)
            self._apply_live()

        slider_w.valueChanged.connect(on_w)
        slider_h.valueChanged.connect(on_h)

        vlayout.addWidget(row_w)
        vlayout.addWidget(link_row)
        vlayout.addWidget(row_h)

        return container, slider_w, slider_h

    def _create_spacing_sliders(self, val_h, val_v, min_v, max_v):
        """Erstellt Slider-Paar für Abstände"""
        container = QWidget()
        vlayout = QVBoxLayout(container)
        vlayout.setContentsMargins(0, 0, 0, 0)
        vlayout.setSpacing(4)

        row_h, slider_h, _ = self._create_slider_row("H:", val_h, min_v, max_v)
        row_v, slider_v, _ = self._create_slider_row("V:", val_v, min_v, max_v)

        slider_h.valueChanged.connect(lambda: self._apply_live())
        slider_v.valueChanged.connect(lambda: self._apply_live())

        vlayout.addWidget(row_h)
        vlayout.addWidget(row_v)

        return container, slider_h, slider_v

    def _setup_ui(self):
        layout = QVBoxLayout(self)
        c = self._config
        d = self._defaults()

        # Scrollable content
        scroll = QScrollArea()
        scroll.setWidgetResizable(True)
        scroll.setStyleSheet("QScrollArea { background: transparent; border: none; }")
        content = QWidget()
        clayout = QVBoxLayout(content)

        # Kleine Kachel
        g1 = QGroupBox("Kleine Kachel (eingeklappt)")
        g1l = QVBoxLayout(g1)
        r1, self.coll_tile_w, self.coll_tile_h = self._create_linked_slider_pair(
            "Breite:", c.get("collapsed_tile_w", d["collapsed_tile_w"]),
            "Höhe:", c.get("collapsed_tile_h", d["collapsed_tile_h"]), 60, 400)
        g1l.addWidget(r1)
        clayout.addWidget(g1)

        # Große Kachel
        g2 = QGroupBox("Große Kachel (ausgeklappt)")
        g2l = QVBoxLayout(g2)
        r2, self.exp_tile_w, self.exp_tile_h = self._create_linked_slider_pair(
            "Breite:", c.get("expanded_tile_w", d["expanded_tile_w"]),
            "Höhe:", c.get("expanded_tile_h", d["expanded_tile_h"]), 120, 800)
        g2l.addWidget(r2)
        clayout.addWidget(g2)

        # Icon-Größe (eingeklappt)
        g3 = QGroupBox("Icon-Größe (eingeklappt)")
        g3l = QVBoxLayout(g3)
        r3, self.coll_icon_w, self.coll_icon_h = self._create_linked_slider_pair(
            "Breite:", c.get("collapsed_icon_w", d["collapsed_icon_w"]),
            "Höhe:", c.get("collapsed_icon_h", d["collapsed_icon_h"]), 16, 128)
        g3l.addWidget(r3)
        clayout.addWidget(g3)

        # Icon-Größe (ausgeklappt)
        g4 = QGroupBox("Icon-Größe (ausgeklappt)")
        g4l = QVBoxLayout(g4)
        r4, self.exp_icon_w, self.exp_icon_h = self._create_linked_slider_pair(
            "Breite:", c.get("expanded_icon_w", d["expanded_icon_w"]),
            "Höhe:", c.get("expanded_icon_h", d["expanded_icon_h"]), 16, 128)
        g4l.addWidget(r4)
        clayout.addWidget(g4)

        # Abstände (eingeklappt)
        g5 = QGroupBox("Abstände (eingeklappt)")
        g5l = QVBoxLayout(g5)
        r5, self.coll_spacing_h, self.coll_spacing_v = self._create_spacing_sliders(
            c.get("collapsed_spacing_h", d["collapsed_spacing_h"]),
            c.get("collapsed_spacing_v", d["collapsed_spacing_v"]), 0, 40)
        g5l.addWidget(r5)
        clayout.addWidget(g5)

        # Abstände (ausgeklappt)
        g6 = QGroupBox("Abstände (ausgeklappt)")
        g6l = QVBoxLayout(g6)
        r6, self.exp_spacing_h, self.exp_spacing_v = self._create_spacing_sliders(
            c.get("expanded_spacing_h", d["expanded_spacing_h"]),
            c.get("expanded_spacing_v", d["expanded_spacing_v"]), 0, 40)
        g6l.addWidget(r6)
        clayout.addWidget(g6)

        scroll.setWidget(content)
        layout.addWidget(scroll, stretch=1)

        # Buttons
        buttons = QDialogButtonBox(
            QDialogButtonBox.StandardButton.Ok | QDialogButtonBox.StandardButton.Cancel)
        buttons.accepted.connect(self.accept)
        buttons.rejected.connect(self.reject)
        layout.addWidget(buttons)

    def _apply_live(self):
        """Wendet die aktuellen Slider-Werte sofort auf die Kachel an"""
        tile = self._tile
        tile.collapsed_width = self.coll_tile_w.value()
        tile.collapsed_height = self.coll_tile_h.value()
        tile.expanded_width = self.exp_tile_w.value()
        tile.expanded_height = self.exp_tile_h.value()
        tile.icon_w_collapsed = self.coll_icon_w.value()
        tile.icon_h_collapsed = self.coll_icon_h.value()
        tile.icon_w_expanded = self.exp_icon_w.value()
        tile.icon_h_expanded = self.exp_icon_h.value()
        tile.spacing_h_collapsed = self.coll_spacing_h.value()
        tile.spacing_v_collapsed = self.coll_spacing_v.value()
        tile.spacing_h_expanded = self.exp_spacing_h.value()
        tile.spacing_v_expanded = self.exp_spacing_v.value()
        # Spacings sofort anwenden
        tile.collapsed_grid_layout.setHorizontalSpacing(tile.spacing_h_collapsed)
        tile.collapsed_grid_layout.setVerticalSpacing(tile.spacing_v_collapsed)
        tile.icons_layout.setHorizontalSpacing(tile.spacing_h_expanded)
        tile.icons_layout.setVerticalSpacing(tile.spacing_v_expanded)
        # Kachel-Größe und Icons aktualisieren
        if tile.is_expanded:
            tile.resize(tile.expanded_width, tile.expanded_height)
            tile._update_expanded_icons()
        else:
            tile.resize(tile.collapsed_width, tile.collapsed_height)
            tile._update_collapsed_icons()

    def get_values(self):
        """Gibt alle Einstellungen als Dict zurück"""
        return {
            "collapsed_tile_w": self.coll_tile_w.value(),
            "collapsed_tile_h": self.coll_tile_h.value(),
            "expanded_tile_w": self.exp_tile_w.value(),
            "expanded_tile_h": self.exp_tile_h.value(),
            "collapsed_icon_w": self.coll_icon_w.value(),
            "collapsed_icon_h": self.coll_icon_h.value(),
            "expanded_icon_w": self.exp_icon_w.value(),
            "expanded_icon_h": self.exp_icon_h.value(),
            "collapsed_spacing_h": self.coll_spacing_h.value(),
            "collapsed_spacing_v": self.coll_spacing_v.value(),
            "expanded_spacing_h": self.exp_spacing_h.value(),
            "expanded_spacing_v": self.exp_spacing_v.value(),
        }

    def get_original(self):
        """Gibt die Original-Werte vor dem Dialog zurück"""
        return self._original


# ============================================================================
# Folder Tile (Kachel)
# ============================================================================

class FolderTile(FrostedGlassWidget):
    """Eine einzelne Ordner-Kachel auf dem Desktop"""
    
    def __init__(self, manager, tile_id, config):
        super().__init__()
        self.manager = manager
        self.tile_id = tile_id
        self.config = config
        
        self.is_expanded = False
        self.animation_running = False
        
        self._drag_start_pos = None
        self._is_dragging = False
        self._icon_images = []
        
        # Größen
        self.collapsed_width = config.get("collapsed_tile_w", 120)
        self.collapsed_height = config.get("collapsed_tile_h", 120)
        self.expanded_width = config.get("expanded_tile_w", 260)
        self.expanded_height = config.get("expanded_tile_h", 320)
        self.icon_w_collapsed = config.get("collapsed_icon_w", 40)
        self.icon_h_collapsed = config.get("collapsed_icon_h", 40)
        self.icon_w_expanded = config.get("expanded_icon_w", 36)
        self.icon_h_expanded = config.get("expanded_icon_h", 36)
        self.spacing_h_collapsed = config.get("collapsed_spacing_h", 4)
        self.spacing_v_collapsed = config.get("collapsed_spacing_v", 4)
        self.spacing_h_expanded = config.get("expanded_spacing_h", 8)
        self.spacing_v_expanded = config.get("expanded_spacing_v", 8)
        
        # Position
        x = config.get("pos_x", DESKTOP_MARGIN_X)
        y = config.get("pos_y", DESKTOP_MARGIN_Y)
        x, y = WindowsDesktopAPI.snap_to_grid(x, y)
        
        self.setGeometry(x, y, self.collapsed_width, self.collapsed_height)
        
        # Layouts
        self._setup_ui()
        
        # Hover-Collapse Timer
        self._collapse_timer = QTimer()
        self._collapse_timer.setSingleShot(True)
        self._collapse_timer.timeout.connect(self._on_collapse_timeout)
        
        # Drag & Drop
        self.setAcceptDrops(True)
    
    def _setup_ui(self):
        """Erstellt die UI-Komponenten"""
        # Alle Kind-Widgets transparent machen
        self.setStyleSheet("QWidget { background: transparent; }")

        self.main_layout = QVBoxLayout(self)
        self.main_layout.setContentsMargins(8, 8, 8, 8)
        self.main_layout.setSpacing(4)

        # Collapsed View
        self.collapsed_widget = QWidget()
        self.collapsed_layout = QVBoxLayout(self.collapsed_widget)
        self.collapsed_layout.setContentsMargins(0, 0, 0, 0)

        # Icon Grid für Collapsed
        self.collapsed_grid = QWidget()
        self.collapsed_grid_layout = QGridLayout(self.collapsed_grid)
        self.collapsed_grid_layout.setContentsMargins(0, 0, 0, 0)
        self.collapsed_grid_layout.setHorizontalSpacing(self.spacing_h_collapsed)
        self.collapsed_grid_layout.setVerticalSpacing(self.spacing_v_collapsed)
        self.collapsed_layout.addWidget(self.collapsed_grid, stretch=1)
        
        # Name Label
        self.name_label = QLabel(self.config.get("name", "Ordner"))
        self.name_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.name_label.setStyleSheet("color: #ffffff; font-family: 'Segoe UI'; font-size: 9pt;")
        self.collapsed_layout.addWidget(self.name_label)
        
        self.main_layout.addWidget(self.collapsed_widget)
        
        # Expanded View (initially hidden)
        self.expanded_widget = QWidget()
        self.expanded_widget.hide()
        self.expanded_layout = QVBoxLayout(self.expanded_widget)
        self.expanded_layout.setContentsMargins(0, 0, 0, 0)
        self.expanded_layout.setSpacing(4)
        
        # Scroll Area für Icons
        self.scroll_area = QScrollArea()
        self.scroll_area.setWidgetResizable(True)
        self.scroll_area.setHorizontalScrollBarPolicy(Qt.ScrollBarPolicy.ScrollBarAlwaysOff)
        self.scroll_area.setStyleSheet("""
            QScrollArea { background: transparent; border: none; }
            QScrollBar:vertical { width: 6px; background: transparent; }
            QScrollBar::handle:vertical { background: rgba(255,255,255,60); border-radius: 3px; min-height: 20px; }
            QScrollBar::handle:vertical:hover { background: rgba(255,255,255,90); }
            QScrollBar::add-line:vertical, QScrollBar::sub-line:vertical { height: 0; }
        """)
        
        self.icons_container = QWidget()
        self.icons_container.setStyleSheet("background: transparent;")
        self.icons_layout = QGridLayout(self.icons_container)
        self.icons_layout.setContentsMargins(4, 4, 4, 4)
        self.icons_layout.setHorizontalSpacing(self.spacing_h_expanded)
        self.icons_layout.setVerticalSpacing(self.spacing_v_expanded)
        
        self.scroll_area.setWidget(self.icons_container)
        self.expanded_layout.addWidget(self.scroll_area, stretch=1)
        
        # Footer mit Name (klickbar zum Umbenennen)
        self.footer_label = QLabel(self.config.get("name", "Ordner"))
        self.footer_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
        self.footer_label.setStyleSheet("""
            color: #ffffff;
            font-family: 'Segoe UI Semibold';
            font-size: 10pt;
            padding: 6px;
            border-top: 1px solid rgba(255, 255, 255, 30);
        """)
        self.footer_label.setCursor(Qt.CursorShape.PointingHandCursor)
        self.footer_label.mousePressEvent = self._start_rename
        self.expanded_layout.addWidget(self.footer_label)
        
        self.main_layout.addWidget(self.expanded_widget)
        
        # Icons zeichnen
        self._update_collapsed_icons()
    
    def _update_collapsed_icons(self):
        """Aktualisiert die Icons in der eingeklappten Ansicht"""
        # Grid leeren
        while self.collapsed_grid_layout.count():
            item = self.collapsed_grid_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        self._icon_images.clear()
        shortcuts = self.config.get("shortcuts", [])[:4]
        
        if not shortcuts:
            # Leerer Ordner - Ordner-Icon anzeigen
            folder_label = QLabel("📁")
            folder_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            folder_label.setStyleSheet("font-size: 48px;")
            self.collapsed_grid_layout.addWidget(folder_label, 0, 0, 2, 2)
        else:
            for i, shortcut in enumerate(shortcuts):
                row = i // 2
                col = i % 2
                
                icon_label = QLabel()
                icon_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
                icon_label.setFixedSize(self.icon_w_collapsed, self.icon_h_collapsed)

                pixmap = IconExtractor.get_icon(shortcut["path"], max(self.icon_w_collapsed, self.icon_h_collapsed))
                if pixmap and not pixmap.isNull():
                    icon_label.setPixmap(pixmap.scaled(
                        self.icon_w_collapsed, self.icon_h_collapsed,
                        Qt.AspectRatioMode.KeepAspectRatio,
                        Qt.TransformationMode.SmoothTransformation
                    ))
                    self._icon_images.append(pixmap)
                
                self.collapsed_grid_layout.addWidget(icon_label, row, col)
    
    def _update_expanded_icons(self):
        """Aktualisiert die Icons in der erweiterten Ansicht"""
        # Grid leeren
        while self.icons_layout.count():
            item = self.icons_layout.takeAt(0)
            if item.widget():
                item.widget().deleteLater()
        
        shortcuts = self.config.get("shortcuts", [])
        
        if not shortcuts:
            empty_label = QLabel("Leer\n\nDateien hierher ziehen")
            empty_label.setAlignment(Qt.AlignmentFlag.AlignCenter)
            empty_label.setStyleSheet("color: rgba(255, 255, 255, 120); font-size: 10pt;")
            self.icons_layout.addWidget(empty_label, 0, 0, 1, 3)
        else:
            cols = 3
            for i, shortcut in enumerate(shortcuts):
                row = i // cols
                col = i % cols
                
                icon_widget = IconWidget(shortcut, i, self.icon_w_expanded, self.icon_h_expanded)
                icon_widget.doubleClicked.connect(lambda s=shortcut: self._launch_shortcut(s["path"]))
                icon_widget.rightClicked.connect(lambda pos, idx=i: self._show_item_menu(pos, idx))
                icon_widget.dragStarted.connect(self._start_icon_drag)
                
                self.icons_layout.addWidget(icon_widget, row, col)
    
    def expand(self):
        """Kachel expandieren"""
        if self.is_expanded or self.animation_running:
            return
        
        self.animation_running = True
        self.is_expanded = True
        self._collapse_timer.stop()
        
        # Nach vorne bringen
        self.setWindowFlags(self.windowFlags() | Qt.WindowType.WindowStaysOnTopHint)
        self.show()
        self.raise_()
        
        # Animation
        self._animate_size(self.expanded_width, self.expanded_height, self._show_expanded)
    
    def collapse(self):
        """Kachel einklappen"""
        if not self.is_expanded or self.animation_running:
            return
        
        self.animation_running = True
        self.is_expanded = False
        
        # Animation
        self._animate_size(self.collapsed_width, self.collapsed_height, self._show_collapsed)
    
    def _animate_size(self, target_width, target_height, callback):
        """Animiert die Größenänderung"""
        self.size_anim = QPropertyAnimation(self, b"size")
        self.size_anim.setDuration(200)
        self.size_anim.setStartValue(self.size())
        self.size_anim.setEndValue(QSize(target_width, target_height))
        self.size_anim.setEasingCurve(QEasingCurve.Type.OutCubic)
        self.size_anim.finished.connect(callback)
        self.size_anim.start()
    
    def _show_expanded(self):
        """Zeigt die erweiterte Ansicht"""
        self.collapsed_widget.hide()
        self._update_expanded_icons()
        self.expanded_widget.show()
        self.animation_running = False
    
    def _show_collapsed(self):
        """Zeigt die eingeklappte Ansicht"""
        self.expanded_widget.hide()
        self._update_collapsed_icons()
        self.collapsed_widget.show()
        self.animation_running = False
        
        # Position auf Grid einrasten
        x, y = WindowsDesktopAPI.snap_to_grid(self.x(), self.y())
        self.move(x, y)
        self.config["pos_x"] = x
        self.config["pos_y"] = y
        self.manager.save_config()
    
    def _launch_shortcut(self, path):
        """Startet eine Verknüpfung"""
        try:
            os.startfile(path)
        except Exception as e:
            print(f"Fehler beim Öffnen: {e}")
    
    def _show_item_menu(self, pos, index):
        """Zeigt Kontextmenü für ein Item"""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: rgba(240, 240, 255, 180);
                color: #1a1a2e;
                border: 1px solid rgba(255, 255, 255, 120);
                border-radius: 8px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: rgba(255, 255, 255, 120);
            }
        """)

        open_action = menu.addAction("▶️ Öffnen")
        restore_action = menu.addAction("📤 Auf Desktop wiederherstellen")
        menu.addSeparator()
        remove_action = menu.addAction("🗑️ Aus Ordner entfernen")
        
        action = menu.exec(pos)
        
        shortcuts = self.config.get("shortcuts", [])
        if action == open_action and 0 <= index < len(shortcuts):
            self._launch_shortcut(shortcuts[index]["path"])
        elif action == restore_action:
            self._restore_to_desktop(index)
        elif action == remove_action:
            self._remove_shortcut(index)
    
    def _restore_to_desktop(self, index):
        """Stellt Shortcut auf Desktop wieder her"""
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            filepath = shortcuts[index]["path"]
            WindowsDesktopAPI.set_file_hidden(filepath, False)
            WindowsDesktopAPI.refresh_desktop()
            del self.config["shortcuts"][index]
            self.manager.save_config()
            self._update_expanded_icons()
    
    def _remove_shortcut(self, index):
        """Entfernt Shortcut aus Ordner (ohne Desktop-Wiederherstellung)"""
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            del self.config["shortcuts"][index]
            self.manager.save_config()
            self._update_expanded_icons()
    
    def _start_icon_drag(self, index):
        """Startet Drag eines Icons"""
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            # Auf Desktop wiederherstellen
            self._restore_to_desktop(index)
    
    def _start_rename(self, event):
        """Startet die Umbenennung"""
        name, ok = QInputDialog.getText(
            self, "Umbenennen", "Neuer Name:",
            QLineEdit.EchoMode.Normal, self.config.get("name", "Ordner")
        )
        if ok and name:
            self.config["name"] = name
            self.name_label.setText(name)
            self.footer_label.setText(name)
            self.manager.save_config()
    
    def mousePressEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            self._drag_start_pos = event.pos()
            self._is_dragging = False
        elif event.button() == Qt.MouseButton.RightButton:
            self._show_context_menu(event.globalPosition().toPoint())
    
    def mouseMoveEvent(self, event):
        if self._drag_start_pos:
            if (event.pos() - self._drag_start_pos).manhattanLength() > 10:
                self._is_dragging = True
                new_pos = self.mapToGlobal(event.pos()) - self._drag_start_pos
                self.move(new_pos)
    
    def mouseReleaseEvent(self, event):
        if event.button() == Qt.MouseButton.LeftButton:
            if self._is_dragging:
                # Auf Grid einrasten
                x, y = WindowsDesktopAPI.snap_to_grid(self.x(), self.y())
                self.move(x, y)
                self.config["pos_x"] = x
                self.config["pos_y"] = y
                self.manager.save_config()
            elif not self.is_expanded:
                self.expand()
            
            self._drag_start_pos = None
            self._is_dragging = False
    
    def enterEvent(self, event):
        """Maus betritt Widget"""
        self._collapse_timer.stop()
        if not self.is_expanded:
            # Auto-expand nach kurzer Verzögerung wenn Maus gedrückt
            if QApplication.mouseButtons() & Qt.MouseButton.LeftButton:
                self.expand()
    
    def leaveEvent(self, event):
        """Maus verlässt Widget"""
        if self.is_expanded and not self.animation_running:
            self._collapse_timer.start(800)  # 800ms Verzögerung
    
    def _on_collapse_timeout(self):
        """Collapse nach Timeout"""
        if self.is_expanded and not self.rect().contains(self.mapFromGlobal(QCursor.pos())):
            self.collapse()
    
    def _show_context_menu(self, pos):
        """Zeigt das Kontextmenü"""
        menu = QMenu(self)
        menu.setStyleSheet("""
            QMenu {
                background-color: rgba(240, 240, 255, 180);
                color: #1a1a2e;
                border: 1px solid rgba(255, 255, 255, 120);
                border-radius: 8px;
                padding: 4px;
            }
            QMenu::item {
                padding: 6px 20px;
                border-radius: 4px;
            }
            QMenu::item:selected {
                background-color: rgba(255, 255, 255, 120);
            }
            QMenu::separator {
                height: 1px;
                background: rgba(0, 0, 0, 20);
                margin: 4px 8px;
            }
        """)
        
        rename_action = menu.addAction("✏️ Umbenennen")
        settings_action = menu.addAction("⚙️ Einstellungen")
        menu.addSeparator()
        new_tile_action = menu.addAction("➕ Neue Kachel")
        delete_action = menu.addAction("🗑️ Kachel löschen")
        menu.addSeparator()
        quit_action = menu.addAction("❌ Beenden")

        action = menu.exec(pos)

        if action == rename_action:
            self._start_rename(None)
        elif action == settings_action:
            self._show_settings()
        elif action == new_tile_action:
            self.manager.create_new_tile()
        elif action == delete_action:
            self.manager.delete_tile(self.tile_id)
        elif action == quit_action:
            self.manager.quit()
    
    def _show_settings(self):
        """Öffnet den Einstellungs-Dialog mit Live-Vorschau"""
        dlg = TileSettingsDialog(self)
        if dlg.exec() == QDialog.DialogCode.Accepted:
            # Werte übernehmen und speichern
            self.config.update(dlg.get_values())
            self.manager.save_config()
        else:
            # Bei Cancel: Originalwerte wiederherstellen
            original = dlg.get_original()
            self.config.update(original)
            self.collapsed_width = original["collapsed_tile_w"]
            self.collapsed_height = original["collapsed_tile_h"]
            self.expanded_width = original["expanded_tile_w"]
            self.expanded_height = original["expanded_tile_h"]
            self.icon_w_collapsed = original["collapsed_icon_w"]
            self.icon_h_collapsed = original["collapsed_icon_h"]
            self.icon_w_expanded = original["expanded_icon_w"]
            self.icon_h_expanded = original["expanded_icon_h"]
            self.spacing_h_collapsed = original["collapsed_spacing_h"]
            self.spacing_v_collapsed = original["collapsed_spacing_v"]
            self.spacing_h_expanded = original["expanded_spacing_h"]
            self.spacing_v_expanded = original["expanded_spacing_v"]
            self.collapsed_grid_layout.setHorizontalSpacing(self.spacing_h_collapsed)
            self.collapsed_grid_layout.setVerticalSpacing(self.spacing_v_collapsed)
            self.icons_layout.setHorizontalSpacing(self.spacing_h_expanded)
            self.icons_layout.setVerticalSpacing(self.spacing_v_expanded)
            if self.is_expanded:
                self.resize(self.expanded_width, self.expanded_height)
                self._update_expanded_icons()
            else:
                self.resize(self.collapsed_width, self.collapsed_height)
                self._update_collapsed_icons()

    def dragEnterEvent(self, event):
        if event.mimeData().hasUrls():
            event.acceptProposedAction()
            if not self.is_expanded:
                self.expand()
    
    def dropEvent(self, event):
        """Verarbeitet Drop-Events"""
        if event.mimeData().hasUrls():
            desktop_path = WindowsDesktopAPI.get_desktop_path()
            added = 0
            
            for url in event.mimeData().urls():
                filepath = url.toLocalFile()
                if not os.path.exists(filepath):
                    continue
                
                # Prüfen ob schon vorhanden
                if any(s["path"] == filepath for s in self.config.get("shortcuts", [])):
                    continue
                
                name = Path(filepath).stem
                
                if "shortcuts" not in self.config:
                    self.config["shortcuts"] = []
                
                self.config["shortcuts"].append({
                    "name": name,
                    "path": filepath
                })
                
                # Desktop-Icon verstecken wenn von Desktop
                if Path(filepath).parent == Path(desktop_path):
                    WindowsDesktopAPI.set_file_hidden(filepath, True)
                
                added += 1
            
            if added > 0:
                WindowsDesktopAPI.refresh_desktop()
                self.manager.save_config()
                self._update_expanded_icons()
                self._update_collapsed_icons()
            
            event.acceptProposedAction()
    
    def close(self):
        """Schließt die Kachel"""
        super().close()


# ============================================================================
# Desktop Folder Manager
# ============================================================================

class DesktopFolderManager:
    """Verwaltet alle Ordner-Kacheln"""
    
    CONFIG_FILE = Path.home() / ".desktop_folder_widget.json"
    
    def __init__(self):
        self.app = QApplication.instance() or QApplication(sys.argv)
        self.app.setQuitOnLastWindowClosed(False)
        
        self.tiles = {}
        self.config = self.load_config()
        
        # Mindestens eine Kachel erstellen
        if not self.config.get("tiles"):
            self.config["tiles"] = {
                "0": {
                    "name": "Apps",
                    "shortcuts": [],
                    "pos_x": DESKTOP_MARGIN_X,
                    "pos_y": DESKTOP_MARGIN_Y
                }
            }
            self.save_config()
        
        # Kacheln erstellen
        for tile_id, tile_config in self.config["tiles"].items():
            self.tiles[tile_id] = FolderTile(self, tile_id, tile_config)
            self.tiles[tile_id].show()
        
        print(f"✓ {len(self.tiles)} Kachel(n) erstellt")
    
    def load_config(self):
        """Lädt Konfiguration"""
        if self.CONFIG_FILE.exists():
            try:
                with open(self.CONFIG_FILE, "r", encoding="utf-8") as f:
                    return json.load(f)
            except:
                pass
        return {"tiles": {}}
    
    def save_config(self):
        """Speichert Konfiguration"""
        try:
            with open(self.CONFIG_FILE, "w", encoding="utf-8") as f:
                json.dump(self.config, f, indent=2, ensure_ascii=False)
        except Exception as e:
            print(f"Speicherfehler: {e}")
    
    def create_new_tile(self):
        """Neue Kachel erstellen"""
        existing = [int(i) for i in self.config["tiles"].keys()]
        new_id = str(max(existing) + 1 if existing else 0)
        
        # Position neben der letzten Kachel
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
            "pos_y": y
        }
        
        self.config["tiles"][new_id] = tile_config
        self.tiles[new_id] = FolderTile(self, new_id, tile_config)
        self.tiles[new_id].show()
        self.save_config()
    
    def delete_tile(self, tile_id):
        """Kachel löschen"""
        if tile_id in self.tiles:
            self.tiles[tile_id].close()
            del self.tiles[tile_id]
        
        if tile_id in self.config["tiles"]:
            del self.config["tiles"][tile_id]
        
        self.save_config()
        
        if not self.tiles:
            self.quit()
    
    def quit(self):
        """Beenden - Alle versteckten Dateien wiederherstellen"""
        print("\n" + "=" * 50)
        print("  Widget wird beendet...")
        print("=" * 50)
        
        restored = 0
        for tile_config in self.config.get("tiles", {}).values():
            for shortcut in tile_config.get("shortcuts", []):
                filepath = shortcut.get("path", "")
                if filepath and os.path.exists(filepath):
                    if WindowsDesktopAPI.set_file_hidden(filepath, False):
                        print(f"  ✓ Wiederhergestellt: {shortcut.get('name', 'Unbekannt')}")
                        restored += 1
        
        WindowsDesktopAPI.refresh_desktop()
        print(f"\n{restored} Icons wiederhergestellt.")
        print("=" * 50)
        
        self.save_config()
        
        for tile in list(self.tiles.values()):
            tile.close()
        
        self.app.quit()
    
    def run(self):
        """Hauptschleife"""
        return self.app.exec()


# ============================================================================
# Cleanup & Main
# ============================================================================

_app_instance = None
_cleanup_done = False


def cleanup_on_exit():
    """Wird beim Beenden aufgerufen"""
    global _app_instance, _cleanup_done
    
    if _cleanup_done:
        return
    _cleanup_done = True
    
    if _app_instance and hasattr(_app_instance, 'config'):
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
        except Exception as e:
            print(f"[Cleanup] Fehler: {e}")


def main():
    global _app_instance
    
    print("=" * 55)
    print("  Desktop Folder Widget - PyQt6 Frosted Glass Edition")
    print("=" * 55)
    print()
    print("Features:")
    print("  • Echter Frosted Glass / Acrylic Blur Effekt")
    print("  • Windows 11 Mica Unterstützung")
    print("  • Flüssige Animationen")
    print()
    print("Bedienung:")
    print("  • Linksklick         → Kachel öffnen")
    print("  • Rechtsklick        → Kontextmenü")
    print("  • Dateien hinziehen  → In Kachel verschieben")
    print("  • Doppelklick Icon   → Öffnen")
    print()
    print("WICHTIG: Beim Beenden werden alle Icons wiederhergestellt!")
    print("-" * 55)
    print()
    
    atexit.register(cleanup_on_exit)
    
    try:
        _app_instance = DesktopFolderManager()
        sys.exit(_app_instance.run())
    except KeyboardInterrupt:
        print("\n[Beendet durch Benutzer]")
    except Exception as e:
        print(f"\n[Fehler] {e}")
        import traceback
        traceback.print_exc()
    finally:
        cleanup_on_exit()


if __name__ == "__main__":
    main()
