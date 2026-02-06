"""
Desktop Folder Widget für Windows - Version 3.0
================================================
Tiefe Desktop-Integration:
- Widget wird in Desktop-Ebene (WorkerW) eingebettet
- Verschobene Verknüpfungen werden auf Desktop versteckt (Hidden-Attribut)
- Im Explorer bleiben sie sichtbar
- Kacheln rasten auf Desktop-Icon-Grid ein
- Icon-Ansicht wie auf dem Desktop

Autor: Claude
"""

import tkinter as tk
from tkinter import filedialog, messagebox, simpledialog
import os
import json
import subprocess
import sys
from pathlib import Path
import ctypes
from ctypes import wintypes
import tempfile
import shutil
import stat
import atexit
import signal

# Für Drag & Drop
try:
    import windnd
    HAS_WINDND = True
except ImportError:
    HAS_WINDND = False

# Für Icon-Extraktion
try:
    from PIL import Image, ImageTk, ImageDraw, ImageFont
    import win32gui
    import win32ui
    import win32con
    import win32api
    HAS_WIN32 = True
except ImportError:
    HAS_WIN32 = False

# Für Shell-Operationen
try:
    import pythoncom
    from win32com.shell import shell, shellcon
    from win32com.client import Dispatch
    HAS_SHELL = True
except ImportError:
    HAS_SHELL = False

# Windows DPI Awareness
try:
    from ctypes import windll
    windll.shcore.SetProcessDpiAwareness(1)
except:
    pass


# ============================================================================
# Windows API Definitionen
# ============================================================================

user32 = ctypes.windll.user32
kernel32 = ctypes.windll.kernel32

# Typen für 64-Bit Kompatibilität
if ctypes.sizeof(ctypes.c_void_p) == 8:
    # 64-Bit
    HWND = ctypes.c_uint64
    LONG_PTR = ctypes.c_int64
else:
    # 32-Bit
    HWND = ctypes.c_uint32
    LONG_PTR = ctypes.c_int32

# Funktions-Signaturen definieren
user32.FindWindowW.argtypes = [ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowW.restype = HWND

user32.FindWindowExW.argtypes = [HWND, HWND, ctypes.c_wchar_p, ctypes.c_wchar_p]
user32.FindWindowExW.restype = HWND

user32.SetParent.argtypes = [HWND, HWND]
user32.SetParent.restype = HWND

user32.GetParent.argtypes = [HWND]
user32.GetParent.restype = HWND

user32.GetWindowLongPtrW = user32.GetWindowLongPtrW if hasattr(user32, 'GetWindowLongPtrW') else user32.GetWindowLongW
user32.GetWindowLongPtrW.argtypes = [HWND, ctypes.c_int]
user32.GetWindowLongPtrW.restype = LONG_PTR

user32.SetWindowLongPtrW = user32.SetWindowLongPtrW if hasattr(user32, 'SetWindowLongPtrW') else user32.SetWindowLongW
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

# Prozess-Memory Funktionen für Desktop-Icon-Positionierung
kernel32.OpenProcess.argtypes = [ctypes.c_uint32, ctypes.c_bool, ctypes.c_uint32]
kernel32.OpenProcess.restype = ctypes.c_void_p

kernel32.CloseHandle.argtypes = [ctypes.c_void_p]
kernel32.CloseHandle.restype = ctypes.c_bool

kernel32.VirtualAllocEx.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_uint32, ctypes.c_uint32]
kernel32.VirtualAllocEx.restype = ctypes.c_void_p

kernel32.VirtualFreeEx.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.c_uint32]
kernel32.VirtualFreeEx.restype = ctypes.c_bool

kernel32.WriteProcessMemory.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
kernel32.WriteProcessMemory.restype = ctypes.c_bool

kernel32.ReadProcessMemory.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_void_p, ctypes.c_size_t, ctypes.POINTER(ctypes.c_size_t)]
kernel32.ReadProcessMemory.restype = ctypes.c_bool

user32.GetWindowThreadProcessId.argtypes = [ctypes.c_void_p, ctypes.POINTER(ctypes.c_ulong)]
user32.GetWindowThreadProcessId.restype = ctypes.c_uint32

# GetWindowRect für Multi-Monitor-Unterstützung
class RECT(ctypes.Structure):
    _fields_ = [
        ("left", ctypes.c_long),
        ("top", ctypes.c_long),
        ("right", ctypes.c_long),
        ("bottom", ctypes.c_long),
    ]

user32.GetWindowRect.argtypes = [ctypes.c_void_p, ctypes.POINTER(RECT)]
user32.GetWindowRect.restype = ctypes.c_bool

# SetWindowRgn für abgerundete Fenster
try:
    user32.SetWindowRgn.argtypes = [ctypes.c_void_p, ctypes.c_void_p, ctypes.c_bool]
    user32.SetWindowRgn.restype = ctypes.c_int
except:
    pass

# EnumWindows callback type
WNDENUMPROC = ctypes.WINFUNCTYPE(ctypes.c_bool, HWND, ctypes.c_void_p)
user32.EnumWindows.argtypes = [WNDENUMPROC, ctypes.c_void_p]
user32.EnumWindows.restype = ctypes.c_bool

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
HWND_BOTTOM = HWND(1)

# File Attributes
FILE_ATTRIBUTE_HIDDEN = 0x02
FILE_ATTRIBUTE_NORMAL = 0x80

# ============================================================================
# Windows Acrylic/Blur API für Glaseffekt
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

# Accent States
ACCENT_DISABLED = 0
ACCENT_ENABLE_BLURBEHIND = 3           # Windows 10 Blur
ACCENT_ENABLE_ACRYLICBLURBEHIND = 4    # Windows 10 1803+ Acrylic
ACCENT_ENABLE_HOSTBACKDROP = 5         # Windows 11 Mica

# Window Composition Attribute
WCA_ACCENT_POLICY = 19

def enable_acrylic_blur(hwnd, gradient_color=0x80000000):
    """
    Aktiviert Acrylic Blur (Glaseffekt) für ein Fenster.
    gradient_color: AABBGGRR Format (Alpha, Blue, Green, Red)
    Standard: 0x80000000 = halbtransparentes Schwarz
    """
    try:
        accent = ACCENT_POLICY()
        accent.AccentState = ACCENT_ENABLE_ACRYLICBLURBEHIND
        accent.AccentFlags = 2  # ACCENT_FLAG_DRAW_ALL
        accent.GradientColor = gradient_color
        accent.AnimationId = 0

        data = WINDOWCOMPOSITIONATTRIBDATA()
        data.Attribute = WCA_ACCENT_POLICY
        data.Data = ctypes.pointer(accent)
        data.SizeOfData = ctypes.sizeof(accent)

        result = ctypes.windll.user32.SetWindowCompositionAttribute(
            int(hwnd), ctypes.byref(data)
        )
        return bool(result)
    except Exception as e:
        print(f"    Acrylic Blur fehlgeschlagen: {e}")
        # Fallback: Standard-Blur versuchen
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

            result = ctypes.windll.user32.SetWindowCompositionAttribute(
                int(hwnd), ctypes.byref(data)
            )
            return bool(result)
        except Exception as e2:
            print(f"    Blur-Fallback fehlgeschlagen: {e2}")
            return False


def set_rounded_region(hwnd, width, height, radius=16):
    """
    Setzt eine abgerundete Fensterregion (zusätzlich zu -transparentcolor).
    Belt-and-suspenders: transparentcolor macht Ecken visuell transparent,
    SetWindowRgn verhindert Maus-Events in den Ecken.
    """
    try:
        gdi32 = ctypes.windll.gdi32
        # CreateRoundRectRgn(left, top, right, bottom, widthEllipse, heightEllipse)
        rgn = gdi32.CreateRoundRectRgn(0, 0, width + 1, height + 1, radius, radius)
        if rgn:
            # SetWindowRgn(hWnd, hRgn, bRedraw)
            user32.SetWindowRgn(int(hwnd), rgn, True)
            return True
    except Exception as e:
        print(f"    SetWindowRgn fehlgeschlagen: {e}")
    return False


def create_3d_tile_background(width, height, base_color=(26, 26, 46), corner_radius=16):
    """
    Erstellt ein 3D-Kachel-Hintergrundbild mit Licht-, Schatten- und Glaseffekten.
    Gibt ein PIL Image im RGBA-Modus zurück.
    """
    try:
        from PIL import Image, ImageDraw, ImageFilter
    except ImportError:
        return None

    # Größeres Bild für Anti-Aliasing
    scale = 2
    w, h = width * scale, height * scale
    r = corner_radius * scale

    img = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    # --- Äußerer Schatten (Drop Shadow) ---
    shadow = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    shadow_draw = ImageDraw.Draw(shadow)
    shadow_offset = 6 * scale
    shadow_draw.rounded_rectangle(
        [shadow_offset, shadow_offset, w - 2 * scale, h - 2 * scale],
        radius=r,
        fill=(0, 0, 0, 100)
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=8 * scale))
    img = Image.alpha_composite(img, shadow)

    # --- Hauptform (dunkler Hintergrund mit leichtem Gradient) ---
    main = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    main_draw = ImageDraw.Draw(main)

    # Gradient von oben (heller) nach unten (dunkler) simulieren
    br, bg_c, bb = base_color
    for i in range(h):
        t = i / h
        # Oben: etwas heller, unten: etwas dunkler
        cr = int(br + (1 - t) * 18)
        cg = int(bg_c + (1 - t) * 18)
        cb = int(bb + (1 - t) * 24)
        cr = min(255, max(0, cr))
        cg = min(255, max(0, cg))
        cb = min(255, max(0, cb))

    # Einfacher Gradient: Zwei Rechtecke überlagert
    # Obere Hälfte heller
    main_draw.rounded_rectangle(
        [0, 0, w - 1, h - 1],
        radius=r,
        fill=(br + 12, bg_c + 12, bb + 18, 210)
    )
    # Untere Hälfte dunkler (Overlay)
    gradient_overlay = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    go_draw = ImageDraw.Draw(gradient_overlay)
    go_draw.rounded_rectangle(
        [0, h // 3, w - 1, h - 1],
        radius=r,
        fill=(0, 0, 0, 40)
    )
    main = Image.alpha_composite(main, gradient_overlay)
    img = Image.alpha_composite(img, main)

    # --- Innerer Lichtrand oben (3D-Highlight) ---
    highlight = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    hl_draw = ImageDraw.Draw(highlight)
    # Heller Strich oben
    hl_draw.rounded_rectangle(
        [2 * scale, 2 * scale, w - 2 * scale, 6 * scale],
        radius=r,
        fill=(255, 255, 255, 35)
    )
    # Heller Rand oben und links (Licht von oben links)
    hl_draw.rounded_rectangle(
        [1 * scale, 1 * scale, w - 1 * scale, h // 4],
        radius=r,
        fill=(255, 255, 255, 15)
    )
    highlight = highlight.filter(ImageFilter.GaussianBlur(radius=3 * scale))
    img = Image.alpha_composite(img, highlight)

    # --- Glasglanz (Specular Highlight) ---
    specular = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    spec_draw = ImageDraw.Draw(specular)
    # Elliptischer Glanz oben
    spec_draw.ellipse(
        [w // 6, -h // 3, w * 5 // 6, h // 4],
        fill=(255, 255, 255, 20)
    )
    specular = specular.filter(ImageFilter.GaussianBlur(radius=12 * scale))
    img = Image.alpha_composite(img, specular)

    # --- Rand (feiner heller Rand oben, dunkler unten = 3D) ---
    border = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    bd_draw = ImageDraw.Draw(border)
    # Äußerer Rand - helle Kante oben
    bd_draw.rounded_rectangle(
        [0, 0, w - 1, h - 1],
        radius=r,
        outline=(255, 255, 255, 40),
        width=scale
    )
    # Dunkle Kante unten rechts
    dark_edge = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    de_draw = ImageDraw.Draw(dark_edge)
    de_draw.rounded_rectangle(
        [1, h // 2, w - 1, h - 1],
        radius=r,
        outline=(0, 0, 0, 50),
        width=scale
    )
    dark_edge = dark_edge.filter(ImageFilter.GaussianBlur(radius=2))
    border = Image.alpha_composite(border, dark_edge)
    img = Image.alpha_composite(img, border)

    # Herunterskalieren für Anti-Aliasing
    if hasattr(Image, 'Resampling'):
        img = img.resize((width, height), Image.Resampling.LANCZOS)
    else:
        img = img.resize((width, height), Image.LANCZOS)

    return img


def create_3d_folder_icon(width, height):
    """Erstellt ein 3D-Ordner-Icon mit Licht und Schatten"""
    try:
        from PIL import Image, ImageDraw, ImageFilter
    except ImportError:
        return None

    scale = 2
    w, h = width * scale, height * scale
    img = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    draw = ImageDraw.Draw(img)

    cx, cy = w // 2, h // 2 - 15 * scale

    fw = 80 * scale
    fh = 60 * scale
    tab_w = 34 * scale
    tab_h = 14 * scale
    r = 6 * scale

    # --- Schatten unter dem Ordner ---
    shadow = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    s_draw = ImageDraw.Draw(shadow)
    s_draw.rounded_rectangle(
        [cx - fw//2 + 4*scale, cy - fh//2 + 6*scale,
         cx + fw//2 + 4*scale, cy + fh//2 + 6*scale],
        radius=r,
        fill=(0, 0, 0, 80)
    )
    shadow = shadow.filter(ImageFilter.GaussianBlur(radius=6*scale))
    img = Image.alpha_composite(img, shadow)

    # --- Tab (Lasche) mit 3D ---
    tab = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    t_draw = ImageDraw.Draw(tab)
    # Dunklere Basis
    t_draw.rounded_rectangle(
        [cx - fw//2, cy - fh//2 - tab_h,
         cx - fw//2 + tab_w, cy - fh//2 + 2*scale],
        radius=r//2,
        fill=(230, 160, 0, 255)
    )
    # Heller Glanz
    t_draw.rounded_rectangle(
        [cx - fw//2 + 2*scale, cy - fh//2 - tab_h + 2*scale,
         cx - fw//2 + tab_w - 2*scale, cy - fh//2 - tab_h//2],
        radius=r//3,
        fill=(255, 210, 80, 100)
    )
    img = Image.alpha_composite(img, tab)

    # --- Ordner-Körper ---
    body = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    b_draw = ImageDraw.Draw(body)

    # Hauptfläche mit Gradient-Simulation
    # Obere Hälfte heller
    b_draw.rounded_rectangle(
        [cx - fw//2, cy - fh//2, cx + fw//2, cy + fh//2],
        radius=r,
        fill=(255, 200, 30, 255)
    )
    # Untere Hälfte dunkler
    b_draw.rounded_rectangle(
        [cx - fw//2, cy, cx + fw//2, cy + fh//2],
        radius=r,
        fill=(235, 175, 10, 255)
    )
    img = Image.alpha_composite(img, body)

    # --- Glanz oben auf dem Ordner ---
    gloss = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    g_draw = ImageDraw.Draw(gloss)
    g_draw.rounded_rectangle(
        [cx - fw//2 + 4*scale, cy - fh//2 + 3*scale,
         cx + fw//2 - 4*scale, cy - fh//2 + fh//3],
        radius=r - 2*scale,
        fill=(255, 255, 255, 60)
    )
    gloss = gloss.filter(ImageFilter.GaussianBlur(radius=3*scale))
    img = Image.alpha_composite(img, gloss)

    # --- Lichtreflexion (Specular) ---
    spec = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    sp_draw = ImageDraw.Draw(spec)
    sp_draw.ellipse(
        [cx - fw//4, cy - fh//2 - 2*scale, cx + fw//4, cy - fh//6],
        fill=(255, 255, 255, 30)
    )
    spec = spec.filter(ImageFilter.GaussianBlur(radius=6*scale))
    img = Image.alpha_composite(img, spec)

    # --- Feiner Rand ---
    edge = Image.new('RGBA', (w, h), (0, 0, 0, 0))
    e_draw = ImageDraw.Draw(edge)
    e_draw.rounded_rectangle(
        [cx - fw//2, cy - fh//2, cx + fw//2, cy + fh//2],
        radius=r,
        outline=(200, 150, 0, 80),
        width=scale
    )
    img = Image.alpha_composite(img, edge)

    # Herunterskalieren
    if hasattr(Image, 'Resampling'):
        img = img.resize((width, height), Image.Resampling.LANCZOS)
    else:
        img = img.resize((width, height), Image.LANCZOS)

    return img

# Desktop Grid (Windows 11 typische Werte bei 100% Skalierung)
DESKTOP_GRID_X = 75  # Horizontaler Abstand
DESKTOP_GRID_Y = 75  # Vertikaler Abstand  
DESKTOP_MARGIN_X = 0  # Linker Rand (kein Offset)
DESKTOP_MARGIN_Y = 0  # Oberer Rand (kein Offset)


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
            # Progman finden
            cls._progman = user32.FindWindowW("Progman", None)
            
            if not cls._progman:
                return None
            
            # Nachricht senden um WorkerW zu erstellen
            result = ctypes.c_ulong()
            user32.SendMessageTimeoutW(
                HWND(cls._progman),
                0x052C,  # Spezielle Nachricht für Desktop
                0, 0,
                0x0000,  # SMTO_NORMAL
                1000,
                ctypes.byref(result)
            )
            
            # WorkerW mit SHELLDLL_DefView finden
            workerw_list = []
            
            def enum_callback(hwnd, lparam):
                shell_view = user32.FindWindowExW(hwnd, HWND(0), "SHELLDLL_DefView", None)
                if shell_view:
                    # Das nächste WorkerW nach diesem ist unser Ziel
                    next_worker = user32.FindWindowExW(HWND(0), hwnd, "WorkerW", None)
                    if next_worker:
                        workerw_list.append(next_worker)
                return True
            
            enum_func = WNDENUMPROC(enum_callback)
            user32.EnumWindows(enum_func, None)
            
            if workerw_list:
                cls._workerw = workerw_list[0]
            else:
                # Fallback: Progman verwenden
                cls._workerw = cls._progman
            
            return cls._workerw
            
        except Exception as e:
            print(f"Fehler beim Finden des Desktop-Fensters: {e}")
            return None
    
    @staticmethod
    def set_parent_to_desktop(hwnd):
        """Setzt ein Fenster als Kind des Desktops"""
        desktop_hwnd = WindowsDesktopAPI.find_desktop_window()
        if desktop_hwnd:
            try:
                # HWND als korrekten Typ
                hwnd = HWND(hwnd) if not isinstance(hwnd, HWND) else hwnd
                desktop_hwnd = HWND(desktop_hwnd) if not isinstance(desktop_hwnd, HWND) else desktop_hwnd
                
                # Als Child des Desktops setzen
                user32.SetParent(hwnd, desktop_hwnd)
                
                # Style anpassen
                style = user32.GetWindowLongPtrW(hwnd, GWL_STYLE)
                style = (style & ~WS_POPUP) | WS_CHILD
                user32.SetWindowLongPtrW(hwnd, GWL_STYLE, LONG_PTR(style))
                
                # Extended Style
                ex_style = user32.GetWindowLongPtrW(hwnd, GWL_EXSTYLE)
                ex_style = ex_style | WS_EX_TOOLWINDOW | WS_EX_NOACTIVATE
                user32.SetWindowLongPtrW(hwnd, GWL_EXSTYLE, LONG_PTR(ex_style))
                
                return True
            except Exception as e:
                print(f"Fehler beim Setzen des Parents: {e}")
        return False
    
    @staticmethod
    def set_file_hidden(filepath, hidden=True):
        """Setzt oder entfernt das Hidden-Attribut einer Datei"""
        try:
            attrs = kernel32.GetFileAttributesW(filepath)
            if attrs == 0xFFFFFFFF:
                print(f"    Warnung: Konnte Attribute nicht lesen für {filepath}")
                return False
            
            if hidden:
                new_attrs = attrs | FILE_ATTRIBUTE_HIDDEN
            else:
                new_attrs = attrs & ~FILE_ATTRIBUTE_HIDDEN
            
            result = kernel32.SetFileAttributesW(filepath, new_attrs)
            if not result:
                print(f"    Warnung: SetFileAttributes fehlgeschlagen für {filepath}")
                return False
            
            return True
        except Exception as e:
            print(f"    Fehler in set_file_hidden: {e}")
            return False
    
    @staticmethod
    def refresh_desktop():
        """Aktualisiert die Desktop-Ansicht"""
        try:
            # Methode 1: SHChangeNotify
            SHCNE_ASSOCCHANGED = 0x08000000
            SHCNE_UPDATEDIR = 0x00001000
            SHCNF_IDLIST = 0x0000
            SHCNF_PATH = 0x0005
            
            ctypes.windll.shell32.SHChangeNotify(SHCNE_ASSOCCHANGED, SHCNF_IDLIST, None, None)
            
            # Methode 2: Desktop-Verzeichnis aktualisieren
            desktop_path = WindowsDesktopAPI.get_desktop_path()
            if desktop_path:
                ctypes.windll.shell32.SHChangeNotify(
                    SHCNE_UPDATEDIR, 
                    SHCNF_PATH, 
                    desktop_path.encode('utf-16-le') + b'\x00\x00',
                    None
                )
        except Exception as e:
            print(f"    Warnung bei refresh_desktop: {e}")
    
    @staticmethod
    def get_desktop_path():
        """Gibt den Desktop-Pfad zurück"""
        try:
            if HAS_SHELL:
                from win32com.shell import shell, shellcon
                return shell.SHGetFolderPath(0, shellcon.CSIDL_DESKTOP, None, 0)
        except:
            pass
        
        # Fallback
        return str(Path.home() / "Desktop")
    
    @staticmethod
    def snap_to_grid(x, y):
        """Rastet Koordinaten auf dem Desktop-Grid ein"""
        grid_x = round((x - DESKTOP_MARGIN_X) / DESKTOP_GRID_X) * DESKTOP_GRID_X + DESKTOP_MARGIN_X
        grid_y = round((y - DESKTOP_MARGIN_Y) / DESKTOP_GRID_Y) * DESKTOP_GRID_Y + DESKTOP_MARGIN_Y
        return max(DESKTOP_MARGIN_X, grid_x), max(DESKTOP_MARGIN_Y, grid_y)
    
    @staticmethod
    def set_window_bottom(hwnd):
        """Setzt Fenster in den Hintergrund"""
        try:
            hwnd = HWND(hwnd) if not isinstance(hwnd, HWND) else hwnd
            user32.SetWindowPos(
                hwnd, HWND_BOTTOM, 0, 0, 0, 0,
                SWP_NOMOVE | SWP_NOSIZE | SWP_NOACTIVATE
            )
        except:
            pass
    
    @staticmethod
    def set_desktop_icon_position(filename, screen_x, screen_y):
        """
        Setzt die Position eines Desktop-Icons über die ListView API.
        Unterstützt Multi-Monitor-Setups.
        
        Args:
            filename: Name der Datei (z.B. "Chrome.lnk")
            screen_x, screen_y: Absolute Bildschirmkoordinaten
        
        Returns:
            True bei Erfolg, False bei Fehler
        """
        try:
            # Desktop ListView-Fenster finden
            progman = user32.FindWindowW("Progman", None)
            if not progman:
                print("  ⚠ Progman nicht gefunden")
                return False
            
            # SHELLDLL_DefView finden (kann unter Progman oder WorkerW sein)
            defview = user32.FindWindowExW(HWND(progman), HWND(0), "SHELLDLL_DefView", None)
            
            if not defview:
                # Unter WorkerW suchen
                def find_defview_callback(hwnd, lparam):
                    shell = user32.FindWindowExW(hwnd, HWND(0), "SHELLDLL_DefView", None)
                    if shell:
                        find_defview_callback.result = shell
                        return False
                    return True
                
                find_defview_callback.result = None
                enum_func = WNDENUMPROC(find_defview_callback)
                user32.EnumWindows(enum_func, None)
                defview = find_defview_callback.result
            
            if not defview:
                print("  ⚠ SHELLDLL_DefView nicht gefunden")
                return False
            
            # SysListView32 finden
            listview = user32.FindWindowExW(HWND(defview), HWND(0), "SysListView32", None)
            if not listview:
                print("  ⚠ SysListView32 nicht gefunden")
                return False
            
            listview_int = int(listview)
            
            # ListView-Position auf dem Bildschirm ermitteln für Multi-Monitor
            listview_rect = RECT()
            user32.GetWindowRect(listview_int, ctypes.byref(listview_rect))
            
            # Absolute Koordinaten in ListView-relative Koordinaten umrechnen
            # Die ListView deckt alle Monitore ab, also müssen wir die Position
            # relativ zum ListView-Ursprung berechnen
            relative_x = int(screen_x) - listview_rect.left
            relative_y = int(screen_y) - listview_rect.top
            
            print(f"  ℹ ListView bei ({listview_rect.left}, {listview_rect.top})")
            print(f"  ℹ Screen-Koordinaten: ({screen_x}, {screen_y})")
            print(f"  ℹ Relative Koordinaten: ({relative_x}, {relative_y})")
            
            # Negative Koordinaten sind bei Multi-Monitor normal (Monitor links)
            # aber die ListView-Koordinaten sollten nicht negativ sein für Icons
            # auf dem aktuellen Monitor
            
            # ListView Konstanten
            LVM_FIRST = 0x1000
            LVM_GETITEMCOUNT = LVM_FIRST + 4
            LVM_SETITEMPOSITION = LVM_FIRST + 15
            LVM_GETITEMTEXTW = LVM_FIRST + 115
            LVM_FINDITEMW = LVM_FIRST + 83
            
            # Prozess-ID des Explorer ermitteln
            pid = ctypes.c_ulong()
            user32.GetWindowThreadProcessId(listview_int, ctypes.byref(pid))
            
            # Prozess öffnen
            PROCESS_ALL_ACCESS = 0x1F0FFF
            process = kernel32.OpenProcess(PROCESS_ALL_ACCESS, False, pid.value)
            if not process:
                print("  ⚠ Konnte Explorer-Prozess nicht öffnen")
                return False
            
            try:
                # Anzahl der Icons
                count = ctypes.windll.user32.SendMessageW(listview_int, LVM_GETITEMCOUNT, 0, 0)
                if count <= 0:
                    print("  ⚠ Keine Icons auf Desktop")
                    return False
                
                # LVFINDINFO Struktur für die Suche
                class LVFINDINFOW(ctypes.Structure):
                    _fields_ = [
                        ("flags", ctypes.c_uint),
                        ("psz", ctypes.c_wchar_p),
                        ("lParam", ctypes.c_void_p),
                        ("pt_x", ctypes.c_int),
                        ("pt_y", ctypes.c_int),
                        ("vkDirection", ctypes.c_uint),
                    ]
                
                # Speicher im Explorer-Prozess allozieren
                MEM_COMMIT = 0x1000
                MEM_RESERVE = 0x2000
                MEM_RELEASE = 0x8000
                PAGE_READWRITE = 0x04
                
                # Dateinamen vorbereiten (ohne .lnk Extension für Suche)
                search_name = Path(filename).stem if filename.endswith('.lnk') else filename
                search_name_with_ext = filename
                
                # Buffer für Text (520 Bytes für Unicode-String)
                buffer_size = 520
                remote_buffer = kernel32.VirtualAllocEx(
                    process, None, buffer_size, 
                    MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
                )
                
                if not remote_buffer:
                    print("  ⚠ Konnte keinen Speicher im Explorer allozieren")
                    return False
                
                try:
                    # LVITEM Struktur
                    class LVITEMW(ctypes.Structure):
                        _fields_ = [
                            ("mask", ctypes.c_uint),
                            ("iItem", ctypes.c_int),
                            ("iSubItem", ctypes.c_int),
                            ("state", ctypes.c_uint),
                            ("stateMask", ctypes.c_uint),
                            ("pszText", ctypes.c_void_p),
                            ("cchTextMax", ctypes.c_int),
                            ("iImage", ctypes.c_int),
                            ("lParam", ctypes.c_void_p),
                            ("iIndent", ctypes.c_int),
                            ("iGroupId", ctypes.c_int),
                            ("cColumns", ctypes.c_uint),
                            ("puColumns", ctypes.c_void_p),
                            ("piColFmt", ctypes.c_void_p),
                            ("iGroup", ctypes.c_int),
                        ]
                    
                    LVIF_TEXT = 0x0001
                    
                    # Struktur für LVITEM im Remote-Prozess
                    lvitem_size = ctypes.sizeof(LVITEMW)
                    remote_lvitem = kernel32.VirtualAllocEx(
                        process, None, lvitem_size + buffer_size,
                        MEM_COMMIT | MEM_RESERVE, PAGE_READWRITE
                    )
                    
                    if not remote_lvitem:
                        print("  ⚠ Konnte LVITEM-Speicher nicht allozieren")
                        return False
                    
                    try:
                        found_index = -1
                        
                        # Durch alle Items iterieren und Namen vergleichen
                        for i in range(count):
                            # LVITEM vorbereiten
                            lvitem = LVITEMW()
                            lvitem.mask = LVIF_TEXT
                            lvitem.iItem = i
                            lvitem.iSubItem = 0
                            lvitem.pszText = remote_lvitem + lvitem_size  # Text-Buffer nach Struktur
                            lvitem.cchTextMax = 260
                            
                            # LVITEM in Remote-Prozess schreiben
                            written = ctypes.c_size_t()
                            kernel32.WriteProcessMemory(
                                process, remote_lvitem, 
                                ctypes.byref(lvitem), lvitem_size,
                                ctypes.byref(written)
                            )
                            
                            # Text abrufen
                            ctypes.windll.user32.SendMessageW(
                                listview_int, LVM_GETITEMTEXTW, i, remote_lvitem
                            )
                            
                            # Text aus Remote-Prozess lesen
                            text_buffer = ctypes.create_unicode_buffer(260)
                            kernel32.ReadProcessMemory(
                                process, remote_lvitem + lvitem_size,
                                text_buffer, 520, ctypes.byref(written)
                            )
                            
                            item_name = text_buffer.value
                            
                            # Vergleichen (mit und ohne Extension)
                            if item_name.lower() == search_name.lower() or \
                               item_name.lower() == search_name_with_ext.lower():
                                found_index = i
                                print(f"  ✓ Icon '{item_name}' gefunden bei Index {i}")
                                break
                        
                        if found_index >= 0:
                            # Position setzen mit LVM_SETITEMPOSITION
                            # Verwende die relativen Koordinaten
                            # MAKELPARAM: y in high word, x in low word
                            # Für Multi-Monitor: Verwende screen_x/screen_y direkt
                            # da die ListView den gesamten virtuellen Desktop abdeckt
                            
                            # Versuche erst mit relativen Koordinaten
                            x_pos = relative_x
                            y_pos = relative_y
                            
                            # Falls relative Koordinaten negativ, verwende absolute
                            if x_pos < 0 or y_pos < 0:
                                x_pos = int(screen_x)
                                y_pos = int(screen_y)
                                print(f"  ℹ Verwende absolute Koordinaten: ({x_pos}, {y_pos})")
                            
                            pos_lparam = (int(y_pos) << 16) | (int(x_pos) & 0xFFFF)
                            
                            result = ctypes.windll.user32.SendMessageW(
                                listview_int, LVM_SETITEMPOSITION, found_index, pos_lparam
                            )
                            
                            if result:
                                print(f"  ✓ Icon-Position gesetzt auf ({x_pos}, {y_pos})")
                                return True
                            else:
                                print(f"  ⚠ LVM_SETITEMPOSITION fehlgeschlagen")
                                print(f"    (Tipp: 'Icons automatisch anordnen' deaktivieren)")
                                return False
                        else:
                            print(f"  ⚠ Icon '{search_name}' nicht auf Desktop gefunden")
                            return False
                            
                    finally:
                        kernel32.VirtualFreeEx(process, remote_lvitem, 0, MEM_RELEASE)
                        
                finally:
                    kernel32.VirtualFreeEx(process, remote_buffer, 0, MEM_RELEASE)
                    
            finally:
                kernel32.CloseHandle(process)
                
        except Exception as e:
            print(f"  ⚠ Fehler beim Setzen der Icon-Position: {e}")
            import traceback
            traceback.print_exc()
            return False


class IconExtractor:
    """Extrahiert echte Windows-Icons aus Dateien"""
    
    ICON_CACHE = {}
    
    @staticmethod
    def get_icon(filepath, size=48):
        """Holt das Icon — zuerst echtes Windows-Icon, dann Fallback"""
        cache_key = f"{filepath}_{size}"
        if cache_key in IconExtractor.ICON_CACHE:
            return IconExtractor.ICON_CACHE[cache_key]
        
        # Versuche echtes Windows-Icon zu extrahieren
        img = IconExtractor.extract_windows_icon(filepath, size)
        
        # Fallback: generiertes Icon
        if not img:
            img = IconExtractor.get_default_icon(filepath, size)
        
        if img:
            IconExtractor.ICON_CACHE[cache_key] = img
        return img
    
    @staticmethod
    def extract_windows_icon(filepath, size=48):
        """Extrahiert das echte Windows-Icon über SHGetFileInfoW"""
        if not HAS_WIN32:
            return None
        try:
            import win32gui, win32ui, win32con, win32api
            from PIL import Image
            
            # SHGetFileInfo — funktioniert mit .lnk, .exe, Ordnern, etc.
            class SHFILEINFOW(ctypes.Structure):
                _fields_ = [
                    ('hIcon', ctypes.c_void_p),
                    ('iIcon', ctypes.c_int),
                    ('dwAttributes', ctypes.c_uint),
                    ('szDisplayName', ctypes.c_wchar * 260),
                    ('szTypeName', ctypes.c_wchar * 80),
                ]
            
            info = SHFILEINFOW()
            SHGFI_ICON = 0x100
            SHGFI_LARGEICON = 0x0
            
            result = ctypes.windll.shell32.SHGetFileInfoW(
                filepath, 0, ctypes.byref(info), ctypes.sizeof(info),
                SHGFI_ICON | SHGFI_LARGEICON
            )
            
            if not result or not info.hIcon:
                return None
            
            hicon = info.hIcon
            
            # Icon-Größe (System-Standard, meist 32x32)
            ico_x = win32api.GetSystemMetrics(win32con.SM_CXICON)
            ico_y = win32api.GetSystemMetrics(win32con.SM_CYICON)
            
            # Device Context + Bitmap erstellen
            hdc_screen = win32gui.GetDC(0)
            hdc = win32ui.CreateDCFromHandle(hdc_screen)
            hdc_mem = hdc.CreateCompatibleDC()
            
            hbmp = win32ui.CreateBitmap()
            hbmp.CreateCompatibleBitmap(hdc, ico_x, ico_y)
            hdc_mem.SelectObject(hbmp)
            
            # Schwarzer Hintergrund
            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0)
            
            # Icon zeichnen
            hdc_mem.DrawIcon((0, 0), hicon)
            
            # Bitmap-Daten auslesen
            bmpstr = hbmp.GetBitmapBits(True)
            img_black = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr, 'raw', 'BGRX', 0, 1)
            
            # Weißer Hintergrund für Alpha-Berechnung
            hdc_mem.FillSolidRect((0, 0, ico_x, ico_y), 0x00FFFFFF)
            hdc_mem.DrawIcon((0, 0), hicon)
            bmpstr2 = hbmp.GetBitmapBits(True)
            img_white = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr2, 'raw', 'BGRX', 0, 1)
            
            # Alpha-Kanal aus Differenz berechnen
            # alpha = 255 - (white_pixel - black_pixel)
            import numpy as np
            black_arr = np.array(img_black, dtype=np.float32)
            white_arr = np.array(img_white, dtype=np.float32)
            
            diff = white_arr - black_arr
            alpha = 255.0 - np.mean(diff, axis=2)
            alpha = np.clip(alpha, 0, 255).astype(np.uint8)
            
            # RGB aus schwarzem Hintergrund mit Alpha-Korrektur
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
            
            # Auf Zielgröße skalieren
            if ico_x != size or ico_y != size:
                resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                img = img.resize((size, size), resample)
            
            # Aufräumen
            win32gui.DestroyIcon(hicon)
            hdc_mem.DeleteDC()
            win32gui.ReleaseDC(0, hdc_screen)
            
            return img
        except Exception as e:
            # Numpy nicht verfügbar — einfacher Fallback ohne Alpha
            try:
                import win32gui, win32ui, win32con, win32api
                from PIL import Image
                
                class SHFILEINFOW2(ctypes.Structure):
                    _fields_ = [
                        ('hIcon', ctypes.c_void_p),
                        ('iIcon', ctypes.c_int),
                        ('dwAttributes', ctypes.c_uint),
                        ('szDisplayName', ctypes.c_wchar * 260),
                        ('szTypeName', ctypes.c_wchar * 80),
                    ]
                
                info = SHFILEINFOW2()
                result = ctypes.windll.shell32.SHGetFileInfoW(
                    filepath, 0, ctypes.byref(info), ctypes.sizeof(info), 0x100
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
                img = Image.frombuffer('RGB', (ico_x, ico_y), bmpstr, 'raw', 'BGRX', 0, 1)
                
                resample = Image.Resampling.LANCZOS if hasattr(Image, 'Resampling') else Image.LANCZOS
                if ico_x != size or ico_y != size:
                    img = img.resize((size, size), resample)
                
                win32gui.DestroyIcon(hicon)
                hdc_mem.DeleteDC()
                win32gui.ReleaseDC(0, hdc_screen)
                
                return img.convert('RGBA')
            except:
                return None
    
    @staticmethod
    def get_default_icon(filepath, size=48):
        """Erstellt ein 3D-Icon mit Licht, Schatten und Glaseffekt"""
        try:
            from PIL import Image, ImageDraw, ImageFont, ImageFilter
        except:
            return None
        
        # Größer rendern für Anti-Aliasing
        scale = 2
        s = size * scale
        
        img = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        ext = Path(filepath).suffix.lower() if filepath else ""
        name = Path(filepath).stem if filepath else ""
        
        # Windows 11 ähnliche Farben pro Dateityp
        type_colors = {
            '.exe': (0, 120, 212), '.msi': (0, 120, 212), '.lnk': (0, 120, 212),
            '.bat': (255, 165, 0), '.cmd': (255, 165, 0), '.ps1': (1, 36, 86),
            '.py': (55, 118, 171), '.txt': (107, 107, 107), '.pdf': (220, 30, 30),
            '.doc': (43, 87, 154), '.docx': (43, 87, 154),
            '.xls': (33, 115, 70), '.xlsx': (33, 115, 70),
            '.ppt': (210, 71, 38), '.pptx': (210, 71, 38),
            '.jpg': (0, 188, 242), '.jpeg': (0, 188, 242), '.png': (0, 188, 242),
            '.mp3': (255, 64, 129), '.mp4': (255, 64, 129),
            '.zip': (255, 215, 0), '.rar': (255, 215, 0), '.7z': (255, 215, 0),
            '.html': (228, 77, 38), '.css': (38, 77, 228), '.js': (247, 223, 30),
        }
        
        cr, cg, cb = type_colors.get(ext, (0, 120, 212))
        
        margin = 4 * scale
        radius = 8 * scale
        
        # --- Drop Shadow ---
        shadow = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        s_draw = ImageDraw.Draw(shadow)
        s_draw.rounded_rectangle(
            [margin + 3*scale, margin + 3*scale, s - margin + 1*scale, s - margin + 1*scale],
            radius=radius,
            fill=(0, 0, 0, 70)
        )
        shadow = shadow.filter(ImageFilter.GaussianBlur(radius=3*scale))
        img = Image.alpha_composite(img, shadow)
        
        # --- Hauptrechteck mit Gradient ---
        main = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        m_draw = ImageDraw.Draw(main)
        # Basis
        m_draw.rounded_rectangle(
            [margin, margin, s - margin, s - margin],
            radius=radius,
            fill=(cr, cg, cb, 240)
        )
        # Dunkler unten
        dark = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        d_draw = ImageDraw.Draw(dark)
        d_draw.rounded_rectangle(
            [margin, s//2, s - margin, s - margin],
            radius=radius,
            fill=(0, 0, 0, 40)
        )
        main = Image.alpha_composite(main, dark)
        img = Image.alpha_composite(img, main)
        
        # --- Glasglanz oben ---
        gloss = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        g_draw = ImageDraw.Draw(gloss)
        g_draw.rounded_rectangle(
            [margin + 2*scale, margin + 2*scale, s - margin - 2*scale, margin + s//3],
            radius=radius - 2*scale,
            fill=(255, 255, 255, 45)
        )
        gloss = gloss.filter(ImageFilter.GaussianBlur(radius=2*scale))
        img = Image.alpha_composite(img, gloss)
        
        # --- Feiner heller Rand oben (3D-Kante) ---
        edge = Image.new('RGBA', (s, s), (0, 0, 0, 0))
        e_draw = ImageDraw.Draw(edge)
        e_draw.rounded_rectangle(
            [margin, margin, s - margin, s - margin],
            radius=radius,
            outline=(255, 255, 255, 50),
            width=scale
        )
        img = Image.alpha_composite(img, edge)
        
        # --- Buchstabe ---
        letter = name[0].upper() if name else "?"
        
        font = None
        font_size = s // 2
        try:
            font = ImageFont.truetype("segoeui.ttf", font_size)
        except:
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                try:
                    font = ImageFont.load_default()
                except:
                    pass
        
        if font:
            # Temporäres Bild für Text
            txt_layer = Image.new('RGBA', (s, s), (0, 0, 0, 0))
            txt_draw = ImageDraw.Draw(txt_layer)
            
            bbox = txt_draw.textbbox((0, 0), letter, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            x = (s - text_width) // 2
            y = (s - text_height) // 2 - 2 * scale
            
            # Text-Schatten
            txt_draw.text((x + scale, y + scale), letter, fill=(0, 0, 0, 80), font=font)
            # Text
            txt_draw.text((x, y), letter, fill=(255, 255, 255, 230), font=font)
            img = Image.alpha_composite(img, txt_layer)
        
        # Herunterskalieren für Anti-Aliasing
        if hasattr(Image, 'Resampling'):
            img = img.resize((size, size), Image.Resampling.LANCZOS)
        else:
            img = img.resize((size, size), Image.LANCZOS)
        
        return img


class FolderTile:
    """Eine einzelne Ordner-Kachel auf dem Desktop"""
    
    def __init__(self, manager, tile_id, config):
        self.manager = manager
        self.tile_id = tile_id
        self.config = config
        
        self.is_expanded = False
        self.animation_running = False
        self.drag_data = {"x": 0, "y": 0, "dragging": False}
        self.icon_images = []
        self.hwnd = None
        self.is_embedded = False
        self._footer_label = None
        self._name_entry = None

        # Zoom-Einstellungen: getrennt für verkleinert/expandiert (10-150%)
        self.collapsed_scale = self.config.get("collapsed_scale", 100)
        self.expanded_scale = self.config.get("expanded_scale", 100)

        # Verknüpfungsnamen in der verkleinerten Ansicht ausblenden (Standard: an)
        self.hide_shortcut_names = self.config.get("hide_shortcut_names", True)

        # Basis-Größen (bei 100%)
        self._base_tile_width = DESKTOP_GRID_X * 2   # 150
        self._base_tile_height = DESKTOP_GRID_Y * 2   # 150
        self._base_expanded_width = 245
        self._base_expanded_height = 280

        # Aktuelle Größen (skaliert)
        self.apply_scale()

        self.create_window()

    def apply_scale(self):
        """Berechnet die skalierten Größen basierend auf collapsed_scale / expanded_scale"""
        sc = max(10, min(150, self.collapsed_scale)) / 100.0
        se = max(10, min(150, self.expanded_scale)) / 100.0
        self.tile_width = max(40, int(self._base_tile_width * sc))
        self.tile_height = max(40, int(self._base_tile_height * sc))
        self.expanded_width = max(80, int(self._base_expanded_width * se))
        self.expanded_height = max(100, int(self._base_expanded_height * se))
    
    def create_window(self):
        """Erstellt das Kachel-Fenster mit Glaseffekt und echt transparenten Ecken"""
        self.window = tk.Toplevel(self.manager.root)
        self.window.title(f"DesktopFolder_{self.tile_id}")
        self.window.overrideredirect(True)
        self.window.attributes("-alpha", 0.92)
        
        # Transparente Farbe für echte durchsichtige Ecken.
        # Nur die winzigen Eck-Dreiecke (außerhalb der abgerundeten 3D-Hintergrund-
        # grafik) haben diese Farbe → nur dort klickdurchlässig.
        # Der gesamte Kachel-Körper wird vom 3D-Bild abgedeckt → Drag&Drop funktioniert.
        self._transparent_color = "#010101"
        self.window.config(bg=self._transparent_color)
        self.window.attributes("-transparentcolor", self._transparent_color)
        
        # Position auf Grid snappen
        x = self.config.get("pos_x", DESKTOP_MARGIN_X + int(self.tile_id) * self.tile_width)
        y = self.config.get("pos_y", DESKTOP_MARGIN_Y)
        x, y = WindowsDesktopAPI.snap_to_grid(x, y)
        
        x = max(10, min(x, 1800))
        y = max(10, min(y, 1000))
        
        self.window.geometry(f"{self.tile_width}x{self.tile_height}+{x}+{y}")
        
        # Hauptframe — transparent für Ecken
        self.main_frame = tk.Frame(
            self.window,
            bg=self._transparent_color,
            highlightthickness=0,
        )
        self.main_frame.pack(fill="both", expand=True)
        
        # Canvas — transparent für Ecken, 3D-Bild deckt Körper ab
        self.canvas = tk.Canvas(
            self.main_frame,
            width=self.tile_width,
            height=self.tile_height,
            bg=self._transparent_color,
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)
        
        # Expanded Frame (anfangs None)
        self.expanded_frame = None
        
        # Hintergrundbild-Cache (normal + hover)
        self._bg_image = None
        self._normal_bg_photo = None
        self._hover_bg_photo = None
        self._is_hovered = False
        
        # Icon zeichnen
        self.draw_tile_icon()
        
        # Bindings
        self.setup_bindings()
        
        # Drag & Drop
        self.setup_drag_drop()
        
        # Fenster sichtbar machen
        self.window.after(100, self.setup_window_mode)
    
    def setup_window_mode(self):
        """Macht das Fenster sichtbar, aktiviert Glaseffekt, runde Ecken und hält es im Hintergrund"""
        try:
            self.window.update_idletasks()
            self.window.deiconify()
            self.window.lift()
            
            # HWND holen
            self.hwnd = user32.GetParent(HWND(self.window.winfo_id()))
            
            # Abgerundete Ecken über SetWindowRgn
            self.apply_rounded_corners()
            
            # Acrylic Blur (Glaseffekt) aktivieren
            # Gradient: AABBGGRR - halbtransparentes dunkles Blau
            blur_color = 0xB0201A0D  # Alpha=0xB0, B=0x20, G=0x1A, R=0x0D
            if enable_acrylic_blur(self.hwnd, blur_color):
                print(f"Kachel {self.tile_id}: Acrylic Blur aktiviert ✓")
                # Bei Blur weniger Fenster-Transparenz nötig
                self.window.attributes("-alpha", 0.98)
            else:
                print(f"Kachel {self.tile_id}: Acrylic Blur nicht verfügbar (Fallback)")
            
            print(f"Kachel {self.tile_id} erstellt bei ({self.window.winfo_x()}, {self.window.winfo_y()})")
            
            # Nach kurzer Verzögerung in Hintergrund
            self.window.after(500, self.move_to_background)
            
        except Exception as e:
            print(f"Fehler: {e}")
            import traceback
            traceback.print_exc()
    
    def apply_rounded_corners(self, width=None, height=None, radius=16):
        """Wendet abgerundete Ecken auf das Fenster an"""
        if not self.hwnd:
            return
        w = width or self.window.winfo_width()
        h = height or self.window.winfo_height()
        set_rounded_region(self.hwnd, w, h, radius)
    
    def move_to_background(self):
        """Bewegt das Fenster in den Hintergrund"""
        try:
            self.window.attributes("-topmost", False)
            if self.hwnd:
                WindowsDesktopAPI.set_window_bottom(self.hwnd)
            
            # Timer um im Hintergrund zu bleiben
            self.keep_background_timer()
        except:
            pass
    
    def keep_background_timer(self):
        """Hält das Fenster regelmäßig im Hintergrund"""
        try:
            if not self.is_expanded and self.hwnd:
                WindowsDesktopAPI.set_window_bottom(self.hwnd)
            
            if hasattr(self, 'window') and self.window.winfo_exists():
                self.window.after(2000, self.keep_background_timer)
        except:
            pass
    
    def embed_in_desktop(self):
        """Bettet das Fenster in den Desktop ein (experimentell)"""
        try:
            self.window.update_idletasks()
            
            # HWND des Fensters holen
            raw_hwnd = user32.GetParent(HWND(self.window.winfo_id()))
            self.hwnd = raw_hwnd
            
            # In Desktop einbetten
            if WindowsDesktopAPI.set_parent_to_desktop(self.hwnd):
                self.is_embedded = True
                print(f"Kachel {self.tile_id} in Desktop eingebettet")
            else:
                print(f"Kachel {self.tile_id}: Einbettung fehlgeschlagen")
                
        except Exception as e:
            print(f"Fehler beim Einbetten: {e}")
            import traceback
            traceback.print_exc()
    
    def setup_bindings(self):
        """Event-Bindings"""
        self.canvas.bind("<ButtonPress-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.do_drag)
        self.canvas.bind("<ButtonRelease-1>", self.stop_drag)
        self.canvas.bind("<Button-3>", self.show_context_menu)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        
        # Hover: Expand/Collapse über Maus-Enter/Leave
        self._hover_expand_timer = None
        self._hover_collapse_timer = None
        self._mouse_inside = False
        
        self.window.bind("<Enter>", self._on_window_enter)
        self.window.bind("<Leave>", self._on_window_leave)
    
    def _on_window_enter(self, event):
        """Maus betritt das Fenster — Hover-Expand starten"""
        self._mouse_inside = True
        
        # Collapse-Timer abbrechen falls aktiv
        if self._hover_collapse_timer:
            self.window.after_cancel(self._hover_collapse_timer)
            self._hover_collapse_timer = None
        
        if self.is_expanded or self.animation_running:
            return
        
        # Visuellen Hover-Effekt sofort zeigen
        self._draw_hover_state(True)
        
        # Expand nach kurzer Verzögerung (150ms — schnell)
        self._hover_expand_timer = self.window.after(150, self._hover_expand)
    
    def _on_window_leave(self, event):
        """Maus verlässt das Fenster — Hover-Collapse starten"""
        try:
            x, y = self.window.winfo_pointerxy()
            wx = self.window.winfo_rootx()
            wy = self.window.winfo_rooty()
            ww = self.window.winfo_width()
            wh = self.window.winfo_height()
            
            if wx <= x <= wx + ww and wy <= y <= wy + wh:
                return
        except:
            pass
        
        self._mouse_inside = False
        
        if self._hover_expand_timer:
            self.window.after_cancel(self._hover_expand_timer)
            self._hover_expand_timer = None
        
        if not self.is_expanded and not self.animation_running:
            self._draw_hover_state(False)
        
        # Collapse nach 350ms
        if self.is_expanded and not self.animation_running:
            self._hover_collapse_timer = self.window.after(350, self._hover_collapse)
    
    def _hover_expand(self):
        """Wird nach Hover-Verzögerung aufgerufen — expandiert die Kachel"""
        self._hover_expand_timer = None
        if not self._mouse_inside or self.is_expanded or self.animation_running:
            return
        self.expand()
    
    def _hover_collapse(self):
        """Wird nach Leave-Verzögerung aufgerufen — klappt zusammen"""
        self._hover_collapse_timer = None
        if self._mouse_inside or not self.is_expanded or self.animation_running:
            return
        self.collapse()
    
    def _draw_hover_state(self, hovered):
        """Tauscht NUR den Hintergrund aus — Icons/Text bleiben unverändert (kein Zittern)"""
        try:
            if hovered:
                # Hover-Hintergrund erzeugen (einmalig cachen)
                if not self._hover_bg_photo:
                    bg_img = create_3d_tile_background(
                        self.tile_width, self.tile_height,
                        base_color=(22, 22, 44),
                        corner_radius=14
                    )
                    if bg_img:
                        from PIL import Image, ImageDraw
                        glow = Image.new('RGBA', bg_img.size, (0, 0, 0, 0))
                        g_draw = ImageDraw.Draw(glow)
                        g_draw.rounded_rectangle(
                            [0, 0, bg_img.width - 1, bg_img.height - 1],
                            radius=14, outline=(120, 140, 255, 60), width=2
                        )
                        g_draw.rounded_rectangle(
                            [1, 1, bg_img.width - 2, bg_img.height - 2],
                            radius=13, outline=(180, 190, 255, 30), width=1
                        )
                        bg_img = Image.alpha_composite(bg_img, glow)
                        self._hover_bg_photo = ImageTk.PhotoImage(bg_img)
                
                if self._hover_bg_photo:
                    self.canvas.delete("bg_layer")
                    self.canvas.create_image(0, 0, anchor="nw", image=self._hover_bg_photo, tags="bg_layer")
                    self.canvas.tag_lower("bg_layer")
            else:
                # Normalen Hintergrund zurücksetzen (gecacht aus draw_tile_icon)
                if self._normal_bg_photo:
                    self.canvas.delete("bg_layer")
                    self.canvas.create_image(0, 0, anchor="nw", image=self._normal_bg_photo, tags="bg_layer")
                    self.canvas.tag_lower("bg_layer")
        except Exception:
            pass
    
    def setup_drag_drop(self):
        """Drag & Drop aktivieren"""
        if HAS_WINDND:
            windnd.hook_dropfiles(self.window, func=self.on_drop_files)
    
    def on_drop_files(self, files):
        """Dateien wurden auf die Kachel gezogen — mit Positionserkennung"""
        desktop_path = WindowsDesktopAPI.get_desktop_path()
        added_count = 0
        
        # Einfüge-Position bestimmen
        insert_index = len(self.config.get("shortcuts", []))  # Standard: am Ende
        
        if self.is_expanded and hasattr(self, 'icons_frame') and self.icons_frame:
            try:
                mx, my = self.window.winfo_pointerxy()
                frame_x = self.icons_frame.winfo_rootx()
                frame_y = self.icons_frame.winfo_rooty()
                rel_x = mx - frame_x
                rel_y = my - frame_y
                
                _s = self.expanded_scale / 100.0
                cols = 3
                cell_width = max(30, int(70 * _s))
                cell_height = max(30, int(65 * _s))
                
                col = max(0, min(int(rel_x // cell_width), cols - 1))
                row = max(0, int(rel_y // cell_height))
                calc_index = row * cols + col
                insert_index = min(calc_index, len(self.config.get("shortcuts", [])))
            except:
                pass
        
        for file_bytes in files:
            try:
                filepath = file_bytes.decode('utf-8')
            except:
                try:
                    filepath = file_bytes.decode('gbk')
                except:
                    continue
            
            if not os.path.exists(filepath):
                continue
            
            already_exists = False
            for shortcut in self.config.get("shortcuts", []):
                if shortcut["path"] == filepath:
                    already_exists = True
                    break
            
            if already_exists:
                continue
            
            name = Path(filepath).stem
            
            if "shortcuts" not in self.config:
                self.config["shortcuts"] = []
            
            # An der berechneten Position einfügen
            self.config["shortcuts"].insert(insert_index, {
                "name": name,
                "path": filepath
            })
            insert_index += 1  # Nächste Datei danach
            
            if Path(filepath).parent == Path(desktop_path):
                WindowsDesktopAPI.set_file_hidden(filepath, True)
                print(f"Desktop-Icon versteckt: {name}")
            
            added_count += 1
        
        if added_count > 0:
            WindowsDesktopAPI.refresh_desktop()
            self.manager.save_config()
            
            # Caches invalidieren (neue Icons)
            self._hover_bg_photo = None
            self._normal_bg_photo = None

            if self.is_expanded:
                # Expandierte Ansicht aktualisieren
                self.refresh_expanded_view()
            else:
                # Kachel-Icon aktualisieren und automatisch expandieren
                self.draw_tile_icon()
                self.canvas.update_idletasks()
                self.window.update()
                self.window.after(200, self.expand)
            
            print(f"{added_count} Verknüpfung(en) hinzugefügt")
    
    def draw_tile_icon(self):
        """Zeichnet das Kachel-Icon mit 3D-Hintergrund, Licht und Schatten"""
        self.canvas.delete("all")
        self.icon_images.clear()

        s = self.collapsed_scale / 100.0
        shortcuts = self.config.get("shortcuts", [])
        width = self.tile_width
        height = self.tile_height
        
        # --- 3D-Hintergrund zeichnen (mit Tag für Hover-Swap) ---
        try:
            bg_img = create_3d_tile_background(width, height, base_color=(13, 13, 26), corner_radius=14)
            if bg_img:
                self._normal_bg_photo = ImageTk.PhotoImage(bg_img)
                self.canvas.create_image(0, 0, anchor="nw", image=self._normal_bg_photo, tags="bg_layer")
        except Exception as e:
            pass
        
        if not shortcuts:
            self.draw_empty_folder(width, height)
        else:
            self.draw_icon_grid(shortcuts[:4], width, height)
        
        # Ordnername unten mit Schatten (skalierte Schriftgröße)
        name = self.config.get("name", "Ordner")
        if len(name) > 12:
            name = name[:11] + "…"

        name_font_size = max(6, int(9 * s))
        self.canvas.create_text(
            width // 2 + 1, height - 9,
            text=name, fill="#000000",
            font=("Segoe UI Semibold", name_font_size), anchor="s"
        )
        self.canvas.create_text(
            width // 2, height - 10,
            text=name, fill="#e0e0e0",
            font=("Segoe UI Semibold", name_font_size), anchor="s"
        )
    
    def draw_empty_folder(self, width, height):
        """Zeichnet 3D-Ordner-Icon mit Licht und Schatten"""
        try:
            folder_img = create_3d_folder_icon(width, height)
            if folder_img:
                self._folder_photo = ImageTk.PhotoImage(folder_img)
                self.canvas.create_image(0, 0, anchor="nw", image=self._folder_photo)
                return
        except:
            pass
        
        # Fallback: einfaches Ordner-Icon
        cx, cy = width // 2, height // 2 - 15
        folder_color = "#FFC107"
        w, h = 80, 60
        tab_w, tab_h = 32, 12
        
        self.canvas.create_rectangle(
            cx - w//2, cy - h//2 - tab_h,
            cx - w//2 + tab_w, cy - h//2,
            fill="#FFA000", outline=""
        )
        self.canvas.create_rectangle(
            cx - w//2, cy - h//2, cx + w//2, cy + h//2,
            fill=folder_color, outline=""
        )
    
    def draw_icon_grid(self, shortcuts, width, height):
        """Zeichnet 2x2 Icon-Grid wie Desktop-Icons"""
        # Verfügbarer Platz — Name-Bereich unten nur abziehen wenn sichtbar
        name_reserve = 0 if self.hide_shortcut_names else 25
        available_height = height - name_reserve

        # Icon-Größe skaliert (Basis 48 bei 100%)
        s = self.collapsed_scale / 100.0
        icon_size = max(16, int(48 * s))
        cell_width = width // 2
        cell_height = available_height // 2

        for i, shortcut in enumerate(shortcuts[:4]):
            row = i // 2
            col = i % 2

            # Zentrierte Position in der Zelle
            cx = col * cell_width + cell_width // 2
            cy = row * cell_height + cell_height // 2

            # Icon laden
            icon_img = None
            try:
                pil_img = IconExtractor.get_icon(shortcut["path"], icon_size)
                if pil_img:
                    if hasattr(Image, 'Resampling'):
                        pil_img = pil_img.resize((icon_size, icon_size), Image.Resampling.LANCZOS)
                    else:
                        pil_img = pil_img.resize((icon_size, icon_size), Image.LANCZOS)
                    icon_img = ImageTk.PhotoImage(pil_img)
                    self.icon_images.append(icon_img)
            except:
                pass

            if icon_img:
                self.canvas.create_image(cx, cy, image=icon_img)
            else:
                # Fallback: Farbiges Rechteck
                ext = Path(shortcut["path"]).suffix.lower()
                colors = {'.exe': '#0078D4', '.lnk': '#0078D4', '.bat': '#FFA500'}
                color = colors.get(ext, '#0078D4')

                self.canvas.create_rectangle(
                    cx - icon_size//2, cy - icon_size//2,
                    cx + icon_size//2, cy + icon_size//2,
                    fill=color, outline=""
                )

                letter = shortcut["name"][0].upper() if shortcut["name"] else "?"
                letter_font_size = max(8, int(16 * s))
                self.canvas.create_text(
                    cx, cy,
                    text=letter,
                    fill="white",
                    font=("Segoe UI", letter_font_size, "bold")
                )

            # Name unter Icon (nur wenn nicht ausgeblendet)
            if not self.hide_shortcut_names:
                name = shortcut["name"]
                if len(name) > 8:
                    name = name[:7] + "…"

                grid_name_font = max(6, int(8 * s))
                self.canvas.create_text(
                    cx, cy + icon_size//2,
                    text=name,
                    fill="white",
                    font=("Segoe UI", grid_name_font),
                    anchor="n"
                )
    
    def on_click(self, event):
        """Klick-Handler — bei collapsed wird manuell expandiert (Fallback)"""
        if self.drag_data.get("dragging"):
            return
        
        if not self.is_expanded:
            self.expand()
    
    def start_drag(self, event):
        """Drag starten"""
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.drag_data["dragging"] = False
        self.drag_data["start_x"] = event.x_root
        self.drag_data["start_y"] = event.y_root
    
    def do_drag(self, event):
        """Drag durchführen"""
        dx = abs(event.x_root - self.drag_data.get("start_x", event.x_root))
        dy = abs(event.y_root - self.drag_data.get("start_y", event.y_root))
        
        if dx > 5 or dy > 5:
            self.drag_data["dragging"] = True
            x = self.window.winfo_x() + (event.x - self.drag_data["x"])
            y = self.window.winfo_y() + (event.y - self.drag_data["y"])
            self.window.geometry(f"+{x}+{y}")
    
    def stop_drag(self, event):
        """Drag beenden - auf Grid einrasten"""
        was_dragging = self.drag_data.get("dragging", False)
        self.drag_data["dragging"] = False
        
        if was_dragging:
            # Auf Grid einrasten
            x = self.window.winfo_x()
            y = self.window.winfo_y()
            snap_x, snap_y = WindowsDesktopAPI.snap_to_grid(x, y)
            
            self.window.geometry(f"+{snap_x}+{snap_y}")
            
            self.config["pos_x"] = snap_x
            self.config["pos_y"] = snap_y
            self.manager.save_config()
        else:
            self.window.after(50, lambda: self.on_click(event))
    
    def expand(self):
        """Kachel expandieren"""
        if self.is_expanded or self.animation_running:
            return
        
        self.animation_running = True
        self.is_expanded = True
        
        # Collapse-Timer abbrechen falls noch aktiv
        if self._hover_collapse_timer:
            self.window.after_cancel(self._hover_collapse_timer)
            self._hover_collapse_timer = None
        
        # Fenster nach vorne bringen
        self.window.attributes("-topmost", True)
        self.window.attributes("-alpha", 0.96)
        self.window.lift()
        self.window.focus_force()
        
        # Fenster temporär aus Desktop lösen wenn eingebettet
        if self.is_embedded and self.hwnd:
            try:
                user32.SetParent(HWND(self.hwnd), HWND(0))
            except:
                pass
        
        x = self.window.winfo_x()
        y = self.window.winfo_y()
        
        self.animate_size(
            self.tile_width, self.tile_height,
            self.expanded_width, self.expanded_height,
            x, y,
            callback=self.show_expanded_content
        )
    
    def show_expanded_content(self):
        """Zeigt Desktop-ähnliche Icon-Ansicht — Titel unten wie collapsed"""
        self.canvas.pack_forget()

        glass_bg = "#0d0d1a"

        self.expanded_frame = tk.Frame(self.main_frame, bg=glass_bg)
        self.expanded_frame.pack(fill="both", expand=True)

        # Acrylic Blur für expandiertes Fenster
        if self.hwnd:
            blur_color = 0xC0281E10
            enable_acrylic_blur(self.hwnd, blur_color)

        # === Icon-Grid Container (kein Header — homogen mit collapsed) ===
        grid_container = tk.Frame(self.expanded_frame, bg=glass_bg)
        grid_container.pack(fill="both", expand=True, padx=5, pady=(8, 0))

        canvas = tk.Canvas(grid_container, bg=glass_bg, highlightthickness=0)
        scrollbar = tk.Scrollbar(grid_container, orient="vertical", command=canvas.yview)

        self.icons_frame = tk.Frame(canvas, bg=glass_bg)
        self.icons_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))

        canvas.create_window((0, 0), window=self.icons_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)

        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")

        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))

        shortcuts = self.config.get("shortcuts", [])

        if not shortcuts:
            empty = tk.Label(
                self.icons_frame,
                text="Leer\n\nDateien vom Desktop\nhierher ziehen",
                font=("Segoe UI", 10), bg=glass_bg, fg="#555570", justify="center",
                wraplength=200
            )
            empty.pack(pady=40)
            empty.bind("<Button-3>", self.show_context_menu)
            empty.bind("<ButtonPress-1>", self._start_bg_drag)
            empty.bind("<B1-Motion>", self._do_bg_drag)
            empty.bind("<ButtonRelease-1>", self._stop_bg_drag)
        else:
            self.create_desktop_icon_grid(shortcuts)

        # === Trennlinie + Titel unten (konsistent mit collapsed) ===
        tk.Frame(self.expanded_frame, bg="#3a3a5a", height=1).pack(fill="x", padx=10, pady=(2, 0))

        name = self.config.get("name", "Ordner")
        self._footer_label = tk.Label(
            self.expanded_frame, text=name,
            font=("Segoe UI Semibold", 9), bg=glass_bg, fg="#e0e0e0",
            anchor="center", cursor="hand2"
        )
        self._footer_label.pack(fill="x", pady=(3, 8))
        self._footer_label.bind("<Button-1>", lambda e: self._start_name_edit())

        # Drag per Klicken-und-Halten auf freie Flächen
        for bg_widget in [self.expanded_frame, grid_container, canvas, self.icons_frame]:
            bg_widget.bind("<ButtonPress-1>", self._start_bg_drag)
            bg_widget.bind("<B1-Motion>", self._do_bg_drag)
            bg_widget.bind("<ButtonRelease-1>", self._stop_bg_drag)

        # Rechtsklick-Kontextmenü auf Hintergrundflächen der expandierten Ansicht
        for bg_widget in [self.expanded_frame, grid_container, canvas, self.icons_frame, self._footer_label]:
            bg_widget.bind("<Button-3>", self.show_context_menu)

        self.animation_running = False
    
    def create_desktop_icon_grid(self, shortcuts):
        """Erstellt Desktop-ähnliches Icon-Grid mit Glas-Hover-Effekten"""
        glass_bg = "#0d0d1a"
        hover_bg = "#1e1e3a"
        active_bg = "#2a2a4a"

        s = self.expanded_scale / 100.0
        cols = 3
        icon_size = max(16, int(36 * s))
        cell_width = max(30, int(70 * s))
        cell_height = max(30, int(65 * s))
        
        for i, shortcut in enumerate(shortcuts):
            row = i // cols
            col = i % cols
            
            # Container wie auf dem Desktop - mit Hover-Rand
            icon_frame = tk.Frame(
                self.icons_frame,
                bg=glass_bg,
                width=cell_width,
                height=cell_height,
                highlightthickness=1,
                highlightbackground=glass_bg,  # Unsichtbar im Normal-Zustand
                highlightcolor=hover_bg,
            )
            icon_frame.grid(row=row, column=col, padx=2, pady=2)
            icon_frame.grid_propagate(False)
            
            # Icon zentriert oben
            icon_container = tk.Frame(icon_frame, bg=glass_bg)
            icon_container.pack(pady=(3, 1))
            
            icon_label = None
            try:
                pil_img = IconExtractor.get_icon(shortcut["path"], icon_size)
                if pil_img:
                    if hasattr(Image, 'Resampling'):
                        pil_img = pil_img.resize((icon_size, icon_size), Image.Resampling.LANCZOS)
                    else:
                        pil_img = pil_img.resize((icon_size, icon_size), Image.LANCZOS)
                    icon_img = ImageTk.PhotoImage(pil_img)
                    self.icon_images.append(icon_img)
                    icon_label = tk.Label(icon_container, image=icon_img, bg=glass_bg)
            except Exception as e:
                print(f"Icon-Fehler für {shortcut['name']}: {e}")
            
            if icon_label:
                icon_label.pack()
            else:
                # Fallback: Canvas-Icon mit Buchstabe
                canvas = tk.Canvas(
                    icon_container, 
                    width=icon_size, 
                    height=icon_size, 
                    bg=glass_bg, 
                    highlightthickness=0
                )
                canvas.pack()
                
                # Abgerundetes Rechteck
                canvas.create_rectangle(
                    2, 2, icon_size-2, icon_size-2,
                    fill="#0078D4", outline="#0078D4"
                )
                
                # Buchstabe
                letter = shortcut["name"][0].upper() if shortcut["name"] else "?"
                exp_letter_font = max(8, int(16 * s))
                canvas.create_text(
                    icon_size//2, icon_size//2,
                    text=letter,
                    fill="white",
                    font=("Segoe UI", exp_letter_font, "bold")
                )
                icon_label = canvas

            # Name (kürzer, keine Umbrüche)
            name = shortcut["name"]
            if len(name) > 10:
                name = name[:9] + "…"

            exp_name_font = max(6, int(8 * s))
            name_label = tk.Label(
                icon_frame,
                text=name,
                font=("Segoe UI", exp_name_font),
                bg=glass_bg,
                fg="#d0d0e0",
                anchor="center"
            )
            name_label.pack()
            
            # Event-Handler mit Drag-Out Funktion und Geister-Fenster
            def make_handlers(idx, path, shortcut_name, frame, name_lbl, icon_lbl):
                drag_data = {
                    'dragging': False, 
                    'start_x': 0, 
                    'start_y': 0,
                    'ghost_window': None,
                    'active': False  # Verhindert Leave-Reset während Drag
                }
                
                def on_enter(e):
                    if not drag_data['active']:
                        frame.config(bg="#1e1e3a", highlightbackground="#4a4a8a")
                        name_lbl.config(bg="#1e1e3a")
                        try:
                            icon_lbl.config(bg="#1e1e3a")
                        except:
                            pass
                
                def on_leave(e):
                    if not drag_data['active']:
                        frame.config(bg="#0d0d1a", highlightbackground="#0d0d1a")
                        name_lbl.config(bg="#0d0d1a")
                        try:
                            icon_lbl.config(bg="#0d0d1a")
                        except:
                            pass
                
                def on_press(e):
                    drag_data['start_x'] = e.x_root
                    drag_data['start_y'] = e.y_root
                    drag_data['dragging'] = False
                    drag_data['active'] = True
                
                def create_ghost_window():
                    """Erstellt ein halbtransparentes Geister-Fenster mit Glaseffekt"""
                    ghost = tk.Toplevel(self.window)
                    ghost.overrideredirect(True)
                    ghost.attributes("-alpha", 0.75)
                    ghost.attributes("-topmost", True)
                    ghost.geometry("60x70")
                    ghost.config(bg="#1a1a3a")
                    
                    # Icon und Name im Geisterfenster
                    ghost_frame = tk.Frame(ghost, bg="#1a1a3a")
                    ghost_frame.pack(fill="both", expand=True, padx=3, pady=3)
                    
                    # Mini-Icon
                    ghost_canvas = tk.Canvas(ghost_frame, width=40, height=40, 
                                            bg="#1a1a3a", highlightthickness=0)
                    ghost_canvas.pack(pady=(2, 0))
                    ghost_canvas.create_rectangle(2, 2, 38, 38, fill="#0078D4", outline="")
                    letter = shortcut_name[0].upper() if shortcut_name else "?"
                    ghost_canvas.create_text(20, 20, text=letter, fill="white", 
                                           font=("Segoe UI", 14, "bold"))
                    
                    # Name
                    short_name = shortcut_name[:8] + "…" if len(shortcut_name) > 8 else shortcut_name
                    tk.Label(ghost_frame, text=short_name, font=("Segoe UI", 7),
                            bg="#1a1a3a", fg="#d0d0e0").pack()
                    
                    return ghost
                
                def on_motion(e):
                    dx = abs(e.x_root - drag_data['start_x'])
                    dy = abs(e.y_root - drag_data['start_y'])
                    
                    if dx > 15 or dy > 15:
                        if not drag_data['dragging']:
                            # Drag beginnt - Geister-Fenster erstellen
                            drag_data['dragging'] = True
                            drag_data['ghost_window'] = create_ghost_window()
                            
                            # Ursprüngliches Icon markieren
                            frame.config(bg="#2a2a4a", highlightbackground="#6a6aaa")
                            name_lbl.config(bg="#2a2a4a")
                            try:
                                icon_lbl.config(bg="#2a2a4a")
                            except:
                                pass
                        
                        # Geister-Fenster folgt der Maus
                        if drag_data['ghost_window']:
                            drag_data['ghost_window'].geometry(
                                f"+{e.x_root - 30}+{e.y_root - 35}"
                            )
                            # Andere Kacheln expandieren wenn Cursor darüber
                            self.expand_tile_under_cursor(e.x_root, e.y_root)
                
                def on_release(e):
                    was_dragging = drag_data['dragging']
                    
                    # Geister-Fenster zerstören
                    if drag_data['ghost_window']:
                        drag_data['ghost_window'].destroy()
                        drag_data['ghost_window'] = None
                    
                    # Farben zurücksetzen
                    frame.config(bg="#0d0d1a", highlightbackground="#0d0d1a")
                    name_lbl.config(bg="#0d0d1a")
                    try:
                        icon_lbl.config(bg="#0d0d1a")
                    except:
                        pass
                    
                    drag_data['dragging'] = False
                    drag_data['active'] = False
                    
                    if was_dragging:
                        # Drag beendet - auf Desktop an Mausposition wiederherstellen
                        drop_x = e.x_root
                        drop_y = e.y_root
                        print(f"Drag-Out erkannt für: {shortcut_name} an Position ({drop_x}, {drop_y})")
                        self.restore_to_desktop_at_position(idx, drop_x, drop_y)
                    else:
                        # Normaler Klick - Programm starten
                        self.launch_shortcut(path)
                
                def on_right_click(e):
                    self.show_item_context_menu(e, idx, path)
                
                return on_enter, on_leave, on_press, on_motion, on_release, on_right_click
            
            enter, leave, press, motion, release, right_click = make_handlers(
                i, shortcut["path"], shortcut["name"], icon_frame, name_label, icon_label
            )
            
            # Bindings für alle Elemente
            for widget in [icon_frame, name_label, icon_container, icon_label]:
                try:
                    widget.bind("<Enter>", enter)
                    widget.bind("<Leave>", leave)
                    widget.bind("<ButtonPress-1>", press)
                    widget.bind("<B1-Motion>", motion)
                    widget.bind("<ButtonRelease-1>", release)
                    widget.bind("<Button-3>", right_click)
                    widget.config(cursor="hand2")
                except:
                    pass
    
    def show_item_context_menu(self, event, index, path):
        """Kontextmenü für einzelnes Item"""
        menu = tk.Menu(self.window, tearoff=0, bg="#12122a", fg="#d0d0e0",
                       activebackground="#2a2a5a", activeforeground="white",
                       relief="flat", bd=0)
        
        menu.add_command(label="▶️ Öffnen", command=lambda: self.launch_shortcut(path))
        menu.add_command(label="📤 Auf Desktop wiederherstellen",
                        command=lambda: self.restore_to_desktop(index))
        menu.add_separator()
        menu.add_command(label="🗑️ Aus Ordner entfernen",
                        command=lambda: self.remove_shortcut(index))
        
        menu.tk_popup(event.x_root, event.y_root)
    
    def restore_to_desktop(self, index):
        """Stellt Verknüpfung auf Desktop wieder her (macht sie sichtbar)"""
        self.restore_to_desktop_at_position(index, None, None)
    
    def restore_to_desktop_at_position(self, index, drop_x=None, drop_y=None):
        """Stellt Verknüpfung auf Desktop an bestimmter Position wieder her"""
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            filepath = shortcuts[index]["path"]
            name = shortcuts[index]["name"]
            
            print(f"Stelle wieder her: {name}")
            
            # Hidden-Attribut entfernen
            result = WindowsDesktopAPI.set_file_hidden(filepath, False)
            if result:
                print(f"  ✓ Hidden-Attribut entfernt")
            else:
                print(f"  ✗ Fehler beim Entfernen des Hidden-Attributs")
            
            # Versuche Icon-Position auf dem Desktop zu setzen
            if drop_x is not None and drop_y is not None:
                print(f"  → Drop-Position: ({drop_x}, {drop_y})")
                
                # Versuche Icon-Position zu setzen (Windows ListView API)
                # Übergebe die Raw Screen-Koordinaten - die Funktion handhabt Multi-Monitor
                success = WindowsDesktopAPI.set_desktop_icon_position(
                    Path(filepath).name, drop_x, drop_y
                )
                if success:
                    print(f"  ✓ Icon-Position gesetzt")
                else:
                    print(f"  ⚠ Icon-Position konnte nicht gesetzt werden (Windows-Einschränkung)")
            
            WindowsDesktopAPI.refresh_desktop()
            
            # Aus Kachel entfernen
            del self.config["shortcuts"][index]
            self.manager.save_config()
            
            # UI sofort aktualisieren
            self.refresh_expanded_view()
            
            print(f"  ✓ Auf Desktop wiederhergestellt: {name}")
    
    def refresh_expanded_view(self):
        """Aktualisiert die expandierte Ansicht"""
        if not self.is_expanded:
            self.draw_tile_icon()
            return
        
        # icons_frame leeren und neu aufbauen
        if hasattr(self, 'icons_frame') and self.icons_frame:
            # Alte Widgets löschen
            for widget in self.icons_frame.winfo_children():
                widget.destroy()
            
            # Icon-Images leeren um Speicher freizugeben
            self.icon_images.clear()
            
            shortcuts = self.config.get("shortcuts", [])
            if not shortcuts:
                empty = tk.Label(
                    self.icons_frame,
                    text="Leer\n\nDateien vom Desktop\nhierher ziehen\noder Icon nach außen\nziehen zum Wiederherstellen",
                    font=("Segoe UI", 10), bg="#0d0d1a", fg="#555570", justify="center",
                    wraplength=200
                )
                empty.pack(pady=50)
            else:
                self.create_desktop_icon_grid(shortcuts)
            
            # Tkinter zwingen, sofort zu aktualisieren
            self.icons_frame.update_idletasks()
            self.window.update()
        
        # Auch das Kachel-Icon aktualisieren
        self.draw_tile_icon()
    
    def start_header_drag(self, event):
        """Start des Header-Drags (für Verschieben oder Klick-zum-Schließen)"""
        self.drag_data["x"] = event.x
        self.drag_data["y"] = event.y
        self.drag_data["dragging"] = False
        self.drag_data["start_x"] = event.x_root
        self.drag_data["start_y"] = event.y_root
    
    def do_header_drag(self, event):
        """Führt Header-Drag durch (Fenster verschieben)"""
        dx = abs(event.x_root - self.drag_data.get("start_x", event.x_root))
        dy = abs(event.y_root - self.drag_data.get("start_y", event.y_root))
        
        if dx > 5 or dy > 5:
            self.drag_data["dragging"] = True
            x = self.window.winfo_x() + (event.x - self.drag_data["x"])
            y = self.window.winfo_y() + (event.y - self.drag_data["y"])
            self.window.geometry(f"+{x}+{y}")
    
    def stop_header_drag(self, event):
        """Beendet Header-Drag - bei Klick (kein Drag) wird geschlossen"""
        was_dragging = self.drag_data.get("dragging", False)
        self.drag_data["dragging"] = False
        
        if was_dragging:
            # Es wurde gedraggt - Position speichern
            x = self.window.winfo_x()
            y = self.window.winfo_y()
            snap_x, snap_y = WindowsDesktopAPI.snap_to_grid(x, y)
            self.window.geometry(f"+{snap_x}+{snap_y}")
            self.config["pos_x"] = snap_x
            self.config["pos_y"] = snap_y
            self.manager.save_config()
        else:
            # Es war ein Klick - Kachel schließen
            self.collapse()
    
    def stop_drag_header(self, event):
        """Alias für Kompatibilität"""
        self.stop_header_drag(event)

    # === Drag auf freie Fläche der expandierten Kachel ===

    def _start_bg_drag(self, event):
        """Drag starten bei Klick auf freie Fläche (expandiert)"""
        self.drag_data["x"] = event.x_root - self.window.winfo_x()
        self.drag_data["y"] = event.y_root - self.window.winfo_y()
        self.drag_data["dragging"] = False
        self.drag_data["start_x"] = event.x_root
        self.drag_data["start_y"] = event.y_root

    def _do_bg_drag(self, event):
        """Drag durchführen auf freie Fläche (expandiert)"""
        dx = abs(event.x_root - self.drag_data.get("start_x", event.x_root))
        dy = abs(event.y_root - self.drag_data.get("start_y", event.y_root))

        if dx > 5 or dy > 5:
            self.drag_data["dragging"] = True
            x = event.x_root - self.drag_data["x"]
            y = event.y_root - self.drag_data["y"]
            self.window.geometry(f"+{x}+{y}")

    def _stop_bg_drag(self, event):
        """Drag beenden auf freie Fläche — nur Position speichern, kein Collapse"""
        was_dragging = self.drag_data.get("dragging", False)
        self.drag_data["dragging"] = False

        if was_dragging:
            x = self.window.winfo_x()
            y = self.window.winfo_y()
            snap_x, snap_y = WindowsDesktopAPI.snap_to_grid(x, y)
            self.window.geometry(f"+{snap_x}+{snap_y}")
            self.config["pos_x"] = snap_x
            self.config["pos_y"] = snap_y
            self.manager.save_config()

    # === Inline-Umbenennung per Klick auf den Namen ===

    def _start_name_edit(self):
        """Ersetzt das Footer-Label durch ein Eingabefeld zum Umbenennen"""
        if not hasattr(self, '_footer_label') or not self._footer_label:
            return

        glass_bg = "#0d0d1a"
        current_name = self.config.get("name", "Ordner")

        # Label verstecken
        self._footer_label.pack_forget()

        # Entry-Widget an gleicher Stelle einfügen
        self._name_entry = tk.Entry(
            self.expanded_frame,
            font=("Segoe UI Semibold", 9),
            bg="#1a1a34", fg="#e0e0e0",
            insertbackground="#e0e0e0",
            relief="flat", justify="center",
            highlightthickness=1, highlightcolor="#5566aa",
            highlightbackground="#3a3a5a"
        )
        self._name_entry.insert(0, current_name)
        self._name_entry.pack(fill="x", padx=10, pady=(3, 8))
        self._name_entry.select_range(0, "end")
        self._name_entry.focus_set()

        self._name_entry.bind("<Return>", lambda e: self._finish_name_edit())
        self._name_entry.bind("<Escape>", lambda e: self._cancel_name_edit())
        self._name_entry.bind("<FocusOut>", lambda e: self._finish_name_edit())

    def _finish_name_edit(self):
        """Übernimmt den neuen Namen aus dem Eingabefeld"""
        if not hasattr(self, '_name_entry') or not self._name_entry:
            return

        new_name = self._name_entry.get().strip()
        if new_name:
            self.config["name"] = new_name
            self.manager.save_config()

        self._name_entry.destroy()
        self._name_entry = None

        # Footer-Label mit neuem Namen wiederherstellen
        glass_bg = "#0d0d1a"
        display_name = self.config.get("name", "Ordner")
        self._footer_label = tk.Label(
            self.expanded_frame, text=display_name,
            font=("Segoe UI Semibold", 9), bg=glass_bg, fg="#e0e0e0",
            anchor="center", cursor="hand2"
        )
        self._footer_label.pack(fill="x", pady=(3, 8))
        self._footer_label.bind("<Button-1>", lambda e: self._start_name_edit())
        self._footer_label.bind("<Button-3>", self.show_context_menu)

        # Auch das minimierte Icon aktualisieren
        self.draw_tile_icon()

    def _cancel_name_edit(self):
        """Bricht die Umbenennung ab, stellt altes Label wieder her"""
        if hasattr(self, '_name_entry') and self._name_entry:
            self._name_entry.destroy()
            self._name_entry = None

        glass_bg = "#0d0d1a"
        name = self.config.get("name", "Ordner")
        self._footer_label = tk.Label(
            self.expanded_frame, text=name,
            font=("Segoe UI Semibold", 9), bg=glass_bg, fg="#e0e0e0",
            anchor="center", cursor="hand2"
        )
        self._footer_label.pack(fill="x", pady=(3, 8))
        self._footer_label.bind("<Button-1>", lambda e: self._start_name_edit())
        self._footer_label.bind("<Button-3>", self.show_context_menu)

    def launch_shortcut(self, path):
        """Startet Verknüpfung"""
        try:
            os.startfile(path)
        except Exception as e:
            messagebox.showerror("Fehler", f"Konnte nicht öffnen:\n{e}")
    
    def remove_shortcut(self, index):
        """Entfernt Verknüpfung (ohne wiederherzustellen)"""
        shortcuts = self.config.get("shortcuts", [])
        if 0 <= index < len(shortcuts):
            del self.config["shortcuts"][index]
            self.manager.save_config()
            self.refresh_expanded_view()
    
    def add_shortcut_dialog(self):
        """Dialog zum Hinzufügen"""
        filepath = filedialog.askopenfilename(
            title="Verknüpfung auswählen",
            filetypes=[("Alle Dateien", "*.*"), ("Programme", "*.exe"), ("Verknüpfungen", "*.lnk")]
        )
        
        if filepath:
            name = Path(filepath).stem
            
            if "shortcuts" not in self.config:
                self.config["shortcuts"] = []
            
            # Prüfen ob schon vorhanden
            for s in self.config["shortcuts"]:
                if s["path"] == filepath:
                    return
            
            self.config["shortcuts"].append({"name": name, "path": filepath})
            
            # Falls auf Desktop, verstecken
            desktop = WindowsDesktopAPI.get_desktop_path()
            if Path(filepath).parent == Path(desktop):
                WindowsDesktopAPI.set_file_hidden(filepath, True)
            
            WindowsDesktopAPI.refresh_desktop()
            self.manager.save_config()
            self.collapse()
            self.draw_tile_icon()
            self.window.after(100, self.expand)
    
    def collapse(self):
        """Kachel zusammenklappen"""
        if not self.is_expanded or self.animation_running:
            return
        
        self.animation_running = True
        self.is_expanded = False
        
        if self.expanded_frame:
            self.expanded_frame.destroy()
            self.expanded_frame = None
        
        self.canvas.pack(fill="both", expand=True)
        
        x = self.window.winfo_x()
        y = self.window.winfo_y()
        snap_x, snap_y = WindowsDesktopAPI.snap_to_grid(x, y)
        
        self.animate_size(
            self.expanded_width, self.expanded_height,
            self.tile_width, self.tile_height,
            snap_x, snap_y,
            callback=self.finish_collapse
        )
    
    def finish_collapse(self):
        """Collapse abschließen"""
        # Hover-Cache invalidieren (Shortcuts können sich geändert haben)
        self._hover_bg_photo = None
        
        self.draw_tile_icon()
        self.animation_running = False
        
        # Abgerundete Ecken für die kleine Kachel
        self.apply_rounded_corners(self.tile_width, self.tile_height)
        
        # Blur für collapsed-Zustand erneut aktivieren
        if self.hwnd:
            blur_color = 0xB0201A0D
            enable_acrylic_blur(self.hwnd, blur_color)
        
        # Zurück in den Hintergrund
        self.window.attributes("-topmost", False)
        self.window.attributes("-alpha", 0.92)
        self.window.after(100, self.move_to_background)
    
    def animate_size(self, from_w, from_h, to_w, to_h, x, y, callback=None):
        """Größenanimation mit Ease-Out-Kurve für flüssigen 3D-Effekt"""
        steps = 12
        step_time = 10
        
        import math
        
        def ease_out(t):
            """Cubic ease-out für natürliche Verzögerung"""
            return 1 - (1 - t) ** 3
        
        def step(i):
            if i <= steps:
                t = ease_out(i / steps)
                new_w = int(from_w + (to_w - from_w) * t)
                new_h = int(from_h + (to_h - from_h) * t)
                self.window.geometry(f"{new_w}x{new_h}+{x}+{y}")
                # Abgerundete Ecken bei jedem Resize-Schritt aktualisieren
                self.apply_rounded_corners(new_w, new_h)
                self.window.after(step_time, lambda: step(i + 1))
            else:
                if callback:
                    callback()
        
        step(1)
    
    def expand_tile_under_cursor(self, mx, my):
        """Expandiert eine andere Kachel, wenn der Cursor während eines Drags darüber ist"""
        for tile_id, tile in self.manager.tiles.items():
            if tile is self:
                continue
            if tile.is_expanded or tile.animation_running:
                continue
            try:
                tx = tile.window.winfo_rootx()
                ty = tile.window.winfo_rooty()
                tw = tile.window.winfo_width()
                th = tile.window.winfo_height()
                if tx <= mx <= tx + tw and ty <= my <= ty + th:
                    tile.expand()
                    return
            except:
                pass

    def show_context_menu(self, event):
        """Kontextmenü der Kachel (collapsed und expanded)"""
        menu = tk.Menu(self.window, tearoff=0, bg="#12122a", fg="#d0d0e0",
                       activebackground="#2a2a5a", activeforeground="white",
                       relief="flat", bd=0)

        # Toggle für Verknüpfungsnamen in verkleinerter Ansicht
        names_label = "✅ Icon-Namen ausblenden" if self.hide_shortcut_names else "⬜ Icon-Namen ausblenden"

        if self.is_expanded:
            # === Kontextmenü für expandierte Kachel ===
            menu.add_command(label="📏 Größe anpassen…", command=self.show_size_dialog)
            menu.add_separator()
            menu.add_command(label="✏️ Umbenennen", command=self.rename)
            menu.add_command(label=names_label, command=self._toggle_hide_shortcut_names)
            menu.add_separator()
            menu.add_command(label="📤 Alle wiederherstellen", command=self.restore_all_to_desktop)
            menu.add_command(label="🗑️ Kachel löschen", command=self.delete_tile)
            menu.add_separator()
            menu.add_command(label="❌ Widget beenden", command=self.manager.quit)
        else:
            # === Kontextmenü für eingeklappte Kachel ===
            menu.add_command(label="📂 Öffnen", command=self.expand)
            menu.add_command(label="✏️ Umbenennen", command=self.rename)
            menu.add_separator()
            menu.add_command(label="📏 Größe anpassen…", command=self.show_size_dialog)
            menu.add_command(label=names_label, command=self._toggle_hide_shortcut_names)
            menu.add_separator()
            menu.add_command(label="🆕 Neue Kachel", command=self.manager.create_new_tile)
            menu.add_separator()
            menu.add_command(label="📤 Alle wiederherstellen", command=self.restore_all_to_desktop)
            menu.add_command(label="🗑️ Kachel löschen", command=self.delete_tile)
            menu.add_separator()
            menu.add_command(label="❌ Widget beenden", command=self.manager.quit)

        menu.tk_popup(event.x_root, event.y_root)

    def _toggle_hide_shortcut_names(self):
        """Schaltet das Ausblenden der Verknüpfungsnamen in der verkleinerten Ansicht um"""
        self.hide_shortcut_names = not self.hide_shortcut_names
        self.config["hide_shortcut_names"] = self.hide_shortcut_names
        self.manager.save_config()
        self._normal_bg_photo = None
        self._hover_bg_photo = None
        self.draw_tile_icon()

    def show_size_dialog(self):
        """Öffnet Slider-Dialog zur Größeneinstellung (verkleinert + expandiert getrennt)"""
        dlg = tk.Toplevel(self.window)
        dlg.title("Kachelgröße")
        dlg.overrideredirect(True)
        dlg.attributes("-topmost", True)
        dlg.config(bg="#12122a")

        # Zentriert neben der Kachel positionieren
        dlg_w, dlg_h = 260, 240
        wx = self.window.winfo_x() + self.window.winfo_width() + 8
        wy = self.window.winfo_y()
        dlg.geometry(f"{dlg_w}x{dlg_h}+{wx}+{wy}")

        style_fg = "#d0d0e0"
        style_bg = "#12122a"
        slider_trough = "#2a2a5a"
        slider_fg = "#6a6aff"

        # --- Titel ---
        tk.Label(
            dlg, text="Kachelgröße", font=("Segoe UI Semibold", 11),
            bg=style_bg, fg=style_fg
        ).pack(pady=(10, 6))

        tk.Frame(dlg, bg="#3a3a5a", height=1).pack(fill="x", padx=15)

        # --- Slider: Verkleinert ---
        frame_c = tk.Frame(dlg, bg=style_bg)
        frame_c.pack(fill="x", padx=18, pady=(10, 2))

        collapsed_label = tk.Label(
            frame_c, text=f"Verkleinert: {self.collapsed_scale}%",
            font=("Segoe UI", 9), bg=style_bg, fg=style_fg, anchor="w"
        )
        collapsed_label.pack(fill="x")

        collapsed_var = tk.IntVar(value=self.collapsed_scale)
        collapsed_slider = tk.Scale(
            frame_c, from_=10, to=150, orient="horizontal",
            variable=collapsed_var, showvalue=False,
            bg=style_bg, fg=style_fg, troughcolor=slider_trough,
            highlightthickness=0, bd=0, sliderrelief="flat",
            activebackground=slider_fg, length=220
        )
        collapsed_slider.pack(fill="x")

        # --- Slider: Expandiert ---
        frame_e = tk.Frame(dlg, bg=style_bg)
        frame_e.pack(fill="x", padx=18, pady=(6, 2))

        expanded_label = tk.Label(
            frame_e, text=f"Expandiert: {self.expanded_scale}%",
            font=("Segoe UI", 9), bg=style_bg, fg=style_fg, anchor="w"
        )
        expanded_label.pack(fill="x")

        expanded_var = tk.IntVar(value=self.expanded_scale)
        expanded_slider = tk.Scale(
            frame_e, from_=10, to=150, orient="horizontal",
            variable=expanded_var, showvalue=False,
            bg=style_bg, fg=style_fg, troughcolor=slider_trough,
            highlightthickness=0, bd=0, sliderrelief="flat",
            activebackground=slider_fg, length=220
        )
        expanded_slider.pack(fill="x")

        # --- Checkbox: Icon-Namen ausblenden ---
        tk.Frame(dlg, bg="#3a3a5a", height=1).pack(fill="x", padx=15, pady=(8, 0))

        hide_names_var = tk.BooleanVar(value=self.hide_shortcut_names)
        hide_names_cb = tk.Checkbutton(
            dlg, text="Icon-Namen ausblenden",
            variable=hide_names_var,
            font=("Segoe UI", 9), bg=style_bg, fg=style_fg,
            activebackground=style_bg, activeforeground=style_fg,
            selectcolor="#2a2a5a", highlightthickness=0, bd=0
        )
        hide_names_cb.pack(padx=18, pady=(6, 2), anchor="w")

        def on_hide_names_change(*_):
            self.hide_shortcut_names = hide_names_var.get()
            self.config["hide_shortcut_names"] = self.hide_shortcut_names
            self.manager.save_config()
            self._normal_bg_photo = None
            self._hover_bg_photo = None
            self.draw_tile_icon()

        hide_names_var.trace_add("write", on_hide_names_change)

        # --- Live-Update bei Slider-Änderung ---
        def on_collapsed_change(*_):
            val = collapsed_var.get()
            collapsed_label.config(text=f"Verkleinert: {val}%")
            self._apply_collapsed_scale(val)

        def on_expanded_change(*_):
            val = expanded_var.get()
            expanded_label.config(text=f"Expandiert: {val}%")
            self._apply_expanded_scale(val)

        collapsed_var.trace_add("write", on_collapsed_change)
        expanded_var.trace_add("write", on_expanded_change)

        # --- Schließen bei Klick außerhalb ---
        def on_focus_out(e):
            try:
                # Prüfe ob der Fokus noch im Dialog ist
                focused = dlg.focus_get()
                if focused and str(focused).startswith(str(dlg)):
                    return
            except:
                pass
            dlg.destroy()

        dlg.bind("<FocusOut>", on_focus_out)
        dlg.bind("<Escape>", lambda e: dlg.destroy())
        dlg.focus_force()

    def _apply_collapsed_scale(self, val):
        """Wendet den Verkleinert-Zoom live an"""
        val = max(10, min(150, val))
        self.collapsed_scale = val
        self.config["collapsed_scale"] = val
        self.manager.save_config()

        # Caches invalidieren
        self._normal_bg_photo = None
        self._hover_bg_photo = None

        self.apply_scale()

        # Canvas-Größe und Icon immer aktualisieren (auch wenn expandiert)
        self.canvas.config(width=self.tile_width, height=self.tile_height)
        self.draw_tile_icon()

        if not self.is_expanded:
            # Collapsed: sofort Fenster anpassen
            x = self.window.winfo_x()
            y = self.window.winfo_y()
            self.window.geometry(f"{self.tile_width}x{self.tile_height}+{x}+{y}")
            self.apply_rounded_corners(self.tile_width, self.tile_height)
            if self.hwnd:
                enable_acrylic_blur(self.hwnd, 0xB0201A0D)

    def _apply_expanded_scale(self, val):
        """Wendet den Expandiert-Zoom live an"""
        val = max(10, min(150, val))
        self.expanded_scale = val
        self.config["expanded_scale"] = val
        self.manager.save_config()

        self.apply_scale()

        if self.is_expanded:
            # Expandiert: Fenster anpassen und Inhalt neu aufbauen
            x = self.window.winfo_x()
            y = self.window.winfo_y()
            self.window.geometry(f"{self.expanded_width}x{self.expanded_height}+{x}+{y}")
            self.apply_rounded_corners(self.expanded_width, self.expanded_height)
            # Icon-Grid im expanded frame neu aufbauen
            self.refresh_expanded_view()
    
    def restore_all_to_desktop(self):
        """Stellt alle Verknüpfungen auf Desktop wieder her"""
        for shortcut in self.config.get("shortcuts", []):
            WindowsDesktopAPI.set_file_hidden(shortcut["path"], False)
        
        WindowsDesktopAPI.refresh_desktop()
        self.config["shortcuts"] = []
        self.manager.save_config()
        self.draw_tile_icon()
    
    def rename(self):
        """Kachel umbenennen"""
        new_name = simpledialog.askstring(
            "Umbenennen", "Neuer Name:",
            initialvalue=self.config.get("name", "Ordner"),
            parent=self.window
        )
        if new_name:
            self.config["name"] = new_name
            self.manager.save_config()
            self.draw_tile_icon()
    
    def delete_tile(self):
        """Kachel löschen"""
        if messagebox.askyesno("Löschen",
            "Kachel löschen?\n\nVerknüpfungen werden auf dem Desktop wiederhergestellt."):
            self.restore_all_to_desktop()
            self.manager.delete_tile(self.tile_id)
    
    def close(self):
        """Fenster schließen"""
        self.window.destroy()


class DesktopFolderManager:
    """Verwaltet alle Ordner-Kacheln"""
    
    CONFIG_FILE = Path.home() / ".desktop_folder_widget_v3.json"
    
    def __init__(self):
        self.root = tk.Tk()
        self.root.withdraw()
        
        self.tiles = {}
        self.config = self.load_config()
        
        # Erste Kachel erstellen falls keine vorhanden
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

        self.check_dependencies()

        # Drag-Erkennung: Periodisch prüfen ob Maus mit gedrückter Taste über Kachel ist
        self._drag_expand_timer = None
        self.start_drag_detection()
    
    def check_dependencies(self):
        """Prüft Abhängigkeiten"""
        missing = []
        if not HAS_WINDND:
            missing.append("windnd")
        if not HAS_WIN32:
            missing.append("pywin32 + Pillow")
        if not HAS_SHELL:
            missing.append("pywin32 (Shell)")

        if missing:
            print(f"\n⚠️ Fehlende Abhängigkeiten: {', '.join(missing)}")
            print("   pip install windnd pywin32 Pillow\n")

    def start_drag_detection(self):
        """Startet periodische Drag-Erkennung für externe Datei-Drags"""
        self._check_drag_over_tiles()

    def _check_drag_over_tiles(self):
        """Prüft ob die Maus mit gedrückter Taste über einer eingeklappten Kachel ist"""
        try:
            # Linke Maustaste gedrückt? (GetAsyncKeyState Bit 0x8000 = gedrückt)
            lmb_state = user32.GetAsyncKeyState(0x01)  # VK_LBUTTON
            if lmb_state & 0x8000:
                # Mausposition holen
                point = wintypes.POINT()
                user32.GetCursorPos(ctypes.byref(point))
                mx, my = point.x, point.y

                # Prüfe ob irgendeine Kachel gerade gedraggt wird (Tile-Verschiebung)
                any_dragging = any(
                    t.drag_data.get("dragging", False) for t in self.tiles.values()
                )
                if not any_dragging:
                    # Prüfe alle Kacheln
                    for tile_id, tile in self.tiles.items():
                        if tile.is_expanded or tile.animation_running:
                            continue
                        try:
                            tx = tile.window.winfo_rootx()
                            ty = tile.window.winfo_rooty()
                            tw = tile.window.winfo_width()
                            th = tile.window.winfo_height()
                            if tx <= mx <= tx + tw and ty <= my <= ty + th:
                                tile.expand()
                                break
                        except:
                            pass
        except:
            pass

        # Alle 200ms erneut prüfen
        self._drag_expand_timer = self.root.after(200, self._check_drag_over_tiles)
    
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
        
        # Position: Neben der letzten Kachel
        last = list(self.tiles.values())[-1] if self.tiles else None
        if last:
            x = last.window.winfo_x() + DESKTOP_GRID_X
            y = last.window.winfo_y()
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
        
        # Alle versteckten Dateien wiederherstellen
        restored_count = 0
        for tile_id, tile_config in self.config.get("tiles", {}).items():
            for shortcut in tile_config.get("shortcuts", []):
                filepath = shortcut.get("path", "")
                name = shortcut.get("name", "Unbekannt")
                
                if filepath and os.path.exists(filepath):
                    try:
                        # Hidden-Attribut entfernen
                        result = WindowsDesktopAPI.set_file_hidden(filepath, False)
                        if result:
                            print(f"  ✓ Wiederhergestellt: {name}")
                            restored_count += 1
                        else:
                            print(f"  ✗ Fehler bei: {name}")
                    except Exception as e:
                        print(f"  ✗ Fehler bei {name}: {e}")
                else:
                    print(f"  ? Datei nicht gefunden: {name}")
        
        # Desktop aktualisieren
        WindowsDesktopAPI.refresh_desktop()
        
        print(f"\n{restored_count} Icons wiederhergestellt.")
        print("=" * 50)
        
        self.save_config()
        
        for tile in list(self.tiles.values()):
            tile.close()
        
        self.root.quit()
    
    def run(self):
        """Hauptschleife"""
        self.root.mainloop()


# Globale Variable für Cleanup
_app_instance = None
_cleanup_done = False


def cleanup_on_exit():
    """Wird beim Beenden aufgerufen - stellt alle Icons wieder her"""
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
    print("  Desktop Folder Widget v3.0")
    print("=" * 55)
    print()
    print("Die Kachel sollte jetzt oben links erscheinen")
    print("(tuerkiser Rand zur besseren Sichtbarkeit)")
    print()
    print("Bedienung:")
    print("  • Linksklick         → Kachel oeffnen")
    print("  • Rechtsklick        → Kontextmenue")
    print("  • Dateien hinziehen  → In Kachel verschieben")
    print("  • Icon wegziehen     → Auf Desktop wiederherstellen")
    print()
    print("WICHTIG: Beim Beenden werden alle Icons wiederhergestellt!")
    print("-" * 55)
    print()
    
    # Cleanup-Handler registrieren
    atexit.register(cleanup_on_exit)
    
    try:
        _app_instance = DesktopFolderManager()
        _app_instance.run()
    except KeyboardInterrupt:
        print("\n[Beendet durch Benutzer]")
    except Exception as e:
        print(f"\n[Fehler] {e}")
    finally:
        cleanup_on_exit()


if __name__ == "__main__":
    main()
