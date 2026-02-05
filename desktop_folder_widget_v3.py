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

# Desktop Grid (typische Windows-Werte)
DESKTOP_GRID_X = 80  # Horizontaler Abstand
DESKTOP_GRID_Y = 90  # Vertikaler Abstand
DESKTOP_MARGIN_X = 10  # Linker Rand
DESKTOP_MARGIN_Y = 10  # Oberer Rand


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


class IconExtractor:
    """Extrahiert Icons aus Dateien - vereinfachte Version"""
    
    ICON_CACHE = {}
    
    @staticmethod
    def get_icon(filepath, size=48):
        """Erstellt ein schönes Icon basierend auf Dateityp"""
        cache_key = f"{filepath}_{size}"
        if cache_key in IconExtractor.ICON_CACHE:
            return IconExtractor.ICON_CACHE[cache_key]
        
        img = IconExtractor.get_default_icon(filepath, size)
        if img:
            IconExtractor.ICON_CACHE[cache_key] = img
        return img
    
    @staticmethod
    def get_default_icon(filepath, size=48):
        """Erstellt ein schönes Standard-Icon basierend auf Dateityp"""
        try:
            from PIL import Image, ImageDraw, ImageFont
        except:
            return None
        
        img = Image.new('RGBA', (size, size), (0, 0, 0, 0))
        draw = ImageDraw.Draw(img)
        
        ext = Path(filepath).suffix.lower() if filepath else ""
        name = Path(filepath).stem if filepath else ""
        
        # Windows 11 ähnliche Farben pro Dateityp
        type_colors = {
            '.exe': '#0078D4', '.msi': '#0078D4', '.lnk': '#0078D4',
            '.bat': '#FFA500', '.cmd': '#FFA500', '.ps1': '#012456',
            '.py': '#3776AB', '.txt': '#6B6B6B', '.pdf': '#FF0000',
            '.doc': '#2B579A', '.docx': '#2B579A',
            '.xls': '#217346', '.xlsx': '#217346',
            '.ppt': '#D24726', '.pptx': '#D24726',
            '.jpg': '#00BCF2', '.jpeg': '#00BCF2', '.png': '#00BCF2',
            '.mp3': '#FF4081', '.mp4': '#FF4081',
            '.zip': '#FFD700', '.rar': '#FFD700', '.7z': '#FFD700',
            '.html': '#E44D26', '.css': '#264DE4', '.js': '#F7DF1E',
        }
        
        color = type_colors.get(ext, '#0078D4')
        
        # Modernes abgerundetes Rechteck
        margin = 3
        radius = 6
        
        # Schatten
        draw.rounded_rectangle(
            [margin + 2, margin + 2, size - margin + 1, size - margin + 1],
            radius=radius,
            fill='#00000040'
        )
        
        # Hauptrechteck
        draw.rounded_rectangle(
            [margin, margin, size - margin, size - margin],
            radius=radius,
            fill=color
        )
        
        # Glanzeffekt oben
        draw.rounded_rectangle(
            [margin + 2, margin + 2, size - margin - 2, margin + size//4],
            radius=radius - 2,
            fill='#FFFFFF30'
        )
        
        # Ersten Buchstaben des Dateinamens zentrieren
        letter = name[0].upper() if name else "?"
        
        # Versuche eine Schriftart zu laden
        font = None
        font_size = size // 2
        try:
            # Windows Schriftart
            font = ImageFont.truetype("segoeui.ttf", font_size)
        except:
            try:
                font = ImageFont.truetype("arial.ttf", font_size)
            except:
                try:
                    font = ImageFont.load_default()
                except:
                    pass
        
        # Text zeichnen
        if font:
            # Textgröße ermitteln
            bbox = draw.textbbox((0, 0), letter, font=font)
            text_width = bbox[2] - bbox[0]
            text_height = bbox[3] - bbox[1]
            
            x = (size - text_width) // 2
            y = (size - text_height) // 2 - 2
            
            draw.text((x, y), letter, fill='white', font=font)
        else:
            # Fallback ohne Font - größerer Text
            draw.text((size//3, size//4), letter, fill='white')
        
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
        
        # Größen - 2x2 Desktop-Icons
        self.tile_width = DESKTOP_GRID_X * 2  # 160
        self.tile_height = DESKTOP_GRID_Y * 2  # 180
        self.expanded_width = 350
        self.expanded_height = 400
        
        self.create_window()
    
    def create_window(self):
        """Erstellt das Kachel-Fenster"""
        self.window = tk.Toplevel(self.manager.root)
        self.window.title(f"DesktopFolder_{self.tile_id}")
        self.window.overrideredirect(True)
        self.window.attributes("-alpha", 0.95)
        
        # Position auf Grid snappen
        x = self.config.get("pos_x", DESKTOP_MARGIN_X + int(self.tile_id) * self.tile_width)
        y = self.config.get("pos_y", DESKTOP_MARGIN_Y)
        x, y = WindowsDesktopAPI.snap_to_grid(x, y)
        
        # Mindestens sichtbar auf dem Bildschirm
        x = max(10, min(x, 1800))
        y = max(10, min(y, 1000))
        
        self.window.geometry(f"{self.tile_width}x{self.tile_height}+{x}+{y}")
        
        # Hauptframe - mit sichtbarem Rand für bessere Erkennbarkeit
        self.main_frame = tk.Frame(
            self.window,
            bg="#1e1e2e",
            highlightthickness=1,
            highlightbackground="#3a3a4a"
        )
        self.main_frame.pack(fill="both", expand=True)
        
        # Canvas für Kachel-Icon
        self.canvas = tk.Canvas(
            self.main_frame,
            width=self.tile_width - 2,
            height=self.tile_height - 2,
            bg="#1e1e2e",
            highlightthickness=0
        )
        self.canvas.pack(fill="both", expand=True)
        
        # Expanded Frame (anfangs None)
        self.expanded_frame = None
        
        # Icon zeichnen
        self.draw_tile_icon()
        
        # Bindings
        self.setup_bindings()
        
        # Drag & Drop
        self.setup_drag_drop()
        
        # Fenster sichtbar machen und im Hintergrund halten
        self.window.after(100, self.setup_window_mode)
    
    def setup_window_mode(self):
        """Macht das Fenster sichtbar und hält es im Hintergrund"""
        try:
            self.window.update_idletasks()
            self.window.deiconify()  # Sicherstellen dass es sichtbar ist
            self.window.lift()  # Kurz nach vorne bringen
            
            # HWND holen
            self.hwnd = user32.GetParent(HWND(self.window.winfo_id()))
            
            print(f"Kachel {self.tile_id} erstellt bei ({self.window.winfo_x()}, {self.window.winfo_y()})")
            
            # Nach kurzer Verzögerung in Hintergrund
            self.window.after(500, self.move_to_background)
            
        except Exception as e:
            print(f"Fehler: {e}")
            import traceback
            traceback.print_exc()
    
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
        self.canvas.bind("<Button-1>", self.on_click)
        self.canvas.bind("<ButtonPress-1>", self.start_drag)
        self.canvas.bind("<B1-Motion>", self.do_drag)
        self.canvas.bind("<ButtonRelease-1>", self.stop_drag)
        self.canvas.bind("<Button-3>", self.show_context_menu)
        self.window.protocol("WM_DELETE_WINDOW", self.close)
        
        # Focus-Loss: Kachel zusammenklappen wenn anderswo geklickt wird
        self.window.bind("<FocusOut>", self.on_focus_out)
    
    def on_focus_out(self, event):
        """Wird aufgerufen wenn das Fenster den Fokus verliert"""
        # Kurze Verzögerung um zu verhindern dass es bei internen Klicks schließt
        if self.is_expanded:
            self.window.after(100, self.check_and_collapse)
    
    def check_and_collapse(self):
        """Prüft ob das Fenster noch fokussiert ist und klappt ggf. zusammen"""
        try:
            # Prüfen ob das Fenster noch existiert und nicht fokussiert ist
            if self.is_expanded and not self.animation_running:
                # Prüfen ob der Fokus auf einem anderen Fenster liegt
                focused = self.window.focus_get()
                if focused is None or not str(focused).startswith(str(self.window)):
                    self.collapse()
        except:
            pass
    
    def setup_drag_drop(self):
        """Drag & Drop aktivieren"""
        if HAS_WINDND:
            windnd.hook_dropfiles(self.window, func=self.on_drop_files)
    
    def on_drop_files(self, files):
        """Dateien wurden auf die Kachel gezogen"""
        desktop_path = WindowsDesktopAPI.get_desktop_path()
        added_count = 0
        
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
            
            # Prüfen ob bereits vorhanden
            already_exists = False
            for shortcut in self.config.get("shortcuts", []):
                if shortcut["path"] == filepath:
                    already_exists = True
                    break
            
            if already_exists:
                continue
            
            # Verknüpfung hinzufügen
            name = Path(filepath).stem
            
            if "shortcuts" not in self.config:
                self.config["shortcuts"] = []
            
            self.config["shortcuts"].append({
                "name": name,
                "path": filepath
            })
            
            # Wenn Datei auf Desktop liegt, verstecken
            if Path(filepath).parent == Path(desktop_path):
                WindowsDesktopAPI.set_file_hidden(filepath, True)
                print(f"Desktop-Icon versteckt: {name}")
            
            added_count += 1
        
        if added_count > 0:
            WindowsDesktopAPI.refresh_desktop()
            self.manager.save_config()
            
            # UI sofort aktualisieren
            self.draw_tile_icon()
            self.canvas.update_idletasks()
            self.window.update()
            
            print(f"{added_count} Verknüpfung(en) hinzugefügt")
    
    def draw_tile_icon(self):
        """Zeichnet das Kachel-Icon wie echte Desktop-Icons (2x2)"""
        self.canvas.delete("all")
        self.icon_images.clear()
        
        shortcuts = self.config.get("shortcuts", [])
        width = self.tile_width
        height = self.tile_height
        
        if not shortcuts:
            # Leeres Ordner-Icon in der Mitte
            self.draw_empty_folder(width, height)
        else:
            # 2x2 Grid wie echte Desktop-Icons
            self.draw_icon_grid(shortcuts[:4], width, height)
        
        # Ordnername unten
        name = self.config.get("name", "Ordner")
        if len(name) > 12:
            name = name[:11] + "…"
        
        self.canvas.create_text(
            width // 2, height - 10,
            text=name,
            fill="white",
            font=("Segoe UI", 9),
            anchor="s"
        )
    
    def draw_empty_folder(self, width, height):
        """Zeichnet leeres Ordner-Icon (Windows-Stil)"""
        cx, cy = width // 2, height // 2 - 15
        
        # Windows-Ordner-Farben
        folder_color = "#FFC107"
        folder_dark = "#FFA000"
        
        w = 80
        h = 60
        tab_w = 32
        tab_h = 12
        
        # Ordner-Tab (oben)
        self.canvas.create_rectangle(
            cx - w//2, cy - h//2 - tab_h,
            cx - w//2 + tab_w, cy - h//2,
            fill=folder_dark, outline=""
        )
        
        # Ordner-Körper
        self.canvas.create_rectangle(
            cx - w//2, cy - h//2,
            cx + w//2, cy + h//2,
            fill=folder_color, outline=""
        )
        
        # Glanzeffekt
        self.canvas.create_rectangle(
            cx - w//2 + 4, cy - h//2 + 4,
            cx + w//2 - 4, cy - h//2 + 12,
            fill="#FFE082", outline=""
        )
    
    def draw_icon_grid(self, shortcuts, width, height):
        """Zeichnet 2x2 Icon-Grid wie Desktop-Icons"""
        # Verfügbarer Platz (ohne Name unten)
        available_height = height - 25
        
        # Icon-Größe wie auf Desktop
        icon_size = 48
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
                self.canvas.create_image(cx, cy - 8, image=icon_img)
            else:
                # Fallback: Farbiges Rechteck
                ext = Path(shortcut["path"]).suffix.lower()
                colors = {'.exe': '#0078D4', '.lnk': '#0078D4', '.bat': '#FFA500'}
                color = colors.get(ext, '#0078D4')
                
                self.canvas.create_rectangle(
                    cx - icon_size//2, cy - icon_size//2 - 8,
                    cx + icon_size//2, cy + icon_size//2 - 8,
                    fill=color, outline=""
                )
                
                letter = shortcut["name"][0].upper() if shortcut["name"] else "?"
                self.canvas.create_text(
                    cx, cy - 8,
                    text=letter,
                    fill="white",
                    font=("Segoe UI", 16, "bold")
                )
            
            # Name unter Icon
            name = shortcut["name"]
            if len(name) > 8:
                name = name[:7] + "…"
            
            self.canvas.create_text(
                cx, cy + icon_size//2,
                text=name,
                fill="white",
                font=("Segoe UI", 8),
                anchor="n"
            )
    
    def on_click(self, event):
        """Klick-Handler"""
        if self.drag_data.get("dragging"):
            return
        
        if self.is_expanded:
            self.collapse()
        else:
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
        
        # Fenster nach vorne bringen
        self.window.attributes("-topmost", True)
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
        """Zeigt Desktop-ähnliche Icon-Ansicht"""
        self.canvas.pack_forget()
        
        self.expanded_frame = tk.Frame(self.main_frame, bg="#1a1a2e")
        self.expanded_frame.pack(fill="both", expand=True)
        
        # Header mit Drag UND Klick-zum-Schließen
        header = tk.Frame(self.expanded_frame, bg="#1a1a2e", cursor="hand2")
        header.pack(fill="x", padx=10, pady=(8, 5))
        
        header.bind("<ButtonPress-1>", self.start_header_drag)
        header.bind("<B1-Motion>", self.do_header_drag)
        header.bind("<ButtonRelease-1>", self.stop_header_drag)
        
        title = tk.Label(
            header, text=self.config.get("name", "Ordner"),
            font=("Segoe UI", 12, "bold"), bg="#1a1a2e", fg="white", cursor="hand2"
        )
        title.pack(side="left")
        title.bind("<ButtonPress-1>", self.start_header_drag)
        title.bind("<B1-Motion>", self.do_header_drag)
        title.bind("<ButtonRelease-1>", self.stop_header_drag)
        
        close_btn = tk.Label(
            header, text="✕", font=("Segoe UI", 12),
            bg="#1a1a2e", fg="#888", cursor="hand2"
        )
        close_btn.pack(side="right")
        close_btn.bind("<Button-1>", lambda e: self.collapse())
        close_btn.bind("<Enter>", lambda e: close_btn.config(fg="#e94560"))
        close_btn.bind("<Leave>", lambda e: close_btn.config(fg="#888"))
        
        # Trennlinie
        tk.Frame(self.expanded_frame, bg="#333", height=1).pack(fill="x", padx=10, pady=5)
        
        # Icon-Grid Container (Desktop-Stil)
        grid_container = tk.Frame(self.expanded_frame, bg="#1a1a2e")
        grid_container.pack(fill="both", expand=True, padx=5, pady=5)
        
        # Canvas mit Scrollbar
        canvas = tk.Canvas(grid_container, bg="#1a1a2e", highlightthickness=0)
        scrollbar = tk.Scrollbar(grid_container, orient="vertical", command=canvas.yview)
        
        self.icons_frame = tk.Frame(canvas, bg="#1a1a2e")
        self.icons_frame.bind("<Configure>", lambda e: canvas.configure(scrollregion=canvas.bbox("all")))
        
        canvas.create_window((0, 0), window=self.icons_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
        # Mausrad
        canvas.bind("<MouseWheel>", lambda e: canvas.yview_scroll(int(-1 * (e.delta / 120)), "units"))
        
        shortcuts = self.config.get("shortcuts", [])
        
        if not shortcuts:
            empty = tk.Label(
                self.icons_frame,
                text="Leer\n\nDateien vom Desktop hierher ziehen",
                font=("Segoe UI", 10), bg="#1a1a2e", fg="#666", justify="center"
            )
            empty.pack(pady=50)
        else:
            self.create_desktop_icon_grid(shortcuts)
        
        # Footer
        footer = tk.Frame(self.expanded_frame, bg="#1a1a2e")
        footer.pack(fill="x", padx=10, pady=8)
        
        add_btn = tk.Label(
            footer, text="➕ Hinzufügen", font=("Segoe UI", 9),
            bg="#2a2a3e", fg="white", padx=12, pady=6, cursor="hand2"
        )
        add_btn.pack(side="left")
        add_btn.bind("<Button-1>", lambda e: self.add_shortcut_dialog())
        add_btn.bind("<Enter>", lambda e: add_btn.config(bg="#3a3a4e"))
        add_btn.bind("<Leave>", lambda e: add_btn.config(bg="#2a2a3e"))
        
        self.animation_running = False
    
    def create_desktop_icon_grid(self, shortcuts):
        """Erstellt Desktop-ähnliches Icon-Grid"""
        cols = 4
        icon_size = 42
        cell_width = 78
        cell_height = 72
        
        for i, shortcut in enumerate(shortcuts):
            row = i // cols
            col = i % cols
            
            # Container wie auf dem Desktop
            icon_frame = tk.Frame(
                self.icons_frame,
                bg="#1a1a2e",
                width=cell_width,
                height=cell_height
            )
            icon_frame.grid(row=row, column=col, padx=1, pady=1)
            icon_frame.grid_propagate(False)
            
            # Icon zentriert oben
            icon_container = tk.Frame(icon_frame, bg="#1a1a2e")
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
                    icon_label = tk.Label(icon_container, image=icon_img, bg="#1a1a2e")
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
                    bg="#1a1a2e", 
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
                canvas.create_text(
                    icon_size//2, icon_size//2,
                    text=letter,
                    fill="white",
                    font=("Segoe UI", 16, "bold")
                )
                icon_label = canvas
            
            # Name (kürzer, keine Umbrüche)
            name = shortcut["name"]
            if len(name) > 10:
                name = name[:9] + "…"
            
            name_label = tk.Label(
                icon_frame,
                text=name,
                font=("Segoe UI", 8),
                bg="#1a1a2e",
                fg="white",
                anchor="center"
            )
            name_label.pack()
            
            # Event-Handler mit Drag-Out Funktion
            def make_handlers(idx, path, shortcut_name, frame, name_lbl, icon_lbl):
                drag_data = {'dragging': False, 'start_x': 0, 'start_y': 0}
                
                def on_enter(e):
                    frame.config(bg="#3a3a5e")
                    name_lbl.config(bg="#3a3a5e")
                    try:
                        icon_lbl.config(bg="#3a3a5e")
                    except:
                        pass
                
                def on_leave(e):
                    frame.config(bg="#1a1a2e")
                    name_lbl.config(bg="#1a1a2e")
                    try:
                        icon_lbl.config(bg="#1a1a2e")
                    except:
                        pass
                
                def on_press(e):
                    drag_data['start_x'] = e.x_root
                    drag_data['start_y'] = e.y_root
                    drag_data['dragging'] = False
                
                def on_motion(e):
                    dx = abs(e.x_root - drag_data['start_x'])
                    dy = abs(e.y_root - drag_data['start_y'])
                    if dx > 15 or dy > 15:
                        drag_data['dragging'] = True
                        # Visuelles Feedback
                        frame.config(bg="#e94560")
                        name_lbl.config(bg="#e94560")
                
                def on_release(e):
                    if drag_data['dragging']:
                        # Drag beendet - auf Desktop wiederherstellen
                        print(f"Drag-Out erkannt für: {shortcut_name}")
                        self.restore_to_desktop(idx)
                    else:
                        # Normaler Klick - Programm starten
                        self.launch_shortcut(path)
                    drag_data['dragging'] = False
                
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
        menu = tk.Menu(self.window, tearoff=0, bg="#1a1a2e", fg="white",
                       activebackground="#e94560")
        
        menu.add_command(label="▶️ Öffnen", command=lambda: self.launch_shortcut(path))
        menu.add_command(label="📤 Auf Desktop wiederherstellen",
                        command=lambda: self.restore_to_desktop(index))
        menu.add_separator()
        menu.add_command(label="🗑️ Aus Ordner entfernen",
                        command=lambda: self.remove_shortcut(index))
        
        menu.tk_popup(event.x_root, event.y_root)
    
    def restore_to_desktop(self, index):
        """Stellt Verknüpfung auf Desktop wieder her (macht sie sichtbar)"""
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
                    text="Leer\n\nDateien vom Desktop hierher ziehen\noder Icon nach außen ziehen zum Wiederherstellen",
                    font=("Segoe UI", 10), bg="#1a1a2e", fg="#666", justify="center"
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
        self.draw_tile_icon()
        self.animation_running = False
        
        # Zurück in den Hintergrund
        self.window.attributes("-topmost", False)
        self.window.after(100, self.move_to_background)
    
    def animate_size(self, from_w, from_h, to_w, to_h, x, y, callback=None):
        """Größenanimation"""
        steps = 8
        step_time = 12
        
        dw = (to_w - from_w) / steps
        dh = (to_h - from_h) / steps
        
        def step(i):
            if i <= steps:
                new_w = int(from_w + dw * i)
                new_h = int(from_h + dh * i)
                self.window.geometry(f"{new_w}x{new_h}+{x}+{y}")
                self.window.after(step_time, lambda: step(i + 1))
            else:
                if callback:
                    callback()
        
        step(1)
    
    def show_context_menu(self, event):
        """Kontextmenü der Kachel"""
        menu = tk.Menu(self.window, tearoff=0, bg="#1a1a2e", fg="white",
                       activebackground="#e94560")
        
        menu.add_command(label="📂 Öffnen", command=self.expand)
        menu.add_command(label="✏️ Umbenennen", command=self.rename)
        menu.add_separator()
        menu.add_command(label="🆕 Neue Kachel", command=self.manager.create_new_tile)
        menu.add_separator()
        menu.add_command(label="📤 Alle wiederherstellen", command=self.restore_all_to_desktop)
        menu.add_command(label="🗑️ Kachel löschen", command=self.delete_tile)
        menu.add_separator()
        menu.add_command(label="❌ Widget beenden", command=self.manager.quit)
        
        menu.tk_popup(event.x_root, event.y_root)
    
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
