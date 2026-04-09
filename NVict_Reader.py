# -*- coding: utf-8 -*-
"""
NVict Reader (Modern UI Style) - Optimized Edition
Gebaseerd op de UI-stijl van NV Sync
Ontwikkeld door NVict Service

Website: www.nvict.nl
Versie: 2.2
"""

import sys
import os
import webbrowser
import tkinter as tk
from tkinter import filedialog, messagebox, ttk, simpledialog
# Lazy imports for heavy modules - imported when needed
# import fitz  # PyMuPDF - NOW LAZY LOADED
# from PIL import Image, ImageTk, ImageOps, ImageDraw - NOW LAZY LOADED
import io
import tempfile
import subprocess
import platform
from datetime import datetime
import urllib.request
import urllib.error
import json
import threading
import socket
import time

# Applicatie versie
APP_VERSION = "2.2"
UPDATE_CHECK_URL = "https://www.nvict.nl/software/updates/nvict_reader_version.json"

# ====================================================================
# LAZY IMPORTS - Heavy modules loaded only when needed
# ====================================================================

_fitz = None
_PIL_modules = None

def get_fitz():
    """Lazy import van PyMuPDF (fitz) - alleen laden wanneer PDF wordt geopend"""
    global _fitz
    if _fitz is None:
        import fitz
        _fitz = fitz
    return _fitz

def get_PIL():
    """Lazy import van PIL modules - alleen laden wanneer nodig"""
    global _PIL_modules
    if _PIL_modules is None:
        from PIL import Image, ImageTk, ImageOps, ImageDraw
        _PIL_modules = (Image, ImageTk, ImageOps, ImageDraw)
    return _PIL_modules

try:
    import winreg
except ImportError:
    winreg = None

# ====================================================================
# DEFAULT PDF HANDLER - Set as Default Functionality
# ====================================================================

class DefaultPDFHandler:
    """Handles setting NVict Reader as default PDF viewer"""
    
    @staticmethod
    def is_default_pdf_handler():
        """Check if NVict Reader is currently the default PDF handler"""
        try:
            # Haal de huidige executable naam op
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
            else:
                current_exe = os.path.abspath(sys.argv[0])

            current_exe_lower = current_exe.lower()
            exe_name = os.path.basename(current_exe_lower)

            # Check UserChoice ProgId
            try:
                key = winreg.OpenKey(
                    winreg.HKEY_CURRENT_USER,
                    r"Software\Microsoft\Windows\CurrentVersion\Explorer\FileExts\.pdf\UserChoice",
                    0,
                    winreg.KEY_READ
                )
                prog_id, _ = winreg.QueryValueEx(key, "ProgId")
                winreg.CloseKey(key)

                prog_id_lower = prog_id.lower()

                # Directe match op onze ProgID
                if prog_id == "NVictReader.PDF":
                    return True

                # Match op exe-naam in ProgId (Applications\NVict Reader.exe, etc.)
                if "nvict" in prog_id_lower or "nvictreader" in prog_id_lower:
                    return True

                # Zoek het command-pad dat bij deze ProgId hoort
                for root_key in (winreg.HKEY_CURRENT_USER, winreg.HKEY_CLASSES_ROOT):
                    for sub in (f"Software\\Classes\\{prog_id}\\shell\\open\\command",
                                f"{prog_id}\\shell\\open\\command"):
                        try:
                            cmd_key = winreg.OpenKey(root_key, sub, 0, winreg.KEY_READ)
                            command, _ = winreg.QueryValueEx(cmd_key, "")
                            winreg.CloseKey(cmd_key)
                            if exe_name in command.lower() or current_exe_lower in command.lower():
                                return True
                        except Exception:
                            continue
            except Exception:
                pass

            return False
        except Exception as e:
            print(f"Error checking default handler: {e}")
            return False
    
    @staticmethod
    def open_windows_default_apps_pdf():
        """Open Windows Settings directly to PDF file association"""
        try:
            # Probeert direct naar de .pdf instelling te gaan (Windows 10/11)
            subprocess.run(['start', 'ms-settings:defaultapps'], shell=True)
            return True
        except:
            return False

    @staticmethod
    def register_open_with():
        """
        Registreer NVict Reader in het register.
        CHECK: Als Inno Setup het al in HKLM heeft gezet, doen we hier NIETS om dubbele items te voorkomen.
        """
        try:
            # 1. CHECK: Is de app al globaal geïnstalleerd via Inno Setup?
            # We kijken of de ProgID in HKEY_LOCAL_MACHINE bestaat.
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\NVictReader.PDF", 0, winreg.KEY_READ)
                winreg.CloseKey(key)
                # Gevonden! De installer heeft zijn werk gedaan.
                # Wij doen niets in Python om duplicaten te voorkomen.
                return True
            except OSError:
                # Niet gevonden in HKLM, dus we draaien waarschijnlijk portable.
                # Ga door met registreren in HKCU.
                pass

            # ---------------------------------------------------------
            # Code voor Portable Versie (Schrijft naar HKEY_CURRENT_USER)
            # ---------------------------------------------------------
            
            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(sys.argv[0])
            
            prog_id = "NVictReader.PDF"
            hkcu = winreg.HKEY_CURRENT_USER
            
            # RegisteredApplications pad
            cap_path = f"Software\\NVict Service\\NVict Reader\\Capabilities"
            
            # Maak de Capabilities sleutels
            key = winreg.CreateKey(hkcu, cap_path)
            winreg.SetValueEx(key, "ApplicationName", 0, winreg.REG_SZ, "NVict Reader")
            winreg.SetValueEx(key, "ApplicationDescription", 0, winreg.REG_SZ, "NVict Reader PDF Viewer")
            winreg.CloseKey(key)
            
            # FileAssociations binnen Capabilities
            key = winreg.CreateKey(hkcu, f"{cap_path}\\FileAssociations")
            winreg.SetValueEx(key, ".pdf", 0, winreg.REG_SZ, prog_id)
            winreg.CloseKey(key)
            
            # Voeg toe aan RegisteredApplications
            key = winreg.CreateKey(hkcu, "Software\\RegisteredApplications")
            winreg.SetValueEx(key, "NVictReader", 0, winreg.REG_SZ, cap_path)
            winreg.CloseKey(key)
            
            # De ProgID
            classes_path = f"Software\\Classes\\{prog_id}"
            
            key = winreg.CreateKey(hkcu, classes_path)
            winreg.SetValue(key, "", winreg.REG_SZ, "NVict Reader PDF")
            winreg.CloseKey(key)
            
            # Gebruik PDF_File_icon.ico voor Windows Verkenner, anders exe zelf
            key = winreg.CreateKey(hkcu, f"{classes_path}\\DefaultIcon")
            icon_dir = os.path.dirname(exe_path)
            pdf_icon_path = os.path.join(icon_dir, "PDF_File_icon.ico")
            if os.path.exists(pdf_icon_path):
                winreg.SetValue(key, "", winreg.REG_SZ, f'"{pdf_icon_path}",0')
            else:
                winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}",0')
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\open\\command")
            winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}" "%1"')
            winreg.CloseKey(key)
            
            # Voeg "Afdrukken" toe aan rechtsklik menu
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print")
            winreg.SetValue(key, "", winreg.REG_SZ, "Afdrukken")
            winreg.CloseKey(key)
            
            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print\\command")
            winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}" --print "%1"')
            winreg.CloseKey(key)
            
            # Koppel .pdf aan ProgID
            key = winreg.CreateKey(hkcu, "Software\\Classes\\.pdf\\OpenWithProgids")
            winreg.SetValueEx(key, prog_id, 0, winreg.REG_NONE, b'')
            winreg.CloseKey(key)

            # Refresh de shell iconen
            try:
                import ctypes
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, 0, 0)
            except:
                pass
            
            return True
        except Exception as e:
            print(f"Error registering: {e}")
            return False

    @staticmethod
    def prompt_set_as_default(parent):
        """Vraagt de gebruiker om de app als standaard in te stellen"""
        # Eerst registreren in het register om zeker te zijn dat we in de lijst staan
        DefaultPDFHandler.register_open_with()
        
        msg = (
            "Om NVict Reader als standaard in te stellen, opent Windows nu het instellingen menu.\n\n"
            "1. Zoek '.pdf' in de lijst of klik op de huidige standaard app.\n"
            "2. Selecteer 'NVict Reader' in de lijst.\n"
            "3. Klik op 'Als standaard instellen'.\n\n"
            "Wilt u de instellingen nu openen?"
        )
        
        if messagebox.askyesno("Standaard App Instellen", msg, parent=parent):
            DefaultPDFHandler.open_windows_default_apps_pdf()

    @staticmethod
    def show_first_run_dialog(parent):
        """Toon dialoog bij eerste keer opstarten met 'Nooit meer vragen' optie"""
        # Check of we al standaard zijn
        if DefaultPDFHandler.is_default_pdf_handler():
            return "already_default"

        # Maak custom dialoog window
        dialog = tk.Toplevel(parent)
        dialog.title("Welkom bij NVict Reader")
        dialog.geometry("450x220")
        dialog.resizable(False, False)
        dialog.transient(parent)
        dialog.grab_set()

        # Center dialoog op parent window
        dialog.update_idletasks()
        x = parent.winfo_x() + (parent.winfo_width() - dialog.winfo_width()) // 2
        y = parent.winfo_y() + (parent.winfo_height() - dialog.winfo_height()) // 2
        dialog.geometry(f"+{x}+{y}")

        # Bericht
        msg = (
            "Welkom bij NVict Reader!\n\n"
            "Wilt u NVict Reader instellen als uw standaard PDF programma?\n"
            "Dit kunt u later altijd nog wijzigen via Instellingen."
        )

        label = tk.Label(dialog, text=msg, font=("Arial", 10),
                        justify=tk.LEFT, padx=20, pady=20)
        label.pack(fill=tk.BOTH, expand=True)

        # Checkbox voor "Nooit meer vragen"
        never_ask_var = tk.BooleanVar(value=False)
        checkbox = tk.Checkbutton(dialog, text="Niet meer vragen",
                                 variable=never_ask_var, font=("Arial", 9),
                                 padx=20, pady=5)
        checkbox.pack(fill=tk.X)

        result = [None]

        def on_yes():
            result[0] = "never" if never_ask_var.get() else "yes"
            DefaultPDFHandler.prompt_set_as_default(parent)
            dialog.destroy()

        def on_no():
            result[0] = "never" if never_ask_var.get() else "no"
            dialog.destroy()

        # Knoppen
        btn_frame = tk.Frame(dialog)
        btn_frame.pack(fill=tk.X, padx=20, pady=15)

        btn_yes = tk.Button(btn_frame, text="Ja", width=10, command=on_yes)
        btn_yes.pack(side=tk.LEFT, padx=5)

        btn_no = tk.Button(btn_frame, text="Nee", width=10, command=on_no)
        btn_no.pack(side=tk.LEFT, padx=5)

        parent.wait_window(dialog)
        return result[0] if result[0] else "no"

# ====================================================================

# ====================================================================

class SingleInstance:
    """Zorgt ervoor dat er maar één instance van de applicatie draait."""
    def __init__(self, port=52847):
        self.port = port
        self.sock = None
        self.server_thread = None
        self.app = None
        self.running = False
        
    def is_already_running(self):
        """Controleer of er al een instance draait."""
        try:
            # Probeer te connecteren met bestaande instance
            test_sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            test_sock.settimeout(1)
            test_sock.connect(('127.0.0.1', self.port))
            test_sock.close()
            return True
        except (socket.error, socket.timeout):
            return False
    
    def send_to_existing_instance(self, filepath):
        """Stuur bestandspad naar bestaande instance."""
        try:
            sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            sock.settimeout(2)
            sock.connect(('127.0.0.1', self.port))
            sock.sendall(filepath.encode('utf-8'))
            sock.close()
            return True
        except Exception as e:
            print(f"Fout bij versturen naar bestaande instance: {e}")
            return False
    
    def start_server(self, app):
        """Start socket server om berichten van andere instances te ontvangen."""
        self.app = app
        self.running = True
        
        try:
            self.sock = socket.socket(socket.AF_INET, socket.SOCK_STREAM)
            self.sock.setsockopt(socket.SOL_SOCKET, socket.SO_REUSEADDR, 1)
            self.sock.bind(('127.0.0.1', self.port))
            self.sock.listen(5)
            self.sock.settimeout(1)
            
            # Start server thread
            self.server_thread = threading.Thread(target=self._server_loop, daemon=True)
            self.server_thread.start()
            
            return True
        except Exception as e:
            print(f"Kon single instance server niet starten: {e}")
            return False
    
    def _server_loop(self):
        """Luister naar berichten van andere instances."""
        while self.running:
            try:
                conn, addr = self.sock.accept()
                conn.settimeout(2)
                
                # Ontvang bestandspad
                data = conn.recv(4096).decode('utf-8')
                conn.close()
                
                if data and self.app:
                    # Open bestand in bestaande instance (in main thread)
                    self.app.root.after(0, lambda path=data: self.app.add_new_tab(path))
                    
                    # Breng window naar voren (ook als geminimaliseerd)
                    def bring_to_front():
                        if self.app.root.state() == 'iconic':
                            self.app.root.deiconify()
                        self.app.root.lift()
                        self.app.root.focus_force()
                    
                    self.app.root.after(10, bring_to_front)
                    
            except socket.timeout:
                continue
            except Exception as e:
                if self.running:
                    print(f"Fout in server loop: {e}")
    
    def stop(self):
        """Stop de server."""
        self.running = False
        if self.sock:
            try:
                self.sock.close()
            except:
                pass

def get_resource_path(relative_path):
    """Geef het absolute pad naar resource bestanden (werkt met PyInstaller)"""
    try:
        # PyInstaller maakt een temp folder en slaat path op in _MEIPASS
        base_path = sys._MEIPASS
    except Exception:
        base_path = os.path.abspath(".")
    
    return os.path.join(base_path, relative_path)

def get_settings_path():
    """Geef pad naar settings bestand in gebruikers map"""
    if platform.system() == "Windows":
        app_data = os.environ.get('APPDATA', os.path.expanduser('~'))
        settings_dir = os.path.join(app_data, 'NVict PDF Reader')
    else:
        settings_dir = os.path.join(os.path.expanduser('~'), '.nvict_pdf_reader')
    
    # Maak directory als die niet bestaat
    os.makedirs(settings_dir, exist_ok=True)
    return os.path.join(settings_dir, 'settings.json')

class Theme:
    """Bevat de kleurenschema's voor lichte en donkere thema's."""
    LIGHT = {
        "BG_PRIMARY": "#f3f3f3", "BG_SECONDARY": "#ffffff",
        "TEXT_PRIMARY": "#1c1c1c", "TEXT_SECONDARY": "#737373",
        "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
        "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
        "SELECTION_COLOR": "#FFD700"  # Goud/geel voor betere zichtbaarheid
    }
    DARK = {
        "BG_PRIMARY": "#1e1e1e", "BG_SECONDARY": "#2d2d2d",
        "TEXT_PRIMARY": "#f0f0f0", "TEXT_SECONDARY": "#a0a0a0",
        "ACCENT_COLOR": "#10a2dd", "SUCCESS_COLOR": "#28a745",
        "WARNING_COLOR": "#ff8c00", "ERROR_COLOR": "#d13438",
        "SELECTION_COLOR": "#FFD700"  # Goud/geel
    }
    FONT_MAIN = ("Segoe UI Variable", 10)
    FONT_HEADING = ("Segoe UI Variable", 12, "bold")
    FONT_SMALL = ("Segoe UI Variable", 9)

class Tooltip:
    """Toon een tooltip bij hover over een widget."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tip_window = None
        self._after_id = None
        widget.bind('<Enter>', self._schedule, add='+')
        widget.bind('<Leave>', self._hide, add='+')
        widget.bind('<ButtonPress>', self._hide, add='+')

    def _schedule(self, event=None):
        self._cancel()
        self._after_id = self.widget.after(500, self._show)

    def _cancel(self):
        if self._after_id:
            self.widget.after_cancel(self._after_id)
            self._after_id = None

    def _show(self):
        if self.tip_window or not self.text:
            return
        x = self.widget.winfo_rootx() + self.widget.winfo_width() // 2
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 6
        self.tip_window = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        tk.Label(tw, text=self.text, justify=tk.LEFT,
                 background="#ffffe0", foreground="#1c1c1c",
                 relief="solid", borderwidth=1,
                 font=Theme.FONT_SMALL, padx=6, pady=3).pack()

    def _hide(self, event=None):
        self._cancel()
        if self.tip_window:
            self.tip_window.destroy()
            self.tip_window = None


class PDFTab(tk.Frame):
    """Een enkel tabblad dat een PDF-document beheert."""
    def __init__(self, master, file_path, theme, password=None):
        super().__init__(master, bg=theme["BG_PRIMARY"])
        self.theme = theme
        
        # Document state
        self.file_path = file_path
        fitz = get_fitz()  # Lazy load PyMuPDF
        self.pdf_document = fitz.open(file_path)
        
        # Authenticeer met wachtwoord indien nodig
        if password and self.pdf_document.needs_pass:
            self.pdf_document.authenticate(password)

        # Detecteer beveiliging en ondertekening
        self.is_encrypted = self.pdf_document.is_encrypted
        self.is_signed = self._detect_signatures()
        self.security_info = self._get_security_info()

        self.current_page = 0
        self.zoom_level = 1.0
        self.zoom_mode = "fit_width"
        
        # UI elements
        self.canvas = tk.Canvas(self, bg=theme["BG_PRIMARY"], relief="flat", bd=0, 
                               highlightthickness=0)
        v_scrollbar = ttk.Scrollbar(self, orient=tk.VERTICAL, command=self.canvas.yview)
        h_scrollbar = ttk.Scrollbar(self, orient=tk.HORIZONTAL, command=self.canvas.xview)
        self.canvas.configure(yscrollcommand=v_scrollbar.set, xscrollcommand=h_scrollbar.set)
        
        v_scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        h_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
        self.canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)

        # Selection data
        self.text_words = []
        self.drag_start = None
        self.drag_rect = None
        self.selection_rects = []
        self.selected_text = ""
        
        # Images
        self.current_image = None
        self.highlighted_image = None
        self.page_offset_x = 0
        self.page_offset_y = 0
        self.page_images = []  # Voor continuous scroll
        self.page_pil_images = {}  # PIL images voor elke pagina (voor highlighting)
        self.page_positions = []  # Y-positie van elke pagina
        self.scroll_to_page = None  # Flag voor initiële scroll

        # Links (nieuw voor hyperlinks)
        self.links = []  # Lijst van (page_num, rect, uri/dest) tuples
        self.current_cursor = "arrow"  # Track cursor state

        # Weergave modi
        self.book_mode = False   # Boek-modus: twee pagina's naast elkaar
        self.page_regions = {}   # {page_num: (x0, y0, x1, y1)} canvas bounding box per pagina

        # Formulier invulvelden
        self.form_mode = False           # Formuliermodus aan/uit
        self.form_widgets = []           # Lijst van (canvas_window_id, widget, field_info) tuples
        self.form_field_values = {}      # {field_xref: waarde} opgeslagen waarden

        # Vrije tekst annotaties
        self.text_annotate_mode = False  # Tekst-annotatiemodus aan/uit
        self.text_annotations = []       # [(page_num, pdf_x, pdf_y, text, font_size), ...]
        self.text_annot_widgets = []     # Actieve canvas overlay widgets

    def _detect_signatures(self):
        """Detecteer of het PDF digitaal ondertekend is"""
        try:
            # Controleer voor signature fields in alle pagina's
            for page_num in range(len(self.pdf_document)):
                page = self.pdf_document[page_num]
                annots = page.annots()
                if annots:
                    for annot in annots:
                        if annot and annot.info.get('subtype') == 'Sig':
                            return True
        except:
            pass
        return False

    def _get_security_info(self):
        """Haal beveiligingsinformatie op"""
        info = []

        try:
            # Controleer op encryptie/wachtwoord
            if self.is_encrypted:
                info.append("🔒 Beveiligd (wachtwoord)")

            # Controleer op permission restrictions
            try:
                perms = self.pdf_document.permissions
                restrictions = []

                if not (perms & 4):      # Bit 2: Print
                    restrictions.append("afdrukken")
                if not (perms & 8):      # Bit 3: Modify
                    restrictions.append("bewerken")
                if not (perms & 16):     # Bit 4: Copy
                    restrictions.append("kopiëren")
                if not (perms & 32):     # Bit 5: Add annotations
                    restrictions.append("annotaties")

                if restrictions:
                    info.append(f"🔐 Beperkt ({', '.join(restrictions)})")
            except:
                pass

            # Controleer op digitale handtekeningen
            if self.is_signed:
                info.append("🔏 Digitaal ondertekend")

        except:
            pass

        return " | ".join(info) if info else ""

    def close_document(self):
        if self.pdf_document:
            self.pdf_document.close()
            self.pdf_document = None

class NVictReader:
    def __init__(self):
        self.root = tk.Tk()
        self.root.title("NVict Reader")
        self.root.geometry("1366x768")
        self.root.minsize(800, 600)

        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                self.root.iconbitmap(icon_path)
        except:
            pass

        self.config = {"theme": "Systeemstandaard"}
        self.highlight_mode = False    # Markeermodus (geel markeren)
        self.thumbnail_visible = True  # Thumbnail-paneel zichtbaar
        self.thumbnail_images = []     # Houdt foto-referenties bij (GC preventie)

        # Update instellingen laden
        self.load_update_settings()
        
        # Herstel opgeslagen schermgrootte en positie
        if self.update_settings.get('window_geometry'):
            try:
                self.root.geometry(self.update_settings['window_geometry'])
            except:
                self.root.geometry("1366x768")  # Fallback naar standaard
        
        # Herstel window state (maximized of normal)
        if self.update_settings.get('window_state') == 'zoomed':
            self.root.state('zoomed')
        
        self.apply_theme()
        
        self.load_icons()
        
        self.setup_ui()
        
        self.setup_shortcuts()
        
        self.update_ui_state()
        
        # Initialiseer drag-and-drop ondersteuning
        self.setup_drag_and_drop()
        
        # Check first run en vraag om default PDF viewer te worden
        if self.update_settings.get('first_run', True):
            self.root.after(1000, self.check_first_run)

    def check_first_run(self):
        """Check first run and show welcome dialog"""
        # Check eerst of we al de standaard PDF viewer zijn
        if DefaultPDFHandler.is_default_pdf_handler():
            # We zijn al de standaard, geen dialoog tonen
            self.update_settings['first_run'] = False
            self.update_settings['ask_default'] = False
            self.save_update_settings()
            return
        
        # We zijn nog niet de standaard, vraag of gebruiker dit wil instellen
        if self.update_settings.get('ask_default', True):
            result = DefaultPDFHandler.show_first_run_dialog(self.root)
            if result == "never":
                self.update_settings['ask_default'] = False
        
        self.update_settings['first_run'] = False
        self.save_update_settings()

    def setup_drag_and_drop(self):
        """Configureer drag-and-drop ondersteuning voor PDF bestanden"""
        try:
            import windnd

            def on_drop(file_list):
                """Verwerk gesleepte bestanden"""
                for file_bytes in file_list:
                    # windnd geeft bytes terug, decodeer naar string
                    if isinstance(file_bytes, bytes):
                        file_path = file_bytes.decode('utf-8', errors='replace')
                    else:
                        file_path = str(file_bytes)

                    # Verwijder eventuele aanhalingstekens
                    file_path = file_path.strip('"').strip("'")

                    if file_path.lower().endswith('.pdf') and os.path.isfile(file_path):
                        self.add_new_tab(file_path)

            windnd.hook_dropfiles(self.root, func=on_drop)

        except ImportError:
            # windnd niet beschikbaar — probeer tkinterdnd2 als fallback
            try:
                from tkinterdnd2 import DND_FILES
                self.root.drop_target_register(DND_FILES)
                self.root.dnd_bind('<<Drop>>', self._on_tkdnd_drop)
            except Exception:
                pass  # Geen drag-and-drop beschikbaar, programma werkt verder zonder

    def _on_tkdnd_drop(self, event):
        """Fallback drop handler voor tkinterdnd2"""
        files = self.root.tk.splitlist(event.data)
        for f in files:
            f = f.strip('"').strip("'")
            if f.lower().endswith('.pdf') and os.path.isfile(f):
                self.add_new_tab(f)

    def apply_theme(self):
        saved = getattr(self, 'update_settings', {}).get('theme', 'Systeemstandaard')
        theme_name = self.get_windows_theme() if saved == 'Systeemstandaard' else saved
        self.theme = Theme.LIGHT if theme_name == "Licht" else Theme.DARK
        self.root.configure(bg=self.theme["BG_PRIMARY"])
        
        style = ttk.Style(self.root)
        style.theme_use('clam')
        style.configure("TNotebook", background=self.theme["BG_PRIMARY"], borderwidth=0)
        style.configure("TNotebook.Tab", background=self.theme["BG_PRIMARY"], 
                       foreground=self.theme["TEXT_SECONDARY"], borderwidth=0, padding=[10, 5])
        style.map("TNotebook.Tab", background=[("selected", self.theme["BG_SECONDARY"])], 
                 foreground=[("selected", self.theme["TEXT_PRIMARY"])])
        style.configure("TScrollbar", background=self.theme["BG_PRIMARY"], 
                       troughcolor=self.theme["BG_SECONDARY"], bordercolor=self.theme["BG_PRIMARY"], 
                       arrowcolor=self.theme["TEXT_PRIMARY"])
        style.map("TScrollbar", background=[('active', self.theme["ACCENT_COLOR"])])

    def get_windows_theme(self):
        try:
            if winreg:
                key = winreg.OpenKey(winreg.HKEY_CURRENT_USER, 
                                    r"Software\Microsoft\Windows\CurrentVersion\Themes\Personalize")
                value, _ = winreg.QueryValueEx(key, "AppsUseLightTheme")
                return "Licht" if value == 1 else "Donker"
        except (FileNotFoundError, AttributeError):
            pass
        return "Licht"
        
    _ICON_FILES = {
        "open": "open.png", "close": "close.png", "print": "print.png",
        "zoom-in": "zoom-in.png", "zoom-out": "zoom-out.png",
        "reset": "reset.png", "prev-page": "prev-page.png", "next-page": "next-page.png",
        "first-page": "first-page.png", "last-page": "last-page.png", "info": "info.png",
        "copy": "copy.png", "search": "search.png", "pdf": "pdf.png",
        "fit-width": "fit-width.png", "save": "save.png", "toolbox": "toolbox.png",
        "full-screen": "full-screen.png", "theme": "theme.png",
        "send": "send.png", "form": "form.png", "type-text": "type-text.png",
        "pages": "pages.png", "marker": "marker.png", "book": "book.png"
    }

    def load_icons(self):
        """Laad icons asynchroon in achtergrond voor snellere startup"""
        self.icons = {name: None for name in self._ICON_FILES}
        self.icons_loaded = False
        self.toolbar_buttons = {}
        self._start_icon_load_thread()

    def _start_icon_load_thread(self):
        """(Her)start de achtergrond thread die icons laadt"""
        self.icons = {name: None for name in self._ICON_FILES}
        self.icons_loaded = False

        def load_icons_thread():
            try:
                Image, ImageTk, ImageOps, ImageDraw = get_PIL()
                is_dark_theme = self.theme == Theme.DARK
                for name, filename in self._ICON_FILES.items():
                    try:
                        path = get_resource_path(os.path.join('icons', filename))
                        if os.path.exists(path):
                            image = Image.open(path).resize((18, 18), Image.Resampling.LANCZOS)
                            if is_dark_theme:
                                if image.mode == 'RGBA':
                                    r, g, b, a = image.split()
                                    inverted = ImageOps.invert(Image.merge('RGB', (r, g, b)))
                                    r2, g2, b2 = inverted.split()
                                    image = Image.merge('RGBA', (r2, g2, b2, a))
                                else:
                                    image = ImageOps.invert(image)
                            self.icons[name] = ImageTk.PhotoImage(image)
                    except Exception:
                        pass
                self.icons_loaded = True
                try:
                    self.root.after(0, self.update_toolbar_icons)
                except:
                    pass
            except Exception as e:
                print(f"Icon loading error: {e}")

        threading.Thread(target=load_icons_thread, daemon=True).start()
    
    def update_toolbar_icons(self):
        """Update toolbar icons nadat ze in achtergrond zijn geladen"""
        if hasattr(self, 'toolbar_buttons') and self.icons_loaded:
            for btn_name, (btn_widget, original_text) in self.toolbar_buttons.items():
                if btn_name in self.icons and self.icons[btn_name]:
                    try:
                        btn_widget.config(image=self.icons[btn_name], text=original_text, compound=tk.LEFT)
                    except:
                        pass

    def setup_ui(self):
        # Maak menubar
        self.create_menubar()
        
        self.create_modern_toolbar()
        
        # ── Hoofd-container: thumbnail-paneel links + notebook rechts ──
        self.content_frame = tk.Frame(self.root, bg=self.theme["BG_PRIMARY"])
        self.content_frame.pack(fill=tk.BOTH, expand=True, padx=20, pady=(10, 20))
        self.content_frame.grid_columnconfigure(1, weight=1)
        self.content_frame.grid_rowconfigure(0, weight=1)

        # ── Thumbnail-paneel ──
        self.thumbnail_panel = tk.Frame(
            self.content_frame, bg=self.theme["BG_SECONDARY"],
            width=114, relief="flat",
            highlightbackground=self.theme["TEXT_SECONDARY"], highlightthickness=1
        )
        self.thumbnail_panel.grid(row=0, column=0, sticky="ns", padx=(0, 5))
        self.thumbnail_panel.grid_propagate(False)

        _thumb_hdr = tk.Label(
            self.thumbnail_panel, text="Pagina's",
            font=Theme.FONT_SMALL, bg=self.theme["BG_SECONDARY"],
            fg=self.theme["TEXT_SECONDARY"]
        )
        _thumb_hdr.pack(side=tk.TOP, fill=tk.X, pady=(4, 2))

        self.thumbnail_canvas = tk.Canvas(
            self.thumbnail_panel, bg=self.theme["BG_SECONDARY"],
            width=114, highlightthickness=0
        )
        _thumb_sb = ttk.Scrollbar(
            self.thumbnail_panel, orient=tk.VERTICAL,
            command=self.thumbnail_canvas.yview
        )
        self.thumbnail_canvas.configure(yscrollcommand=_thumb_sb.set)
        _thumb_sb.pack(side=tk.RIGHT, fill=tk.Y)
        self.thumbnail_canvas.pack(side=tk.LEFT, fill=tk.BOTH, expand=True)
        # Muiswiel op thumbnail canvas
        self.thumbnail_canvas.bind(
            "<MouseWheel>",
            lambda e: self.thumbnail_canvas.yview_scroll(
                -1 if e.delta > 0 else 1, "units"
            )
        )

        # ── Notebook (tabbladen) ──
        self.notebook = ttk.Notebook(self.content_frame)
        self.notebook.grid(row=0, column=1, sticky="nsew")
        self.notebook.bind("<<NotebookTabChanged>>", self.on_tab_change)

        self.welcome_frame = tk.Frame(self.notebook, bg=self.theme["BG_PRIMARY"])
        
        # Laad logo asynchroon in achtergrond
        def load_logo_async():
            try:
                Image, ImageTk, _, _ = get_PIL()  # Lazy load PIL
                logo_path = get_resource_path('Logo.png')
                if os.path.exists(logo_path):
                    logo_image = Image.open(logo_path)
                    # Resize logo als het te groot is (max 180x180)
                    logo_image.thumbnail((180, 180), Image.Resampling.LANCZOS)
                    self.logo_photo = ImageTk.PhotoImage(logo_image)
                    
                    # Update UI in main thread
                    def update_logo():
                        logo_label = tk.Label(self.welcome_frame, image=self.logo_photo, 
                                             bg=self.theme["BG_PRIMARY"])
                        logo_label.place(relx=0.5, rely=0.35, anchor="center")
                    
                    self.root.after(0, update_logo)
            except Exception as e:
                print(f"Logo loading error: {e}")
        
        threading.Thread(target=load_logo_async, daemon=True).start()
        
        welcome_text = "Welkom bij NVict Reader\n\nKlik op 'Openen' of druk op Ctrl+O om een PDF te laden."
        self.welcome_label = tk.Label(self.welcome_frame, text=welcome_text,
                                      font=Theme.FONT_HEADING, fg=self.theme["TEXT_SECONDARY"],
                                      bg=self.theme["BG_PRIMARY"], justify=tk.CENTER)
        # Positioneer tekst lager op kleine schermen om overlap met logo te vermijden
        text_y = 0.55 if self.root.winfo_height() > 600 else 0.65
        self.welcome_label.place(relx=0.5, rely=text_y, anchor="center")

        # Initialiseer de recente bestanden sectie (wordt ingevuld na laden instellingen)
        self.welcome_recent_frame = tk.Frame(self.welcome_frame, bg=self.theme["BG_PRIMARY"])
        self.root.after(100, self.create_welcome_recent_section)

        self.create_status_bar()

    def create_menubar(self):
        """Maak menubar met bewerken opties"""
        menubar = tk.Menu(self.root)
        self.root.config(menu=menubar)
        
        # Bestand menu
        file_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bestand", menu=file_menu)
        file_menu.add_command(label="Openen...", command=self.open_pdf, accelerator="Ctrl+O")
        # Recente bestanden submenu
        self.recent_menu = tk.Menu(file_menu, tearoff=0)
        file_menu.add_cascade(label="Recente bestanden", menu=self.recent_menu)
        self.update_recent_files_menu()
        file_menu.add_separator()
        file_menu.add_command(label="Afdrukken...", command=self.print_pdf, accelerator="Ctrl+P")
        file_menu.add_separator()
        file_menu.add_command(label="Sluiten", command=self.close_active_tab, accelerator="Ctrl+W")
        file_menu.add_command(label="Afsluiten", command=self.exit_application, accelerator="Ctrl+Q")
        
        # Bewerken menu
        edit_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Bewerken", menu=edit_menu)
        edit_menu.add_command(label="Kopieer tekst", command=self.copy_text, accelerator="Ctrl+C")
        edit_menu.add_command(label="Zoeken...", command=self.show_search_dialog, accelerator="Ctrl+F")
        edit_menu.add_separator()
        edit_menu.add_command(label="Pagina's exporteren...", command=self.export_pages)
        edit_menu.add_command(label="PDF's samenvoegen...", command=self.merge_pdfs)
        edit_menu.add_command(label="Pagina's roteren...", command=self.rotate_pages)
        
        # Beeld menu
        view_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Beeld", menu=view_menu)
        view_menu.add_command(label="Zoom in", command=self.zoom_in, accelerator="Ctrl++")
        view_menu.add_command(label="Zoom uit", command=self.zoom_out, accelerator="Ctrl+-")
        view_menu.add_command(label="Pasbreedte", command=lambda: self.set_zoom_mode("fit_width"))
        view_menu.add_separator()
        view_menu.add_command(label="Eerste pagina", command=self.first_page)
        view_menu.add_command(label="Vorige pagina", command=self.prev_page, accelerator="←")
        view_menu.add_command(label="Volgende pagina", command=self.next_page, accelerator="→")
        view_menu.add_command(label="Laatste pagina", command=self.last_page)
        view_menu.add_separator()
        view_menu.add_command(label="Volledig scherm", command=self.toggle_fullscreen, accelerator="F11")
        view_menu.add_separator()
        view_menu.add_command(label="Pagina's paneel", command=self.toggle_thumbnail_panel, accelerator="Ctrl+T")
        view_menu.add_command(label="Boek-modus", command=self.toggle_book_mode, accelerator="Ctrl+B")
        view_menu.add_separator()
        view_menu.add_command(label="Markeermodus", command=self.toggle_highlight_mode, accelerator="Ctrl+H")
        
        # Instellingen menu
        settings_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Instellingen", menu=settings_menu)
        settings_menu.add_command(label="Instellen als standaard PDF viewer", command=self.set_as_default_pdf)
        
        # Help menu
        help_menu = tk.Menu(menubar, tearoff=0)
        menubar.add_cascade(label="Help", menu=help_menu)
        help_menu.add_command(label="PDF Info", command=self.show_pdf_info)
        help_menu.add_separator()
        help_menu.add_command(label="Controleer op updates...", command=lambda: self.check_for_updates(silent=False))
        help_menu.add_separator()
        help_menu.add_command(label="Over NVict Reader", command=self.show_about)

    def create_modern_toolbar(self):
        # Buitenste container met vaste hoogte
        self.toolbar_outer = tk.Frame(self.root, bg=self.theme["BG_SECONDARY"], height=60,
                                      highlightbackground="#e0e0e0", highlightthickness=1)
        self.toolbar_outer.pack(side=tk.TOP, fill=tk.X, padx=20, pady=(20, 0))
        self.toolbar_outer.pack_propagate(False)

        # Canvas voor horizontaal scrollen
        self.toolbar_canvas = tk.Canvas(self.toolbar_outer, bg=self.theme["BG_SECONDARY"],
                                        height=58, highlightthickness=0, bd=0)
        self.toolbar_canvas.pack(side=tk.TOP, fill=tk.BOTH, expand=True)

        # Binnenframe waar de knoppen in komen
        self.toolbar_frame = tk.Frame(self.toolbar_canvas, bg=self.theme["BG_SECONDARY"], height=58)
        self._toolbar_win_id = self.toolbar_canvas.create_window(
            (0, 0), window=self.toolbar_frame, anchor="nw", height=58
        )

        # Scrollregio bijwerken als de inhoud verandert
        def _on_toolbar_configure(e):
            self.toolbar_canvas.configure(scrollregion=self.toolbar_canvas.bbox("all"))
            # Toon/verberg scrollbar op basis van beschikbare breedte
            canvas_w = self.toolbar_canvas.winfo_width()
            frame_w = self.toolbar_frame.winfo_reqwidth()
            if frame_w > canvas_w:
                self.toolbar_scrollbar.pack(side=tk.BOTTOM, fill=tk.X)
            else:
                self.toolbar_scrollbar.pack_forget()

        self.toolbar_frame.bind("<Configure>", _on_toolbar_configure)
        self.toolbar_canvas.bind("<Configure>",
            lambda e: self.toolbar_canvas.configure(scrollregion=self.toolbar_canvas.bbox("all")))

        # Muiswiel horizontaal scrollen op de toolbar
        def _on_toolbar_mousewheel(e):
            self.toolbar_canvas.xview_scroll(int(-1 * (e.delta / 120)), "units")

        self.toolbar_canvas.bind("<MouseWheel>", _on_toolbar_mousewheel)
        self.toolbar_frame.bind("<MouseWheel>", _on_toolbar_mousewheel)

        # Dunne horizontale scrollbar (alleen zichtbaar als nodig)
        self.toolbar_scrollbar = tk.Scrollbar(self.toolbar_outer, orient=tk.HORIZONTAL,
                                              command=self.toolbar_canvas.xview)
        self.toolbar_canvas.configure(xscrollcommand=self.toolbar_scrollbar.set)

        self._fill_toolbar()

    def _fill_toolbar(self):
        """Vul de toolbar met knoppen (kan herhaald worden bij themawissel)"""
        # Verwijder bestaande inhoud
        for widget in self.toolbar_frame.winfo_children():
            widget.destroy()
        self.toolbar_buttons = {}
        self.toolbar_frame.config(bg=self.theme["BG_SECONDARY"])

        f = self.toolbar_frame

        # ── Groep 1: Bestand ──
        self.open_btn = self.create_toolbar_button(f, " Openen", "open",
                                                   self.open_pdf, self.theme["ACCENT_COLOR"])
        self.save_btn = self.create_toolbar_button(f, " Opslaan", "save",
                                                   self.save_changes_to_pdf, self.theme["BG_SECONDARY"])
        self.send_btn = self.create_toolbar_button(f, " Doorsturen", "send",
                                                   self.send_pdf, self.theme["BG_SECONDARY"])
        self.print_btn = self.create_toolbar_button(f, " Printen", "print",
                                                    self.print_pdf, self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(f)

        # ── Groep 2: Bewerking ──
        self.copy_btn = self.create_toolbar_button(f, " Kopiëren", "copy",
                                                   self.copy_text, self.theme["BG_SECONDARY"])
        self.highlight_btn = self.create_toolbar_button(
            f, " Markeer", "marker", self.toggle_highlight_mode,
            self.theme["BG_SECONDARY"]
        )
        self.edit_btn = self.create_toolbar_button(f, " Bewerken", "toolbox",
                                                   self.show_edit_menu, self.theme["BG_SECONDARY"])
        self.type_text_btn = self.create_toolbar_button(f, " Tekst", "type-text",
                                                        self.toggle_text_annotate_mode, self.theme["BG_SECONDARY"])
        self.form_btn = self.create_toolbar_button(f, " Formulier", "form",
                                                   self.toggle_form_mode, self.theme["BG_SECONDARY"])
        self.search_btn = self.create_toolbar_button(f, " Zoeken", "search",
                                                     self.show_search_dialog, self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(f)

        # ── Groep 3: Weergave ──
        self.fullscreen_btn = self.create_toolbar_button(f, " Volledig scherm", "full-screen",
                                                         self.toggle_fullscreen, self.theme["BG_SECONDARY"])
        self.fit_width_btn = self.create_toolbar_button(f, "", "fit-width",
                                                        lambda: self.set_zoom_mode("fit_width"),
                                                        self.theme["BG_SECONDARY"])
        self.zoom_in_btn = self.create_toolbar_button(f, "", "zoom-in",
                                                      self.zoom_in, self.theme["BG_SECONDARY"])
        self.zoom_out_btn = self.create_toolbar_button(f, "", "zoom-out",
                                                       self.zoom_out, self.theme["BG_SECONDARY"])
        self.add_toolbar_separator(f)

        # ── Groep 4: Navigatie ──
        self.thumb_btn = self.create_toolbar_button(
            f, " Pagina's", "pages", self.toggle_thumbnail_panel,
            self.theme["ACCENT_COLOR"] if self.thumbnail_visible else self.theme["BG_SECONDARY"]
        )
        self.book_btn = self.create_toolbar_button(
            f, " Boek", "book", self.toggle_book_mode,
            self.theme["BG_SECONDARY"]
        )
        self.prev_btn = self.create_toolbar_button(f, "", "prev-page",
                                                   self.prev_page, self.theme["BG_SECONDARY"])
        self.next_btn = self.create_toolbar_button(f, "", "next-page",
                                                   self.next_page, self.theme["BG_SECONDARY"])

        page_frame = tk.Frame(f, bg=self.theme["BG_SECONDARY"])
        page_frame.pack(side=tk.LEFT, padx=5)

        self.page_var = tk.StringVar(value="1")
        page_entry = tk.Entry(page_frame, textvariable=self.page_var, width=5,
                              font=Theme.FONT_MAIN, justify=tk.CENTER,
                              bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                              relief="flat", bd=1)
        page_entry.pack(side=tk.LEFT)
        page_entry.bind("<Return>", self.go_to_page)

        self.total_pages_label = tk.Label(page_frame, text="/ 0", font=Theme.FONT_MAIN,
                                          fg=self.theme["TEXT_SECONDARY"],
                                          bg=self.theme["BG_SECONDARY"])
        self.total_pages_label.pack(side=tk.LEFT, padx=(5, 0))

    _TOOLTIPS = {
        "open":        "PDF openen  (Ctrl+O)",
        "close":       "Tab sluiten  (Ctrl+W)",
        "save":        "Opslaan  (Ctrl+S)",
        "print":       "Afdrukken  (Ctrl+P)",
        "search":      "Zoeken in document  (Ctrl+F)",
        "copy":        "Tekst kopiëren  (Ctrl+C)",
        "full-screen": "Volledig scherm  (F11)",
        "info":        "PDF-informatie",
        "toolbox":     "PDF bewerken",
        "theme":       "Wisselen tussen licht en donker thema",
        "zoom-in":     "Inzoomen  (Ctrl++)",
        "zoom-out":    "Uitzoomen  (Ctrl+-)",
        "fit-width":   "Pasbreedte",
        "prev-page":   "Vorige pagina  (←)",
        "next-page":   "Volgende pagina  (→)",
        "pages":       "Pagina's paneel tonen/verbergen  (Ctrl+T)",
        "marker":      "Markeermodus: selecteer tekst om te markeren  (Ctrl+H)",
        "book":        "Boek-modus: twee pagina's naast elkaar  (Ctrl+B)",
        "send":        "PDF doorsturen per e-mail",
        "form":        "Formuliervelden invullen aan/uit",
        "type-text":   "Tekst toevoegen op PDF  (dubbelklik op gewenste positie)",
    }

    def create_toolbar_button(self, parent, text, icon_name, command, bg_color):
        btn_frame = tk.Frame(parent, bg=bg_color, highlightthickness=0)
        btn_frame.pack(side=tk.LEFT, padx=5, pady=10)

        # Controleer of icon bestaat
        icon_image = self.icons.get(icon_name)

        # Emoji fallbacks voor ontbrekende iconen
        emoji_fallbacks = {
            "open": "📂",
            "close": "✕",
            "save": "💾",
            "print": "🖨️",
            "copy": "📋",
            "search": "🔍",
            "zoom-in": "🔍+",
            "zoom-out": "🔍-",
            "reset": "↺",
            "fit-width": "⬌",
            "prev-page": "◄",
            "next-page": "►",
            "first-page": "⏮",
            "last-page": "⏭",
            "info": "ℹ️",
            "toolbox": "🛠️",
            "full-screen": "⛶",
            "theme": "◐",
            "send": "✉",
            "form": "📝",
            "type-text": "T",
            "pages": "📄",
            "marker": "🖍️",
            "book": "📖"
        }

        # Als icon bestaat, gebruik icon + tekst
        if icon_image:
            btn = tk.Button(btn_frame, text=text, image=icon_image,
                           compound=tk.LEFT, command=command, font=Theme.FONT_MAIN,
                           bg=bg_color, fg=self.theme["TEXT_PRIMARY"],
                           activebackground=self.theme["ACCENT_COLOR"],
                           activeforeground="#ffffff", relief="flat", bd=0,
                           padx=10, pady=5, cursor="hand2")
        else:
            # Geen icon - gebruik emoji fallback
            display_text = text
            if not text and icon_name in emoji_fallbacks:
                display_text = emoji_fallbacks[icon_name]
            elif text and icon_name in emoji_fallbacks:
                # Voeg emoji toe voor de tekst
                display_text = emoji_fallbacks[icon_name] + text

            btn = tk.Button(btn_frame, text=display_text, command=command, font=Theme.FONT_MAIN,
                           bg=bg_color, fg=self.theme["TEXT_PRIMARY"],
                           activebackground=self.theme["ACCENT_COLOR"],
                           activeforeground="#ffffff", relief="flat", bd=0,
                           padx=10, pady=5, cursor="hand2")

        btn.pack()

        # Accentlijn onder de knop als actieve indicator
        indicator = tk.Frame(btn_frame, bg=bg_color, height=3)
        indicator.pack(fill=tk.X, padx=2)

        btn._is_active = False
        btn._default_bg = bg_color
        btn._indicator = indicator

        def on_enter(e):
            try:
                btn.configure(bg=self.theme["ACCENT_COLOR"], fg="#ffffff")
            except (tk.TclError, Exception):
                pass
        def on_leave(e):
            try:
                if getattr(btn, '_is_active', False):
                    btn.configure(bg=self.theme["ACCENT_COLOR"], fg="#ffffff")
                else:
                    btn.configure(bg=btn._default_bg, fg=self.theme["TEXT_PRIMARY"])
            except (tk.TclError, Exception):
                pass

        btn.bind("<Enter>", on_enter)
        btn.bind("<Leave>", on_leave)

        # Muiswiel doorsturen naar toolbar canvas voor horizontaal scrollen
        def _fwd_mousewheel(e):
            self.toolbar_canvas.xview_scroll(int(-1 * (e.delta / 120)), "units")
        btn.bind("<MouseWheel>", _fwd_mousewheel)
        btn_frame.bind("<MouseWheel>", _fwd_mousewheel)
        indicator.bind("<MouseWheel>", _fwd_mousewheel)

        self.toolbar_buttons[icon_name] = (btn, text)

        if icon_name in self._TOOLTIPS:
            Tooltip(btn, self._TOOLTIPS[icon_name])

        return btn

    def _set_toolbar_button_active(self, icon_name, active):
        """Zet een toolbar-knop visueel actief of inactief."""
        try:
            btn_tuple = self.toolbar_buttons.get(icon_name)
            if btn_tuple:
                btn = btn_tuple[0]
                btn._is_active = active
                indicator = getattr(btn, '_indicator', None)
                if active:
                    btn.config(bg=self.theme["ACCENT_COLOR"], fg="white")
                    if indicator:
                        indicator.config(bg=self.theme["WARNING_COLOR"])
                else:
                    btn.config(bg=btn._default_bg, fg=self.theme["TEXT_PRIMARY"])
                    if indicator:
                        indicator.config(bg=btn._default_bg)
        except (tk.TclError, Exception):
            pass

    def add_toolbar_separator(self, parent):
        separator = tk.Frame(parent, bg=self.theme["TEXT_SECONDARY"], width=1, height=40)
        separator.pack(side=tk.LEFT, padx=10, pady=10)

    def create_status_bar(self):
        self.status_bar = tk.Frame(self.root, bg=self.theme["BG_SECONDARY"], height=30,
                                   highlightbackground="#e0e0e0", highlightthickness=1)
        self.status_bar.pack(side=tk.BOTTOM, fill=tk.X, padx=20, pady=(0, 20))
        self.status_bar.pack_propagate(False)
        self._fill_status_bar()

    def _fill_status_bar(self):
        """Vul de statusbalk met labels (kan herhaald worden bij themawissel)"""
        for widget in self.status_bar.winfo_children():
            widget.destroy()
        self.status_bar.config(bg=self.theme["BG_SECONDARY"])

        self.status_label = tk.Label(self.status_bar, text="Klaar", font=Theme.FONT_SMALL,
                                     fg=self.theme["TEXT_SECONDARY"], bg=self.theme["BG_SECONDARY"],
                                     anchor="w")
        self.status_label.pack(side=tk.LEFT, padx=10, fill=tk.X, expand=True)

        copyright_label = tk.Label(self.status_bar,
                                   text=f"© {self.get_current_year()} NVict Service - www.nvict.nl",
                                   font=Theme.FONT_SMALL, fg=self.theme["TEXT_SECONDARY"],
                                   bg=self.theme["BG_SECONDARY"], anchor="e", cursor="hand2")
        copyright_label.pack(side=tk.RIGHT, padx=10)
        copyright_label.bind("<Button-1>", lambda e: webbrowser.open("https://www.nvict.nl/software.html"))

    def get_current_year(self):
        """Get current year"""
        from datetime import datetime
        return datetime.now().year

    def setup_shortcuts(self):
        self.root.bind("<Control-o>", lambda e: self.open_pdf())
        self.root.bind("<Control-p>", lambda e: self.print_pdf())
        self.root.bind("<Control-w>", lambda e: self.close_active_tab())
        self.root.bind("<Control-q>", lambda e: self.exit_application())
        self.root.bind("<Control-plus>", lambda e: self.zoom_in())
        self.root.bind("<Control-minus>", lambda e: self.zoom_out())
        self.root.bind("<Control-c>", lambda e: self.copy_text())
        self.root.bind("<Control-f>", lambda e: self.show_search_dialog())
        self.root.bind("<Control-s>", lambda e: self.save_changes_to_pdf())
        self.root.bind("<Left>", lambda e: self.prev_page())
        self.root.bind("<Right>", lambda e: self.next_page())
        self.root.bind("<F11>", lambda e: self.toggle_fullscreen())
        self.root.bind("<Escape>", lambda e: self.exit_fullscreen())
        self.root.bind("<Control-t>", lambda e: self.toggle_thumbnail_panel())
        self.root.bind("<Control-h>", lambda e: self.toggle_highlight_mode())
        self.root.bind("<Control-b>", lambda e: self.toggle_book_mode())

    def toggle_fullscreen(self):
        """Schakel PDF-presentatiemodus aan of uit"""
        if getattr(self, 'is_fullscreen', False):
            self.exit_fullscreen()
        else:
            self.enter_fullscreen()

    def enter_fullscreen(self):
        """Open de huidige PDF-pagina in een volledig scherm presentatiemodus"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showinfo("Geen PDF", "Open eerst een PDF om de presentatiemodus te gebruiken.")
            return

        self.is_fullscreen = True
        if hasattr(self, 'fullscreen_btn'):
            self.fullscreen_btn.config(text=" Volledig scherm")

        self.fs_tab = tab
        self.fs_page = tab.current_page

        # Maak volledig scherm venster
        self.fs_window = tk.Toplevel(self.root)
        self.fs_window.attributes('-fullscreen', True)
        self.fs_window.configure(bg='black')
        self.fs_window.focus_set()
        self.fs_window.lift()

        # Canvas voor PDF weergave
        self.fs_canvas = tk.Canvas(self.fs_window, bg='black', highlightthickness=0)
        self.fs_canvas.pack(fill=tk.BOTH, expand=True)

        # Hint-banner bovenaan
        self._fs_hint_after_id = None
        self.fs_hint_frame = tk.Frame(self.fs_window, bg='#1e1e1e')
        self.fs_hint_frame.place(relx=0, rely=0, relwidth=1)
        tk.Label(
            self.fs_hint_frame,
            text="  Druk op  Escape  of  F11  om volledig scherm te verlaten"
                 "     ·     ←  →  voor pagina navigatie  ",
            bg='#1e1e1e', fg='#dddddd',
            font=('Segoe UI', 14), pady=14
        ).pack()
        self._fs_schedule_hide(4000)

        # Paginanummer indicatie (subtiel onderaan)
        self.fs_label = tk.Label(
            self.fs_window, text="", bg='black', fg='#666666',
            font=('Segoe UI', 10)
        )
        self.fs_label.place(relx=0.5, rely=0.97, anchor='s')

        # Sneltoetsen
        self.fs_window.bind('<Escape>', lambda e: self.exit_fullscreen())
        self.fs_window.bind('<F11>', lambda e: self.exit_fullscreen())
        self.fs_window.bind('<Left>', lambda e: self._fs_navigate(-1))
        self.fs_window.bind('<Right>', lambda e: self._fs_navigate(1))
        self.fs_window.bind('<Prior>', lambda e: self._fs_navigate(-1))
        self.fs_window.bind('<Next>', lambda e: self._fs_navigate(1))
        self.fs_window.bind('<Configure>', lambda e: self._render_fs_page())
        self.fs_window.bind('<Motion>', self._fs_on_mouse_move)
        self.fs_window.protocol('WM_DELETE_WINDOW', self.exit_fullscreen)

        self._render_fs_page()

    def _fs_schedule_hide(self, delay_ms=3000):
        """Plan het verbergen van de hint-banner na een vertraging"""
        if self._fs_hint_after_id:
            self.fs_window.after_cancel(self._fs_hint_after_id)
        self._fs_hint_after_id = self.fs_window.after(delay_ms, self._hide_fs_hint)

    def _hide_fs_hint(self):
        """Verberg de hint-banner"""
        self._fs_hint_after_id = None
        if hasattr(self, 'fs_hint_frame') and self.fs_hint_frame.winfo_exists():
            self.fs_hint_frame.place_forget()

    def _fs_on_mouse_move(self, event):
        """Toon de hint-banner opnieuw als de muis de bovenkant nadert"""
        if event.y < 80:
            if hasattr(self, 'fs_hint_frame') and self.fs_hint_frame.winfo_exists():
                self.fs_hint_frame.place(relx=0, rely=0, relwidth=1)
                self._fs_schedule_hide(3000)

    def _render_fs_page(self):
        """Render de huidige pagina gecentreerd in het volledig scherm venster"""
        if not hasattr(self, 'fs_window') or not self.fs_window.winfo_exists():
            return
        if not hasattr(self, 'fs_tab') or not self.fs_tab.pdf_document:
            return

        screen_w = self.fs_window.winfo_width()
        screen_h = self.fs_window.winfo_height()
        if screen_w < 10 or screen_h < 10:
            return

        fitz = get_fitz()
        Image, ImageTk, _, _ = get_PIL()

        page = self.fs_tab.pdf_document[self.fs_page]
        page_rect = page.bound()

        # Zoom berekenen zodat de pagina het scherm vult (letterbox)
        zoom = min(screen_w / page_rect.width, screen_h / page_rect.height)

        mat = fitz.Matrix(zoom, zoom)
        pix = page.get_pixmap(matrix=mat)
        pil_image = Image.open(io.BytesIO(pix.tobytes("ppm")))

        self.fs_photo = ImageTk.PhotoImage(pil_image)

        self.fs_canvas.delete("all")
        self.fs_canvas.create_image(screen_w // 2, screen_h // 2, anchor='center', image=self.fs_photo)

        total = len(self.fs_tab.pdf_document)
        self.fs_label.config(text=f"Pagina {self.fs_page + 1} / {total}  ·  Escape om te sluiten")

    def _fs_navigate(self, delta):
        """Navigeer naar vorige of volgende pagina in presentatiemodus"""
        if not hasattr(self, 'fs_tab') or not self.fs_tab.pdf_document:
            return
        new_page = self.fs_page + delta
        if 0 <= new_page < len(self.fs_tab.pdf_document):
            self.fs_page = new_page
            self._render_fs_page()

    def exit_fullscreen(self):
        """Verlaat de presentatiemodus"""
        if not getattr(self, 'is_fullscreen', False):
            return
        self.is_fullscreen = False
        if hasattr(self, 'fullscreen_btn'):
            self.fullscreen_btn.config(text=" Volledig scherm")
        if hasattr(self, 'fs_window') and self.fs_window.winfo_exists():
            self.fs_window.destroy()
        # Synchroniseer de hoofdtab naar de bekeken pagina
        if hasattr(self, 'fs_tab') and hasattr(self, 'fs_page'):
            self.fs_tab.current_page = self.fs_page
            self.scroll_to_page(self.fs_tab, self.fs_page)
            self.update_ui_state()

    def toggle_theme(self):
        """Wissel tussen licht en donker thema en pas de UI direct aan"""
        old_theme = self.theme

        new_name = "Donker" if self.theme == Theme.LIGHT else "Licht"
        self.update_settings['theme'] = new_name
        self.save_update_settings()
        self.apply_theme()  # update self.theme

        # Toolbar containers herkleurent + vullen
        self.toolbar_outer.config(bg=self.theme["BG_SECONDARY"],
                                  highlightbackground="#e0e0e0")
        self.toolbar_canvas.config(bg=self.theme["BG_SECONDARY"])
        self._fill_toolbar()
        self._fill_status_bar()

        # Herthem welcome frame en open PDF-tabs
        self._retheme_widget_tree(self.welcome_frame, old_theme, self.theme)
        for tab_id in self.notebook.tabs():
            tab = self.notebook.nametowidget(tab_id)
            if isinstance(tab, PDFTab):
                tab.theme = self.theme
                self._retheme_widget_tree(tab, old_theme, self.theme)
                self.display_page(tab)

        # Herlaad icons met juiste kleur (licht/donker inversie)
        self._start_icon_load_thread()
        self.update_ui_state()

    def _retheme_widget_tree(self, widget, old_theme, new_theme):
        """Pas kleuren van alle widgets recursief aan op basis van kleurmapping"""
        color_map = {
            old_theme["BG_PRIMARY"].lower(): new_theme["BG_PRIMARY"],
            old_theme["BG_SECONDARY"].lower(): new_theme["BG_SECONDARY"],
            old_theme["TEXT_PRIMARY"].lower(): new_theme["TEXT_PRIMARY"],
            old_theme["TEXT_SECONDARY"].lower(): new_theme["TEXT_SECONDARY"],
        }
        for attr in ('bg', 'fg'):
            try:
                current = widget.cget(attr).lower()
                if current in color_map:
                    widget.config(**{attr: color_map[current]})
            except Exception:
                pass
        for child in widget.winfo_children():
            self._retheme_widget_tree(child, old_theme, new_theme)

    def load_update_settings(self):
        """Laad update instellingen van bestand"""
        self.update_settings = {
            'auto_check': True,  # Automatisch controleren bij opstarten
            'auto_download': False,  # Automatisch downloaden (standaard uit)
            'last_check': None,
            'window_geometry': None,  # Laatst gebruikte schermgrootte
            'window_state': 'normal',  # normal of zoomed (maximized)
            'recent_files': []  # Lijst van recent geopende bestanden
        }
        
        try:
            settings_path = get_settings_path()
            if os.path.exists(settings_path):
                with open(settings_path, 'r') as f:
                    saved_settings = json.load(f)
                    self.update_settings.update(saved_settings)
        except Exception:
            pass  # Gebruik default settings
    
    def save_update_settings(self):
        """Sla update instellingen op naar bestand"""
        try:
            settings_path = get_settings_path()
            with open(settings_path, 'w') as f:
                json.dump(self.update_settings, f, indent=2)
        except Exception:
            pass  # Stille fout

    def add_to_recent_files(self, file_path):
        """Voeg een bestandspad toe aan de lijst van recente bestanden (max 10)"""
        recent = self.update_settings.get('recent_files', [])
        # Verwijder als het al in de lijst staat (om duplicaten te voorkomen)
        if file_path in recent:
            recent.remove(file_path)
        # Voeg bovenaan de lijst toe
        recent.insert(0, file_path)
        # Beperk tot 10 bestanden
        self.update_settings['recent_files'] = recent[:10]
        self.save_update_settings()
        # Ververs het menu
        self.update_recent_files_menu()

    def update_recent_files_menu(self):
        """Herbouw het 'Recente bestanden' submenu op basis van huidige lijst"""
        if not hasattr(self, 'recent_menu'):
            return
        self.recent_menu.delete(0, 'end')
        recent = self.update_settings.get('recent_files', [])
        # Filter bestanden die niet meer bestaan
        recent = [f for f in recent if os.path.exists(f)]
        if not recent:
            self.recent_menu.add_command(label="(Geen recente bestanden)", state='disabled')
        else:
            for path in recent:
                # Toon alleen bestandsnaam, maar open het volledige pad
                label = os.path.basename(path)
                self.recent_menu.add_command(
                    label=label,
                    command=lambda p=path: self.add_new_tab(p)
                )
            self.recent_menu.add_separator()
            self.recent_menu.add_command(
                label="Lijst wissen",
                command=self.clear_recent_files
            )
        # Ververs ook het welkomstscherm
        if hasattr(self, 'welcome_frame') and self.welcome_frame.winfo_exists():
            self.create_welcome_recent_section()

    def clear_recent_files(self):
        """Leeg de lijst van recente bestanden"""
        self.update_settings['recent_files'] = []
        self.save_update_settings()
        self.update_recent_files_menu()

    def create_welcome_recent_section(self):
        """Maak of ververs de 'Recente bestanden' sectie op het welkomstscherm"""
        # Verwijder bestaande sectie als die al bestaat
        if hasattr(self, 'welcome_recent_frame') and self.welcome_recent_frame.winfo_exists():
            self.welcome_recent_frame.destroy()

        recent = [f for f in self.update_settings.get('recent_files', []) if os.path.exists(f)]

        if not recent:
            return  # Geen recente bestanden - sectie niet tonen

        # Hoofdcontainer, gecentreerd onder de welkomsttekst
        self.welcome_recent_frame = tk.Frame(self.welcome_frame, bg=self.theme["BG_PRIMARY"])
        self.welcome_recent_frame.place(relx=0.5, rely=0.63, anchor="n", relwidth=0.52)

        # Sectie-titel
        tk.Label(
            self.welcome_recent_frame,
            text="Recente bestanden",
            font=Theme.FONT_HEADING,
            fg=self.theme["TEXT_SECONDARY"],
            bg=self.theme["BG_PRIMARY"],
            anchor="w"
        ).pack(fill=tk.X, pady=(0, 6))

        # Toon maximaal 5 recente bestanden
        for path in recent[:5]:
            name = os.path.basename(path)
            # Verkorte mapweergave
            directory = os.path.dirname(path)
            if len(directory) > 55:
                directory = "..." + directory[-52:]

            item_frame = tk.Frame(
                self.welcome_recent_frame,
                bg=self.theme["BG_SECONDARY"],
                cursor="hand2"
            )
            item_frame.pack(fill=tk.X, pady=2)

            name_label = tk.Label(
                item_frame,
                text=f"📄  {name}",
                font=Theme.FONT_MAIN,
                fg=self.theme["ACCENT_COLOR"],
                bg=self.theme["BG_SECONDARY"],
                anchor="w",
                cursor="hand2"
            )
            name_label.pack(side=tk.LEFT, padx=10, pady=6)

            dir_label = tk.Label(
                item_frame,
                text=directory,
                font=Theme.FONT_SMALL,
                fg=self.theme["TEXT_SECONDARY"],
                bg=self.theme["BG_SECONDARY"],
                anchor="e",
                cursor="hand2"
            )
            dir_label.pack(side=tk.RIGHT, padx=10, pady=6)

            # Klik opent het bestand
            for widget in (item_frame, name_label, dir_label):
                widget.bind("<Button-1>", lambda e, p=path: self.add_new_tab(p))

            # Hover-effect
            def on_enter(e, f=item_frame, nl=name_label, dl=dir_label):
                hover = self.theme["BG_PRIMARY"]
                f.config(bg=hover)
                nl.config(bg=hover)
                dl.config(bg=hover)

            def on_leave(e, f=item_frame, nl=name_label, dl=dir_label):
                normal = self.theme["BG_SECONDARY"]
                f.config(bg=normal)
                nl.config(bg=normal)
                dl.config(bg=normal)

            item_frame.bind("<Enter>", on_enter)
            item_frame.bind("<Leave>", on_leave)
            name_label.bind("<Enter>", on_enter)
            name_label.bind("<Leave>", on_leave)
            dir_label.bind("<Enter>", on_enter)
            dir_label.bind("<Leave>", on_leave)

    def check_for_updates_on_startup(self):
        """Controleer automatisch op updates bij opstarten (in achtergrond)"""
        if not self.update_settings.get('auto_check', True):
            return
        
        # Doe check in aparte thread om UI niet te blokkeren
        def background_check():
            import time
            time.sleep(2)  # Wacht 2 seconden na opstarten
            self.root.after(0, lambda: self.check_for_updates(silent=True))
        
        thread = threading.Thread(target=background_check, daemon=True)
        thread.start()

    def get_active_tab(self):
        try:
            current_tab_id = self.notebook.select()
            if current_tab_id:
                return self.notebook.nametowidget(current_tab_id)
        except:
            pass
        return None

    def on_tab_change(self, event=None):
        self.update_ui_state()
        self.root.after(50, self.update_thumbnails)

    def update_ui_state(self):
        tab = self.get_active_tab()
        has_pdf = isinstance(tab, PDFTab)

        ui_btns = [self.print_btn, self.zoom_in_btn,
                   self.zoom_out_btn, self.fit_width_btn, self.prev_btn,
                   self.next_btn, self.copy_btn, self.search_btn, self.edit_btn,
                   self.highlight_btn, self.book_btn,
                   self.send_btn, self.form_btn, self.type_text_btn,
                   self.fullscreen_btn]
        for btn in ui_btns:
            try:
                btn.config(state=tk.NORMAL if has_pdf else tk.DISABLED)
            except (tk.TclError, Exception):
                pass

        # Opslaan-knop: alleen actief als er wijzigingen zijn
        self._update_save_button_state()

        if has_pdf:
            self.page_var.set(str(tab.current_page + 1))
            self.total_pages_label.config(text=f"/ {len(tab.pdf_document)}")
            # Update statusbar met zoom en beveiligingsinformatie
            self.update_status_with_security(tab)
        else:
            self.page_var.set("1")
            self.total_pages_label.config(text="/ 0")
            self.status_label.config(text="Geen document geopend")

    def update_status_with_security(self, tab):
        """Update statusbar met zoom en beveiligingsinformatie"""
        status_text = f"Zoom: {int(tab.zoom_level * 100)}%"

        # Voeg beveiligingsinformatie toe indien beschikbaar
        if hasattr(tab, 'security_info') and tab.security_info:
            status_text += f"  ·  {tab.security_info}"

        self.status_label.config(text=status_text)

    def open_pdf(self):
        file_path = filedialog.askopenfilename(
            title="Selecteer een PDF bestand",
            filetypes=[("PDF Bestanden", "*.pdf"), ("Alle Bestanden", "*.*")]
        )
        if file_path:
            self.add_new_tab(file_path)

    def add_new_tab(self, file_path):
        try:
            fitz = get_fitz()  # Lazy load PyMuPDF
            # Normaliseer bestandspad voor vergelijking
            file_path = os.path.abspath(file_path)
            
            # Check of dit bestand al open is in een bestaande tab
            for tab_id in self.notebook.tabs():
                tab = self.notebook.nametowidget(tab_id)
                if isinstance(tab, PDFTab):
                    if os.path.abspath(tab.file_path) == file_path:
                        # Bestand is al open - switch naar die tab
                        self.notebook.select(tab)
                        
                        # Breng venster naar voren als het geminimaliseerd is
                        if self.root.state() == 'iconic':
                            self.root.deiconify()
                        
                        # Breng venster naar voren
                        self.root.lift()
                        self.root.focus_force()
                        
                        # Toon melding
                        self.status_label.config(text=f"Bestand is al geopend: {os.path.basename(file_path)}")
                        return
            
            # Breng venster naar voren als het geminimaliseerd is (voor nieuwe bestanden)
            if self.root.state() == 'iconic':
                self.root.deiconify()
            
            # Breng venster naar voren
            self.root.lift()
            self.root.focus_force()
            
            # Probeer eerst te openen om te controleren of wachtwoord nodig is
            test_doc = fitz.open(file_path)
            
            # Controleer of document beveiligd is
            if test_doc.needs_pass:
                test_doc.close()
                
                # Vraag wachtwoord
                password = self.ask_password(file_path)
                
                if password is None:
                    # Gebruiker heeft geannuleerd
                    return
                
                # Probeer te openen met wachtwoord
                test_doc = fitz.open(file_path)
                auth_result = test_doc.authenticate(password)
                
                if not auth_result:
                    test_doc.close()
                    messagebox.showerror("Fout", 
                        "Onjuist wachtwoord!\n\nKan het PDF bestand niet openen.")
                    return
                
                # Wachtwoord is correct, sla het op voor later gebruik
                # (wordt gebruikt bij PDFTab aanmaak)
                self.temp_password = password
            else:
                self.temp_password = None
            
            test_doc.close()
            
            # Nu de tab aanmaken
            if self.welcome_frame.winfo_ismapped():
                self.notebook.forget(self.welcome_frame)

            tab = PDFTab(self.notebook, file_path, self.theme, getattr(self, 'temp_password', None))
            self.notebook.add(tab, text=os.path.basename(file_path), padding=5)
            self.notebook.select(tab)
            self.display_page(tab)
            # Voeg toe aan recente bestanden
            self.add_to_recent_files(file_path)
            # Thumbnails bijwerken (na korte vertraging zodat canvas op orde is)
            self.root.after(200, self.update_thumbnails)
            
            # Bind events
            tab.canvas.bind("<Configure>", lambda e, t=tab: self.on_resize(e, t))
            tab.canvas.bind("<Button-1>", lambda e, t=tab: self.on_click(e, t))
            tab.canvas.bind("<B1-Motion>", lambda e, t=tab: self.on_drag(e, t))
            tab.canvas.bind("<ButtonRelease-1>", lambda e, t=tab: self.on_release(e, t))
            tab.canvas.bind("<Motion>", lambda e, t=tab: self.on_mouse_move(e, t))
            
            # Muiswiel
            tab.canvas.bind("<MouseWheel>", lambda e, t=tab: self.on_mousewheel(e, t))
            tab.canvas.bind("<Button-4>", lambda e, t=tab: self.on_mousewheel(e, t))
            tab.canvas.bind("<Button-5>", lambda e, t=tab: self.on_mousewheel(e, t))
            
            # Wis tijdelijk wachtwoord
            self.temp_password = None
            
        except Exception as e:
            messagebox.showerror("Fout", f"Kan PDF niet openen:\n{str(e)}")
    
    def ask_password(self, file_path):
        """Toon dialoog om wachtwoord op te vragen voor beveiligde PDF"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Wachtwoord vereist")
        dialog.geometry("450x280")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur
        header_frame = tk.Frame(dialog, bg=self.theme["WARNING_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔒 Wachtwoord Vereist", font=("Segoe UI", 14, "bold"),
                bg=self.theme["WARNING_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        tk.Label(content_frame, 
                text=f"Dit PDF bestand is beveiligd met een wachtwoord.\n\nBestand: {os.path.basename(file_path)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_PRIMARY"],
                justify=tk.LEFT,
                wraplength=380).pack(pady=(0, 20))
        
        tk.Label(content_frame, text="Wachtwoord:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        password_var = tk.StringVar()
        password_entry = tk.Entry(content_frame, textvariable=password_var, 
                                 font=("Segoe UI", 10), show="●", width=40)
        password_entry.pack(pady=8, fill=tk.X)
        password_entry.focus()
        
        result = {"password": None}
        
        def on_ok():
            result["password"] = password_var.get()
            dialog.destroy()
        
        def on_cancel():
            result["password"] = None
            dialog.destroy()
        
        # Footer met knoppen
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="OK", command=on_ok,
                 bg=self.theme["WARNING_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=on_cancel,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Enter key binding
        password_entry.bind("<Return>", lambda e: on_ok())
        
        # Wacht tot dialoog gesloten is
        self.root.wait_window(dialog)
        
        return result["password"]

    def on_mousewheel(self, event, tab):
        if event.num == 4 or event.delta > 0:
            tab.canvas.yview_scroll(-1, "units")
        elif event.num == 5 or event.delta < 0:
            tab.canvas.yview_scroll(1, "units")

    def close_active_tab(self):
        active_tab = self.get_active_tab()
        if isinstance(active_tab, PDFTab):
            active_tab.close_document()
            self.notebook.forget(active_tab)
            if len(self.notebook.tabs()) == 0:
                self.notebook.add(self.welcome_frame)
                # Ververs recente bestanden op welkomstscherm
                self.create_welcome_recent_section()
            self.update_ui_state()
    
    def on_resize(self, event, tab):
        if tab.zoom_mode == "fit_width":
            self.display_page(tab)

    def display_page(self, tab):
        if not tab or not tab.pdf_document:
            return

        fitz = get_fitz()
        Image, ImageTk, _, _ = get_PIL()

        # ── Clear ──
        tab.canvas.delete("all")
        tab.text_words = []
        tab.selected_text = ""
        tab.links = []
        tab.page_regions = {}
        book_mode   = getattr(tab, 'book_mode', False)
        page_gap    = 12   # horizontale ruimte tussen pagina's in boek-modus
        x_margin    = 20
        y_start     = 20
        page_spacing = 20

        n = len(tab.pdf_document)

        # ── Zoom berekening ──
        if tab.zoom_mode == "fit_width":
            canvas_width = max(tab.canvas.winfo_width() - 40, 100)
            if book_mode and n > 1:
                # Twee pagina's naast elkaar: zoom zo dat beide passen
                max_w = max(tab.pdf_document[i].bound().width for i in range(n))
                if max_w > 0:
                    tab.zoom_level = canvas_width / (max_w * 2 + page_gap)
            else:
                page_width = tab.pdf_document[0].bound().width
                if page_width > 0:
                    tab.zoom_level = canvas_width / page_width

        # ── Bereken pagina-layout: (page_num, x, y) per pagina ──
        page_layout = []   # [(page_num, x_offset, y_offset), ...]
        tab.page_positions = []  # index == page_num → y canvas positie

        if book_mode:
            y = y_start
            p = 0
            while p < n:
                left_num  = p
                right_num = p + 1 if p + 1 < n else None
                left_h  = tab.pdf_document[left_num].bound().height  * tab.zoom_level
                left_w  = tab.pdf_document[left_num].bound().width   * tab.zoom_level
                right_h = tab.pdf_document[right_num].bound().height * tab.zoom_level if right_num is not None else 0
                row_height = max(left_h, right_h)

                page_layout.append((left_num, x_margin, y))
                tab.page_positions.append(y)

                if right_num is not None:
                    rx = x_margin + int(left_w) + page_gap
                    page_layout.append((right_num, rx, y))
                    tab.page_positions.append(y)
                    p += 2
                else:
                    p += 1

                y += int(row_height) + page_spacing
            total_y = y
        else:
            y = y_start
            for page_num in range(n):
                page_layout.append((page_num, x_margin, y))
                tab.page_positions.append(y)
                ph = tab.pdf_document[page_num].bound().height * tab.zoom_level
                y += int(ph) + page_spacing
            total_y = y

        # ── Initialiseer page_images ──
        if not hasattr(tab, 'page_pil_images'):
            tab.page_pil_images = {}
        if not hasattr(tab, 'page_images'):
            tab.page_images = []
        # Zorg dat de lijst lang genoeg is
        while len(tab.page_images) < n:
            tab.page_images.append(None)

        # ── Render elke pagina ──
        for (page_num, x_offset, y_offset) in page_layout:
            page = tab.pdf_document[page_num]
            mat  = fitz.Matrix(tab.zoom_level, tab.zoom_level)
            pix  = page.get_pixmap(matrix=mat)
            pil_image = Image.open(io.BytesIO(pix.tobytes("ppm")))
            img_width, img_height = pil_image.size

            # Bewaar voor highlights / navigatie
            if page_num == tab.current_page:
                tab.current_image  = pil_image.copy()
                tab.page_offset_x  = x_offset
                tab.page_offset_y  = y_offset
            tab.page_pil_images[page_num] = pil_image.copy()
            tab.page_regions[page_num]    = (x_offset, y_offset,
                                              x_offset + img_width,
                                              y_offset + img_height)

            photo = ImageTk.PhotoImage(pil_image)
            tab.page_images[page_num] = photo

            # Teken pagina
            tab.canvas.create_image(x_offset, y_offset, anchor="nw",
                                    image=photo, tags=f"page_{page_num}")

            # Paginanummer label
            tab.canvas.create_text(
                x_offset + img_width // 2, y_offset - 5,
                text=f"Pagina {page_num + 1} / {n}",
                font=("Segoe UI", 9), fill=self.theme["TEXT_SECONDARY"]
            )

            # Extract tekst
            for wi in page.get_text("words"):
                x0, y0, x1, y1, text = wi[0], wi[1], wi[2], wi[3], wi[4]
                tab.text_words.append((
                    text,
                    x0 * tab.zoom_level + x_offset,
                    y0 * tab.zoom_level + y_offset,
                    x1 * tab.zoom_level + x_offset,
                    y1 * tab.zoom_level + y_offset,
                ))

            # Extract links
            for link in page.get_links():
                if 'from' in link:
                    rect = link['from']
                    tab.links.append({
                        'page_num': page_num,
                        'rect': (
                            rect.x0 * tab.zoom_level + x_offset,
                            rect.y0 * tab.zoom_level + y_offset,
                            rect.x1 * tab.zoom_level + x_offset,
                            rect.y1 * tab.zoom_level + y_offset,
                        ),
                        'uri':  link.get('uri',  None),
                        'page': link.get('page', None),
                        'kind': link.get('kind', 0),
                    })

            # Scheidingslijn onder pagina
            sep_y = y_offset + img_height + page_spacing // 2
            tab.canvas.create_line(
                x_offset, sep_y, x_offset + img_width, sep_y,
                fill=self.theme["TEXT_SECONDARY"], width=1, dash=(2, 4)
            )

        # ── Scrollregion ──
        max_width = max(
            (tab.page_regions[p][2] for p in tab.page_regions),
            default=400
        ) + x_margin
        tab.canvas.configure(scrollregion=(0, 0, max_width, total_y + 20))

        # ── Navigeer naar gewenste pagina ──
        if hasattr(tab, 'scroll_to_page') and tab.scroll_to_page is not None:
            self.scroll_to_page(tab, tab.scroll_to_page)
            tab.scroll_to_page = None

        self.update_ui_state()
        # Thumbnails bijwerken na render (actuele pagina markeren)
        self.root.after(50, self.update_thumbnails)

        # Formuliervelden: teken lichtblauwe markeringen als er velden zijn
        if hasattr(tab, 'form_mode'):
            self._draw_form_field_highlights(tab)
            # Overlay widgets opnieuw tekenen als formuliermodus actief is
            if tab.form_mode:
                self._save_form_widget_values(tab)
                self._create_form_overlays(tab)

        # Tekst-annotaties opnieuw tekenen
        if hasattr(tab, 'text_annotations') and tab.text_annotations:
            self._draw_text_annotations(tab)

    def scroll_to_page(self, tab, page_num):
        """Scroll naar een specifieke pagina"""
        if not hasattr(tab, 'page_positions') or page_num >= len(tab.page_positions):
            return
        
        y_pos = tab.page_positions[page_num]
        
        # Scroll canvas naar deze positie
        total_height = int(tab.canvas.cget("scrollregion").split()[3])
        canvas_height = tab.canvas.winfo_height()
        
        if total_height > canvas_height:
            # Bereken fractie voor scrollpositie (0.0 - 1.0)
            fraction = y_pos / total_height
            tab.canvas.yview_moveto(fraction)

    def on_click(self, event, tab):
        # In tekst-annotatiemodus: geen drag-selectie starten
        if hasattr(tab, 'text_annotate_mode') and tab.text_annotate_mode:
            return

        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)

        # Check eerst of er op een link is geklikt
        link_data = self.is_over_link(x, y, tab)
        if link_data:
            self.open_link(link_data, tab)
            return  # Stop hier, geen tekst selectie
        
        # Clear oude selectie
        tab.selected_text = ""
        
        tab.drag_start = (x, y)
        
        # Herstel alle originele pagina afbeeldingen (verwijder highlighting)
        if hasattr(tab, 'page_pil_images') and hasattr(tab, 'page_positions'):
            Image, ImageTk, ImageOps, ImageDraw = get_PIL()  # Lazy load PIL
            for page_num in tab.page_pil_images.keys():
                original_image = tab.page_pil_images[page_num]
                photo = ImageTk.PhotoImage(original_image)
                
                # Bewaar photo reference
                if len(tab.page_images) > page_num:
                    tab.page_images[page_num] = photo
                
                # Update de afbeelding op canvas
                page_y_offset = tab.page_positions[page_num]
                page_x_offset = tab.page_offset_x
                
                tab.canvas.delete(f"page_{page_num}")
                tab.canvas.create_image(page_x_offset, page_y_offset, 
                                       anchor="nw", image=photo, tags=f"page_{page_num}")
        
        # Breng annotaties en form highlights terug naar de voorgrond
        tab.canvas.tag_raise("form_highlight")
        tab.canvas.tag_raise("form_values")
        tab.canvas.tag_raise("text_annotation")
        tab.canvas.tag_raise("form_overlay")
        tab.canvas.tag_raise("text_annotation_editor")

        # Teken drag rectangle
        if tab.drag_rect:
            tab.canvas.delete(tab.drag_rect)
        tab.drag_rect = tab.canvas.create_rectangle(
            x, y, x, y, outline="red", width=2, tags="drag_rect"
        )

        self.status_label.config(text="Selecteren...")

    def on_drag(self, event, tab):
        if not tab.drag_start:
            return
        
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        
        if tab.drag_rect:
            tab.canvas.coords(tab.drag_rect, tab.drag_start[0], tab.drag_start[1], x, y)

    def on_release(self, event, tab):
        if not tab.drag_start:
            return
        
        Image, ImageTk, ImageOps, ImageDraw = get_PIL()  # Lazy load PIL
        
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        
        # Verwijder drag rectangle
        if tab.drag_rect:
            tab.canvas.delete(tab.drag_rect)
            tab.drag_rect = None
        
        # Selection bounds
        x1, y1 = tab.drag_start
        x2, y2 = x, y
        
        left = min(x1, x2)
        right = max(x1, x2)
        top = min(y1, y2)
        bottom = max(y1, y2)
        
        # Detecteer op welke pagina(s) de selectie zich bevindt
        selected_pages = set()
        page_regions = getattr(tab, 'page_regions', {})
        for word_data in tab.text_words:
            text, wx0, wy0, wx1, wy1 = word_data
            if not (wx1 < left or wx0 > right or wy1 < top or wy0 > bottom):
                if page_regions:
                    # Gebruik nauwkeurige bounding-box detectie (werkt ook in boek-modus)
                    for page_num, (rx0, ry0, rx1, ry1) in page_regions.items():
                        if rx0 <= wx0 <= rx1 and ry0 <= wy0 <= ry1:
                            selected_pages.add(page_num)
                            break
                else:
                    # Fallback: Y-positie gebaseerde detectie
                    for page_num, page_y_pos in enumerate(tab.page_positions):
                        if page_num + 1 < len(tab.page_positions):
                            next_page_y = tab.page_positions[page_num + 1]
                            if page_y_pos <= wy0 < next_page_y:
                                selected_pages.add(page_num)
                                break
                        else:
                            if wy0 >= page_y_pos:
                                selected_pages.add(page_num)
                                break
        
        # Vind geselecteerde woorden
        selected_words = []
        for word_data in tab.text_words:
            text, wx0, wy0, wx1, wy1 = word_data
            if not (wx1 < left or wx0 > right or wy1 < top or wy0 > bottom):
                selected_words.append(word_data)
        
        if len(selected_words) > 0:
            # Sorteer woorden
            selected_words.sort(key=lambda w: (w[2], w[1]))

            # Markeermodus: direct gele annotatie toevoegen, geen blauwe selectie
            if getattr(self, 'highlight_mode', False):
                # Verzamel tekst
                tab.selected_text = ""
                last_y = None
                for word_data in selected_words:
                    text_w, wx0, wy0, wx1, wy1 = word_data
                    if last_y is not None and abs(wy0 - last_y) > 5:
                        tab.selected_text += "\n"
                    elif tab.selected_text:
                        tab.selected_text += " "
                    tab.selected_text += text_w
                    last_y = wy0
                self.apply_highlight_annotation(tab, selected_words)
                self._update_save_button_state()
                tab.drag_start = None
                return

            # Normale modus: teken blauwe selectie-highlight
            _page_regions = getattr(tab, 'page_regions', {})
            for page_num in selected_pages:
                if not hasattr(tab, 'page_pil_images') or page_num not in tab.page_pil_images:
                    continue

                # Haal canvas-positie van deze pagina op
                if page_num in _page_regions:
                    page_x_offset, page_y_offset = _page_regions[page_num][0], _page_regions[page_num][1]
                else:
                    page_y_offset = tab.page_positions[page_num]
                    page_x_offset = tab.page_offset_x

                highlighted = tab.page_pil_images[page_num].copy()
                draw = ImageDraw.Draw(highlighted, 'RGBA')

                for word_data in selected_words:
                    text_w, wx0, wy0, wx1, wy1 = word_data

                    # Check of dit woord op deze pagina valt
                    if page_num in _page_regions:
                        rx0, ry0, rx1, ry1 = _page_regions[page_num]
                        if not (rx0 <= wx0 <= rx1 and ry0 <= wy0 <= ry1):
                            continue
                    else:
                        if page_num + 1 < len(tab.page_positions):
                            if not (page_y_offset <= wy0 < tab.page_positions[page_num + 1]):
                                continue
                        else:
                            if wy0 < page_y_offset:
                                continue

                    rel_x0 = wx0 - page_x_offset
                    rel_y0 = wy0 - page_y_offset
                    rel_x1 = wx1 - page_x_offset
                    rel_y1 = wy1 - page_y_offset

                    draw.rectangle(
                        [rel_x0, rel_y0, rel_x1, rel_y1],
                        fill=(173, 216, 230, 100),
                        outline=(100, 150, 200, 200),
                        width=1
                    )

                photo = ImageTk.PhotoImage(highlighted)
                if len(tab.page_images) > page_num:
                    tab.page_images[page_num] = photo
                tab.canvas.delete(f"page_{page_num}")
                tab.canvas.create_image(page_x_offset, page_y_offset,
                                        anchor="nw", image=photo, tags=f"page_{page_num}")

            # Verzamel tekst
            tab.selected_text = ""
            last_y = None
            for word_data in selected_words:
                text_w, wx0, wy0, wx1, wy1 = word_data

                if last_y is not None and abs(wy0 - last_y) > 5:
                    tab.selected_text += "\n"
                elif tab.selected_text:
                    tab.selected_text += " "

                tab.selected_text += text_w
                last_y = wy0

            self.status_label.config(text=f"Geselecteerd: {len(tab.selected_text)} tekens")
        else:
            self.status_label.config(text="Geen tekst geselecteerd")

        # Breng annotaties en form highlights terug naar de voorgrond
        tab.canvas.tag_raise("form_highlight")
        tab.canvas.tag_raise("form_values")
        tab.canvas.tag_raise("text_annotation")
        tab.canvas.tag_raise("form_overlay")
        tab.canvas.tag_raise("text_annotation_editor")

        tab.drag_start = None

    def is_over_link(self, x, y, tab):
        """Check of de cursor over een link is"""
        if not hasattr(tab, 'links'):
            return None
            
        for link_data in tab.links:
            lx0, ly0, lx1, ly1 = link_data['rect']
            if lx0 <= x <= lx1 and ly0 <= y <= ly1:
                return link_data
        return None
    
    def on_mouse_move(self, event, tab):
        """Handle mouse beweging voor link hover effecten"""
        # Check of tab links heeft
        if not hasattr(tab, 'links'):
            return
            
        x = tab.canvas.canvasx(event.x)
        y = tab.canvas.canvasy(event.y)
        
        link_data = self.is_over_link(x, y, tab)
        
        if link_data:
            # Cursor veranderen naar hand
            if not hasattr(tab, 'current_cursor'):
                tab.current_cursor = "arrow"
            
            if tab.current_cursor != "hand2":
                tab.canvas.config(cursor="hand2")
                tab.current_cursor = "hand2"
                
                # Toon link URL in statusbalk
                if link_data['uri']:
                    self.status_label.config(text=f"Link: {link_data['uri']}")
                elif link_data['page'] is not None:
                    self.status_label.config(text=f"Interne link naar pagina {link_data['page'] + 1}")
        else:
            # Terug naar normale cursor
            if not hasattr(tab, 'current_cursor'):
                tab.current_cursor = "arrow"
                
            if tab.current_cursor != "arrow":
                tab.canvas.config(cursor="arrow")
                tab.current_cursor = "arrow"
                if hasattr(self, 'status_label'):
                    self.status_label.config(text="Gereed")
    
    def show_link_warning(self, url):
        """Toon veiligheidswaarschuwing voordat een link wordt geopend"""
        # Maak een custom dialog
        dialog = tk.Toplevel(self.root)
        dialog.title("Link Openen")
        dialog.geometry("550x220")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Centreer op parent window
        dialog.update_idletasks()
        x = self.root.winfo_x() + (self.root.winfo_width() // 2) - (dialog.winfo_width() // 2)
        y = self.root.winfo_y() + (self.root.winfo_height() // 2) - (dialog.winfo_height() // 2)
        dialog.geometry(f"+{x}+{y}")
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(expand=True, fill="both", padx=30, pady=25)
        
        # URL label (in een frame met border)
        url_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], relief="solid", bd=1)
        url_frame.pack(fill="x", pady=(0, 20))
        
        url_label = tk.Label(
            url_frame, 
            text=url,
            font=("Segoe UI", 9),
            bg=self.theme["BG_SECONDARY"],
            fg=self.theme["TEXT_PRIMARY"],
            wraplength=480,
            padx=10,
            pady=8
        )
        url_label.pack()
        
        # Waarschuwing tekst
        warning_label = tk.Label(
            content_frame,
            text="Weet je zeker dat je deze link wilt openen?",
            font=("Segoe UI", 11, "bold"),
            bg=self.theme["BG_PRIMARY"],
            fg=self.theme["TEXT_PRIMARY"]
        )
        warning_label.pack(pady=(0, 25))
        
        # Result variabele
        result = [False]
        
        def on_yes():
            result[0] = True
            dialog.destroy()
        
        def on_no():
            result[0] = False
            dialog.destroy()
        
        # Button frame
        button_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        button_frame.pack()
        
        # Ja button (groen - altijd duidelijk zichtbaar)
        yes_button = tk.Button(
            button_frame,
            text="Ja",
            command=on_yes,
            font=("Segoe UI", 11, "bold"),
            bg="#28a745",
            fg="white",
            activebackground="#218838",
            activeforeground="white",
            relief="flat",
            width=10,
            height=1,
            cursor="hand2",
            borderwidth=0,
            highlightthickness=0
        )
        yes_button.pack(side="left", padx=10)
        
        # Nee button (rood - altijd duidelijk zichtbaar)
        no_button = tk.Button(
            button_frame,
            text="Nee",
            command=on_no,
            font=("Segoe UI", 11, "bold"),
            bg="#dc3545",
            fg="white",
            activebackground="#c82333",
            activeforeground="white",
            relief="flat",
            width=10,
            height=1,
            cursor="hand2",
            borderwidth=0,
            highlightthickness=0
        )
        no_button.pack(side="left", padx=10)
        
        # Enter key voor Ja, Escape voor Nee
        dialog.bind("<Return>", lambda e: on_yes())
        dialog.bind("<Escape>", lambda e: on_no())
        
        # Focus op Nee button (veiliger default)
        no_button.focus_set()
        
        # Wacht tot dialoog wordt gesloten
        dialog.wait_window()
        
        return result[0]
    
    def open_link(self, link_data, tab):
        """Open een link na bevestiging"""
        # Externe link (URI)
        if link_data['uri']:
            url = link_data['uri']
            # Toon waarschuwing
            if self.show_link_warning(url):
                try:
                    webbrowser.open(url)
                    self.status_label.config(text=f"Link geopend: {url}")
                except Exception as e:
                    messagebox.showerror("Fout", f"Kon link niet openen:\n{str(e)}")
            else:
                self.status_label.config(text="Link niet geopend")
        
        # Interne link (naar andere pagina)
        elif link_data['page'] is not None:
            target_page = link_data['page']
            if 0 <= target_page < len(tab.pdf_document):
                tab.current_page = target_page
                tab.scroll_to_page = target_page
                self.display_page(tab)
                self.status_label.config(text=f"Naar pagina {target_page + 1}")

    def copy_text(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab) and tab.selected_text:
            self.root.clipboard_clear()
            self.root.clipboard_append(tab.selected_text)
            self.status_label.config(text=f"Gekopieerd: {len(tab.selected_text)} tekens")
        else:
            self.status_label.config(text="Geen tekst geselecteerd")

    def show_search_dialog(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            search_window = tk.Toplevel(self.root)
            search_window.title("Zoeken in PDF")
            search_window.geometry("450x240")
            search_window.configure(bg=self.theme["BG_PRIMARY"])
            search_window.transient(self.root)
            search_window.grab_set()
            search_window.resizable(False, False)
            
            # Icon toevoegen
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    search_window.iconbitmap(icon_path)
            except:
                pass
            
            # Header met accent kleur (moderne stijl)
            header_frame = tk.Frame(search_window, bg=self.theme["ACCENT_COLOR"], height=60)
            header_frame.pack(fill=tk.X)
            header_frame.pack_propagate(False)
            
            tk.Label(header_frame, text="🔍 Zoeken in PDF", font=("Segoe UI", 14, "bold"),
                    bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
            
            # Content frame
            content_frame = tk.Frame(search_window, bg=self.theme["BG_PRIMARY"])
            content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
            
            tk.Label(content_frame, text="Zoek tekst:", font=Theme.FONT_MAIN,
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(pady=(0, 10))
            
            search_var = tk.StringVar()
            search_entry = tk.Entry(content_frame, textvariable=search_var, 
                                   font=Theme.FONT_MAIN, width=40)
            search_entry.pack(pady=5)
            search_entry.focus()
            
            # Footer met knoppen (moderne stijl)
            footer_frame = tk.Frame(search_window, bg=self.theme["BG_SECONDARY"], height=70)
            footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
            footer_frame.pack_propagate(False)
            
            btn_frame = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
            btn_frame.pack(expand=True)
            
            def do_search():
                search_text = search_var.get()
                if search_text:
                    self.search_in_pdf(tab, search_text)
                    search_window.destroy()
            
            tk.Button(btn_frame, text="Zoeken", command=do_search, 
                     bg=self.theme["ACCENT_COLOR"], fg="white", 
                     font=("Segoe UI", 10), padx=25, pady=10,
                     relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
            
            tk.Button(btn_frame, text="Annuleren", command=search_window.destroy,
                     bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                     font=("Segoe UI", 10), padx=25, pady=10,
                     relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
            
            search_entry.bind("<Return>", lambda e: do_search())

    def search_in_pdf(self, tab, search_text):
        fitz = get_fitz()  # Lazy load voor fitz.Rect
        Image, ImageTk, ImageOps, ImageDraw = get_PIL()  # Lazy load PIL
        found = False
        start_page = tab.current_page

        for offset in range(len(tab.pdf_document)):
            page_num = (start_page + offset) % len(tab.pdf_document)
            page = tab.pdf_document[page_num]
            instances = page.search_for(search_text)

            if instances:
                if page_num != tab.current_page:
                    tab.current_page = page_num
                    self.display_page(tab)

                # Gebruik page_pil_images (multi-page) of current_image als fallback
                pil_images = getattr(tab, 'page_pil_images', {})
                source_image = pil_images.get(page_num) or getattr(tab, 'current_image', None)
                if source_image is None:
                    # Render pagina on-the-fly als er geen afbeelding beschikbaar is
                    mat = fitz.Matrix(tab.zoom_level, tab.zoom_level)
                    pix = page.get_pixmap(matrix=mat)
                    source_image = Image.open(io.BytesIO(pix.tobytes("ppm")))

                # Highlight op de afbeelding
                highlighted = source_image.copy()
                draw = ImageDraw.Draw(highlighted, 'RGBA')

                for inst in instances:
                    rect = fitz.Rect(inst)
                    x0 = rect.x0 * tab.zoom_level
                    y0 = rect.y0 * tab.zoom_level
                    x1 = rect.x1 * tab.zoom_level
                    y1 = rect.y1 * tab.zoom_level

                    draw.rectangle(
                        [x0, y0, x1, y1],
                        outline=(255, 140, 0, 255),  # Oranje
                        width=3
                    )

                photo = ImageTk.PhotoImage(highlighted)
                tab.highlighted_image = photo

                # Gebruik page_regions voor correcte positie, of fallback
                regions = getattr(tab, 'page_regions', {})
                if page_num in regions:
                    px, py = regions[page_num][0], regions[page_num][1]
                else:
                    px = getattr(tab, 'page_offset_x', 0)
                    py = getattr(tab, 'page_offset_y', 0)

                tab.canvas.delete(f"page_{page_num}")
                tab.canvas.create_image(px, py, anchor="nw",
                                        image=photo, tags=f"page_{page_num}")

                # Breng overlays terug naar voorgrond
                tab.canvas.tag_raise("form_highlight")
                tab.canvas.tag_raise("form_values")
                tab.canvas.tag_raise("text_annotation")
                tab.canvas.tag_raise("form_overlay")

                found = True
                self.status_label.config(
                    text=f"Gevonden: '{search_text}' ({len(instances)}x op pagina {page_num + 1})"
                )
                break

        if not found:
            messagebox.showinfo("Zoeken", f"'{search_text}' niet gevonden in document")

    # Navigatie functies
    def navigate(self, delta):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            new_page = tab.current_page + delta
            if 0 <= new_page < len(tab.pdf_document):
                tab.current_page = new_page
                self.scroll_to_page(tab, tab.current_page)
                self.update_ui_state()

    def first_page(self): 
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.current_page = 0
            self.scroll_to_page(tab, 0)
            self.update_ui_state()

    def prev_page(self): 
        self.navigate(-1)
        
    def next_page(self): 
        self.navigate(1)
    
    def last_page(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.current_page = len(tab.pdf_document) - 1
            self.scroll_to_page(tab, tab.current_page)
            self.update_ui_state()

    def go_to_page(self, event=None):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            try:
                page_num = int(self.page_var.get()) - 1
                if 0 <= page_num < len(tab.pdf_document):
                    tab.current_page = page_num
                    self.scroll_to_page(tab, page_num)
                    self.update_ui_state()
            except ValueError:
                self.update_ui_state()

    # Zoom functies
    def zoom(self, factor):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.zoom_mode = "manual"
            new_zoom = tab.zoom_level * factor
            if 0.2 < new_zoom < 5.0:
                tab.zoom_level = new_zoom
                self.display_page(tab)
    
    def zoom_in(self): 
        self.zoom(1.2)
        
    def zoom_out(self): 
        self.zoom(1/1.2)

    def set_zoom_mode(self, mode):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            tab.zoom_mode = mode
            self.display_page(tab)

    def print_pdf(self):
        """Toon ingebouwde print dialoog met printer selectie"""
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            print_dialog = tk.Toplevel(self.root)
            print_dialog.title("Afdrukken")
            print_dialog.geometry("650x730")
            print_dialog.configure(bg=self.theme["BG_PRIMARY"])
            print_dialog.transient(self.root)
            print_dialog.grab_set()
            print_dialog.resizable(False, False)
            
            # Voeg logo toe aan taakbalk
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    print_dialog.iconbitmap(icon_path)
            except:
                pass
            
            # Header
            header_frame = tk.Frame(print_dialog, bg=self.theme["ACCENT_COLOR"], height=60)
            header_frame.pack(fill=tk.X)
            header_frame.pack_propagate(False)
            
            tk.Label(header_frame, text="🖨️ PDF Afdrukken", font=("Segoe UI", 14, "bold"),
                    bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
            
            # Main content with scrollbar
            canvas = tk.Canvas(print_dialog, bg=self.theme["BG_PRIMARY"], highlightthickness=0)
            scrollbar = ttk.Scrollbar(print_dialog, orient="vertical", command=canvas.yview)
            content_frame = tk.Frame(canvas, bg=self.theme["BG_PRIMARY"])

            content_frame.bind(
                "<Configure>",
                lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
            )

            canvas.create_window((0, 0), window=content_frame, anchor="nw")
            canvas.configure(yscrollcommand=scrollbar.set)

            # Knoppen ALTIJD ZICHTBAAR ONDERAAN (voor canvas packing!)
            btn_frame = tk.Frame(print_dialog, bg=self.theme["BG_SECONDARY"], height=70)
            btn_frame.pack(fill=tk.X, side=tk.BOTTOM)
            btn_frame.pack_propagate(False)

            button_container = tk.Frame(btn_frame, bg=self.theme["BG_SECONDARY"])
            button_container.pack(expand=True)

            # Canvas met scrollbar bovenaan
            canvas.pack(side="left", fill="both", expand=True, padx=20, pady=20)
            scrollbar.pack(side="right", fill="y", padx=(0, 20), pady=20)

            # Mousewheel scrolling
            def _on_mousewheel(event):
                canvas.yview_scroll(int(-1*(event.delta/120)), "units")
            canvas.bind_all("<MouseWheel>", _on_mousewheel)
            
            # Document info
            tk.Label(content_frame, text="Document:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 2))
            tk.Label(content_frame, text=os.path.basename(tab.file_path), 
                    font=("Segoe UI", 9),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], anchor="w").pack(fill=tk.X, pady=(0, 15))
            
            # Printer selectie
            tk.Label(content_frame, text="Printer:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            printer_var = tk.StringVar(value="Printers laden...")

            # Frame voor combobox met border voor betere zichtbaarheid
            combo_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"],
                                  highlightbackground=self.theme["TEXT_SECONDARY"],
                                  highlightthickness=1)
            combo_frame.pack(fill=tk.X, pady=(0, 15))

            # Combobox - tijdelijk uitgeschakeld tot printers geladen zijn
            printer_dropdown = ttk.Combobox(combo_frame, textvariable=printer_var,
                                           values=[], state="disabled",
                                           font=("Segoe UI", 10), width=40, height=10)
            printer_dropdown.pack(fill=tk.X, padx=2, pady=2)

            def load_printers_async():
                printers = self.get_available_printers()
                default_printer = "Standaard printer"
                try:
                    import win32print
                    default_printer = win32print.GetDefaultPrinter()
                except:
                    pass
                default_value = default_printer if default_printer in printers else (printers[0] if printers else "Standaard printer")

                def update_ui():
                    try:
                        printer_dropdown.config(values=printers, state="readonly")
                        printer_var.set(default_value)
                    except:
                        pass  # Dialoog al gesloten

                if print_dialog.winfo_exists():
                    print_dialog.after(0, update_ui)

            threading.Thread(target=load_printers_async, daemon=True).start()
            
            # Override combobox kleuren voor beter contrast
            self.root.option_add('*TCombobox*Listbox.background', 'white')
            self.root.option_add('*TCombobox*Listbox.foreground', 'black')
            self.root.option_add('*TCombobox*Listbox.selectBackground', self.theme["ACCENT_COLOR"])
            self.root.option_add('*TCombobox*Listbox.selectForeground', 'white')
            
            # Pagina selectie
            tk.Label(content_frame, text="Pagina's:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            page_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            page_frame.pack(fill=tk.X, pady=(0, 15))
            
            page_option = tk.StringVar(value="all")
            
            # Alle pagina's
            tk.Radiobutton(page_frame, text=f"Alle pagina's (1-{len(tab.pdf_document)})", 
                          variable=page_option, value="all",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w")
            
            # Huidige pagina
            current_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            current_frame.pack(anchor="w", pady=5)
            tk.Radiobutton(current_frame, text="Huidige pagina", 
                          variable=page_option, value="current",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)
            tk.Label(current_frame, text=f"(pagina {tab.current_page + 1})",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], 
                    font=("Segoe UI", 8)).pack(side=tk.LEFT, padx=5)
            
            # Aangepaste pagina's
            custom_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            custom_frame.pack(anchor="w", pady=5)
            
            tk.Radiobutton(custom_frame, text="Aangepaste pagina's:", 
                          variable=page_option, value="custom",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)
            
            custom_pages_var = tk.StringVar(value="1,3")
            
            # Frame voor entry met border
            entry_frame = tk.Frame(custom_frame, bg=self.theme["BG_SECONDARY"],
                                  highlightbackground=self.theme["TEXT_SECONDARY"],
                                  highlightthickness=1)
            entry_frame.pack(side=tk.LEFT, padx=5)
            
            custom_entry = tk.Entry(entry_frame, textvariable=custom_pages_var,
                                   font=("Segoe UI", 10), width=18,
                                   bg="white",
                                   fg="black",
                                   relief="flat",
                                   insertbackground="black")
            custom_entry.pack(padx=2, pady=2)
            
            # Uitleg voor aangepaste pagina's
            help_frame = tk.Frame(page_frame, bg=self.theme["BG_PRIMARY"])
            help_frame.pack(anchor="w", padx=20, pady=2)
            tk.Label(help_frame, text="(bijv: 1,3,5 of 1-3,5 of 2-5)",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"], 
                    font=("Segoe UI", 8)).pack(anchor="w")
            
            # Aantal kopieën
            tk.Label(content_frame, text="Aantal kopieën:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            copies_var = tk.StringVar(value="1")
            copies_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            copies_frame.pack(fill=tk.X, pady=(0, 15))
            
            # Frame voor spinbox met border
            spinbox_frame = tk.Frame(copies_frame, bg=self.theme["BG_SECONDARY"],
                                    highlightbackground=self.theme["TEXT_SECONDARY"],
                                    highlightthickness=1)
            spinbox_frame.pack(side=tk.LEFT)
            
            tk.Spinbox(spinbox_frame, from_=1, to=99, textvariable=copies_var,
                      font=("Segoe UI", 10), width=8,
                      bg="white",
                      fg="black",
                      buttonbackground=self.theme["BG_SECONDARY"],
                      relief="flat",
                      insertbackground="black").pack(padx=2, pady=2)
            
            # Passend maken optie
            tk.Label(content_frame, text="Afdruk opties:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))
            
            fit_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            fit_frame.pack(fill=tk.X, pady=(0, 10))
            
            fit_to_page_var = tk.BooleanVar(value=True)
            tk.Checkbutton(fit_frame, text="Passend maken op pagina",
                          variable=fit_to_page_var,
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w")
            tk.Label(fit_frame, text="(schaalt document om op papier te passen)",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"],
                    font=("Segoe UI", 8)).pack(anchor="w", padx=20)

            # Dubbelzijdig printen
            duplex_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            duplex_frame.pack(fill=tk.X, pady=(0, 10))

            duplex_var = tk.BooleanVar(value=False)
            tk.Checkbutton(duplex_frame, text="Dubbelzijdig printen (beide zijden)",
                          variable=duplex_var,
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w")
            tk.Label(duplex_frame, text="(drukt op beide zijden van het papier - papier besparing)",
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"],
                    font=("Segoe UI", 8)).pack(anchor="w", padx=20)

            # Kleur modus
            tk.Label(content_frame, text="Kleur:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))

            color_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            color_frame.pack(fill=tk.X, pady=(0, 10))

            color_mode_var = tk.StringVar(value="kleur")
            tk.Radiobutton(color_frame, text="Kleur", variable=color_mode_var, value="kleur",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 20))
            tk.Radiobutton(color_frame, text="Zwart/Wit", variable=color_mode_var, value="zwart_wit",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)

            # Rotatie
            tk.Label(content_frame, text="Rotatie:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))

            rot_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            rot_frame.pack(fill=tk.X, pady=(0, 10))

            rotation_var = tk.StringVar(value="0")
            for rot_text, rot_val in [("Geen", "0"), ("90° rechts", "90"), ("180°", "180"), ("90° links", "270")]:
                tk.Radiobutton(rot_frame, text=rot_text, variable=rotation_var, value=rot_val,
                              bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                              selectcolor=self.theme["BG_SECONDARY"],
                              activebackground=self.theme["BG_PRIMARY"],
                              activeforeground=self.theme["TEXT_PRIMARY"],
                              font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 12))

            # Pagina oriëntatie
            tk.Label(content_frame, text="Oriëntatie:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))

            orient_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            orient_frame.pack(fill=tk.X, pady=(0, 10))

            orientation_var = tk.StringVar(value="staand")
            tk.Radiobutton(orient_frame, text="Staand", variable=orientation_var, value="staand",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 20))
            tk.Radiobutton(orient_frame, text="Liggend", variable=orientation_var, value="liggend",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)

            # Snelheid / Kwaliteit
            tk.Label(content_frame, text="Kwaliteit:", font=("Segoe UI", 9, "bold"),
                    bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"], anchor="w").pack(fill=tk.X, pady=(0, 5))

            quality_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
            quality_frame.pack(fill=tk.X, pady=(0, 10))

            quality_var = tk.StringVar(value="normaal")
            tk.Radiobutton(quality_frame, text="Snel (besparing)", variable=quality_var, value="snel",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 15))
            tk.Radiobutton(quality_frame, text="Normaal", variable=quality_var, value="normaal",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT, padx=(0, 15))
            tk.Radiobutton(quality_frame, text="Hoog (beste kwaliteit)", variable=quality_var, value="hoog",
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(side=tk.LEFT)
            def open_windows_print():
                try:
                    import win32print
                    printer_name = printer_var.get()
                    if not printer_name or printer_name in ("Printers laden...", "Standaard printer"):
                        try:
                            printer_name = win32print.GetDefaultPrinter()
                        except:
                            messagebox.showerror("Fout", "Geen standaard printer gevonden.")
                            return

                    hPrinter = win32print.OpenPrinter(printer_name)
                    try:
                        DM_IN_PROMPT = 4
                        win32print.DocumentProperties(0, hPrinter, printer_name, None, None, DM_IN_PROMPT)
                    finally:
                        win32print.ClosePrinter(hPrinter)

                except Exception as e:
                    messagebox.showerror("Fout", f"Kan printer instellingen niet openen:\n{str(e)}")

            def do_print():
                try:
                    printer = printer_var.get()
                    copies = int(copies_var.get())
                    page_opt = page_option.get()
                    fit_to_page = fit_to_page_var.get()
                    color_mode = color_mode_var.get()
                    rotation = int(rotation_var.get())
                    duplex = duplex_var.get()
                    orientation = orientation_var.get()
                    quality = quality_var.get()

                    # Waarschuwing als duplex is geselecteerd
                    if duplex:
                        if not messagebox.askyesno("Dubbelzijdig Printen",
                            "Dubbelzijdig printen is geselecteerd.\n\n"
                            "Let op: Niet alle printers ondersteunen automatisch dubbelzijdig printen.\n\n"
                            "Als uw printer dit niet ondersteunt, ziet u een dialoog om\n"
                            "het papier handmatig om te draaien.\n\n"
                            "Wilt u doorgaan?"):
                            return

                    # Bepaal welke pagina's te printen
                    if page_opt == "current":
                        pages_to_print = [tab.current_page]
                    elif page_opt == "custom":
                        # Parse aangepaste pagina selectie
                        custom_pages = custom_pages_var.get()
                        pages_to_print = self.parse_page_range(custom_pages, len(tab.pdf_document))

                        if not pages_to_print:
                            messagebox.showerror("Ongeldige pagina's",
                                "Ongeldige pagina selectie.\n\n" +
                                "Gebruik formaat zoals:\n" +
                                "• 1,3,5 (specifieke pagina's)\n" +
                                "• 1-5 (bereik)\n" +
                                "• 1-3,5,7-9 (combinatie)")
                            return
                    else:  # all
                        pages_to_print = list(range(len(tab.pdf_document)))

                    print_dialog.destroy()
                    self.execute_print(tab, printer, pages_to_print, copies, fit_to_page, color_mode, rotation, duplex, orientation, quality)

                except ValueError as e:
                    messagebox.showerror("Invoer Fout", f"Ongeldig aantal kopieën: {str(e)}")
                except Exception as e:
                    print_dialog.destroy()
                    messagebox.showerror("Print Fout", f"Kan niet printen:\n{str(e)}")

            print_btn = tk.Button(button_container, text="Afdrukken",
                                 command=do_print,
                                 bg=self.theme["ACCENT_COLOR"], fg="white",
                                 font=("Segoe UI", 10, "bold"),
                                 padx=30, pady=10,
                                 relief="flat", cursor="hand2")
            print_btn.pack(side=tk.LEFT, padx=5)

            cancel_btn = tk.Button(button_container, text="Annuleren",
                                  command=print_dialog.destroy,
                                  bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                                  font=("Segoe UI", 10),
                                  padx=20, pady=10,
                                  relief="flat", cursor="hand2")
            cancel_btn.pack(side=tk.LEFT, padx=5)

            windows_btn = tk.Button(button_container, text="Printer instellingen...",
                                   command=open_windows_print,
                                   bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                                   font=("Segoe UI", 9),
                                   padx=15, pady=10,
                                   relief="flat", cursor="hand2")
            windows_btn.pack(side=tk.LEFT, padx=5)

            def on_enter_print(e):
                print_btn.config(bg="#0d8cbd")
            def on_leave_print(e):
                print_btn.config(bg=self.theme["ACCENT_COLOR"])
            def on_enter_cancel(e):
                cancel_btn.config(bg=self.theme["ACCENT_COLOR"], fg="white")
            def on_leave_cancel(e):
                cancel_btn.config(bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"])
            def on_enter_windows(e):
                windows_btn.config(bg=self.theme["ACCENT_COLOR"], fg="white")
            def on_leave_windows(e):
                windows_btn.config(bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"])

            print_btn.bind("<Enter>", on_enter_print)
            print_btn.bind("<Leave>", on_leave_print)
            cancel_btn.bind("<Enter>", on_enter_cancel)
            cancel_btn.bind("<Leave>", on_leave_cancel)
            windows_btn.bind("<Enter>", on_enter_windows)
            windows_btn.bind("<Leave>", on_leave_windows)

    def get_available_printers(self):
        """Haal beschikbare printers op"""
        printers = []
        try:
            if platform.system() == "Windows":
                try:
                    import win32print
                    printers = [printer[2] for printer in win32print.EnumPrinters(2)]
                except ImportError:
                    result = subprocess.run(
                        ["powershell", "-Command", "Get-Printer | Select-Object -ExpandProperty Name"],
                        capture_output=True, text=True, timeout=5, creationflags=subprocess.CREATE_NO_WINDOW
                    )
                    if result.returncode == 0:
                        printers = [p.strip() for p in result.stdout.split('\n') if p.strip()]
        except Exception as e:
            print(f"Kon printers niet ophalen: {e}")
        
        if not printers:
            printers = ["Standaard printer"]
        
        return printers

    def parse_page_range(self, page_string, total_pages):
        """Parse pagina bereik string zoals '1,3,5' of '1-5,7' naar lijst van pagina nummers (0-indexed)"""
        pages = set()
        
        try:
            # Verwijder spaties
            page_string = page_string.replace(" ", "")
            
            # Split op komma's
            parts = page_string.split(',')
            
            for part in parts:
                if '-' in part:
                    # Bereik zoals '1-5'
                    start, end = part.split('-')
                    start = int(start)
                    end = int(end)
                    
                    # Valideer bereik
                    if start < 1 or end > total_pages or start > end:
                        return None
                    
                    # Voeg alle pagina's in bereik toe (convert naar 0-indexed)
                    for page_num in range(start, end + 1):
                        pages.add(page_num - 1)
                else:
                    # Enkele pagina zoals '3'
                    page_num = int(part)
                    
                    # Valideer pagina nummer
                    if page_num < 1 or page_num > total_pages:
                        return None
                    
                    # Voeg pagina toe (convert naar 0-indexed)
                    pages.add(page_num - 1)
            
            # Sorteer en return als lijst
            return sorted(list(pages))
            
        except (ValueError, AttributeError):
            return None

    def execute_print(self, tab, printer, pages, copies, fit_to_page=True, color_mode="kleur", rotation=0, duplex=False, orientation="portret", quality="normaal"):
        """Print DIRECT naar printer via Windows GDI met goede error handling"""
        try:
            fitz = get_fitz()  # Lazy load voor fitz.Matrix
            import win32print
            import win32ui
            import win32con
            from PIL import Image, ImageWin
            
        except ImportError:
            messagebox.showerror("Module Ontbreekt",
                "De 'pywin32' module is vereist voor printen.\n\n"
                "Installeer met: pip install pywin32\n\n"
                "Start daarna NVict Reader opnieuw op.")
            return
        
        try:
            # Krijg printer naam
            if printer == "Standaard printer" or not printer:
                printer = win32print.GetDefaultPrinter()
            
            # Verificeer dat printer bestaat en beschikbaar is
            try:
                printer_info = win32print.GetPrinter(win32print.OpenPrinter(printer))
            except:
                messagebox.showerror("Printer Niet Gevonden",
                    f"Kan printer '{printer}' niet vinden.\n\n"
                    f"Mogelijke oorzaken:\n"
                    f"• Printer is offline\n"
                    f"• Printer is niet geïnstalleerd\n"
                    f"• Geen toegang tot netwerkprinter\n\n"
                    f"Check de printer in Windows Instellingen.")
                return
            
            # Maak printer device context
            try:
                hDC = win32ui.CreateDC()
                hDC.CreatePrinterDC(printer)
            except Exception as e:
                messagebox.showerror("Kan Niet Verbinden met Printer",
                    f"Kan geen verbinding maken met printer '{printer}'.\n\n"
                    f"Mogelijke oorzaken:\n"
                    f"• Printer driver probleem\n"
                    f"• Printer is in gebruik\n"
                    f"• Onvoldoende rechten\n\n"
                    f"Probeer de printer opnieuw te installeren of\n"
                    f"gebruik een andere printer.")
                return
            
            # Krijg printer eigenschappen
            try:
                printer_width = hDC.GetDeviceCaps(win32con.HORZRES)
                printer_height = hDC.GetDeviceCaps(win32con.VERTRES)
                printer_ppi_x = hDC.GetDeviceCaps(win32con.LOGPIXELSX)
                printer_ppi_y = hDC.GetDeviceCaps(win32con.LOGPIXELSY)
            except Exception as e:
                hDC.DeleteDC()
                messagebox.showerror("Printer Eigenschappen Fout",
                    f"Kan printer eigenschappen niet ophalen.\n\n"
                    f"Dit kan betekenen dat de printer driver\n"
                    f"niet correct is geïnstalleerd.\n\n"
                    f"Herinstalleer de printer driver.")
                return
            
            # Start print job
            try:
                hDC.StartDoc("NVict Reader")
            except Exception as e:
                hDC.DeleteDC()
                error_msg = str(e).lower()
                
                if "access" in error_msg or "denied" in error_msg:
                    messagebox.showerror("Geen Toegang",
                        f"Geen toegang tot printer '{printer}'.\n\n"
                        f"Mogelijke oorzaken:\n"
                        f"• Onvoldoende gebruikersrechten\n"
                        f"• Printer is vergrendeld door admin\n"
                        f"• Printer is in gebruik door ander programma\n\n"
                        f"Neem contact op met uw systeembeheerder.")
                elif "offline" in error_msg or "not ready" in error_msg:
                    messagebox.showerror("Printer Offline",
                        f"Printer '{printer}' is offline of niet gereed.\n\n"
                        f"Controleer:\n"
                        f"• Is de printer aangezet?\n"
                        f"• Is de printer verbonden (USB/netwerk)?\n"
                        f"• Heeft de printer papier/toner?\n"
                        f"• Zijn er error lampjes op de printer?")
                else:
                    messagebox.showerror("Kan Print Job Niet Starten",
                        f"Kan print job niet starten.\n\n"
                        f"Printer: {printer}\n"
                        f"Fout: {str(e)}\n\n"
                        f"Probeer:\n"
                        f"• Check of printer werkt in andere programma's\n"
                        f"• Herstart de printer\n"
                        f"• Check Windows printer wachtrij")
                return
            
            print_success = False
            
            try:
                # Print alle kopieën
                for copy_num in range(copies):
                    # Print elke geselecteerde pagina
                    for page_num in pages:
                        try:
                            # Start nieuwe pagina
                            hDC.StartPage()
                            
                            # Haal PDF pagina op
                            page = tab.pdf_document[page_num]
                            
                            # Render pagina naar hoge resolutie image
                            dpi_scale = max(printer_ppi_x, printer_ppi_y) / 72
                            zoom = min(dpi_scale, 4.0)
                            mat = fitz.Matrix(zoom, zoom)
                            pix = page.get_pixmap(matrix=mat, alpha=False)
                            
                            # Converteer naar PIL Image
                            img_data = pix.tobytes("ppm")
                            img = Image.open(io.BytesIO(img_data))

                            # Rotatie toepassen
                            if rotation == 90:
                                img = img.rotate(-90, expand=True)
                            elif rotation == 180:
                                img = img.rotate(180, expand=True)
                            elif rotation == 270:
                                img = img.rotate(90, expand=True)

                            # Kleur modus toepassen
                            if color_mode == "zwart_wit":
                                img = img.convert('L').convert('RGB')

                            # Bereken afmetingen voor printer
                            img_width, img_height = img.size
                            aspect_ratio = img_width / img_height
                            
                            if fit_to_page:
                                printer_aspect = printer_width / printer_height
                                
                                if aspect_ratio > printer_aspect:
                                    print_width = printer_width
                                    print_height = int(printer_width / aspect_ratio)
                                else:
                                    print_height = printer_height
                                    print_width = int(printer_height * aspect_ratio)
                            else:
                                print_width = int(img_width * 72 / printer_ppi_x * printer_width / printer_width)
                                print_height = int(img_height * 72 / printer_ppi_y * printer_height / printer_height)
                                
                                if print_width > printer_width or print_height > printer_height:
                                    scale = min(printer_width / print_width, printer_height / print_height)
                                    print_width = int(print_width * scale)
                                    print_height = int(print_height * scale)
                            
                            # Centreer op pagina
                            x = (printer_width - print_width) // 2
                            y = (printer_height - print_height) // 2
                            
                            # Print image naar printer DC
                            dib = ImageWin.Dib(img)
                            dib.draw(hDC.GetHandleOutput(), (x, y, x + print_width, y + print_height))
                            
                            # Einde pagina
                            hDC.EndPage()
                            
                        except Exception as page_error:
                            print(f"Fout bij printen pagina {page_num + 1}: {page_error}")
                            # Probeer door te gaan met volgende pagina
                            try:
                                hDC.EndPage()
                            except:
                                pass
                
                # Einde document - als we hier komen is alles geslaagd
                hDC.EndDoc()
                print_success = True
                
            except Exception as print_error:
                # Print proces gefaald
                try:
                    hDC.AbortDoc()
                except:
                    pass
                
                messagebox.showerror("Print Fout",
                    f"Fout tijdens het printen:\n{str(print_error)}\n\n"
                    f"De print job is geannuleerd.\n\n"
                    f"Probeer opnieuw of gebruik een andere printer.")
                
            finally:
                # Sluit printer DC
                try:
                    hDC.DeleteDC()
                except:
                    pass
            
            # Alleen success message tonen als echt gelukt
            if print_success:
                # Maak leesbare pagina info
                if len(pages) == 1:
                    page_info = f"Pagina {pages[0] + 1}"
                elif len(pages) <= 5:
                    page_info = f"Pagina's {', '.join(str(p + 1) for p in pages)}"
                else:
                    page_info = f"{len(pages)} pagina's"
                
                self.status_label.config(text=f"Print job verzonden naar {printer}")
                
                # GEEN messagebox - alleen status update
                # Gebruiker ziet zelf of het werkt
                
        except Exception as e:
            import traceback
            traceback.print_exc()
            messagebox.showerror("Onverwachte Fout",
                f"Er is een onverwachte fout opgetreden:\n{str(e)}\n\n"
                f"Controleer of pywin32 correct is geïnstalleerd:\n"
                f"pip install pywin32")

    def cleanup_temp_file(self, filepath):
        """Verwijder tijdelijk bestand"""
        try:
            if os.path.exists(filepath):
                os.remove(filepath)
        except:
            pass

    def split_pdf(self):
        """Splits PDF in losse pagina's"""
        fitz = get_fitz()  # Lazy load
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Vraag output folder
        folder_path = filedialog.askdirectory(
            title="Selecteer map voor gesplitste pagina's"
        )
        
        if not folder_path:
            return
        
        try:
            base_name = os.path.splitext(os.path.basename(tab.file_path))[0]
            
            # Splits elke pagina
            for page_num in range(len(tab.pdf_document)):
                output_doc = fitz.open()
                output_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                
                output_path = os.path.join(folder_path, f"{base_name}_pagina_{page_num + 1}.pdf")
                output_doc.save(output_path)
                output_doc.close()
            
            messagebox.showinfo("Succes",
                f"PDF succesvol gesplitst!\n\n"
                f"Aantal pagina's: {len(tab.pdf_document)}\n"
                f"Opgeslagen in: {folder_path}")
            
            # Open de folder
            if messagebox.askyesno("Map Openen", "Wil je de map met bestanden openen?"):
                if platform.system() == "Windows":
                    os.startfile(folder_path)
                elif platform.system() == "Darwin":
                    subprocess.run(["open", folder_path])
                else:
                    subprocess.run(["xdg-open", folder_path])
            
        except Exception as e:
            messagebox.showerror("Fout", f"Kan PDF niet splitsen:\n{str(e)}")

    def _build_modified_pdf(self, tab):
        """Maak een tijdelijk PDF-bestand met alle wijzigingen (annotaties, markeringen, formuliervelden).
        Retourneert het pad naar het tijdelijke bestand, of None als er geen wijzigingen zijn."""
        # Verzamel waarden uit actieve form widgets
        if tab.form_widgets:
            self._save_form_widget_values(tab)

        has_changes = (
            bool(tab.form_field_values)
            or bool(tab.text_annotations)
            or bool(getattr(tab, 'highlight_annotations', []))
        )
        if not has_changes:
            return None

        try:
            fitz = get_fitz()
            import tempfile
            doc = fitz.open(tab.file_path)

            # 1. Formuliervelden invullen
            if tab.form_field_values:
                for page_num in range(len(doc)):
                    page = doc[page_num]
                    for widget in page.widgets():
                        if widget is None:
                            continue
                        xref = widget.xref
                        if xref in tab.form_field_values:
                            val = tab.form_field_values[xref]
                            if widget.field_type in (2, 5):  # Checkbox / Radio
                                widget.field_value = bool(val)
                            else:
                                widget.field_value = str(val) if val else ""
                            widget.update()

            # 2. Tekst-annotaties toevoegen
            color_map = {
                "black": (0, 0, 0), "red": (1, 0, 0),
                "#0066CC": (0, 0.4, 0.8), "blue": (0, 0, 1),
                "darkgreen": (0, 0.4, 0),
            }
            for annot in tab.text_annotations:
                page = doc[annot["page_num"]]
                pdf_x = annot["pdf_x"]
                pdf_y = annot["pdf_y"]
                text = annot["text"]
                font_size = annot["font_size"]
                color = color_map.get(annot.get("color", "black"), (0, 0, 0))

                lines = text.split("\n")
                max_line_len = max(len(line) for line in lines) if lines else 10
                approx_width = max_line_len * font_size * 0.55
                approx_height = len(lines) * font_size * 1.4

                rect = fitz.Rect(pdf_x, pdf_y,
                                 pdf_x + approx_width + 10,
                                 pdf_y + approx_height + 5)
                fontname = annot.get("fontname", "helv")
                annot_obj = page.add_freetext_annot(
                    rect, text,
                    fontsize=font_size, fontname=fontname,
                    text_color=color, fill_color=(1, 1, 1),
                )
                annot_obj.update()

            # 3. Highlight-markeringen toevoegen
            if hasattr(tab, 'highlight_annotations'):
                for ha in tab.highlight_annotations:
                    page = doc[ha["page_num"]]
                    quads = ha.get("quads", [])
                    if quads:
                        annot_obj = page.add_highlight_annot(quads=quads)
                        annot_obj.update()

            # Sla op in tijdelijk bestand
            base_name = os.path.splitext(os.path.basename(tab.file_path))[0]
            tmp = tempfile.NamedTemporaryFile(
                suffix=".pdf", prefix=f"{base_name}_bewerkt_", delete=False
            )
            tmp_path = tmp.name
            tmp.close()
            doc.save(tmp_path)
            doc.close()
            return tmp_path
        except Exception as e:
            print(f"Error building modified PDF: {e}")
            return None

    def send_pdf(self):
        """Verstuur/deel de actieve PDF via e-mail (Outlook of standaard mailprogramma)"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showinfo("Doorsturen", "Geen PDF geopend om door te sturen.")
            return

        # Bouw een gewijzigde PDF als er aanpassingen zijn, anders gebruik origineel
        modified_path = self._build_modified_pdf(tab)
        pdf_path = modified_path if modified_path else tab.file_path
        filename = os.path.basename(pdf_path)

        # Probeer eerst via Outlook COM object (Windows)
        try:
            import win32com.client
            outlook = win32com.client.Dispatch("Outlook.Application")
            mail = outlook.CreateItem(0)  # 0 = olMailItem
            mail.Subject = f"PDF: {os.path.basename(tab.file_path)}"
            mail.Body = f"Hierbij stuur ik de PDF '{os.path.basename(tab.file_path)}' door."
            mail.Attachments.Add(pdf_path)
            mail.Display(True)
            return
        except ImportError:
            pass
        except Exception:
            pass

        # Fallback: open standaard mailprogramma via mailto: link
        try:
            import urllib.parse
            subject = urllib.parse.quote(f"PDF: {os.path.basename(tab.file_path)}")
            body = urllib.parse.quote(f"Hierbij stuur ik de PDF '{os.path.basename(tab.file_path)}' door.")
            mailto = f"mailto:?subject={subject}&body={body}"
            import webbrowser
            webbrowser.open(mailto)

            # Kopieer bestandspad naar klembord als handige hint
            self.root.clipboard_clear()
            self.root.clipboard_append(pdf_path)
            messagebox.showinfo(
                "Doorsturen",
                f"Je standaard e-mailprogramma is geopend.\n\n"
                f"Het bestandspad is gekopieerd naar het klembord:\n{pdf_path}\n\n"
                f"Voeg het bestand handmatig toe als bijlage."
            )
        except Exception as e:
            messagebox.showerror("Fout", f"Kan e-mailprogramma niet openen:\n{str(e)}")

    # ─── Opslaan (formuliervelden + annotaties) ─────────────────────────

    def _has_unsaved_changes(self):
        """Controleer of de actieve tab onopgeslagen wijzigingen heeft"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return False
        # Formuliervelden ingevuld?
        if tab.form_widgets:
            self._save_form_widget_values(tab)
        if tab.form_field_values:
            return True
        # Tekst-annotaties geplaatst?
        if tab.text_annotations:
            return True
        # Markeringen (highlight annotaties)?
        if hasattr(tab, 'highlight_annotations') and tab.highlight_annotations:
            return True
        return False

    def _update_save_button_state(self):
        """Update de opslaan-knop: actief als er wijzigingen zijn, anders uitgegreyed"""
        try:
            if self._has_unsaved_changes():
                self.save_btn.config(state=tk.NORMAL)
            else:
                self.save_btn.config(state=tk.DISABLED)
        except (tk.TclError, Exception):
            pass

    def save_changes_to_pdf(self):
        """Sla alle wijzigingen (formuliervelden + annotaties + markeringen) op als nieuw PDF"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return

        if not self._has_unsaved_changes():
            return  # Niets op te slaan, knop zou disabled moeten zijn

        original = tab.file_path
        base, ext = os.path.splitext(original)
        suggested = f"{base}_bewerkt{ext}"
        save_path = filedialog.asksaveasfilename(
            title="Gewijzigde PDF opslaan als",
            initialfile=os.path.basename(suggested),
            initialdir=os.path.dirname(original),
            defaultextension=".pdf",
            filetypes=[("PDF bestanden", "*.pdf"), ("Alle bestanden", "*.*")]
        )
        if not save_path:
            return

        try:
            tmp_path = self._build_modified_pdf(tab)
            if tmp_path:
                import shutil
                shutil.move(tmp_path, save_path)
            else:
                # Geen wijzigingen gevonden, kopieer origineel
                import shutil
                shutil.copy2(original, save_path)

            # Wis de wijzigingen na succesvol opslaan
            tab.form_field_values = {}
            tab.text_annotations = []
            if hasattr(tab, 'highlight_annotations'):
                tab.highlight_annotations = []

            self._update_save_button_state()
            messagebox.showinfo("Opgeslagen",
                               f"PDF opgeslagen als:\n{os.path.basename(save_path)}")

        except Exception as e:
            messagebox.showerror("Fout", f"Kan PDF niet opslaan:\n{str(e)}")

    # ─── Formulier invulvelden ───────────────────────────────────────────

    def toggle_form_mode(self):
        """Schakel formuliermodus in/uit voor de actieve tab"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showinfo("Formulier", "Geen PDF geopend.")
            return

        tab.form_mode = not tab.form_mode

        if tab.form_mode:
            # Controleer of er formuliervelden bestaan
            has_fields = False
            for page_num in range(len(tab.pdf_document)):
                page = tab.pdf_document[page_num]
                widgets = page.widgets()
                if widgets:
                    for w in widgets:
                        has_fields = True
                        break
                if has_fields:
                    break

            if not has_fields:
                tab.form_mode = False
                messagebox.showinfo("Formulier", "Deze PDF bevat geen invulbare formuliervelden.")
                return

            # Deactiveer conflicterende modi
            if tab.text_annotate_mode:
                tab.text_annotate_mode = False
                tab.canvas.unbind("<Button-1>")
                tab.canvas.bind("<Button-1>", lambda e, t=tab: self.on_click(e, t))
                tab.canvas.config(cursor="arrow")
                self._set_toolbar_button_active("type-text", False)
            if self.highlight_mode:
                self.highlight_mode = False
                self._set_toolbar_button_active("marker", False)

            self._create_form_overlays(tab)
            self._set_toolbar_button_active("form", True)
        else:
            self._save_form_widget_values(tab)
            self._clear_form_overlays(tab)
            # Herteken pagina om ingevulde waarden als tekst te tonen
            self.display_page(tab)
            self._update_save_button_state()
            self._set_toolbar_button_active("form", False)

    def _draw_form_field_highlights(self, tab):
        """Teken lichtblauwe markeringen op formulierveld-posities (altijd zichtbaar)"""
        tab.canvas.delete("form_highlight")  # Verwijder oude markeringen
        tab.canvas.delete("form_values")     # Verwijder oude waarde-teksten
        zoom = tab.zoom_level
        doc = tab.pdf_document
        has_fields = False

        for page_num in range(len(doc)):
            page = doc[page_num]
            if page_num not in tab.page_regions:
                continue
            px0, py0, px1, py1 = tab.page_regions[page_num]

            for widget in page.widgets():
                if widget is None:
                    continue
                has_fields = True
                rect = widget.rect
                cx0 = rect.x0 * zoom + px0
                cy0 = rect.y0 * zoom + py0
                cx1 = rect.x1 * zoom + px0
                cy1 = rect.y1 * zoom + py0

                # Lichtblauwe achtergrond (semi-transparant effect via stipple)
                tab.canvas.create_rectangle(
                    cx0, cy0, cx1, cy1,
                    fill="#D0E8FF", outline="#80B8FF", width=1,
                    stipple="gray50",
                    tags="form_highlight"
                )

                # Toon opgeslagen waarden als tekst als formuliermodus UIT is
                if not tab.form_mode and widget.xref in tab.form_field_values:
                    val = tab.form_field_values[widget.xref]
                    display_val = ""
                    if widget.field_type in (2, 5):  # Checkbox / Radio
                        display_val = "☑" if val else "☐"
                    else:
                        display_val = str(val) if val else ""

                    if display_val:
                        font_size = max(int(9 * zoom), 7)
                        tab.canvas.create_text(
                            cx0 + 3, cy0 + 2, anchor="nw",
                            text=display_val,
                            font=("Segoe UI", font_size),
                            fill="#0055AA",
                            tags="form_values"
                        )

        # Klik op een gemarkeerd veld → activeer formuliermodus automatisch
        if has_fields and not tab.form_mode:
            tab.canvas.tag_bind("form_highlight", "<Button-1>",
                               lambda e: self._activate_form_mode_from_click(tab))

    def _activate_form_mode_from_click(self, tab):
        """Activeer formuliermodus wanneer op een blauw gemarkeerd veld wordt geklikt"""
        if not tab.form_mode:
            tab.form_mode = True
            self._create_form_overlays(tab)
            # Verberg de blauwe markeringen en waarde-teksten (overlays zitten er nu overheen)
            tab.canvas.delete("form_highlight")
            tab.canvas.delete("form_values")
            try:
                btn_widget = self.toolbar_buttons.get("form", (None,))[0]
                if btn_widget:
                    btn_widget.config(bg=self.theme["ACCENT_COLOR"], fg="white")
            except:
                pass

    def _create_form_overlays(self, tab):
        """Maak Tkinter widgets als overlay op de formuliervelden"""
        self._clear_form_overlays(tab)

        zoom = tab.zoom_level
        doc = tab.pdf_document

        for page_num in range(len(doc)):
            page = doc[page_num]

            # Haal canvas offset voor deze pagina op
            if page_num not in tab.page_regions:
                continue
            px0, py0, px1, py1 = tab.page_regions[page_num]

            for widget in page.widgets():
                if widget is None:
                    continue

                field_type = widget.field_type
                field_name = widget.field_name or ""
                field_xref = widget.xref
                rect = widget.rect  # fitz.Rect in PDF-coordinates

                # Bereken positie op canvas
                cx = rect.x0 * zoom + px0
                cy = rect.y0 * zoom + py0
                cw = max((rect.x1 - rect.x0) * zoom, 30)
                ch = max((rect.y1 - rect.y0) * zoom, 20)

                # Gebruik eerder opgeslagen waarde of huidige PDF-waarde
                current_value = tab.form_field_values.get(field_xref,
                                widget.field_value or "")

                tk_widget = None
                var = None

                # ── Tekstveld (type 7 = Text) ──
                if field_type == 7:  # PDF_WIDGET_TYPE_TEXT
                    is_multiline = bool(widget.field_flags & 4096) or ch > 30 * zoom
                    if is_multiline:
                        txt = tk.Text(tab.canvas, font=("Segoe UI", max(int(9 * zoom), 7)),
                                      bg="white", fg="black", relief="solid", bd=1,
                                      highlightcolor=self.theme["ACCENT_COLOR"],
                                      highlightthickness=1, wrap=tk.WORD, undo=True)
                        if current_value:
                            txt.insert("1.0", str(current_value))
                        tk_widget = txt
                    else:
                        entry = tk.Entry(tab.canvas, font=("Segoe UI", max(int(9 * zoom), 7)),
                                         bg="white", fg="black", relief="solid", bd=1,
                                         highlightcolor=self.theme["ACCENT_COLOR"],
                                         highlightthickness=1)
                        if current_value:
                            entry.insert(0, str(current_value))
                        tk_widget = entry

                # ── Checkbox (type 2 = CheckBox) ──
                elif field_type == 2:  # PDF_WIDGET_TYPE_CHECKBOX
                    var = tk.BooleanVar(value=current_value in ("Yes", "On", True, "true", "/Yes"))
                    cb = tk.Checkbutton(tab.canvas, variable=var,
                                       bg="white", activebackground="white",
                                       highlightthickness=0, bd=1, relief="solid")
                    tk_widget = cb

                # ── Radio button (type 5 = RadioButton) ──
                elif field_type == 5:  # PDF_WIDGET_TYPE_RADIOBUTTON
                    var = tk.BooleanVar(value=current_value in ("Yes", "On", True, "true", "/Yes"))
                    rb = tk.Checkbutton(tab.canvas, variable=var,
                                       bg="white", activebackground="white",
                                       indicatoron=True, selectcolor="white",
                                       highlightthickness=0, bd=1, relief="solid")
                    tk_widget = rb

                # ── Keuzelijst / Dropdown (type 3 = Choice / Listbox, type 4 = ComboBox) ──
                elif field_type in (3, 4):  # PDF_WIDGET_TYPE_LISTBOX / COMBOBOX
                    choices = widget.choice_values or []
                    combo = ttk.Combobox(tab.canvas, values=choices,
                                        font=("Segoe UI", max(int(9 * zoom), 7)),
                                        state="readonly" if choices else "normal")
                    if current_value and current_value in choices:
                        combo.set(current_value)
                    elif current_value:
                        combo.set(str(current_value))
                    tk_widget = combo

                # ── Onbekend type → tekstveld als fallback ──
                else:
                    entry = tk.Entry(tab.canvas, font=("Segoe UI", max(int(9 * zoom), 7)),
                                     bg="#FFFFDD", fg="black", relief="solid", bd=1)
                    if current_value:
                        entry.insert(0, str(current_value))
                    tk_widget = entry

                if tk_widget:
                    # Plaats widget op canvas
                    win_id = tab.canvas.create_window(
                        cx, cy, anchor="nw",
                        window=tk_widget,
                        width=int(cw), height=int(ch),
                        tags="form_overlay"
                    )
                    tab.form_widgets.append({
                        "win_id": win_id,
                        "widget": tk_widget,
                        "var": var,
                        "field_xref": field_xref,
                        "field_name": field_name,
                        "field_type": field_type,
                        "page_num": page_num,
                    })

    def _clear_form_overlays(self, tab):
        """Verwijder alle formulier overlay widgets"""
        for fw in tab.form_widgets:
            try:
                tab.canvas.delete(fw["win_id"])
                fw["widget"].destroy()
            except:
                pass
        tab.form_widgets = []

    def _save_form_widget_values(self, tab):
        """Sla de huidige waarden van de overlay widgets op in tab.form_field_values"""
        for fw in tab.form_widgets:
            xref = fw["field_xref"]
            ftype = fw["field_type"]
            var = fw["var"]
            widget = fw["widget"]

            try:
                if ftype == 2 or ftype == 5:  # Checkbox / Radio
                    tab.form_field_values[xref] = var.get()
                elif ftype in (3, 4):  # Combobox
                    tab.form_field_values[xref] = widget.get()
                elif isinstance(widget, tk.Text):  # Multiline tekstveld
                    tab.form_field_values[xref] = widget.get("1.0", "end-1c")
                else:  # Enkel-regel tekstveld
                    tab.form_field_values[xref] = widget.get()
            except:
                pass

    def save_form_to_pdf(self):
        """Sla de ingevulde formuliervelden op in een nieuw PDF-bestand"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return

        # Eerst waarden ophalen uit widgets als die nog bestaan
        if tab.form_widgets:
            self._save_form_widget_values(tab)

        if not tab.form_field_values:
            messagebox.showinfo("Formulier opslaan", "Geen ingevulde velden om op te slaan.")
            return

        # Vraag bestandsnaam
        original = tab.file_path
        base, ext = os.path.splitext(original)
        suggested = f"{base}_ingevuld{ext}"
        save_path = filedialog.asksaveasfilename(
            title="Formulier opslaan als",
            initialfile=os.path.basename(suggested),
            initialdir=os.path.dirname(original),
            defaultextension=".pdf",
            filetypes=[("PDF bestanden", "*.pdf"), ("Alle bestanden", "*.*")]
        )
        if not save_path:
            return

        try:
            fitz = get_fitz()
            # Open een verse kopie om te bewerken
            doc = fitz.open(original)

            for page_num in range(len(doc)):
                page = doc[page_num]
                for widget in page.widgets():
                    if widget is None:
                        continue
                    xref = widget.xref
                    if xref in tab.form_field_values:
                        val = tab.form_field_values[xref]
                        ftype = widget.field_type

                        if ftype in (2, 5):  # Checkbox / Radio
                            widget.field_value = bool(val)
                        else:  # Tekst / Combo / List
                            widget.field_value = str(val) if val else ""

                        widget.update()

            doc.save(save_path)
            doc.close()

            messagebox.showinfo("Formulier opgeslagen",
                               f"Het ingevulde formulier is opgeslagen als:\n{save_path}")

        except Exception as e:
            messagebox.showerror("Fout", f"Kan formulier niet opslaan:\n{str(e)}")

    # ─── Vrije tekst annotaties ──────────────────────────────────────────

    def toggle_text_annotate_mode(self):
        """Schakel tekst-annotatiemodus in/uit"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showinfo("Tekst toevoegen", "Geen PDF geopend.")
            return

        tab.text_annotate_mode = not tab.text_annotate_mode

        if tab.text_annotate_mode:
            # Formuliermodus uitzetten als die actief is
            if tab.form_mode:
                tab.form_mode = False
                self._save_form_widget_values(tab)
                self._clear_form_overlays(tab)
                self._set_toolbar_button_active("form", False)

            # Markeermodus uitzetten als die actief is
            if self.highlight_mode:
                self.highlight_mode = False
                self._set_toolbar_button_active("marker", False)

            # Bind klik-event op canvas (enkelklik)
            tab.canvas.bind("<Button-1>", self._on_text_annotate_click)
            tab.canvas.config(cursor="crosshair")

            self._set_toolbar_button_active("type-text", True)

            try:
                self.status_label.config(text="Tekst-modus: klik op de PDF om tekst te plaatsen")
            except:
                pass
        else:
            tab.canvas.unbind("<Button-1>")
            # Herstel de standaard klik-handler
            tab.canvas.bind("<Button-1>", lambda e, t=tab: self.on_click(e, t))
            tab.canvas.config(cursor="arrow")

            self._set_toolbar_button_active("type-text", False)

            try:
                self.status_label.config(text="Gereed")
            except:
                pass

    def _on_text_annotate_click(self, event):
        """Verwerk klik op canvas voor tekst-annotatie plaatsing"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab) or not tab.text_annotate_mode:
            return

        # Sluit bestaande editor(s) — slechts één tegelijk
        for w_info in list(tab.text_annot_widgets):
            try:
                tab.canvas.delete(w_info["win_id"])
                w_info["frame"].destroy()
            except Exception:
                pass
        tab.text_annot_widgets.clear()

        # Canvas-coördinaten (rekening houdend met scroll)
        cx = tab.canvas.canvasx(event.x)
        cy = tab.canvas.canvasy(event.y)

        # Bepaal op welke pagina geklikt is
        target_page = None
        for page_num, (px0, py0, px1, py1) in tab.page_regions.items():
            if px0 <= cx <= px1 and py0 <= cy <= py1:
                target_page = page_num
                break

        if target_page is None:
            return

        px0, py0, _, _ = tab.page_regions[target_page]
        zoom = tab.zoom_level

        # ── Laad iconen voor knoppen (verkleind naar 20x20) ──
        Image, ImageTk, _, _ = get_PIL()
        check_icon = None
        close_icon = None
        try:
            check_path = get_resource_path(os.path.join('icons', 'check.png'))
            close_path = get_resource_path(os.path.join('icons', 'close.png'))
            if os.path.exists(check_path):
                check_img = Image.open(check_path).resize((20, 20), Image.LANCZOS)
                check_icon = ImageTk.PhotoImage(check_img)
            if os.path.exists(close_path):
                close_img = Image.open(close_path).convert('RGBA').resize((20, 20), Image.LANCZOS)
                # Inverteer icoon in donker thema
                if self.theme == Theme.DARK:
                    from PIL import ImageOps
                    r, g, b, a = close_img.split()
                    inv = ImageOps.invert(Image.merge('RGB', (r, g, b)))
                    close_img = Image.merge('RGBA', (*inv.split(), a))
                close_icon = ImageTk.PhotoImage(close_img)
        except:
            pass

        # ── Sluit eventueel bestaand editor-venster (max 1 tegelijk) ──
        for old in list(tab.text_annot_widgets):
            try:
                tab.canvas.delete(old["win_id"])
                old["frame"].destroy()
            except Exception:
                pass
        tab.text_annot_widgets.clear()

        # ── Hoofdframe voor de editor ──
        is_dark = self.theme == Theme.DARK
        editor_bg = self.theme["BG_PRIMARY"]
        editor_fg = self.theme["TEXT_PRIMARY"]
        toolbar_bg = self.theme["BG_SECONDARY"]

        annot_frame = tk.Frame(tab.canvas, bg=editor_bg, bd=0, relief="flat",
                              highlightbackground=self.theme["ACCENT_COLOR"],
                              highlightthickness=2)

        # ── Toolbar BOVENIN met opties (twee rijen) ──
        toolbar = tk.Frame(annot_frame, bg=toolbar_bg)
        toolbar.pack(fill=tk.X, side=tk.TOP)

        # Rij 1: Lettertype en grootte
        row1 = tk.Frame(toolbar, bg=toolbar_bg)
        row1.pack(fill=tk.X, padx=4, pady=(4, 0))

        # Lettertype keuze
        font_map = {
            "Helvetica": "helv",
            "Times Roman": "tiro",
            "Courier": "cour",
        }
        font_var = tk.StringVar(value="Helvetica")
        tk.Label(row1, text="Lettertype:", font=("Segoe UI", 9, "bold"),
                bg=toolbar_bg, fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT, padx=(2, 2))
        font_combo = ttk.Combobox(row1, textvariable=font_var, width=13,
                                  values=list(font_map.keys()), state="readonly",
                                  font=("Segoe UI", 9))
        font_combo.pack(side=tk.LEFT, padx=(0, 10))

        # Lettergrootte
        size_var = tk.IntVar(value=11)
        tk.Label(row1, text="Grootte:", font=("Segoe UI", 9, "bold"),
                bg=toolbar_bg, fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT, padx=(0, 2))
        size_spin = tk.Spinbox(row1, from_=6, to=36, width=3, textvariable=size_var,
                              font=("Segoe UI", 10),
                              bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                              buttonbackground=toolbar_bg)
        size_spin.pack(side=tk.LEFT, padx=(0, 4))

        # Rij 2: Kleuren en knoppen
        row2 = tk.Frame(toolbar, bg=toolbar_bg)
        row2.pack(fill=tk.X, padx=4, pady=(2, 4))

        # Kleurkeuze met gekleurde knoppen
        color_var = tk.StringVar(value="black")
        colors_cfg = [
            ("Zwart", "black",     "#333333" if is_dark else "#222222", "white" if is_dark else "white"),
            ("Rood",  "red",       "#cc3333",                          "white"),
            ("Blauw", "#0066CC",   "#2266bb",                          "white"),
            ("Groen", "darkgreen", "#228833",                          "white"),
        ]
        for label, color_val, btn_bg, btn_fg in colors_cfg:
            rb = tk.Radiobutton(row2, text=f" {label} ", variable=color_var, value=color_val,
                               font=("Segoe UI", 9, "bold"),
                               bg=btn_bg, fg=btn_fg,
                               selectcolor=btn_bg,
                               activebackground=btn_bg,
                               activeforeground=btn_fg,
                               indicatoron=False, relief="raised", bd=1,
                               overrelief="groove", cursor="hand2")
            rb.pack(side=tk.LEFT, padx=2)

        # Bepaal of dit een bewerking is van een bestaande annotatie
        edit_index = getattr(self, '_editing_annot_index', None)
        self._editing_annot_index = None  # Reset

        # Akkoord/Annuleer/Verwijder knoppen RECHTS in rij 2
        def save_annotation():
            text_content = txt.get("1.0", "end-1c").strip()
            if text_content:
                pdf_x = (cx - px0) / zoom
                pdf_y = (cy - py0) / zoom
                font_size = size_var.get()
                text_color = color_var.get()
                chosen_font = font_map.get(font_var.get(), "helv")

                annot_data = {
                    "page_num": target_page,
                    "pdf_x": pdf_x,
                    "pdf_y": pdf_y,
                    "text": text_content,
                    "font_size": font_size,
                    "color": text_color,
                    "fontname": chosen_font,
                }
                if edit_index is not None and 0 <= edit_index < len(tab.text_annotations):
                    tab.text_annotations[edit_index] = annot_data
                else:
                    tab.text_annotations.append(annot_data)
            tab.canvas.delete(win_id)
            annot_frame.destroy()
            tab.text_annot_widgets = [w for w in tab.text_annot_widgets if w["win_id"] != win_id]
            self.display_page(tab)
            self._update_save_button_state()

        def delete_annotation():
            if edit_index is not None and 0 <= edit_index < len(tab.text_annotations):
                tab.text_annotations.pop(edit_index)
            tab.canvas.delete(win_id)
            annot_frame.destroy()
            tab.text_annot_widgets = [w for w in tab.text_annot_widgets if w["win_id"] != win_id]
            self.display_page(tab)
            self._update_save_button_state()

        def cancel_annotation():
            tab.canvas.delete(win_id)
            annot_frame.destroy()
            tab.text_annot_widgets = [w for w in tab.text_annot_widgets if w["win_id"] != win_id]

        # Akkoord knop met check.png icoon
        save_btn = tk.Button(row2, command=save_annotation,
                            bg=self.theme["ACCENT_COLOR"], fg="white",
                            relief="flat", cursor="hand2", padx=6, pady=2)
        if check_icon:
            save_btn.config(image=check_icon, text=" OK", compound=tk.LEFT)
            save_btn._icon = check_icon  # Voorkom garbage collection
        else:
            save_btn.config(text="OK", font=("Segoe UI", 9, "bold"))
        save_btn.pack(side=tk.RIGHT, padx=(2, 6))

        # Verwijder knop (alleen bij bewerken van bestaande annotatie)
        if edit_index is not None:
            del_btn = tk.Button(row2, text=" Verwijder", command=delete_annotation,
                               bg=self.theme["ERROR_COLOR"], fg="white",
                               font=("Segoe UI", 9, "bold"),
                               relief="flat", cursor="hand2", padx=6, pady=2)
            del_btn.pack(side=tk.RIGHT, padx=2)

        # Annuleer knop met close.png icoon
        cancel_btn = tk.Button(row2, command=cancel_annotation,
                              bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                              relief="flat", cursor="hand2", padx=6, pady=2)
        if close_icon:
            cancel_btn.config(image=close_icon)
            cancel_btn._icon = close_icon  # Voorkom garbage collection
        else:
            cancel_btn.config(text="✕", font=("Segoe UI", 9))
        cancel_btn.pack(side=tk.RIGHT, padx=2, pady=4)

        # ── Tekstveld ONDER de toolbar ──
        txt = tk.Text(annot_frame, font=("Segoe UI", max(int(10 * zoom), 9)),
                      bg=editor_bg, fg=editor_fg, wrap=tk.WORD, undo=True,
                      insertbackground=editor_fg,
                      bd=0, highlightthickness=0, padx=6, pady=4)
        txt.pack(fill=tk.BOTH, expand=True)

        # Vul editor met bestaande annotatie-gegevens bij bewerking
        if edit_index is not None and 0 <= edit_index < len(tab.text_annotations):
            existing = tab.text_annotations[edit_index]
            txt.insert("1.0", existing.get("text", ""))
            size_var.set(existing.get("font_size", 11))
            color_var.set(existing.get("color", "black"))
            # Herstel lettertype
            reverse_font_map = {v: k for k, v in font_map.items()}
            saved_font = existing.get("fontname", "helv")
            font_var.set(reverse_font_map.get(saved_font, "Helvetica"))

        # ── Positie berekenen: binnen zichtbaar venster houden ──
        editor_w = max(int(320 * zoom), 280)
        editor_h = max(int(140 * zoom), 130)

        # Zichtbare canvas-area
        vis_x0 = tab.canvas.canvasx(0)
        vis_y0 = tab.canvas.canvasy(0)
        vis_x1 = vis_x0 + tab.canvas.winfo_width()
        vis_y1 = vis_y0 + tab.canvas.winfo_height()

        # Corrigeer positie als editor buiten beeld zou vallen
        place_x = cx
        place_y = cy
        if place_x + editor_w > vis_x1 - 10:
            place_x = max(vis_x0 + 10, vis_x1 - editor_w - 10)
        if place_y + editor_h > vis_y1 - 10:
            place_y = max(vis_y0 + 10, vis_y1 - editor_h - 10)

        win_id = tab.canvas.create_window(place_x, place_y, anchor="nw",
                                          window=annot_frame,
                                          width=editor_w, height=editor_h,
                                          tags="text_annotation_editor")

        tab.text_annot_widgets.append({"win_id": win_id, "frame": annot_frame})
        txt.focus_set()

    def _edit_text_annotation(self, event, annot_index):
        """Markeer een bestaande annotatie voor bewerking.
        De daadwerkelijke editor wordt geopend door _on_text_annotate_click
        die daarna via de canvas <Button-1> binding wordt aangeroepen."""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab) or not tab.text_annotate_mode:
            return
        if annot_index < 0 or annot_index >= len(tab.text_annotations):
            return
        # Zet index; _on_text_annotate_click pikt dit op
        self._editing_annot_index = annot_index

    def _draw_text_annotations(self, tab):
        """Teken opgeslagen tekst-annotaties op de canvas"""
        tab.canvas.delete("text_annotation")
        zoom = tab.zoom_level

        for idx, annot in enumerate(tab.text_annotations):
            page_num = annot["page_num"]
            if page_num not in tab.page_regions:
                continue

            px0, py0, _, _ = tab.page_regions[page_num]
            cx = annot["pdf_x"] * zoom + px0
            cy = annot["pdf_y"] * zoom + py0
            font_size = max(int(annot["font_size"] * zoom), 7)
            color = annot.get("color", "black")
            display_color = color

            # Vertaal PDF-fontnaam naar Tkinter-fontnaam
            pdf_fontname = annot.get("fontname", "helv")
            tk_font_map = {"helv": "Helvetica", "tiro": "Times New Roman", "cour": "Courier New"}
            tk_font = tk_font_map.get(pdf_fontname, "Helvetica")

            annot_tag = f"text_annot_{idx}"
            item_id = tab.canvas.create_text(
                cx, cy, anchor="nw",
                text=annot["text"],
                font=(tk_font, font_size),
                fill=display_color,
                tags=("text_annotation", annot_tag)
            )

            # Klik-binding voor bewerken/verwijderen (alleen in tekst-modus)
            tab.canvas.tag_bind(annot_tag, "<Button-1>",
                                lambda e, i=idx: self._edit_text_annotation(e, i))

    def save_annotations_to_pdf(self):
        """Sla alle tekst-annotaties op in een nieuw PDF-bestand"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return

        if not tab.text_annotations:
            messagebox.showinfo("Annotaties opslaan", "Geen tekst-annotaties om op te slaan.")
            return

        original = tab.file_path
        base, ext = os.path.splitext(original)
        suggested = f"{base}_annotaties{ext}"
        save_path = filedialog.asksaveasfilename(
            title="PDF met annotaties opslaan als",
            initialfile=os.path.basename(suggested),
            initialdir=os.path.dirname(original),
            defaultextension=".pdf",
            filetypes=[("PDF bestanden", "*.pdf"), ("Alle bestanden", "*.*")]
        )
        if not save_path:
            return

        try:
            fitz = get_fitz()
            doc = fitz.open(original)

            # Kleur-mapping naar RGB tuples (0-1 bereik)
            color_map = {
                "black": (0, 0, 0),
                "red": (1, 0, 0),
                "blue": (0, 0, 1),
                "darkgreen": (0, 0.4, 0),
            }

            for annot in tab.text_annotations:
                page = doc[annot["page_num"]]
                pdf_x = annot["pdf_x"]
                pdf_y = annot["pdf_y"]
                text = annot["text"]
                font_size = annot["font_size"]
                color = color_map.get(annot.get("color", "black"), (0, 0, 0))

                # Bereken tekst breedte/hoogte voor de rect
                lines = text.split("\n")
                max_line_len = max(len(line) for line in lines) if lines else 10
                approx_width = max_line_len * font_size * 0.55
                approx_height = len(lines) * font_size * 1.4

                rect = fitz.Rect(pdf_x, pdf_y,
                                pdf_x + approx_width + 10,
                                pdf_y + approx_height + 5)

                fontname = annot.get("fontname", "helv")
                annot_obj = page.add_freetext_annot(
                    rect, text,
                    fontsize=font_size,
                    fontname=fontname,
                    text_color=color,
                    fill_color=(1, 1, 1),  # Witte achtergrond
                    border_color=None,
                )
                annot_obj.update()

            doc.save(save_path)
            doc.close()

            messagebox.showinfo("Annotaties opgeslagen",
                               f"PDF met annotaties opgeslagen als:\n{save_path}")
        except Exception as e:
            messagebox.showerror("Fout", f"Kan annotaties niet opslaan:\n{str(e)}")

    def show_pdf_info(self):
        tab = self.get_active_tab()
        if isinstance(tab, PDFTab):
            metadata = tab.pdf_document.metadata
            info_text = (
                f"Titel: {metadata.get('title', 'N/A')}\n"
                f"Auteur: {metadata.get('author', 'N/A')}\n"
                f"Onderwerp: {metadata.get('subject', 'N/A')}\n"
                f"Trefwoorden: {metadata.get('keywords', 'N/A')}\n"
                f"Creator: {metadata.get('creator', 'N/A')}\n"
                f"Producer: {metadata.get('producer', 'N/A')}\n"
                f"Gemaakt: {metadata.get('creationDate', 'N/A')}\n"
                f"Gewijzigd: {metadata.get('modDate', 'N/A')}\n"
                f"Pagina's: {len(tab.pdf_document)}\n"
                f"Bestandsgrootte: {os.path.getsize(tab.file_path) / 1024:.1f} KB"
            )
            messagebox.showinfo("PDF Informatie", info_text)

    def show_edit_menu(self):
        """Toon bewerkingsmenu met opties"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak popup menu
        edit_menu = tk.Toplevel(self.root)
        edit_menu.title("PDF Bewerken")
        edit_menu.geometry("400x500")
        edit_menu.configure(bg=self.theme["BG_PRIMARY"])
        edit_menu.transient(self.root)
        edit_menu.grab_set()
        edit_menu.resizable(False, False)
        
        # Probeer icoon te laden
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                edit_menu.iconbitmap(icon_path)
        except:
            pass
        
        # Header
        header = tk.Frame(edit_menu, bg=self.theme["ACCENT_COLOR"], height=60)
        header.pack(fill=tk.X)
        header.pack_propagate(False)
        
        tk.Label(header, text="📝 PDF Bewerken", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content
        content = tk.Frame(edit_menu, bg=self.theme["BG_PRIMARY"])
        content.pack(fill=tk.BOTH, expand=True, padx=20, pady=20)
        
        # Opties
        tk.Label(content, text="Kies een bewerkingsoptie:", font=("Segoe UI", 10, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(0, 15))
        
        # Pagina's exporteren
        export_frame = self.create_menu_option(content, 
                                               "📄 Pagina's Exporteren",
                                               "Sla geselecteerde pagina's op als nieuw PDF bestand",
                                               lambda: [edit_menu.destroy(), self.export_pages()])
        export_frame.pack(fill=tk.X, pady=5)
        
        # PDF's samenvoegen
        merge_frame = self.create_menu_option(content,
                                              "📑 PDF's Samenvoegen",
                                              "Voeg meerdere PDF bestanden samen tot één",
                                              lambda: [edit_menu.destroy(), self.merge_pdfs()])
        merge_frame.pack(fill=tk.X, pady=5)
        
        # Pagina roteren
        rotate_frame = self.create_menu_option(content,
                                              "🔄 Pagina Roteren",
                                              "Roteer geselecteerde pagina's 90° / 180° / 270°",
                                              lambda: [edit_menu.destroy(), self.rotate_pages()])
        rotate_frame.pack(fill=tk.X, pady=5)

        # Footer met knop (moderne stijl)
        footer_frame = tk.Frame(edit_menu, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        tk.Button(footer_frame, text="Sluiten", command=edit_menu.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(pady=15)

    def create_menu_option(self, parent, title, description, command):
        """Maak een menu optie frame"""
        frame = tk.Frame(parent, bg=self.theme["BG_SECONDARY"], 
                        highlightbackground=self.theme["TEXT_SECONDARY"],
                        highlightthickness=1, cursor="hand2")
        
        inner = tk.Frame(frame, bg=self.theme["BG_SECONDARY"])
        inner.pack(fill=tk.BOTH, expand=True, padx=15, pady=10)
        
        tk.Label(inner, text=title, font=("Segoe UI", 10, "bold"),
                bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                anchor="w").pack(fill=tk.X)
        
        tk.Label(inner, text=description, font=("Segoe UI", 8),
                bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_SECONDARY"],
                anchor="w", wraplength=340).pack(fill=tk.X, pady=(5, 0))
        
        frame.bind("<Button-1>", lambda e: command())
        for widget in frame.winfo_children():
            widget.bind("<Button-1>", lambda e: command())
            for child in widget.winfo_children():
                child.bind("<Button-1>", lambda e: command())
        
        def on_enter(e):
            frame.config(bg=self.theme["ACCENT_COLOR"])
            inner.config(bg=self.theme["ACCENT_COLOR"])
            for w in inner.winfo_children():
                w.config(bg=self.theme["ACCENT_COLOR"], fg="white")
        
        def on_leave(e):
            frame.config(bg=self.theme["BG_SECONDARY"])
            inner.config(bg=self.theme["BG_SECONDARY"])
            for i, w in enumerate(inner.winfo_children()):
                w.config(bg=self.theme["BG_SECONDARY"])
                if i == 0:
                    w.config(fg=self.theme["TEXT_PRIMARY"])
                else:
                    w.config(fg=self.theme["TEXT_SECONDARY"])
        
        frame.bind("<Enter>", on_enter)
        frame.bind("<Leave>", on_leave)
        
        return frame

    def export_pages(self):
        """Exporteer geselecteerde pagina's"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak moderne dialoog met header
        dialog = tk.Toplevel(self.root)
        dialog.title("Pagina's Exporteren")
        dialog.geometry("500x480")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📄 Pagina's Exporteren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Document info
        tk.Label(content_frame, 
                text=f"Document: {os.path.basename(tab.file_path)}\nTotaal aantal pagina's: {len(tab.pdf_document)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_SECONDARY"],
                justify=tk.LEFT).pack(pady=(0, 20), anchor="w")
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's wilt u exporteren?", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        pages_var = tk.StringVar()
        entry = tk.Entry(content_frame, textvariable=pages_var, font=("Segoe UI", 10), width=40)
        entry.pack(pady=8, fill=tk.X)
        entry.focus()
        
        # Voorbeelden in een nette box
        examples_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], 
                                 relief="flat", bd=1)
        examples_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(examples_frame, text="Voorbeelden:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", padx=15, pady=(10, 5))
        
        examples = [
            "• 1,3,5 (specifieke pagina's)",
            "• 1-5 (bereik)",
            "• 1-3,5,7-9 (combinatie)"
        ]
        
        for example in examples:
            tk.Label(examples_frame, text=example, 
                    font=("Segoe UI", 9),
                    bg=self.theme["BG_SECONDARY"], 
                    fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w", padx=25, pady=2)
        
        tk.Label(examples_frame, text=" ", bg=self.theme["BG_SECONDARY"]).pack(pady=5)
        
        def do_export():
            fitz = get_fitz()  # Lazy load
            pages_input = pages_var.get()
            
            if not pages_input:
                messagebox.showwarning("Geen invoer", "Voer pagina nummers in")
                return
            
            # Parse pagina's
            pages = self.parse_page_range(pages_input, len(tab.pdf_document))
            
            if not pages:
                messagebox.showerror("Ongeldige invoer", "Ongeldige pagina selectie!")
                return
            
            # Vraag opslag locatie
            save_path = filedialog.asksaveasfilename(
                title="Exporteer pagina's als",
                defaultextension=".pdf",
                filetypes=[("PDF Bestanden", "*.pdf"), ("Alle Bestanden", "*.*")]
            )
            
            if not save_path:
                return
            
            try:
                # Maak nieuw PDF document
                new_doc = fitz.open()
                
                for page_num in pages:
                    new_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                
                new_doc.save(save_path)
                new_doc.close()
                
                dialog.destroy()
                messagebox.showinfo("Succes", 
                    f"{len(pages)} pagina('s) succesvol geëxporteerd naar:\n{os.path.basename(save_path)}")
                
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet exporteren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Exporteren", command=do_export,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        # Enter key binding
        entry.bind("<Return>", lambda e: do_export())

    def rotate_pages(self):
        """Roteer geselecteerde pagina's"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        
        # Maak rotatie dialoog
        rotate_dialog = tk.Toplevel(self.root)
        rotate_dialog.title("Pagina's Roteren")
        rotate_dialog.geometry("500x420")
        rotate_dialog.configure(bg=self.theme["BG_PRIMARY"])
        rotate_dialog.transient(self.root)
        rotate_dialog.grab_set()
        rotate_dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                rotate_dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(rotate_dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🔄 Pagina's Roteren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(rotate_dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's?", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        page_var = tk.StringVar(value=f"{tab.current_page + 1}")
        tk.Entry(content_frame, textvariable=page_var, font=("Segoe UI", 10),
                width=40).pack(pady=8, fill=tk.X)
        
        tk.Label(content_frame, text="(bijv: 1,3,5 of 1-5)", font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w")
        
        # Rotatie hoek
        tk.Label(content_frame, text="Rotatie:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(15, 5))
        
        rotation_var = tk.IntVar(value=90)
        
        for angle in [90, 180, 270]:
            tk.Radiobutton(content_frame, text=f"{angle}° (rechtsom)" if angle == 90 else f"{angle}°",
                          variable=rotation_var, value=angle,
                          bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                          selectcolor=self.theme["BG_SECONDARY"],
                          activebackground=self.theme["BG_PRIMARY"],
                          activeforeground=self.theme["TEXT_PRIMARY"],
                          font=("Segoe UI", 9)).pack(anchor="w", padx=20, pady=2)
        
        def do_rotate():
            pages_str = page_var.get()
            pages = self.parse_page_range(pages_str, len(tab.pdf_document))
            
            if not pages:
                messagebox.showerror("Ongeldige invoer", "Ongeldige pagina selectie!")
                return
            
            rotation = rotation_var.get()
            
            try:
                for page_num in pages:
                    page = tab.pdf_document[page_num]
                    page.set_rotation(rotation)
                
                # Ververs weergave
                self.display_page(tab)
                rotate_dialog.destroy()
                
                messagebox.showinfo("Succes", 
                    f"{len(pages)} pagina('s) geroteerd met {rotation}°\n\n" +
                    "Vergeet niet op te slaan om wijzigingen te behouden!")
                
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet roteren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(rotate_dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Roteren", command=do_rotate,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=rotate_dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)

    def exit_application(self):
        # Sla window geometry en state op voordat we afsluiten
        try:
            self.update_settings['window_geometry'] = self.root.geometry()
            self.update_settings['window_state'] = self.root.state()
            self.save_update_settings()
        except:
            pass  # Als opslaan mislukt, sluit gewoon af
        
        num_tabs = sum(1 for tab_id in self.notebook.tabs() 
                      if isinstance(self.notebook.nametowidget(tab_id), PDFTab))

        if num_tabs > 1:
            answer = messagebox.askyesno(
                "Afsluiten bevestigen",
                f"Er zijn {num_tabs} documenten geopend. Weet u zeker dat u wilt afsluiten?"
            )
            if not answer:
                return

        for tab_id in self.notebook.tabs():
            tab = self.notebook.nametowidget(tab_id)
            if isinstance(tab, PDFTab):
                tab.close_document()
        
        self.root.quit()
        self.root.destroy()
        sys.exit(0)

    def extract_pages(self):
        """Extraheer specifieke pagina's naar een nieuwe PDF"""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            messagebox.showwarning("Geen document", "Open eerst een PDF document")
            return
        
        # Maak moderne dialoog met header
        dialog = tk.Toplevel(self.root)
        dialog.title("Pagina's extraheren")
        dialog.geometry("500x380")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📄 Pagina's Extraheren", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Document info
        tk.Label(content_frame, text=f"Document: {os.path.basename(tab.file_path)}", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 20))
        
        # Pagina selectie
        tk.Label(content_frame, text="Welke pagina's?", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        pages_var = tk.StringVar()
        tk.Entry(content_frame, textvariable=pages_var, font=("Segoe UI", 10),
                width=40).pack(pady=5, fill=tk.X)
        
        tk.Label(content_frame, text="(bijv: 1,3,5 of 1-5)", font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(anchor="w")
        
        # Bereik selectie
        tk.Label(content_frame, text="Of selecteer bereik:", font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(15, 5))
        
        from_var = tk.StringVar(value="1")
        to_var = tk.StringVar(value=str(len(tab.pdf_document)))
        
        # Bereik layout
        range_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        range_frame.pack(anchor="w", pady=5)
        
        tk.Label(range_frame, text="Van:", font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=from_var, width=8,
                font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=(5, 15))
        
        tk.Label(range_frame, text="Tot:", font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(side=tk.LEFT)
        tk.Entry(range_frame, textvariable=to_var, width=8,
                font=("Segoe UI", 10)).pack(side=tk.LEFT, padx=5)
        
        def do_extract():
            fitz = get_fitz()  # Lazy load
            try:
                # Bepaal welke pagina's te extraheren
                if pages_var.get():
                    pages = self.parse_page_range(pages_var.get(), len(tab.pdf_document))
                else:
                    from_page = int(from_var.get()) - 1
                    to_page = int(to_var.get()) - 1
                    pages = list(range(from_page, to_page + 1))
                
                if not pages:
                    messagebox.showerror("Fout", "Ongeldige pagina selectie")
                    return
                
                # Vraag waar op te slaan
                save_path = filedialog.asksaveasfilename(
                    defaultextension=".pdf",
                    filetypes=[("PDF Bestanden", "*.pdf")]
                )
                
                if save_path:
                    # Maak nieuwe PDF met geselecteerde pagina's
                    new_doc = fitz.open()
                    for page_num in pages:
                        new_doc.insert_pdf(tab.pdf_document, from_page=page_num, to_page=page_num)
                    
                    new_doc.save(save_path)
                    new_doc.close()
                    
                    dialog.destroy()
                    messagebox.showinfo("Succes", 
                        f"{len(pages)} pagina's geëxtraheerd naar:\n{os.path.basename(save_path)}")
                    
            except Exception as e:
                messagebox.showerror("Fout", f"Kan pagina's niet extraheren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Extraheren", command=do_extract,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)

    def merge_pdfs(self):
        """Combineer meerdere PDF bestanden"""
        dialog = tk.Toplevel(self.root)
        dialog.title("PDF's combineren")
        dialog.geometry("550x570")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(dialog, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="📑 PDF's Samenvoegen", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Optie voor openstaande bestanden
        option_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        option_frame.pack(fill=tk.X, pady=(0, 10))
        
        def add_open_tabs():
            """Voeg alle openstaande PDF's toe aan de lijst"""
            added_count = 0
            for tab_id in self.notebook.tabs():
                tab = self.notebook.nametowidget(tab_id)
                if isinstance(tab, PDFTab):
                    if tab.file_path not in pdf_files:
                        pdf_files.append(tab.file_path)
                        listbox.insert(tk.END, os.path.basename(tab.file_path))
                        added_count += 1
            
            if added_count > 0:
                messagebox.showinfo("Toegevoegd", 
                    f"{added_count} openstaande PDF{'s' if added_count > 1 else ''} toegevoegd aan de lijst")
            else:
                messagebox.showinfo("Info", "Alle openstaande PDF's staan al in de lijst")
        
        tk.Button(option_frame, text="📂 Voeg openstaande PDF's toe", command=add_open_tabs,
                 bg=self.theme["SUCCESS_COLOR"], fg="white",
                 font=("Segoe UI", 9, "bold"), padx=15, pady=8,
                 relief="flat", cursor="hand2").pack(anchor="w")
        
        # Lijst van bestanden
        tk.Label(content_frame, text="Geselecteerde PDF bestanden:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w")
        
        listbox = tk.Listbox(content_frame, height=10, font=("Segoe UI", 9))
        listbox.pack(fill=tk.BOTH, expand=True, pady=5)
        
        scrollbar = tk.Scrollbar(listbox)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)
        listbox.config(yscrollcommand=scrollbar.set)
        scrollbar.config(command=listbox.yview)
        
        pdf_files = []
        
        def add_files():
            files = filedialog.askopenfilenames(
                title="Selecteer PDF bestanden",
                filetypes=[("PDF Bestanden", "*.pdf")]
            )
            for file in files:
                if file not in pdf_files:
                    pdf_files.append(file)
                    listbox.insert(tk.END, os.path.basename(file))
        
        def remove_file():
            selection = listbox.curselection()
            if selection:
                index = selection[0]
                listbox.delete(index)
                pdf_files.pop(index)
        
        def move_up():
            selection = listbox.curselection()
            if selection and selection[0] > 0:
                index = selection[0]
                # Swap in list
                pdf_files[index], pdf_files[index-1] = pdf_files[index-1], pdf_files[index]
                # Swap in listbox
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index-1, item)
                listbox.selection_set(index-1)
        
        def move_down():
            selection = listbox.curselection()
            if selection and selection[0] < len(pdf_files) - 1:
                index = selection[0]
                # Swap in list
                pdf_files[index], pdf_files[index+1] = pdf_files[index+1], pdf_files[index]
                # Swap in listbox
                item = listbox.get(index)
                listbox.delete(index)
                listbox.insert(index+1, item)
                listbox.selection_set(index+1)
        
        # Knoppen voor lijst beheer
        list_btn_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        list_btn_frame.pack(pady=10)
        
        tk.Button(list_btn_frame, text="➕ Toevoegen", command=add_files,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="➖ Verwijderen", command=remove_file,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="⬆ Omhoog", command=move_up,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        tk.Button(list_btn_frame, text="⬇ Omlaag", command=move_down,
                 bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 9), padx=10, pady=5,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=2)
        
        def do_merge():
            fitz = get_fitz()  # Lazy load
            if len(pdf_files) < 2:
                messagebox.showwarning("Niet genoeg bestanden", 
                    "Selecteer tenminste 2 PDF bestanden om te combineren")
                return
            
            save_path = filedialog.asksaveasfilename(
                defaultextension=".pdf",
                filetypes=[("PDF Bestanden", "*.pdf")],
                title="Gecombineerde PDF opslaan als"
            )
            
            if save_path:
                try:
                    merged_doc = fitz.open()
                    for pdf_file in pdf_files:
                        pdf_doc = fitz.open(pdf_file)
                        merged_doc.insert_pdf(pdf_doc)
                        pdf_doc.close()
                    
                    merged_doc.save(save_path)
                    merged_doc.close()
                    
                    dialog.destroy()
                    messagebox.showinfo("Succes", 
                        f"{len(pdf_files)} PDF's gecombineerd naar:\n{os.path.basename(save_path)}")
                    
                    # Vraag of gebruiker het gecombineerde bestand wil openen
                    if messagebox.askyesno("Openen?", "Wilt u het gecombineerde bestand openen?"):
                        self.add_new_tab(save_path)
                        
                except Exception as e:
                    messagebox.showerror("Fout", f"Kan PDF's niet combineren:\n{str(e)}")
        
        # Footer met knoppen (moderne stijl)
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        tk.Button(btn_container, text="Combineren", command=do_merge,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)
        tk.Button(btn_container, text="Annuleren", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=25, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=5)


    def show_about(self):
        """Toon Over dialoog"""
        about = tk.Toplevel(self.root)
        about.title("Over NVict Reader")
        about.geometry("450x600")
        about.configure(bg=self.theme["BG_PRIMARY"])
        about.transient(self.root)
        about.resizable(False, False)
        
        # Voeg favicon toe aan taakbalk
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                about.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur (moderne stijl)
        header_frame = tk.Frame(about, bg=self.theme["ACCENT_COLOR"], height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="Over NVict Reader", font=("Segoe UI", 14, "bold"),
                bg=self.theme["ACCENT_COLOR"], fg="white").pack(pady=15)
        
        # Content frame
        content_frame = tk.Frame(about, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Logo in content (met witte achtergrond voor betere zichtbaarheid)
        try:
            Image, ImageTk, ImageOps, ImageDraw = get_PIL()  # Lazy load PIL
            logo_path = get_resource_path('Logo.png')
            if os.path.exists(logo_path):
                logo_image = Image.open(logo_path)
                logo_image.thumbnail((80, 80), Image.Resampling.LANCZOS)
                
                # Maak witte achtergrond voor logo
                bg_size = 100
                background = Image.new('RGB', (bg_size, bg_size), 'white')
                
                # Centreer logo op witte achtergrond
                offset = ((bg_size - logo_image.size[0]) // 2, (bg_size - logo_image.size[1]) // 2)
                if logo_image.mode == 'RGBA':
                    background.paste(logo_image, offset, logo_image)
                else:
                    background.paste(logo_image, offset)
                
                logo_photo = ImageTk.PhotoImage(background)
                logo_label = tk.Label(content_frame, image=logo_photo, bg=self.theme["BG_PRIMARY"])
                logo_label.image = logo_photo  # Keep reference
                logo_label.pack(pady=(10, 15))
        except Exception as e:
            print(f"Logo laden mislukt: {e}")
            pass
        
        # Titel
        tk.Label(content_frame, text="NVict Reader", 
                font=("Segoe UI", 16, "bold"),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"]).pack(pady=(0, 5))
        
        tk.Label(content_frame, text=f"Versie {APP_VERSION}", 
                font=("Segoe UI", 10),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 5))
        
        tk.Label(content_frame, text=f"© {self.get_current_year()} NVict Service", 
                font=("Segoe UI", 9),
                bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_SECONDARY"]).pack(pady=(0, 20))
        
        # Features in een mooie box
        features_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], 
                                 relief="flat", bd=1)
        features_frame.pack(fill=tk.X, pady=10)
        
        tk.Label(features_frame, text="Functies:", 
                font=("Segoe UI", 9, "bold"),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", padx=20, pady=(15, 5))
        
        feature_list = [
            "✓ PDF's openen en bekijken",
            "✓ Tekst selecteren en kopiëren",
            "✓ Formulieren invullen",
            "✓ Pagina's exporteren",
            "✓ PDF's samenvoegen",
            "✓ Pagina's roteren",
            "✓ Afdrukken met opties",
            "✓ Zoeken in documenten"
        ]
        
        for feature in feature_list:
            tk.Label(features_frame, text=feature, font=("Segoe UI", 9),
                    bg=self.theme["BG_SECONDARY"], fg=self.theme["TEXT_PRIMARY"],
                    anchor="w").pack(anchor="w", padx=30, pady=2)
        
        tk.Label(features_frame, text=" ", bg=self.theme["BG_SECONDARY"]).pack(pady=5)
        
        def open_website():
            webbrowser.open("https://www.nvict.nl/software.html")
        
        link_label = tk.Label(content_frame, text="www.nvict.nl",
                             font=("Segoe UI", 9, "underline"),
                             bg=self.theme["BG_PRIMARY"],
                             fg=self.theme["ACCENT_COLOR"],
                             cursor="hand2")
        link_label.pack(pady=(15, 5))
        link_label.bind("<Button-1>", lambda e: open_website())

        # Iconen credit
        icon_credit = tk.Label(content_frame,
                              text="Iconen door Freepik - Flaticon",
                              font=("Segoe UI", 8, "underline"),
                              bg=self.theme["BG_PRIMARY"],
                              fg=self.theme["TEXT_SECONDARY"],
                              cursor="hand2")
        icon_credit.pack(pady=(0, 10))
        icon_credit.bind("<Button-1>", lambda e: webbrowser.open(
            "https://www.flaticon.com/free-icons/page"))

        # Footer met knop (moderne stijl)
        footer_frame = tk.Frame(about, bg=self.theme["BG_SECONDARY"], height=70)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        tk.Button(footer_frame, text="Sluiten", command=about.destroy,
                 bg=self.theme["ACCENT_COLOR"], fg="white", 
                 font=("Segoe UI", 10), padx=30, pady=10,
                 relief="flat", cursor="hand2").pack(pady=15)

    def set_as_default_pdf(self):
        """Prompt user to set NVict Reader as default PDF viewer"""
        DefaultPDFHandler.prompt_set_as_default(self.root)

    def check_for_updates(self, silent=False):
        """Controleer of er updates beschikbaar zijn"""
        try:
            # Download versie info van server
            with urllib.request.urlopen(UPDATE_CHECK_URL, timeout=5) as response:
                data = json.loads(response.read().decode('utf-8'))
                
                latest_version = data.get("version", "0.0")
                download_url = data.get("download_url", "")
                release_notes = data.get("release_notes", "")
                
                # Vergelijk versies
                current_parts = [int(x) for x in APP_VERSION.split('.')]
                latest_parts = [int(x) for x in latest_version.split('.')]
                
                # Pad version parts als ze verschillende lengtes hebben
                max_length = max(len(current_parts), len(latest_parts))
                current_parts += [0] * (max_length - len(current_parts))
                latest_parts += [0] * (max_length - len(latest_parts))
                
                update_available = latest_parts > current_parts
                
                if update_available:
                    self.show_update_dialog(latest_version, download_url, release_notes)
                else:
                    if not silent:
                        messagebox.showinfo("Geen updates", 
                            f"U gebruikt al de nieuwste versie ({APP_VERSION})")
                        
        except urllib.error.URLError:
            if not silent:
                messagebox.showerror("Verbindingsfout", 
                    "Kan niet verbinden met de update server.\n\n"
                    "Controleer uw internetverbinding en probeer het later opnieuw.")
        except Exception as e:
            if not silent:
                messagebox.showerror("Fout", 
                    f"Fout bij controleren op updates:\n{str(e)}")

    def show_update_dialog(self, new_version, download_url, release_notes):
        """Toon dialoog met update informatie en automatische download/installatie optie"""
        dialog = tk.Toplevel(self.root)
        dialog.title("Update Beschikbaar")
        dialog.geometry("520x550")
        dialog.configure(bg=self.theme["BG_PRIMARY"])
        dialog.transient(self.root)
        dialog.grab_set()
        dialog.resizable(False, False)
        
        # Icon toevoegen
        try:
            icon_path = get_resource_path('favicon.ico')
            if os.path.exists(icon_path):
                dialog.iconbitmap(icon_path)
        except:
            pass
        
        # Header met accent kleur
        header_frame = tk.Frame(dialog, bg=self.theme["SUCCESS_COLOR"], height=80)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        tk.Label(header_frame, text="🎉 Update Beschikbaar!", 
                font=("Segoe UI", 16, "bold"),
                bg=self.theme["SUCCESS_COLOR"], fg="white").pack(pady=25)
        
        # Content frame
        content_frame = tk.Frame(dialog, bg=self.theme["BG_PRIMARY"])
        content_frame.pack(fill=tk.BOTH, expand=True, padx=30, pady=20)
        
        # Versie informatie
        version_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"], relief="flat")
        version_frame.pack(fill=tk.X, pady=(0, 20))
        
        info_text = f"Huidige versie: {APP_VERSION}\nNieuwe versie: {new_version}"
        tk.Label(version_frame, text=info_text, 
                font=("Segoe UI", 10),
                bg=self.theme["BG_SECONDARY"], 
                fg=self.theme["TEXT_PRIMARY"],
                justify=tk.LEFT).pack(padx=20, pady=15)
        
        # Release notes
        if release_notes:
            tk.Label(content_frame, text="Wat is er nieuw:", 
                    font=("Segoe UI", 10, "bold"),
                    bg=self.theme["BG_PRIMARY"], 
                    fg=self.theme["TEXT_PRIMARY"]).pack(anchor="w", pady=(0, 5))
            
            notes_frame = tk.Frame(content_frame, bg=self.theme["BG_SECONDARY"])
            notes_frame.pack(fill=tk.BOTH, expand=True)
            
            # Scrollbar toevoegen
            scrollbar = tk.Scrollbar(notes_frame, bg=self.theme["BG_SECONDARY"])
            scrollbar.pack(side=tk.RIGHT, fill=tk.Y, padx=(0, 2), pady=10)
            
            notes_text = tk.Text(notes_frame, wrap=tk.WORD, 
                               font=("Segoe UI", 9),
                               bg=self.theme["BG_SECONDARY"],
                               fg=self.theme["TEXT_PRIMARY"],
                               relief="flat", height=8,
                               yscrollcommand=scrollbar.set)
            notes_text.pack(side=tk.LEFT, fill=tk.BOTH, expand=True, padx=10, pady=10)
            notes_text.insert("1.0", release_notes)
            notes_text.config(state=tk.DISABLED)
            
            # Koppel scrollbar aan text widget
            scrollbar.config(command=notes_text.yview)
        
        # Installatie instructie
        install_frame = tk.Frame(content_frame, bg=self.theme["BG_PRIMARY"])
        install_frame.pack(fill=tk.X, pady=(10, 0))
        
        tk.Label(install_frame, 
                text="ℹ️ Na download wordt de installer automatisch geopend.\nSluit eerst NVict Reader af voordat u installeert.",
                font=("Segoe UI", 8),
                bg=self.theme["BG_PRIMARY"], 
                fg=self.theme["TEXT_SECONDARY"],
                justify=tk.LEFT).pack(anchor="w")
        
        # Footer met knoppen
        footer_frame = tk.Frame(dialog, bg=self.theme["BG_SECONDARY"], height=80)
        footer_frame.pack(fill=tk.X, side=tk.BOTTOM)
        footer_frame.pack_propagate(False)
        
        btn_container = tk.Frame(footer_frame, bg=self.theme["BG_SECONDARY"])
        btn_container.pack(expand=True)
        
        def download_and_install():
            """Download update en start installatie automatisch"""
            if download_url:
                dialog.destroy()
                self.download_and_install_update(download_url, new_version)
        
        def download_only():
            """Open alleen de download pagina in browser"""
            if download_url:
                webbrowser.open(download_url)
                dialog.destroy()
        
        tk.Button(btn_container, text="Download & Installeer", command=download_and_install,
                 bg=self.theme["SUCCESS_COLOR"], fg="white",
                 font=("Segoe UI", 10, "bold"), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
        
        tk.Button(btn_container, text="Alleen Download", command=download_only,
                 bg=self.theme["ACCENT_COLOR"], fg="white",
                 font=("Segoe UI", 10), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
        
        tk.Button(btn_container, text="Later", command=dialog.destroy,
                 bg=self.theme["BG_PRIMARY"], fg=self.theme["TEXT_PRIMARY"],
                 font=("Segoe UI", 10), padx=20, pady=10,
                 relief="flat", cursor="hand2").pack(side=tk.LEFT, padx=3)
    
    def download_and_install_update(self, download_url, version):
        """Download update en start installatie automatisch"""
        try:
            # Toon voortgang dialoog
            progress_dialog = tk.Toplevel(self.root)
            progress_dialog.title("Update Downloaden")
            progress_dialog.geometry("400x150")
            progress_dialog.configure(bg=self.theme["BG_PRIMARY"])
            progress_dialog.transient(self.root)
            progress_dialog.resizable(False, False)
            
            try:
                icon_path = get_resource_path('favicon.ico')
                if os.path.exists(icon_path):
                    progress_dialog.iconbitmap(icon_path)
            except:
                pass
            
            tk.Label(progress_dialog, text="Update downloaden...", 
                    font=("Segoe UI", 12, "bold"),
                    bg=self.theme["BG_PRIMARY"], 
                    fg=self.theme["TEXT_PRIMARY"]).pack(pady=20)
            
            status_label = tk.Label(progress_dialog, text="Bezig met downloaden...",
                                   font=("Segoe UI", 9),
                                   bg=self.theme["BG_PRIMARY"],
                                   fg=self.theme["TEXT_SECONDARY"])
            status_label.pack(pady=10)
            
            progress_dialog.update()
            
            # Download in achtergrond thread
            def download_thread():
                try:
                    # Download naar temp directory
                    temp_dir = tempfile.gettempdir()
                    filename = f"NVict_Reader_v{version}_Setup.exe"
                    filepath = os.path.join(temp_dir, filename)
                    
                    # Download bestand
                    urllib.request.urlretrieve(download_url, filepath)
                    
                    # Update UI in main thread
                    self.root.after(0, lambda: self._finish_download(progress_dialog, filepath))
                    
                except Exception as e:
                    self.root.after(0, lambda: self._download_error(progress_dialog, str(e)))
            
            thread = threading.Thread(target=download_thread, daemon=True)
            thread.start()
            
        except Exception as e:
            messagebox.showerror("Download Fout", 
                f"Kan update niet downloaden:\n{str(e)}\n\nProbeer handmatig te downloaden via de website.")
    
    def _finish_download(self, progress_dialog, filepath):
        """Voltooi download en start installer"""
        try:
            progress_dialog.destroy()
            
            if os.path.exists(filepath):
                # Vraag bevestiging om installer te starten
                if messagebox.askyesno("Update Downloaden Voltooid",
                    f"Update succesvol gedownload!\n\n"
                    f"Wilt u de installer nu starten?\n\n"
                    f"Let op: Sluit eerst NVict Reader af voordat u de installatie voltooit."):
                    
                    # Start installer
                    if platform.system() == "Windows":
                        os.startfile(filepath)
                    elif platform.system() == "Darwin":
                        subprocess.run(["open", filepath])
                    else:
                        subprocess.run(["xdg-open", filepath])
                    
                    # Sluit applicatie
                    self.root.after(1000, self.exit_application)
            else:
                messagebox.showerror("Fout", "Download bestand niet gevonden")
                
        except Exception as e:
            messagebox.showerror("Fout", f"Kan installer niet starten:\n{str(e)}")
    
    def _download_error(self, progress_dialog, error_msg):
        """Toon download fout"""
        progress_dialog.destroy()
        messagebox.showerror("Download Fout",
            f"Kan update niet downloaden:\n{error_msg}\n\n"
            f"Probeer handmatig te downloaden via de website.")

    # ====================================================================
    # THUMBNAIL-PANEEL
    # ====================================================================

    def toggle_thumbnail_panel(self):
        """Toon of verberg het thumbnail-paneel links van het notebook."""
        if self.thumbnail_visible:
            self.thumbnail_panel.grid_remove()
            self.thumbnail_visible = False
            self._set_toolbar_button_active("pages", False)
        else:
            self.thumbnail_panel.grid()
            self.thumbnail_visible = True
            self._set_toolbar_button_active("pages", True)
            self.update_thumbnails()

    def update_thumbnails(self):
        """Herrender thumbnails voor het actieve tabblad (in achtergrondthread)."""
        if not self.thumbnail_visible:
            return
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            self.thumbnail_canvas.delete("all")
            self.thumbnail_images = []
            return

        # Render in achtergrondthread om UI niet te blokkeren
        def _render():
            fitz_mod = get_fitz()
            Image, ImageTk, _, _ = get_PIL()
            thumb_w = 86
            items = []
            y = 8
            try:
                n = len(tab.pdf_document)
                for pn in range(n):
                    page = tab.pdf_document[pn]
                    rect = page.bound()
                    scale = thumb_w / rect.width if rect.width > 0 else 1.0
                    mat = fitz_mod.Matrix(scale, scale)
                    pix = page.get_pixmap(matrix=mat)
                    pil_img = Image.open(io.BytesIO(pix.tobytes("ppm")))
                    items.append((pn, pil_img, y))
                    y += pil_img.height + 4 + 14  # thumb + gap + label
            except Exception as e:
                print(f"Thumbnail render error: {e}")
                return
            self.root.after(0, lambda: self._draw_thumbnails(tab, items, y))

        threading.Thread(target=_render, daemon=True).start()

    def _draw_thumbnails(self, tab, items, total_y):
        """Teken thumbnails op het thumbnail-canvas (wordt aangeroepen vanuit main thread)."""
        Image, ImageTk, _, _ = get_PIL()
        self.thumbnail_canvas.delete("all")
        self.thumbnail_images = []
        cx = 57  # horizontaal midden in 114px breed paneel
        current_page = getattr(tab, 'current_page', 0)

        for (page_num, pil_img, y) in items:
            photo = ImageTk.PhotoImage(pil_img)
            self.thumbnail_images.append(photo)
            w, h = pil_img.size
            is_current = (page_num == current_page)
            border_color = self.theme["ACCENT_COLOR"] if is_current else self.theme["TEXT_SECONDARY"]
            bw = 2 if is_current else 1

            # Rand-rechthoek
            self.thumbnail_canvas.create_rectangle(
                cx - w // 2 - bw, y - bw,
                cx + w // 2 + bw, y + h + bw,
                outline=border_color, width=bw,
                tags=f"tb_border_{page_num}"
            )
            # Afbeelding
            self.thumbnail_canvas.create_image(
                cx, y, anchor="n", image=photo,
                tags=f"tb_img_{page_num}"
            )
            # Paginanummer
            self.thumbnail_canvas.create_text(
                cx, y + h + 2, anchor="n",
                text=str(page_num + 1),
                font=("Segoe UI", 7),
                fill=self.theme["TEXT_SECONDARY"],
                tags=f"tb_lbl_{page_num}"
            )
            # Klik-binding op alle elementen van dit thumbnail
            for tag in (f"tb_border_{page_num}", f"tb_img_{page_num}", f"tb_lbl_{page_num}"):
                self.thumbnail_canvas.tag_bind(
                    tag, "<Button-1>",
                    lambda _e, p=page_num, t=tab: self._thumbnail_click(p, t)
                )

        self.thumbnail_canvas.configure(scrollregion=(0, 0, 114, total_y + 10))

    def _thumbnail_click(self, page_num, tab):
        """Navigeer naar de geklickte pagina via het thumbnail-paneel."""
        current_tab = self.get_active_tab()
        if current_tab is not tab:
            return
        tab.current_page = page_num
        self.scroll_to_page(tab, page_num)
        self.update_ui_state()
        # Thumbnail-markering bijwerken zonder volledige herrender
        self.root.after(30, self.update_thumbnails)

    # ====================================================================
    # MARKEERMODUS
    # ====================================================================

    def toggle_highlight_mode(self):
        """Zet de markeermodus aan of uit."""
        self.highlight_mode = not self.highlight_mode
        if self.highlight_mode:
            # Deactiveer conflicterende modi
            tab = self.get_active_tab()
            if isinstance(tab, PDFTab) and tab.text_annotate_mode:
                tab.text_annotate_mode = False
                tab.canvas.unbind("<Button-1>")
                tab.canvas.bind("<Button-1>", lambda e, t=tab: self.on_click(e, t))
                tab.canvas.config(cursor="arrow")
                self._set_toolbar_button_active("type-text", False)
            if isinstance(tab, PDFTab) and tab.form_mode:
                tab.form_mode = False
                self._save_form_widget_values(tab)
                self._clear_form_overlays(tab)
                self.display_page(tab)
                self._set_toolbar_button_active("form", False)
            self._set_toolbar_button_active("marker", True)
            self.status_label.config(
                text="✏ Markeermodus aan — sleep over tekst om te markeren"
            )
        else:
            self._set_toolbar_button_active("marker", False)
            self.status_label.config(text="Gereed")

    def apply_highlight_annotation(self, tab, selected_words):
        """
        Voeg een gele highlight-annotatie toe aan het PDF-document voor de
        geselecteerde woorden. De annotatie wordt opgeslagen bij 'Opslaan'.
        """
        try:
            fitz_mod = get_fitz()
            page_regions = getattr(tab, 'page_regions', {})

            # Groepeer woorden per pagina
            words_by_page = {}
            for word_data in selected_words:
                text, wx0, wy0, wx1, wy1 = word_data
                found_page = None
                if page_regions:
                    for pn, (rx0, ry0, rx1, ry1) in page_regions.items():
                        if rx0 <= wx0 <= rx1 and ry0 <= wy0 <= ry1:
                            found_page = pn
                            break
                else:
                    for pn, page_y in enumerate(tab.page_positions):
                        if pn + 1 < len(tab.page_positions):
                            if page_y <= wy0 < tab.page_positions[pn + 1]:
                                found_page = pn
                                break
                        else:
                            if wy0 >= page_y:
                                found_page = pn
                                break
                if found_page is not None:
                    words_by_page.setdefault(found_page, []).append(word_data)

            total_words = 0
            for page_num, page_words in words_by_page.items():
                page = tab.pdf_document[page_num]
                # Haal canvas-offset op voor deze pagina
                if page_num in page_regions:
                    px_off, py_off = page_regions[page_num][0], page_regions[page_num][1]
                else:
                    px_off = tab.page_offset_x
                    py_off = tab.page_positions[page_num]

                quads = []
                for _text, wx0, wy0, wx1, wy1 in page_words:
                    # Canvas-coördinaten terug naar PDF-coördinaten
                    pdf_x0 = (wx0 - px_off) / tab.zoom_level
                    pdf_y0 = (wy0 - py_off) / tab.zoom_level
                    pdf_x1 = (wx1 - px_off) / tab.zoom_level
                    pdf_y1 = (wy1 - py_off) / tab.zoom_level
                    quads.append(
                        fitz_mod.Rect(pdf_x0, pdf_y0, pdf_x1, pdf_y1).quad
                    )
                    total_words += 1

                if quads:
                    annot = page.add_highlight_annot(quads)
                    annot.set_colors(stroke=(1.0, 0.85, 0.0))  # geel
                    annot.update()

            # Herrender om annotaties direct te tonen
            self.display_page(tab)
            self.status_label.config(
                text=f"✅ Markering toegevoegd ({total_words} woorden). Sla op met Ctrl+S."
            )
        except Exception as e:
            self.status_label.config(text=f"Markering mislukt: {e}")

    # ====================================================================
    # BOEK-MODUS
    # ====================================================================

    def toggle_book_mode(self):
        """Schakel boek-modus (twee pagina's naast elkaar) in of uit."""
        tab = self.get_active_tab()
        if not isinstance(tab, PDFTab):
            return
        tab.book_mode = not getattr(tab, 'book_mode', False)
        if tab.book_mode:
            self._set_toolbar_button_active("book", True)
            self.status_label.config(text="Boek-modus aan")
        else:
            self._set_toolbar_button_active("book", False)
            self.status_label.config(text="Boek-modus uit")
        self.display_page(tab)

    # ====================================================================

    def run(self):
        self.notebook.add(self.welcome_frame)
        # WM_DELETE_WINDOW wordt nu in main() ingesteld met single instance cleanup
        
        # Start automatische update check in achtergrond (na 2 seconden)
        self.root.after(2000, self.check_for_updates_on_startup)
        
        self.root.mainloop()

def main():
    try:
        # Controleer voor single instance
        single_instance = SingleInstance()
        
        # Check of er een bestand is meegegeven als argument
        file_to_open = None
        print_mode = False
        
        # Parse command line argumenten
        if len(sys.argv) > 1:
            if sys.argv[1] == "--print" and len(sys.argv) > 2:
                # Format: NVictReader.exe --print "bestand.pdf"
                print_mode = True
                if os.path.exists(sys.argv[2]):
                    file_to_open = os.path.abspath(sys.argv[2])
            elif os.path.exists(sys.argv[1]):
                # Format: NVictReader.exe "bestand.pdf"
                file_to_open = os.path.abspath(sys.argv[1])
        
        # Check of er al een instance draait
        if single_instance.is_already_running():
            # Stuur bestand naar bestaande instance als er een is
            if file_to_open:
                if single_instance.send_to_existing_instance(file_to_open):
                    print(f"Bestand verzonden naar bestaande instance: {file_to_open}")
                    return  # Sluit deze instance af
                else:
                    print("Kon niet communiceren met bestaande instance, start nieuwe instance")
            else:
                # Geen bestand om te openen, gewoon een nieuwe instance starten
                # (gebruiker wil misschien een tweede venster)
                pass
        
        # Start de applicatie
        app = NVictReader()
        
        # Start single instance server
        single_instance.start_server(app)
        
        # Open bestand als er een is meegegeven
        if file_to_open:
            app.root.after(100, lambda: app.add_new_tab(file_to_open))
            
            # Als print mode, open automatisch het print dialoog
            if print_mode:
                app.root.after(500, lambda: app.print_pdf())
        
        # Zorg dat server wordt gestopt bij afsluiten
        def on_closing():
            single_instance.stop()
            app.exit_application()
        
        app.root.protocol("WM_DELETE_WINDOW", on_closing)
        
        app.run()
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        input("\nDruk op Enter om af te sluiten...")

if __name__ == "__main__":
    main()
