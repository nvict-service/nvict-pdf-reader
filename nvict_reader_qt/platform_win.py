# -*- coding: utf-8 -*-
"""Windows-registry-integratie voor "als standaard PDF-viewer instellen".

Overgenomen uit NVict_Reader.py (class DefaultPDFHandler, regel 122+), maar
alleen de tkinter-onafhankelijke statische methodes (registry-logica). De
UI-methodes (prompt_set_as_default, show_first_run_dialog) gebruikten
tkinter.messagebox/Toplevel en horen bij een latere fase, waarin ze als
Qt-dialoog opnieuw gebouwd worden. Deze module wordt in fase 1 nog niet
aangeroepen (geen first-run-dialoog), maar staat al klaar zodat een latere
fase zonder refactor kan aanhaken.
"""

import os
import subprocess
import sys

try:
    import winreg
except ImportError:
    winreg = None


class DefaultPDFHandler:
    """Handelt de registratie van NVict Reader als standaard PDF-viewer af."""

    @staticmethod
    def is_default_pdf_handler():
        """Check of NVict Reader momenteel de standaard PDF-handler is."""
        try:
            if getattr(sys, 'frozen', False):
                current_exe = sys.executable
            else:
                current_exe = os.path.abspath(sys.argv[0])

            current_exe_lower = current_exe.lower()
            exe_name = os.path.basename(current_exe_lower)

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

                if prog_id == "NVictReader.PDF":
                    return True

                if "nvict" in prog_id_lower or "nvictreader" in prog_id_lower:
                    return True

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
        """Open Windows-instellingen direct bij de .pdf-koppeling."""
        try:
            subprocess.run(['start', 'ms-settings:defaultapps'], shell=True)
            return True
        except Exception:
            return False

    @staticmethod
    def register_open_with():
        """Registreer NVict Reader in het register.

        Als Inno Setup het al in HKLM heeft gezet, doen we hier niets om
        dubbele items te voorkomen.
        """
        try:
            try:
                key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, "Software\\Classes\\NVictReader.PDF", 0, winreg.KEY_READ)
                winreg.CloseKey(key)
                return True
            except OSError:
                pass

            if getattr(sys, 'frozen', False):
                exe_path = sys.executable
            else:
                exe_path = os.path.abspath(sys.argv[0])

            prog_id = "NVictReader.PDF"
            hkcu = winreg.HKEY_CURRENT_USER

            cap_path = r"Software\NVict Service\NVict Reader\Capabilities"

            key = winreg.CreateKey(hkcu, cap_path)
            winreg.SetValueEx(key, "ApplicationName", 0, winreg.REG_SZ, "NVict Reader")
            winreg.SetValueEx(key, "ApplicationDescription", 0, winreg.REG_SZ, "NVict Reader PDF Viewer")
            winreg.CloseKey(key)

            key = winreg.CreateKey(hkcu, f"{cap_path}\\FileAssociations")
            winreg.SetValueEx(key, ".pdf", 0, winreg.REG_SZ, prog_id)
            winreg.CloseKey(key)

            key = winreg.CreateKey(hkcu, "Software\\RegisteredApplications")
            winreg.SetValueEx(key, "NVictReader", 0, winreg.REG_SZ, cap_path)
            winreg.CloseKey(key)

            classes_path = f"Software\\Classes\\{prog_id}"

            key = winreg.CreateKey(hkcu, classes_path)
            winreg.SetValue(key, "", winreg.REG_SZ, "NVict Reader PDF")
            winreg.CloseKey(key)

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

            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print")
            winreg.SetValue(key, "", winreg.REG_SZ, "Afdrukken")
            winreg.CloseKey(key)

            key = winreg.CreateKey(hkcu, f"{classes_path}\\shell\\print\\command")
            winreg.SetValue(key, "", winreg.REG_SZ, f'"{exe_path}" --print "%1"')
            winreg.CloseKey(key)

            key = winreg.CreateKey(hkcu, "Software\\Classes\\.pdf\\OpenWithProgids")
            winreg.SetValueEx(key, prog_id, 0, winreg.REG_NONE, b'')
            winreg.CloseKey(key)

            try:
                import ctypes
                ctypes.windll.shell32.SHChangeNotify(0x08000000, 0, 0, 0)
            except Exception:
                pass

            return True
        except Exception as e:
            print(f"Error registering: {e}")
            return False
