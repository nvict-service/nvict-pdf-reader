# -*- coding: utf-8 -*-
"""Entrypoint voor de PyInstaller-bundel van de PySide6-editie.

Los top-level script (buiten het nvict_reader_qt-package) omdat
PyInstaller's Analysis geen relatieve imports kan volgen als het script
zelf binnen het package staat.
"""

import sys

from nvict_reader_qt.app import main

if __name__ == "__main__":
    sys.exit(main())
