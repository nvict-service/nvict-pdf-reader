# -*- coding: utf-8 -*-
"""Resource-pad-resolutie die ook in een PyInstaller-bundel werkt.

Poort van get_resource_path uit NVict_Reader.py (regel 465-473): in een
gebundelde build wijst sys._MEIPASS naar de map met meegepakte data
(icons/, favicon.ico); tijdens development is dat de projectroot (één
niveau boven dit package).
"""

import os
import sys

_PROJECT_ROOT = os.path.abspath(os.path.join(os.path.dirname(__file__), ".."))


def get_resource_path(relative_path: str) -> str:
    base_path = getattr(sys, "_MEIPASS", _PROJECT_ROOT)
    return os.path.join(base_path, relative_path)
