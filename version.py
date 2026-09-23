"""Central version information voor de release-pijplijn (_nvict_build).

Vanaf 2026-09-23 volgt dit bestand de PySide6 (Qt)-release-lijn (zie
app_meta.yaml) - loopt daardoor bewust NIET meer gelijk met de hardcoded
APP_VERSION in NVict_Reader.py (tkinter), die op 2.5 blijft staan omdat dat
bestand ongewijzigd blijft.
"""

APP_NAME = "NVict Reader"
APP_VERSION = "3.0"
APP_PUBLISHER = "NVict Service"
APP_URL = "https://www.nvict.nl/software/NVict_Reader/"
APP_ID = "NVictReader"
APP_COPYRIGHT = "Copyright 2026 NVict Service"

# Update-check endpoints
DOWNLOAD_URL = "https://www.nvict.nl/software/NVict_Reader/NVict_Reader_Setup.exe"
UPDATE_CHECK_URL = "https://www.nvict.nl/software/updates/nvict_reader_version.json"
