# Build & Release — NVict Reader

NVict Reader heeft **twee** build-paden:

1. **Nieuw (aanbevolen):** `python build.py` — gebruikt de gedeelde [`_nvict_build`](../_nvict_build) toolkit
2. **Oud (backup):** `Release_Complete_v3_0.bat` + `Upload_FTP.ps1` — blijft werken zolang je `ftp_config.ini` hebt

De nieuwe flow is consistent met `NVict-CDtoSpotify` en andere NVict apps, en zou geleidelijk de batch-scripts moeten vervangen.

## Nieuwe flow

```bash
# Volledige release (EXE + installer + FTPS upload)
python build.py

# Alleen lokaal bouwen
python build.py --no-upload

# Zonder signing (dev)
python build.py --no-sign --no-upload
```

### Wat er gebeurt

1. **clean** — verwijder `dist/`, `build/`, `Output/`
2. **version-info** — schrijf `version_info.txt` uit `version.py`
3. **exe** — PyInstaller → `dist/NVict_Reader.exe`
4. **sign-exe** — signtool + DigiCert timestamp
5. **installer** — Inno Setup → `Output/NVict_Reader_Setup.exe` (ook gesigneerd)
6. **version-json** — schrijf `dist/updates/nvict_reader_version.json`
7. **upload** — FTPS upload (3 bestanden)

### FTP configuratie

Credentials staan centraal in `../_nvict_build/.env.build` (gedeeld met andere NVict apps):

```ini
CERT_PATH=C:\Users\NVict Service\certs\codesign.pfx
CERT_PASSWORD=...
FTP_HOST=ftp.nvict.nl
FTP_USER=softwareupload@nvict.nl
FTP_PASSWORD=...
FTP_TLS=true
```

De doelmappen staan in `app_meta.yaml`:

```yaml
ftp_subdir: /NVict_Reader
# ftp_updates_subdir wordt afgeleid als /NVict_Reader/updates
```

## Versie bijwerken

`APP_VERSION` staat op **twee** plekken die in sync moeten blijven:

- `NVict_Reader.py` regel ~33: `APP_VERSION = "..."` (gebruikt door de app intern)
- `version.py`: `APP_VERSION = "..."` (gebruikt door `_nvict_build`)

Werk beide bij, werk `release_notes.txt` bij, en run `python build.py`.

## Oude batch-flow (fallback)

`Release_Complete_v3_0.bat` blijft werken voor de PS1-gebaseerde release. Gebruik deze alleen als de Python-flow faalt. De oude `Create_Version_JSON.ps1` / `Upload_FTP.ps1` / `Update_Version_Info.ps1` scripts blijven in de repo staan als backup.
