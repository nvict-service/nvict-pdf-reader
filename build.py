#!/usr/bin/env python3
"""
Thin build-wrapper voor NVict Reader.

Delegeert aan ../_nvict_build/release.py. Alle CLI-flags worden doorgegeven.

Voorbeelden:
    python build.py                  # complete release (build + sign + upload)
    python build.py --no-upload      # alleen lokaal bouwen
    python build.py --no-sign        # zonder code-signing (dev)
    python build.py --only clean,exe # alleen specifieke stappen
"""
from __future__ import annotations

import subprocess
import sys
from pathlib import Path


def _load_app_id(app_meta: Path) -> str:
    for line in app_meta.read_text(encoding="utf-8").splitlines():
        s = line.strip()
        if s.startswith("app_id:"):
            v = s.split(":", 1)[1].strip()
            return v.strip("\"'")
    raise RuntimeError("app_id niet gevonden in app_meta.yaml")


def main() -> int:
    here = Path(__file__).resolve().parent
    tools = here.parent / "_nvict_build"
    release = tools / "release.py"
    meta = here / "app_meta.yaml"

    if not release.exists():
        sys.stderr.write(
            f"[!] _nvict_build niet gevonden op {tools}\n"
            "    Zorg dat deze map naast de app-map staat "
            "(C:\\...\\NVictPython\\_nvict_build).\n"
        )
        return 1
    if not meta.exists():
        sys.stderr.write(f"[!] app_meta.yaml niet gevonden in {here}\n")
        return 1

    app_id = _load_app_id(meta)
    cmd = [sys.executable, str(release), app_id, *sys.argv[1:]]
    return subprocess.call(cmd, cwd=str(tools))


if __name__ == "__main__":
    sys.exit(main())
