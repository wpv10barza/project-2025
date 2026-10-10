from __future__ import annotations
import os
import sys
from pathlib import Path

from config import (
    DATA_PILOTO_DIR, OBJETIVO_TXT, TRANSCRIPCION_TXT, TRANSCRIPCION_PDF,
    DB_CONFIG_FILE, EVALUACION_ID, OPENAI_MODEL
)

def _exists(p: Path) -> str:
    return "OK" if p.exists() else "MISSING"

def run_diagnostics() -> bool:
    ok = True
    print("=== DIAGNOSTICO P13 ===")
    print("Python:", sys.version.replace("\n", " "))
    print("Executable:", sys.executable)
    print("DATA_PILOTO_DIR:", DATA_PILOTO_DIR)
    print("OBJETIVO_FINAL.txt:", _exists(OBJETIVO_TXT), OBJETIVO_TXT)
    print("Transcripcion.pdf:", _exists(TRANSCRIPCION_PDF), TRANSCRIPCION_PDF)
    print("db_config.json:", _exists(DB_CONFIG_FILE), DB_CONFIG_FILE)
    print("EVALUACION_ID:", EVALUACION_ID)
    print("OPENAI_MODEL:", OPENAI_MODEL)
    print("OPENAI_API_KEY:", "SET" if os.getenv("OPENAI_API_KEY") else "NOT SET")
    print("OBJETIVO_FINAL.txt:", _exists(OBJETIVO_TXT))
    print("Transcripcion.txt:", _exists(TRANSCRIPCION_TXT))
    print("Transcripcion.pdf:", _exists(TRANSCRIPCION_PDF))

    if not OBJETIVO_TXT.exists():
        ok = False
    if not (TRANSCRIPCION_TXT.exists() or TRANSCRIPCION_PDF.exists()):
        ok = False
    if not DB_CONFIG_FILE.exists():
        ok = False

    try:
        import openai  # noqa
        print("openai import:", "OK")
    except Exception as e:
        print("openai import:", f"FAIL ({e})")

    try:
        import pyodbc  # noqa
        print("pyodbc import:", "OK")
    except Exception as e:
        print("pyodbc import:", f"FAIL ({e})")
        ok = False

    print("=======================")
    return ok
