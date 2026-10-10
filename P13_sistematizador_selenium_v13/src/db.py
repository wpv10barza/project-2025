# -*- coding: utf-8 -*-
from __future__ import annotations

import json
from pathlib import Path
import pyodbc

from config import DRIVER if False else None  # placeholder for type-checkers

ODBC_DRIVER = "{ODBC Driver 18 for SQL Server}"

def read_db_config(path: Path) -> dict:
    p = Path(path)
    if p.exists() and p.is_dir():
        p = p / "db_config.json"
    if not p.exists():
        raise FileNotFoundError(f"No existe db_config.json en: {p}")
    cfg = json.loads(p.read_text(encoding="utf-8-sig"))
    for k in ("server","database","username","password"):
        if k not in cfg or not str(cfg[k]).strip():
            raise ValueError(f"Falta '{k}' en {p}")
    return cfg

def connect(cfg: dict) -> pyodbc.Connection:
    conn_str = (
        f"DRIVER={ODBC_DRIVER};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
        "Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str)
