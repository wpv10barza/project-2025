from __future__ import annotations
import json
from pathlib import Path
from typing import Any, Dict

import pyodbc

# Nota: mantenemos un driver por defecto, pero se puede sobreescribir en db_config.json
DEFAULT_DRIVER = "ODBC Driver 18 for SQL Server"

def read_db_config(path: str | Path) -> Dict[str, Any]:
    """Lee db_config.json soportando UTF-8 con BOM (utf-8-sig)."""
    p = Path(path).expanduser().resolve()
    if not p.exists():
        raise FileNotFoundError(f"No existe db_config.json en: {p}")
    raw = p.read_text(encoding="utf-8-sig")  # soporta BOM
    data = json.loads(raw)
    required = ["server", "database", "username", "password"]
    missing = [k for k in required if not data.get(k)]
    if missing:
        raise ValueError(f"db_config.json incompleto. Faltan: {', '.join(missing)}")
    if not data.get("driver"):
        data["driver"] = DEFAULT_DRIVER
    return data

def connect(cfg: Dict[str, Any]) -> pyodbc.Connection:
    """Crea conexión pyodbc a SQL Server Azure."""
    driver = cfg.get("driver") or DEFAULT_DRIVER
    server = cfg["server"]
    database = cfg["database"]
    username = cfg["username"]
    password = cfg["password"]
    encrypt = cfg.get("encrypt", "yes")
    trust = cfg.get("trust_server_certificate", "no")
    timeout = int(cfg.get("timeout", 30))

    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={username};"
        f"PWD={password};"
        f"Encrypt={encrypt};"
        f"TrustServerCertificate={trust};"
        f"Connection Timeout={timeout};"
    )
    return pyodbc.connect(conn_str)

def sql_literal_newlines_to_db(value: str) -> str:
    """Convierte \n literal (dos caracteres) a salto real, antes de MERGE/INSERT."""
    return value.replace("\\r\\n", "\\n").replace("\\n", "\n")

