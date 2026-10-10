from __future__ import annotations
import json
from pathlib import Path
from typing import Dict

import pyodbc
import pandas as pd

SQL_PREGUNTAS = """
SELECT ID_PREGUNTA, PREGUNTA_CODIGO, PREGUNTA_DESCRIPCION, PREGUNTA_SITUACION_DESEADA
FROM dbo.VW_PREGUNTA_POR_EVALUACION
WHERE EvaluacionId = ?
ORDER BY PREGUNTA_CODIGO;
"""

def read_db_config(config_path: Path) -> Dict[str, str]:
    p = Path(config_path)
    if p.exists() and p.is_dir():
        p = p / "db_config.json"
    if not p.exists():
        raise FileNotFoundError(
            f"No existe {p}. Crea db_config.json al lado de config.py.\n"
            'Ejemplo: { "server":"...", "database":"...", "username":"...", "password":"..." }'
        )
    cfg = json.loads(p.read_text(encoding="utf-8-sig"))
    for k in ["server", "database", "username", "password"]:
        if k not in cfg or not str(cfg[k]).strip():
            raise ValueError(f"Falta '{k}' en {p}.")
    return cfg

def get_connection(db_config_file: Path, driver: str) -> pyodbc.Connection:
    cfg = read_db_config(db_config_file)
    conn_str = (
        f"DRIVER={driver};"
        f"SERVER={cfg['server']};"
        f"DATABASE={cfg['database']};"
        f"UID={cfg['username']};"
        f"PWD={cfg['password']};"
        "Encrypt=yes;"
        "TrustServerCertificate=no;"
        "Connection Timeout=30;"
    )
    return pyodbc.connect(conn_str)

def load_df_preguntas(conn: pyodbc.Connection, evaluacion_id: int) -> pd.DataFrame:
    df = pd.read_sql(SQL_PREGUNTAS, conn, params=[evaluacion_id])
    df["PREGUNTA_CODIGO"] = df["PREGUNTA_CODIGO"].astype(str).str.strip()
    df["PREGUNTA_DESCRIPCION"] = df["PREGUNTA_DESCRIPCION"].fillna("").astype(str)
    df["PREGUNTA_SITUACION_DESEADA"] = df["PREGUNTA_SITUACION_DESEADA"].fillna("").astype(str)
    return df
