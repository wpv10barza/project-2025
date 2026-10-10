# -*- coding: utf-8 -*-
from __future__ import annotations
from pathlib import Path
import pandas as pd

from src.db import read_db_config, connect
from src.sql import SQL_APLICACION

def get_aplicacion_id(conn, evaluacion_id: int, informante_id: int) -> int:
    row = pd.read_sql(SQL_APLICACION, conn, params=[int(evaluacion_id), int(informante_id)])
    if row.empty:
        raise RuntimeError(f"No existe Aplicacion para EvaluacionId={evaluacion_id} e InformanteId={informante_id}")
    return int(row.iloc[0]["Id"])

def load_stub(db_config_path: Path, evaluacion_id: int, informante_id: int) -> int:
    cfg = read_db_config(db_config_path)
    with connect(cfg) as conn:
        return get_aplicacion_id(conn, evaluacion_id, informante_id)
