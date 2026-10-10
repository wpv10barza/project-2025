from pathlib import Path
import pandas as pd
from .db_questions import load_questions
from .excel_reader import read_excel_matrix
from .mapping import map_row_to_sql
from .db import connect, read_db_config
from .sql import SQL_APLICACION

def get_aplicacion_id(evaluacion_id: int, informante_id: int, db_config_path: Path) -> int:
    cfg = read_db_config(db_config_path)
    conn = connect(cfg)
    try:
        cur = conn.cursor()
        cur.execute(SQL_APLICACION, (evaluacion_id, informante_id))
        row = cur.fetchone()
        if not row:
            raise RuntimeError(f"No se encontró Aplicacion para EvaluacionId={evaluacion_id} InformanteId={informante_id}")
        return int(row[0])
    finally:
        conn.close()
