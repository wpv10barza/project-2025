# -*- coding: utf-8 -*-
from __future__ import annotations

import pandas as pd
from pathlib import Path

SQL_PREGUNTAS = """
SELECT
  ID_PREGUNTA,
  PREGUNTA_CODIGO,
  PREGUNTA_DESCRIPCION,
  PREGUNTA_SITUACION_DESEADA,
  EvaluacionId
FROM dbo.VW_PREGUNTA_POR_EVALUACION
WHERE EvaluacionId = ?
ORDER BY PREGUNTA_CODIGO;
"""

def fetch_questions(conn, evaluacion_id: int) -> pd.DataFrame:
    # Nota: pandas muestra un warning con pyodbc; es normal y NO afecta el resultado.
    return pd.read_sql(SQL_PREGUNTAS, conn, params=[evaluacion_id])
