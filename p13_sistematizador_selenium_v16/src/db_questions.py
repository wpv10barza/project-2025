# -*- coding: utf-8 -*-
import pandas as pd
from src.sql import SQL_PREGUNTAS

def load_questions_df(conn, evaluacion_id: int) -> pd.DataFrame:
    df = pd.read_sql(SQL_PREGUNTAS, conn, params=[int(evaluacion_id)])
    df["PREGUNTA_CODIGO"] = df["PREGUNTA_CODIGO"].astype(str).str.strip()
    return df
