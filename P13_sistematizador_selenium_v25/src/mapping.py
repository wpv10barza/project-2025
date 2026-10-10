from __future__ import annotations
import pandas as pd
from typing import Any, Dict

def _get(row: pd.Series, *candidates: str) -> str:
    for c in candidates:
        if c in row and row[c] is not None:
            v = str(row[c]).strip()
            if v and v.lower() != "nan":
                return v
    return ""

def normalize_newlines_to_literal(s: str) -> str:
    """
    Convierte saltos reales (\r\n o \n) a secuencia literal \\n para que en SQL
    se vea como texto "\n" (y no corte la fila / no rompa el MERGE).
    Si el texto ya trae \\n literal, se mantiene.
    """
    if not s:
        return ""
    s = s.replace("\r\n", "\n").replace("\r", "\n")
    # solo convierte saltos reales:
    return s.replace("\n", r"\n")

def map_row_to_respuesta(row: pd.Series, pregunta_id: int, aplicacion_id: int) -> Dict[str, Any]:
    """
    Mapea una fila del Excel a las columnas REALES de dbo.Respuesta.
    Pedido:
      - SituacionActual  <= Columna F (ANÁLISIS_DE_COBERTURA)
      - Notas            <= Columna B (CONVERSACION_ENTREVISTADOR_1)
      - RespuestaLiteral <= Columna C (CONVERSACION_INFORMANTE_1)
    Además:
      - PREGUNTA_CODIGO se usa para encontrar PreguntaId (fuera de esta función)
    """
    notas = _get(row, "CONVERSACION_ENTREVISTADOR_1", "Notas", "NOTAS")
    resp_lit = _get(row, "CONVERSACION_INFORMANTE_1", "RespuestaLiteral", "RESPUESTA_LITERAL", "RESPUESTA")
    situacion = _get(row, "ANÁLISIS_DE_COBERTURA", "ANALISIS_DE_COBERTURA", "SituacionActual", "SITUACIONACTUAL")

    # Reemplazo recomendado (a los 3)
    notas = normalize_newlines_to_literal(notas)
    resp_lit = normalize_newlines_to_literal(resp_lit)
    situacion = normalize_newlines_to_literal(situacion)

    return {
        "AplicacionId": aplicacion_id,
        "PreguntaId": pregunta_id,
        "SituacionActual": situacion,
        "Notas": notas,
        "RespuestaLiteral": resp_lit,
    }
