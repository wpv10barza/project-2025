from __future__ import annotations
from pathlib import Path

# Este código vive en: ...\01.01 Código de sistematizador\[carpeta del zip]
CODE_DIR = Path(__file__).resolve().parent

def _find_data_piloto_dir() -> Path:
    """
    Encuentra automáticamente la carpeta DATA PILOTO aunque el código esté
    en subcarpetas.

    Busca Transcripcion.pdf + OBJETIVO_FINAL.txt en:
    1) Ancestros directos del CODE_DIR
    2) Carpetas hermanas con "DATA PILOTO" en el nombre
    3) Fallback: CODE_DIR.parent
    """
    required = ["Transcripcion.pdf", "OBJETIVO_FINAL.txt"]

    p = CODE_DIR
    for _ in range(8):
        if all((p / f).exists() for f in required):
            return p
        p = p.parent

    p = CODE_DIR
    for _ in range(8):
        parent = p.parent
        try:
            for child in parent.iterdir():
                if child.is_dir() and "DATA PILOTO" in child.name.upper():
                    if all((child / f).exists() for f in required):
                        return child
        except Exception:
            pass
        p = parent

    return CODE_DIR.parent

DATA_PILOTO_DIR = _find_data_piloto_dir()

# =========================
# INSUMOS
# =========================
TRANSCRIPCION_TXT = DATA_PILOTO_DIR / "Transcripcion.txt"
TRANSCRIPCION_PDF = DATA_PILOTO_DIR / "Transcripcion.pdf"  # fallback
OBJETIVO_TXT = DATA_PILOTO_DIR / "OBJETIVO_FINAL.txt"

# (Opcional) si algún piloto aún usa guía PDF
GUIA_PDF = DATA_PILOTO_DIR / "PREGUNTA_CODIGO.pdf"

# =========================
# SALIDAS
# =========================
PROMP_DIR = DATA_PILOTO_DIR / "01.02 Promp"
# Copia adicional del C8 (si existe la unidad/ruta)
PROMP_EXTRA_DIR = Path(r"G:\Mi unidad\PROMP")
RESULT_DIR = DATA_PILOTO_DIR / "01.03 Resultado ChatGPT"
PROMP_DIR.mkdir(parents=True, exist_ok=True)
RESULT_DIR.mkdir(parents=True, exist_ok=True)

PROMPT_TXT_OUT = PROMP_DIR / "prompt_completo.txt"
PROMPT_META_OUT = PROMP_DIR / "prompt_metadata.json"

RESULT_RAW_TXT = RESULT_DIR / "resultado_raw.txt"
RESULT_TSV = RESULT_DIR / "resultado.tsv"
RESULT_CSV = RESULT_DIR / "resultado.csv"
RESULT_XLSX = RESULT_DIR / "resultado.xlsx"

# =========================
# OPENAI
# =========================
import os
OPENAI_MODEL = os.getenv("OPENAI_MODEL", "gpt-5.2")
TEMPERATURE = float(os.getenv("OPENAI_TEMPERATURE", "0.2"))

# =========================
# PILOTO P13: preguntas desde SQL (sin PDF guía)
# =========================
USE_SQL_PREGUNTAS = True
NO_DOCUMENTACION_PDF_INCLUIDA = True

# Parámetro del piloto
EVALUACION_ID = int(os.getenv("EVALUACION_ID", "17"))

# DB config local (NO subir a repos)
DB_CONFIG_FILE = CODE_DIR / "db_config.json"
SQL_DRIVER = "{ODBC Driver 18 for SQL Server}"

# Opcional: fijar roles acá (si lo dejas vacío, no se agrega nada extra)
DECLARACION_PARTICIPANTES = ""
