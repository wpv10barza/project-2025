\
# -*- coding: utf-8 -*-
from pathlib import Path
import os

REPO_DIR = Path(__file__).resolve().parent

DEFAULT_DATA_PILOTO_DIR = Path(os.environ.get("DATA_PILOTO_DIR", str(REPO_DIR.parent)))

DEFAULT_OBJETIVO_TXT      = DEFAULT_DATA_PILOTO_DIR / "OBJETIVO_FINAL.txt"
DEFAULT_TRANSCRIPCION_TXT = DEFAULT_DATA_PILOTO_DIR / "Transcripcion.txt"
DEFAULT_TRANSCRIPCION_PDF = DEFAULT_DATA_PILOTO_DIR / "Transcripcion.pdf"

PROMP_DIR  = DEFAULT_DATA_PILOTO_DIR / "01.02 Promp"
RESULT_DIR = DEFAULT_DATA_PILOTO_DIR / "01.03 Resultado ChatGPT"
PROMP_DIR.mkdir(parents=True, exist_ok=True)
RESULT_DIR.mkdir(parents=True, exist_ok=True)

PROMPT_TXT_OUT  = PROMP_DIR / "prompt_completo.txt"
PROMPT_META_OUT = PROMP_DIR / "prompt_metadata.json"

RESULT_RAW_TXT = RESULT_DIR / "resultado_raw.txt"
RESULT_TSV     = RESULT_DIR / "resultado.tsv"
RESULT_CSV     = RESULT_DIR / "resultado.csv"
RESULT_XLSX    = RESULT_DIR / "resultado.xlsx"

DEFAULT_PROMPT_TRABAJO = REPO_DIR / "prompt_trabajo.txt"

CHATGPT_URL = os.environ.get("CHATGPT_URL", "https://chatgpt.com/")

EDGE_USER_DATA_DIR = os.environ.get("EDGE_USER_DATA_DIR", "")
EDGE_PROFILE_DIR   = os.environ.get("EDGE_PROFILE_DIR", "Default")
EDGE_BINARY = os.environ.get("EDGE_BINARY", r"C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe")

DRIVER = "{ODBC Driver 18 for SQL Server}"
DEFAULT_DB_CONFIG = REPO_DIR / "db_config.json"
