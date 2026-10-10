from __future__ import annotations
from pathlib import Path

REPO_DIR = Path(__file__).resolve().parent

# Si no pasas --data-piloto-dir, se usa este default.
# Recomendación: SIEMPRE usa --data-piloto-dir.
DEFAULT_DATA_PILOTO_DIR = REPO_DIR.parent / "03.00 DATA PILOTO"

# Archivos esperados en DATA PILOTO
OBJETIVO_FINAL = "OBJETIVO_FINAL.txt"
TRANSCRIPCION_TXT = "Transcripcion.txt"
TRANSCRIPCION_PDF = "Transcripcion.pdf"
TRANSCRIPCION_DOCX = "Transcripcion.docx"

PROMP_SUBDIR = "01.02 Promp"
RESULT_SUBDIR = "01.03 Resultado ChatGPT"

PROMPT_TXT_NAME = "prompt_completo.txt"
PROMPT_META_NAME = "prompt_metadata.json"

# Copia adicional del prompt para flujo Selenium
PROMPT_TRABAJO_NAME = "prompt_trabajo.txt"
PROMPT_TRABAJO_DIRNAME = "Sistematizador"  # DATA_PILOTO/01.02 Promp/Sistematizador/prompt_trabajo.txt

# DB / ODBC
ODBC_DRIVERS = [
    "ODBC Driver 18 for SQL Server",
    "ODBC Driver 17 for SQL Server",
]

# Selenium / Edge
EDGE_USER_DATA_DIR = r"C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data"
EDGE_PROFILE_DIR = "Default"
CHATGPT_URL = "https://chatgpt.com/"

DEFAULT_WAIT_SECONDS = 420
