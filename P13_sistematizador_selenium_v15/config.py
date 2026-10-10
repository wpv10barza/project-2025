from __future__ import annotations

from pathlib import Path

# Este proyecto se descomprime en una carpeta (ej: ...\03.02 PRUEBA CON REGISTROS\p13_sistematizador_...).
# DATA_PILOTO_DIR por defecto es la carpeta PADRE del código.
#
# ✅ Ajuste fijo solicitado:
# - Si existe la ruta estándar del proyecto (03.00 DATA PILOTO), úsala como insumo.
# - Si no existe (otro equipo), cae al comportamiento anterior (padre del código).
CODE_DIR = Path(__file__).resolve().parent
DEFAULT_DATA_PILOTO_DIR = Path(
    r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.00 DATA PILOTO"
)
DATA_PILOTO_DIR = DEFAULT_DATA_PILOTO_DIR if DEFAULT_DATA_PILOTO_DIR.exists() else CODE_DIR.parent

# -------------------------
# INSUMOS (defaults)
# -------------------------
OBJETIVO_TXT = DATA_PILOTO_DIR / "OBJETIVO_FINAL.txt"

# Prioriza TXT (lo solicitado). Si no existe, cae a PDF.
TRANSCRIPCION_TXT = DATA_PILOTO_DIR / "Transcripcion.txt"
TRANSCRIPCION_PDF = DATA_PILOTO_DIR / "Transcripcion.pdf"
TRANSCRIPCIONES_DIR = DATA_PILOTO_DIR / "Transcripciones"  # opcional (varios .txt/.pdf)

# -------------------------
# SALIDAS (defaults)
# -------------------------
PROMP_DIR = DATA_PILOTO_DIR / "01.02 Promp"
RESULT_DIR = DATA_PILOTO_DIR / "01.03 Resultado ChatGPT"
PROMP_DIR.mkdir(parents=True, exist_ok=True)
RESULT_DIR.mkdir(parents=True, exist_ok=True)

# Salida "normal" del C8
PROMPT_TXT_OUT = PROMP_DIR / "prompt_completo.txt"
PROMPT_META_OUT = PROMP_DIR / "prompt_metadata.json"

# ✅ Ajuste fijo solicitado: además, escribir una copia en:
# ...\01.02 Promp\Sistematizador\prompt_trabajo.txt
PROMPT_TRABAJO_DIR = Path(
    r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.02 PRUEBA CON REGISTROS\01.02 Promp\Sistematizador"
)
PROMPT_TRABAJO_DIR.mkdir(parents=True, exist_ok=True)
PROMPT_TRABAJO_OUT = PROMPT_TRABAJO_DIR / "prompt_trabajo.txt"
PROMPT_TRABAJO_META_OUT = PROMPT_TRABAJO_DIR / "prompt_trabajo_metadata.json"

# Copia adicional del prompt para Drive/compartidos (se copia si existe la ruta)
PROMP_COPY_DIR = Path(r"G:\Mi unidad\PROMP")

# Resultado capturado desde Sheets (o export manual)
RESULT_TSV = RESULT_DIR / "resultado.tsv"
RESULT_CSV = RESULT_DIR / "resultado.csv"
RESULT_XLSX = RESULT_DIR / "resultado.xlsx"
RESULT_RAW_TXT = RESULT_DIR / "resultado_raw.txt"

# OpenAI (ya no se usa en run; run ahora está pensado para Selenium)
OPENAI_MODEL = "gpt-5"
TEMPERATURE = 0.2

# DB
DRIVER = "{ODBC Driver 18 for SQL Server}"
DB_CONFIG_FILE = CODE_DIR / "db_config.json"
DB_CONFIG_FILE_FALLBACK = DATA_PILOTO_DIR / "db_config.json"

# Parámetros de entrevista (editar en CLI si aplica)
DECLARACION_PARTICIPANTES = ""  # opcional

# Selenium / Edge
EDGE_PROFILE_DIR = r"C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data"
EDGE_PROFILE_NAME = "Default"

# Google Sheets destino para pegar TSV (opcional)
GOOGLE_SHEETS_URL = ""  # pégalo si quieres automatizar el pegado

# ChatGPT Web URL (Selenium)
CHATGPT_URL = "https://chatgpt.com/"

# Copia adicional del resultado (si existe la ruta)
RESULT_COPY_DIR = Path(r"G:\Mi unidad\TGA_PROYECTO\Matriz de sistematización")
