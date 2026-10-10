from pathlib import Path

# Este código vive en: 02.00 DATA PILOTO\01.01 Código de sistematizador
CODE_DIR = Path(__file__).resolve().parent
DATA_PILOTO_DIR = CODE_DIR.parent  # => 02.00 DATA PILOTO

# Insumos (nombres exactos)
GUIA_PDF = DATA_PILOTO_DIR / "PREGUNTA_CODIGO.pdf"
TRANSCRIPCION_PDF = DATA_PILOTO_DIR / "Transcripcion.pdf"
OBJETIVO_TXT = DATA_PILOTO_DIR / "OBJETIVO_FINAL.txt"

# Salidas
PROMP_DIR = DATA_PILOTO_DIR / "01.02 Promp"
RESULT_DIR = DATA_PILOTO_DIR / "01.03 Resultado ChatGPT"
PROMP_DIR.mkdir(parents=True, exist_ok=True)
RESULT_DIR.mkdir(parents=True, exist_ok=True)

PROMPT_TXT_OUT = PROMP_DIR / "prompt_completo.txt"
PROMPT_META_OUT = PROMP_DIR / "prompt_metadata.json"

RESULT_RAW_TXT = RESULT_DIR / "resultado_raw.txt"
RESULT_TSV = RESULT_DIR / "resultado.tsv"
RESULT_CSV = RESULT_DIR / "resultado.csv"
RESULT_XLSX = RESULT_DIR / "resultado.xlsx"

# OpenAI
OPENAI_MODEL = "gpt-4o"
TEMPERATURE = 0.2

# Opcional: fijar roles acá (si lo dejas vacío, no se agrega nada extra)
DECLARACION_PARTICIPANTES = ""
