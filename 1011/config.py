from pathlib import Path

# =========================
# AJUSTA ESTA RUTA
# =========================
BASE_DIR = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\02.00 DATA PILOTO")

# Carpeta de salida solicitada
OUTPUT_DIR = BASE_DIR / "Resultado sistematizado"

# Archivos de entrada (el loader también busca variantes de nombre)
GUIDE_PDF_NAME = "PREGUNTA_CODIGO.pdf"
TRANSCRIPT_PDF_NAME = "Transcripcion.pdf"
OBJECTIVE_TXT_NAME = "OBJETIVO_FINAL.txt"

# Modelo (Structured Outputs soportado en snapshots GPT‑4o recientes)
MODEL = "gpt-4o-2024-08-06"

# Lotes: evita pedir TODO en una sola respuesta (más estable)
BATCH_SIZE = 10

# Columnas a construir en la matriz (ajusta según tus roles)
COLUMNS = [
    "CONVERSACION_ENTREVISTADOR_1",
    "CONVERSACION_ENTREVISTADOR_2",
    "CONVERSACION_INFORMANTE_1",
]
