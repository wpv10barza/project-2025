\
from pathlib import Path

# === CONFIGURACIÓN BASE ===
# Ajusta SOLO si tu carpeta DATA PILOTO está en otra ruta.
BASE_DIR = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\02.00 DATA PILOTO")

# Archivos de entrada
OBJETIVO_FINAL_TXT = BASE_DIR / "OBJETIVO_FINAL.txt"
GUIA_PREGUNTA_CODIGO_PDF = BASE_DIR / "PREGUNTA_CODIGO.pdf"
TRANSCRIPCION_PDF = BASE_DIR / "Transcripcion.pdf"

# Archivo de roles (edítalo luego de crearlo con --init-roles)
ROLES_JSON = BASE_DIR / "roles.json"

# Carpeta de salida para revisión del prompt
OUT_DIR = BASE_DIR / "PRUEBA"
