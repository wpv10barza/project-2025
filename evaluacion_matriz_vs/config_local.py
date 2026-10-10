\
"""
CONFIGURACIÓN LOCAL (edítalo para tu máquina)

Este archivo está pensado para:
- Visual Studio / VS Code en Windows
- Ejecutar leyendo PDFs desde tu carpeta DATA PILOTO
- Escribir resultados en Resultado ChatGPT

NOTA: No guardes tu API key aquí. Usa variable de entorno OPENAI_API_KEY o un .env.
"""

from pathlib import Path

# ========= RUTAS (Windows) =========
BASE_INPUT_DIR = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\02.00 DATA PILOTO")
OUTPUT_DIR = Path(r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\02.00 DATA PILOTO\Resultado ChatGPT")

# Pega tu OBJETIVO FINAL completo en un .txt (sin resumir)
OBJECTIVE_FINAL_PATH = BASE_INPUT_DIR / "OBJETIVO_FINAL.txt"

# ========= ARCHIVOS (búsqueda por patrón) =========
# El script intentará ubicar los PDFs automáticamente buscando estas palabras clave en el nombre del archivo:
TRANSCRIPT_KEYWORDS = ["transcripcion", "transcripción"]
GUIDE_KEYWORDS_PRIMARY = ["pregunta_codigo", "pregunta codigo", "pregunta-codigo"]
GUIDE_KEYWORDS_SECONDARY = ["pregunta_descripcion", "situacion_deseada", "pregunta_descripcion-pregunta_situacion_deseada"]

# ========= ROLES / PARTICIPANTES =========
# Asignación de roles por speaker exacto (debe coincidir con lo que aparece en la transcripción)
SPEAKER_ROLES = {
    "E1": "Entrevistador principal",
    "E2": "Entrevistador secundario",
    "I1": "Informante 1",
}

# Define quiénes cuentan como ENTREVISTADORES e INFORMANTES para crear columnas dinámicas
INTERVIEWER_SPEAKERS = ["E1", "E2"]
INFORMANT_SPEAKERS = ["I1"]

# ========= MODELO / API =========
MODEL_NAME = "gpt-4o"       # alternativos: "gpt-4o-mini"
TEMPERATURE = 0
MAX_TOKENS = 4096           # NO usar 1,000,000 (es inválido / costoso). Ajusta según tu plan.

# ========= CONTROLES DE PROMPT =========
# Si tu transcripción es muy larga, limita lo que se envía al modelo (en caracteres)
MAX_TRANSCRIPT_CHARS = 200_000

# Guardar prompts para auditoría
SAVE_PROMPTS = True

# Reintentos ante rate-limit/transitorios
MAX_RETRIES = 6
RETRY_BACKOFF_BASE_SECONDS = 2.0
