# Matriz de Sistematización de Entrevistas (Visual Studio / Windows)

Este proyecto genera una **Matriz de Sistematización** a partir de:
- Un PDF de **Transcripción** (ej. `Transcripcion.pdf`) con líneas tipo `L001 E1 ...`
- Uno o más PDFs con **guía de preguntas** (códigos X.X.X, descripción y situación deseada)

El script:
1) Extrae texto de PDFs (sin modificar los archivos).
2) Parsear transcripción a una tabla estructurada `LÍNEA	SPEAKER	ROL	TEXTO`.
3) Extraer preguntas (PREGUNTA_CODIGO, DESCRIPCION, SITUACION_DESEADA).
4) Para cada PREGUNTA_CODIGO, arma un prompt con tu **Objetivo Final** y pide al modelo una **línea TSV**.
5) Guarda:
   - `resultado_rows.tsv` (una fila TSV por pregunta)
   - `systematization_matrix.csv` y `systematization_matrix.xlsx` (matriz consolidada)
   en la carpeta `Resultado ChatGPT`.

---

## 1) Requisitos
- Python 3.10+ (recomendado 3.11)
- Una API key válida en variable de entorno: `OPENAI_API_KEY`

> **Importante**: NO pegues tu API key dentro del código. Usa variables de entorno o un `.env` local.

---

## 2) Instalación (Windows PowerShell)
```powershell
cd ruta\al\proyecto
python -m venv .venv
.\.venv\Scripts\Activate.ps1
pip install -r requirements.txt
```

---

## 3) Configuración rápida (rutas y archivos)
Edita `config_local.py` y ajusta:

- `BASE_INPUT_DIR` -> carpeta donde están tus PDFs (DATA PILOTO)
- `OUTPUT_DIR` -> carpeta de salida (Resultado ChatGPT)
- `OBJECTIVE_FINAL_PATH` -> archivo `.txt` donde pegarás tu **OBJETIVO FINAL** completo (sin resumir)
- `SPEAKER_ROLES` -> roles por speaker (E1/E2/I1, etc.)
- `INTERVIEWER_SPEAKERS` y `INFORMANT_SPEAKERS` -> para columnas dinámicas

---

## 4) Ejecutar
```powershell
python main.py
```

Opcional:
```powershell
python main.py --limit 5
python main.py --codes 1.1.1,1.2.1
python main.py --dry-run
```

---

## 5) Outputs
Por defecto se genera en:
`OUTPUT_DIR\run_YYYYMMDD_HHMMSS\`

Archivos:
- `resultado_rows.tsv`
- `systematization_matrix.csv`
- `systematization_matrix.xlsx`
- `prompts/` (prompts enviados por pregunta, para auditoría)

---

## 6) Notas sobre tamaño de prompt y costos
La transcripción puede ser grande. El proyecto incluye:
- `MAX_TRANSCRIPT_CHARS` para truncar si se necesita.
- Reintentos automáticos ante RateLimitError.

Si tu cuota es limitada, prueba primero con:
- `MODEL_NAME="gpt-4o-mini"` (más barato) y luego pasas a `gpt-4o`.



## Inventario de migración desde Google Drive

- Nombre: evaluacion_matriz_vs.
- Propósito documentado: Matriz de Sistematización de Entrevistas (Visual Studio / Windows).
- Periodo de los archivos: 2025-12-16 a 2025-12-16.
- Tecnologías y formatos presentes: .bat, .md, .py, .txt.
- Estructura original: `README.md`, `config_local.py`, `main.py`, `requirements.txt`, `run_windows.bat`, `utils`.
- Módulos principales: `utils/matrix_builder.py`, `utils/parse_transcript.py`, `utils/api_client.py`, `utils/parse_questions.py`, `utils/prompt_builder.py`, `utils/pdf_utils.py`, `main.py`, `config_local.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1-i9a3W6n0gj9dvRS4lDMzw1_MzVZFtkX`.
- Archivos originales seleccionados: 11.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
