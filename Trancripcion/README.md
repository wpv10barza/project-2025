# AI Transcripción (Whisper local) + Limpieza (OpenAI API) + Prompt final

Este proyecto corre **en Windows** y está pensado para esta carpeta (puede ser cualquier ruta):
`...\AI transcripcion\`

## Carpeta de trabajo
- `audio/`    → pon aquí tus `.mp3`
- `texto/`    → aquí se guardan transcripciones **crudas** de Whisper (`<audio>.txt`)
- `limpios/`  → aquí se guardan transcripciones **estructuradas** por IA (`<audio>_LIMPIO.txt`)
- `prompts/`  → aquí se guardan prompts finales listos (`PROMPT_FINAL_<audio>.txt`)

## Requisitos (1 vez)
1) **FFmpeg instalado y en PATH**
- Verifica en PowerShell:
  `ffmpeg -version`
- Si no existe y tienes winget:
  `winget install Gyan.FFmpeg`

2) Crear venv + instalar dependencias
```powershell
cd "TU_RUTA\AI transcripcion"
python -m venv venv
.\venv\Scripts\Activate.ps1
python -m pip install --upgrade pip
python -m pip install -r requirements.txt
```

3) Configurar OpenAI API key (solo para el paso de limpieza IA)
```powershell
$env:OPENAI_API_KEY="sk-...."
# opcional:
$env:OPENAI_MODEL="gpt-5.2"   # puedes cambiarlo
```

## Uso rápido
### A) Transcribir MP3 → TXT (Whisper local)
```powershell
.\venv\Scripts\Activate.ps1
python main.py transcribe --model medium
```
- Solo transcribe audios que **aún no tengan** su TXT en `texto/`.

### B) Estructurar/ordenar TXT con OpenAI → `limpios/`
```powershell
python main.py clean --model gpt-5.2
```
- Solo procesa TXT que **aún no tengan** su `_LIMPIO.txt`.

### C) Construir Prompt Final → `prompts/`
```powershell
python main.py prompt
```

### D) Todo en 1 comando
```powershell
python main.py all --whisper-model medium --openai-model gpt-5.2
```

## Notas importantes
- Whisper **no hace diarización** (no sabe quién habla). El limpiador IA usa el texto para **ordenar** y **poner columna HABLANTE** sin inventar. Si no puede inferir, coloca `[?]`.
- Si tu cuenta OpenAI devuelve `429 insufficient_quota`, necesitas revisar facturación/cuota en tu cuenta.


## Inventario de migración desde Google Drive

- Nombre: Trancripcion.
- Propósito documentado: AI Transcripción (Whisper local) + Limpieza (OpenAI API) + Prompt final.
- Periodo de los archivos: 2025-12-23 a 2025-12-23.
- Tecnologías y formatos presentes: .md, .py, .toml, .txt.
- Estructura original: `README.md`, `main.py`, `pyproject.toml`, `requirements.txt`, `src`.
- Módulos principales: `src/ai_transcripcion/cli.py`, `main.py`.
- Origen: Google Drive, Proyectos_Exportados; ID de carpeta `1sH8CD3IGl9bdeaIa5ARLx4bDsBWy66vv`.
- Archivos originales seleccionados: 10.
- Observación: se conserva la estructura y el contenido original de los archivos; la copia se verifica por hash antes de eliminar el origen.
