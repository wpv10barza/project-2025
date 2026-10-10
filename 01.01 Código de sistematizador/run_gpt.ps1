# Requiere: $env:OPENAI_API_KEY="TU_API_KEY"
$ErrorActionPreference = "Stop"

. .\.venv\Scripts\Activate.ps1
python main.py run
