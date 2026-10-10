# Ejecuta GPT (requiere OPENAI_API_KEY)
Set-Location -Path $PSScriptRoot
if (!(Test-Path ".\.venv\Scripts\Activate.ps1")) {
  python -m venv .venv
}
.\.venv\Scripts\Activate.ps1
python -m pip install -r requirements.txt
python .\main.py run
