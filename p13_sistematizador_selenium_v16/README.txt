\
P13 Sistematizador Selenium v16 (FIX Edge crash / DevToolsActivePort)

✅ Fix incluido para:
selenium.common.exceptions.SessionNotCreatedException:
  DevToolsActivePort file doesn't exist
  Microsoft Edge failed to start: crashed

Qué hace este fix:
- Por defecto cierra Edge y msedgedriver antes de iniciar (evita bloqueo del perfil).
- Permite usar tu perfil real (para ya estar logueado en chatgpt.com):
    --edge-user-data-dir "C:\Users\...\Microsoft\Edge\User Data"
    --edge-profile-dir "Default"
- Si aun así Edge crashea, reintenta automáticamente con un perfil temporal (.edge_profile_tmp)
  (ojo: con perfil temporal tendrás que loguearte 1 vez).

Uso rápido:
  python -m venv .venv
  .\.venv\Scripts\Activate.ps1
  python -m pip install -r .\requirements.txt

  python .\main.py preview --evaluacion-id 17 --db-config ".\db_config.json"
  python .\main.py run --wait-seconds 420 --edge-user-data-dir "C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data" --edge-profile-dir "Default"

