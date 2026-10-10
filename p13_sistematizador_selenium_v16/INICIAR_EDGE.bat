@echo off
echo --- CERRANDO PROCESOS DE EDGE... ---
taskkill /F /IM msedge.exe /T >nul 2>&1
taskkill /F /IM msedgedriver.exe /T >nul 2>&1
timeout /t 2 >nul

echo --- ABRIENDO EDGE EN MODO DEPURACION (PUERTO 9222)... ---
start "" "C:\Program Files (x86)\Microsoft\Edge\Application\msedge.exe" --remote-debugging-port=9222 --user-data-dir="C:\Users\ASUS\AppData\Local\Microsoft\Edge\User Data" --profile-directory="Default"

echo --- LISTO. NO CIERRES LA VENTANA NEGRA HASTA QUE EDGE ABRA. ---
timeout /t 5
exit