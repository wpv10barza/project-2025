import time
import os
import winsound
import pyperclip
from datetime import datetime
from selenium import webdriver
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

# --- RUTAS ---
URL_SHEETS = "https://docs.google.com/spreadsheets/d/1fBb6cqdwNel5VgwWb4A4YQ6U4v2r4XgPt7-o4_vpQ7c/edit#gid=0"
ARCHIVO_PROMPT = r"C:\Users\ASUS\GESTIONA ACTIVOS S.A.C\TGA_GACSAC - TGA_TI - 03. Proyectos\P13. Sistema Evaluación v2\03.00 EJECUCION\03.00 DATA PILOTO\01.02 Promp\prompt_completo.txt"

def ejecutar_proceso_dp():
    print(f"\n[{datetime.now().strftime('%H:%M:%S')}] >>> Conectando a Edge (Puerto 9222)...")
    
    # Lectura del prompt
    try:
        with open(ARCHIVO_PROMPT, "r", encoding="utf-8") as f:
            prompt_input = f.read()
    except Exception as e:
        print(f"!!! Error leyendo TXT: {e}")
        return

    # CONFIGURACIÓN CLAVE: Conectarse a navegador existente
    options = Options()
    options.add_experimental_option("debuggerAddress", "127.0.0.1:9222")

    driver = None
    try:
        # Iniciamos driver conectando al puerto, SIN abrir ventana nueva
        driver = webdriver.Edge(options=options)
        
        # Ir a ChatGPT (usa la pestaña actual o abre una en la misma ventana)
        driver.get("https://chatgpt.com/")
        wait = WebDriverWait(driver, 30)
        
        # 1. PEGAR PROMPT
        print(">>> Pegando Prompt...")
        box = wait.until(EC.element_to_be_clickable((By.ID, "prompt-textarea")))
        pyperclip.copy(prompt_input)
        box.click()
        box.send_keys(Keys.CONTROL, 'v')
        time.sleep(2)
        box.send_keys(Keys.ENTER)
        
        # 2. ESPERA (7 MINUTOS)
        print(">>> ChatGPT trabajando (7 min)... (Puedes minimizar Edge)")
        time.sleep(420) 
        
        # 3. COPIAR CÓDIGO TSV
        print(">>> Buscando botón 'Copiar código'...")
        try:
            # Buscar botón específico de bloques de código
            boton_copiar = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Copiar código')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", boton_copiar)
            time.sleep(1)
            boton_copiar.click()
            print(">>> ¡Matriz copiada!")
        except:
            print("!!! No se halló el botón 'Copiar código'. Intentando fallback...")
            try:
                # Intento secundario por icono
                btns = driver.find_elements(By.CSS_SELECTOR, "button.flex.ml-auto")
                if btns: btns[-1].click()
            except:
                print("!!! ERROR CRÍTICO: No se pudo copiar.")

        # 4. PEGAR EN SHEETS
        print(">>> Pegando en Sheets...")
        driver.get(URL_SHEETS)
        time.sleep(15) # Espera carga de hoja
        
        webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        
        print(">>> CICLO COMPLETADO.")
        winsound.Beep(1000, 500)
        
    except Exception as e:
        print(f"!!! Error de conexión: {e}")
        print("IMPORTANTE: ¿Ejecutaste primero el archivo INICIAR_EDGE.bat?")

def monitor():
    print("--- MONITOR ACTIVO (MODO PUERTO 9222) ---")
    print(f"Ruta Prompt: {ARCHIVO_PROMPT}")
    
    if not os.path.exists(ARCHIVO_PROMPT):
        print("!!! EL ARCHIVO DE TEXTO NO EXISTE.")
        return
    
    ultima_m = os.path.getmtime(ARCHIVO_PROMPT)
    while True:
        try:
            actual_m = os.path.getmtime(ARCHIVO_PROMPT)
            if actual_m != ultima_m:
                time.sleep(2) # Espera a que termine de guardar
                ejecutar_proceso_dp()
                ultima_m = os.path.getmtime(ARCHIVO_PROMPT)
                print("\n--- Esperando cambios... ---")
            time.sleep(5)
        except KeyboardInterrupt: break
        except: time.sleep(5)

if __name__ == "__main__":
    monitor()
