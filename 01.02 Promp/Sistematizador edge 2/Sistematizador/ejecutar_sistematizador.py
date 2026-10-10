import time
import pyperclip
import os
from selenium import webdriver
from selenium.webdriver.chrome.service import Service
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC
from webdriver_manager.chrome import ChromeDriverManager

# --- CONFIGURACIÓN ---
# IMPORTANTE: Reemplaza con tu ruta de usuario de Windows
CHROME_PROFILE_PATH = r"C:\Users\ASUS\AppData\Local\Google\Chrome\User Data"
SHEET_URL = "TU_URL_DE_GOOGLE_SHEETS_AQUI"

def iniciar_bot():
    options = webdriver.ChromeOptions()
    options.add_argument(f"user-data-dir={CHROME_PROFILE_PATH}")
    options.add_argument("--profile-directory=Default") 
    options.add_experimental_option("excludeSwitches", ["enable-automation"])
    
    driver = webdriver.Chrome(service=Service(ChromeDriverManager().install()), options=options)
    wait = WebDriverWait(driver, 20)

    try:
        # 1. LEER PROMPT LOCAL
        with open("prompt_base.txt", "r", encoding="utf-8") as f:
            prompt = f.read()

        # 2. CHATGPT
        driver.get("https://chatgpt.com/")
        text_area = wait.until(EC.presence_of_element_located((By.ID, "prompt-textarea")))
        
        # Inyectar texto vía JS (más rápido para prompts largos)
        driver.execute_script("arguments[0].value = arguments[1];", text_area, prompt)
        text_area.send_keys(" ") # Trigger para el botón
        text_area.send_keys(Keys.ENTER)
        
        print("Esperando respuesta de GPT... (ajusta el tiempo si es necesario)")
        time.sleep(30) # Tiempo para que GPT genere la tabla TSV

        # 3. COPIAR RESULTADO (Buscamos el botón de copiar del último bloque de código)
        copy_buttons = driver.find_elements(By.XPATH, "//button[contains(., 'Copy')]")
        if copy_buttons:
            copy_buttons[-1].click()
            time.sleep(1)
        
        # 4. PEGAR EN GOOGLE SHEETS
        driver.get(SHEET_URL)
        time.sleep(8) # Esperar carga de Sheets
        webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()
        print("¡Datos pegados en Sheets!")

    except Exception as e:
        print(f"Error: {e}")
    finally:
        print("Proceso terminado.")

if __name__ == "__main__":
    iniciar_bot()
