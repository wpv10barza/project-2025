from __future__ import annotations
import time, os
from pathlib import Path
from datetime import datetime
import pyperclip

from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import EDGE_PROFILE_DIR, EDGE_PROFILE_NAME, CHATGPT_URL, DEFAULT_WAIT_SECONDS

def _kill_edge():
    os.system('taskkill /f /im msedge.exe /t >nul 2>&1')
    os.system('taskkill /f /im msedgedriver.exe /t >nul 2>&1')
    time.sleep(2)

def run_chatgpt_via_edge(prompt_path: Path, wait_seconds: int = DEFAULT_WAIT_SECONDS, sheets_url: str|None=None, headless: bool=False) -> None:
    prompt = Path(prompt_path).read_text(encoding="utf-8", errors="ignore")

    _kill_edge()
    print(f"[{datetime.now().strftime('%H:%M:%S')}] >>> Abriendo Edge/ChatGPT...")

    options = Options()
    options.add_argument(f"user-data-dir={EDGE_PROFILE_DIR}")
    options.add_argument(f"--profile-directory={EDGE_PROFILE_NAME}")
    options.add_argument("--log-level=3")
    options.add_experimental_option("detach", True)
    if headless:
        options.add_argument("--headless=new")

    ruta_driver = str(Path.cwd() / "msedgedriver.exe")
    if not Path(ruta_driver).exists():
        raise FileNotFoundError(f"No se encontró msedgedriver.exe en: {ruta_driver}")

    driver = webdriver.Edge(service=Service(executable_path=ruta_driver), options=options)
    try:
        driver.get(CHATGPT_URL)
        wait = WebDriverWait(driver, 60)

        box = wait.until(EC.element_to_be_clickable((By.ID, "prompt-textarea")))
        pyperclip.copy(prompt)
        box.click()
        box.send_keys(Keys.CONTROL, 'v')
        time.sleep(1)
        box.send_keys(Keys.ENTER)

        print(f">>> ChatGPT generando (espera {wait_seconds}s)...")
        time.sleep(wait_seconds)

        if sheets_url:
            print(">>> Abriendo Google Sheets para pegar portapapeles...")
            driver.get(sheets_url)
            time.sleep(15)
            webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

        print(">>> RUN OK (revisa el resultado en ChatGPT / Sheets).")
    finally:
        driver.quit()
        _kill_edge()
