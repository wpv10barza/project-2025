# -*- coding: utf-8 -*-
from __future__ import annotations

import time
import shutil
from pathlib import Path
from typing import Dict, Optional

import pyperclip
from selenium import webdriver
from selenium.webdriver.edge.service import Service
from selenium.webdriver.edge.options import Options
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

import config as cfg
from src.output_utils import extract_codeblock, write_table_files

def _try_click_copy_code(driver, wait: WebDriverWait) -> bool:
    # Botón puede estar en español o inglés. Probamos varios XPATH.
    candidates = [
        "//button[contains(., 'Copiar código')]",
        "//button[contains(., 'Copy code')]",
        "//button[contains(@aria-label, 'Copy code')]",
        "//button[contains(@aria-label, 'Copiar código')]",
    ]
    for xp in candidates:
        try:
            btn = wait.until(EC.element_to_be_clickable((By.XPATH, xp)))
            driver.execute_script("arguments[0].scrollIntoView({block:'center'});", btn)
            time.sleep(1)
            btn.click()
            return True
        except Exception:
            continue
    return False

def _extract_last_code_text(driver) -> str | None:
    # Fallback: buscar el último <pre><code> visible
    pres = driver.find_elements(By.CSS_SELECTOR, "pre code")
    if pres:
        try:
            return pres[-1].text
        except Exception:
            return None
    return None

def run_chatgpt_via_edge(wait_seconds: int = 420, sheets_url: Optional[str] = None, headless: bool = False) -> Dict[str, str]:
    """
    1) Lee prompt_completo.txt
    2) Abre chatgpt.com en Edge (perfil existente)
    3) Pega prompt, ENTER
    4) Espera
    5) Intenta copiar el bloque TSV (Copy code). Si falla, extrae texto del último codeblock.
    6) Guarda resultado en 01.03 Resultado ChatGPT (TSV/CSV/XLSX) y copia a G: si existe.
    """
    prompt = cfg.PROMPT_TXT_OUT.read_text(encoding="utf-8", errors="ignore")

    options = Options()
    options.add_argument(f"user-data-dir={str(cfg.EDGE_PROFILE_DIR)}")
    options.add_argument(f"--profile-directory={cfg.EDGE_PROFILE_NAME}")
    options.add_argument("--log-level=3")
    if headless:
        options.add_argument("--headless=new")

    # Driver: si existe msedgedriver.exe en la carpeta actual, úsalo; si no, Selenium Manager lo resuelve.
    driver_path = Path.cwd() / "msedgedriver.exe"
    if driver_path.exists():
        service = Service(executable_path=str(driver_path))
        driver = webdriver.Edge(service=service, options=options)
    else:
        driver = webdriver.Edge(options=options)
    wait = WebDriverWait(driver, 90)

    try:
        driver.get(getattr(cfg, 'CHATGPT_URL', 'https://chatgpt.com/'))

        # Caja de prompt (ID suele ser prompt-textarea)
        box = wait.until(EC.element_to_be_clickable((By.ID, "prompt-textarea")))
        box.click()
        pyperclip.copy(prompt)
        box.send_keys(Keys.CONTROL, "v")
        time.sleep(1)
        box.send_keys(Keys.ENTER)

        # Espera generación
        time.sleep(max(5, int(wait_seconds)))

        # Intentar copiar código
        copied = _try_click_copy_code(driver, wait)
        time.sleep(1)

        raw = ""
        if copied:
            try:
                raw = pyperclip.paste() or ""
            except Exception:
                raw = ""
        if not raw:
            raw = _extract_last_code_text(driver) or ""

        cfg.RESULT_RAW_TXT.write_text(raw, encoding="utf-8")

        table = extract_codeblock(raw) or raw.strip()
        if not table:
            table = "NO_TSV_GENERADO"

        write_table_files(table, cfg.RESULT_TSV, cfg.RESULT_CSV, cfg.RESULT_XLSX)

        # Copia a G:\Mi unidad\TGA_PROYECTO\Matriz de sistematización
        try:
            getattr(cfg, 'RESULT_COPY_DIR', None).mkdir(parents=True, exist_ok=True)
            shutil.copy2(cfg.RESULT_TSV, getattr(cfg, 'RESULT_COPY_DIR', None) / cfg.RESULT_TSV.name)
            shutil.copy2(cfg.RESULT_CSV, getattr(cfg, 'RESULT_COPY_DIR', None) / cfg.RESULT_CSV.name)
            shutil.copy2(cfg.RESULT_XLSX, getattr(cfg, 'RESULT_COPY_DIR', None) / cfg.RESULT_XLSX.name)
        except Exception:
            pass

        # Opcional: pegar en Sheets
        su = sheets_url or getattr(cfg, 'GOOGLE_SHEETS_URL', '')
        if su:
            driver.get(su)
            time.sleep(15)
            # intenta pegar
            webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys('v').key_up(Keys.CONTROL).perform()

        return {
            "raw": str(cfg.RESULT_RAW_TXT),
            "tsv": str(cfg.RESULT_TSV),
            "csv": str(cfg.RESULT_CSV),
            "xlsx": str(cfg.RESULT_XLSX),
        }
    finally:
        try:
            driver.quit()
        except Exception:
            pass
