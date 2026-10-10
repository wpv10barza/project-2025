# -*- coding: utf-8 -*-
from __future__ import annotations

import time
import subprocess
from pathlib import Path
from typing import Optional

import pyperclip
from selenium import webdriver
from selenium.common.exceptions import SessionNotCreatedException
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from config import (
    CHATGPT_URL,
    EDGE_USER_DATA_DIR, EDGE_PROFILE_DIR, EDGE_BINARY,
    DEFAULT_PROMPT_TRABAJO
)

def _taskkill(names: list[str]) -> None:
    for n in names:
        try:
            subprocess.run(["taskkill", "/f", "/im", n, "/t"], stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        except Exception:
            pass

def _build_options(
    headless: bool,
    edge_user_data_dir: str | None,
    edge_profile_dir: str | None,
    edge_binary: str | None,
    use_temp_profile: bool,
) -> webdriver.EdgeOptions:
    options = webdriver.EdgeOptions()

    if edge_binary and Path(edge_binary).exists():
        options.binary_location = edge_binary

    # Estabilidad
    options.add_argument("--disable-gpu")
    options.add_argument("--disable-extensions")
    options.add_argument("--no-first-run")
    options.add_argument("--no-default-browser-check")
    options.add_argument("--disable-features=Translate,BackForwardCache")
    options.add_argument("--remote-allow-origins=*")
    options.add_argument("--log-level=3")

    if headless:
        options.add_argument("--headless=new")

    if use_temp_profile:
        tmp = Path.cwd() / ".edge_profile_tmp"
        tmp.mkdir(parents=True, exist_ok=True)
        options.add_argument(f"--user-data-dir={tmp}")
    else:
        if edge_user_data_dir:
            options.add_argument(f"--user-data-dir={edge_user_data_dir}")
            if edge_profile_dir:
                options.add_argument(f"--profile-directory={edge_profile_dir}")

    return options

def run_chatgpt_via_edge(
    wait_seconds: int = 420,
    sheets_url: Optional[str] = None,
    headless: bool = False,
    prompt_path: str | None = None,
    kill_edge: bool = True,
    edge_user_data_dir: str | None = None,
    edge_profile_dir: str | None = None,
    edge_binary: str | None = None,
) -> dict:
    prompt_file = Path(prompt_path).expanduser().resolve() if prompt_path else DEFAULT_PROMPT_TRABAJO.resolve()
    if not prompt_file.exists():
        raise FileNotFoundError(
            f"No existe prompt_trabajo.txt en: {prompt_file}\n"
            "Primero ejecuta: python main.py preview ..."
        )

    prompt_text = prompt_file.read_text(encoding="utf-8", errors="ignore")

    edge_user_data_dir = edge_user_data_dir if edge_user_data_dir is not None else (EDGE_USER_DATA_DIR or None)
    edge_profile_dir   = edge_profile_dir if edge_profile_dir is not None else EDGE_PROFILE_DIR
    edge_binary        = edge_binary if edge_binary is not None else EDGE_BINARY

    # Evita bloqueo del perfil => crash (DevToolsActivePort)
    if kill_edge:
        _taskkill(["msedge.exe", "msedgedriver.exe"])
        time.sleep(2)

    def _start(use_temp_profile: bool):
        options = _build_options(
            headless=headless,
            edge_user_data_dir=edge_user_data_dir,
            edge_profile_dir=edge_profile_dir,
            edge_binary=edge_binary,
            use_temp_profile=use_temp_profile,
        )
        return webdriver.Edge(options=options)  # Selenium Manager

    use_temp = False if edge_user_data_dir else True

    try:
        driver = _start(use_temp_profile=use_temp)
    except SessionNotCreatedException as e:
        msg = str(e)
        if "DevToolsActivePort" in msg or "failed to start" in msg.lower():
            print("[WARN] Edge falló al iniciar con tu perfil. Reintentando con perfil temporal (.edge_profile_tmp).")
            driver = _start(use_temp_profile=True)
        else:
            raise

    try:
        wait = WebDriverWait(driver, 60)
        driver.get(CHATGPT_URL)

        box = wait.until(EC.element_to_be_clickable((By.ID, "prompt-textarea")))
        pyperclip.copy(prompt_text)
        box.click()
        box.send_keys(Keys.CONTROL, "v")
        time.sleep(1.5)
        box.send_keys(Keys.ENTER)

        time.sleep(int(wait_seconds))

        try:
            btn = wait.until(EC.presence_of_element_located((By.XPATH, "//button[contains(., 'Copiar código')]")))
            driver.execute_script("arguments[0].scrollIntoView({block: 'center'});", btn)
            time.sleep(1)
            btn.click()
        except Exception:
            print("[WARN] No encontré el botón 'Copiar código'. Copia manualmente el bloque TSV si es necesario.")

        if sheets_url:
            driver.get(sheets_url)
            time.sleep(15)
            webdriver.ActionChains(driver).key_down(Keys.CONTROL).send_keys("v").key_up(Keys.CONTROL).perform()

    finally:
        try:
            driver.quit()
        except Exception:
            pass
        if kill_edge:
            _taskkill(["msedgedriver.exe"])
    return {"prompt_used": str(prompt_file)}
