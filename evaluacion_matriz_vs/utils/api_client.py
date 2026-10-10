from __future__ import annotations

import os
import time
from dataclasses import dataclass
from typing import Optional

from openai import OpenAI
from openai import RateLimitError, APIError, APIConnectionError, Timeout


@dataclass
class OpenAIConfig:
    model_name: str
    temperature: float = 0.0
    max_tokens: int = 4096
    max_retries: int = 6
    backoff_base_seconds: float = 2.0


def get_client() -> OpenAI:
    api_key = os.environ.get("OPENAI_API_KEY", "").strip()
    if not api_key:
        raise RuntimeError(
            "No se encontró OPENAI_API_KEY en variables de entorno.\n"
            "En Windows PowerShell: $env:OPENAI_API_KEY='sk-...'"
        )
    return OpenAI(api_key=api_key)


def chat_completion_text(client: OpenAI, prompt: str, cfg: OpenAIConfig) -> str:
    """Llama a Chat Completions con reintentos y devuelve texto plano."""
    last_err: Optional[Exception] = None

    for attempt in range(cfg.max_retries + 1):
        try:
            resp = client.chat.completions.create(
                model=cfg.model_name,
                messages=[{"role": "user", "content": prompt}],
                temperature=cfg.temperature,
                max_tokens=cfg.max_tokens,
            )
            return (resp.choices[0].message.content or "").strip()

        except RateLimitError as e:
            # Incluye 'insufficient_quota' y rate limits
            last_err = e
            # Si es insufficient_quota, reintentar no ayuda
            msg = str(e)
            if "insufficient_quota" in msg or "exceeded your current quota" in msg:
                raise

            sleep_s = cfg.backoff_base_seconds * (2 ** attempt)
            time.sleep(sleep_s)

        except (APIConnectionError, Timeout, APIError) as e:
            last_err = e
            sleep_s = cfg.backoff_base_seconds * (2 ** attempt)
            time.sleep(sleep_s)

        except Exception as e:
            # Otros errores inesperados: no reintentar mucho
            raise

    # Si agotamos reintentos
    raise RuntimeError(f"Fallo tras reintentos. Último error: {last_err}")
