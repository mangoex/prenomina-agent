from __future__ import annotations

import json
import os
import ast
from pathlib import Path
from typing import Any

from dotenv import load_dotenv

from app.tools_prenomina import calcular_prenomina_excel

try:
    from agents import Agent, ModelSettings, RunConfig, Runner
except ImportError:
    Agent = None  # type: ignore[assignment]
    ModelSettings = None  # type: ignore[assignment]
    RunConfig = None  # type: ignore[assignment]
    Runner = None  # type: ignore[assignment]


PROJECT_ROOT = Path(__file__).resolve().parents[1]
load_dotenv(PROJECT_ROOT / ".env", override=True)


AGENT_INSTRUCTIONS = """
Eres el Agente de Prenomina.

Tu responsabilidad:
- Usar la herramienta calcular_prenomina_excel para leer, validar y calcular la prenomina.
- No inventar calculos ni modificar reglas de negocio.
- No calcular importes mentalmente: los importes se calculan exclusivamente en Python.
- Explicar de forma breve las inconsistencias y resumir el resultado operativo.
- Si hay errores criticos, informar que no se genero total final.
""".strip()


def _get_model_provider() -> str:
    return os.getenv("MODEL_PROVIDER", "openai").strip().lower()


def _get_model_name(provider: str) -> str | None:
    if provider == "openrouter":
        return os.getenv("OPENROUTER_MODEL", "openrouter/openai/gpt-4o-mini").strip()
    return os.getenv("OPENAI_MODEL")


def _validate_provider_config(provider: str) -> None:
    if provider == "openrouter":
        if not os.getenv("OPENROUTER_API_KEY"):
            raise RuntimeError(
                "Falta OPENROUTER_API_KEY en el archivo .env o en las variables de entorno."
            )
        return

    if provider == "openai":
        if not os.getenv("OPENAI_API_KEY"):
            raise RuntimeError(
                "Falta OPENAI_API_KEY en el archivo .env o en las variables de entorno."
            )
        return

    raise RuntimeError("MODEL_PROVIDER debe ser 'openai' u 'openrouter'.")


def _build_run_config(provider: str):
    if RunConfig is None:
        raise RuntimeError(
            "OpenAI Agents SDK no esta instalado. Ejecuta: pip install -r requirements.txt"
        )

    if provider == "openrouter":
        try:
            from agents.extensions.models.litellm_provider import LitellmProvider
        except ImportError as exc:
            raise RuntimeError(
                "Falta LiteLLM. Ejecuta: pip install -r requirements.txt"
            ) from exc

        return RunConfig(
            model_provider=LitellmProvider(),
            tracing_disabled=True,
        )

    return RunConfig()


def _build_agent(provider: str):
    if Agent is None or ModelSettings is None:
        raise RuntimeError(
            "OpenAI Agents SDK no esta instalado. Ejecuta: pip install -r requirements.txt"
        )

    kwargs: dict[str, Any] = {
        "name": "Agente de Prenomina",
        "instructions": AGENT_INSTRUCTIONS,
        "tools": [calcular_prenomina_excel],
        "model_settings": ModelSettings(tool_choice="calcular_prenomina_excel"),
        "tool_use_behavior": "stop_on_first_tool",
    }

    model = _get_model_name(provider)
    if model:
        kwargs["model"] = model

    return Agent(**kwargs)


def _coerce_agent_output(value: Any) -> dict[str, Any]:
    if isinstance(value, dict):
        return value

    if isinstance(value, str):
        try:
            parsed = json.loads(value)
            if isinstance(parsed, dict):
                return parsed
        except json.JSONDecodeError:
            pass

        try:
            parsed = ast.literal_eval(value)
            if isinstance(parsed, dict):
                return parsed
        except (SyntaxError, ValueError):
            pass

        return {
            "ok": False,
            "mensaje": value,
            "archivo_generado": None,
            "empleados_procesados": 0,
            "total_prenomina": None,
            "inconsistencias": [],
        }

    return {
        "ok": False,
        "mensaje": "El agente devolvio un formato no esperado.",
        "archivo_generado": None,
        "empleados_procesados": 0,
        "total_prenomina": None,
        "inconsistencias": [],
    }


async def ejecutar_agente_prenomina(input_path: str, output_path: str) -> dict[str, Any]:
    if Runner is None:
        raise RuntimeError(
            "OpenAI Agents SDK no esta instalado. Ejecuta: pip install -r requirements.txt"
        )

    load_dotenv(PROJECT_ROOT / ".env", override=True)
    provider = _get_model_provider()
    _validate_provider_config(provider)

    agent = _build_agent(provider)
    run_config = _build_run_config(provider)
    prompt = (
        "Procesa este archivo Excel usando la herramienta calcular_prenomina_excel.\n"
        f"input_path: {Path(input_path)}\n"
        f"output_path: {Path(output_path)}\n"
        "Devuelve exactamente el resultado de la herramienta."
    )
    result = await Runner.run(agent, prompt, max_turns=3, run_config=run_config)
    return _coerce_agent_output(result.final_output)
