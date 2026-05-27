from __future__ import annotations

from typing import Any, Literal

from pydantic import BaseModel, Field


class Inconsistencia(BaseModel):
    severidad: Literal["critico", "advertencia"]
    tipo: str
    fila: int | None = None
    columna: str | None = None
    mensaje: str


class ProcesarPrenominaResponse(BaseModel):
    ok: bool
    mensaje: str
    archivo_generado: str | None = None
    download_url: str | None = None
    empleados_procesados: int = 0
    total_prenomina: float | None = None
    inconsistencias: list[dict[str, Any]] = Field(default_factory=list)
    analisis_agente: dict[str, Any] | None = None
    analisis_agente_error: str | None = None
