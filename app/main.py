from __future__ import annotations

import os
import uuid
from pathlib import Path

from dotenv import load_dotenv
from fastapi import Depends, File, Header, HTTPException, Request, UploadFile
from fastapi import FastAPI
from fastapi.responses import FileResponse

from app.agent import ejecutar_agente_prenomina
from app.schemas import ProcesarPrenominaResponse


PROJECT_ROOT = Path(__file__).resolve().parents[1]
load_dotenv(PROJECT_ROOT / ".env")

INPUT_DIR = PROJECT_ROOT / "input"
OUTPUT_DIR = PROJECT_ROOT / "output"
MAX_UPLOAD_MB = int(os.getenv("MAX_UPLOAD_MB", "25"))
MAX_UPLOAD_BYTES = MAX_UPLOAD_MB * 1024 * 1024

app = FastAPI(
    title="Prenomina Agent",
    description="MVP de agente de prenomina con OpenAI Agents SDK y FastAPI.",
    version="0.1.0",
)


@app.get("/health")
def health() -> dict[str, str]:
    return {"status": "ok"}


def require_api_key(x_api_key: str | None = Header(default=None, alias="X-API-Key")) -> None:
    expected = os.getenv("PRENOMINA_API_KEY")
    if expected and x_api_key != expected:
        raise HTTPException(status_code=401, detail="API key invalida o faltante.")


async def save_upload_file(upload: UploadFile, destination: Path) -> None:
    total = 0
    with destination.open("wb") as buffer:
        while True:
            chunk = await upload.read(1024 * 1024)
            if not chunk:
                break
            total += len(chunk)
            if total > MAX_UPLOAD_BYTES:
                raise HTTPException(
                    status_code=413,
                    detail=f"El archivo excede el limite de {MAX_UPLOAD_MB} MB.",
                )
            buffer.write(chunk)


@app.post(
    "/procesar-prenomina",
    response_model=ProcesarPrenominaResponse,
    dependencies=[Depends(require_api_key)],
)
async def procesar_prenomina(request: Request, file: UploadFile = File(...)) -> dict:
    filename = file.filename or ""
    if not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        raise HTTPException(status_code=400, detail="Sube un archivo Excel valido.")

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    safe_name = Path(filename).name
    temp_path = INPUT_DIR / f"{uuid.uuid4().hex}_{safe_name}"

    try:
        await save_upload_file(file, temp_path)

        result_file = OUTPUT_DIR / f"prenomina_resultado_{uuid.uuid4().hex[:8]}.xlsx"
        resultado = await ejecutar_agente_prenomina(
            input_path=str(temp_path),
            output_path=str(result_file),
        )
        generated_path = resultado.get("archivo_generado")
        generated_name = Path(generated_path).name if generated_path else None
        download_url = (
            str(request.url_for("descargar_resultado", filename=generated_name))
            if generated_name
            else None
        )
        return {
            "ok": bool(resultado.get("ok", False)),
            "mensaje": resultado.get("mensaje", ""),
            "archivo_generado": generated_path,
            "download_url": download_url,
            "empleados_procesados": int(resultado.get("empleados_procesados") or 0),
            "total_prenomina": resultado.get("total_prenomina"),
            "inconsistencias": resultado.get("inconsistencias") or [],
        }
    except RuntimeError as exc:
        raise HTTPException(status_code=500, detail=str(exc)) from exc
    except Exception as exc:
        raise HTTPException(status_code=500, detail=f"Error procesando prenomina: {exc}") from exc
    finally:
        await file.close()
        if temp_path.exists():
            temp_path.unlink()


@app.get("/descargar/{filename}", dependencies=[Depends(require_api_key)])
def descargar_resultado(filename: str) -> FileResponse:
    safe_name = Path(filename).name
    if safe_name != filename or not safe_name.startswith("prenomina_resultado_"):
        raise HTTPException(status_code=400, detail="Nombre de archivo invalido.")

    file_path = OUTPUT_DIR / safe_name
    if not file_path.exists():
        raise HTTPException(status_code=404, detail="Archivo no encontrado.")

    return FileResponse(
        file_path,
        media_type="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        filename=safe_name,
    )
