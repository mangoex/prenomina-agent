from __future__ import annotations

import os
import uuid
from pathlib import Path
from html import escape
from urllib.parse import urlencode

from dotenv import load_dotenv
from fastapi import Depends, File, Form, Header, HTTPException, Query, Request, UploadFile
from fastapi import FastAPI
from fastapi.responses import FileResponse, HTMLResponse

from app.agent import ejecutar_agente_prenomina
from app.schemas import ProcesarPrenominaResponse
from app.tools_prenomina import procesar_prenomina_excel


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


def _web_access_enabled() -> bool:
    return bool(os.getenv("WEB_ACCESS_KEY"))


def _require_web_access(access_key: str | None) -> None:
    expected = os.getenv("WEB_ACCESS_KEY")
    if expected and access_key != expected:
        raise HTTPException(status_code=401, detail="Clave de acceso invalida o faltante.")


def _render_web_page(
    *,
    result: dict | None = None,
    error: str | None = None,
    download_url: str | None = None,
    form_values: dict | None = None,
) -> HTMLResponse:
    values = {
        "fondo_ahorro_factor": "0.11",
        "uma_diaria": "117.31",
        "fondo_ahorro_tope_mode": "mensual",
        "use_excel_fondo_ahorro": True,
        "dias_base_periodo": "30.4",
    }
    if form_values:
        values.update(form_values)

    issues = result.get("inconsistencias", []) if result else []
    warning_count = sum(1 for issue in issues if issue.get("severidad") == "advertencia")
    critical_count = sum(1 for issue in issues if issue.get("severidad") == "critico")
    rows = "\n".join(
        f"""
        <tr>
          <td>{escape(str(issue.get("severidad", "")))}</td>
          <td>{escape(str(issue.get("tipo", "")))}</td>
          <td>{escape(str(issue.get("mensaje", "")))}</td>
        </tr>
        """
        for issue in issues[:12]
    )
    if not rows:
        rows = "<tr><td colspan=\"3\">Sin advertencias para mostrar.</td></tr>"

    access_field = ""
    if _web_access_enabled():
        access_field = """
        <label class="field">
          <span>Clave de acceso</span>
          <input name="access_key" type="password" autocomplete="current-password" required>
        </label>
        """

    selected = {
        mode: "selected" if values["fondo_ahorro_tope_mode"] == mode else ""
        for mode in ["mensual", "quincenal", "proporcional", "none"]
    }
    checked = "checked" if values.get("use_excel_fondo_ahorro") else ""

    result_panel = ""
    if error:
        result_panel = f"""
        <section class="result error">
          <h2>No se pudo generar la prenomina</h2>
          <p>{escape(error)}</p>
        </section>
        """
    elif result:
        total = result.get("total_prenomina")
        total_text = "Pendiente" if total is None else f"${float(total):,.2f}"
        download_button = (
            f'<a class="button primary" href="{escape(download_url or "")}">Descargar Excel</a>'
            if download_url
            else ""
        )
        result_panel = f"""
        <section class="result">
          <div class="result-head">
            <div>
              <h2>{escape(str(result.get("mensaje", "Prenomina procesada.")))}</h2>
              <p>Revisa advertencias antes de usar el archivo como calculo final.</p>
            </div>
            {download_button}
          </div>
          <div class="metrics" aria-label="Resumen del procesamiento">
            <div><span>Empleados</span><strong>{int(result.get("empleados_procesados") or 0)}</strong></div>
            <div><span>Total</span><strong>{total_text}</strong></div>
            <div><span>Advertencias</span><strong>{warning_count}</strong></div>
            <div><span>Criticos</span><strong>{critical_count}</strong></div>
          </div>
          <div class="table-wrap">
            <table>
              <thead><tr><th>Severidad</th><th>Tipo</th><th>Mensaje</th></tr></thead>
              <tbody>{rows}</tbody>
            </table>
          </div>
        </section>
        """

    html = f"""
    <!doctype html>
    <html lang="es">
    <head>
      <meta charset="utf-8">
      <meta name="viewport" content="width=device-width, initial-scale=1">
      <title>Generador de prenomina</title>
      <link rel="icon" href="data:,">
      <style>
        :root {{
          --bg: #f5f7fb;
          --surface: #ffffff;
          --surface-2: #eef3f8;
          --text: #172033;
          --muted: #647087;
          --border: #d8e0ea;
          --accent: #126c68;
          --accent-strong: #0d504d;
          --danger: #9f2d35;
          --shadow: 0 20px 50px rgba(28, 39, 58, .10);
        }}
        * {{ box-sizing: border-box; }}
        body {{
          margin: 0;
          font-family: Inter, ui-sans-serif, system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
          background: var(--bg);
          color: var(--text);
        }}
        main {{
          width: min(1180px, calc(100% - 32px));
          margin: 0 auto;
          padding: 28px 0 48px;
        }}
        header {{
          display: flex;
          justify-content: space-between;
          align-items: center;
          gap: 20px;
          padding: 12px 0 26px;
        }}
        .brand {{
          display: flex;
          align-items: center;
          gap: 12px;
          font-weight: 750;
          letter-spacing: 0;
        }}
        .mark {{
          width: 38px;
          height: 38px;
          border-radius: 8px;
          background: var(--accent);
          display: grid;
          place-items: center;
          color: white;
          font-weight: 800;
        }}
        .status {{
          color: var(--accent-strong);
          font-size: 14px;
          font-weight: 650;
        }}
        .intro {{
          display: grid;
          grid-template-columns: minmax(0, 1.1fr) minmax(320px, .9fr);
          gap: 28px;
          align-items: start;
        }}
        h1 {{
          font-size: 42px;
          line-height: 1.08;
          margin: 0 0 12px;
          letter-spacing: 0;
        }}
        .lead {{
          color: var(--muted);
          font-size: 17px;
          line-height: 1.6;
          margin: 0 0 24px;
          max-width: 680px;
        }}
        .steps {{
          display: grid;
          gap: 10px;
          margin-top: 24px;
        }}
        .step {{
          display: grid;
          grid-template-columns: 32px 1fr;
          gap: 12px;
          align-items: start;
          color: var(--muted);
          font-size: 15px;
          line-height: 1.45;
        }}
        .step b {{
          color: var(--text);
        }}
        .number {{
          width: 32px;
          height: 32px;
          border-radius: 8px;
          border: 1px solid var(--border);
          display: grid;
          place-items: center;
          color: var(--accent-strong);
          font-weight: 750;
          background: white;
        }}
        form, .result {{
          background: var(--surface);
          border: 1px solid var(--border);
          border-radius: 8px;
          box-shadow: var(--shadow);
        }}
        form {{
          padding: 22px;
          display: grid;
          gap: 18px;
        }}
        .field {{
          display: grid;
          gap: 7px;
        }}
        label span, .toggle span {{
          color: #334058;
          font-size: 14px;
          font-weight: 700;
        }}
        input, select {{
          width: 100%;
          border: 1px solid var(--border);
          border-radius: 8px;
          padding: 12px 13px;
          color: var(--text);
          background: #fff;
          font-size: 15px;
          line-height: 1.25;
        }}
        input[type="file"] {{
          padding: 11px;
          background: var(--surface-2);
        }}
        input[type="file"]::file-selector-button {{
          border: 1px solid var(--border);
          border-radius: 6px;
          padding: 9px 12px;
          margin-right: 10px;
          background: white;
          color: var(--text);
          font-weight: 700;
          cursor: pointer;
        }}
        input:focus, select:focus {{
          outline: 3px solid rgba(18, 108, 104, .16);
          border-color: var(--accent);
        }}
        .grid {{
          display: grid;
          grid-template-columns: repeat(2, minmax(0, 1fr));
          gap: 14px;
        }}
        .toggle {{
          display: flex;
          gap: 10px;
          align-items: flex-start;
          color: var(--muted);
          font-size: 14px;
          line-height: 1.4;
          padding: 12px;
          border: 1px solid var(--border);
          border-radius: 8px;
          background: #fbfcfe;
        }}
        .toggle input {{
          width: 18px;
          height: 18px;
          margin-top: 2px;
        }}
        .button {{
          appearance: none;
          border: 0;
          border-radius: 8px;
          padding: 13px 16px;
          font-size: 15px;
          font-weight: 750;
          text-align: center;
          text-decoration: none;
          cursor: pointer;
          display: inline-flex;
          align-items: center;
          justify-content: center;
          min-height: 46px;
        }}
        .primary {{
          color: white;
          background: var(--accent);
        }}
        .primary:hover {{
          background: var(--accent-strong);
        }}
        .result {{
          margin-top: 28px;
          padding: 22px;
        }}
        .result.error {{
          border-color: rgba(159, 45, 53, .35);
        }}
        .result.error h2 {{
          color: var(--danger);
        }}
        .result-head {{
          display: flex;
          justify-content: space-between;
          gap: 18px;
          align-items: flex-start;
          margin-bottom: 18px;
        }}
        h2 {{
          font-size: 22px;
          line-height: 1.2;
          margin: 0 0 6px;
          letter-spacing: 0;
        }}
        .result p {{
          margin: 0;
          color: var(--muted);
        }}
        .metrics {{
          display: grid;
          grid-template-columns: repeat(4, minmax(0, 1fr));
          gap: 10px;
          margin: 18px 0;
        }}
        .metrics div {{
          background: var(--surface-2);
          border-radius: 8px;
          padding: 12px;
        }}
        .metrics span {{
          display: block;
          color: var(--muted);
          font-size: 13px;
          margin-bottom: 5px;
        }}
        .metrics strong {{
          font-size: 20px;
        }}
        .table-wrap {{
          overflow-x: auto;
          border: 1px solid var(--border);
          border-radius: 8px;
        }}
        table {{
          width: 100%;
          border-collapse: collapse;
          font-size: 14px;
        }}
        th, td {{
          text-align: left;
          padding: 11px 12px;
          border-bottom: 1px solid var(--border);
          vertical-align: top;
        }}
        th {{
          background: #f7f9fc;
          color: #334058;
          font-size: 13px;
        }}
        tr:last-child td {{
          border-bottom: 0;
        }}
        @media (max-width: 820px) {{
          main {{ width: min(100% - 20px, 680px); padding-top: 18px; }}
          header {{ align-items: flex-start; flex-direction: column; gap: 8px; }}
          .intro {{ grid-template-columns: 1fr; }}
          h1 {{ font-size: 32px; }}
          .grid, .metrics {{ grid-template-columns: 1fr; }}
          .result-head {{ flex-direction: column; }}
          .button {{ width: 100%; }}
        }}
      </style>
    </head>
    <body>
      <main>
        <header>
          <div class="brand"><div class="mark">P</div><span>Generador de prenomina</span></div>
          <div class="status">API activa</div>
        </header>
        <section class="intro">
          <div>
            <h1>Sube el Excel, valida los criterios y descarga la prenomina.</h1>
            <p class="lead">La pantalla usa el motor de calculo del agente y deja trazabilidad de UMA, fondo de ahorro, SDI, diferencias contra Excel y advertencias para revision.</p>
            <div class="steps">
              <div class="step"><div class="number">1</div><div><b>Selecciona el archivo.</b><br>Usa un Excel de asistencia o prenomina administrativa.</div></div>
              <div class="step"><div class="number">2</div><div><b>Confirma los criterios.</b><br>Ajusta UMA, factor y modo de tope antes de generar.</div></div>
              <div class="step"><div class="number">3</div><div><b>Descarga el resultado.</b><br>El archivo incluye hojas de calculo, diferencias, inconsistencias y resumen.</div></div>
            </div>
          </div>
          <form method="post" action="/web/procesar-prenomina" enctype="multipart/form-data">
            {access_field}
            <label class="field">
              <span>Archivo Excel</span>
              <input name="file" type="file" accept=".xlsx,.xlsm,.xls" required>
            </label>
            <div class="grid">
              <label class="field">
                <span>Factor fondo de ahorro</span>
                <input name="fondo_ahorro_factor" type="number" min="0" step="0.0001" value="{escape(str(values["fondo_ahorro_factor"]))}" required>
              </label>
              <label class="field">
                <span>UMA diaria</span>
                <input name="uma_diaria" type="number" min="0" step="0.01" value="{escape(str(values["uma_diaria"]))}" required>
              </label>
              <label class="field">
                <span>Modo de tope</span>
                <select name="fondo_ahorro_tope_mode">
                  <option value="mensual" {selected["mensual"]}>Mensual</option>
                  <option value="quincenal" {selected["quincenal"]}>Quincenal</option>
                  <option value="proporcional" {selected["proporcional"]}>Proporcional</option>
                  <option value="none" {selected["none"]}>Sin tope</option>
                </select>
              </label>
              <label class="field">
                <span>Dias base periodo</span>
                <input name="dias_base_periodo" type="number" min="0" step="0.1" value="{escape(str(values["dias_base_periodo"]))}" required>
              </label>
            </div>
            <label class="toggle">
              <input name="use_excel_fondo_ahorro" type="checkbox" value="true" {checked}>
              <span>Conservar el fondo de ahorro del Excel en el calculo principal mientras se valida el criterio.</span>
            </label>
            <button class="button primary" type="submit">Generar prenomina</button>
          </form>
        </section>
        {result_panel}
      </main>
    </body>
    </html>
    """
    return HTMLResponse(html)


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


@app.get("/", response_class=HTMLResponse)
def web_home() -> HTMLResponse:
    return _render_web_page()


@app.post("/web/procesar-prenomina", response_class=HTMLResponse)
async def web_procesar_prenomina(
    request: Request,
    file: UploadFile = File(...),
    access_key: str = Form(default=""),
    fondo_ahorro_factor: float = Form(default=0.11),
    uma_diaria: float = Form(default=117.31),
    fondo_ahorro_tope_mode: str = Form(default="mensual"),
    dias_base_periodo: float = Form(default=30.4),
    use_excel_fondo_ahorro: str = Form(default="false"),
) -> HTMLResponse:
    _require_web_access(access_key)
    filename = file.filename or ""
    form_values = {
        "fondo_ahorro_factor": fondo_ahorro_factor,
        "uma_diaria": uma_diaria,
        "fondo_ahorro_tope_mode": fondo_ahorro_tope_mode,
        "use_excel_fondo_ahorro": use_excel_fondo_ahorro == "true",
        "dias_base_periodo": dias_base_periodo,
    }
    if not filename.lower().endswith((".xlsx", ".xlsm", ".xls")):
        return _render_web_page(
            error="Sube un archivo Excel valido.",
            form_values=form_values,
        )

    INPUT_DIR.mkdir(parents=True, exist_ok=True)
    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)

    safe_name = Path(filename).name
    temp_path = INPUT_DIR / f"{uuid.uuid4().hex}_{safe_name}"
    result_file = OUTPUT_DIR / f"prenomina_resultado_{uuid.uuid4().hex[:8]}.xlsx"

    try:
        await save_upload_file(file, temp_path)
        resultado = procesar_prenomina_excel(
            input_path=str(temp_path),
            output_path=str(result_file),
            admin_config_overrides={
                "fondo_ahorro_factor": fondo_ahorro_factor,
                "uma_diaria": uma_diaria,
                "fondo_ahorro_tope_mode": fondo_ahorro_tope_mode,
                "use_excel_fondo_ahorro": use_excel_fondo_ahorro == "true",
                "dias_base_periodo": dias_base_periodo,
            },
        )
        generated_path = resultado.get("archivo_generado")
        generated_name = Path(generated_path).name if generated_path else None
        download_url = None
        if generated_name:
            download_url = f"/web/descargar/{generated_name}"
            if _web_access_enabled():
                download_url = f"{download_url}?{urlencode({'access_key': access_key})}"
        return _render_web_page(
            result=resultado,
            download_url=download_url,
            form_values=form_values,
        )
    except Exception as exc:
        return _render_web_page(
            error=f"Error procesando prenomina: {exc}",
            form_values=form_values,
        )
    finally:
        await file.close()
        if temp_path.exists():
            temp_path.unlink()


@app.get("/web/descargar/{filename}")
def web_descargar_resultado(
    filename: str,
    access_key: str | None = Query(default=None),
) -> FileResponse:
    _require_web_access(access_key)
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
