from pathlib import Path

from fastapi.testclient import TestClient

from app.main import app


def test_web_home_renders_form(monkeypatch):
    monkeypatch.delenv("WEB_ACCESS_KEY", raising=False)
    client = TestClient(app)

    response = client.get("/")

    assert response.status_code == 200
    assert "Generador de prenomina" in response.text
    assert "name=\"file\"" in response.text
    assert "name=\"uma_diaria\"" in response.text


def test_web_upload_processes_sample_file(monkeypatch):
    monkeypatch.delenv("WEB_ACCESS_KEY", raising=False)
    client = TestClient(app)
    sample_path = Path("samples/asistencia_ejemplo.xlsx")

    with sample_path.open("rb") as sample_file:
        response = client.post(
            "/web/procesar-prenomina",
            data={
                "fondo_ahorro_factor": "0.11",
                "uma_diaria": "117.31",
                "fondo_ahorro_tope_mode": "mensual",
                "dias_base_periodo": "30.4",
                "use_excel_fondo_ahorro": "true",
            },
            files={
                "file": (
                    sample_path.name,
                    sample_file,
                    "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )
            },
        )

    assert response.status_code == 200
    assert "Prenomina calculada correctamente" in response.text
    assert "Descargar Excel" in response.text
