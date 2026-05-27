# Windows Desktop Installer Implementation Plan

> **For agentic workers:** REQUIRED SUB-SKILL: Use superpowers:subagent-driven-development (recommended) or superpowers:executing-plans to implement this plan task-by-task. Steps use checkbox (`- [ ]`) syntax for tracking.

**Goal:** Build a Windows desktop installer for the Prenomina Agent so the accountant can install and run it as a local desktop app without installing Python manually.

**Architecture:** Add a small desktop launcher that starts the existing FastAPI app on `127.0.0.1`, opens it in a desktop window with `pywebview`, and shuts the local server down when the window closes. Package that launcher with PyInstaller on a Windows GitHub Actions runner, then wrap the executable with an Inno Setup installer artifact.

**Tech Stack:** Python 3.12, FastAPI, Uvicorn, pywebview, PyInstaller, Inno Setup, GitHub Actions on `windows-latest`.

---

## File Structure

- Create `desktop_app.py`: desktop entrypoint that loads `.env`, creates local working folders, starts Uvicorn in a background thread, waits for `/health`, opens pywebview, and shuts down cleanly.
- Modify `requirements.txt`: add `pywebview` for the desktop window and `pyinstaller` for local/manual packaging.
- Create `installer/prenomina.iss`: Inno Setup script that installs `Prenomina.exe`, creates desktop/start-menu shortcuts, and creates user-facing documents/output folders.
- Create `.github/workflows/build-windows.yml`: manual GitHub Actions workflow that builds `Prenomina.exe`, compiles `Prenomina Setup.exe`, and uploads it as an artifact.
- Create `docs/windows_desktop_build.md`: operator instructions for running the GitHub Action, downloading the artifact, and testing it on Windows.
- Test with existing `tests/test_web_ui.py` and `samples/asistencia_ejemplo.xlsx`; add no new business-logic tests because this work wraps the existing local web app rather than changing calculations.

## Task 1: Desktop Entrypoint

**Files:**
- Create: `desktop_app.py`
- Modify: `requirements.txt`

- [ ] **Step 1: Add desktop dependencies**

Append these lines to `requirements.txt`:

```text
pywebview
pyinstaller
```

- [ ] **Step 2: Create the desktop app entrypoint**

Create `desktop_app.py` with this implementation:

```python
from __future__ import annotations

import socket
import threading
import time
import urllib.request
from pathlib import Path

import uvicorn
import webview


APP_NAME = "Prenomina"
HOST = "127.0.0.1"
PROJECT_ROOT = Path(__file__).resolve().parent


def find_free_port() -> int:
    with socket.socket(socket.AF_INET, socket.SOCK_STREAM) as sock:
        sock.bind((HOST, 0))
        return int(sock.getsockname()[1])


def wait_for_health(port: int, timeout_seconds: float = 20.0) -> None:
    deadline = time.monotonic() + timeout_seconds
    url = f"http://{HOST}:{port}/health"
    last_error: Exception | None = None
    while time.monotonic() < deadline:
        try:
            with urllib.request.urlopen(url, timeout=1.0) as response:
                if response.status == 200:
                    return
        except Exception as exc:
            last_error = exc
            time.sleep(0.25)
    raise RuntimeError(f"No se pudo iniciar {APP_NAME}. Ultimo error: {last_error}")


def run_server(port: int, server_holder: dict[str, uvicorn.Server]) -> None:
    config = uvicorn.Config(
        "app.main:app",
        host=HOST,
        port=port,
        log_level="warning",
        reload=False,
    )
    server = uvicorn.Server(config)
    server_holder["server"] = server
    server.run()


def main() -> None:
    port = find_free_port()
    server_holder: dict[str, uvicorn.Server] = {}
    server_thread = threading.Thread(
        target=run_server,
        args=(port, server_holder),
        daemon=True,
    )
    server_thread.start()
    wait_for_health(port)

    window = webview.create_window(
        APP_NAME,
        f"http://{HOST}:{port}",
        width=1180,
        height=820,
        min_size=(980, 680),
    )

    def stop_server() -> None:
        server = server_holder.get("server")
        if server is not None:
            server.should_exit = True

    window.events.closed += stop_server
    webview.start()


if __name__ == "__main__":
    main()
```

- [ ] **Step 3: Run a syntax check**

Run:

```bash
python3 -m py_compile desktop_app.py
```

Expected: command exits with code 0 and prints no errors.

- [ ] **Step 4: Install dependencies locally only if needed for smoke test**

If local Mac environment has a virtualenv ready:

```bash
python3 -m pip install -r requirements.txt
```

Expected: dependencies install. If the system Python cannot install packages, skip local pywebview smoke test and rely on GitHub Actions Windows build.

- [ ] **Step 5: Commit**

```bash
git add desktop_app.py requirements.txt
git commit -m "feat: add desktop app entrypoint"
```

## Task 2: Windows Installer Script

**Files:**
- Create: `installer/prenomina.iss`

- [ ] **Step 1: Create installer folder**

Run:

```bash
mkdir -p installer
```

- [ ] **Step 2: Add Inno Setup script**

Create `installer/prenomina.iss`:

```ini
#define MyAppName "Prenomina"
#define MyAppVersion "0.1.0"
#define MyAppPublisher "Humanio"
#define MyAppExeName "Prenomina.exe"

[Setup]
AppId={{8B18F3E0-2255-4676-A710-6B2E5EB9C780}
AppName={#MyAppName}
AppVersion={#MyAppVersion}
AppPublisher={#MyAppPublisher}
DefaultDirName={autopf}\{#MyAppName}
DefaultGroupName={#MyAppName}
DisableProgramGroupPage=yes
OutputDir=output
OutputBaseFilename=Prenomina Setup
Compression=lzma
SolidCompression=yes
WizardStyle=modern
ArchitecturesAllowed=x64
ArchitecturesInstallIn64BitMode=x64

[Languages]
Name: "spanish"; MessagesFile: "compiler:Languages\Spanish.isl"

[Tasks]
Name: "desktopicon"; Description: "Crear icono en el escritorio"; GroupDescription: "Accesos directos:"; Flags: checkedonce

[Files]
Source: "..\dist\Prenomina.exe"; DestDir: "{app}"; Flags: ignoreversion
Source: "..\.env.example"; DestDir: "{userdocs}\Prenomina"; DestName: ".env.example"; Flags: ignoreversion

[Dirs]
Name: "{userdocs}\Prenomina"
Name: "{userdocs}\Prenomina\input"
Name: "{userdocs}\Prenomina\output"

[Icons]
Name: "{group}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"
Name: "{autodesktop}\{#MyAppName}"; Filename: "{app}\{#MyAppExeName}"; Tasks: desktopicon

[Run]
Filename: "{app}\{#MyAppExeName}"; Description: "Abrir {#MyAppName}"; Flags: nowait postinstall skipifsilent
```

- [ ] **Step 3: Commit**

```bash
git add installer/prenomina.iss
git commit -m "build: add windows installer script"
```

## Task 3: GitHub Actions Windows Build

**Files:**
- Create: `.github/workflows/build-windows.yml`

- [ ] **Step 1: Create workflow folder**

Run:

```bash
mkdir -p .github/workflows
```

- [ ] **Step 2: Add manual Windows build workflow**

Create `.github/workflows/build-windows.yml`:

```yaml
name: Build Windows Installer

on:
  workflow_dispatch:

jobs:
  build-windows:
    runs-on: windows-latest

    steps:
      - name: Checkout repository
        uses: actions/checkout@v4

      - name: Set up Python
        uses: actions/setup-python@v5
        with:
          python-version: "3.12"

      - name: Install Python dependencies
        shell: pwsh
        run: |
          python -m pip install --upgrade pip
          pip install -r requirements.txt

      - name: Run tests
        shell: pwsh
        run: |
          pytest -q

      - name: Build desktop executable
        shell: pwsh
        run: |
          pyinstaller --name Prenomina --onefile --windowed --collect-all webview desktop_app.py

      - name: Install Inno Setup
        shell: pwsh
        run: |
          choco install innosetup -y

      - name: Build installer
        shell: pwsh
        run: |
          & "C:\Program Files (x86)\Inno Setup 6\ISCC.exe" installer\prenomina.iss

      - name: Upload installer artifact
        uses: actions/upload-artifact@v4
        with:
          name: Prenomina-Windows-Installer
          path: installer/output/*.exe
```

- [ ] **Step 3: Commit**

```bash
git add .github/workflows/build-windows.yml
git commit -m "ci: add windows installer build workflow"
```

## Task 4: Build Instructions

**Files:**
- Create: `docs/windows_desktop_build.md`

- [ ] **Step 1: Add operator documentation**

Create `docs/windows_desktop_build.md`:

```markdown
# Build Windows Installer

This project builds the Windows desktop installer with GitHub Actions.

## Run the build

1. Push the latest code to GitHub.
2. Open the repository on GitHub.
3. Go to Actions.
4. Select Build Windows Installer.
5. Click Run workflow.
6. Wait for the run to finish.
7. Download the artifact named Prenomina-Windows-Installer.

## Output

The artifact contains:

```text
Prenomina Setup.exe
```

## Test checklist

- Install on a Windows machine without Python installed.
- Open Prenomina from the desktop icon.
- Confirm the app window opens.
- Open `http://127.0.0.1:<dynamic-port>/health` is not required for the user, but the app should load the local web UI.
- Upload `samples/asistencia_ejemplo.xlsx`.
- Confirm the result screen shows employees, total, warnings, and critical issues.
- Download the generated Excel.
- Close the app and confirm the local server stops.

## Notes

The desktop app runs locally. Railway is not used for normal desktop operation.
The `/procesar-prenomina` API path still requires model credentials if used.
The web form path `/` uses the deterministic local calculation engine.
```

- [ ] **Step 2: Commit**

```bash
git add docs/windows_desktop_build.md
git commit -m "docs: add windows desktop build instructions"
```

## Task 5: GitHub Build Verification

**Files:**
- No code changes expected unless workflow fails.

- [ ] **Step 1: Push branch**

Run:

```bash
git push origin HEAD
```

Expected: branch appears on GitHub.

- [ ] **Step 2: Run workflow manually**

In GitHub:

```text
Repository -> Actions -> Build Windows Installer -> Run workflow
```

Expected: workflow starts on `windows-latest`.

- [ ] **Step 3: Verify tests pass in Windows**

Expected GitHub Actions log:

```text
pytest -q
...
8 passed
```

If the number of tests differs, verify there are no failures.

- [ ] **Step 4: Verify installer artifact**

Expected artifact:

```text
Prenomina-Windows-Installer
```

It should contain:

```text
Prenomina Setup.exe
```

- [ ] **Step 5: If build fails because pywebview hidden imports are missing**

Update the PyInstaller command in `.github/workflows/build-windows.yml` to:

```powershell
pyinstaller --name Prenomina --onefile --windowed --collect-all webview --hidden-import=webview.platforms.edgechromium desktop_app.py
```

Then commit:

```bash
git add .github/workflows/build-windows.yml
git commit -m "ci: include pywebview windows hidden imports"
git push origin HEAD
```

## Task 6: Windows Acceptance Test

**Files:**
- No code changes expected unless acceptance test fails.

- [ ] **Step 1: Install on Windows**

Download `Prenomina Setup.exe` from GitHub Actions and run it on a Windows test machine.

Expected:

- Installer completes.
- Desktop icon `Prenomina` is created if selected.
- Start menu entry `Prenomina` is created.

- [ ] **Step 2: Launch app**

Double-click `Prenomina`.

Expected:

- A desktop window opens.
- The web UI title says `Generador de prenomina`.
- No terminal window is required for the accountant.

- [ ] **Step 3: Process sample Excel**

Use:

```text
samples/asistencia_ejemplo.xlsx
```

Expected:

- Result message says prenomina was calculated.
- Employee count is `2`.
- Total is `$2,480.00`.
- Download button appears.

- [ ] **Step 4: Close app**

Close the desktop window.

Expected:

- App closes.
- Reopening the desktop icon starts a fresh local server.

- [ ] **Step 5: Record release notes**

Create a short note with:

```text
Windows version tested:
Installer filename:
Build run URL:
Sample file result:
Known issues:
```

Commit only if this note is stored in the repo:

```bash
git add docs/windows_desktop_build.md
git commit -m "docs: record windows installer acceptance notes"
```

## Self-Review

- Spec coverage: covers desktop wrapper, Windows executable, installer, GitHub Actions artifact, docs, and Windows acceptance testing.
- Placeholder scan: no TBD/TODO placeholders remain.
- Type consistency: `desktop_app.py`, `Prenomina.exe`, `installer/prenomina.iss`, and `build-windows.yml` names match across tasks.
