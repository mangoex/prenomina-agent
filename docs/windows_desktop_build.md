# Build Windows Installer

Este documento explica como generar el instalador Windows del Agente de Prenomina usando GitHub Actions.

## Que se genera

GitHub Actions construye un archivo descargable:

```text
Prenomina Setup.exe
```

Ese archivo es el instalador que se puede copiar a una computadora Windows.

## Como correr el build en GitHub

1. Sube los cambios del repo a GitHub.
2. Abre el repositorio en GitHub.
3. Entra a la pestana `Actions`.
4. En la lista izquierda, selecciona `Build Windows Installer`.
5. Presiona `Run workflow`.
6. Confirma la rama correcta.
7. Espera a que termine el proceso.

GitHub usara una maquina temporal con Windows para:

- Instalar Python 3.12.
- Instalar dependencias desde `requirements-desktop.txt`.
- Correr pruebas.
- Crear `Prenomina.exe` con PyInstaller.
- Crear `Prenomina Setup.exe` con Inno Setup.
- Publicar el instalador como artifact.

## Como descargar el instalador

1. Entra al run terminado dentro de `Actions`.
2. Baja hasta la seccion `Artifacts`.
3. Descarga `Prenomina-Windows-Installer`.
4. Descomprime el archivo descargado.
5. Dentro estara:

```text
Prenomina Setup.exe
```

## Como instalar en Windows

1. Copia `Prenomina Setup.exe` a la computadora Windows.
2. Da doble clic sobre el instalador.
3. Sigue el asistente.
4. Deja activada la opcion de crear icono en escritorio.
5. Al terminar, abre `Prenomina`.

La app se instala por usuario en una ruta parecida a:

```text
C:\Users\<usuario>\AppData\Local\Programs\Prenomina
```

Los archivos de trabajo quedan en:

```text
C:\Users\<usuario>\Documents\Prenomina
```

## Checklist de prueba en Windows

- Instalar en una computadora Windows sin Python instalado.
- Abrir `Prenomina` desde el icono del escritorio.
- Confirmar que se abre una ventana de escritorio.
- Confirmar que la pantalla muestra `Generador de prenomina`.
- Subir `samples\asistencia_ejemplo.xlsx` si se tiene una copia del repo o un Excel real de prueba.
- Confirmar que aparece resultado.
- Descargar el Excel generado.
- En app de escritorio, si la descarga no abre una ventana del navegador, usar `Abrir carpeta`.
- El Excel generado queda en `C:\Users\<usuario>\Documents\Prenomina\output`.
- Cerrar la app.
- Volver a abrir la app y confirmar que inicia de nuevo.

## Configuracion local

El instalador copia una plantilla:

```text
C:\Users\<usuario>\Documents\Prenomina\.env.example
```

Si se necesita configurar variables, copiar esa plantilla como:

```text
C:\Users\<usuario>\Documents\Prenomina\.env
```

En Windows se puede hacer desde el Explorador de archivos:

1. Abrir `Documentos`.
2. Entrar a la carpeta `Prenomina`.
3. Copiar `.env.example`.
4. Pegar la copia en la misma carpeta.
5. Cambiar el nombre de la copia a `.env`.
6. Abrir `.env` con Bloc de notas.
7. Guardar el archivo y volver a abrir la aplicacion.

Importante: el archivo debe llamarse exactamente `.env`, no `.env.txt`.

Variables comunes:

```env
MAX_UPLOAD_MB=25
ENABLE_AI_ANALYSIS=false
FONDO_AHORRO_FACTOR=0.11
UMA_DIARIA=117.31
FONDO_AHORRO_TOPE_MODE=mensual
USE_EXCEL_FONDO_AHORRO=true
DIAS_BASE_PERIODO=30.4
```

Para uso normal de escritorio, no se necesita Railway.

Para que funcione como agente IA con analisis, comparaciones y sugerencias, configurar:

```env
MODEL_PROVIDER=openrouter
OPENROUTER_API_KEY=tu_clave_openrouter
OPENROUTER_MODEL=openrouter/openai/gpt-4o-mini
PRENOMINA_API_KEY=una_clave_larga_para_la_api
ENABLE_AI_ANALYSIS=true
```

## Notas tecnicas

- La app corre localmente en `127.0.0.1` usando un puerto libre automatico.
- La ventana de escritorio carga la app local con `pywebview`.
- El calculo normal usa `app/tools_prenomina.py`.
- Si `ENABLE_AI_ANALYSIS=true`, la app agrega una capa de analisis IA despues del calculo.
- Railway no participa en la ejecucion local.
- GitHub solo se usa para construir y descargar el instalador.

## Si el build falla

Revisar primero:

- El paso `Run tests`.
- El paso `Build desktop executable`.
- El paso `Build installer`.

Si falla PyInstaller por imports de `pywebview`, revisar el comando del workflow:

```powershell
pyinstaller --name Prenomina --onefile --windowed --collect-all webview --collect-all openpyxl --hidden-import=webview.platforms.edgechromium desktop_app.py
```

Si falla Inno Setup, revisar:

```text
installer/prenomina.iss
```
