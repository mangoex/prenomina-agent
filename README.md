# prenomina-agent

MVP local de un agente de prenomina construido con FastAPI, pandas, openpyxl y OpenAI Agents SDK.

El sistema recibe un Excel de asistencia/incidencias, valida datos, calcula una prenomina operativa con reglas deterministicas en Python y genera un archivo en `output/`.

## Que hace el MVP

- Crea un agente unico llamado `Agente de Prenomina`.
- Usa OpenAI Agents SDK para ejecutar la tool `calcular_prenomina_excel`.
- La tool lee el Excel, valida columnas y datos, calcula en Python y genera el archivo final.
- El modelo no inventa calculos: solo llama la tool y resume el resultado.
- Expone una API FastAPI para que luego n8n pueda llamar `POST /procesar-prenomina`.
- Detecta dos formatos: asistencia por horas y prenomina administrativa del cliente.

## Columnas operativas para calcular

Para el MVP, el Excel debe incluir estas columnas o equivalentes:

- `empleado_id`
- `nombre`
- `fecha`
- `horas_trabajadas`
- `sueldo_hora`
- `bono`
- `descuento`

La herramienta tambien reconoce alias comunes como `Cod.`, `Nombre completo`, `Horas trabajadas`, `Sueldo hora`, `Bonos` y `Descuentos`.

## Reglas de calculo MVP

- `total_horas = suma de horas_trabajadas por empleado`
- `subtotal = total_horas * sueldo_hora`
- `total_bonos = suma de bono`
- `total_descuentos = suma de descuento`
- `total_pagar = subtotal + total_bonos - total_descuentos`

Si hay errores criticos, el sistema genera el reporte de inconsistencias pero no calcula total final.

## Formato prenomina administrativa

Para el archivo real del cliente, el sistema detecta columnas como:

- `SUELDO NOMINAL`
- `PUNTUALIDAD`
- `ASISTENCIA`
- `VALES DE DESPENSA`
- `FONDO DE AHORRO`
- `PERCEPCION SUELDOS`
- `HONORARIOS ASIMILADOS`
- `GASOLINA SUELDO`
- `SOCIO`
- `Efectivo`
- `FACTURADO`
- `DEUDA DE CARRO`
- `SUELDO BRUTO MENSUAL`
- `SUELDO BRUTO QUINCENAL`
- `OBSERVACIONES`

Reglas actuales para este formato:

- `percepcion_sueldos = sueldo_nominal + puntualidad + asistencia + vales_despensa + fondo_ahorro_para_calculo`
- `sueldo_bruto_mensual = percepcion_sueldos + honorarios_asimilados + gasolina_sueldo + socio + efectivo + facturado + deuda_carro`
- `sueldo_bruto_quincenal = sueldo_bruto_mensual / 2`
- `sueldo_bruto_mensual_final = sueldo_bruto_mensual - descuento`
- `sueldo_bruto_quincenal_final = sueldo_bruto_mensual_final / 2`
- Si `OBSERVACIONES` dice `DESCONTAR X DIA`, aplica el descuento de dias sobre la quincena.
- Valida `salario_diario_integrado = salario_diario * factor_integracion_imss`.
- Valida Fondo de ahorro contra factor, UMA y tope parametrizables.

Observaciones como prestamos, adeudos, fondo de ahorro, comisiones o pagos por fecha se reportan como advertencias para revision humana.

## Variables de entorno de prenomina administrativa

```env
FONDO_AHORRO_FACTOR=0.11
UMA_DIARIA=117.31
FONDO_AHORRO_TOPE_MODE=mensual
USE_EXCEL_FONDO_AHORRO=true
DIAS_BASE_PERIODO=30.4
```

- `FONDO_AHORRO_FACTOR`: factor usado para validar el fondo de ahorro. Default temporal: `0.11`.
- `UMA_DIARIA`: UMA diaria usada para calcular el tope de fondo de ahorro. Default temporal: `117.31`; requiere validacion y actualizacion anual.
- `FONDO_AHORRO_TOPE_MODE`: modo del tope. Acepta `none`, `mensual`, `quincenal` o `proporcional`. Default: `mensual`.
- `USE_EXCEL_FONDO_AHORRO`: si es `true`, el calculo principal conserva el valor de Excel y solo reporta la diferencia contra el calculo parametrizado. Si es `false`, usa `fondo_ahorro_calc`.
- `DIAS_BASE_PERIODO`: dias usados para la validacion exploratoria de puntualidad y asistencia contra 10% de SDI. Default: `30.4`.

La formula externa del Excel para fondo de ahorro apunta a un archivo que el sistema no puede leer, por ejemplo `[0.BASE DE DATOS LAFHER NUEVA 2023.xlsx]UMA Y SMG!$F$26`. Por eso el agente muestra explicitamente el factor, UMA, tope, valor de Excel, valor calculado y diferencia. El detalle completo esta en `docs/criterios_prenomina_administrativa.md`.

## Preparacion local

```bash
python -m venv .venv
.venv\Scripts\activate
pip install -r requirements.txt
copy .env.example .env
```

Despues edita `.env` y agrega tu proveedor.

Opcion OpenAI directo:

```env
MODEL_PROVIDER=openai
OPENAI_API_KEY=tu_clave
```

Opcion OpenRouter:

```env
MODEL_PROVIDER=openrouter
OPENROUTER_API_KEY=tu_clave_openrouter
OPENROUTER_MODEL=openrouter/openai/gpt-4o-mini
PRENOMINA_API_KEY=una_clave_larga_para_proteger_la_api
MAX_UPLOAD_MB=25
```

OpenRouter se usa a traves de LiteLLM con el extra `openai-agents[litellm]`.

## Correr la API

```bash
uvicorn app.main:app --reload --host 0.0.0.0 --port 8000
```

## Pantalla web para usuario final

La ruta `/` abre una pantalla web simple para usuarios no tecnicos:

- Subir archivo Excel.
- Validar `UMA_DIARIA`, factor de fondo de ahorro y modo de tope.
- Generar la prenomina con el motor deterministico.
- Revisar resumen y advertencias.
- Descargar el Excel generado.

La pantalla web ejecuta el motor deterministico directamente para evitar que el usuario tenga que manejar llaves o headers. El endpoint `/procesar-prenomina` se mantiene separado para integraciones y sigue usando OpenAI Agents SDK con el proveedor configurado, incluyendo OpenRouter.

La pantalla web no requiere que el usuario capture `X-API-Key`. Para protegerla en Railway, configura:

```env
WEB_ACCESS_KEY=una_clave_para_el_formulario
```

Si `WEB_ACCESS_KEY` esta definida, el formulario pedira esa clave y tambien la validara al descargar el archivo. Si no esta definida, la pantalla queda abierta.

Prueba salud:

```bash
curl http://localhost:8000/health
```

Procesar un Excel:

```bash
curl -X POST "http://localhost:8000/procesar-prenomina" ^
  -H "X-API-Key: tu_clave_de_prenomina" ^
  -F "file=@samples/asistencia_ejemplo.xlsx"
```

Si `PRENOMINA_API_KEY` no esta definida, la API queda sin proteccion local. En Railway debe definirse.

El resultado se guarda con nombre unico para evitar bloqueos si el archivo anterior esta abierto:

```text
output/prenomina_resultado_<id>.xlsx
```

La respuesta incluye `download_url`. Para descargar:

```bash
curl -L "https://tu-app.railway.app/descargar/prenomina_resultado_<id>.xlsx" ^
  -H "X-API-Key: tu_clave_de_prenomina" ^
  -o prenomina_resultado.xlsx
```

## Hojas del Excel generado

1. `Base original`
2. `Prenomina`
3. `Inconsistencias`
4. `Diferencias`
5. `Resumen`

## Despliegue en Railway

1. Sube el repositorio a GitHub.
2. Crea un nuevo proyecto en Railway desde el repositorio.
3. Configura las variables de entorno segun el proveedor:
   - OpenAI: `MODEL_PROVIDER=openai` y `OPENAI_API_KEY`.
   - OpenRouter: `MODEL_PROVIDER=openrouter`, `OPENROUTER_API_KEY` y `OPENROUTER_MODEL`.
4. Configura tambien:
   - `PRENOMINA_API_KEY`: llave privada para proteger la API.
   - `WEB_ACCESS_KEY`: llave opcional para proteger la pantalla web.
   - `MAX_UPLOAD_MB`: limite maximo de archivo, por ejemplo `25`.
5. Railway usara el `Procfile`:

```bash
uvicorn app.main:app --host 0.0.0.0 --port $PORT
```

Railway instalara dependencias desde `requirements.txt`.

## Seguridad

- No subas `.env`.
- No subas archivos reales de nomina o prenomina.
- `input/*.xlsx` y `output/*.xlsx` estan excluidos del repositorio.
- `samples/asistencia_ejemplo.xlsx` es sintetico y sirve solo para pruebas.

## Integracion futura con n8n

n8n podra llamar el endpoint:

```text
POST /procesar-prenomina
```

enviando el Excel como `multipart/form-data` con el campo `file`.

Configuracion sugerida en n8n:

- HTTP Method: `POST`
- URL: `https://tu-app.railway.app/procesar-prenomina`
- Header: `X-API-Key: <PRENOMINA_API_KEY>`
- Body Content Type: `multipart/form-data`
- Campo del archivo: `file`
- Despues, usar `download_url` para descargar el Excel generado o enviarlo por correo.
