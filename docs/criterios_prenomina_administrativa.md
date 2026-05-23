# Criterios de prenomina administrativa

Este documento describe los criterios actuales del motor de prenomina administrativa. El objetivo es que el calculo sea auditable: el sistema no solo genera importes, tambien muestra que valores uso, que diferencias encontro contra Excel y que criterios siguen pendientes de validacion con el cliente.

## Memoria de conversacion y proyecto

Repositorio del agente: `mangoex/prenomina-agent`.

URL Railway actual: `https://web-production-7b4d.up.railway.app/`.

Decision principal: el agente debe servir dos modos de uso. El modo API se mantiene para n8n/integraciones con `POST /procesar-prenomina` protegido por `X-API-Key`. El modo usuario final vive en `/` y permite subir el Excel, validar UMA/factor/modo de tope, generar la prenomina y descargar el resultado sin que un usuario no tecnico tenga que manejar headers o comandos.

Proveedor de modelo: el proyecto usa OpenRouter en produccion, no OpenAI directo. `OPENROUTER_API_KEY` y `PRENOMINA_API_KEY` son llaves distintas: OpenRouter sirve para el proveedor de modelo del agente; `PRENOMINA_API_KEY` protege el endpoint API. La pantalla web ejecuta el motor deterministico directo para reducir friccion del usuario y conservar la API para automatizaciones.

Cambios ya publicados:

- PR #1: validaciones auditables de prenomina administrativa, fondo de ahorro, UMA, SDI, diferencias contra Excel y `runtime.txt`.
- PR #2: pantalla web para subir Excel, validar criterios y descargar resultado.
- PR #3: correccion de enlaces de descarga web con rutas relativas para Railway.

Estado validado:

- `/health` responde correctamente en Railway.
- `/` muestra la pantalla web.
- `/web/procesar-prenomina` procesa el Excel sintetico/de muestra y entrega boton de descarga.
- `/procesar-prenomina` sigue protegido por `X-API-Key` y queda reservado para API/n8n.

Variable recomendada para la pantalla web:

```env
WEB_ACCESS_KEY=una_clave_para_el_formulario
```

Si `WEB_ACCESS_KEY` existe, el formulario y la descarga web piden esa clave. Si no existe, la pantalla web queda abierta.

## Campos leidos del Excel

El motor detecta columnas administrativas como `No.`, `Cod.`, `Empresa`, `Nombre completo`, `Area`, `Departamento`, `Puesto`, `Lugar de trabajo`, `Fecha de ingreso`, `Cuenta con fondo de ahorro`, `Salario diario`, `Factor integracion IMSS`, `Salario diario integrado`, `Sueldo nominal`, `Puntualidad`, `Asistencia`, `Vales de despensa`, `Fondo de ahorro`, `Percepcion sueldos`, `Honorarios asimilados`, `Gasolina sueldo`, `Socio`, `Efectivo`, `Facturado`, `Deuda de carro`, `Sueldo bruto mensual`, `Sueldo bruto quincenal`, `Descuento`, `Sueldo bruto mensual final`, `Sueldo bruto quincenal final` y `Observaciones`.

## Campos recalculados

El sistema recalcula y compara contra Excel:

- Salario diario integrado.
- Fondo de ahorro parametrizado.
- Percepcion sueldos.
- Sueldo bruto mensual.
- Sueldo bruto quincenal.
- Sueldo bruto mensual final.
- Sueldo bruto quincenal final.
- Descuentos por observaciones tipo `DESCONTAR X DIA`.

## Salario diario integrado

El documento de negocio indica:

```text
FI = 1 + (Aguinaldo / 365) + [(Vacaciones x Prima vacacional) / 365]
SDI = Salario diario x Factor de integracion IMSS
```

Como el Excel ya trae `Salario diario`, `Factor integracion IMSS` y `Salario diario integrado`, el motor valida:

```text
salario_diario_integrado_calc = salario_diario x factor_integracion_imss
```

Si la diferencia contra `Salario diario integrado` de Excel supera `0.03`, se agrega una advertencia `diferencia_sdi`.

## Percepcion sueldos

La percepcion de sueldos se calcula como:

```text
percepcion_sueldos = sueldo_nominal + puntualidad + asistencia + vales_despensa + fondo_ahorro_para_calculo
```

Mientras `USE_EXCEL_FONDO_AHORRO=true`, `fondo_ahorro_para_calculo` conserva el valor de Excel. Cuando `USE_EXCEL_FONDO_AHORRO=false`, usa el calculo propio `fondo_ahorro_calc`.

## Sueldo bruto

El sueldo bruto mensual se calcula como:

```text
sueldo_bruto_mensual =
  percepcion_sueldos
  + honorarios_asimilados
  + gasolina_sueldo
  + socio
  + efectivo
  + facturado
  + deuda_carro
```

El sueldo bruto quincenal se calcula como:

```text
sueldo_bruto_quincenal = sueldo_bruto_mensual / 2
```

El sueldo bruto mensual final descuenta `Descuento`:

```text
sueldo_bruto_mensual_final = sueldo_bruto_mensual - descuento
sueldo_bruto_quincenal_final = sueldo_bruto_mensual_final / 2
```

## Observaciones y descuentos por dias

Si `OBSERVACIONES` contiene un texto como `DESCONTAR X DIA`, el motor interpreta `X` como dias a descontar sobre la quincena:

```text
descuento_por_dias = (sueldo_bruto_mensual_final / 2) / 15 x dias
sueldo_bruto_quincenal_final = (sueldo_bruto_mensual_final / 2) - descuento_por_dias
```

Observaciones sobre prestamos, adeudos, comisiones, fondo de ahorro o pagos por fecha se reportan como advertencias para revision humana.

## Fondo de ahorro

El Excel original puede traer una formula externa parecida a:

```text
=SI(L6="Si",P6*'/srv-conta/00..archivos/Prestaciones/[0.BASE DE DATOS LAFHER NUEVA 2023.xlsx]UMA Y SMG'!$F$26,0)
```

En formato ingles, `openpyxl` puede verla como:

```text
=IF(L6="Si",P6*'[1]UMA Y SMG'!$F$26,0)
```

El sistema no puede leer el archivo externo `[0.BASE DE DATOS LAFHER NUEVA 2023.xlsx]UMA Y SMG!$F$26`. Por eso usa variables de entorno auditables:

```env
FONDO_AHORRO_FACTOR=0.11
UMA_DIARIA=117.31
FONDO_AHORRO_TOPE_MODE=mensual
USE_EXCEL_FONDO_AHORRO=true
```

La validacion base es:

```text
si cuenta_fondo_ahorro = Si:
  fondo_ahorro_base_calc = sueldo_nominal x FONDO_AHORRO_FACTOR
si no:
  fondo_ahorro_base_calc = 0
```

El tope se calcula segun `FONDO_AHORRO_TOPE_MODE`:

```text
none:
  fondo_ahorro_calc = fondo_ahorro_base_calc

mensual:
  tope = 1.3 x UMA_DIARIA x 365 / 12
  fondo_ahorro_calc = min(fondo_ahorro_base_calc, tope)

quincenal:
  tope = 1.3 x UMA_DIARIA x 365 / 24
  fondo_ahorro_calc = min(fondo_ahorro_base_calc, tope)

proporcional:
  tope = (1.3 x UMA_DIARIA x 365 / 12 / dias_mes) x dias_trabajados
```

Si el modo proporcional no tiene `dias_mes` o `dias_trabajados` confiables, el sistema usa modo mensual como fallback y agrega advertencia.

El reporte siempre muestra:

- Factor de fondo de ahorro usado.
- UMA diaria usada.
- Modo de tope usado.
- Valor de fondo de ahorro de Excel.
- Valor base calculado.
- Tope calculado.
- Valor final calculado.
- Diferencia contra Excel.
- Si requiere validacion.

Mientras `USE_EXCEL_FONDO_AHORRO=true`, las diferencias de fondo de ahorro no detienen la prenomina. Se reportan como advertencias. Cuando `USE_EXCEL_FONDO_AHORRO=false`, el calculo principal usa `fondo_ahorro_calc`.

## Puntualidad y asistencia

El documento de negocio dice que los premios por puntualidad y asistencia se calculan como 10% del salario diario integrado. El Excel real puede traer importes mensualizados o calculados con otro criterio, por lo que el motor no cambia todavia el calculo principal.

Solo agrega una validacion exploratoria:

```text
puntualidad_validacion_10_sdi = salario_diario_integrado_calc x 0.10 x DIAS_BASE_PERIODO
asistencia_validacion_10_sdi = salario_diario_integrado_calc x 0.10 x DIAS_BASE_PERIODO
```

`DIAS_BASE_PERIODO` tiene default `30.4`. Las diferencias se reportan como advertencia informativa `validacion_puntualidad_asistencia`.

## Criterios pendientes de validacion con cliente

- Confirmar si el factor externo observado en `[UMA Y SMG]F26` debe ser siempre `0.11` o si cambia por periodo.
- Confirmar la UMA diaria aplicable por ejercicio y su proceso anual de actualizacion.
- Confirmar si el tope de fondo de ahorro debe aplicarse mensual, quincenal o proporcional por dias trabajados.
- Confirmar si `USE_EXCEL_FONDO_AHORRO` puede pasar a `false` una vez validado el criterio.
- Confirmar si puntualidad y asistencia se calculan con SDI diario, sueldo nominal mensual u otro criterio historico del cliente.
- Confirmar como inferir dias trabajados cuando el Excel no trae una columna explicita.
