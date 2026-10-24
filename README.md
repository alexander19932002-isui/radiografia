# Radiografía del sistema bancario chileno

Primera versión local con datos reales del reporte CMF a agosto de 2026: 17 bancos y una fila separada de Sistema Bancario. No necesita cuentas, claves ni servicios de pago para ejecutarse en tu computador.

## Abrir en Windows con VS Code

1. Extrae todo el ZIP y abre la carpeta `radiografia_bancaria` desde VS Code.
2. Abre Terminal > Nueva terminal. Necesitas Python instalado.
3. Ejecuta, una línea a la vez:

```powershell
python -m venv .venv
.venv\Scripts\python.exe -m pip install -r requirements.txt
.venv\Scripts\python.exe -m streamlit run app.py
```

Abre en el navegador la dirección local que muestra la terminal, normalmente http://localhost:8501. Mantén la terminal abierta. Para detener, presiona Ctrl+C. En ejecuciones posteriores solo necesitas la última instrucción. No se usa Live Server para esta aplicación Python.

## Qué contiene

- `app.py`: pantalla, filtros y gráficos.
- `scripts/importar_cmf.py`: extracción validada y actualización de CSV y SQLite.
- `datos/originales/`: reporte CMF original, conservado sin cambios.
- `datos/procesados/bancos_cmf.csv`: tabla para Excel, Python o Power BI, una fila por banco y mes, con el sistema identificado aparte.
- `datos/procesados/trazabilidad.csv`: hoja y celda de origen de cada indicador.
- `datos/procesados/bancos.sqlite`: base SQL con la tabla `indicadores_bancarios`.

El CSV usa coma como separador y punto decimal. En Excel utiliza Datos > Desde texto/CSV y selecciona UTF-8, delimitador coma y configuración regional Inglés (Estados Unidos) para interpretar decimales. Los porcentajes ya están en escala 0–100: 2.33 significa 2.33%, no aplicar formato porcentual de Excel directamente. Los saldos están en millones de CLP.

## Agregar meses permanentemente

Coloca nuevos reportes XLSX del mismo formato en `datos/originales`, sin repetir un mes. Ejecuta:

```powershell
.venv\Scripts\python.exe scripts/importar_cmf.py
.venv\Scripts\python.exe -m streamlit run app.py
```

El importador valida encabezados y detiene la carga si cambia el formato. No descarga datos automáticamente. Las cargas desde la barra lateral solo duran durante esa sesión. La primera versión contiene una foto de agosto de 2026 por banco; los comparativos agregados de otras fechas en la hoja Indicadores Sistema no se convierten en una historia ficticia por banco.

## Definiciones y origen

Fuente: `articles-113883_recurso_1.xlsx`, título “REPORTE DE INFORMACIÓN FINANCIERA MENSUAL DEL SISTEMA BANCARIO A AGOSTO DE 2026”, entregado por el usuario. CMF, Reportes Financieros Mensuales: https://www.cmfchile.cl/portal/estadisticas/626/w4-propertyvalue-28910.html

Se extraen valores publicados, sin recalcular indicadores:

| Campo | Hoja | Columna | Unidad / alcance |
|---|---|---|---|
| colocaciones_mmclp | Est. Situación Financ. Bancos | D | Millones CLP, total |
| colocaciones_costo_amortizado_mmclp | Est. Situación Financ. Bancos | E | Millones CLP, costo amortizado |
| morosidad_90_pct | Ind. Mora y Deteriorada | C | Porcentaje sobre colocaciones totales |
| provisiones_clasificacion_pct | Ind. R. Crédito Provisiones | C | Porcentaje por clasificación sobre costo amortizado |
| eficiencia_gastos_ingresos_pct | Ind. de Rentab. y Eficiencia | M | Gastos / ingresos operacionales (%) |

Todas las cifras corresponden al alcance consolidado del reporte. Las provisiones mostradas son un índice, no un importe. No incluyen las otras variantes adicionales del reporte. Eficiencia utiliza los resultados acumulados al cierre, no solo el flujo del mes. El sistema es una serie oficial separada: nunca sumar esa fila a los bancos ni promediar tasas para recrearla. No se asignan códigos bancarios porque no se extrajeron del archivo. Solo se retiran llamadas numéricas al pie del nombre, como “(4)”. Se conserva el cero real y no se transforma un dato no disponible en cero.

Nota de la CMF: fusión de Banco Bice y Banco Security materializada el 1 de noviembre de 2025. Debe considerarse si se agregan períodos anteriores. Un ratio de eficiencia más bajo no basta para juzgar la solvencia o calidad global de un banco.

## Ejemplo SQL

```sql
SELECT banco, colocaciones_mmclp, morosidad_90_pct
FROM indicadores_bancarios
WHERE fecha = '2026-08-31' AND tipo_entidad = 'banco'
ORDER BY colocaciones_mmclp DESC;
```
