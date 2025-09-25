# FalconSolutions Report Generator

Este repositorio contiene una versión en Python del proceso que anteriormente se
realizaba en SQL Server utilizando tablas temporales dentro de `tempdb`.  El
script `report_generator.py` ejecuta la misma lógica en memoria y genera un
archivo Excel con el detalle de las órdenes de compra.

## Requisitos

Instale las dependencias necesarias (idealmente en un entorno virtual):

```bash
pip install -r requirements.txt
```

## Ejecución

Proporcione una cadena de conexión ODBC válida y los parámetros requeridos. Se
pueden definir mediante variables de entorno o flags de línea de comandos.

```bash
export FALCON_SQL_CONNECTION="Driver={ODBC Driver 18 for SQL Server};Server=mi_servidor;Database=FalicTwo;UID=usuario;PWD=clave;"
python report_generator.py --company-id 4849 --group-code BSC --output reporte.xlsx
```

El script obtiene todos los filtros de análisis (`AnalysisFilter`) que contienen
el código de grupo indicado, ejecuta el procedimiento almacenado
`dbo.task_PODetail` para cada filtro y consolida los resultados en un único
archivo Excel.
