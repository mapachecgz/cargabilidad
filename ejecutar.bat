@echo off
setlocal DisableDelayedExpansion
pushd "%~dp0"
set "CARGA_INPUT=%~1"
set "CARGA_OUTPUT=%~2"
set "CARGA_DATE=%~3"
set "CARGA_CONFIG=%~4"
if not defined CARGA_INPUT goto ask_input
goto input_ready
:ask_input
echo Directorio de entrada: debe contener exactamente un CSV.
set /p "CARGA_INPUT=Entrada [entrada]: "
if not defined CARGA_INPUT set "CARGA_INPUT=%~dp0entrada"
:input_ready
if not defined CARGA_OUTPUT goto ask_output
goto output_ready
:ask_output
set /p "CARGA_OUTPUT=Salida [salida]: "
if not defined CARGA_OUTPUT set "CARGA_OUTPUT=%~dp0salida"
:output_ready
if not defined CARGA_DATE goto ask_date
goto date_ready
:ask_date
set /p "CARGA_DATE=Fecha de corte YYYY-MM-DD [hoy]: "
:date_ready
if not defined CARGA_CONFIG set "CARGA_CONFIG=%~dp0config"
set "CARGA_PYTHON=python"
python --version >nul 2>&1
if not errorlevel 1 goto python_ready
set "CARGA_PYTHON=py -3"
py -3 --version >nul 2>&1
if not errorlevel 1 goto python_ready
echo ERROR: Python 3 no esta instalado o no esta en PATH.
set "CARGA_EXIT=1"
goto finish
:python_ready
%CARGA_PYTHON% -c "import openpyxl" >nul 2>&1
if not errorlevel 1 goto dependencies_ready
echo ERROR: falta openpyxl. Ejecuta: %CARGA_PYTHON% -m pip install -r requirements.txt
set "CARGA_EXIT=1"
goto finish
:dependencies_ready
if defined CARGA_DATE goto run_with_date
%CARGA_PYTHON% src/generar_reporte.py --entrada "%CARGA_INPUT%" --directorio-salida "%CARGA_OUTPUT%" --config "%CARGA_CONFIG%"
goto capture_exit
:run_with_date
%CARGA_PYTHON% src/generar_reporte.py --entrada "%CARGA_INPUT%" --directorio-salida "%CARGA_OUTPUT%" --config "%CARGA_CONFIG%" --fecha-corte "%CARGA_DATE%"
:capture_exit
set "CARGA_EXIT=%ERRORLEVEL%"
if "%CARGA_EXIT%"=="0" echo Reportes generados correctamente. Las rutas se muestran arriba.
if not "%CARGA_EXIT%"=="0" echo ERROR: no fue posible generar los reportes. Revisa el mensaje anterior.
:finish
if "%~1"=="" pause
popd
exit /b %CARGA_EXIT%
