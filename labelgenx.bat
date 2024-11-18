@echo off
SET VENV_DIR=lgenv
SET REQUIREMENTS_FILE=requirements.txt

REM Check if the virtual environment directory exists
echo Buscando ambiente virtual...
echo.
IF NOT EXIST "%VENV_DIR%" (
    echo El ambiente virtual necesita ser creado. El proceso puede demorar un poco.
    echo Creando ambiente virtual...
    python -m venv %VENV_DIR%
    echo Intentando inicializar ambiente...
    CALL "%VENV_DIR%\Scripts\activate.bat"
        IF EXIST "%REQUIREMENTS_FILE%" (
        echo.
        echo Instalando requerimientos del programa desde %REQUIREMENTS_FILE%...
        pip install -r %REQUIREMENTS_FILE%
        echo.
        echo Los requerimientos fueron instalados exitosamente.
    ) ELSE (
        echo No se encontro requirements.txt. Creando archivo...
        echo et_xmlfile==2.0.0 >> requirements.txt
        echo openpyxl==3.1.5 >> requirements.txt
        echo Archivo requirements.txt creado exitosamente.
        echo.
        echo Instalando requerimientos del programa desde %REQUIREMENTS_FILE%...
        pip install -r %REQUIREMENTS_FILE%
        echo.
        echo Los requerimientos fueron instalados exitosamente.
    )
) ELSE (
    echo Ambiente virtual localizado. Salteando creacion de ambiente virtual.
)

echo Inicializando ambiente virtual...
CALL "%VENV_DIR%\Scripts\activate.bat"

echo Lanzando programa...
python main.py

CALL deactivate

echo.
echo Presione cualquier tecla para salir...
pause > nul
exit