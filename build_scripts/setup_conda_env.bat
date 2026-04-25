@echo off
setlocal

cd /d "%~dp0\.."

echo ===============================================
echo Gamboa Transcriptor - Crear entorno Anaconda
echo ===============================================
echo.
echo Este script crea un entorno separado para evitar conflictos
echo de DLLs/Qt en el entorno base de Anaconda.
echo.

conda env create -f environment.yml

echo.
echo Entorno creado.
echo Para activarlo:
echo conda activate gamboa-transcriptor
echo.
echo Para ejecutar:
echo python main.py
pause
