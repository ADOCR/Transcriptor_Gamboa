@echo off
setlocal

cd /d "%~dp0\.."

echo ===============================================
echo Gamboa Transcriptor - Compilacion .exe
echo ===============================================
echo.

python -m pip install --upgrade pip
python -m pip install faster-whisper python-docx requests pyinstaller

python -c "from PySide6.QtCore import Qt; print('PySide6 OK')" || (
  echo.
  echo ERROR: PySide6 no carga en este entorno.
  echo Use el entorno recomendado:
  echo   conda env create -f environment.yml
  echo   conda activate gamboa-transcriptor
  echo   build_scripts\build_exe.bat
  pause
  exit /b 1
)

if exist build rmdir /s /q build
if exist dist\GamboaTranscriptor rmdir /s /q dist\GamboaTranscriptor

pyinstaller ^
  --noconsole ^
  --onedir ^
  --clean ^
  --name "GamboaTranscriptor" ^
  --icon "assets\icon.ico" ^
  --add-data "assets;assets" ^
  --collect-all faster_whisper ^
  --collect-all ctranslate2 ^
  --collect-all av ^
  --hidden-import docx ^
  main.py

echo.
echo Compilacion finalizada.
echo Ejecutable:
echo dist\GamboaTranscriptor\GamboaTranscriptor.exe
pause
