@echo off
REM Build script for PyInstaller (adjusts flags for robust packaging)
REM Usage: build_exe.bat [debug]

REM Use a shell-safe name for the generated executable (avoid &)
SET EXE_NAME=Sistema_Generador_Documentos_VC
SET EXE_DISPLAY_NAME="Sistema Generador de Documentos V&C"
SET MAIN=app.py

REM Choose build mode: onedir (recommended) or onefile
SET BUILD_MODE=onedir

REM Set WINDOWED=1 to hide console (final builds). Set WINDOWED=0 to keep console (debug).
SET WINDOWED=1

IF "%1"=="debug" (
    ECHO Debug mode: enabling console and BUILD_MODE=onedir
    SET WINDOWED=0
    SET BUILD_MODE=onedir
)

REM Ensure PyInstaller is available (install if missing)
REM Install requirements to ensure runtime libs are available for the bundle
echo Installing Python requirements (this may take a few minutes)...
py -m pip install -r requirements.txt

py -m pip show pyinstaller >nul 2>&1 || (
    echo Installing PyInstaller...
    py -m pip install pyinstaller --upgrade
)

REM Common hidden imports that PyInstaller sometimes misses for this project
SET HIDDEN_IMPORTS=--hidden-import=fitz --hidden-import=docx --hidden-import=PIL

REM We'll pass each --add-data explicitly to avoid batch quoting issues
REM Data folders to include next to the exe (Windows syntax: src;dest)
REM Add or remove entries below as needed
SET ADD1=--add-data "data;data"
SET ADD2=--add-data "Documentos Inspeccion;Documentos Inspeccion"
SET ADD3=--add-data "Pegado de Evidenvia Fotografica;Pegado de Evidenvia Fotografica"
SET ADD4=--add-data "Firmas;Firmas"
SET ADD5=--add-data "img;img"


REM Construct mode flags
IF "%BUILD_MODE%"=="onefile" (
    SET MODE_FLAG=--onefile
) ELSE (
    SET MODE_FLAG=--onedir
)

REM Windowed vs console
IF "%WINDOWED%"=="1" (
    SET WINDOW_FLAG=--windowed
) ELSE (
    SET WINDOW_FLAG=--console
)

REM Run PyInstaller using the same Python interpreter (py launcher)
echo Building %EXE_NAME% (%MODE_FLAG%, %WINDOW_FLAG%) ...
py -m PyInstaller --noconfirm %MODE_FLAG% %WINDOW_FLAG% --icon "img\icono.ico" --name %EXE_NAME% %HIDDEN_IMPORTS% %ADD1% %ADD2% %ADD3% %ADD4% %ADD5% %ADD6% %ADD7% %MAIN%

echo.
echo Build finished. Review the "dist\%EXE_NAME%" folder (for --onedir) or "dist\%EXE_NAME%.exe" (for --onefile).
echo If the application fails at runtime, rebuild with the 'debug' argument to keep the console open:
echo    build_exe.bat debug
pause