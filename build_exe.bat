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

SET ADD7=--add-data "data\tabla_de_relacion.json;data\tabla_de_relacion.json"
REM Detectar automáticamente la DLL de Python usada por el intérprete y agregarla al bundle
FOR /F "usebackq delims=" %%p IN (`py -c "import sys,os; print(os.path.join(os.path.dirname(sys.executable), f'python{sys.version_info.major}{sys.version_info.minor}.dll'))"`) DO SET PY_DLL=%%p
IF EXIST "%PY_DLL%" (
    ECHO Including Python DLL: %PY_DLL%
    SET ADD6=--add-binary "%PY_DLL%;."
) ELSE (
    ECHO WARNING: Python DLL not found at %PY_DLL% - you may need to add it manually with --add-binary
    SET ADD6=
)


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
REM Ensure script runs from its own directory so relative paths work
pushd "%~dp0" >nul

REM Allow user to pass extra --add-data entries via environment variable EXTRA_ADDS
IF NOT DEFINED EXTRA_ADDS SET EXTRA_ADDS=

REM Run PyInstaller with quoted arguments to handle spaces
py -m PyInstaller --noconfirm %MODE_FLAG% %WINDOW_FLAG% --icon "img\icono.ico" --name "%EXE_NAME%" %HIDDEN_IMPORTS% %ADD1% %ADD2% %ADD3% %ADD4% %ADD5% %ADD6% %ADD7% %ADD8% %EXTRA_ADDS% "%MAIN%"

REM Restore original directory
popd >nul

echo.
echo Build finished. Review the "dist\%EXE_NAME%" folder (for --onedir) or "dist\%EXE_NAME%.exe" (for --onefile).
echo If the application fails at runtime, rebuild with the 'debug' argument to keep the console open:
echo    build_exe.bat debug
pause