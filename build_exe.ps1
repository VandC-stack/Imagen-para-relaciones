# PowerShell build script for creating a single-file .exe with PyInstaller
# Usage: Open PowerShell as Administrator (if necessary) and run:
#   .\build_exe.ps1

$exeName = 'Sistema Generador de Documentos V&C'
$main = 'app.py'

# Ensure PyInstaller is installed
try {
    python -m pip show pyinstaller > $null 2>&1
} catch {
    Write-Host "Installing PyInstaller..."
    python -m pip install pyinstaller --upgrade
}

# Build command
$addData = @(
    '"data;data"',
    '"Firmas;Firmas"',
    '"img;img"',
    '"Plantillas PDF;Plantillas PDF"',
    '"Documentos Inspeccion;Documentos Inspeccion"',
    '"Pegado de Evidenvia Fotografica;Pegado de Evidenvia Fotografica"',
    '"etiquetas_generadas;etiquetas_generadas"',
    '"Otros archivos;Otros archivos"',
    '"dictamenes_prueba;dictamenes_prueba"',
    '"tools;tools"'
) -join ' '

$iconPath = "img\icono.ico"
$cmd = "python -m PyInstaller --noconfirm --onefile --windowed --icon `$iconPath` --name `"$exeName`" $addData $main"
Write-Host "Running: $cmd"
Invoke-Expression $cmd

Write-Host "Build finished. Find the exe in the dist\ directory."