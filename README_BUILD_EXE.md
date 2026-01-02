Build instructions for creating a single-file Windows executable

Goal
- Create a single .exe named exactly: "Sistema Generador de Documentos V&C.exe"
- The folder `data` must be external (created next to the .exe on first run) so the user can edit it.
- Other asset folders (Firmas, img, Plantillas PDF, Otros archivos, Documentos Inspeccion) are packaged inside the exe and will be available to the app at runtime via the PyInstaller extraction path (no external access required).

Prerequisites
- Windows machine with Python 3.10+ installed and on PATH.
- pip available.

Steps (PowerShell)
1. Open PowerShell in this project folder.
2. Run:

```powershell
.\build_exe.ps1
```

Steps (CMD)
1. Open Command Prompt in this project folder.
2. Run:

```
build_exe.bat
```

What the scripts do
- Ensure PyInstaller is installed (installs if missing).
- Run PyInstaller with `--onefile --windowed` and `--add-data` for the specified folders.
 - Run PyInstaller with `--onefile --windowed`, `--icon img/icono.ico` and `--add-data` for the specified folders.
- Produces a single `Sistema Generador de Documentos V&C.exe` inside `dist\`.

Notes about resource paths
- The application code now distinguishes between two locations:
  - `APP_DIR`: directory where the executable lives (used to create the external `data` folder).
  - `BASE_DIR` / runtime resource directory: when running frozen PyInstaller, bundled files are extracted to a temporary folder; the existing code uses `BASE_DIR` (and `sys._MEIPASS`) to find bundled resources.
- The `data` folder is intentionally NOT bundled: the app will create `data` next to the exe on first run so the user can inspect/edit JSON data.

Security / access
- Bundled folders are included in the exe and aren’t exposed as editable folders on disk by default (they are extracted at runtime to a temporary folder). If you need stronger protection, consider code or resource encryption, but PyInstaller's onefile is generally sufficient for casual bundling.

If you want me to run the build here, tell me and I will attempt to run the PowerShell script (it will install PyInstaller if missing).