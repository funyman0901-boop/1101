@echo off
setlocal
cd /d "%~dp0"
py -3.11 -m venv .venv
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m pip install -r requirements-build.txt
if errorlevel 1 goto fail
.venv\Scripts\python.exe -m PyInstaller --clean --noconfirm --onefile --windowed --name KompasDimensions24 --hidden-import win32timezone --hidden-import win32com.gen_py --collect-submodules win32com.client kompas_dimensions_gui.py
if errorlevel 1 goto fail
echo Ready: dist\KompasDimensions24.exe
pause
exit /b 0
:fail
echo Build failed. See the error above.
pause
exit /b 1
