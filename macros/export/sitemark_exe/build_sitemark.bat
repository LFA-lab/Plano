@echo off
REM Compilation Sitemark.exe avec PyInstaller (Windows)
REM À lancer depuis ce dossier (macros/export/sitemark_exe)

set SEP=;
pyinstaller --onefile ^
  --hidden-import=pandas ^
  --hidden-import=plotly ^
  --hidden-import=plotly.express ^
  --hidden-import=plotly.graph_objects ^
  --hidden-import=openpyxl ^
  --hidden-import=openpyxl.cell._writer ^
  --hidden-import=numpy ^
  --add-data "logoomexom.png%SEP%." ^
  --add-data "vue aerienne centrale solaire.png%SEP%." ^
  --distpath . ^
  --name Sitemark ^
  sitemark.py

echo.
echo Exécutable : Sitemark.exe (dans ce dossier)
