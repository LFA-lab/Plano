@echo off
title Sitemark — Build EXE
echo.
echo  Construction du fichier Sitemark.exe...
echo.

:: Installer les dependances
python -m pip install pymupdf pdfplumber openpyxl pillow requests pyinstaller -q --disable-pip-version-check

:: Construire l'exe (--onefile = un seul fichier, --noconsole = pas de fenetre noire)
python -m PyInstaller --onefile --noconsole --name "Sitemark" sitemark.py

echo.
if exist "dist\Sitemark.exe" (
    echo  EXE cree : dist\Sitemark.exe
    python inject_exe_build_date.py
    echo  Copiez ce fichier ou vous voulez, il fonctionne seul.
    explorer dist
) else (
    echo  ERREUR : le build a echoue.
)
pause
