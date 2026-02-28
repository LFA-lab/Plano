"""
Met à jour sitemark.html avec la date de build de l'exe (dist/Sitemark.exe).
À lancer après PyInstaller pour que la page affiche le bon build.
"""
import os
import sys
from datetime import datetime

PLACEHOLDER = "__EXE_BUILD_DATE__"
SCRIPT_DIR = os.path.dirname(os.path.abspath(__file__))
HTML_PATH = os.path.join(SCRIPT_DIR, "sitemark.html")
EXE_PATH = os.path.join(SCRIPT_DIR, "dist", "Sitemark.exe")


def main():
    if not os.path.isfile(EXE_PATH):
        print("Exe non trouvé :", EXE_PATH)
        return 1
    if not os.path.isfile(HTML_PATH):
        print("HTML non trouvé :", HTML_PATH)
        return 1

    mtime = os.path.getmtime(EXE_PATH)
    build_date = datetime.fromtimestamp(mtime).strftime("%d/%m/%Y %H:%M")

    with open(HTML_PATH, "r", encoding="utf-8") as f:
        content = f.read()
    if PLACEHOLDER not in content:
        print("Placeholder", PLACEHOLDER, "absent de sitemark.html")
        return 1
    content = content.replace(PLACEHOLDER, build_date)

    with open(HTML_PATH, "w", encoding="utf-8") as f:
        f.write(content)
    print("sitemark.html mis à jour : Build de l'exe", build_date)
    return 0


if __name__ == "__main__":
    sys.exit(main())
