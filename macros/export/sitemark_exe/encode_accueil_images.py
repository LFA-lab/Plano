"""
Script one-shot : lit les 3 PNG de l'Accueil, les encode en base64,
et met à jour les constantes B64_LOGO, B64_VUE, B64_BLOC dans sitemark.py.
À lancer depuis le dossier sitemark_exe (ou depuis n'importe où : le script
détecte son propre dossier via __file__).
"""
import base64
import os

DIR = os.path.dirname(os.path.abspath(__file__))
FILES = [
    ("logoomexom.png", "B64_LOGO", "  # logoomexom.png"),
    ("vue aerienne centrale solaire.png", "B64_VUE", "   # vue aerienne centrale solaire.png"),
    ("bloc couleur.png", "B64_BLOC", "  # bloc couleur.png"),
]

sitemark_path = os.path.join(DIR, "sitemark.py")
with open(sitemark_path, "r", encoding="utf-8") as f:
    lines = f.readlines()

for filename, var, comment in FILES:
    path = os.path.join(DIR, filename)
    if not os.path.isfile(path):
        print(f"Manquant: {path}")
        continue
    with open(path, "rb") as f:
        b64 = base64.b64encode(f.read()).decode("ascii")
    new_line = f'{var} = """{b64}"""{comment}\n'
    for i, line in enumerate(lines):
        if line.strip().startswith(var + " ="):
            lines[i] = new_line
            break
    print(f"OK: {filename} -> {var} ({len(b64)} car.)")

with open(sitemark_path, "w", encoding="utf-8") as f:
    f.writelines(lines)
print("sitemark.py mis à jour.")
