# Build sitemark.exe (PyInstaller)

## Prérequis

- Python 3 avec venv activé
- Dépendances installées : `pip install -r requirements.txt pyinstaller`
- Fichiers dans ce dossier : `sitemark.py`, `dashboard.py`, `logoomexom.png`, `vue aerienne centrale solaire.png`, `bloc couleur.png`

## Commande complète (à taper dans le terminal)

**Depuis ce dossier** (`macros/export/sitemark_exe`) :

### Windows (cmd / PowerShell)

```cmd
pyinstaller --onefile --hidden-import=pandas --hidden-import=plotly --hidden-import=plotly.express --hidden-import=plotly.graph_objects --hidden-import=openpyxl --hidden-import=openpyxl.cell._writer --hidden-import=numpy --add-data "logoomexom.png;." --add-data "vue aerienne centrale solaire.png;." --add-data "bloc couleur.png;." --name sitemark sitemark.py
```

### Linux / WSL (séparateur `:` pour --add-data)

```bash
pyinstaller --onefile \
  --hidden-import=pandas \
  --hidden-import=plotly \
  --hidden-import=plotly.express \
  --hidden-import=plotly.graph_objects \
  --hidden-import=openpyxl \
  --hidden-import=openpyxl.cell._writer \
  --hidden-import=numpy \
  --add-data "logoomexom.png:." \
  --add-data "vue aerienne centrale solaire.png:." \
  --add-data "bloc couleur.png:." \
  --name sitemark \
  sitemark.py
```

## Script fourni

- **Windows** : `build_sitemark.bat` (double-clic ou `build_sitemark.bat` dans cmd)

## Résultat

- Exécutable : `dist/sitemark.exe` (Windows) ou `dist/sitemark` (Linux)
- Les logos sont inclus dans le .exe ; au lancement, PyInstaller les décompresse dans un dossier temporaire et le script utilise `sys._MEIPASS` pour les charger.

## Si un module manque encore

Ajouter d’autres `--hidden-import=` selon l’erreur, par exemple :

- `--hidden-import=openpyxl.styles`
- `--hidden-import=kaleido` (si export d’images Plotly)
