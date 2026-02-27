# Audit comparatif — Gestion des images de l’onglet Accueil (sitemark.exe)

## Contexte
- **Objectif** : un seul fichier `sitemark.exe` téléchargeable par l’utilisateur final.
- **Images** : `logoomexom.png`, `vue aerienne centrale solaire.png`, `bloc couleur.png` (page de garde).
- **Critères** : simplicité pour l’utilisateur final, robustesse de l’exécutable.

---

## Solution 1 — Fichiers PNG externes (même dossier que l’exe)

| Aspect | Détail |
|--------|--------|
| **Déploiement** | L’utilisateur doit placer les 3 PNG à côté de `sitemark.exe`. |
| **Build** | Aucune option PyInstaller spécifique. |
| **Code** | Lecture via `os.path.join(script_folder, "nom.png")` avec `script_folder = os.path.dirname(sys.executable)` en frozen. |

### Avantages
- Exe plus léger (pas d’images dedans).
- Images modifiables sans recompiler (charte graphique, logos).
- Mise à jour des visuels sans rebuild PyInstaller.
- Pas de dépendance à un dossier temporaire.

### Inconvénients
- **Pas un seul fichier** : 4 éléments à distribuer (exe + 3 PNG).
- Risque d’oubli, de déplacement ou de suppression des PNG par l’utilisateur.
- Chemins sensibles (espaces, accents dans noms de fichiers).
- Certains déploiements (réseau, GPO) compliquent la copie de plusieurs fichiers.
- Antivirus / stratégies de sécurité peuvent bloquer l’écriture ou la lecture du dossier d’installation.

**Verdict** : Simple pour le développeur, moins pour l’utilisateur et peu robuste (un seul fichier non respecté).

---

## Solution 2 — PyInstaller `--add-data` + `sys._MEIPASS`

| Aspect | Détail |
|--------|--------|
| **Déploiement** | Un seul exe ; les images sont extraites au premier lancement dans un dossier temporaire. |
| **Build** | `pyinstaller --add-data "logoomexom.png;." --add-data "vue aerienne centrale solaire.png;." --add-data "bloc couleur.png;." sitemark.py` (sous Windows `;`, sous Linux/OSX `:`). Ou un spec file avec `datas=[('assets', 'assets')]`. |
| **Code** | `script_folder = getattr(sys, '_MEIPASS', os.path.dirname(__file__))` puis lecture des PNG comme en solution 1. |

### Avantages
- Un seul fichier exe à distribuer.
- Images toujours présentes (pas de dépendance au répertoire de l’exe).
- Comportement identique en dev (fichiers à côté du script) et en frozen.

### Inconvénients
- Build plus complexe : plusieurs `--add-data` ou gestion d’un dossier/spec.
- Noms de fichiers avec espaces/accents à échapper ou à mettre dans un spec.
- `_MEIPASS` est un dossier temporaire (nettoyage possible par l’OS, mais lecture seule pendant l’exécution suffit).
- L’exe grossit de la taille des 3 PNG.

**Verdict** : Bon compromis “un seul exe” et maintenabilité des visuels (fichiers séparés dans le repo), au prix d’un build un peu plus soigné.

---

## Solution 3 — Base64 encodé dans le script + `io.BytesIO`

| Aspect | Détail |
|--------|--------|
| **Déploiement** | Un seul exe ; aucune ressource externe. |
| **Build** | Aucune option PyInstaller ; les chaînes base64 sont dans le .py. |
| **Code** | Constantes `B64_LOGO`, `B64_VUE`, `B64_BLOC` ; décodage `base64.b64decode(s)` → `io.BytesIO` → `openpyxl.drawing.image.Image(buf)` → `ws.add_image(img, anchor)`. |

### Avantages
- **Un seul fichier** exe, sans aucun fichier à côté.
- Expérience utilisateur maximale : télécharger l’exe, l’exécuter.
- Aucun chemin disque, aucun risque de fichier manquant ou déplacé.
- Robuste face aux politiques de sécurité (pas d’écriture, pas de dossier “assets” à protéger).
- Build PyInstaller inchangé (pas de `--add-data`).

### Inconvénients
- Le fichier .py (et donc l’exe) grossit de ~33 % par rapport à la taille binaire des PNG (overhead base64).
- Modifier une image impose de ré-encoder en base64 et de mettre à jour le script.
- Le code source contient de longues chaînes (peu lisible en diff).

**Verdict** : Simplicité et robustesse maximales pour l’utilisateur final ; coût en taille et en process de mise à jour des visuels.

---

## Conclusion argumentée

- **Simplicité pour l’utilisateur final** : Solution 3 > Solution 2 > Solution 1 (un seul exe, rien à installer à côté).
- **Robustesse** : Solution 3 (aucune ressource externe) ≥ Solution 2 (ressources dans l’exe) > Solution 1 (fichiers externes fragiles).

**Recommandation : Solution 3 (Base64)** si l’objectif prioritaire est “un seul exe, zéro fichier à gérer” et que les images de la page de garde changent rarement. Le surcoût de taille et la procédure de mise à jour des images (ré-encodage + remplacement des constantes) restent acceptables pour un outil interne ou livré en un seul exe.

**Alternative** : Solution 2 si l’on préfère garder les PNG dans le dépôt et éviter les grosses chaînes base64 dans le code, en acceptant un spec PyInstaller un peu plus complexe.

---

## Méthode exacte : chaîne Base64 → cellule Excel (image)

```python
import base64
import io
from openpyxl.drawing.image import Image as XLImage

def image_from_b64(b64_string, width_px=None, height_px=None):
    """Construit une Image openpyxl à partir d'une chaîne base64 (PNG/JPEG)."""
    if not b64_string or not b64_string.strip():
        return None
    try:
        raw = base64.b64decode(b64_string)
        buf = io.BytesIO(raw)
        buf.seek(0)
        img = XLImage(buf)
        if width_px is not None:
            img.width = width_px
        if height_px is not None:
            img.height = height_px
        return img
    except Exception:
        return None

# Utilisation dans fill_onglet_accueil :
# img = image_from_b64(B64_LOGO, ACCUEIL_LOGO_W, ACCUEIL_LOGO_H)
# if img:
#     ws.add_image(img, "L2")
```

**Chaîne de transformation** :  
`base64_string` → `base64.b64decode()` → `bytes` → `io.BytesIO(bytes)` → `buf.seek(0)` → `XLImage(buf)` → réglage `width`/`height` → `ws.add_image(img, anchor)`.

---

## Génération des constantes base64 (pour Solution 3)

Une fois les PNG prêts, exécuter dans le dossier du script :

```python
import base64
for name, var in [
    ("logoomexom.png", "B64_LOGO"),
    ("vue aerienne centrale solaire.png", "B64_VUE"),
    ("bloc couleur.png", "B64_BLOC"),
]:
    with open(name, "rb") as f:
        b64 = base64.b64encode(f.read()).decode()
    print(f'{var} = """{b64}"""')
```

Copier les sorties dans `sitemark.py` à la place des `B64_LOGO = ""`, etc. Puis rebuild PyInstaller : un seul `sitemark.exe` contiendra les images.
