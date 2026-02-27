#!/bin/bash
# À exécuter depuis la racine du repo (WSL/Linux) pour retirer du suivi Git
# les fichiers Zone.Identifier dont le chemin contient ":" (invalide sur Windows).
# Usage: depuis la racine du repo: bash macros/export/sitemark_exe/fix_zone_identifier_in_repo.sh

set -e
cd "$(git rev-parse --show-toplevel)"

echo "Recherche des fichiers avec ':' dans le chemin..."
# Lister les fichiers Zone.Identifier (nom avec ":" invalide sur Windows)
FILES=$(git ls-files | grep -i 'Zone\.Identifier' || true)
if [ -z "$FILES" ]; then
  echo "Aucun fichier avec ':' dans le chemin trouvé dans l'index."
  exit 0
fi

echo "Fichiers à retirer du suivi Git:"
echo "$FILES"
# Lecture ligne par ligne (gère les espaces dans les noms)
echo "$FILES" | while IFS= read -r f; do
  [ -n "$f" ] && git rm --cached "$f" 2>/dev/null || true
done
# Chemins exacts rapportés par le runner Windows (au cas où)
git rm --cached "macros/export/modal_app.py:Zone.Identifier" 2>/dev/null || true
git rm --cached "samples/dashboard_mecaelec 1.html:Zone.Identifier" 2>/dev/null || true
echo "Fait. Puis exécuter:"
echo "  git add .gitignore"
echo "  git commit -m 'chore: remove Zone.Identifier and other colon paths for Windows CI'"
echo "  git push"
