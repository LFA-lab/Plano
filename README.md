# 1. Pars propre depuis main
git checkout main
git pull origin main

# 2. Crée ta branche de travail (change le nom évidemment)
git checkout -b feature-x

# 3. Tu fais tes modifs de code...

# 4. Ajoute et commit tes changements
git add .
git commit -m "implémentation de X (décris ce que tu as fait)"

# 5. Pousse ta branche sur GitHub
git push origin feature-x

# 6. Crée automatiquement la Pull Request vers main
gh pr create --base main --head feature-x --title "Feature X" --body "Description rapide de ce que fait cette feature."

#   À ce stade :
#   - La branche feature-x existe sur GitHub
#   - La PR feature-x → main est ouverte
#   - Tu peux aller sur GitHub et cliquer "Merge pull request"
#   (Pas besoin de faire quoi que ce soit d'autre en local pour ouvrir la PR)
#
# 7. Après avoir mergé la PR sur GitHub, tu synchronises ton main local
git checkout main
git pull origin main

# 8. Nettoyage local
git branch -d feature-x

# 9. Nettoyage distant (si la branche n'a pas déjà été supprimée par GitHub)
git push origin --delete feature-x


# après tes modifs
git add .
git commit -m "description claire de la nouvelle tâche"

# envoie la branche distante
git push -u origin feature/nouvelle-tache

# ouvre la PR automatiquement
gh pr create --base main --head feature/nouvelle-tache \
  --title "Titre de la nouvelle tâche" \
  --body "Ce que cette branche change, pourquoi, impact."
