# Plano - Pilotage Chantier

## Qu'est-ce que Plano ?

Plano transforme des données terrain (tâches, quantités, heures) en tableaux de bord d'avancement financier et physique via Microsoft Project.

## Workflow

1. Client télécharge `ModèleImport.mpt` depuis la page d'onboarding
2. Client ouvre le .mpt dans MS Project
3. Macro génère le fichier Excel template (ou client utilise `FichierTypearemplir.xlsx`)
4. Client remplit le Excel avec ses données projet (colonnes A:L)
5. Client exécute la macro d'import (Alt+F8 → Import_Taches_Simples_AvecTitre_OPTIMISE)
6. Planning généré automatiquement
7. Export JSON → Dashboard HTML pour visualisation

## Structure du repo

```
/templates     - Fichiers template (.mpt, .xlsx)
/macros        - Macros VBA organisées par fonction
  /import      - Import Excel → MS Project
  /export      - Export vers JSON
  /reports     - Génération rapports Word/PNG
  /utils       - Utilitaires divers
/scripts       - Scripts PowerShell (build .mpt)
/dashboard     - Dashboard HTML
/onboarding    - Page d'onboarding client
/docs          - Documentation
/samples       - Fichiers de test
/_archive      - Anciennes versions
```

## Prérequis

- Microsoft Project 2019+
- Microsoft Excel 2019+
- Macros VBA activées

## Colonnes Excel attendues (import)

| Colonne | Contenu |
|---------|---------|
| A | Nom de la tâche |
| B | Quantités |
| C | Nb personnes |
| D | Heures |
| E | Zone |
| F | Sous-Zone |
| G | Tranche |
| H | Lot |
| I | Entreprise |
| J | Qualité |
| K | Niveau |
| L | Onduleur |

## Mise à jour des macros

Après modification d'un fichier .vb dans `/macros` :

```powershell
./scripts/build_mpt.ps1
```

Cela injecte automatiquement les macros dans `ModèleImport.mpt`.
