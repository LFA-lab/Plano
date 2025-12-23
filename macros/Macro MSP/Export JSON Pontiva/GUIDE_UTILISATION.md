# Guide d'utilisation - Export JSON Pontiva & Dashboard

## 📋 Vue d'ensemble

Ce système permet d'exporter un projet MS Project au format JSON et de visualiser les données dans un dashboard web avec 3 vues différentes selon votre rôle.

---

## 🔧 ÉTAPE 1 : Exporter le JSON depuis MS Project

### Installation de la macro

1. **Ouvrir MS Project** avec votre projet actif
2. **Appuyer sur `Alt + F11`** pour ouvrir l'éditeur VBA
3. **Fichier → Importer un fichier** (ou `Ctrl + M`)
4. **Sélectionner le fichier** : `ExportPontivaJson.bas`
   - Chemin : `macros/Macro MSP/Export JSON Pontiva/ExportPontivaJson.bas`
5. **Fermer l'éditeur VBA** (`Alt + F11`)

### Utilisation de la macro

1. **Ouvrir votre projet MS Project**
2. **Appuyer sur `Alt + F11`** pour ouvrir l'éditeur VBA
3. **Dans la fenêtre VBA**, appuyer sur `F5` ou cliquer sur le bouton "Exécuter"
4. **Sélectionner la macro** `ExportProjectToJson` dans la liste
5. **Cliquer sur "Exécuter"**

✅ **Le fichier JSON est automatiquement créé dans votre dossier Téléchargements !**

Le nom du fichier suit le format : `Pontiva_[NomProjet]_[DateHeure].json`

**Exemple** : `Pontiva_AUC5005-0-HS-IEL_20251205_143022.json`

---

## 🌐 ÉTAPE 2 : Importer le JSON sur le site web

1. **Ouvrir le fichier** `dashboard.html` dans votre navigateur
   - Vous pouvez double-cliquer sur le fichier ou l'ouvrir via un serveur web local
2. **Cliquer sur le bouton** "📥 Importer le JSON Pontiva"
3. **Sélectionner le fichier JSON** que vous venez d'exporter depuis MS Project
4. **Le dashboard s'affiche automatiquement** avec les données de votre projet

---

## 👥 ÉTAPE 3 : Utiliser les différentes vues

Le dashboard propose **3 vues** accessibles via des onglets en haut de la page :

### 🎯 Vue Responsable d'affaires

**Accès** : Cliquer sur l'onglet "Responsable d'affaires"

**Fonctionnalités** :
- **Filtres par activité** :
  - **Toutes** : Affiche toutes les tâches
  - **Mécanique** : Filtre les tâches liées à la mécanique
  - **Électrique** : Filtre les tâches liées à l'électrique
  - **Qualité** : Filtre les tâches liées à la qualité

**Informations affichées** :
- Nom de la tâche
- % d'avancement
- Date de début
- Date de fin prévue
- Date de fin réelle (si disponible)
- État de la tâche (À l'heure / En retard / À venir / Terminé)

**Comment ça fonctionne** :
- Le système détecte automatiquement la catégorie d'une tâche en analysant :
  - Les champs personnalisés (Text1, Text2, etc.)
  - Le nom de la tâche (recherche des mots-clés "mécanique", "électrique", "qualité")

---

### 👷 Vue Responsable d'activités

**Accès** : Cliquer sur l'onglet "Responsable d'activités"

**Fonctionnalités** :
- **Vue agrégée par groupe de ressources**
- Chaque groupe affiche :
  - Nombre de tâches
  - Somme des heures réelles travaillées
  - Somme des heures restantes
  - % moyen d'avancement du groupe

**Comment ça fonctionne** :
- Les tâches sont regroupées selon le champ "Groupe" des ressources assignées
- Si une ressource n'a pas de groupe, elle apparaît dans "Sans groupe"

---

### 🏢 Vue Client

**Accès** : Cliquer sur l'onglet "Vue Client"

**Fonctionnalités** :

1. **Histogramme d'avancement global**
   - Affiche le % d'avancement par grande catégorie (Mécanique, Électrique, Qualité)
   - Graphique en barres visuel et clair

2. **3 prochaines tâches qui seront finies**
   - Tableau avec :
     - Nom de la tâche
     - % d'avancement
     - Date de fin prévue
   - Triées par date de fin croissante
   - Filtrées sur les tâches non terminées

3. **3 prochaines tâches qui vont démarrer**
   - Tableau avec :
     - Nom de la tâche
     - Date de démarrage prévue
   - Triées par date de début croissante
   - Filtrées sur les tâches à venir

4. **Indicateur d'avancement de la semaine**
   - Affiche le % de tâches terminées dans la semaine courante
   - Compte le nombre de tâches terminées sur le total des tâches de la semaine

---

## 📊 Données exportées

Le fichier JSON contient pour chaque tâche :

### Informations de base
- UID (identifiant unique)
- Nom de la tâche
- Durée
- Dates (début, fin, fin prévue, fin réelle)

### Avancement
- % achevé
- % physique achevé
- % travail achevé
- État (on_time, late, not_started, completed)

### Référence (Baseline)
- Début de référence
- Fin de référence
- Durée de référence

### Durées
- Durée planifiée
- Durée réelle
- Durée restante

### Travail
- Travail réel (en heures)
- Travail restant (en heures)

### Ressources
Pour chaque ressource assignée :
- Type (work, material, cost)
- Nom
- Groupe
- Heures réelles travaillées
- Heures restantes

### Relations
- Prédécesseurs (liste des UID)
- Successeurs (liste des UID)

### Champs personnalisés
- Tous les champs personnalisés non vides (Text1-30, Number1-20, Flag1-20, Date1-10)

---

## ⚠️ Notes importantes

### Catégorisation automatique

Pour que le filtrage par activité fonctionne correctement, vous pouvez :

1. **Utiliser les champs personnalisés** :
   - Remplir un champ Text (Text1, Text2, etc.) avec "Mécanique", "Électrique" ou "Qualité"
   
2. **Utiliser le nom de la tâche** :
   - Inclure le mot "mécanique", "électrique" ou "qualité" dans le nom de la tâche

### Groupes de ressources

Pour que la vue "Responsable d'activités" fonctionne :
- Assurez-vous que vos ressources ont un **Groupe** défini dans MS Project
- Le groupe peut être défini dans la vue Ressources de MS Project

### Champs personnalisés

- Seuls les champs personnalisés **non vides** sont exportés
- Les champs vides, à 0, ou avec la date par défaut (01/01/1984) sont ignorés

---

## 🐛 Dépannage

### La macro ne s'exécute pas
- Vérifiez que les macros sont activées dans MS Project
- Vérifiez que le projet est bien ouvert et actif

### Le fichier JSON n'est pas créé
- Vérifiez que le dossier Téléchargements existe
- Vérifiez les permissions d'écriture
- Si le dossier Téléchargements n'existe pas, le fichier sera créé sur le Bureau

### Le dashboard n'affiche pas les données
- Vérifiez que le fichier JSON est bien formaté
- Ouvrez la console du navigateur (F12) pour voir les erreurs éventuelles
- Vérifiez que le fichier contient bien les champs `version`, `project_name` et `tasks`

### Les filtres ne fonctionnent pas
- Vérifiez que vos tâches ont bien une catégorie détectable (dans le nom ou un champ personnalisé)
- Les catégories sont détectées automatiquement, mais vous pouvez les forcer via les champs personnalisés

---

## 📝 Exemple de structure JSON

```json
{
  "version": "0.2",
  "project_name": "Mon Projet",
  "export_date": "2025-12-05",
  "tasks": [
    {
      "uid": 11,
      "name": "Battage",
      "duration": "109.4",
      "start": "2025-01-15",
      "finish": "2025-02-20",
      "percent_complete": 50,
      "status": "on_time",
      "resources": [
        {
          "type": "work",
          "name": "Battage Ouest",
          "group": "Mécanique",
          "actual_work_hours": 54.7,
          "remaining_work_hours": 54.7
        }
      ],
      "custom_fields": {
        "Text1": "Mécanique"
      }
    }
  ]
}
```

---

## ✅ Checklist rapide

- [ ] Macro VBA importée dans MS Project
- [ ] Projet MS Project ouvert
- [ ] Macro exécutée avec succès
- [ ] Fichier JSON créé dans Téléchargements
- [ ] Dashboard HTML ouvert dans le navigateur
- [ ] Fichier JSON importé dans le dashboard
- [ ] Vues testées (Responsable d'affaires, Responsable d'activités, Client)

---

**Besoin d'aide ?** Vérifiez la console du navigateur (F12) pour les erreurs JavaScript, ou consultez les messages d'erreur dans MS Project lors de l'export.

