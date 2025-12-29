# 📊 TRAÇABILITÉ DES DONNÉES - RAPPORT PREVENCHERES

## 🎯 Objectif

Cette fonctionnalité génère **automatiquement** un fichier `.txt` détaillé lors de l'exécution de `BuildWeeklyReport()`. Ce fichier permet de **tracer l'origine de chaque donnée** affichée dans les graphiques et tableaux du rapport Word.

---

## 🚀 Utilisation

### Exécution automatique

Lorsque vous exécutez la macro `BuildWeeklyReport()` dans MS Project, **2 fichiers** sont générés sur votre Bureau :

1. **`Rapport_Hebdo_Prevencheres_2025-12-29_1445.docx`** - Rapport Word habituel
2. **`Rapport_Data_Trace_2025-12-29_1445.txt`** ✨ - Fichier de traçabilité (NOUVEAU)

> 💡 **Aucune action supplémentaire requise** - La traçabilité est générée automatiquement.

---

## 📋 Contenu du fichier de traçabilité

Le fichier `.txt` est structuré en **6 parties** :

### **PARTIE 1 : Liste brute de toutes les tâches**
- Format tabulaire avec toutes les propriétés
- Colonnes : ID, Nom, Zone, SousZone, Métier, Work, ActualWork, %Complete, Ressources, Summary

**Exemple :**
```
[12] | Pose de câbles Zone 1 | 1 | 1-Nord | Électricité | 8.00 | 4.00 | 50.0% | Jean, Pierre, CQ | NON
[13] | Montage supports Zone 2 | 2 |  | Montage | 12.50 | 0.00 | 0.0% | Marie | NON
```

### **PARTIE 2 : Section 2 - Graphique 2.1**
**Avancement par Zone et Métier (% tâches)**

Pour chaque combinaison (Zone|Métier), vous voyez :
- Le **résultat final** (pourcentage qui apparaît dans le graphique)
- La **liste des tâches** qui contribuent au calcul
- Le **détail du calcul** étape par étape

**Exemple :**
```
--------------------------------------------------------------------------------
📊 1 | ÉLECTRICITÉ => 45.0%
   Nombre de tâches : 5
   Détail des tâches :
   ├─ [12] Pose de câbles Zone 1 : 50.0% (Work=8.00h, ActualWork=4.00h)
   ├─ [15] Raccordement électrique : 20.0% (Work=6.00h, ActualWork=1.20h)
   ├─ [18] Tests électriques : 60.0% (Work=4.00h, ActualWork=2.40h)
   ├─ [21] Installation armoires : 70.0% (Work=10.00h, ActualWork=7.00h)
   └─ [24] Câblage local : 25.0% (Work=3.00h, ActualWork=0.75h)
   
   Calcul (moyenne % Complete) :
   = 225.0 / 5
   = 45.0%
```

### **PARTIE 3 : Section 2 - Graphique 2.2**
**Avancement par Zone et Métier (% ressources)**

Même structure, mais avec calcul basé sur ActualWork/Work :

**Exemple :**
```
📊 1 | ÉLECTRICITÉ => 48.4%
   ...
   Calcul (% ressources) :
   = (15.00h / 31.00h) × 100
   = 48.4%
```

### **PARTIE 4 : Section 2 - Graphique 2.3**
**Avancement par Sous-Zone et Métier (% tâches)**

Groupement par Text3 (Sous-Zone) au lieu de Text2 (Zone)

### **PARTIE 5 : Section 2 - Graphique 2.4**
**Avancement par Sous-Zone et Métier (% ressources)**

Groupement par Text3 (Sous-Zone) avec calcul % ressources

### **PARTIE 6 : Section 3 - Contrôles Qualité**
**Tableau et Graphique CQ par Zone et Métier**

Pour chaque combinaison (Zone|Métier) avec tâches CQ :

**Exemple :**
```
--------------------------------------------------------------------------------
📊 1 | ÉLECTRICITÉ
   Nb CQ Total : 3
   Nb CQ Terminés (100%) : 2
   % Complet Moyen : 83.3%
   Détail des tâches CQ :
   ├─ [45] Contrôle Qualité - Pose de câbles : 100.0% ✓
   ├─ [48] Contrôle Qualité - Tests électriques : 100.0% ✓
   └─ [51] Contrôle installation armoires : 50.0%
   
   Calcul (% moyen) :
   = 250.0 / 3
   = 83.3%
```

---

## 🔍 Cas d'usage

### ✅ **Validation des calculs**
Vous pouvez vérifier que les pourcentages affichés dans le rapport Word correspondent bien aux données MS Project.

### ✅ **Débogage**
Si un graphique montre des valeurs inattendues, le fichier de traçabilité permet d'identifier rapidement :
- Quelles tâches sont prises en compte
- Quelles tâches sont ignorées (et pourquoi)
- Comment le calcul est effectué

### ✅ **Audit et documentation**
Le fichier `.txt` sert de preuve documentaire pour montrer la provenance des données du rapport.

### ✅ **Compréhension de la logique**
Si vous reprenez le projet plus tard ou si quelqu'un d'autre utilise le code, le fichier de traçabilité explique clairement la logique de calcul.

---

## 🛠️ Architecture technique

### **Fonction principale**
```vba
ExportProjectDataTrace(outFolder)
```
- **Rôle** : Orchestrateur principal qui génère le fichier `.txt`
- **Appelée depuis** : `BuildWeeklyReport()` (ligne 45)
- **Emplacement** : Fin du fichier `RapportPrevencheres.vb` (section dédiée)

### **Fonctions auxiliaires**

#### `TraceExportRawTaskList(txtFile)`
- Génère la liste brute de toutes les tâches
- Format tabulaire simple

#### `TraceExportProgressDetails(txtFile, groupBy, useTaskPercent)`
- Génère le détail des calculs pour Section 2 (avancement)
- **Paramètres** :
  - `groupBy` : "Zone" ou "SousZone"
  - `useTaskPercent` : True (% tâches) ou False (% ressources)

#### `TraceExportQualityDetails(txtFile)`
- Génère le détail des calculs pour Section 3 (CQ)
- Filtre les tâches avec ressource "CQ"

### **Séparation du code**

✅ **Les fonctions de traçabilité sont isolées** dans une section dédiée à la fin du fichier

✅ **Pas de modification** des fonctions existantes de calcul (`ExtractProgressData`, `ExtractQualityData`, etc.)

✅ **Maintenabilité** : Si vous modifiez les sections 2 ou 3, vous devez simplement mettre à jour les fonctions `TraceExport*` correspondantes

---

## 🚨 Erreurs possibles

### **Erreur : "ActiveProject Is Nothing"**
**Cause** : Aucun projet MS Project n'est ouvert  
**Solution** : Ouvrir un fichier `.mpp` avant d'exécuter la macro

### **"[Aucune donnée disponible pour ce graphique]"**
**Cause** : Aucune tâche ne satisfait les critères (Text2/Text3/Text4 vides, Work=0, Summary=True)  
**Solution** : Vérifier que les champs personnalisés Text2, Text3, Text4 sont bien remplis dans MS Project

### **"[Aucune tâche CQ détectée]"**
**Cause** : Aucune tâche n'a de ressource nommée "CQ"  
**Solution** : Affecter la ressource "CQ" aux tâches concernées dans MS Project

---

## 📝 Exemple de workflow

1. **Ouvrir MS Project** avec votre fichier `.mpp`
2. **Vérifier** que les champs Text2, Text3, Text4 sont bien remplis
3. **Exécuter** `BuildWeeklyReport()` (Alt+F11 > F5)
4. **Consulter** le fichier `Rapport_Data_Trace_XXXX.txt` sur le Bureau
5. **Comparer** les valeurs du `.txt` avec celles du rapport Word `.docx`

---

## ✅ Avantages

| Aspect | Avantage |
|--------|----------|
| **Transparence** | Chaque valeur du rapport est traçable jusqu'aux tâches sources |
| **Validation** | Permet de vérifier les calculs manuellement |
| **Débogage** | Identification rapide des problèmes de données |
| **Documentation** | Preuve de l'origine des données pour audits |
| **Maintenabilité** | Code séparé, facile à modifier sans casser les sections |

---

## 📞 Support

Pour toute question sur la traçabilité des données :
1. Consulter ce README
2. Ouvrir le fichier `.txt` généré et analyser les logs
3. Vérifier la section "TRAÇABILITÉ" à la fin de `RapportPrevencheres.vb`

---

**Date de création** : 29 décembre 2025  
**Version** : 1.0  
**Auteur** : Implémentation via Cursor AI

