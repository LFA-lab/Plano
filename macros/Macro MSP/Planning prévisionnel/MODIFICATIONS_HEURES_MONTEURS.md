# Modification VBA : Peak Units → Heures Monteurs

## 📋 Résumé des modifications

Le module VBA `Planningheures.bas` a été modifié pour remplacer l'export des unités de pointe (Peak Units) par l'export des heures journalières de la ressource "Monteurs".

## 🔄 Principales modifications effectuées

### 1. Nouveau point d'entrée principal
- **Avant** : `Sub PlanningHeures()`
- **Après** : `Sub Export_Plan_Et_HeuresMonteurs()`
- **Compatibilité** : L'ancienne procédure `PlanningHeures()` redirige maintenant vers la nouvelle version

### 2. Nouvelle procédure d'injection
- **Avant** : `Sub Injecter_UnitesDePointe_Dans_Planning(xlApp As Object, xlBook As Object)`
- **Après** : `Sub Injecter_HeuresMonteurs_Dans_Planning(xlApp As Object, xlBook As Object)`

## 🎯 Changements fonctionnels

### Filtrage des ressources
```vb
' AVANT - Toutes les ressources de type Work
If assign.Resource.Type = pjResourceTypeWork Then
    Set tsData = assign.TimeScaleData(debut, fin, pjAssignmentTimescaledPeakUnits, pjTimescaleDays)

' APRÈS - Uniquement la ressource "Monteurs"
If assign.Resource.Name = "Monteurs" And assign.Resource.Type = pjResourceTypeWork Then
    Set tsData = assign.TimeScaleData(debut, fin, pjTimescaleDays, pjTimescaleWork)
```

### Conversion des données
```vb
' AVANT - Peak Units (valeur directe)
arr(arrRowIdx, arrColIdx) = Round(currentValue + CDbl(tsValue.Value), 2)

' APRÈS - Work en minutes converti en heures
heuresJour = Round(CDbl(tsValue.Value) / 60, 2)
arr(arrRowIdx, arrColIdx) = Round(currentValue + heuresJour, 2)
```

## 📊 APIs MS Project utilisées

### Nouvelles APIs pour les heures Monteurs
- `assign.Resource.Name = "Monteurs"` : Filtrage par nom de ressource
- `pjTimescaleWork` : Extraction du travail en minutes
- `Round(CDbl(tsValue.Value) / 60, 2)` : Conversion minutes → heures (2 décimales)

### APIs conservées
- `assign.TimeScaleData()` : Extraction des données temps-phasées
- Gestion des dictionnaires pour l'optimisation
- Écriture par plages Excel pour les performances

## ✅ Fonctionnalités maintenues

1. **Optimisation Excel** : Désactivation temporaire des événements
2. **Gestion d'erreurs** : Restauration de l'état en cas d'erreur
3. **Multi-affectations** : Agrégation des heures pour plusieurs affectations "Monteurs" le même jour
4. **Performance** : Une seule lecture/écriture de plage Excel
5. **Totaux** : Calcul automatique des totaux par colonne avec formatage

## 🚀 Résultat attendu

L'export Excel affiche maintenant :
- Les heures "Monteurs" par jour au lieu des Peak Units
- Agrégation automatique des multi-affectations
- Respect des calendriers projet
- Conversion précise : minutes → heures (2 décimales)
- Totaux par colonne avec formatage (ligne 3, fond jaune, gras)

## 📝 Points de test recommandés

1. **Tâche avec Monteurs** : Vérifier l'affichage des heures correctes
2. **Tâche sans Monteurs** : Vérifier que les colonnes restent vides
3. **Multi-affectations** : Vérifier la somme des heures multiples
4. **Performance** : Chronométrer l'exécution sur gros projets
5. **Compatibilité** : Tester l'appel à l'ancienne procédure `PlanningHeures()`

## 🔧 Maintenance

La procédure `PlanningHeures()` originale est conservée pour la compatibilité ascendante et redirige automatiquement vers la nouvelle version `Export_Plan_Et_HeuresMonteurs()`.
