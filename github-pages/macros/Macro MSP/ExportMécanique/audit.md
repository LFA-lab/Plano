# Audit Macro Export Mécanique - Adaptation Text2 Assignments

## 🔍 Analyse du code existant

### 1. Localisation des boucles Resource actuelles

La lecture des ressources est principalement effectuée dans :

```vb
' Function GetSortedMechanicalResources (ligne ~400)
For Each res In proj.Resources
    If Not res Is Nothing Then
        ' Lecture directe de Text2 sur la ressource :
        cleanTexte2 = Trim(res.Text2)  
```

### 2. Analyse de l'utilisation de Text2

- ❌ **Actuellement** : `res.Text2` lu directement sur l'objet Resource
- ✅ **Nécessaire** : `assn.Text2` à lire sur chaque Assignment

### 3. Problème identifié

La macro actuelle :
- Parcourt uniquement les ressources (Resources)
- Lit Text2 au mauvais niveau (Resource au lieu de Assignment)
- Ne peut pas détecter les ressources utilisées dans plusieurs zones (Est/Ouest)

## 📋 Plan de modification

### 1. Nouvelle fonction principale proposée

```vb
Function GetSortedMechanicalAssignments(texte2Filter As String) As Collection
    ' Init collection résultat
    Set result = New Collection
    
    ' Pour chaque tâche du projet
    For Each tsk In ActiveProject.Tasks
        ' Pour chaque assignation de la tâche  
        For Each assn In tsk.Assignments
            
            ' Récupérer la ressource
            Set res = assn.Resource
            
            ' Filtres :
            If res.Type = pjResourceTypeMaterial And _
               Trim(res.Group) = "Mecanique" And _
               Trim(assn.Text2) = texte2Filter Then
                
                ' Ajouter l'assignation à la collection
                result.Add assn
            End If
        Next assn
    Next tsk
    
    Set GetSortedMechanicalAssignments = result
End Function
```

## 🔄 Impacts à gérer

### 1. Gestion des doublons
- Une même ressource peut avoir des assignations "Est" ET "Ouest"
- Solution : accepter cette situation car elle reflète la réalité du planning

### 2. Modifications WriteDetailSheet
Options possibles :
1. Modification profonde pour utiliser des assignations
2. Création d'un wrapper assignations → ressources

### 3. Adaptation des calculs
Fonctions à modifier :
- `ComputeTotalPlannedWork()`
- `ComputeDailyActualWork()`
- `ComputeCumulativeActual()`

## 💡 Recommandation

### Option recommandée : Approche wrapper

1. Créer une structure intermédiaire :
```vb
Type PseudoResource
    Name As String             ' Nom ressource original  
    Assignments As Collection  ' Collection d'assignations filtrées
End Type
```

2. Avantages :
   - Préserve la structure existante
   - Facilite les tests et le débogage
   - Permet une transition progressive
   - Minimise les risques de régression

### Plan d'implémentation proposé

1. Créer les nouvelles fonctions de collecte
2. Implémenter le wrapper PseudoResource
3. Adapter progressivement WriteDetailSheet
4. Mettre à jour les fonctions de calcul
5. Tester avec des plannings variés

## ⚠️ Points d'attention

1. Performance
   - Double parcours tâches/assignations
   - Impact faible car nombre limité de ressources mécaniques

2. Gestion erreurs
   - Valider Text2 null/vide
   - Vérifier existence ressource
   - Journalisation détaillée

3. Messages utilisateur
   - Adapter pour parler d'assignations
   - Indiquer nombre d'assignations trouvées
