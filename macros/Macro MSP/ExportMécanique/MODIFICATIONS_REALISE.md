# Modifications réalisées sur exportmécanique.bas

## Résumé des changements

La macro a été adaptée pour générer deux feuilles distinctes "Données Est" et "Données Ouest" au lieu d'une seule feuille "Données détaillées", en filtrant les ressources consommables selon leur champ `Texte2`.

## Modifications apportées

### 1. Ajout de constantes
- **Ligne 9** : Ajout de `Const pjResourceTypeMaterial = 1`
- **Mise à jour des commentaires d'en-tête** pour refléter les nouvelles feuilles

### 2. Nouvelle fonction de filtrage
- **Fonction `GetFilteredConsumableResources`** (lignes ~468-545) :
  - Filtre les ressources par type (`pjResourceTypeMaterial`) ET par `Texte2` ("Est" ou "Ouest")
  - Conserve le tri par ID de tâche comme dans `GetSortedMechanicalResources`
  - Vérifie aussi que le groupe est "Mécanique"

### 3. Nouvelle fonction de création des feuilles
- **Fonction `CreateDetailSheets`** (lignes ~910-970) :
  - Remplace l'ancienne logique de création d'une seule feuille
  - Crée une feuille "Données Est" si des ressources Est existent
  - Crée une feuille "Données Ouest" si des ressources Ouest existent
  - Retourne un message détaillé pour l'affichage final
  - Réutilise les fonctions existantes `WriteDetailSheet` et `FormatDetailSheet`

### 4. Modifications du code principal
- **Suppression** de la déclaration `xlDetailSheet` (ligne ~56)
- **Suppression** de la création de "Données détaillées" (lignes ~205-210)
- **Remplacement** de l'appel `WriteDetailSheet` par `CreateDetailSheets` (lignes ~275-280)
- **Adaptation** du message final pour afficher les informations des feuilles créées (lignes ~303-307)

## Logique de filtrage implémentée

```vb
If res.Type = pjResourceTypeMaterial Then
    Dim texte2Value As String
    texte2Value = Trim(res.Text2)
    
    If texte2Value = filter Then ' "Est" ou "Ouest"
        ' Vérifier aussi que c'est une ressource "Mécanique"
        ' Ajouter à la collection filtrée
    End If
End If
```

## Comportement de la macro modifiée

### Avant
- Feuille 1 : "Récapitulatif" (toutes les ressources mécaniques)
- Feuille 2 : "Données détaillées" (toutes les ressources mécaniques)

### Après
- Feuille 1 : "Récapitulatif" (toutes les ressources mécaniques - **inchangé**)
- Feuille 2 : "Données Est" (uniquement ressources consommables avec Texte2 = "Est")
- Feuille 3 : "Données Ouest" (uniquement ressources consommables avec Texte2 = "Ouest")

### Gestion des cas particuliers
- Si aucune ressource Est : pas de feuille "Données Est"
- Si aucune ressource Ouest : pas de feuille "Données Ouest"
- Si Texte2 est vide ou autre valeur : ressource ignorée
- Si Type ≠ Material : ressource ignorée

## Message final adapté

Le message de fin d'export affiche maintenant :
```
Export terminé :
Fichier Excel : [chemin]
Onglet 1 : Récapitulatif (X ressources)
Données Est (Y ressources)
Données Ouest (Z ressources)
Total: W ressource(s) consommable(s)
(X dates réelles)
```

## Test recommandé

Pour tester les modifications :
1. Créer des ressources de type "Material" avec Group="Mécanique"
2. Affecter Texte2 = "Est" à certaines, "Ouest" à d'autres, laisser vide pour d'autres
3. Créer des assignations avec heures réelles
4. Exécuter la macro et vérifier que :
   - Les feuilles sont créées uniquement si des ressources correspondantes existent
   - La structure est identique à l'ancienne "Données détaillées"
   - Le filtrage fonctionne correctement
