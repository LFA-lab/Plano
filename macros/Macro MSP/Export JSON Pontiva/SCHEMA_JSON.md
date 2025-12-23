# Schéma JSON Pontiva v0.2

## Structure globale

```json
{
  "version": "0.2",
  "project_name": "string",
  "export_date": "YYYY-MM-DD",
  "tasks": [ ... ]
}
```

## Structure d'une tâche

```json
{
  "uid": 123,
  "name": "Nom de la tâche",
  "duration": "109.4",
  "start": "2025-01-15",
  "finish": "2025-02-20",
  "predecessors": "11, 12",
  "successors": "14, 15",
  "resources": [ ... ],
  "percent_complete": 50,
  "physical_percent_complete": 30,
  "percent_work_complete": 45,
  "constraint_date": "2025-01-20",
  "baseline_start": "2025-01-10",
  "baseline_finish": "2025-02-15",
  "baseline_duration": "100.0",
  "planned_duration": "109.4",
  "actual_duration": "54.7",
  "remaining_duration": "54.7",
  "status": "on_time | late | not_started | completed",
  "scheduled_finish": "2025-02-20",
  "actual_finish": "2025-02-18",
  "notes": "Remarques sur la tâche",
  "actual_work_hours": 54.7,
  "remaining_work_hours": 54.7,
  "custom_fields": { ... }
}
```

## Structure d'une ressource

```json
{
  "type": "work | material | cost",
  "name": "Nom de la ressource",
  "group": "Groupe de ressources",
  "actual_work_hours": 54.7,
  "remaining_work_hours": 54.7
}
```

## Structure des champs personnalisés

```json
{
  "custom_fields": {
    "Text1": "Valeur texte",
    "Text2": "Autre valeur",
    "Number1": 123.45,
    "Number2": 0,
    "Flag1": true,
    "Flag2": false,
    "Date1": "2025-01-15",
    "Date2": "2025-02-20"
  }
}
```

## Types de données

### Dates
- Format : `YYYY-MM-DD` (ISO 8601)
- Valeur vide : `""` (chaîne vide)
- Dates invalides ou par défaut MS Project sont exclues

### Durées
- Format : nombre décimal en heures (point décimal)
- Exemple : `"109.4"` pour 109 heures et 24 minutes

### Statut
- `"on_time"` : Tâche à l'heure
- `"late"` : Tâche en retard
- `"not_started"` : Tâche non démarrée
- `"completed"` : Tâche terminée

### Type de ressource
- `"work"` : Ressource de travail
- `"material"` : Ressource matérielle
- `"cost"` : Ressource de coût

### Prédécesseurs / Successeurs
- Format : liste d'UID séparés par des virgules
- Exemple : `"11, 12, 13"`
- Valeur vide si aucun : `""`

## Notes importantes

1. **Champs personnalisés** : Seuls les champs non vides sont inclus
2. **Dates** : Les dates par défaut MS Project (01/01/1984, 31/12/2049) sont exclues
3. **Ressources** : Seules les ressources avec des heures > 0 sont incluses
4. **Tâches** : Seules les tâches non récapitulatives (Summary = False) avec un nom sont exportées

