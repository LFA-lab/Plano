# 📚 Documentation : Import Excel → MS Project

## 🎯 C'est quoi ce programme ?

Imagine que tu as une liste de tâches dans Excel (comme un tableau de devoirs), et tu veux les mettre dans MS Project (un logiciel pour gérer des projets). Ce programme fait ça automatiquement pour toi !

**En résumé** : Il prend ton fichier Excel et crée automatiquement un beau planning dans MS Project avec toutes les tâches, les heures de travail, et qui fait quoi.

---

## 📂 Structure du fichier Excel attendu

Ton fichier Excel doit ressembler à ça :

| Colonne | Nom | Description | Exemple |
|---------|-----|-------------|---------|
| **A** | Nom de la tâche | Ce qu'il faut faire | "Installer l'électricité" |
| **B** | Quantité | Combien de matériel | 100 |
| **C** | Personnes | Combien de personnes travaillent | 2 |
| **D** | Heures | Combien d'heures de travail | 32 |
| **E** | Zone | Dans quelle zone | "Zone 1" |
| **F** | Sous-Zone | Détail de la zone | "Bâtiment A" |
| **G** | Tranche | Quelle phase du projet | "Tranche A" |
| **H** | Métier | Type de travail | "Électricité" |
| **I** | Entreprise | Quelle entreprise | "OMEXOM" |
| **J** | Qualité | Contrôle qualité ? | "CQ" ou "TACHE" ou vide |

**Important** :
- La **ligne 1** : Titre des colonnes (pas utilisé par le programme)
- La **ligne 2** : Le titre du projet (colonne A uniquement)
- **À partir de la ligne 3** : Les tâches à importer

---

## 🚀 Comment ça marche ? (Vue d'ensemble)

```
┌─────────────────┐
│  Fichier Excel  │
└────────┬────────┘
         │
         ↓
┌────────────────────────────┐
│  1. Ouvrir le fichier      │
│  2. Créer MS Project       │
│  3. Lire chaque ligne      │
│  4. Créer les tâches       │
│  5. Forcer les heures      │
│  6. Calculer tout          │
└────────┬───────────────────┘
         │
         ↓
┌─────────────────────┐
│  MS Project prêt !  │
│  + Fichier log      │
└─────────────────────┘
```

---

## 📝 Explication détaillée : Étape par étape

### 🔧 1. Préparation (Lignes 3-8)

```vba
Dim xlApp As Object, xlBook As Object, xlSheet As Object
Dim pjApp As MSProject.Application, pjProj As MSProject.Project
```

**En français simple** : On prépare des "boîtes" pour stocker Excel et MS Project.

- `xlApp` = L'application Excel
- `xlBook` = Le fichier Excel ouvert
- `xlSheet` = La feuille du fichier
- `pjApp` = L'application MS Project
- `pjProj` = Le projet créé dans MS Project

---

### 📁 2. Sélection du fichier (Lignes 10-32)

```vba
With xlTempApp.FileDialog(msoFileDialogFilePicker)
    .Title = "Sélectionnez le fichier Excel à importer"
    .InitialFileName = Environ$("USERPROFILE") & "\Downloads\"
```

**En français simple** : On ouvre une fenêtre pour que tu puisses choisir ton fichier Excel.

**Pourquoi** : Le programme ne sait pas où est ton fichier, donc il te demande de le montrer.

**Astuce** : La fenêtre s'ouvre directement dans ton dossier "Téléchargements" !

---

### 📖 3. Ouverture d'Excel (Lignes 34-39)

```vba
Set xlApp = CreateObject("Excel.Application")
xlApp.Visible = False
Set xlBook = xlApp.Workbooks.Open(FileName:=fichierExcel, ReadOnly:=True)
```

**En français simple** : On ouvre Excel en mode invisible (tu ne le vois pas) et on lit ton fichier.

**Pourquoi invisible ?** Pour aller plus vite et ne pas te déranger avec des fenêtres qui s'ouvrent.

---

### 🏗️ 4. Création du projet MS Project (Lignes 41-58)

```vba
Set pjApp = MSProject.Application
pjApp.Visible = True
pjApp.FileNew
```

**En français simple** : On ouvre MS Project et on crée un nouveau projet vide.

Ensuite on configure les **noms des colonnes personnalisées** :

```vba
pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
```

Ça permet d'avoir des colonnes avec des noms clairs comme "Tranche" au lieu de "Texte1".

---

### 📅 5. Configuration du calendrier (Lignes 65-80)

```vba
For j = 2 To 6 ' Lundi à vendredi
    With .WeekDays(j)
        .Shift1.Start = "09:00"
        .Shift1.Finish = "18:00"
```

**En français simple** : On dit à MS Project que les gens travaillent :
- Du lundi au vendredi
- De 9h à 18h
- Pas le week-end

**Pourquoi ?** Pour que MS Project calcule bien les durées des tâches.

---

### 👷 6. Création des ressources (Lignes 82-95)

```vba
Set rMonteurs = GetOrCreateWorkResource("Monteurs")
rMonteurs.MaxUnits = 10 ' 10 personnes max
```

**En français simple** : On crée une "équipe" appelée "Monteurs" qui peut avoir jusqu'à 10 personnes.

On crée aussi une ressource "CQ" pour le Contrôle Qualité.

**Astuce importante** : On désactive le calcul automatique pour éviter les popups embêtants !

---

### 📝 7. Création du fichier LOG (Lignes 99-111)

```vba
logFile = Replace(fichierExcel, ".xlsx", "_import_log.txt")
Set logStream = fso.CreateTextFile(logFile, True)
```

**En français simple** : On crée un fichier texte à côté de ton Excel pour noter tout ce qui se passe.

**Exemple** : Si ton fichier s'appelle `MonProjet.xlsx`, le log sera `MonProjet_import_log.txt`

**Pourquoi ?** Pour pouvoir vérifier si tout s'est bien passé et débugger en cas de problème.

---

### 🔄 8. Boucle principale : Lecture des tâches (Lignes 114-242)

C'est la partie la plus importante ! Le programme va lire **chaque ligne** de ton Excel et créer les tâches.

#### 📋 8.1. Lecture des données (Lignes 120-134)

```vba
nom = Trim(CStr(xlSheet.Cells(i, 1).Value))      ' Colonne A
qte = xlSheet.Cells(i, 2).Value                  ' Colonne B
pers = xlSheet.Cells(i, 3).Value                 ' Colonne C
h = xlSheet.Cells(i, 4).Value                    ' Colonne D
zone = Trim(CStr(xlSheet.Cells(i, 5).Value))     ' Colonne E
```

**En français simple** : On lit toutes les colonnes de la ligne actuelle.

#### 🏗️ 8.2. Création de la tâche (Lignes 146-160)

```vba
Set t = pjProj.Tasks.Add(nom)
t.Manual = False
t.LevelingCanSplit = False
```

**En français simple** : On crée une tâche dans MS Project avec le nom qu'on a lu.

**Détails** :
- `Manual = False` : La tâche est automatique (MS Project calcule les dates)
- `LevelingCanSplit = False` : La tâche ne peut PAS être coupée en morceaux

On remplit aussi les **tags** (Tranche, Zone, etc.) :

```vba
t.Text1 = tranche
t.Text2 = zone
t.Text3 = sousZone
```

#### 📦 8.3. Ajout du matériau (Lignes 164-173)

```vba
If IsNumeric(qte) And qte > 0 Then
    Set rMat = GetOrCreateMaterialResource(nom)
    Set a = t.Assignments.Add(ResourceID:=rMat.ID)
    a.Units = CDbl(qte)
End If
```

**En français simple** : Si tu as indiqué une quantité de matériel (colonne B), on l'ajoute à la tâche.

**Exemple** : 100 mètres de câble pour la tâche "Installer l'électricité".

#### ✅ 8.4. Contrôle Qualité (Lignes 175-206)

Il y a **3 cas possibles** :

**Cas 1** : Colonne J = "CQ"
```vba
Set a = t.Assignments.Add(ResourceID:=rCQ.ID)
a.Units = 1
```
→ On ajoute une ressource CQ directement sur la tâche.

**Cas 2** : Colonne J = "TACHE"
```vba
Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
```
→ On crée une **nouvelle tâche** séparée pour le contrôle qualité.

**Cas 3** : Colonne J = vide
→ Pas de contrôle qualité, on ne fait rien.

#### ⏱️ 8.5. Ajout des heures de travail (Lignes 208-236)

**C'est LA partie la plus critique !**

```vba
workMinutes = CLng(CDbl(h) * 60)

Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)

' ÉTAPE 1: Assigner Work EN PREMIER
a.Work = workMinutes

' ÉTAPE 2: Puis assigner Units
a.Units = nbPers

' ÉTAPE 3: FORCER le Work à nouveau après Units
a.Work = workMinutes
```

**En français simple** : On dit à MS Project :
1. Cette tâche prend X heures (on convertit en minutes)
2. Il y a Y personnes qui travaillent dessus
3. On **re-force** les heures une deuxième fois

**Pourquoi 3 étapes ?** Parce que MS Project a tendance à recalculer les heures automatiquement. En forçant 2 fois, on est sûr que ça reste bien à la bonne valeur !

**Exemple** :
- Tu mets 32 heures dans Excel (colonne D)
- 2 personnes (colonne C)
- → Le programme met bien 32h de travail total dans MS Project
- → Durée calculée : 32h ÷ 2 personnes = 16h de temps calendaire

---

### 🔁 9. Forçage final du Work (Lignes 244-326)

**Problème** : Même après l'étape 8, MS Project peut **encore** recalculer les heures.

**Solution** : On **reparcourt TOUTES les tâches** une deuxième fois et on force à nouveau les heures !

```vba
For i = 3 To lastRow
    ' ... trouve la tâche ...
    
    tForce.Type = pjFixedWork
    aForce.Work = workMinutesForce
End For
```

**En français simple** : On re-vérifie toutes les tâches et on s'assure que les heures sont correctes.

---

### ✔️ 10. Vérification finale (Lignes 331-388)

```vba
logStream.WriteLine "Excel=" & Format(hoursCheck, "0.00") & "h | Project=" & Format(hoursInProject, "0.00") & "h"
```

**En français simple** : On compare ce qu'il y a dans Excel avec ce qui est dans MS Project, et on écrit ça dans le log.

**Exemple de log** :
```
Ligne 3 - Raccordement base vie: Excel=32.00h | Project=32.00h
```

Si les deux correspondent = ✅ parfait !

---

### 🧮 11. Calcul final (Lignes 396-402)

```vba
pjApp.Calculation = True
pjProj.Calculate
pjApp.CalculateAll
```

**En français simple** : On dit à MS Project : "Maintenant, recalcule TOUT pour que les totaux soient bons !"

**Pourquoi ?** Pour que la ressource "Monteurs" affiche le total correct de toutes les heures.

---

### 🚪 12. Fermeture (Lignes 404-409)

```vba
xlBook.Close SaveChanges:=False
xlApp.Quit
Set xlApp = Nothing
```

**En français simple** : On ferme Excel sans sauvegarder (on n'a rien modifié de toute façon).

Et on affiche un message : "Import terminé !" 🎉

---

## 🛠️ Les fonctions utilitaires

### `GetOrCreateWorkResource(nom As String)` (Lignes 414-424)

**But** : Créer une ressource "personne" (comme "Monteurs").

**En français simple** :
1. On cherche si la ressource existe déjà
2. Si oui → on la renvoie
3. Si non → on la crée

**Pourquoi ?** Pour éviter de créer plusieurs fois la même ressource.

---

### `GetOrCreateMaterialResource(nom As String)` (Lignes 426-436)

**But** : Créer une ressource "matériau" (comme "CQ" ou "Câbles").

**Même principe** que ci-dessus, mais pour du matériel au lieu de personnes.

---

## 🐛 Problèmes résolus dans ce code

### ❌ Problème 1 : Les heures étaient fausses

**Symptôme** : Excel disait 32h, mais MS Project affichait 9h.

**Cause** : MS Project recalculait le Work après l'ajout des autres ressources (matériau, CQ).

**Solution** : 
1. Ordre des assignments : Matériau → CQ → Travail (EN DERNIER)
2. Forcer Work 2 fois : avant ET après Units
3. Forçage final à la fin de l'import

---

### ❌ Problème 2 : Popups de surutilisation

**Symptôme** : À chaque tâche, MS Project affichait "Impossible de résoudre la surutilisation".

**Cause** : Le calcul automatique était actif + MaxUnits trop bas.

**Solution** :
1. Désactiver le calcul automatique pendant l'import
2. MaxUnits = 10 (au lieu de 1)
3. Réactiver le calcul uniquement à la fin

---

### ❌ Problème 3 : Tâches fractionnées

**Symptôme** : MS Project coupait les tâches en plusieurs morceaux.

**Cause** : Option de nivellement par défaut.

**Solution** : `t.LevelingCanSplit = False` sur toutes les tâches.

---

## 📊 Schéma complet du flux

```
┌─────────────────────────────────────────────────────────┐
│                    DEBUT DU PROGRAMME                   │
└────────────────────┬────────────────────────────────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Sélectionner fichier  │ (FileDialog)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Ouvrir Excel          │ (Mode invisible)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Créer MS Project      │ (Nouveau projet)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Configurer calendrier │ (9h-18h, lun-ven)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Créer ressources      │ (Monteurs, CQ)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────┐
         │ Désactiver calcul auto│ (Évite popups)
         └───────────┬───────────┘
                     │
                     ↓
         ┌───────────────────────────────────┐
         │  BOUCLE: Pour chaque ligne Excel  │
         │  ================================  │
         │  1. Lire données (nom, heures...) │
         │  2. Créer tâche                   │
         │  3. Ajouter tags (Zone, Tranche)  │
         │  4. Ajouter matériau (si qté > 0) │
         │  5. Ajouter CQ (si demandé)       │
         │  6. Ajouter heures de travail     │
         │     → Work → Units → Work (2x)    │
         └────────────────┬──────────────────┘
                          │
                          ↓
         ┌────────────────────────────┐
         │ Forçage final du Work      │ (2ème passage)
         └────────────────┬───────────┘
                          │
                          ↓
         ┌────────────────────────────┐
         │ Vérification Excel vs MSP  │ (Log comparaison)
         └────────────────┬───────────┘
                          │
                          ↓
         ┌────────────────────────────┐
         │ Réactiver calcul auto      │
         │ Calculer projet complet    │
         └────────────────┬───────────┘
                          │
                          ↓
         ┌────────────────────────────┐
         │ Fermer Excel               │
         │ Afficher "Import terminé!" │
         └────────────────────────────┘
```

---

## 💡 Conseils d'utilisation

### ✅ Bonnes pratiques

1. **Fichier Excel bien formaté** : Respecte les colonnes A à J
2. **Ligne 2 = Titre du projet** : Important !
3. **Données à partir de ligne 3** : Les tâches commencent là
4. **Heures en nombre** : 32 (pas "32h" ou "32 heures")
5. **Qualité en majuscules** : "CQ" ou "TACHE" (pas "cq" ou "Tache")

### 🔍 Vérifier que ça a marché

1. **Regarde le fichier log** : À côté de ton Excel, il y a un `.txt`
2. **Cherche les lignes "VERIFICATION FINALE"** : Compare Excel vs MS Project
3. **Vérifie la ressource "Monteurs"** : Le total d'heures doit être correct
4. **Vérifie les tâches** : Colonne "Travail" doit correspondre à ton Excel

### 🐛 Si ça ne marche pas

1. **Ouvre le fichier log** : Il contient tous les détails
2. **Cherche "ERREUR"** ou "IGNORÉ"** : Indices du problème
3. **Vérifie les colonnes Excel** : Bonnes données au bon endroit ?
4. **Vérifie le format des heures** : Nombre pur (pas de texte)

---

## 🎓 Vocabulaire MS Project

| Terme | Explication |
|-------|-------------|
| **Task** | Une tâche (une ligne de travail à faire) |
| **Assignment** | L'affectation d'une ressource à une tâche |
| **Resource** | Une personne ou du matériel |
| **Work** | Le travail total (en heures) |
| **Duration** | La durée calendaire (combien de jours) |
| **Units** | Le nombre de personnes (100% = 1 personne) |
| **Fixed Work** | Le travail est fixe, la durée s'adapte |

---

## 🎯 Résumé ultra-simplifié

**Ce que fait le programme en 5 phrases :**

1. Tu choisis ton fichier Excel
2. Il lit toutes les lignes (tâches, heures, personnes, etc.)
3. Il crée automatiquement un projet MS Project avec tout ça
4. Il force les bonnes valeurs d'heures (pour éviter les bugs de MS Project)
5. Il te donne un fichier log pour vérifier que tout est OK

**Et voilà, ton planning est prêt ! 🎉**

---

## 📞 Support

Si quelque chose ne fonctionne pas :
1. Ouvre le fichier log (`NomFichier_import_log.txt`)
2. Cherche les lignes avec "ERREUR" ou "IGNORÉ"
3. Vérifie que ton Excel est bien formaté
4. Relis la section "Problèmes résolus" ci-dessus

**Bonne utilisation ! 🚀**

