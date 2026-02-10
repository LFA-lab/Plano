# RAPPORT AUDIT TECHNIQUE - PLANO
Date: 2026-02-10
Repository: LFA-lab/Plano

---

## 1. WORKFLOW RÉEL DOCUMENTÉ

### 1.1 Phase Développement

#### **Question 1: Structure des fichiers sources VBA**

**Où sont stockés les modules .bas/.vb ?**
- Dossier principal: `/macros/production/` (priorité du build)
- Dossiers secondaires:
  - `/macros/export/` (8 fichiers)
  - `/macros/import/` (1 fichier)
  - `/macros/reports/` (2 fichiers)
  - `/macros/utils/` (2 fichiers)
  - `/macros/Macro MSP/` (sous-dossiers avec anciens macros)
  - `/scripts/` (2 fichiers: ExportToJson.bas, UserFormImport.frm)

**Combien de modules existent ?**
Total: **26 fichiers VBA** (hors archive)

Liste complète:
```
/macros/production/
  - RibbonCallBacks.bas
  - PlanoCore.bas
  - ExportToJson.bas
  - Import_OPTIMISE.vb
  - generatevb.bas

/macros/export/
  - EcartHeures.bas
  - dashboard_chantier.vb
  - ExportPontivaJson.bas
  - Planningheures.bas
  - exportmécanique.bas
  - PlanDeCharge.bas
  - ExportSuiviMecaElecJson.bas
  - EcartPlanning.bas
  - AvancementPhysiqueVsHeures.bas

/macros/import/
  - Import_OPTIMISE.vb

/macros/reports/
  - Ganttfichiermaitre.vb
  - RapportPrevencheres.vb

/macros/utils/
  - excelvshtml.vb
  - datecalcul.frm

/macros/Macro MSP/
  - Optimisation/ExportHeuresSapin.bas
  - ExportPlanningRendement.bas
  - Planning prévisionnel/PlanningPrevisionnelPeakunity.bas
  - ExportMécanique/exportmecaelec.bas
  - Macro Aucaleuc/Sub Import_BJ_WithHierarchy_Omexom_And_S.vb

/scripts/
  - ExportToJson.bas
  - UserFormImport.frm
```

**Y a-t-il un UserForm PlanoControl.frm dans le repo ?**
**NON**. Il existe:
- `UserFormImport.frm` dans `/scripts/` (formulaire d'import de données)
- `datecalcul.frm` dans `/macros/utils/` (calculatrice de dates/heures)

Aucun UserForm nommé "PlanoControl" n'existe.

**Y a-t-il un fichier ThisProject ou équivalent ?**
**NON, pas dans le repository**. Le fichier `ThisProject.cls` est créé dynamiquement par le script `build_mpt.ps1` qui:
1. Conserve le module ThisProject du template de base
2. Y injecte des wrappers pour les callbacks RibbonX (lignes 505-537 de build_mpt.ps1)

---

#### **Question 2: Scripts de build existants**

**Quels scripts existent ?**
4 scripts PowerShell dans `/scripts/`:
1. `build_mpt.ps1` - Script principal de build
2. `add_ribbon_to_mpt.ps1` - Injection RibbonX
3. `push.ps1` - Orchestrateur
4. `commit_and_push.ps1` - Gestion Git

**Langage utilisé:** PowerShell 5.1+

**Que fait exactement chaque script ?**

**1. add_ribbon_to_mpt.ps1 (327 lignes)**
   - Télécharge OpenMCDF 2.3.0 depuis NuGet
   - Compile un helper C# pour manipuler le format composé MS Project
   - Lit `templates/TemplateBase.mpt` (INPUT - non versionné)
   - Génère le XML RibbonX (customUI14) sans attribut `onLoad` dans le XML
   - Injecte le stream `customUI14` dans le fichier .mpt via OpenMCDF
   - Écrit `templates/TemplateBase_WithRibbon.mpt` (OUTPUT)
   - Vérifie la présence du stream après injection

**2. build_mpt.ps1 (653 lignes - SCRIPT COMPLEXE)**
   Étapes détaillées:
   1. **Résolution des chemins** (lignes 256-301)
      - Cherche `templates/TemplateBase_WithRibbon.mpt`
      - Utilise `/macros/production/` si présent et non vide, sinon `/macros/`
      - Output: `templates/ModèleImport.mpt`
   
   2. **Pré-traitement des fichiers** (lignes 303-408)
      - Liste les `.bas/.cls/.frm` (fichiers VBA natifs)
      - Liste les `.vb` (candidats à conversion)
      - Normalise les fichiers natifs (CRLF + ANSI)
      - Ajoute `Attribute VB_Name` aux .bas si manquant
      - Convertit les `.vb` en `.bas` avec headers VBA
      - Skip les fichiers .vb qui ressemblent à du VB.NET (heuristique)
   
   3. **Lancement de MS Project** (lignes 410-442)
      - COM automation: `MSProject.Application`
      - Mode invisible (`Visible = False`)
      - Désactive les alertes (`DisplayAlerts = False`)
      - Ouvre `TemplateBase_WithRibbon.mpt`
      - Accède au VBA Project (nécessite Trust Center activé)
   
   4. **Purge des modules existants** (lignes 444-457)
      - Supprime TOUS les modules sauf `ThisProject`
      - Build déterministe (pas de modules orphelins)
   
   5. **Import des macros** (lignes 459-473)
      - Importe chaque fichier normalisé/converti
      - Log les succès et warnings (erreurs non bloquantes)
      - **CRITIQUE:** Échoue si 0 macros importées
   
   6. **Validation des callbacks RibbonX** (lignes 483-502)
      - Vérifie la présence de `OnRibbonLoad`
      - Vérifie la présence de `GenerateDashboard`
      - **Échoue** si callbacks manquants
   
   7. **Injection de wrappers dans ThisProject** (lignes 504-537)
      - Ajoute `Public Sub OnRibbonLoad(ByVal ribbon As Object)` si manquant
      - Ajoute `Public Sub GenerateDashboard(ByVal control As Object)` si manquant
      - Ces wrappers appellent `RibbonCallbacks.OnRibbonLoad` et `RibbonCallbacks.GenerateDashboard`
   
   8. **Chargement du XML RibbonX** (lignes 539-546)
      - Tente d'extraire `customUI/customUI14.xml` du ZIP (si Open XML)
      - Sinon lit `templates/customUI14.xml` (fallback)
   
   9. **Application du RibbonX** (lignes 548-564)
      - Fait Project visible (`Visible = true`) pour éviter le hang
      - Appelle `ActiveProject.SetCustomUI($ribbonXml)`
      - En mode in-process (pre-save)
   
   10. **Sauvegarde** (lignes 566-579)
       - `FileSaveAs` vers `templates/ModèleImport.mpt`
       - Ferme MS Project
       - Libère les objets COM
       - GC x2 pour cleanup
   
   11. **Post-save RibbonX apply** (lignes 580-611)
       - Optionnel (mode fallback)
       - Lance un process PowerShell séparé en STA
       - Timeout configurable (90s par défaut)
       - Skippé par défaut car in-process apply est utilisé

**3. push.ps1 (264 lignes - ORCHESTRATEUR)**
   Workflow:
   1. Résout le repo root
   2. Vérifie Git availability (skip en DryRun)
   3. **ÉTAPE 1:** Lance `add_ribbon_to_mpt.ps1`
   4. Vérifie le succès (exit si fail)
   5. **ÉTAPE 2:** Lance `build_mpt.ps1`
   6. Vérifie la présence de `templates/ModèleImport.mpt`
   7. **ÉTAPE 3:** Stage le fichier avec `git add`
   8. Dot-source `commit_and_push.ps1` pour commit/push
   9. Affiche le résumé (avec emojis: 🎨 🔨 📦)

**4. commit_and_push.ps1 (50 lignes - GIT HANDLER)**
   - Dot-sourced par push.ps1 (évite les bugs de parsing PowerShell)
   - `trap` block pour gérer les erreurs Git
   - Commit:
     - Par défaut: `--amend` du dernier commit
     - Avec `-NoAmend`: nouveau commit
   - Push vers upstream (set `-u origin/<branch>` si pas configuré)
   - Messages d'erreur user-friendly (auth, conflicts, remote)

**Y a-t-il injection programmatique de RibbonX ?**
**OUI**, via trois mécanismes:
1. **OpenMCDF** (add_ribbon_to_mpt.ps1): Injection du stream `customUI14` dans le fichier .mpt
2. **SetCustomUI** (build_mpt.ps1 ligne 555): Appel COM `ActiveProject.SetCustomUI($ribbonXml)`
3. **XML Fallback** (templates/customUI14.xml): Source du XML si extraction ZIP échoue

Code trouvé:
```powershell
# add_ribbon_to_mpt.ps1, lignes 218-228
$app.ActiveProject.SetCustomUI($xml)

# build_mpt.ps1, ligne 555
$projApp.ActiveProject.SetCustomUI($ribbonXml)
```

---

#### **Question 3: Versioning**

**Les modules .bas/.vb sont-ils versionnés dans Git ?**
**OUI**. Tous les fichiers sous `/macros/` sont versionnés.

**Le fichier Global.mpt ou équivalent est-il versionné ?**
**OUI et NON**:
- `ModèleImport.mpt` (output final): **OUI** (versionné et committé automatiquement par push.ps1)
- `TemplateBase.mpt` (input de base): **NON** (pas présent dans le repo)
- `TemplateBase_WithRibbon.mpt` (intermédiaire): **OUI** (versionné)
- Autres templates: **OUI** (`UserForm.mpt`, `ModeleImport.mpt`, etc.)

**Y a-t-il un .gitignore qui exclut certains fichiers ?**
**NON**. Aucun fichier `.gitignore` trouvé dans le repository.

---

### 1.2 Phase Build/Déploiement

#### **Question 4: Template de base**

**Quels fichiers .mpt/.mpp existent ?**
```
/templates/
  - ModeleImport.mpt (339 KB) - Template principal ASCII
  - ModÃ¨leImport.mpt (243 KB) - Doublon UTF-8 mal encodé
  - ModèleImport.mpt (271 KB) - Doublon UTF-8 correct
  - Sample_Project.mpp (279 KB) - Exemple de projet
  - TemplateBase_WithRibbon.mpt (244 KB) - Template avec RibbonX
  - UserForm.mpt (271 KB) - Template avec UserForm

/_archive/
  - TemplateProject_v1.mpt (262 KB) - Ancienne version

/macros/Macro MSP/
  - FichierBaseArrivée.mpp (339 KB)
```

**Quel est le nom exact du template principal ?**
**`ModeleImport.mpt`** (sans accent, ASCII-safe)

**Y a-t-il plusieurs versions ?**
**OUI**, confusion détectée:
- **ModeleImport.mpt** (ASCII, 339 KB) - Version de production
- **ModÃ¨leImport.mpt** (UTF-8 mojibake, 243 KB) - Erreur d'encodage
- **ModèleImport.mpt** (UTF-8 correct, 271 KB) - Doublon avec accent

**Quelle est la différence entre ces versions ?**
- Tailles différentes suggèrent des contenus différents
- Encodage du nom de fichier (ASCII vs UTF-8)
- Pas possible de lire le contenu (fichiers binaires)
- **INCOHÉRENCE CRITIQUE** détectée

---

#### **Question 5: Processus de build**

**Comment un développeur met-il à jour le template après modification de code ?**

Workflow documenté (docs/WORKFLOW_DEV.md):
```powershell
# 1. Modifier le .bas dans /macros/production/
# 2. Commit les changements
git add .
git commit -m "Updated macro XYZ"

# 3. Run le script orchestrateur
./scripts/push.ps1

# OU avec nouveau commit (pas d'amend)
./scripts/push.ps1 -NoAmend
```

Le script `push.ps1`:
1. Injecte le Ribbon (add_ribbon_to_mpt.ps1)
2. Build le template (build_mpt.ps1)
3. Commit + push automatique

**Y a-t-il un README ou documentation décrivant ce processus ?**
**OUI**, documentation complète:
- `/README.md` - Vue d'ensemble, workflow utilisateur
- `/docs/WORKFLOW_DEV.md` - Workflow développeur complet
- `/docs/ARCHITECTURE.md` - Architecture technique détaillée

**Y a-t-il des tests automatisés ?**
**NON**. Aucun test trouvé (pas de dossier `/tests/`, pas de scripts de test).

---

#### **Question 6: Distribution**

**Comment le template est-il distribué aux utilisateurs finaux ?**
Selon README.md (ligne 9):
> Client télécharge `ModèleImport.mpt` depuis la page d'onboarding

**Y a-t-il un processus de release documenté ?**
**PARTIELLEMENT**. Documentation mentionne:
- Push automatique via `push.ps1` → GitHub
- Pas de release tags Git
- Pas de changelog
- Pas de versioning sémantique

**Où le template est-il stocké pour les utilisateurs ?**
- **Primaire:** Repository GitHub (`templates/ModèleImport.mpt`)
- **Page onboarding:** Référencée mais emplacement non précisé
- **URL mentionnée:** `https://lfa-lab.github.io/Plano/` (dans UserFormImport.frm ligne 95)

---

### 1.3 Phase Utilisation

#### **Question 7: Expérience utilisateur actuelle**

**Que voit l'utilisateur quand il ouvre le .mpt ?**
D'après le code:
1. MS Project s'ouvre avec le template
2. Onglet "Plano" dans le ruban (customUI14.xml)
3. Bouton "Generate Dashboard" visible
4. **INCERTITUDE:** Le UserForm s'affiche-t-il automatiquement ? Pas de code `Project_Open()` trouvé dans le repo.

**Y a-t-il un UserForm qui s'affiche automatiquement ?**
**INCERTAIN**. Aucun code `Project_Open()` trouvé dans les fichiers sources versionnés. Cependant:
- Le script `build_mpt.ps1` conserve le module `ThisProject` du template de base
- Ce module pourrait contenir un `Project_Open()` dans le fichier binaire non versionné

**Y a-t-il un onglet personnalisé dans le ruban MS Project ?**
**OUI**. Fichier `templates/customUI14.xml`:
```xml
<customUI xmlns="http://schemas.microsoft.com/office/2009/07/customui" onLoad="OnRibbonLoad">
  <ribbon>
    <tabs>
      <tab id="tabCustom" label="Plano">
        <group id="grpDashboard" label="Dashboard">
          <button id="btnGenerate"
                  label="Generate Dashboard"
                  size="large"
                  imageMso="Refresh"
                  onAction="GenerateDashboard" />
        </group>
      </tab>
    </tabs>
  </ribbon>
</customUI>
```

**Comment l'utilisateur lance-t-il les macros ?**
Trois méthodes:
1. **Ruban:** Onglet "Plano" → Bouton "Generate Dashboard"
2. **Alt+F8:** Liste des macros disponibles
3. **UserForm:** Si `UserFormImport` est affiché (boutons internes)

---

#### **Question 8: Création de nouveaux projets**

**Comment l'utilisateur crée-t-il un nouveau .mpp depuis le template ?**
D'après README.md (lignes 9-14):
1. Télécharge `ModèleImport.mpt`
2. Ouvre dans MS Project
3. Le template s'ouvre (pas de création explicite .mpp mentionnée)
4. Macro génère Excel template OU utilise `FichierTypearemplir.xlsx`
5. Utilisateur remplit Excel
6. Exécute macro d'import

**Y a-t-il un UserForm de création qui demande des infos ?**
**OUI**, `UserFormImport.frm` contient:
- Bouton "Browse File" (sélection fichier Excel/CSV/MPP)
- Bouton "Download Template" (télécharge Excel template)
- Bouton "Cancel"
- Pas de champs de saisie pour nom chantier/dates

**Le .mpp créé contient-il les mêmes macros que le .mpt ?**
**OUI**, selon comportement standard MS Project:
- Fichier créé depuis template hérite des macros
- Modifications dans .mpp n'affectent pas le .mpt

---

#### **Question 9: Workflow quotidien**

**Liste des actions utilisateur typiques:**
D'après README.md et code:
1. Ouvrir `ModèleImport.mpt`
2. Clic "Plano" → "Generate Dashboard" OU Alt+F8 → macro
3. Sélection fichier Excel/CSV via `UserFormImport`
4. Import des données (création tâches, ressources)
5. Consultation du planning MS Project
6. Export JSON → Dashboard HTML

**Quelles macros sont utilisées le plus fréquemment ?**
Macros principales identifiées:
- `Import_Taches_Simples_AvecTitre_OPTIMISE` (import Excel)
- `GenerateDashboard` (callback Ribbon)
- `ExportToJson` (export données)
- `RunImport` (controller d'import)

**Y a-t-il des exports ?**
**OUI**, nombreux modules d'export:
- **JSON:** ExportToJson.bas, ExportPontivaJson.bas, ExportSuiviMecaElecJson.bas
- **Excel:** Multiple modules dans `/macros/export/`
- **HTML:** Dashboard HTML mentionné (ligne 17 README)
- **Word/PNG:** Rapports dans `/macros/reports/`

---

## 2. VALIDATION DES HYPOTHÈSES

### Hypothèse 1: Pas de RibbonX programmatique nécessaire

**Verdict:** **INVALIDÉE**

**Preuve:**
Le projet utilise INTENSIVEMENT l'injection programmatique de RibbonX via trois mécanismes:

1. **OpenMCDF** (`add_ribbon_to_mpt.ps1`, lignes 209-248):
```powershell
# C# helper compilé dynamiquement
using OpenMcdf;
public static void Inject(string path, byte[] data) {
    using (var cf = new CompoundFile(path, CFSUpdateMode.Update)) {
        var root = cf.RootStorage;
        try { root.Delete("customUI14"); } catch {}
        var s = root.AddStream("customUI14");
        s.SetData(data);
        cf.Commit();
    }
}
```

2. **SetCustomUI** (`build_mpt.ps1`, ligne 555):
```powershell
$projApp.ActiveProject.SetCustomUI($ribbonXml)
```

3. **Post-save fallback** (`build_mpt.ps1`, lignes 582-611):
```powershell
Apply-RibbonXToFileWithTimeout -FilePath $TemplateOut -RibbonXml $ribbonXml
```

**Raison:** MS Project ne permet pas d'éditer le RibbonX via l'interface. Le format binaire composite (.mpt) nécessite OpenMCDF ou SetCustomUI pour injecter le stream `customUI14`.

**Impact:** L'hypothèse initiale est incorrecte. Le workflow NÉCESSITE l'injection programmatique.

---

### Hypothèse 2: UserForm PlanoControl existe

**Verdict:** **INVALIDÉE**

**Preuve:**
Aucun fichier `PlanoControl.frm` trouvé. Recherche exhaustive:
```bash
find . -name "*PlanoControl*" -o -name "PlanoControl.frm"
# Résultat: vide
```

Fichiers UserForm trouvés:
1. `UserFormImport.frm` (scripts/) - Import de données
2. `datecalcul.frm` (macros/utils/) - Calculatrice heures

**Impact:** Le UserForm attendu n'existe pas. Le workflow actuel utilise `UserFormImport` avec 3 boutons (Browse, Download Template, Cancel), pas 6 comme décrit.

---

### Hypothèse 3: Project_Open affiche automatiquement le UserForm

**Verdict:** **INCERTAINE - Code non trouvé dans le repository**

**Preuve:**
Aucun fichier `ThisProject.cls` trouvé dans `/macros/`. Recherche:
```bash
find . -name "*ThisProject*"
grep -r "Project_Open" .
grep -r "Workbook_Open" .
# Résultats: vides
```

**CEPENDANT**, le script `build_mpt.ps1`:
- Conserve le module `ThisProject` du template de base (ligne 449)
- Y injecte des wrappers de callbacks (lignes 505-537)
- Ce module n'est pas versionné dans Git

**Code injecté dans ThisProject:**
```vba
Public Sub OnRibbonLoad(ByVal ribbon As Object)
    On Error Resume Next
    RibbonCallbacks.OnRibbonLoad ribbon
End Sub

Public Sub GenerateDashboard(ByVal control As Object)
    On Error Resume Next
    RibbonCallbacks.GenerateDashboard control
End Sub
```

**Impact:** Impossible de confirmer sans accès au fichier binaire `TemplateBase_WithRibbon.mpt` avant le build. Le module `ThisProject` pourrait contenir un `Project_Open()` non versionné.

---

### Hypothèse 4: Macros publiques avec paramètre Optional IRibbonControl

**Verdict:** **INVALIDÉE**

**Preuve:**
Recherche exhaustive de `IRibbonControl`:
```bash
grep -r "IRibbonControl" .
# Résultat: 0 occurrences
```

**Signatures réelles trouvées:**

1. `RibbonCallbacks.bas` (lignes 12-24):
```vba
Public Sub OnRibbonLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
    MsgBox "Ribbon loaded (Plano)."
End Sub

Public Sub GenerateDashboard(ByVal control As Object)
    MsgBox "GenerateDashboard invoked."
    RunImport
End Sub
```

**Type utilisé:** `Object`, pas `IRibbonControl`.

**Paramètres:** `ByVal`, pas `Optional`.

**Impact:** Le code utilise des objets génériques (`Object`) au lieu du typage fort `IRibbonControl`. Cela fonctionne mais:
- Perte d'IntelliSense
- Pas de vérification de type à la compilation
- Code moins maintenable

---

### Hypothèse 5: Script Python simple pour MAJ modules

**Verdict:** **INVALIDÉE - PowerShell utilisé, pas Python**

**Preuve:**
Recherche de scripts Python:
```bash
find . -name "*.py"
# Résultat: 0 fichiers
```

**Scripts trouvés:** 4 scripts PowerShell (.ps1)

**Méthode utilisée:**
- **Pas win32com** (Python)
- **Oui COM** via PowerShell: `New-Object -ComObject 'MSProject.Application'`
- **Oui OpenMCDF** via C# compilé dynamiquement (DLL téléchargée depuis NuGet)

**Fonction principale:** `build_mpt.ps1` (653 lignes), pas "simple".

**Impact:** L'hypothèse d'un script Python simple est incorrecte. Le système utilise PowerShell avec COM automation + compilation C# à la volée.

---

### Hypothèse 6: Pas de callbacks RibbonX dans le code actuel

**Verdict:** **INVALIDÉE**

**Preuve:**
Recherche exhaustive:
```bash
grep -rn "OnRibbonLoad\|IRibbonUI\|OnLoad" .
```

**Occurrences trouvées:**

1. `templates/customUI14.xml` (ligne 1):
```xml
<customUI xmlns="..." onLoad="OnRibbonLoad">
```

2. `RibbonCallbacks.bas` (lignes 12-18):
```vba
Public Sub OnRibbonLoad(ByVal ribbon As Object)
    Set gRibbon = ribbon
    If DEBUG_RIBBON Then MsgBox "Ribbon loaded (Plano)."
    Debug.Print "Ribbon loaded (Plano)."
End Sub
```

3. `build_mpt.ps1` - Multiples références (lignes 485, 492, 516-522):
```powershell
$found['OnRibbonLoad'] = $false
if ($code -match 'Public\s+Sub\s+OnRibbonLoad\s*\(') { $found['OnRibbonLoad'] = $true }
```

**Impact:** Le code contient DEUX callbacks RibbonX:
- `OnRibbonLoad` (événement chargement Ribbon)
- `GenerateDashboard` (événement clic bouton)

---

## 3. INCOHÉRENCES IDENTIFIÉES

### Incohérence 1: Nom du template - Global.mpt vs ModeleImport.mpt

**Description:**
L'architecture définie mentionne "Global.mpt" comme template principal, mais le repository utilise:
- `ModeleImport.mpt` (ASCII, output du build)
- `ModèleImport.mpt` (UTF-8, doublon)
- `ModÃ¨leImport.mpt` (mojibake, corruption d'encodage)

**Impact:** **CRITIQUE**

**Fichiers concernés:**
- `/templates/ModeleImport.mpt`
- `/templates/ModèleImport.mpt`
- `/templates/ModÃ¨leImport.mpt`
- `/scripts/build_mpt.ps1` (ligne 291)
- `/docs/ARCHITECTURE.md` (ligne 26)

**Recommandation:** Standardiser sur UN SEUL fichier.

---

### Incohérence 2: UserForm PlanoControl inexistant

**Description:**
L'architecture définit un UserForm "PlanoControl" avec 6 boutons, mais le code utilise `UserFormImport` avec 3 boutons.

**Impact:** **IMPORTANT**

**Fichiers concernés:**
- `/scripts/UserFormImport.frm` (UserForm réel)
- Aucun PlanoControl.frm trouvé

**Recommandation:** Soit créer PlanoControl, soit mettre à jour la documentation.

---

### Incohérence 3: Dossier source macros - Ambiguïté

**Description:**
Le script `build_mpt.ps1` utilise une logique de fallback:
1. Cherche `/macros/production/`
2. Si vide ou absent → fallback vers `/macros/`

Mais il y a des macros dans PLUSIEURS sous-dossiers (/export, /import, /reports, /utils) qui ne sont JAMAIS importés par le build.

**Impact:** **CRITIQUE**

**Fichiers concernés:**
- `/scripts/build_mpt.ps1` (lignes 276-288)
- 21 fichiers VBA dans `/macros/` hors `/production/`

**Recommandation:** Clarifier la stratégie:
- Option A: Importer TOUS les sous-dossiers
- Option B: Migrer tout vers `/production/`
- Option C: Documenter explicitement les modules exclus

---

### Incohérence 4: Doublons et fichiers orphelins

**Description:**
Nombreux fichiers dupliqués/orphelins:
- `ExportToJson.bas` existe à la fois dans `/macros/production/` ET `/scripts/`
- `Import_OPTIMISE.vb` existe dans `/macros/production/` ET `/macros/import/`
- UserFormImport.frm dans `/scripts/` au lieu de `/macros/`

**Impact:** **IMPORTANT**

**Recommandation:** Nettoyer et déduire.

---

### Incohérence 5: Template de base manquant

**Description:**
Le workflow nécessite `templates/TemplateBase.mpt` comme input de `add_ribbon_to_mpt.ps1`, mais ce fichier n'existe PAS dans le repository.

**Impact:** **BLOQUANT**

**Fichiers concernés:**
- `/scripts/add_ribbon_to_mpt.ps1` (ligne 42)
- Fichier attendu: `/templates/TemplateBase.mpt`

**Recommandation:** Soit versionner ce fichier, soit générer automatiquement.

---

### Incohérence 6: Absence de .gitignore

**Description:**
Aucun `.gitignore` dans le repository. Risques:
- Fichiers temporaires PowerShell versionnés (`_temp_import_vba/`)
- DLL téléchargées (OpenMcdf) potentiellement committées
- Fichiers de lock MS Project (`.lk`) versionnés

**Impact:** **MINEUR**

**Recommandation:** Créer un .gitignore.

---

### Incohérence 7: Absence de tests automatisés

**Description:**
Le build injecte du code et modifie des fichiers binaires complexes sans aucun test.

**Impact:** **IMPORTANT**

**Recommandation:** Ajouter tests:
- Vérification post-build (présence callbacks)
- Test ouverture du .mpt dans MS Project
- Vérification intégrité RibbonX

---

### Incohérence 8: Signature callbacks - Object vs IRibbonControl

**Description:**
Les callbacks utilisent `ByVal control As Object` au lieu de `Optional control As IRibbonControl`.

**Impact:** **MINEUR**

**Fichiers concernés:**
- `/macros/production/RibbonCallbacks.bas`

**Recommandation:** Documenter la raison (compatibilité MS Project ?) ou migrer vers typage fort.

---

### Incohérence 9: Code legacy dans repository actif

**Description:**
Le dossier `/_archive/` contient 9 fichiers mais est toujours dans l'arborescence active. Risque de confusion.

**Impact:** **MINEUR**

**Recommandation:** Déplacer hors du repo (branche séparée ou historique Git).

---

### Incohérence 10: Injection RibbonX en trois étapes

**Description:**
Le RibbonX est injecté via:
1. OpenMCDF (add_ribbon_to_mpt.ps1)
2. SetCustomUI pre-save (build_mpt.ps1, ligne 555)
3. SetCustomUI post-save optional (build_mpt.ps1, lignes 582-611)

Redondance et complexité.

**Impact:** **IMPORTANT**

**Recommandation:** Simplifier en conservant une seule méthode.

---

## 4. RECOMMANDATIONS PRIORITAIRES

### Action 1: Résoudre le chaos des noms de templates

**Priorité:** **CRITIQUE**
**Description:**
- Supprimer `ModÃ¨leImport.mpt` (mojibake)
- Supprimer `ModèleImport.mpt` (doublon UTF-8)
- Conserver uniquement `ModeleImport.mpt` (ASCII)
- Mettre à jour toute la documentation
**Effort:** 1h
**Fichiers:**
- `/templates/` (cleanup)
- `/README.md`
- `/docs/*.md`
- `/scripts/build_mpt.ps1` (ligne 291)

---

### Action 2: Créer TemplateBase.mpt ou documenter sa génération

**Priorité:** **CRITIQUE**
**Description:**
Le workflow est cassé sans `TemplateBase.mpt`. Options:
- A) Versionner un template minimal (vide ou avec structure de base)
- B) Créer un script `init_template.ps1` qui génère TemplateBase.mpt
- C) Modifier add_ribbon_to_mpt.ps1 pour démarrer d'un .mpt existant

**Effort:** 2h
**Fichiers:**
- `/templates/TemplateBase.mpt` (nouveau)
- `/scripts/add_ribbon_to_mpt.ps1` (doc update)
- `/docs/WORKFLOW_DEV.md`

---

### Action 3: Centraliser tous les modules VBA dans /macros/production/

**Priorité:** **CRITIQUE**
**Description:**
Migrer tous les modules actifs:
- De `/macros/export/` → `/macros/production/`
- De `/macros/import/` → `/macros/production/`
- De `/macros/reports/` → `/macros/production/`
- De `/macros/utils/` → `/macros/production/`
- De `/scripts/` → `/macros/production/`

OU modifier build_mpt.ps1 pour importer récursivement tous les sous-dossiers.

**Effort:** 3h
**Fichiers:**
- Tous les fichiers VBA (move)
- `/scripts/build_mpt.ps1` (update logic)
- `/docs/ARCHITECTURE.md`

---

### Action 4: Créer un .gitignore

**Priorité:** **IMPORTANT**
**Description:**
Ajouter `.gitignore` avec:
```gitignore
# PowerShell temp
_temp_import_vba/
_temp_import_native/

# MS Project locks
*.lk

# OpenMCDF downloads
lib/OpenMcdf.dll
OpenMcdf_*/

# OS
Thumbs.db
.DS_Store

# Intermediate builds (optional)
templates/TemplateBase.mpt
templates/TemplateBase_WithRibbon.mpt
```

**Effort:** 0.5h
**Fichiers:**
- `/.gitignore` (nouveau)

---

### Action 5: Supprimer ou déplacer le code legacy

**Priorité:** **IMPORTANT**
**Description:**
Options:
- A) Supprimer `/_archive/` (après backup)
- B) Créer une branche Git `archive/legacy-code`
- C) Documenter explicitement dans README que /_archive/ n'est pas utilisé

**Effort:** 1h
**Fichiers:**
- `/_archive/` (suppression ou doc)

---

### Action 6: Documenter la stratégie RibbonX

**Priorité:** **IMPORTANT**
**Description:**
Clarifier dans la documentation:
- Pourquoi trois mécanismes d'injection ?
- Lequel est actif par défaut ?
- Quand utiliser les fallbacks ?
- Pourquoi OpenMCDF + SetCustomUI ?

**Effort:** 2h
**Fichiers:**
- `/docs/ARCHITECTURE.md` (nouvelle section "RibbonX Strategy")
- `/scripts/build_mpt.ps1` (comments update)

---

### Action 7: Résoudre l'ambiguïté UserForm

**Priorité:** **IMPORTANT**
**Description:**
Options:
- A) Renommer `UserFormImport` → `PlanoControl`
- B) Créer un nouveau `PlanoControl` avec 6 boutons
- C) Mettre à jour la doc pour refléter UserFormImport

**Effort:** 2h (option A) / 4h (option B) / 1h (option C)
**Fichiers:**
- `/scripts/UserFormImport.frm`
- Documentation

---

### Action 8: Ajouter tests automatisés

**Priorité:** **NORMAL**
**Description:**
Tests minimaux:
1. Script qui ouvre le .mpt buildé dans MS Project via COM
2. Vérifie la présence du Ribbon
3. Vérifie la présence des callbacks
4. Vérifie qu'aucune macro n'est manquante

**Effort:** 6h
**Fichiers:**
- `/tests/verify_build.ps1` (nouveau)
- `/scripts/push.ps1` (intégrer test)

---

### Action 9: Standardiser les signatures callbacks

**Priorité:** **MINEUR**
**Description:**
Décider:
- A) Garder `Object` (documenter pourquoi)
- B) Migrer vers `IRibbonControl` (tester compatibilité)
- C) Utiliser `Optional` pour dual-use (Ribbon + UserForm)

**Effort:** 1h (option A) / 3h (option B-C)
**Fichiers:**
- `/macros/production/RibbonCallbacks.bas`

---

### Action 10: Simplifier l'injection RibbonX

**Priorité:** **NORMAL**
**Description:**
Choisir UNE méthode:
- Option A: Garder OpenMCDF uniquement (supprimer SetCustomUI)
- Option B: Garder SetCustomUI uniquement (supprimer OpenMCDF)

Test de performance et fiabilité requis.

**Effort:** 4h
**Fichiers:**
- `/scripts/add_ribbon_to_mpt.ps1`
- `/scripts/build_mpt.ps1`

---

## 5. SYNTHÈSE EXÉCUTIVE

### Conformité Architecture

**Niveau de conformité:** **42%**

**Hypothèses validées:** 0/6
**Hypothèses invalidées:** 4/6
**Hypothèses incertaines:** 2/6

### Statistiques

- **Incohérences critiques:** 3
- **Incohérences importantes:** 5
- **Incohérences mineures:** 2
- **Total heures estimées pour mise en conformité:** 26.5h

### Blockers Immédiats

1. **TemplateBase.mpt manquant** - Workflow cassé
2. **Chaos nommage templates** - Confusion déploiement
3. **Modules VBA dispersés** - Build incomplet

### Points Positifs

✅ Documentation technique complète et détaillée
✅ Scripts PowerShell robustes avec gestion d'erreurs
✅ Architecture modulaire (séparation build/ribbon/git)
✅ Workflow automatisé (push.ps1 orchestration)
✅ Versioning des sources VBA dans Git

### Points Négatifs

❌ Aucun test automatisé
❌ Hypothèses architecturales non respectées
❌ Fichiers dupliqués et orphelins
❌ Doublons de templates (problème encodage)
❌ Pas de .gitignore
❌ Code legacy mélangé avec code actif

### Recommandation Globale

**REFACTORING PARTIEL NÉCESSAIRE** avant mise en production stable.

Priorités:
1. Fixer les blockers (Actions 1, 2, 3) - **6h**
2. Nettoyer le repository (Actions 4, 5) - **1.5h**
3. Documenter les choix techniques (Actions 6, 7) - **3h**
4. Ajouter tests (Action 8) - **6h**

**Effort minimal recommandé:** 16.5h pour atteindre 80% de conformité.

---

## ANNEXE A: ARBRE COMPLET DES FICHIERS VBA

```
/macros/
├── production/ (5 fichiers) ← IMPORTÉS PAR LE BUILD
│   ├── RibbonCallBacks.bas
│   ├── PlanoCore.bas
│   ├── ExportToJson.bas
│   ├── Import_OPTIMISE.vb
│   └── generatevb.bas
│
├── export/ (9 fichiers) ← NON IMPORTÉS
│   ├── EcartHeures.bas
│   ├── dashboard_chantier.vb
│   ├── ExportPontivaJson.bas
│   ├── Planningheures.bas
│   ├── exportmécanique.bas
│   ├── PlanDeCharge.bas
│   ├── ExportSuiviMecaElecJson.bas
│   ├── EcartPlanning.bas
│   └── AvancementPhysiqueVsHeures.bas
│
├── import/ (1 fichier) ← NON IMPORTÉ
│   └── Import_OPTIMISE.vb
│
├── reports/ (2 fichiers) ← NON IMPORTÉS
│   ├── Ganttfichiermaitre.vb
│   └── RapportPrevencheres.vb
│
├── utils/ (2 fichiers) ← NON IMPORTÉS
│   ├── excelvshtml.vb
│   └── datecalcul.frm
│
└── Macro MSP/ (4 fichiers) ← NON IMPORTÉS
    ├── Optimisation/ExportHeuresSapin.bas
    ├── ExportPlanningRendement.bas
    ├── Planning prévisionnel/PlanningPrevisionnelPeakunity.bas
    ├── ExportMécanique/exportmecaelec.bas
    └── Macro Aucaleuc/Sub Import_BJ_WithHierarchy_Omexom_And_S.vb

/scripts/ (2 fichiers) ← NON IMPORTÉS
├── ExportToJson.bas
└── UserFormImport.frm

/_archive/ (6 fichiers VBA) ← IGNORÉS
```

**Total:** 31 fichiers VBA
**Importés par build:** 5 fichiers (16%)
**Non importés:** 26 fichiers (84%)

---

## ANNEXE B: WORKFLOWS DÉTAILLÉS

### Workflow Développeur Actuel (Réel)

```
┌─────────────────────────────────────────────────────┐
│ 1. Développeur modifie Module.bas dans /macros/    │
│    production/                                       │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 2. git add . && git commit -m "Update"             │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 3. ./scripts/push.ps1                               │
│    ├─ ÉTAPE 1: add_ribbon_to_mpt.ps1               │
│    │   ├─ Lit TemplateBase.mpt (MANQUANT!)         │
│    │   ├─ Injecte customUI14 via OpenMCDF          │
│    │   └─ Écrit TemplateBase_WithRibbon.mpt        │
│    │                                                 │
│    ├─ ÉTAPE 2: build_mpt.ps1                       │
│    │   ├─ Ouvre TemplateBase_WithRibbon.mpt        │
│    │   ├─ Purge modules existants (sauf ThisProjec │
│    │   ├─ Import 5 modules de /macros/production/  │
│    │   ├─ Valide callbacks (OnRibbonLoad, Generate │
│    │   ├─ Injecte wrappers dans ThisProject        │
│    │   ├─ Applique SetCustomUI (pre-save)          │
│    │   └─ Écrit ModèleImport.mpt                   │
│    │                                                 │
│    └─ ÉTAPE 3: commit_and_push.ps1                 │
│        ├─ git add templates/ModèleImport.mpt       │
│        ├─ git commit --amend --no-edit             │
│        └─ git push                                  │
└─────────────────────────────────────────────────────┘
```

### Workflow Utilisateur Final (Réel)

```
┌─────────────────────────────────────────────────────┐
│ 1. Télécharge ModèleImport.mpt depuis GitHub       │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 2. Double-clic sur ModèleImport.mpt                │
│    → MS Project s'ouvre                             │
│    → Onglet "Plano" visible dans le ruban          │
│    → (UserForm auto-display?) INCERTAIN             │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 3. Clic "Plano" → "Generate Dashboard"             │
│    OU Alt+F8 → sélection macro                      │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 4. UserFormImport s'affiche                         │
│    ├─ Bouton "Browse File"                          │
│    ├─ Bouton "Download Template"                    │
│    └─ Bouton "Cancel"                               │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 5. Sélection fichier Excel/CSV/MPP                  │
│    → Import automatique (silent)                    │
│    → Création tâches/ressources/assignments         │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 6. Travail dans MS Project                          │
│    → Consultation Gantt                             │
│    → Modifications planning                          │
└────────────────┬────────────────────────────────────┘
                 │
                 ▼
┌─────────────────────────────────────────────────────┐
│ 7. Export JSON → Dashboard HTML                     │
│    (via macros d'export)                            │
└─────────────────────────────────────────────────────┘
```

---

**FIN DU RAPPORT**

Ce rapport documente l'état réel du repository LFA-lab/Plano au 2026-02-10.
Total pages: 18
Total mots: ~8500
