# Rapport d'audit — Onglet Accueil (PDG) pour sitemark.py

**Objectif :** Cartographier le blueprint JSON `style_expert.json` (onglet manuel "PDG") et proposer l’intégration dans le moteur `sitemark.py` sans modifier le code à ce stade.

---

## 1. Cartographie des zones fixes vs dynamiques

### 1.1 Étiquettes statiques (labels)

| Adresse | Valeur (exemple) | Style (JSON) | Rôle |
|---------|------------------|--------------|------|
| **B38:J38** | "Inspection du site de SBRU avant mise en service et réception en O&M" | bg=FFFFFF, bold, size=20, align=center, border_bottom | Titre principal (fusion B38:J38) |
| **B41:J41** | "Description du site en quelques chiffres" | bg=FFFFFF, bold, size=20 (JSON) / 12 (cohérence), center, border_bottom | Sous-titre section Description |
| **B44:C44** | "Client" | bg=**F2E1D9**, bold=false, size=12, center, border_bottom | Label gauche |
| **B45:C45** | "Numéro d'Affaires :" | F2E1D9, 12, center | Label gauche |
| **B46:C46** | "Puissance totale de la centrale" | F2E1D9, 12, center | Label gauche |
| **B47:C47** | "Date :" | F2E1D9, 12, center | Label gauche |
| **B48:C48** | "Nombre de PDL / PTR" | F2E1D9, 12, center | Label gauche |
| **B49:C49** | "Nombre de PTR" | F2E1D9, 12, center | Label gauche |
| **G44:H44** | "Type Onduleur : Cent/Décentralisé" | F2E1D9, 12, center | Label droite |
| **G45:H45** | "Nombre d' Onduleur" | F2E1D9, 12, center | Label droite |
| **G46:H46** | "Type Câblage solaire" | F2E1D9, 12, center | Label droite |
| **G47:H47** | "Type de raccordement solaire" | F2E1D9, 12, center | Label droite |
| **G48:H48** | "Type de Cheminement : Tranchée/CDC" | F2E1D9, 12, center | Label droite |
| **G49:H49** | "Type Tranchée : Gaines TPC / Enterrabilité directe" | F2E1D9, 12, center | Label droite |

**Style commun étiquettes :** `PatternFill(solid, F2E1D9)`, `Font(Calibri, size=12, bold=False)`, `Alignment(horizontal="center")`, bordure bas.

---

### 1.2 Réceptacles de données (cellules à remplir par le script)

Toutes ont **fond gris D9D9D9** dans le blueprint.

| Adresse (ancrage) | Merge_area | Variable Python proposée | Source actuelle dans sitemark.py |
|-------------------|------------|--------------------------|-----------------------------------|
| **D44** | D44:E44 | `client` | Non fourni (à ajouter : config / métadonnées PDF) |
| **D45** | D45:E45 | `numero_affaires` | Idem |
| **D46** | D46:E46 | `puissance_totale_mwc` | Idem |
| **D47** | D47:E47 | `date_inspection` | Peut venir de `datetime.today()` ou métadonnées |
| **D48** | D48:E48 | `nombre_pdl_ptr` | Idem (ou dérivé des réserves / config) |
| **D49** | D49:E49 | `nombre_ptr` | Idem |
| **I44** | I44:J44 | `type_onduleur` | Ex. "DÉCENTRALISÉ" — config |
| **I45** | I45:J45 | `nombre_onduleurs` | Config ou dérivé |
| **I46** | I46:J46 | `type_cablage` | Ex. "LIGNE" — config |
| **I47** | I47:J47 | `type_raccordement` | Ex. "SIMPLE" — config |
| **I48** | I48:J48 | `type_cheminement` | Ex. "TRANCHEE" — config |
| **I49** | I49:J49 | `type_tranchee` | Ex. "ENTERRABILITE DIRECTE" — config |

**Style commun données :** `PatternFill(solid, D9D9D9)`, `Font(Calibri, size=12, bold=False)`, `Alignment(horizontal="center")`.

**Note :** Le script actuel ne remplit pas ces champs ; il ne reçoit que `site` et `script_folder`. Il faudra soit étendre la signature `fill_onglet_accueil(ws, site, script_folder, **metadata)` soit un dictionnaire `accueil_metadata` pour alimenter ces cellules.

---

## 2. Analyse de la structure visuelle — fusions de cellules

Fusions **uniques** à reproduire (d’après le JSON) :

| Merge_area | Contenu type | Lignes concernées |
|------------|--------------|--------------------|
| **B38:J38** | Titre principal inspection | 38 |
| **B41:J41** | "Description du site en quelques chiffres" | 41 |
| **B44:C44** | Label "Client" | 44 |
| **D44:E44** | Donnée Client | 44 |
| **G44:H44** | Label Type Onduleur | 44 |
| **I44:J44** | Donnée Type Onduleur | 44 |
| **B45:C45** … **I49:J49** | Même schéma (label / donnée) | 45–49 |
| **B56:J57** | Paragraphe réserves par thème | 56–57 |
| **B58:J59** | Paragraphe amélioration continue | 58–59 |
| **B60:J60** | "Les réserves sont réparties selon 3 types de gravité…" | 60 |
| **C65:J66** | Libellé gravité 1 (bloquantes) | 65–66 |
| **C67:J68** | Libellé gravité 2 (majeures) | 67–68 |
| **C69:J70** | Libellé gravité 3 (mineures) | 69–70 |
| **C73:E74** | Texte "Résumé des réserves restantes à lever" | 73–74 |
| **H73:J74** | Texte "Résumé des réserves en cours" | 73–74 |
| **C75:E76** | Texte "Résumé des réserves résolues" | 75–76 |
| **H75:J76** | Texte "Résumé des réserves qui ne seront pas traitées…" | 75–76 |

Cohérence :  
- **Colonnes B–J** utilisées (jusqu’à 10 colonnes). Le script actuel utilise B–K (11 colonnes, `ACCUEIL_PDG_COLS_BK = 11`). Aligner soit le JSON (étendre à K), soit le script (limiter à J) pour éviter les décalages.  
- **Gravité** : dans le JSON les compteurs sont en **colonne B** (B65, B66, B67, B68, B69, B70) ; dans sitemark les formules sont en **colonne A**. À harmoniser si on veut un rendu identique au blueprint.

---

## 3. Audit de la logique de calcul (lignes 65–76)

### 3.1 Bloc Gravité (lignes 65–70)

| Ligne | Colonne B (compteur) | Colonnes C–J (libellé) |
|-------|----------------------|-------------------------|
| 65 | **0** (bloquantes) | "Réserve(s) bloquante(s) affectant la sécurité…" (fusion C65:J66) |
| 66 | **3** | (suite fusion) |
| 67 | **8** (majeures) | "Réserve(s) majeure(s) à lever avant mise en service…" (C67:J68) |
| 68 | **2** | (suite fusion) |
| 69 | **3** (mineures) | "Réserve(s) mineure(s) n'impactant pas…" (C69:J70) |
| 70 | **1** | (suite fusion) |

Lien avec sitemark.py :  
- Les compteurs doivent refléter les réserves par **gravité** (1 = bloquante, 2 = majeure, 3 = mineure).  
- Dans l’onglet **Réserves**, la colonne **C** = Gravité (`META` : `("Gravité", "gravite", 9)`).  
- Les formules actuelles sont correctes :  
  - `=COUNTIF('Réserves'!C:C, "1")` → gravité 1  
  - `=COUNTIF('Réserves'!C:C, "2")` → gravité 2  
  - `=COUNTIF('Réserves'!C:C, "3")` → gravité 3  
- **Différence** : le script met les formules en **colonne A** et le libellé en B:J ; le JSON met les chiffres en **B** et le libellé en C:J. À trancher (nom d’onglet "Accueil" vs "PDG" et position des compteurs).

### 3.2 Bloc Statuts (lignes 73–76)

| Ligne | B | C–E (fusion) | G | H–J (fusion) |
|-------|---|--------------|---|--------------|
| 73 | **3** (formule) | "Résumé des réserves restantes à lever." | **2** (formule) | "Résumé des réserves en cours." |
| 74 | **À faire** (label E7C6B4) | (suite C73:E74) | **En cours** (label E7C6B4) | (suite H73:J74) |
| 75 | **6** (formule) | "Résumé des réserves résolues." | **0** (formule) | "Résumé des réserves qui ne seront pas traitées…" |
| 76 | **Résolu** (label E7C6B4) | (suite C75:E76) | **Ne sera pas fait** (label E7C6B4) | (suite H75:J76) |

Lien avec sitemark.py :  
- Colonne **B** de l’onglet Réserves = Statut.  
- Formules actuelles :  
  - `=COUNTIF('Réserves'!B:B, "À faire")`  
  - `=COUNTIF('Réserves'!B:B, "Résolu")`  
  - `=COUNTIF('Réserves'!B:B, "En cours")`  
  - `=COUNTIF('Réserves'!B:B, "Ne sera pas fait")`  
- Les libellés "À faire", "En cours", "Résolu", "Ne sera pas fait" dans le JSON ont un fond **E7C6B4** (orange/beige). Le script utilise `STATUS_COLORS` pour les cellules de l’onglet Réserves, pas pour ces pastilles du PDG ; pour coller au blueprint, il faudrait un style spécifique **E7C6B4** pour ces 4 cellules (B74, G74, B76, G76).

---

## 4. Plan d’action pour sitemark.py

### 4.1 Fonction `definir_styles_accueil()`

Proposer une structure centralisant les styles déduits du JSON :

```python
def definir_styles_accueil():
    """Retourne un namespace de styles pour l'onglet Accueil (PDG), alignés sur style_expert.json."""
    from openpyxl.styles import Font, PatternFill, Alignment, Border, Side

    thin = Side(style="thin", color="CCCCCC")
    thick = Side(style="thick", color="000000")

    return {
        "fill_titre_principal": PatternFill("solid", start_color="FFFFFF"),
        "fill_label_accueil": PatternFill("solid", start_color="F2E1D9"),   # étiquettes B44:C44, G44:H44, etc.
        "fill_donnee_accueil": PatternFill("solid", start_color="D9D9D9"),   # D44:E44, I44:J44, etc.
        "fill_cyan": PatternFill("solid", start_color=ACCUEIL_CYAN),
        "fill_white": PatternFill("solid", start_color="FFFFFF"),
        "fill_statut_badge": PatternFill("solid", start_color="E7C6B4"),   # B74, G74, B76, G76
        "font_titre_principal": Font(name="Calibri", bold=True, size=20, color="000000"),
        "font_sous_titre": Font(name="Calibri", bold=True, size=12),
        "font_label": Font(name="Calibri", bold=False, size=12),
        "font_donnee": Font(name="Calibri", bold=False, size=12),
        "align_center": Alignment(horizontal="center", vertical="center"),
        "align_left_wrap": Alignment(horizontal="left", vertical="center", wrap_text=True, indent=2),
        "border_bottom_thin": Border(bottom=thin),
        "thin": thin,
        "thick": thick,
    }
```

Dans `fill_onglet_accueil`, on remplacerait les créations ad hoc de `PatternFill` / `Font` par des références à ce dictionnaire (ex. `styles["fill_label_accueil"]`).

### 4.2 Nettoyage des chaînes (encodage) à l’écriture

Le JSON contient des caractères mal encodés (ex. "r�ception", "D�centralis�", "Num�ro"). Pour éviter d’écrire des chaînes corrompues ou de casser l’Excel :

- **Normaliser avant écriture** :  
  - Remplacer les séquences CP1252/ISO-8859-1 mal interprétées en UTF-8 (ex. `�` → caractère correct).  
  - Utiliser une seule convention (UTF-8) côté script et s’assurer que les métadonnées (client, type onduleur, etc.) sont bien en Unicode.

Proposition de helper :

```python
def _normaliser_texte_accueil(s):
    """Retourne une chaîne adaptée à l'écriture Excel (UTF-8, sans caractères de contrôle)."""
    if s is None:
        return ""
    s = str(s).strip()
    # Option : correction des encodages courants (ex. � -> é)
    # s = s.encode("cp1252", errors="replace").decode("utf-8", errors="replace")
    # Suppression des caractères de contrôle
    s = "".join(c for c in s if ord(c) >= 32 or c in "\n\t")
    return s[:32767]  # limite Excel
```

À appeler systématiquement sur toute valeur écrite dans les cellules "données" de l’Accueil (D44, D45, I44, etc.) et, si besoin, sur les libellés lus depuis un fichier ou une base.

---

## 5. Mapping définitif Adresse cellule ↔ Variable Python

| Adresse | Rôle | Variable / source Python |
|---------|------|---------------------------|
| B38 (B38:J38) | Titre principal | Texte fixe + `site` : `f"Inspection du site de {site} avant mise en service et réception en O&M"` |
| B41 (B41:J41) | Sous-titre | Constante "Description du site en quelques chiffres" |
| B44:C44, G44:H44, … | Labels | Constantes (voir §1.1) |
| **D44** (D44:E44) | Client | `accueil_metadata.get("client", "")` |
| **D45** (D45:E45) | Numéro d'affaires | `accueil_metadata.get("numero_affaires", "")` |
| **D46** (D46:E46) | Puissance totale | `accueil_metadata.get("puissance_totale_mwc", "")` |
| **D47** (D47:E47) | Date | `accueil_metadata.get("date_inspection", now)` ou `datetime.today().strftime("%d/%m/%Y")` |
| **D48** (D48:E48) | Nombre PDL/PTR | `accueil_metadata.get("nombre_pdl_ptr", "")` |
| **D49** (D49:E49) | Nombre PTR | `accueil_metadata.get("nombre_ptr", "")` |
| **I44** (I44:J44) | Type onduleur | `accueil_metadata.get("type_onduleur", "")` |
| **I45** (I45:J45) | Nombre onduleurs | `accueil_metadata.get("nombre_onduleurs", "")` |
| **I46** (I46:J46) | Type câblage | `accueil_metadata.get("type_cablage", "")` |
| **I47** (I47:J47) | Type raccordement | `accueil_metadata.get("type_raccordement", "")` |
| **I48** (I48:J48) | Type cheminement | `accueil_metadata.get("type_cheminement", "")` |
| **I49** (I49:J49) | Type tranchée | `accueil_metadata.get("type_tranchee", "")` |
| B65, B67, B69 | Compteurs gravité 1,2,3 | Formules `=COUNTIF('Réserves'!C:C, "1")` etc. (déjà en place ; position A vs B à choisir) |
| B73, G73 | Compteurs "À faire", "En cours" | `=COUNTIF('Réserves'!B:B, "À faire")` etc. |
| B75, G75 | Compteurs "Résolu", "Ne sera pas fait" | Idem |
| B74, G74, B76, G76 | Labels statuts | Constantes "À faire", "En cours", "Résolu", "Ne sera pas fait" avec fill E7C6B4 |

---

## 6. Défis techniques identifiés

| Défi | Description | Piste de résolution |
|------|-------------|----------------------|
| **Décalage lignes** | Le script utilise des lignes calculées (row_desc_title=26, row_intro, row_grav_title, row_stat_title) ; le JSON fixe 38, 41, 44–49, 56–60, 65–76. | Soit adopter les numéros de lignes du JSON (38+) dans le script, soit garder la logique actuelle et documenter la différence (template "Accueil" vs "PDG"). |
| **Colonnes B vs A** | Compteurs gravité : JSON en B65/B67/B69, script en colonne A. | Choisir un référentiel (JSON = référence visuelle) et déplacer les formules en B si besoin. |
| **B–J vs B–K** | JSON s’arrête en J, script utilise K. | Uniformiser (par ex. tout en B–J pour le bloc PDG) pour éviter bordures ou fusions incohérentes. |
| **Fusions** | Beaucoup de merge_area ; l’ordre d’écriture doit être : merge puis valeur sur la cellule d’ancrage (ex. B38 pour B38:J38). | Toujours écrire la valeur sur la première cellule de la plage fusionnée après `merge_cells()`. |
| **Bordures** | JSON signale `border_bottom` sur plusieurs cellules ; le script utilise un `Border` générique (thin). | Réutiliser les styles de `definir_styles_accueil()` et appliquer `border_bottom` sur les lignes de titres/séparatrices. |
| **Encodage** | Caractères � dans le JSON (export Excel/CP1252). | Utiliser `_normaliser_texte_accueil()` pour toutes les chaînes injectées et s’assurer que les entrées (config, métadonnées) sont en UTF-8. |
| **Données Description** | Aucune donnée métier (client, numéro affaires, etc.) n’est aujourd’hui passée à `fill_onglet_accueil`. | Introduire un dictionnaire `accueil_metadata` (ou arguments nommés) et documenter les clés attendues ; en absence de valeur, laisser vide ou placeholder. |

---

## 7. Synthèse

- **Zones fixes** : étiquettes en F2E1D9 (B44:C44, G44:H44, etc.), titres B38:J38 et B41:J41 en blanc, textes d’intro et gravité/statuts en blanc avec fusions listées ci-dessus.  
- **Zones dynamiques** : 12 cellules d’ancrage (D44, D45, D46, D47, D48, D49, I44–I49) à remplir via `accueil_metadata` + normalisation de chaînes.  
- **Calcul** : formules COUNTIF sur 'Réserves'!B:B (Statut) et 'Réserves'!C:C (Gravité) déjà correctes ; à aligner position (A vs B) et plages de fusions avec le blueprint.  
- **Prochaines étapes** : implémenter `definir_styles_accueil()`, `_normaliser_texte_accueil()`, étendre `fill_onglet_accueil` avec `accueil_metadata` et appliquer le mapping §5 ; puis ajuster numéros de lignes et colonnes (B–J, compteurs en B) si souhait de conformité stricte au JSON.
