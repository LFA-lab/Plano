# Audit de conformité visuelle — Onglet Accueil vs Blueprint (PDG)

**Objectif :** Comparer le fichier Excel d’origine fait à la main (`blueprint_complet.json`, onglet "PDG") avec l’implémentation actuelle de l’onglet "Accueil" dans `sitemark.py`, afin de lister les écarts et permettre le même rendu pour les utilisateurs finaux.

---

## 1. Correspondance des zones (Blueprint → Script)

### 1.1 Vue d’ensemble des numéros de lignes

| Zone | Blueprint (PDG) | Script (Accueil) | Écart |
|------|------------------|------------------|--------|
| Titre principal "Inspection du site..." | **B38:J38** (ligne 38) | **G24** (dans le bloc cyan, ligne 24) | Titre blueprint = ligne blanche dédiée ; script = texte dans le bloc cyan |
| Sous-titre "Description du site en quelques chiffres" | **B41:J41** (41) | **B26:K26** (26) | Ligne différente (26 vs 41) |
| Grille Description (labels + données) | **44–49** (6 lignes) | **27–32** (6 lignes) | Lignes 27–32 vs 44–49 |
| Paragraphes d’intro (thèmes, amélioration continue, gravité) | **B56:J57**, **B58:J59**, **B60:J60** | **B34:K38** (une seule fusion 5 lignes) | Même contenu, fusions différentes (2+2+1 vs 5) |
| Titre "Gravité" | — (pas de ligne dédiée dans le JSON) | **41** | — |
| Bloc Gravité (3 lignes × 2) | **65–70** (B = compteur, C:J = libellé) | **42–44** (A = compteur, B:K = libellé) | Colonne compteur B vs A ; fusion C:J vs B:K |
| Titre "Statuts" | — | **46** | — |
| Bloc Statuts (2×2 lignes) | **73–76** (compte + texte L1, badge L2) | **47–48** (compte + texte sur une seule ligne) | Blueprint : 2 lignes par statut (compte+texte puis badge) ; script : 1 ligne |

Les numéros de lignes du script sont **décalés plus haut** (pas de lignes vides 38–43, 50–55, 61–64, 71–72). Pour une conformité stricte on peut soit **reproduire les numéros de lignes du blueprint** (38, 41, 44…), soit garder le script actuel et **aligner uniquement couleurs, polices et structure des cellules**.

---

## 2. Écarts visuels à corriger pour conformité

### 2.1 Bloc "Description du site en quelques chiffres"

| Élément | Blueprint (PDG) | Script actuel | Action recommandée |
|--------|------------------|---------------|---------------------|
| **Couleur des étiquettes** | **F2E1D9** (beige) | **D9D9D9** (gris) | Utiliser `PatternFill("solid", "F2E1D9")` pour tous les labels (Client, Numéro d'affaires, etc.) |
| **Couleur des cellules données** | **D9D9D9** (gris) | **FFFFFF** (blanc) | Utiliser `PatternFill("solid", "D9D9D9")` pour les cellules de valeur (D44, I44, etc.) |
| **Position des labels** | Gauche : **B44:C44** (fusion 2 col.) ; Droite : **G44:H44** | Gauche : **colonne A** ; Droite : **colonne H** | Passer à une structure **B:C** = label gauche, **D:E** = donnée gauche ; **G:H** = label droite, **I:J** = donnée droite (sans utiliser A pour les labels) |
| **Position des données** | Gauche : **D44:E44** ; Droite : **I44:J44** | Gauche : **B:G** fusionné ; Droite : **I:K** fusionné | Adopter les fusions **D:E** et **I:J** comme dans le blueprint (pas B:G ni I:K) |
| **Colonnes utilisées** | **B à J** (col. F libre entre les deux blocs) | **A à K** (A = label gauche, B–K = reste) | Limiter le bloc à **B–J** et ne pas utiliser A pour ce bloc (ou réserver A pour autre usage) |
| **Libellés exacts** | "Client", "Numéro d'Affaires :", "Puissance totale de la centrale", "Date :", "Nombre de PDL / PTR", "Nombre de PTR" ; "Type Onduleur : Cent/Décentralisé", "Nombre d' Onduleur", "Type Câblage solaire", "Type de raccordement solaire", "Type de Cheminement : Tranchée/CDC", "Type Tranchée : Gaines TPC / Enterrabilité directe " | "Client", "Numero d'Affaires", "Puissance totale", "Date", "Nombre de PDL/PTR" ; "Type Onduleur", "Nombre d'Onduleur", "Type Cablage", etc. | Remplacer par les libellés du blueprint (accents, deux-points, texte complet) |

Résumé : **F2E1D9** pour les labels, **D9D9D9** pour les données ; structure **B:C / D:E** à gauche et **G:H / I:J** à droite ; libellés identiques au blueprint.

---

### 2.2 Titre principal "Inspection du site..."

| Élément | Blueprint | Script | Action |
|--------|-----------|--------|--------|
| Fond | **FFFFFF** (blanc) | **Cyan** (dans le bloc E24/G24) | Dans le blueprint, ce titre est sur fond blanc (ligne 38). Le script le met dans le bloc cyan (ligne 24). Soit on garde le script (tout en cyan), soit on ajoute une ligne dédiée blanche avec ce texte pour coller au manuel. |
| Police | bold, size **20** | bold, size **14**, blanc | Si on aligne sur le blueprint : size 20, couleur noire (fond blanc). |
| Fusion | **B38:J38** | Bloc G24:O24 (texte dans le cyan) | — |

---

### 2.3 Sous-titre "Description du site en quelques chiffres"

| Élément | Blueprint | Script | Action |
|--------|-----------|--------|--------|
| Fusion | **B41:J41** | B26:K26 | Aligner sur **B:J** (pas K) si on veut stricte conformité. |
| Taille | size **20** (ou 12 selon cellules) | size **12** | Garder 12 pour cohérence avec le reste du bloc. |

---

### 2.4 Bloc Gravité (réserves 1, 2, 3)

| Élément | Blueprint | Script | Action |
|--------|-----------|--------|--------|
| **Colonne du compteur** | **B** (B65, B66, B67, B68, B69, B70) | **A** | Mettre les formules COUNTIF en **colonne B** pour conformité. |
| **Fusion du libellé** | **C65:J66**, **C67:J68**, **C69:J70** | **B:K** (une ligne par gravité) | En blueprint chaque gravité = 2 lignes (L1 = compteur B + début texte, L2 = suite texte). Script = 1 ligne. Adopter **C:J** pour le texte (pas B:K). |
| **Fond des cellules compteur** | **FFFFFF** (blanc) | **GRAVITY_COLORS** (rouge / orange / jaune) | Blueprint : fond blanc pour B65, B66… Pour conformité stricte : fond blanc en B. Sinon garder les couleurs en A (choix métier). |
| **Taille police compteur** | size **20** | size **12** | Optionnel : passer à 20 pour le chiffre si on vise le même rendu. |

---

### 2.5 Bloc Statuts (À faire, En cours, Résolu, Ne sera pas fait)

| Élément | Blueprint | Script | Action |
|--------|-----------|--------|--------|
| **Structure** | **2 lignes par paire** : L1 = compteur (B73, G73) + texte (C73:E74, H73:J74) ; L2 = **badge** "À faire" en **B74**, "En cours" en **G74** (bg **E7C6B4**), idem B75/G75, B76/G76 | **1 ligne** : compteur + texte sur la même ligne, pas de ligne "badge" | Ajouter une **2ᵉ ligne par paire** avec uniquement les libellés "À faire", "En cours", "Résolu", "Ne sera pas fait" dans des cellules **E7C6B4**. |
| **Couleur des badges** | **E7C6B4** (beige/orange) pour B74, G74, B76, G76 | Non utilisé (script utilise STATUS_COLORS pour l’onglet Réserves, pas pour l’Accueil) | Définir `fill_statut_badge = PatternFill("solid", "E7C6B4")` et l’appliquer aux 4 cellules de libellé de statut. |
| **Taille police badge** | size **11** | — | Utiliser size **11** pour les cellules "À faire", "En cours", "Résolu", "Ne sera pas fait". |
| **Fusions texte** | C73:E74 (À faire), H73:J74 (En cours), C75:E76, H75:J76 | C:G et I:K | Aligner sur **C:E** et **H:J** pour le texte descriptif. |

---

### 2.6 Paragraphes d’introduction (thèmes, amélioration continue, gravité)

| Élément | Blueprint | Script | Action |
|--------|-----------|--------|--------|
| Fusions | **B56:J57** (1er paragraphe), **B58:J59** (2ᵉ), **B60:J60** (3ᵉ) | **B34:K38** (une seule fusion 5 lignes) | Contenu identique ; pour le même visuel on peut scinder en 3 fusions B:J. |
| Taille | size **20** | size **10** | Optionnel : augmenter à 11–12 pour lisibilité ; 20 dans le blueprint peut être un défaut d’export. |

---

### 2.7 Images et ancres (référence blueprint)

Le blueprint indique des zones d’image / forme :

- **B1** : "Text Placeholder 3" (titre)
- **B10** : "Espace réservé pour une image 5" (vue aérienne)
- **F10** : "Espace réservé pour une image 31"
- **E24** : formes (flèche, freeform) — bloc cyan
- **G25** : "Espace réservé du texte 15"

Le script place : titre en B2, vue aérienne en B10, logo en G11, bloc cyan E24 et texte en G24. Les ancres sont cohérentes (B10, E24, zone G24/G25) ; seule la ligne de titre (B1 vs B2) et le logo (F10 vs G11) diffèrent légèrement.

---

## 3. Synthèse des modifications à prévoir dans `sitemark.py`

Pour obtenir le **même visuel** que l’ancien Excel :

1. **Couleurs**
   - Labels du bloc Description : **F2E1D9** (au lieu de D9D9D9).
   - Cellules de données du bloc Description : **D9D9D9** (au lieu de blanc).
   - Badges Statuts (À faire, En cours, Résolu, Ne sera pas fait) : **E7C6B4**.

2. **Structure du bloc Description**
   - Ne plus utiliser la colonne **A** pour les labels.
   - Gauche : **B:C** = label (fusion), **D:E** = donnée (fusion).
   - Droite : **G:H** = label (fusion), **I:J** = donnée (fusion).
   - Colonne **F** : vide ou séparateur.
   - Utiliser **B–J** (pas K) pour ce bloc si on veut coller au blueprint.

3. **Libellés**
   - Remplacer par les textes exacts du blueprint (avec accents, deux-points, libellés longs : "Puissance totale de la centrale", "Type Onduleur : Cent/Décentralisé", "Type Tranchée : Gaines TPC / Enterrabilité directe ", etc.).

4. **Gravité**
   - Option conforme blueprint : formules en **colonne B**, libellés en **C:J**, fond **blanc** pour les cellules de compteur.
   - Ou garder les compteurs en **A** avec couleurs (rouge/orange/jaune) pour la lisibilité actuelle.

5. **Statuts**
   - Passer à **2 lignes par paire** : ligne 1 = compteur (B et G) + texte (C:E et H:J) ; ligne 2 = uniquement les cellules "À faire", "En cours", "Résolu", "Ne sera pas fait" avec fond **E7C6B4** et police **11**.
   - Fusions texte : **C73:E74**, **H73:J74**, **C75:E76**, **H75:J76** (ou équivalent en numéros de lignes du script).

6. **Titre "Inspection du site..."**
   - Soit le laisser dans le bloc cyan (comportement actuel), soit ajouter une ligne à part avec fond blanc et police 20 comme en B38 du blueprint.

7. **Fusions**
   - Partout où le script utilise **B:K**, envisager **B:J** pour les zones qui s’arrêtent en J dans le blueprint (titres, intro, bloc Description).

8. **Numéros de lignes**
   - Soit on conserve les lignes calculées du script (26, 27–32, 34–38, 41–44, 46–48) : seul le rendu (couleurs, structure, textes) change.
   - Soit on bascule sur les lignes fixes du blueprint (38, 41, 44–49, 56–60, 65–70, 73–76) pour que la feuille soit identique ligne à ligne à l’ancien fichier.

---

## 4. Tableau de correspondance rapide (cellule / style)

| Zone | Blueprint | Couleur / style à reproduire |
|------|-----------|------------------------------|
| Titre "Inspection du site..." | B38:J38 | bg FFFFFF, bold, size 20, center |
| Sous-titre Description | B41:J41 | bg FFFFFF, bold, size 12, center |
| Labels Description (Client, Type Onduleur…) | B44:C44, G44:H44, … | bg **F2E1D9**, size 12, center |
| Données Description (VENSOLAIR, 10, …) | D44:E44, I44:J44, … | bg **D9D9D9**, size 12, center |
| Intro (thèmes, amélioration, gravité) | B56:J60 | bg FFFFFF, size 10–12, left, wrap |
| Compteurs Gravité | B65, B66, B67, B68, B69, B70 | bg FFFFFF, size 20 (ou 12), center ; formule COUNTIF |
| Libellés Gravité | C65:J66, C67:J68, C69:J70 | bg FFFFFF, size 10, left, wrap |
| Compteurs Statuts | B73, G73, B75, G75 | bg FFFFFF, formule COUNTIF |
| Badges Statuts | B74, G74, B76, G76 | bg **E7C6B4**, size **11**, center ("À faire", "En cours", "Résolu", "Ne sera pas fait") |
| Textes Statuts | C73:E74, H73:J74, C75:E76, H75:J76 | bg FFFFFF, left, wrap |

---

## 5. Fichiers concernés

- **Référence visuelle :** `blueprint_complet.json` (onglet `sheet_name: "PDG"`).
- **Code à adapter :** `sitemark.py`, fonction `fill_onglet_accueil` (environ lignes 264–493).
- **Constantes utiles à ajouter :**  
  `ACCUEIL_FILL_LABEL = "F2E1D9"`, `ACCUEIL_FILL_DONNEE = "D9D9D9"`, `ACCUEIL_FILL_BADGE_STATUT = "E7C6B4"`.

Une fois ces points appliqués, l’onglet Accueil généré par le script pourra être aligné visuellement sur l’ancien Excel manuel pour la conformité des utilisateurs finaux.
