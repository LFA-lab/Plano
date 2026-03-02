# Différence entre les deux blueprints JSON

**Fichiers comparés :**
- **Créé (automatique)** : `blueprint_completautomatique.json` — sortie du script sitemark (onglet "Accueil")
- **Ancien (manuel)** : `blueprint_complet.json` — Excel d’origine fait à la main (onglet "PDG")

---

## 1. Métadonnées racine

| Élément | blueprint_completautomatique.json | blueprint_complet.json |
|--------|-----------------------------------|-------------------------|
| **sheet_name** | `"Accueil"` | `"PDG"` |
| **Nombre de cellules** | ~450+ (lignes 2–54, colonnes A–N) | ~200 (lignes 38–76, colonnes B–J uniquement) |
| **Images** | 2 : Picture 1 (B10), Picture 2 (G11) | 6 : Text Placeholder B1, image B10, image F10, formes E24, texte G25 |

---

## 2. Numérotation des lignes (structure)

Le contenu logique est le même, mais **les numéros de lignes ne correspondent pas** :

| Zone | Automatique (créé) | Ancien (manuel) |
|------|--------------------|------------------|
| Titre principal type "Rapport de pré-commissioning" | **B2** (fusion B2:S2) | — (absent à cette position) |
| Texte "Inspection du site... avant mise en service" | **G24** (dans bloc E24:O24) | **B38** (ligne dédiée B38:J38) |
| Sous-titre "Description du site en quelques chiffres" | **B26** (B26:J26) | **B41** (B41:J41) |
| Grille Description (labels + données) | **27–32** (6 lignes) | **44–49** (6 lignes) |
| Paragraphes d’intro (thèmes, amélioration, gravité) | **35–39** (fusion B35:J39) | **56–60** (B56:J57, B58:J59, B60:J60) |
| Titre "Gravité" | **42** (B42:J42) | — (pas de ligne titre explicite dans le JSON) |
| Bloc Gravité (3 niveaux) | **43–45** (1 ligne par niveau) | **65–70** (2 lignes par niveau : compteur L1, suite L2) |
| Titre "Statuts" | **48** (B48:J48) | — |
| Bloc Statuts (4 lignes) | **49–52** | **73–76** |

Résumé : l’automatique **remonte tout** (pas de lignes vides 1–37, 50–55, 61–64, 71–72) et utilise des lignes **26–52** au lieu de **38–76**.

---

## 3. Couleurs (bg) — écarts visuels

| Zone | blueprint_completautomatique.json (créé) | blueprint_complet.json (ancien) |
|------|------------------------------------------|----------------------------------|
| Bande titre du haut | **D1CE00** (jaune/vert) en B2 | — |
| Bloc "Inspection du site..." (ligne 24) | **D1CE00** (E24, G24) | — (à la place : B38 **FFFFFF**) |
| Labels du bloc Description (Client, Type Onduleur…) | **D9E1F2** (bleu clair) | **F2E1D9** (beige) |
| Cellules données du bloc Description | **D9D9D9** (gris) | **D9D9D9** (gris) ✓ |
| Badges Statuts ("À faire", "En cours", "Résolu", "Ne sera pas fait") | **B4C6E7** (bleu) | **E7C6B4** (beige/orange) |
| Bloc "Ne sera pas fait" (texte à droite, 2ᵉ paire) | **4F4737** (gris foncé, texte clair) | **FFFFFF** (blanc) |

À aligner pour conformité : **F2E1D9** pour les labels Description, **E7C6B4** pour les badges Statuts, et selon choix : titre B38 en blanc ou bande D1CE00 en B2.

---

## 4. Fusions et colonnes

| Élément | Automatique (créé) | Ancien (manuel) |
|--------|--------------------|------------------|
| Étendue des colonnes utilisées | **A à N** (14 colonnes, lignes 25+) | **B à J** (9 colonnes) |
| Titre du haut | B2:**S2** (18 colonnes) | — |
| Bloc "Inspection du site..." | G24:**O24** (jusqu’à O) | B38:**J38** |
| Sous-titre Description | B26:**J26** ✓ | B41:**J41** ✓ |
| Intro | B35:**J39** (une fusion 5 lignes) | B56:J57, B58:J59, B60:J60 (3 fusions) |
| Gravité — libellé | **C43:J43**, C44:J44, C45:J45 (1 ligne par niveau) | **C65:J66**, C67:J68, C69:J70 (2 lignes par niveau) |
| Statuts — texte gauche/droite | C49:**E50**, H49:**J50**, C51:**E52**, H51:**J52** ✓ | C73:E74, H73:J74, C75:E76, H75:J76 ✓ |

Structure des fusions Statuts et Description (B:C, D:E, G:H, I:J) est la même ; l’ancien utilise uniquement B–J, l’automatique étend jusqu’à N sur certaines lignes.

---

## 5. Polices et alignements

| Zone | Automatique | Ancien |
|------|-------------|--------|
| Taille par défaut | size **11** (nombreuses cellules) | size **12** ou **20** selon zone |
| Titre "Inspection du site..." | bold, size **14** (en G24) | bold, size **20** (en B38) |
| Sous-titre Description | bold, size **12** | bold, size **20** ou 12 |
| Libellés Description | size **12**, center | size **12**, center ✓ |
| Compteurs Gravité | size **12**, center | size **20**, center |
| Badges Statuts | size **11**, center ✓ | size **11**, center ✓ |
| Alignement défaut | **left** sur beaucoup de cellules | center sur labels/données, left sur textes longs |

---

## 6. Contenu texte (valeurs)

- **Automatique** : pas de données métier dans les cellules Description (valeurs vides en D27:E32, I27:J32) ; nom de site long dans B2 et G24 (ex. `VENSOLAIR_SBRU_MWc_P.0828413.T.01_2026-02-26`).
- **Ancien** : exemples de données (VENSOLAIR, P.0828413.T.01, 4,1 MWC, 09/12/2025, 10, DÉCENTRALISÉ, LIGNE, etc.) en D44:J49.
- Libellés (Client, Numéro d’affaires, Type Onduleur…) : identiques ou très proches (accents / césure selon encodage).

---

## 7. Images

| Automatique | Ancien |
|-------------|--------|
| Picture 1 → **B10** | Espace réservé image 5 → **B10** |
| Picture 2 → **G11** | Espace réservé image 31 → **F10** |
| — | Formes (Arrow, Freeform) → **E24** |
| — | Espace réservé texte 15 → **G25** |
| — | Text Placeholder 3 → **B1** |

L’ancien décrit plus de zones (titres, formes, emplacements texte) ; l’automatique ne liste que les 2 images insérées par le script.

---

## 8. Synthèse des écarts à corriger pour coller à l’ancien

1. **Nom d’onglet** : "Accueil" vs "PDG" (choix métier).
2. **Couleurs** : remplacer **D9E1F2** par **F2E1D9** (labels Description), **B4C6E7** par **E7C6B4** (badges Statuts). Option : bande titre **D1CE00** vs titre B38 **FFFFFF**.
3. **Lignes** : soit garder les lignes 26–52 (automatique), soit insérer des lignes vides pour retrouver 38, 41, 44–49, 56–60, 65–70, 73–76.
4. **Gravité** : ancien = 2 lignes par niveau (fusion C:J sur 2 lignes) ; automatique = 1 ligne par niveau. Option : passer à 2 lignes par gravité pour conformité.
5. **Colonnes** : limiter le bloc PDG à **B:J** (pas de contenu en A, K–N) pour être aligné avec l’ancien.
6. **Taille police** : titre "Inspection du site..." en 20 (ancien) vs 14 (automatique) ; compteurs gravité en 20 (ancien) vs 12 (automatique).
7. **Bloc "Ne sera pas fait"** : ancien en blanc ; automatique en **4F4737**. Choisir selon la charte (conserver ou repasser en blanc).

Ces points peuvent être traités dans `sitemark.py` (couleurs, fusions, tailles, colonnes) pour que le blueprint exporté se rapproche au maximum de `blueprint_complet.json`.
