# Guide de dépannage - Export Excel MS Project

## 📋 Prérequis obligatoires

### 1. Activation des macros
- **Fichier** > **Options** > **Centre de gestion de la confidentialité**
- **Paramètres du centre de gestion de la confidentialité**
- **Paramètres des macros** :
  - ✅ Cocher "Activer toutes les macros"
  - ✅ Cocher "Accès approuvé au modèle d'objet de projet VBA"

### 2. Vérifications Office
- Office installé complètement (pas en mode Click-to-Run bloqué)
- Excel disponible et démarrable manuellement
- MS Project avec accès VBA autorisé

## 🔧 Procédure de diagnostic

### Étape 1 : Lancer le diagnostic
```
Ouvrir MS Project → Macros → Diagnostic_Environnement
```

Le diagnostic vérifie automatiquement :
- ✅ Excel.Application (automation COM)
- ✅ MSXML (décodage Base64 des logos)
- ✅ ADODB.Stream (écriture fichiers temporaires)
- ✅ FileDialog MS Project (sélecteur de dossier)
- ✅ Droits d'écriture (Downloads/Bureau/Documents)
- ✅ Accès VBA autorisé

### Étape 2 : Analyser les résultats

#### ✅ Tout OK
Votre environnement est prêt. Lancez `ExportMecanique`.

#### ❌ Excel.Application - ÉCHEC
**Causes possibles :**
- Excel non installé ou version incomplète
- Version Office Click-to-Run avec restrictions IT
- Excel corrompu

**Solutions :**
1. Réparer Office via Panneau de configuration
2. Redémarrer en tant qu'administrateur
3. Contacter votre service IT

#### ❌ MSXML - ÉCHEC
**Cause :** Composant MSXML absent (logos non insérés)
**Solution :** Windows Update ou installer MSXML manuellement

#### ❌ ADODB.Stream - ÉCHEC
**Cause :** Composant ADO manquant
**Solution :** Installer/réparer MDAC (Microsoft Data Access Components)

#### ❌ Droits d'écriture - ÉCHEC
**Causes possibles :**
- Dossier OneDrive/SharePoint en mode "Fichiers à la demande"
- Droits NTFS insuffisants
- Antivirus bloquant l'écriture
- Politique de sécurité IT

**Solutions :**
1. Choisir un dossier local (ex: C:\Temp)
2. Exécuter MS Project en tant qu'administrateur
3. Désactiver temporairement la synchronisation OneDrive
4. Contacter votre service IT

## 🚨 En cas d'erreur persistante

### Option de secours : Export CSV
Si Excel reste indisponible, utilisez :
```
Macros → ExportCSV_Secours
```
Génère un fichier CSV avec les données de base (ressources, heures prévues, pourcentages).

### Informations à fournir au support
Copiez-collez le rapport complet du diagnostic comprenant :
- ✅/❌ État de chaque composant
- Numéros d'erreur exacts (Err.Number)
- Descriptions d'erreur (Err.Description)
- Chemins testés
- ProgID qui échoue

## 📞 Support technique

Pour toute assistance :
- Envoyer le rapport de diagnostic complet
- Préciser votre version d'Office (32/64 bits)
- Préciser votre environnement (OneDrive, domaine, VPN)

---

## 🛠️ Causes courantes d'erreurs

### "Erreur Automation" générique
- Centre de gestion de la confidentialité mal configuré
- Excel non automatisable (Click-to-Run)
- Processus Excel fantôme en arrière-plan

### "Fichier non trouvé" / "Chemin non valide"
- Dossier OneDrive non synchronisé
- Caractères spéciaux dans le chemin
- Droits insuffisants

### "Composant non disponible"
- Installation Office incomplète
- MSXML manquant (Windows Server minimal)
- ADO non inscrit dans le registre

### Solutions générales
1. **Redémarrer** MS Project et Excel
2. **Réparer Office** via Panneau de configuration
3. **Exécuter en administrateur** temporairement
4. **Mettre à jour Windows** (composants COM)
5. **Contacter IT** si problème persiste
