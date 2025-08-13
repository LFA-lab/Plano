# Instructions d'installation - Macro Export Mécanique

## Prérequis système

### 1. Logiciels requis
- Microsoft Project (version 2010 ou ultérieure)
- Microsoft Excel (version 2010 ou ultérieure)
- Windows 7 ou ultérieur

### 2. Paramètres de sécurité
**Important** : Pour éviter les erreurs d'automation, configurez les paramètres suivants :

#### Dans Microsoft Project :
1. Fichier → Options → Centre de gestion de la confidentialité → Paramètres du centre de gestion
2. Paramètres des macros → Activer toutes les macros
3. Objets ActiveX → Activer tous les contrôles

#### Dans Microsoft Excel (si des erreurs persistent) :
1. Fichier → Options → Approbations → Paramètres des macros
2. Activer toutes les macros

## Installation de la macro

### Méthode 1 : Import direct
1. Ouvrir Microsoft Project
2. Alt + F11 pour ouvrir l'éditeur VBA
3. Fichier → Importer le fichier → Sélectionner `exportmécanique.bas`
4. Fermer l'éditeur VBA

### Méthode 2 : Copier-coller
1. Ouvrir le fichier `exportmécanique.bas` dans un éditeur de texte
2. Copier tout le contenu
3. Dans MS Project : Alt + F11
4. Insertion → Module
5. Coller le code
6. Ctrl + S pour sauvegarder

## Utilisation

### Exécution de la macro
1. Dans MS Project, ouvrir votre projet
2. Alt + F8 → Sélectionner `ExportMecaniqueComplet`
3. Cliquer sur "Exécuter"

### Ou créer un bouton (optionnel)
1. Exécuter la macro `InstallerBoutonExportMeca`
2. Un bouton "Export Mécanique" apparaîtra dans les barres d'outils

## Résolution des problèmes courants

### Erreur "Automation error" ou "Object doesn't support this property"
**Solution :**
1. Fermer complètement Excel et MS Project
2. Redémarrer MS Project
3. Réessayer la macro

### Erreur "Permission denied" ou "File access denied"
**Solution :**
1. Vérifier que le dossier Téléchargements est accessible
2. Exécuter MS Project en tant qu'administrateur
3. La macro utilisera automatiquement le Bureau ou Documents si Téléchargements n'est pas accessible

### Erreur "Dictionary object not found"
**Solution :**
1. Dans l'éditeur VBA (Alt + F11)
2. Outils → Références
3. Cocher "Microsoft Scripting Runtime"
4. OK et réessayer

### Erreur "Excel object not found"
**Solution :**
1. Vérifier qu'Excel est installé
2. Dans VBA : Outils → Références → Cocher "Microsoft Excel Object Library"
3. Redémarrer MS Project

## Support technique

Si les problèmes persistent :
1. Noter le message d'erreur exact
2. Vérifier la version de MS Project/Excel
3. Tester sur une machine similaire
4. Contacter le support IT de votre organisation

## Notes importantes

- La macro exporte uniquement les ressources du groupe "Mécanique"
- Le fichier Excel est sauvé avec la date/heure dans le nom
- Les données incluent le travail prévu, réalisé et les pourcentages d'avancement
- Compatible avec les versions 32 et 64 bits d'Office
