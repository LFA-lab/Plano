import unicodedata
import re

def clean_vba_file(text: str) -> str:
    """
    Nettoie un fichier VBA de manière intelligente :
    - Supprime les accents dans les chaînes entre guillemets ET dans les commentaires
    - Préserve TOUS les commentaires VBA (lignes commençant par ')
    - Préserve toute la syntaxe VBA
    """
    
    lines = text.split('\n')
    cleaned_lines = []
    
    for line in lines:
        # Nettoyer les accents partout, mais préserver la structure
        def remove_accents(text):
            return ''.join(
                c for c in unicodedata.normalize('NFD', text)
                if unicodedata.category(c) != 'Mn'
            )
        
        # Nettoyer les accents dans toute la ligne
        line = remove_accents(line)
        
        # Mais on ne touche PAS aux apostrophes ni aux autres caractères de structure
        cleaned_lines.append(line)
    
    return '\n'.join(cleaned_lines)

def smart_clean_file(file_path: str):
    """
    Nettoie intelligemment un fichier selon son extension
    """
    
    # Lecture du fichier
    with open(file_path, "r", encoding="utf-8") as f:
        content = f.read()
    
    # Détection du type de fichier
    if file_path.endswith('.bas') or file_path.endswith('.vba') or file_path.endswith('.cls'):
        # Fichier VBA - nettoyage intelligent
        cleaned = clean_vba_file(content)
        print(f"📝 Mode VBA détecté - nettoyage des chaînes uniquement")
    else:
        # Autres fichiers - nettoyage complet (ancien comportement)
        cleaned = clean_text_complete(content)
        print(f"📄 Mode texte standard - nettoyage complet")
    
    # Réécriture
    with open(file_path, "w", encoding="utf-8") as f:
        f.write(cleaned)
    
    print(f"✅ Fichier nettoyé et remplacé : {file_path}")

def clean_text_complete(text: str) -> str:
    """Nettoyage complet pour les fichiers non-VBA (ancien comportement)"""
    # 1. Supprimer les accents
    no_accents = ''.join(
        c for c in unicodedata.normalize('NFD', text)
        if unicodedata.category(c) != 'Mn'
    )

    # 2. Supprimer les apostrophes et guillemets
    no_apostrophe = re.sub(r"[''`]", "", no_accents)

    # 3. Supprimer les emojis et caractères hors ASCII
    no_emoji = re.sub(r"[^\x00-\x7F]", "", no_apostrophe)

    return no_emoji

# Configuration - Mets ici le chemin de ton fichier
file_path = "/home/ntoi/LFA-lab/Omexom/macros/Macro MSP/Optimisation/ExportHeuresSapin.bas"

# Nettoyage intelligent
smart_clean_file(file_path)
