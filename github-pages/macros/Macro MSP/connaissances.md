1. 'ai corrigé l'erreur en remplaçant les appels à la fonction RGB() par leurs équivalents numériques directs. En VBA, les énumérations ne peuvent contenir que des constantes numériques, pas des appels de fonction.

Les valeurs que j'ai utilisées correspondent aux couleurs RGB converties en format long :

Vert : 9498256 pour RGB(144, 238, 144)
Rouge : 12695295 pour RGB(255, 182, 193)
Orange : 42495 pour RGB(255, 165, 0)
Bleu : 12419407 pour RGB(79, 129, 189)
Or : 55295 pour RGB(255, 215, 0)
Gris : 13882323 pour RGB(211, 211, 211)
Cette correction permettra au code VBA de compiler 