Sub Import_Taches_Indentation_Auto()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim i As Long, lastRow As Long
    Dim t As Task, a As Assignment
    Dim fichierExcel As String
    Dim estRecapTache As Boolean

    ' ==== SÉLECTION DU FICHIER ====
    Dim xlTempApp As Object
    Set xlTempApp = CreateObject("Excel.Application")

    With xlTempApp.FileDialog(msoFileDialogFilePicker)
        .Title = "Sélectionnez le fichier Excel à importer"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx; *.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            fichierExcel = .SelectedItems(1)
        Else
            MsgBox "Aucun fichier sélectionné.", vbExclamation
            xlTempApp.Quit
            Set xlTempApp = Nothing
            Exit Sub
        End If
    End With

    xlTempApp.Quit
    Set xlTempApp = Nothing

    ' ==== OUVERTURE EXCEL ====
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    Set xlBook = xlApp.Workbooks.Open(fichierExcel)
    Set xlSheet = xlBook.Sheets(1)

    ' ==== OUVERTURE MS PROJECT ====
    Set pjApp = MSProject.Application
    pjApp.Visible = True
    pjApp.FileNew
    Set pjProj = pjApp.ActiveProject

    ' ==== TITRE PROJET (A2) ====
    Dim tRoot As Task
    Set tRoot = pjProj.Tasks.Add(Name:=xlSheet.Cells(2, 1).Value, Before:=1)
    tRoot.Manual = False

    ' ==== CONFIGURATION PROJET ====
    pjProj.DefaultTaskType = pjFixedWork
    pjProj.ScheduleFromStart = True
    pjProj.DefaultEffortDriven = True

    ' ==== CALENDRIER ====
    With ActiveProject.BaseCalendars("Standard").WorkWeeks
        .Add Start:="01/01/2025", Finish:="01/01/2027", Name:="Calendrier Standard"
        With .Item(1)
            Dim j As Integer
            For j = 2 To 6
                With .WeekDays(j)
                    .Shift1.Start = "09:00"
                    .Shift1.Finish = "18:00"
                    .Shift2.Clear: .Shift3.Clear: .Shift4.Clear: .Shift5.Clear
                End With
            Next j
            .WeekDays(1).Default
            .WeekDays(7).Default
        End With
    End With

    ' ==== RESSOURCE MONTEURS ====
    Dim rMonteurs As Resource
    Set rMonteurs = GetOrCreateWorkResource("Monteurs")

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row

    ' -----------------------------------------------------------
    ' ?? ANALYSE PRÉALABLE : Calculer les niveaux
    ' -----------------------------------------------------------
    
    Debug.Print "========================================"
    Debug.Print "=== ANALYSE DE LA STRUCTURE ==="
    Debug.Print "========================================"
    
    ' Tableau pour stocker les niveaux calculés
    Dim niveauxCalcules() As Integer
    ReDim niveauxCalcules(3 To lastRow)
    
    Dim niveauActuel As Integer
    niveauActuel = 2 ' Commence au niveau 2 (sous le titre niveau 1)
    
    For i = 3 To lastRow
        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        
        nom = xlSheet.Cells(i, 1).Value
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value

        If nom <> "" Then
            ' Détection récap : B, C, D vides
            Dim qteVide As Boolean, persVide As Boolean, hVide As Boolean
            qteVide = (qte = "" Or IsEmpty(qte))
            persVide = (pers = "" Or IsEmpty(pers))
            hVide = (h = "" Or IsEmpty(h))
            
            Dim estRecap As Boolean
            estRecap = qteVide And persVide And hVide
            
            If estRecap Then
                ' RÉCAPITULATIVE : toujours niveau 2
                niveauxCalcules(i) = 2
                niveauActuel = 3 ' Les suivantes seront niveau 3
                Debug.Print "Ligne " & i & " [RECAP] " & nom & " ? Niveau 2"
            Else
                ' SUBORDONNÉE : niveau actuel (3)
                niveauxCalcules(i) = niveauActuel
                Debug.Print "Ligne " & i & " [SUB]   " & nom & " ? Niveau " & niveauActuel
            End If
        End If
    Next i

    ' -----------------------------------------------------------
    ' ??? CRÉATION DES TÂCHES AVEC NIVEAUX CALCULÉS
    ' -----------------------------------------------------------
    
    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "=== IMPORT DES TÂCHES ==="
    Debug.Print "========================================"
    
    ' Table (tableaux parallèles) pour stocker les tâches et leurs infos
    Dim tabTaches() As Object
    Dim tabNiveaux() As Integer
    Dim tabNoms() As String
    Dim tabQte() As Variant
    Dim tabPers() As Variant
    Dim tabH() As Variant
    Dim nbTaches As Long
    nbTaches = 0
    
    ' Étape 1 : Compter les tâches valides
    Dim nbTachesValides As Long
    nbTachesValides = 0
    For i = 3 To lastRow
        If xlSheet.Cells(i, 1).Value <> "" Then
            nbTachesValides = nbTachesValides + 1
        End If
    Next i
    
    ' Redimensionner les tableaux
    ReDim tabTaches(1 To nbTachesValides)
    ReDim tabNiveaux(1 To nbTachesValides)
    ReDim tabNoms(1 To nbTachesValides)
    ReDim tabQte(1 To nbTachesValides)
    ReDim tabPers(1 To nbTachesValides)
    ReDim tabH(1 To nbTachesValides)
    
    ' Étape 2 : Créer toutes les tâches et les stocker dans la table
    nbTaches = 0
    For i = 3 To lastRow
        nom = xlSheet.Cells(i, 1).Value
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value

        If nom <> "" Then
            ' Créer la tâche
            Set t = pjProj.Tasks.Add(nom)
            t.Manual = False
            t.Type = pjFixedWork
            t.ConstraintType = pjASAP
            t.Calendar = ActiveProject.BaseCalendars("Standard")
            
            ' Forcer Summary = False pour les tâches non-récapitulatives
            ' (MS Project peut créer automatiquement des summary tasks)
            estRecapTache = (qte = "" Or IsEmpty(qte)) And (pers = "" Or IsEmpty(pers)) And (h = "" Or IsEmpty(h))
            If Not estRecapTache Then
                t.Summary = False
            End If
            
            ' Stocker dans les tableaux
            nbTaches = nbTaches + 1
            Set tabTaches(nbTaches) = t
            tabNiveaux(nbTaches) = niveauxCalcules(i)
            tabNoms(nbTaches) = nom
            tabQte(nbTaches) = qte
            tabPers(nbTaches) = pers
            tabH(nbTaches) = h
        End If
    Next i
    
    ' Étape 3 : Indentation RAPIDE par sélection multiple (beaucoup plus rapide !)
    ' Désactiver les mises à jour d'écran pour accélérer
    pjApp.ScreenUpdating = False
    
    ' Regrouper les tâches par niveau dans des arrays
    Dim tabTachesNiveau2() As Object
    Dim tabTachesNiveau3() As Object
    Dim tabIdxNiveau2() As Long
    Dim tabIdxNiveau3() As Long
    Dim nbNiveau2 As Long, nbNiveau3 As Long
    nbNiveau2 = 0
    nbNiveau3 = 0
    
    ' Compter les tâches par niveau
    For idxTache = 1 To nbTaches
        If tabNiveaux(idxTache) = 2 Then
            nbNiveau2 = nbNiveau2 + 1
        ElseIf tabNiveaux(idxTache) = 3 Then
            nbNiveau3 = nbNiveau3 + 1
        End If
    Next idxTache
    
    ' Redimensionner les arrays par niveau
    If nbNiveau2 > 0 Then
        ReDim tabTachesNiveau2(1 To nbNiveau2)
        ReDim tabIdxNiveau2(1 To nbNiveau2)
    End If
    If nbNiveau3 > 0 Then
        ReDim tabTachesNiveau3(1 To nbNiveau3)
        ReDim tabIdxNiveau3(1 To nbNiveau3)
    End If
    
    ' Remplir les arrays par niveau
    nbNiveau2 = 0
    nbNiveau3 = 0
    For idxTache = 1 To nbTaches
        If tabNiveaux(idxTache) = 2 Then
            nbNiveau2 = nbNiveau2 + 1
            Set tabTachesNiveau2(nbNiveau2) = tabTaches(idxTache)
            tabIdxNiveau2(nbNiveau2) = idxTache
        ElseIf tabNiveaux(idxTache) = 3 Then
            nbNiveau3 = nbNiveau3 + 1
            Set tabTachesNiveau3(nbNiveau3) = tabTaches(idxTache)
            tabIdxNiveau3(nbNiveau3) = idxTache
        End If
    Next idxTache
    
    ' MÉTHODE ULTRA-RAPIDE : Sélection multiple et indentation en une seule opération
    Dim k As Integer
    
    Debug.Print ""
    Debug.Print "=== INDENTATION PAR SÉLECTION MULTIPLE ==="
    
    ' Niveau 2 : Sélectionner toutes les tâches niveau 2 et les indenter ensemble
    If nbNiveau2 > 0 Then
        Debug.Print ">>> Sélection de " & nbNiveau2 & " tâche(s) niveau 2..."
        
        ' Sélectionner la première tâche
        Set t = tabTachesNiveau2(1)
        pjApp.SelectRow t.ID, RowRelative:=False
        Debug.Print "  - Tâche " & t.ID & " sélectionnée"
        
        ' Étendre la sélection aux autres tâches niveau 2
        For k = 2 To nbNiveau2
            Set t = tabTachesNiveau2(k)
            pjApp.SelectRow t.ID, RowRelative:=False, Add:=True
            Debug.Print "  - Tâche " & t.ID & " ajoutée à la sélection"
        Next k
        
        Debug.Print ">>> Indentation de toutes les tâches sélectionnées (niveau 2)..."
        ' Indenter toutes les tâches sélectionnées en une seule fois
        pjApp.OutlineIndent
        Debug.Print "  ✓ " & nbNiveau2 & " tâche(s) indentée(s) en une seule opération"
        
        ' Debug final
        For k = 1 To nbNiveau2
            idxTache = tabIdxNiveau2(k)
            Debug.Print "  [" & tabTachesNiveau2(k).ID & "] " & tabNoms(idxTache) & " -> niveau " & tabTachesNiveau2(k).OutlineLevel
        Next k
    End If
    
    ' Niveau 3 : Sélectionner toutes les tâches niveau 3 et les indenter 2 fois ensemble
    If nbNiveau3 > 0 Then
        Debug.Print ""
        Debug.Print ">>> Sélection de " & nbNiveau3 & " tâche(s) niveau 3..."
        
        ' Sélectionner la première tâche
        Set t = tabTachesNiveau3(1)
        pjApp.SelectRow t.ID, RowRelative:=False
        Debug.Print "  - Tâche " & t.ID & " sélectionnée"
        
        ' Étendre la sélection aux autres tâches niveau 3
        For k = 2 To nbNiveau3
            Set t = tabTachesNiveau3(k)
            pjApp.SelectRow t.ID, RowRelative:=False, Add:=True
            Debug.Print "  - Tâche " & t.ID & " ajoutée à la sélection"
        Next k
        
        Debug.Print ">>> Indentation de toutes les tâches sélectionnées (niveau 3 - 2 fois)..."
        ' Indenter toutes les tâches sélectionnées 2 fois en une seule opération
        pjApp.OutlineIndent
        Debug.Print "  - Première indentation effectuée"
        pjApp.OutlineIndent
        Debug.Print "  ✓ " & nbNiveau3 & " tâche(s) indentée(s) 2 fois en une seule opération"
        
        ' Debug final
        For k = 1 To nbNiveau3
            idxTache = tabIdxNiveau3(k)
            Debug.Print "  [" & tabTachesNiveau3(k).ID & "] " & tabNoms(idxTache) & " -> niveau " & tabTachesNiveau3(k).OutlineLevel
        Next k
    End If
    
    Debug.Print ""
    Debug.Print "=== INDENTATION TERMINÉE ==="
    
    ' Forcer Summary = False pour toutes les tâches non-récapitulatives
    ' (MS Project peut créer automatiquement des summary tasks après l'indentation)
    Debug.Print ""
    Debug.Print ">>> Correction des summary tasks incorrectes..."
    For idxTache = 1 To nbTaches
        Set t = tabTaches(idxTache)
        ' Si la tâche a des données (B, C, D non vides), elle ne doit PAS être une summary task
        estRecapTache = (tabQte(idxTache) = "" Or IsEmpty(tabQte(idxTache))) And _
                        (tabPers(idxTache) = "" Or IsEmpty(tabPers(idxTache))) And _
                        (tabH(idxTache) = "" Or IsEmpty(tabH(idxTache)))
        
        If Not estRecapTache And t.Summary Then
            t.Summary = False
            Debug.Print "  - Tâche [" & t.ID & "] " & tabNoms(idxTache) & " : Summary forcé à False"
        End If
    Next idxTache
    
    ' Réactiver les mises à jour d'écran
    pjApp.ScreenUpdating = True
    
    ' Étape 4 : Assignations (traitement par batch pour les tâches niveau 3)
    Dim niveauCible As Integer
    For k = 1 To nbNiveau3
        idxTache = tabIdxNiveau3(k)
        Set t = tabTachesNiveau3(k)
        niveauCible = 3
        
        ' Assignation Monteurs
        If IsNumeric(tabH(idxTache)) And tabH(idxTache) > 0 Then
            Dim nbPers As Double
            nbPers = IIf(IsNumeric(tabPers(idxTache)) And tabPers(idxTache) > 0, tabPers(idxTache), 1)
            
            Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
            a.Units = nbPers
            a.Work = tabH(idxTache) * 60
        End If

        ' Assignation matériel
        If IsNumeric(tabQte(idxTache)) And tabQte(idxTache) > 0 Then
            Dim rMat As Resource
            Set rMat = GetOrCreateMaterialResource(tabNoms(idxTache))
            Set a = t.Assignments.Add(ResourceID:=rMat.ID)
            a.Units = tabQte(idxTache)
        End If
    Next k

    ' -----------------------------------------------------------
    ' ?? ORDONNANCEMENT
    ' -----------------------------------------------------------
    
    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "=== ORDONNANCEMENT ==="
    Debug.Print "========================================"
    
    Dim capaciteEquipe As Long
    capaciteEquipe = LireCapaciteDepuisG1(xlSheet)
    
    Debug.Print "Capacite equipe (G1) = " & capaciteEquipe & " personnes"
    
    rMonteurs.MaxUnits = capaciteEquipe
    Debug.Print "MaxUnits Monteurs = " & capaciteEquipe
    
    Debug.Print ""
    Debug.Print "Lancement du nivellement..."
    
    ' Désactiver les alertes pour éviter les boîtes de dialogue de surutilisation
    Dim alertsActives As Boolean
    alertsActives = pjApp.DisplayAlerts
    pjApp.DisplayAlerts = False
    
    ' Configurer le nivellement pour ignorer automatiquement les erreurs
    On Error Resume Next
    LevelNow
    On Error GoTo 0
    
    ' Réactiver les alertes
    pjApp.DisplayAlerts = alertsActives
    
    Debug.Print "Nivellement termine (surutilisations ignorees automatiquement)"
    
    Debug.Print ""
    Debug.Print "========================================"
    Debug.Print "? IMPORT TERMINE"
    Debug.Print "========================================"

    ' ==== FERMETURE ====
    xlBook.Close False
    xlApp.Quit
    Set xlApp = Nothing
    
    MsgBox "Import termine !" & vbCrLf & _
           "Capacite equipe : " & capaciteEquipe & " personnes" & vbCrLf & _
           "Structure hierarchique creee automatiquement", _
           vbInformation

End Sub


' -----------------------------------------------------------
' FONCTIONS UTILITAIRES
' -----------------------------------------------------------

Function GetOrCreateWorkResource(nom As String) As Resource
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeWork
    End If
    On Error GoTo 0
    Set GetOrCreateWorkResource = r
End Function

Function GetOrCreateMaterialResource(nom As String) As Resource
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeMaterial
    End If
    On Error GoTo 0
    Set GetOrCreateMaterialResource = r
End Function

Function LireCapaciteDepuisG1(xlSheet As Object) As Long
    Dim valG1 As Variant
    valG1 = xlSheet.Cells(1, 7).Value
    
    If IsNumeric(valG1) And valG1 > 0 Then
        LireCapaciteDepuisG1 = CLng(valG1)
    Else
        LireCapaciteDepuisG1 = 1
    End If
End Function
