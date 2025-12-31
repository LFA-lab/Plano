Sub Import_Prevencheres_WBS()
' ===========================================================================================
' IMPORT PREVENCHERES - VERSION WBS SIMPLIFIÉE
' ===========================================================================================
' Cette version utilise la structure hiérarchique DIRECTEMENT depuis le fichier Excel.
' Les titres (lignes vides) définissent la hiérarchie : pas de calcul de niveau par code.
' 
' Structure Excel attendue :
'   - Ligne avec données (Qté/Pers/Heures) = TACHE
'   - Ligne sans données = TITRE/GROUPE
'   - La numérotation WBS dans le nom (1, 1.1, 1.1.1) définit le niveau automatiquement
' 
' Exemple de structure :
'   PREVENCHERES (root)
'   ├─ ZONE 3A (MECA)
'   │  ├─ Zone 3.1
'   │  │  └─ Tâches...
'   │  └─ Zone 3.2
'   │     └─ Tâches...
'   └─ 1 ELEC – ZONE 3A
'      ├─ 1.1 Pose chemins de câbles BT
'      │  ├─ 1.1.1 OND1
'      │  ├─ 1.1.2 OND2
'      │  └─ 1.1.3 OND3
'      └─ 1.2 Raccordement onduleurs
'         ├─ 1.2.1 OND1
'         ├─ 1.2.2 OND2
'         └─ 1.2.3 OND3
' ===========================================================================================

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim i As Long, lastRow As Long
    Dim t As Task, tCQ As Task, a As Assignment
    Dim fichierExcel As String
    Dim oldCalculation As Boolean

    ' ==== SELECTION DU FICHIER VIA SELECTEUR NATIF ====
    Dim xlTempApp As Object
    Set xlTempApp = CreateObject("Excel.Application")
    xlTempApp.Visible = False

    With xlTempApp.FileDialog(msoFileDialogFilePicker)
        .Title = "Sélectionnez le fichier Excel PREVENCHERES à importer"
        .InitialFileName = Environ$("USERPROFILE") & "\Downloads\"
        .Filters.Clear
        .Filters.Add "Fichiers Excel", "*.xlsx;*.xls"
        .AllowMultiSelect = False
        If .Show = -1 Then
            fichierExcel = .SelectedItems(1)
        Else
            MsgBox "Aucun fichier sélectionné. Import annulé.", vbExclamation
            xlTempApp.Quit
            Set xlTempApp = Nothing
            Exit Sub
        End If
    End With

    xlTempApp.Quit
    Set xlTempApp = Nothing

    ' ==== OUVERTURE D'EXCEL (LECTURE) ====
    Set xlApp = CreateObject("Excel.Application")
    xlApp.Visible = False
    xlApp.DisplayAlerts = False
    Set xlBook = xlApp.Workbooks.Open(FileName:=fichierExcel, ReadOnly:=True, UpdateLinks:=False)
    Set xlSheet = xlBook.Sheets(1)

    ' ==== OUVERTURE DE MS PROJECT ====
    Set pjApp = MSProject.Application
    pjApp.Visible = True
    pjApp.FileNew
    Set pjProj = pjApp.ActiveProject

    ' ==== LIBELLES DES CHAMPS TEXTE POUR L'IHM ====
    pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
    pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
    pjApp.CustomFieldRename pjCustomTaskText3, "Sous-Zone"
    pjApp.CustomFieldRename pjCustomTaskText4, "Metier"
    pjApp.CustomFieldRename pjCustomTaskText5, "Entreprise"
    pjApp.CustomFieldRename pjCustomTaskText6, "Niveau"
    pjApp.CustomFieldRename pjCustomTaskText7, "Onduleur"

    ' ==== AJOUT DU TITRE DE PROJET (CELLULE A2) ====
    Dim tRoot As Task
    Set tRoot = pjProj.Tasks.Add(Name:=xlSheet.Cells(2, 1).Value, Before:=1)
    tRoot.Manual = False
    tRoot.Calendar = ActiveProject.BaseCalendars("Standard")
    tRoot.OutlineLevel = 1
    
    ' Variable pour gérer la hiérarchie des groupes
    Dim tGroup As Task
    Set tGroup = tRoot

    ' ==== CONFIGURATION PROJET ====
    pjProj.DefaultTaskType = pjFixedWork
    pjProj.ScheduleFromStart = True
    pjProj.DefaultEffortDriven = True

    ' ==== MODIFICATION DU CALENDRIER "Standard" ====
    With ActiveProject.BaseCalendars("Standard").WorkWeeks
        .Add Start:="01/01/2025", Finish:="01/01/2027", Name:="Calendrier Standard"
        With .Item(1)
            Dim j As Integer
            For j = 2 To 6 ' Lundi à vendredi
                With .WeekDays(j)
                    .Shift1.Start = "09:00"
                    .Shift1.Finish = "18:00"
                    .Shift2.Clear: .Shift3.Clear: .Shift4.Clear: .Shift5.Clear
                End With
            Next j
            .WeekDays(1).Default ' dimanche
            .WeekDays(7).Default ' samedi
        End With
    End With

    ' ==== RESSOURCES STANDARD ====
    Dim rMonteurs As Resource
    Set rMonteurs = GetOrCreateWorkResource("Monteurs")
    rMonteurs.MaxUnits = 10

    ' Ressource matérielle CQ pour tous les contrôles (OMX et SST)
    Dim rCQMat As Resource
    Set rCQMat = GetOrCreateMaterialResource("CQ")
    
    ' ==== DÉSACTIVER CALCUL AUTOMATIQUE PENDANT L'IMPORT ====
    On Error Resume Next
    oldCalculation = pjApp.Calculation
    pjApp.Calculation = False
    On Error GoTo 0

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row

    ' ==== FICHIER LOG ====
    Dim logFile As String
    logFile = Replace(fichierExcel, ".xlsx", "_import_WBS_log.txt")
    logFile = Replace(logFile, ".xls", "_import_WBS_log.txt")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logStream As Object
    Set logStream = fso.CreateTextFile(logFile, True)
    
    logStream.WriteLine "===== DEBUT IMPORT PREVENCHERES WBS - " & Now & " ====="
    logStream.WriteLine "Fichier source: " & fichierExcel
    logStream.WriteLine "Nombre de lignes: " & lastRow
    logStream.WriteLine ""

    ' ==== BOUCLE PRINCIPALE ====
    For i = 3 To lastRow

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
        Dim qualite As String, niveau As String, onduleur As String
        Dim dateDebutMonteurs As Date, dateFinMonteurs As Date
        Dim hasMonteursAssignment As Boolean

        nom = Trim(CStr(xlSheet.Cells(i, 1).Value))
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value
        hasMonteursAssignment = False
        
        logStream.WriteLine "--- Ligne " & i & " ---"
        logStream.WriteLine "  Nom: " & nom

        zone = Trim(CStr(xlSheet.Cells(i, 5).Value))
        sousZone = Trim(CStr(xlSheet.Cells(i, 6).Value))
        tranche = Trim(CStr(xlSheet.Cells(i, 7).Value))
        typ = Trim(CStr(xlSheet.Cells(i, 8).Value))
        entreprise = Trim(CStr(xlSheet.Cells(i, 9).Value))
        qualite = UCase$(Trim(CStr(xlSheet.Cells(i, 10).Value)))
        niveau = UCase$(Trim(CStr(xlSheet.Cells(i, 11).Value)))
        onduleur = UCase$(Trim(CStr(xlSheet.Cells(i, 12).Value)))

        If nom = "" Then
            logStream.WriteLine "  -> Ligne ignorée (nom vide)"
            logStream.WriteLine ""
            GoTo NextRow
        End If
        
        ' ==== DETECTION TITRE/GROUPE (ligne sans données) ====
        Dim isTitle As Boolean
        isTitle = IsEmptyOrZero(qte) And IsEmptyOrZero(pers) And IsEmptyOrZero(h)
        
        If isTitle Then
            ' ===== CRÉATION D'UN TITRE/GROUPE =====
            Set tGroup = pjProj.Tasks.Add(nom)
            tGroup.Manual = False
            
            ' Détection automatique du niveau par numérotation WBS
            Dim detectedLevel As Integer
            detectedLevel = DetectHierarchyLevel(nom)
            
            ' Forcer le niveau détecté
            Dim maxAttempts As Integer
            maxAttempts = 0
            On Error Resume Next
            Do While tGroup.OutlineLevel < detectedLevel And maxAttempts < 10
                tGroup.OutlineIndent
                maxAttempts = maxAttempts + 1
                If Err.Number <> 0 Then Exit Do
            Loop
            Err.Clear
            maxAttempts = 0
            Do While tGroup.OutlineLevel > detectedLevel And maxAttempts < 10
                tGroup.OutlineOutdent
                maxAttempts = maxAttempts + 1
                If Err.Number <> 0 Then Exit Do
            Loop
            On Error GoTo 0
            
            logStream.WriteLine "  -> TITRE créé: " & nom & " (Niveau détecté: " & detectedLevel & ", Niveau réel: " & tGroup.OutlineLevel & ")"
            logStream.WriteLine ""
            GoTo NextRow
        End If

        ' ===== CRÉATION D'UNE TACHE =====
        Set t = pjProj.Tasks.Add(nom)
        t.Manual = False
        t.Calendar = ActiveProject.BaseCalendars("Standard")
        t.LevelingCanSplit = False
        
        ' ==== NIVEAU HIERARCHIQUE : Toujours groupe + 1 ====
        Dim targetLevel As Integer
        If Not tGroup Is Nothing Then
            targetLevel = tGroup.OutlineLevel + 1
        Else
            targetLevel = 2
        End If
        
        ' Forcer le niveau (relatif au groupe parent)
        maxAttempts = 0
        On Error Resume Next
        Do While t.OutlineLevel < targetLevel And maxAttempts < 10
            t.OutlineIndent
            maxAttempts = maxAttempts + 1
            If Err.Number <> 0 Then
                logStream.WriteLine "  -> ATTENTION: Impossible d'atteindre le niveau " & targetLevel
                Exit Do
            End If
        Loop
        On Error GoTo 0
        
        logStream.WriteLine "  -> Tâche créée: " & t.Name & " (ID: " & t.ID & ", Niveau: " & t.OutlineLevel & ")"

        ' ==== TAGS DANS CHAMPS TEXTE ====
        t.Text1 = tranche
        t.Text2 = zone
        t.Text3 = sousZone
        t.Text4 = typ
        t.Text5 = entreprise
        t.Text6 = niveau
        t.Text7 = onduleur

        ' ==== TRAVAIL (Monteurs) ====
        If IsNumeric(h) And h > 0 Then
            Dim workMinutes As Long
            workMinutes = CLng(CDbl(h) * 60)
            t.Type = pjFixedWork
            t.Work = workMinutes
            
            Dim nbPers As Long
            nbPers = IIf(IsNumeric(pers) And pers > 0, CLng(pers), 1)
            
            Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
            a.Work = workMinutes
            a.Units = nbPers
            a.Work = workMinutes ' Forcer à nouveau
            a.WorkContour = pjFlat
            
            dateDebutMonteurs = a.Start
            dateFinMonteurs = a.Finish
            hasMonteursAssignment = True
            
            ' Copie tags
            a.Text1 = tranche
            a.Text2 = zone
            a.Text3 = sousZone
            a.Text4 = typ
            a.Text5 = entreprise
            a.Text6 = niveau
            a.Text7 = onduleur
            
            logStream.WriteLine "  -> Travail: " & workMinutes & " min (" & CDbl(h) & "h) | Personnes: " & nbPers
        End If
        
        ' ==== QUANTITÉ (Matériau) ====
        If IsNumeric(qte) And qte > 0 Then
            ' Utiliser le nom du groupe parent comme ressource matérielle
            ' Cela permet d'agréger les quantités par activité (ex: "Remontées du serpentins")
            Dim nomRessource As String
            If Not tGroup Is Nothing Then
                nomRessource = tGroup.Name  ' Nom de la tâche récap parente
                logStream.WriteLine "  -> Ressource matérielle: " & nomRessource & " (depuis groupe parent)"
            Else
                nomRessource = nom  ' Fallback: nom de la tâche
                logStream.WriteLine "  -> Ressource matérielle: " & nomRessource & " (nom de tâche)"
            End If
            
            Dim rMat As Resource
            Set rMat = GetOrCreateMaterialResource(nomRessource)
            
            Set a = t.Assignments.Add(ResourceID:=rMat.ID)
            a.Units = CDbl(qte)
            a.WorkContour = pjFlat
            
            If hasMonteursAssignment Then
                a.Start = dateDebutMonteurs
                a.Finish = dateFinMonteurs
            End If
            
            ' Copie tags
            a.Text1 = tranche
            a.Text2 = zone
            a.Text3 = sousZone
            a.Text4 = typ
            a.Text5 = entreprise
            a.Text6 = niveau
            a.Text7 = onduleur
            
            logStream.WriteLine "  -> Quantité: " & CDbl(qte) & " unités de '" & nomRessource & "'"
        End If

        ' ==== QUALITÉ (CQ/TACHE) ====
        Dim isOmx As Boolean
        isOmx = (UCase$(entreprise) = "OMX" Or UCase$(entreprise) = "OMEXOM")
        
        If qualite = "CQ" Then
            If isOmx Then
                ' ===== CQ OMX : Ressource MATERIELLE (consommable) sur la tâche =====
                ' But : Vérifier la cadence des contrôles (intérimaires dédiés)
                Set a = t.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1  ' 1 contrôle prévu (ajustable selon besoin)
                a.WorkContour = pjFlat
                
                If hasMonteursAssignment Then
                    a.Start = dateDebutMonteurs
                    a.Finish = dateFinMonteurs
                End If
                
                a.Text1 = tranche: a.Text2 = zone: a.Text3 = sousZone
                a.Text4 = typ: a.Text5 = entreprise: a.Text6 = niveau: a.Text7 = onduleur
                
                logStream.WriteLine "  -> CQ OMX ajouté (ressource CQ, 1 contrôle)"
            Else
                ' ===== CQ SST : Tâche séparée avec ressource CQ =====
                ' But : Visualiser le besoin de contrôle sur la zone
                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                tCQ.Manual = False
                tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
                tCQ.LevelingCanSplit = False
                
                ' Même niveau que la tâche principale
                On Error Resume Next
                Do While tCQ.OutlineLevel < t.OutlineLevel
                    tCQ.OutlineIndent
                Loop
                On Error GoTo 0
                
                tCQ.Text1 = tranche: tCQ.Text2 = zone: tCQ.Text3 = sousZone
                tCQ.Text4 = "CQ": tCQ.Text5 = "OMEXOM": tCQ.Text6 = niveau: tCQ.Text7 = onduleur
                
                ' Ressource matérielle CQ
                Set a = tCQ.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1  ' 1 contrôle
                a.WorkContour = pjFlat
                
                ' Créer une dépendance DÉBUT-DÉBUT : CQ démarre 1 jour après le début de la tâche
                ' Utiliser LinkSuccessors depuis la tâche principale (plus fiable)
                Dim errNum As Long, errDesc As String
                On Error Resume Next
                t.LinkSuccessors tCQ, pjStartToStart, "1d"
                errNum = Err.Number
                errDesc = Err.Description
                On Error GoTo 0
                
                If errNum = 0 Then
                    logStream.WriteLine "  -> Tâche CQ SST créée (ressource CQ, dépendance DD+1j OK)"
                Else
                    logStream.WriteLine "  -> Tâche CQ SST créée (ressource CQ, ERREUR dépendance: " & errNum & " - " & errDesc & ")"
                End If
            End If
            
        ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
            ' ===== Force tâche CQ séparée (même pour OMX) =====
            Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
            tCQ.Manual = False
            tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
            tCQ.LevelingCanSplit = False
            
            On Error Resume Next
            Do While tCQ.OutlineLevel < t.OutlineLevel
                tCQ.OutlineIndent
            Loop
            On Error GoTo 0
            
            tCQ.Text1 = tranche: tCQ.Text2 = zone: tCQ.Text3 = sousZone
            tCQ.Text4 = "CQ": tCQ.Text5 = "OMEXOM": tCQ.Text6 = niveau: tCQ.Text7 = onduleur
            
            ' Ressource matérielle CQ
            Set a = tCQ.Assignments.Add(ResourceID:=rCQMat.ID)
            a.Units = 1: a.WorkContour = pjFlat
            
            ' Créer une dépendance DÉBUT-DÉBUT +1 jour
            ' Utiliser LinkSuccessors depuis la tâche principale (plus fiable)
            Dim errNum2 As Long, errDesc2 As String
            On Error Resume Next
            t.LinkSuccessors tCQ, pjStartToStart, "1d"
            errNum2 = Err.Number
            errDesc2 = Err.Description
            On Error GoTo 0
            
            If errNum2 = 0 Then
                logStream.WriteLine "  -> Tâche CQ explicite créée (ressource CQ, dépendance DD+1j OK)"
            Else
                logStream.WriteLine "  -> Tâche CQ explicite créée (ressource CQ, ERREUR dépendance: " & errNum2 & " - " & errDesc2 & ")"
            End If
        End If
        
NextRow:
        logStream.WriteLine ""
    Next i

    ' ==== AFFICHAGE STRUCTURE ====
    logStream.WriteLine ""
    logStream.WriteLine "===== STRUCTURE HIERARCHIQUE MS PROJECT ====="
    Dim tDebug As Task
    For Each tDebug In pjProj.Tasks
        If Not tDebug Is Nothing Then
            Dim indent As String
            indent = String((tDebug.OutlineLevel - 1) * 2, " ")
            Dim prefix As String
            prefix = IIf(tDebug.Summary, "[GROUPE]", "[TACHE ]")
            logStream.WriteLine indent & prefix & " [Niv " & tDebug.OutlineLevel & "] " & tDebug.Name
        End If
    Next tDebug
    logStream.WriteLine "===== FIN STRUCTURE ====="

    ' ==== FERMETURE LOG ====
    logStream.WriteLine ""
    logStream.WriteLine "===== FIN IMPORT - " & Now & " ====="
    logStream.Close
    Set logStream = Nothing
    Set fso = Nothing

    ' ==== CALCUL FINAL ====
    On Error Resume Next
    pjApp.Calculation = True
    pjProj.Calculate
    pjApp.CalculateAll
    On Error GoTo 0

    ' ==== FERMETURE ====
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "Import PREVENCHERES terminé avec structure WBS." & vbCrLf & vbCrLf & "Log: " & logFile, vbInformation

End Sub


' ===== FONCTION : Détection niveau hiérarchique par numérotation WBS =====
Private Function DetectHierarchyLevel(nom As String) As Integer
    ' Détecte le niveau en comptant les points dans la numérotation
    ' Exemples :
    '   "ZONE 3A (MECA)"          → niveau 2 (pas de numéro)
    '   "Zone 3.1"                → niveau 2 (pas de numéro)
    '   "1 ELEC - ZONE 3A"        → niveau 2 (1 chiffre seul)
    '   "1.1 Pose chemins"        → niveau 3 (1 point)
    '   "1.1.1 OND1"              → niveau 4 (2 points)
    
    Dim firstWord As String
    Dim pointCount As Integer
    
    ' Extraire le premier mot (la numérotation)
    If InStr(nom, " ") > 0 Then
        firstWord = Trim$(Left$(nom, InStr(nom, " ") - 1))
    Else
        firstWord = nom
    End If
    
    ' Si c'est une numérotation valide (ex: "1", "1.1", "1.1.1")
    If IsNumericPattern(firstWord) Then
        ' Compter les points
        pointCount = Len(firstWord) - Len(Replace(firstWord, ".", ""))
        ' Niveau = nombre de points + 2 (car niveau 1 = root)
        DetectHierarchyLevel = pointCount + 2
    Else
        ' Pas de numérotation détectée → niveau 2 par défaut
        DetectHierarchyLevel = 2
    End If
End Function

Private Function IsNumericPattern(text As String) As Boolean
    ' Vérifie si c'est un pattern numérique type "1", "1.1", "1.1.1"
    Dim i As Integer
    Dim ch As String
    
    If text = "" Then
        IsNumericPattern = False
        Exit Function
    End If
    
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        If Not (ch >= "0" And ch <= "9") And ch <> "." Then
            IsNumericPattern = False
            Exit Function
        End If
    Next i
    
    IsNumericPattern = True
End Function

Private Function IsEmptyOrZero(v As Variant) As Boolean
    If IsEmpty(v) Then
        IsEmptyOrZero = True
    ElseIf Trim$(CStr(v)) = "" Then
        IsEmptyOrZero = True
    ElseIf IsNumeric(v) Then
        IsEmptyOrZero = (CDbl(v) = 0)
    Else
        IsEmptyOrZero = False
    End If
End Function

Function GetOrCreateWorkResource(nom As String) As Resource
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeWork
    End If
    Set GetOrCreateWorkResource = r
End Function

Function GetOrCreateMaterialResource(nom As String) As Resource
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeMaterial
    End If
    Set GetOrCreateMaterialResource = r
End Function

