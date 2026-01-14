Sub Import_Taches_Simples_AvecTitre()

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
        .Title = "Sélectionnez le fichier Excel à importer"
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
    ' Champs des tâches
    pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
    pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
    pjApp.CustomFieldRename pjCustomTaskText3, "Sous-Zone"
    pjApp.CustomFieldRename pjCustomTaskText4, "Metier"
    pjApp.CustomFieldRename pjCustomTaskText5, "Entreprise"
    pjApp.CustomFieldRename pjCustomTaskText6, "Niveau"
    pjApp.CustomFieldRename pjCustomTaskText7, "Onduleur"
    pjApp.CustomFieldRename pjCustomTaskText8, "PTR"
    
    ' NOTE: Les champs d'assignments (Text1-Text7) ne peuvent pas être renommés via CustomFieldRename
    ' Ils utiliseront les noms par défaut "Text1", "Text2", etc. dans l'interface
    ' Mais les DONNÉES seront bien stockées dans Assignment.Text1 à Assignment.Text7

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
    rMonteurs.MaxUnits = 10 ' 1000% = 10 personnes max (large pour éviter surutilisation)

    ' Ressource matérielle CQ pour tous les contrôles (OMX et SST)
    Dim rCQMat As Resource
    Set rCQMat = GetOrCreateMaterialResource("CQ")
    
    ' ==== DÉSACTIVER CALCUL AUTOMATIQUE PENDANT L'IMPORT ====
    ' Évite les popups de surutilisation
    On Error Resume Next
    oldCalculation = pjApp.Calculation
    pjApp.Calculation = False
    On Error GoTo 0

    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' fin de la colonne A

    ' ==== FICHIER LOG ====
    Dim logFile As String
    logFile = Replace(fichierExcel, ".xlsx", "_import_log.txt")
    logFile = Replace(logFile, ".xls", "_import_log.txt")
    Dim fso As Object
    Set fso = CreateObject("Scripting.FileSystemObject")
    Dim logStream As Object
    Set logStream = fso.CreateTextFile(logFile, True)
    
    logStream.WriteLine "===== DEBUT IMPORT - " & Now & " ====="
    logStream.WriteLine "Fichier source: " & fichierExcel
    logStream.WriteLine "Nombre de lignes: " & lastRow
    logStream.WriteLine ""
    
    ' ==== APERCU FICHIER EXCEL ====
    logStream.WriteLine "===== APERCU FICHIER EXCEL (COLONNES A, K, L) ====="
    Dim iPreview As Long
    For iPreview = 2 To lastRow
        Dim nomPreview As String, niveauPreview As String, onduleurPreview As String
        Dim qtePreview As Variant, persPreview As Variant, hPreview As Variant
        
        nomPreview = Trim(CStr(xlSheet.Cells(iPreview, 1).Value))
        qtePreview = xlSheet.Cells(iPreview, 2).Value
        persPreview = xlSheet.Cells(iPreview, 3).Value
        hPreview = xlSheet.Cells(iPreview, 4).Value
        niveauPreview = Trim(CStr(xlSheet.Cells(iPreview, 11).Value))
        onduleurPreview = Trim(CStr(xlSheet.Cells(iPreview, 12).Value))
        
        Dim typePreview As String
        Dim hasData As Boolean
        hasData = (Not IsEmpty(qtePreview) And qtePreview <> "") Or _
                  (Not IsEmpty(persPreview) And persPreview <> "") Or _
                  (Not IsEmpty(hPreview) And hPreview <> "")
        
        If hasData Then
            typePreview = "[TACHE]"
        Else
            typePreview = "[TITRE]"
        End If
        
        Dim niveauDetecte As String
        If Not hasData Then
            ' Simuler la détection de niveau
            Dim firstWordPreview As String
            If InStr(nomPreview, " ") > 0 Then
                firstWordPreview = Trim$(Left$(nomPreview, InStr(nomPreview, " ") - 1))
            Else
                firstWordPreview = nomPreview
            End If
            
            If IsNumericPattern(firstWordPreview) Then
                Dim pointCountPreview As Integer
                pointCountPreview = Len(firstWordPreview) - Len(Replace(firstWordPreview, ".", ""))
                niveauDetecte = " -> Niv " & (pointCountPreview + 2)
            Else
                niveauDetecte = " -> Niv 2 (défaut)"
            End If
        Else
            niveauDetecte = ""
        End If
        
        logStream.WriteLine "Ligne " & Format(iPreview, "00") & " " & typePreview & niveauDetecte & " | " & nomPreview & _
            IIf(niveauPreview <> "", " | K=" & niveauPreview, "") & _
            IIf(onduleurPreview <> "", " | L=" & onduleurPreview, "")
    Next iPreview
    logStream.WriteLine ""
    logStream.WriteLine "===== FIN APERCU EXCEL ====="
    logStream.WriteLine ""

    ' ==== BOUCLE TACHES ====
    For i = 3 To lastRow

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
        Dim qualite As String, niveau As String, onduleur As String, ptr As String
        Dim dateDebutMonteurs As Date, dateFinMonteurs As Date
        Dim hasMonteursAssignment As Boolean

        nom = Trim(CStr(xlSheet.Cells(i, 1).Value))
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value
        hasMonteursAssignment = False
        
        ' LOG LIGNE COURANTE
        logStream.WriteLine "--- Ligne " & i & " ---"
        logStream.WriteLine "  Nom: " & nom
        
        ' ==== LOG DETAILLE POUR DIAGNOSTIC ERREUR 1101 ====
        logStream.WriteLine "  [DIAG] Valeur brute: """ & nom & """"
        logStream.WriteLine "  [DIAG] TypeName: " & TypeName(nom)
        logStream.WriteLine "  [DIAG] Len(nom): " & Len(nom)
        logStream.WriteLine "  [DIAG] Trim(nom) = """": " & CBool(Trim(nom) = "")
        
        ' Analyse caractère par caractère
        If Len(nom) > 0 Then
            Dim analyseChars As String
            analyseChars = AnalyzeStringCharacters(nom)
            logStream.WriteLine "  [DIAG] Codes ASCII/Unicode des caractères:"
            logStream.WriteLine analyseChars
            
            ' Détection caractères invisibles
            If IsInvisibleOnlyString(nom) Then
                logStream.WriteLine "  [DIAG] *** ATTENTION: nom contient UNIQUEMENT des caractères invisibles ***"
            End If
        End If

        zone = Trim(CStr(xlSheet.Cells(i, 5).Value))        ' E
        sousZone = Trim(CStr(xlSheet.Cells(i, 6).Value))    ' F
        tranche = Trim(CStr(xlSheet.Cells(i, 7).Value))     ' G
        typ = Trim(CStr(xlSheet.Cells(i, 8).Value))         ' H
        entreprise = Trim(CStr(xlSheet.Cells(i, 9).Value))  ' I
        qualite = UCase$(Trim(CStr(xlSheet.Cells(i, 10).Value))) ' J : CQ / TACHE / vide
        niveau = UCase$(Trim(CStr(xlSheet.Cells(i, 11).Value)))  ' K : SZ / OND / vide
        onduleur = UCase$(Trim(CStr(xlSheet.Cells(i, 12).Value))) ' L : OND1, OND2...
        
        ' Lecture PTR (colonne 13 / M) - Rétrocompatible si absente
        On Error Resume Next
        ptr = Trim(CStr(xlSheet.Cells(i, 13).Value))        ' M : PTR (optionnel)
        If Err.Number <> 0 Then ptr = "" ' Si erreur (colonne absente), PTR vide
        On Error GoTo 0

        ' LOG VALEURS BRUTES
        logStream.WriteLine "  Qte (col B): " & qte & " | Type: " & TypeName(qte)
        logStream.WriteLine "  Pers (col C): " & pers & " | Type: " & TypeName(pers)
        logStream.WriteLine "  Heures (col D): " & h & " | Type: " & TypeName(h)
        logStream.WriteLine "  Zone: " & zone & " | Tranche: " & tranche
        logStream.WriteLine "  Type: " & typ & " | Entreprise: " & entreprise
        logStream.WriteLine "  Qualité: " & qualite & " | Niveau: " & niveau & " | Onduleur: " & onduleur & " | PTR: " & ptr

        If nom = "" Then
            logStream.WriteLine "  -> Ligne ignorée (nom vide)"
            logStream.WriteLine ""
            GoTo NextRow
        End If
        
        ' ==== DETECTION TITRE (ligne sans données ni tags) ====
        Dim isTitle As Boolean
        isTitle = IsEmptyOrZero(qte) And IsEmptyOrZero(pers) And IsEmptyOrZero(h) _
                  And zone = "" And sousZone = "" And tranche = "" And typ = "" _
                  And entreprise = "" And qualite = "" And niveau = "" And onduleur = ""
        
        If isTitle Then
            ' Créer un groupe/titre - toujours niveau 2
            logStream.WriteLine "  [DIAG] Tentative de création TITRE avec Tasks.Add(nom)..."
            On Error Resume Next
            Set tGroup = pjProj.Tasks.Add(nom)
            If Err.Number <> 0 Then
                logStream.WriteLine "  [DIAG] *** ERREUR Tasks.Add() pour TITRE ***"
                logStream.WriteLine "  [DIAG] Err.Number: " & Err.Number
                logStream.WriteLine "  [DIAG] Err.Description: " & Err.Description
                logStream.WriteLine "  [DIAG] Valeur de nom au moment de l'erreur: """ & nom & """"
                logStream.WriteLine "  [DIAG] Len(nom): " & Len(nom)
                Err.Clear
                On Error GoTo 0
                logStream.WriteLine ""
                GoTo NextRow
            End If
            On Error GoTo 0
            
            tGroup.Manual = False
            tGroup.OutlineLevel = 2  ' Tous les titres au niveau 2
            
            logStream.WriteLine "  -> TITRE créé: " & nom & " (Niveau " & tGroup.OutlineLevel & ")"
            logStream.WriteLine ""
            GoTo NextRow
        End If
        
        ' ==== VALIDATION Niveau/Onduleur ====
        If niveau = "OND" And onduleur = "" Then
            logStream.WriteLine "  -> ATTENTION: Niveau=OND mais Onduleur vide!"
        End If

        If nom <> "" Then

            ' ==== LOG DETAILLE AVANT Tasks.Add() ====
            logStream.WriteLine "  [DIAG] Tentative de création TACHE avec Tasks.Add(nom)..."
            logStream.WriteLine "  [DIAG] Valeur exacte de nom avant Tasks.Add: """ & nom & """"
            
            On Error Resume Next
            Set t = pjProj.Tasks.Add(nom)
            If Err.Number <> 0 Then
                logStream.WriteLine "  [DIAG] *** ERREUR Tasks.Add() pour TACHE ***"
                logStream.WriteLine "  [DIAG] Err.Number: " & Err.Number
                logStream.WriteLine "  [DIAG] Err.Description: " & Err.Description
                logStream.WriteLine "  [DIAG] Valeur de nom au moment de l'erreur: """ & nom & """"
                logStream.WriteLine "  [DIAG] Len(nom): " & Len(nom)
                logStream.WriteLine "  [DIAG] TypeName(nom): " & TypeName(nom)
                Err.Clear
                On Error GoTo 0
                logStream.WriteLine ""
                GoTo NextRow
            End If
            On Error GoTo 0
            logStream.WriteLine "  [DIAG] Tasks.Add() réussi - ID tâche: " & t.ID
            
            t.Manual = False
            t.Calendar = ActiveProject.BaseCalendars("Standard")
            t.LevelingCanSplit = False ' Empêche le fractionnement de la tâche
            
            ' ==== NIVEAU HIERARCHIQUE basé sur colonne K (Niveau) ====
            ' On crée d'abord la tâche, puis on ajuste son niveau avec OutlineIndent
            Dim targetLevel As Integer
            
            If niveau = "OND" Then
                targetLevel = 4  ' Tâches onduleurs au niveau 4
            ElseIf niveau = "SZ" Then
                targetLevel = 3  ' Tâches sous-zone au niveau 3
            ElseIf Not tGroup Is Nothing Then
                targetLevel = tGroup.OutlineLevel + 1
            Else
                targetLevel = 3  ' Par défaut niveau 3
            End If
            
            ' Forcer le niveau avec OutlineIndent/OutlineOutdent
            ' Protection contre erreurs 1101 et boucles infinies
            On Error Resume Next
            Do While t.OutlineLevel < targetLevel And t.OutlineLevel < 9 And Not t.Summary
                t.OutlineIndent
                If Err.Number <> 0 Then
                    logStream.WriteLine "  -> ATTENTION: Impossible d'indenter au niveau " & targetLevel & " (Erreur: " & Err.Number & ")"
                    Exit Do
                End If
            Loop
            Err.Clear
            Do While t.OutlineLevel > targetLevel And t.OutlineLevel > 1 And Not t.Summary
                t.OutlineOutdent
                If Err.Number <> 0 Then
                    logStream.WriteLine "  -> ATTENTION: Impossible de désindenter au niveau " & targetLevel & " (Erreur: " & Err.Number & ")"
                    Exit Do
                End If
            Loop
            On Error GoTo 0
            
            logStream.WriteLine "  -> Tâche créée: " & t.Name & " (ID: " & t.ID & ", Niveau: " & t.OutlineLevel & " - K=" & niveau & ")"

            ' Tags dans champs texte
            ' Convention proposée:
            ' Text1 = Tranche, Text2 = Zone, Text3 = Sous-zone, Text4 = Type, Text5 = Entreprise
            ' Text6 = Niveau, Text7 = Onduleur, Text8 = PTR
            t.Text1 = tranche
            t.Text2 = zone
            t.Text3 = sousZone
            t.Text4 = typ
            t.Text5 = entreprise
            t.Text6 = niveau
            t.Text7 = onduleur
            t.Text8 = ptr

            ' ✅ DÉFINIR LE TRAVAIL DE LA TÂCHE EN PREMIER (avant les assignments)
            ' Cela permet à MS Project de calculer correctement la durée
            If IsNumeric(h) And h > 0 Then
                Dim workMinutes As Long
                workMinutes = CLng(CDbl(h) * 60)
                t.Type = pjFixedWork
                t.Work = workMinutes
                logStream.WriteLine "  -> Travail de la tâche défini: " & workMinutes & " minutes (" & CDbl(h) & "h)"
            End If

            ' ✅ ORDRE CRITIQUE : d'abord TRAVAIL (pour calculer Duration), puis matériau et CQ
            
            ' Travail (Monteurs) - EN PREMIER pour que MS Project calcule t.Duration correctement
            If IsNumeric(h) And h > 0 Then
                Dim nbPers As Long
                nbPers = IIf(IsNumeric(pers) And pers > 0, CLng(pers), 1)
                
                logStream.WriteLine "  -> HEURES: h = " & h
                logStream.WriteLine "     nbPers = " & nbPers
                logStream.WriteLine "     workMinutes calculé = " & workMinutes

                Set a = t.Assignments.Add(ResourceID:=rMonteurs.ID)
                
                ' ÉTAPE 1: Assigner Work EN PREMIER
                a.Work = workMinutes
                
                ' ÉTAPE 2: Puis assigner Units
                a.Units = nbPers ' 1=100%, 2=200%, 3=300% automatiquement
                
                ' ÉTAPE 3: FORCER le Work à nouveau après Units
                a.Work = workMinutes
                
                ' ÉTAPE 4: Profil de travail régulier (répartition uniforme)
                a.WorkContour = pjFlat
                
                ' ÉTAPE 5: Sauvegarder les dates de l'assignment Monteurs
                dateDebutMonteurs = a.Start
                dateFinMonteurs = a.Finish
                hasMonteursAssignment = True
                
                ' ÉTAPE 6: Copie DIRECTE des tags
                a.Text1 = tranche
                a.Text2 = zone
                a.Text3 = sousZone
                a.Text4 = typ
                a.Text5 = entreprise
                a.Text6 = niveau
                a.Text7 = onduleur
                a.Text8 = ptr
                
                logStream.WriteLine "     Assignment.Units = " & a.Units
                logStream.WriteLine "     Assignment.Work FINAL = " & a.Work & " minutes"
                logStream.WriteLine "     Assignment Monteurs - Début: " & Format(a.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(a.Finish, "dd/mm/yyyy hh:nn")
                logStream.WriteLine "     Tags copiés: Tranche=" & tranche & " | Zone=" & zone & " | Type=" & typ & " | Niveau=" & niveau
            Else
                logStream.WriteLine "  -> HEURES IGNORÉES: h = " & h & " | IsNumeric = " & IsNumeric(h) & " | h > 0 = " & (h > 0)
            End If
            
            ' Quantité (matériau) - APRÈS le travail pour avoir la vraie durée
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
                
                Dim qteTotal As Double
                qteTotal = CDbl(qte)  ' Valeur Excel conservée
                
                ' ✅ Injecter simplement la quantité
                a.Units = qteTotal  ' La quantité Excel affichée dans MS Project
                a.WorkContour = pjFlat  ' Répartition régulière sur la durée
                
                ' ✅ FORCER les dates pour correspondre aux Monteurs si présents
                If hasMonteursAssignment Then
                    a.Start = dateDebutMonteurs
                    a.Finish = dateFinMonteurs
                    logStream.WriteLine "  -> QUANTITE: " & qteTotal & " unités de matériau '" & nomRessource & "' (dates synchronisées avec Monteurs)"
                Else
                    logStream.WriteLine "  -> QUANTITE: " & qteTotal & " unités de matériau '" & nomRessource & "' (pas de Monteurs, dates par défaut)"
                End If
                
                ' Copie DIRECTE des tags (sans passer par fonction)
                a.Text1 = tranche
                a.Text2 = zone
                a.Text3 = sousZone
                a.Text4 = typ
                a.Text5 = entreprise
                a.Text6 = niveau
                a.Text7 = onduleur
                a.Text8 = ptr
                
                logStream.WriteLine "     Tags copiés: Tranche=" & tranche & " | Zone=" & zone & " | Type=" & typ
                logStream.WriteLine "     Vérif lecture: a.Text1=" & a.Text1 & " | a.Text2=" & a.Text2
                logStream.WriteLine "     Assignment Matériau - Début: " & Format(a.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(a.Finish, "dd/mm/yyyy hh:nn")
            End If

            ' Qualité (J) : Logique hybride OMX/SST
            ' - Si entreprise = OMX et qualite = CQ : ajoute ressource MATERIELLE "CQ" sur la tâche
            ' - Si entreprise = SST et qualite = CQ : crée tâche CQ séparée avec ressource CQ
            ' - Si qualite = TACHE : force tâche CQ séparée (même pour OMX)
            Dim isOmx As Boolean
            isOmx = (UCase$(entreprise) = "OMX" Or UCase$(entreprise) = "OMEXOM")
            
            If qualite = "CQ" Then
                If isOmx Then
                    ' ===== CQ OMX : ressource MATERIELLE (consommable) =====
                    ' But : Vérifier la cadence des contrôles (intérimaires dédiés)
                    Set a = t.Assignments.Add(ResourceID:=rCQMat.ID)
                    a.Units = 1  ' 1 contrôle prévu (ajustable selon besoin)
                    a.WorkContour = pjFlat
                    
                    If hasMonteursAssignment Then
                        a.Start = dateDebutMonteurs
                        a.Finish = dateFinMonteurs
                        logStream.WriteLine "  -> QUALITE CQ OMX ajoutée (ressource CQ, dates sync)"
                    Else
                        logStream.WriteLine "  -> QUALITE CQ OMX ajoutée (ressource CQ)"
                    End If
                    
                    ' Copie tags
                    a.Text1 = tranche
                    a.Text2 = zone
                    a.Text3 = sousZone
                    a.Text4 = typ
                    a.Text5 = entreprise
                    a.Text6 = niveau
                    a.Text7 = onduleur
                    a.Text8 = ptr
                    
                    logStream.WriteLine "     Tags CQ copiés | Assignment CQ - Début: " & Format(a.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(a.Finish, "dd/mm/yyyy hh:nn")
                Else
                    ' ===== CQ SST : tâche séparée avec ressource CQ =====
                    ' But : Visualiser le besoin de contrôle sur la zone
                    Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                    tCQ.Manual = False
                    tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
                    tCQ.LevelingCanSplit = False
                    
                    ' Forcer le même niveau que la tâche principale
                    ' Protection contre erreurs 1101
                    On Error Resume Next
                    Do While tCQ.OutlineLevel < t.OutlineLevel And tCQ.OutlineLevel < 9 And Not tCQ.Summary
                        tCQ.OutlineIndent
                        If Err.Number <> 0 Then Exit Do
                    Loop
                    Err.Clear
                    Do While tCQ.OutlineLevel > t.OutlineLevel And tCQ.OutlineLevel > 1 And Not tCQ.Summary
                        tCQ.OutlineOutdent
                        If Err.Number <> 0 Then Exit Do
                    Loop
                    On Error GoTo 0
                    
                    ' Tags tâche CQ
                    tCQ.Text1 = tranche
                    tCQ.Text2 = zone
                    tCQ.Text3 = sousZone
                    tCQ.Text4 = "CQ"
                    tCQ.Text5 = "OMEXOM"  ' CQ porté par OMX
                    tCQ.Text6 = niveau
                    tCQ.Text7 = onduleur
                    tCQ.Text8 = ptr
                    
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
                        logStream.WriteLine "  -> TACHE CQ SST créée (ressource CQ, dépendance DD+1j OK)"
                    Else
                        logStream.WriteLine "  -> TACHE CQ SST créée (ressource CQ, ERREUR dépendance: " & errNum & " - " & errDesc & ")"
                    End If
                    
                    logStream.WriteLine "     Tags CQ copiés | Tâche CQ - Début: " & Format(tCQ.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(tCQ.Finish, "dd/mm/yyyy hh:nn")
                End If

            ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
                ' Force une tâche CQ séparée (même pour OMX)
                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                tCQ.Manual = False
                tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
                tCQ.LevelingCanSplit = False
                
                ' Forcer le même niveau que la tâche principale
                ' Protection contre erreurs 1101
                On Error Resume Next
                Do While tCQ.OutlineLevel < t.OutlineLevel And tCQ.OutlineLevel < 9 And Not tCQ.Summary
                    tCQ.OutlineIndent
                    If Err.Number <> 0 Then Exit Do
                Loop
                Err.Clear
                Do While tCQ.OutlineLevel > t.OutlineLevel And tCQ.OutlineLevel > 1 And Not tCQ.Summary
                    tCQ.OutlineOutdent
                    If Err.Number <> 0 Then Exit Do
                Loop
                On Error GoTo 0
                
                ' Tags
                tCQ.Text1 = tranche
                tCQ.Text2 = zone
                tCQ.Text3 = sousZone
                tCQ.Text4 = "CQ"
                tCQ.Text5 = "OMEXOM"
                tCQ.Text6 = niveau
                tCQ.Text7 = onduleur
                tCQ.Text8 = ptr
                
                ' Ressource matérielle CQ
                Set a = tCQ.Assignments.Add(ResourceID:=rCQMat.ID)
                a.Units = 1  ' 1 contrôle
                a.WorkContour = pjFlat
                
                ' Créer une dépendance DÉBUT-DÉBUT +1 jour
                ' Utiliser LinkSuccessors depuis la tâche principale (plus fiable)
                Dim errNum2 As Long, errDesc2 As String
                On Error Resume Next
                t.LinkSuccessors tCQ, pjStartToStart, "1d"
                errNum2 = Err.Number
                errDesc2 = Err.Description
                On Error GoTo 0
                
                If errNum2 = 0 Then
                    logStream.WriteLine "  -> TACHE CQ explicite créée (ressource CQ, dépendance DD+1j OK)"
                Else
                    logStream.WriteLine "  -> TACHE CQ explicite créée (ressource CQ, ERREUR dépendance: " & errNum2 & " - " & errDesc2 & ")"
                End If
                
                logStream.WriteLine "     Tags CQ copiés | Tâche CQ - Début: " & Format(tCQ.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(tCQ.Finish, "dd/mm/yyyy hh:nn")
            End If
        End If
        
NextRow:
        logStream.WriteLine ""
    Next i

    ' ==== AFFICHAGE STRUCTURE MS PROJECT ====
    logStream.WriteLine ""
    logStream.WriteLine "===== STRUCTURE HIERARCHIQUE MS PROJECT CREEE ====="
    logStream.WriteLine ""
    
    Dim tDebug As Task
    For Each tDebug In pjProj.Tasks
        If Not tDebug Is Nothing Then
            Dim indent As String
            indent = String((tDebug.OutlineLevel - 1) * 2, " ")
            
            Dim prefix As String
            If tDebug.Summary Then
                prefix = "[GROUPE]"
            Else
                prefix = "[TACHE ]"
            End If
            
            Dim tagInfo As String
            tagInfo = ""
            If tDebug.Text6 <> "" Then tagInfo = tagInfo & " | Niveau=" & tDebug.Text6
            If tDebug.Text7 <> "" Then tagInfo = tagInfo & " | Ond=" & tDebug.Text7
            If tDebug.Text2 <> "" Then tagInfo = tagInfo & " | Zone=" & tDebug.Text2
            If tDebug.Text3 <> "" Then tagInfo = tagInfo & " | SsZone=" & tDebug.Text3
            If tDebug.Text8 <> "" Then tagInfo = tagInfo & " | PTR=" & tDebug.Text8
            
            logStream.WriteLine indent & prefix & " [Niv " & tDebug.OutlineLevel & "] ID=" & tDebug.ID & " | " & tDebug.Name & tagInfo
        End If
    Next tDebug
    
    logStream.WriteLine ""
    logStream.WriteLine "===== FIN STRUCTURE ====="
    
    ' ==== VÉRIFICATION FINALE ====
    ' Le travail est déjà défini correctement sur les tâches et assignments
    ' Cette section vérifie simplement que tout correspond
    logStream.WriteLine ""
    logStream.WriteLine "===== VERIFICATION FINALE WORK ====="
    For i = 3 To lastRow
        Dim nomCheck As String
        nomCheck = Trim(CStr(xlSheet.Cells(i, 1).Value))
        If nomCheck = "" Then GoTo ContinueCheck
        
        Dim qteCheck As Variant, persCheck As Variant, hCheck As Variant
        qteCheck = xlSheet.Cells(i, 2).Value
        persCheck = xlSheet.Cells(i, 3).Value
        hCheck = xlSheet.Cells(i, 4).Value
        
        Dim isRecapCheck As Boolean
        isRecapCheck = (Trim(CStr(qteCheck)) = "") And (Trim(CStr(persCheck)) = "") And (Trim(CStr(hCheck)) = "")
        If isRecapCheck Then GoTo ContinueCheck
        
        Dim hoursCheck As Double
        hoursCheck = 0#
        If IsNumeric(hCheck) Then hoursCheck = CDbl(hCheck)
        If hoursCheck <= 0# Then GoTo ContinueCheck
        
        ' Trouver la tâche et vérifier
        Dim tCheck As Task
        Set tCheck = Nothing
        Dim tAllCheck As Task
        For Each tAllCheck In pjProj.Tasks
            If Not tAllCheck Is Nothing Then
                If Trim(tAllCheck.Name) = nomCheck And Not tAllCheck.Summary Then
                    Set tCheck = tAllCheck
                    Exit For
                End If
            End If
        Next tAllCheck
        
        If Not tCheck Is Nothing Then
            On Error Resume Next
            Dim hoursInProject As Double
            hoursInProject = 0#
            If tCheck.Assignments.Count > 0 Then
                Dim aCheck As Assignment
                For Each aCheck In tCheck.Assignments
                    If Not aCheck Is Nothing Then
                        If aCheck.Resource.Type = pjResourceTypeWork Then
                            hoursInProject = aCheck.Work / 60#
                            Exit For
                        End If
                    End If
                Next aCheck
            End If
            On Error GoTo 0
            
            logStream.WriteLine "Ligne " & i & " - " & nomCheck & ": Excel=" & Format(hoursCheck, "0.00") & "h | Project=" & Format(hoursInProject, "0.00") & "h"
        End If

ContinueCheck:
    Next i
    logStream.WriteLine "===== FIN VERIFICATION ====="
    logStream.WriteLine ""

    ' ==== FERMETURE LOG ====
    logStream.WriteLine "===== FIN IMPORT - " & Now & " ====="
    logStream.Close
    Set logStream = Nothing
    Set fso = Nothing

    ' ==== CALCUL FINAL DU PROJET ====
    ' Forcer MS Project à recalculer toutes les ressources et tâches
    On Error Resume Next
    pjApp.Calculation = True
    pjProj.Calculate
    pjApp.CalculateAll
    On Error GoTo 0

    ' ==== FERMETURE ====
    xlBook.Close SaveChanges:=False
    xlApp.Quit
    Set xlApp = Nothing

    MsgBox "Import terminé: tâches, ressources, tags (Zone/Sous-zone/Tranche/Type/Entreprise/Niveau/Onduleur/PTR) et Qualité hybride." & vbCrLf & vbCrLf & "Fichier log créé: " & logFile, vbInformation

End Sub


' ===== FONCTIONS HELPER DIAGNOSTIC ERREUR 1101 =====

' Fonction pour analyser les caractères d'une chaîne et retourner leurs codes ASCII/Unicode
Private Function AnalyzeStringCharacters(text As String) As String
    Dim result As String
    Dim i As Integer
    Dim ch As String
    Dim charCode As Long
    Dim charName As String
    
    result = ""
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        charCode = AscW(ch) ' AscW pour Unicode complet
        
        ' Identifier les caractères spéciaux
        Select Case charCode
            Case 9
                charName = "TAB"
            Case 10
                charName = "LF (Line Feed)"
            Case 13
                charName = "CR (Carriage Return)"
            Case 32
                charName = "SPACE"
            Case 160
                charName = "NBSP (Non-Breaking Space)"
            Case 0 To 31
                charName = "CTRL (caractère de contrôle)"
            Case 127 To 159
                charName = "CTRL étendu"
            Case Else
                If charCode > 127 Then
                    charName = "'" & ch & "' (Unicode)"
                Else
                    charName = "'" & ch & "'"
                End If
        End Select
        
        result = result & "    Pos " & Format(i, "00") & ": Code=" & Format(charCode, "000") & " (" & charName & ")" & vbCrLf
    Next i
    
    AnalyzeStringCharacters = result
End Function

' Fonction pour détecter si une chaîne ne contient que des caractères invisibles
Private Function IsInvisibleOnlyString(text As String) As Boolean
    Dim i As Integer
    Dim ch As String
    Dim charCode As Long
    Dim hasVisibleChar As Boolean
    
    hasVisibleChar = False
    
    For i = 1 To Len(text)
        ch = Mid$(text, i, 1)
        charCode = AscW(ch)
        
        ' Caractères invisibles typiques : 0-32 (contrôle + espace), 127-160, etc.
        ' On considère visible : ASCII 33-126 et Unicode > 160
        If (charCode >= 33 And charCode <= 126) Or (charCode > 160) Then
            hasVisibleChar = True
            Exit For
        End If
    Next i
    
    IsInvisibleOnlyString = Not hasVisibleChar
End Function


' ===== FONCTION : Détection niveau hiérarchique par numérotation WBS =====
Private Function DetectHierarchyLevel(nom As String) As Integer
    ' Détecte le niveau en comptant les points dans la numérotation
    ' Exemples :
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
    ' Helper pour détecter les cellules vides ou zéro
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


' ==== FONCTION HELPER: COPIE DES TAGS DE LA TACHE VERS L'ASSIGNMENT ====
' Cette fonction copie automatiquement les champs Text1 à Text7 (tags métier)
' de la tâche source vers l'assignment, permettant de filtrer les ressources
' par Tranche/Zone/Sous-Zone/Type/Entreprise/Niveau/Onduleur au niveau des affectations.
Sub CopyTaskTagsToAssignment(ByVal tSource As Task, ByVal a As Assignment)
    ' Tentative de copie SANS masquage d'erreur pour détecter le problème
    On Error GoTo ErrHandler
    
    ' MS Project nécessite parfois un délai pour que l'assignment soit "prêt"
    DoEvents
    
    ' Copie des champs texte
    a.Text1 = tSource.Text1  ' Tranche
    a.Text2 = tSource.Text2  ' Zone
    a.Text3 = tSource.Text3  ' Sous-Zone
    a.Text4 = tSource.Text4  ' Type/Métier
    a.Text5 = tSource.Text5  ' Entreprise
    a.Text6 = tSource.Text6  ' Niveau
    a.Text7 = tSource.Text7  ' Onduleur
    a.Text8 = tSource.Text8  ' PTR
    
    Exit Sub

ErrHandler:
    ' Si erreur, on continue mais on log dans Debug
    Debug.Print "ERREUR CopyTaskTagsToAssignment: " & Err.Description & " (Tâche: " & tSource.Name & ")"
    Resume Next
End Sub

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

