Sub Import_Taches_Simples_AvecTitre()

    Dim xlApp As Object, xlBook As Object, xlSheet As Object
    Dim pjApp As MSProject.Application, pjProj As MSProject.Project
    Dim i As Long, lastRow As Long
    Dim t As Task, tCQ As Task, a As Assignment
    Dim fichierExcel As String
    Dim oldCalculation As Boolean
    
    ' ========== VARIABLES OPTIMISATION ==========
    Dim dataArr As Variant              ' Array pour lecture Excel en mémoire
    Dim resourceCache As Object         ' Dictionary pour cache ressources (late binding)
    Dim oldScreenUpdating As Boolean    ' État original ScreenUpdating
    ' ============================================

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
    ' Ouverture standard pour fichier XLSX
    Set xlBook = xlApp.Workbooks.Open(FileName:=fichierExcel, ReadOnly:=True, UpdateLinks:=False)
    Set xlSheet = xlBook.Sheets(1)

    ' ==== OUVERTURE DE MS PROJECT ====
    On Error Resume Next
    Set pjApp = GetObject(, "MSProject.Application")
    If pjApp Is Nothing Then
        Set pjApp = CreateObject("MSProject.Application")
    End If
    On Error GoTo 0
    
    If pjApp Is Nothing Then
        MsgBox "Impossible de démarrer Microsoft Project.", vbCritical
        Exit Sub
    End If
    
    pjApp.DisplayAlerts = False ' Désactive les popups de surutilisation pendant l'import
    pjApp.Visible = True
    
    ' ========== OPTIMISATION: DESACTIVER SCREENUPDATING ==========
    On Error Resume Next
    oldScreenUpdating = pjApp.ScreenUpdating
    pjApp.ScreenUpdating = False
    On Error GoTo 0
    ' =============================================================
    
    Dim templatePath As String
    templatePath = pjApp.TemplatesPath & "ModèleImport.mpt"
    If Dir$(templatePath) = "" Then
        MsgBox "Template introuvable: " & templatePath, vbCritical
        Exit Sub
    End If
    pjApp.FileNew Template:=templatePath
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

    ' ==== CALCUL LASTROW ET CHARGEMENT ARRAY ====
    lastRow = xlSheet.Cells(xlSheet.Rows.Count, 1).End(-4162).Row ' fin de la colonne A
    
    ' ========== OPTIMISATION: LECTURE EXCEL EN MEMOIRE (ARRAY) ==========
    ' Charger tout A1:M<lastRow> en une seule opération
    dataArr = xlSheet.Range("A1:M" & lastRow).Value
    ' dataArr est maintenant un tableau 2D : dataArr(ligne, colonne)
    ' Les indices commencent à 1 (ligne 1 = en-tête Excel)
    ' =====================================================================

    ' ==== AJOUT DU TITRE DE PROJET (CELLULE A2 = dataArr(2,1)) ====
    Dim tRoot As Task
    Set tRoot = pjProj.Tasks.Add(Name:=CStr(dataArr(2, 1)), Before:=1)
    tRoot.Manual = False
    tRoot.Calendar = pjProj.BaseCalendars("Standard")
    tRoot.OutlineLevel = 1
    
    ' Variable pour gérer la hiérarchie des groupes
    Dim tGroup As Task
    Set tGroup = tRoot

    ' ==== CONFIGURATION PROJET ====
    pjProj.DefaultTaskType = pjFixedWork
    pjProj.ScheduleFromStart = True
    pjProj.DefaultEffortDriven = True

    ' ==== MODIFICATION DU CALENDRIER "Standard" ====
    With pjProj.BaseCalendars("Standard").WorkWeeks
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

    ' ========== OPTIMISATION: CACHE RESSOURCES VIA DICTIONARY ==========
    ' Créer le cache (late binding - pas besoin de référence)
    Set resourceCache = CreateObject("Scripting.Dictionary")
    ' ===================================================================

    ' ==== RESSOURCES STANDARD ====
    Dim rMonteurs As Resource
    Set rMonteurs = GetOrCreateWorkResourceCached("Monteurs", resourceCache)
    rMonteurs.MaxUnits = 10 ' 1000% = 10 personnes max (large pour éviter surutilisation)

    ' Ressource matérielle CQ pour tous les contrôles (OMX et SST)
    Dim rCQMat As Resource
    Set rCQMat = GetOrCreateMaterialResourceCached("CQ", resourceCache)
    
    ' ==== DÉSACTIVER CALCUL AUTOMATIQUE PENDANT L'IMPORT ====
    ' Évite les popups de surutilisation
    On Error Resume Next
    oldCalculation = pjApp.Calculation
    pjApp.Calculation = False
    On Error GoTo 0

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
    logStream.WriteLine "[OPTIMISE] Lecture Array + Cache Ressources + ScreenUpdating OFF"
    logStream.WriteLine ""
    
    ' ==== APERCU FICHIER EXCEL (UTILISE DATAARRAY) ====
    logStream.WriteLine "===== APERCU FICHIER EXCEL (COLONNES A, K, L) ====="
    Dim iPreview As Long
    For iPreview = 2 To lastRow
        Dim nomPreview As String, niveauPreview As String, onduleurPreview As String
        Dim qtePreview As Variant, persPreview As Variant, hPreview As Variant
        
        ' OPTIMISATION: Lecture depuis Array au lieu de xlSheet.Cells
        nomPreview = Trim(CStr(dataArr(iPreview, 1)))
        qtePreview = dataArr(iPreview, 2)
        persPreview = dataArr(iPreview, 3)
        hPreview = dataArr(iPreview, 4)
        niveauPreview = Trim(CStr(dataArr(iPreview, 11)))
        onduleurPreview = Trim(CStr(dataArr(iPreview, 12)))
        
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

    ' ==== BOUCLE TACHES (UTILISE DATAARRAY) ====
    For i = 3 To lastRow

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
        Dim qualite As String, niveau As String, onduleur As String, ptr As String
        Dim dateDebutMonteurs As Date, dateFinMonteurs As Date
        Dim hasMonteursAssignment As Boolean

        ' OPTIMISATION: Lecture depuis Array
        nom = Trim(CStr(dataArr(i, 1)))
        qte = dataArr(i, 2)
        pers = dataArr(i, 3)
        h = dataArr(i, 4)
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

        ' OPTIMISATION: Lecture depuis Array
        zone = Trim(CStr(dataArr(i, 5)))        ' E
        sousZone = Trim(CStr(dataArr(i, 6)))    ' F
        tranche = Trim(CStr(dataArr(i, 7)))     ' G
        typ = Trim(CStr(dataArr(i, 8)))         ' H
        entreprise = Trim(CStr(dataArr(i, 9)))  ' I
        qualite = UCase$(Trim(CStr(dataArr(i, 10)))) ' J : CQ / TACHE / vide
        niveau = UCase$(Trim(CStr(dataArr(i, 11))))  ' K : SZ / OND / vide
        onduleur = UCase$(Trim(CStr(dataArr(i, 12)))) ' L : OND1, OND2...
        
        ' Lecture PTR (colonne 13 / M) - Rétrocompatible si absente
        On Error Resume Next
        ptr = Trim(CStr(dataArr(i, 13)))        ' M : PTR (optionnel)
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
        
        ' ==== DETECTION TITRE (ligne sans données de travail) ====
        Dim isTitle As Boolean
        ' Un titre est une ligne sans quantité et sans heures
        isTitle = IsEmptyOrZero(qte) And IsEmptyOrZero(h)
        
        If isTitle Then
            ' Créer un groupe/titre
            On Error Resume Next
            Set tGroup = pjProj.Tasks.Add(nom)
            If Err.Number <> 0 Then
                logStream.WriteLine "  [DIAG] *** ERREUR Tasks.Add() pour TITRE ***"
                Err.Clear
                On Error GoTo 0
                GoTo NextRow
            End If
            On Error GoTo 0
            
            tGroup.Manual = False
            
            ' ==== DETERMINATION DU NIVEAU DU TITRE ====
            Dim targetGroupLevel As Integer
            ' Si le nom contient "ZONE" -> Niveau 2, sinon Niveau 3 (sous-groupe)
            If InStr(1, nom, "ZONE", vbTextCompare) > 0 Then
                targetGroupLevel = 2
            Else
                targetGroupLevel = 3
            End If
            
            ' ========== OPTIMISATION: INDENTATION INTELLIGENTE ==========
            ' Vérifier d'abord si on doit indenter/désindenter
            On Error Resume Next
            If tGroup.OutlineLevel < targetGroupLevel Then
                Do While tGroup.OutlineLevel < targetGroupLevel And tGroup.OutlineLevel < 9
                    tGroup.OutlineIndent
                    If Err.Number <> 0 Then Exit Do
                Loop
            ElseIf tGroup.OutlineLevel > targetGroupLevel Then
                Do While tGroup.OutlineLevel > targetGroupLevel And tGroup.OutlineLevel > 1
                    tGroup.OutlineOutdent
                    If Err.Number <> 0 Then Exit Do
                Loop
            End If
            Err.Clear
            On Error GoTo 0
            ' =============================================================
            
            logStream.WriteLine "  -> TITRE créé: " & nom & " (Niveau " & tGroup.OutlineLevel & " auto)"
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
            t.Calendar = pjProj.BaseCalendars("Standard")
            t.LevelingCanSplit = False ' Empêche le fractionnement de la tâche
            
            ' ==== NIVEAU HIERARCHIQUE basé sur colonne K (Niveau) ====
            Dim targetLevel As Integer
            
            If niveau = "OND" Then
                targetLevel = 4  ' Tâches onduleurs au niveau 4
            ElseIf niveau = "SZ" Then
                targetLevel = 3  ' Tâches sous-zone au niveau 3
            Else
                ' Par défaut, on se met un niveau en dessous du dernier titre créé
                If Not tGroup Is Nothing Then
                    targetLevel = tGroup.OutlineLevel + 1
                Else
                    targetLevel = 3
                End If
            End If
            
            ' ========== OPTIMISATION: INDENTATION INTELLIGENTE ==========
            ' Vérifier d'abord si on doit indenter/désindenter (évite boucles inutiles)
            On Error Resume Next
            If t.OutlineLevel < targetLevel Then
                Do While t.OutlineLevel < targetLevel And t.OutlineLevel < 9 And Not t.Summary
                    t.OutlineIndent
                    If Err.Number <> 0 Then
                        logStream.WriteLine "  -> ATTENTION: Impossible d'indenter au niveau " & targetLevel & " (Erreur: " & Err.Number & ")"
                        Exit Do
                    End If
                Loop
            ElseIf t.OutlineLevel > targetLevel Then
                Do While t.OutlineLevel > targetLevel And t.OutlineLevel > 1 And Not t.Summary
                    t.OutlineOutdent
                    If Err.Number <> 0 Then
                        logStream.WriteLine "  -> ATTENTION: Impossible de désindenter au niveau " & targetLevel & " (Erreur: " & Err.Number & ")"
                        Exit Do
                    End If
                Loop
            End If
            Err.Clear
            On Error GoTo 0
            ' =============================================================
            
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
                ' LOGIQUE HYBRIDE : 
                ' Si c'est un niveau Onduleur, on agrège par le titre du groupe
                ' Sinon, on prend le nom propre de la tâche
                Dim nomRessource As String
                If niveau = "OND" And Not tGroup Is Nothing Then
                    nomRessource = tGroup.Name  ' Nom de la tâche récap parente
                    logStream.WriteLine "  -> Ressource matérielle: " & nomRessource & " (depuis groupe parent - Mode OND)"
                Else
                    nomRessource = nom  ' Nom de la tâche
                    logStream.WriteLine "  -> Ressource matérielle: " & nomRessource & " (nom de tâche - Mode Standard)"
                End If
                
                ' OPTIMISATION: Utilise le cache ressources
                Dim rMat As Resource
                Set rMat = GetOrCreateMaterialResourceCached(nomRessource, resourceCache)

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
                    tCQ.Calendar = pjProj.BaseCalendars("Standard")
                    tCQ.LevelingCanSplit = False
                    
                    ' ========== OPTIMISATION: INDENTATION INTELLIGENTE ==========
                    On Error Resume Next
                    If tCQ.OutlineLevel < t.OutlineLevel Then
                        Do While tCQ.OutlineLevel < t.OutlineLevel And tCQ.OutlineLevel < 9 And Not tCQ.Summary
                            tCQ.OutlineIndent
                            If Err.Number <> 0 Then Exit Do
                        Loop
                    ElseIf tCQ.OutlineLevel > t.OutlineLevel Then
                        Do While tCQ.OutlineLevel > t.OutlineLevel And tCQ.OutlineLevel > 1 And Not tCQ.Summary
                            tCQ.OutlineOutdent
                            If Err.Number <> 0 Then Exit Do
                        Loop
                    End If
                    Err.Clear
                    On Error GoTo 0
                    ' =============================================================
                    
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
                    On Error Resume Next
                    t.LinkSuccessors tCQ, pjStartToStart, "1d"
                    If Err.Number <> 0 Then
                        logStream.WriteLine "  -> TACHE CQ SST créée (ressource CQ, ERREUR dépendance: " & Err.Number & " - " & Err.Description & ")"
                        Err.Clear
                    Else
                        logStream.WriteLine "  -> TACHE CQ SST créée (ressource CQ, dépendance DD+1j OK)"
                    End If
                    On Error GoTo 0
                    
                    logStream.WriteLine "     Tags CQ copiés | Tâche CQ - Début: " & Format(tCQ.Start, "dd/mm/yyyy hh:nn") & " | Fin: " & Format(tCQ.Finish, "dd/mm/yyyy hh:nn")
                End If

            ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then
                ' Force une tâche CQ séparée (même pour OMX)
                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)
                tCQ.Manual = False
                tCQ.Calendar = pjProj.BaseCalendars("Standard")
                tCQ.LevelingCanSplit = False
                
                ' ========== OPTIMISATION: INDENTATION INTELLIGENTE ==========
                On Error Resume Next
                If tCQ.OutlineLevel < t.OutlineLevel Then
                    Do While tCQ.OutlineLevel < t.OutlineLevel And tCQ.OutlineLevel < 9 And Not tCQ.Summary
                        tCQ.OutlineIndent
                        If Err.Number <> 0 Then Exit Do
                    Loop
                ElseIf tCQ.OutlineLevel > t.OutlineLevel Then
                    Do While tCQ.OutlineLevel > t.OutlineLevel And tCQ.OutlineLevel > 1 And Not tCQ.Summary
                        tCQ.OutlineOutdent
                        If Err.Number <> 0 Then Exit Do
                    Loop
                End If
                Err.Clear
                On Error GoTo 0
                ' =============================================================
                
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
                On Error Resume Next
                t.LinkSuccessors tCQ, pjStartToStart, "1d"
                If Err.Number <> 0 Then
                    logStream.WriteLine "  -> TACHE CQ explicite créée (ressource CQ, ERREUR dépendance: " & Err.Number & " - " & Err.Description & ")"
                    Err.Clear
                Else
                    logStream.WriteLine "  -> TACHE CQ explicite créée (ressource CQ, dépendance DD+1j OK)"
                End If
                On Error GoTo 0
                
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
    
    ' ==== VÉRIFICATION FINALE (UTILISE DATAARRAY) ====
    ' Le travail est déjà défini correctement sur les tâches et assignments
    ' Cette section vérifie simplement que tout correspond
    logStream.WriteLine ""
    logStream.WriteLine "===== VERIFICATION FINALE WORK ====="
    For i = 3 To lastRow
        Dim nomCheck As String
        ' OPTIMISATION: Lecture depuis Array
        nomCheck = Trim(CStr(dataArr(i, 1)))
        If nomCheck = "" Then GoTo ContinueCheck
        
        Dim qteCheck As Variant, persCheck As Variant, hCheck As Variant
        ' OPTIMISATION: Lecture depuis Array
        qteCheck = dataArr(i, 2)
        persCheck = dataArr(i, 3)
        hCheck = dataArr(i, 4)
        
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
    
    ' ========== RESTAURATION SCREENUPDATING ==========
    On Error Resume Next
    pjApp.ScreenUpdating = oldScreenUpdating
    On Error GoTo 0
    ' =================================================
    
    ' Nettoyage cache ressources
    Set resourceCache = Nothing

    pjApp.DisplayAlerts = True ' Réactive les alertes pour l'utilisateur
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
    ' Détecte le niveau en comptant les points dans la numérotation au DEBUT du nom
    ' Exemples :
    '   "1 ELEC"      -> 0 point -> niveau 2
    '   "1.1 Pose"    -> 1 point -> niveau 3
    '   "1.1.1 OND"   -> 2 points -> niveau 4
    
    Dim i As Integer
    Dim pointCount As Integer
    Dim foundNumeric As Boolean
    Dim char As String
    
    pointCount = 0
    foundNumeric = False
    
    ' On parcourt le nom tant qu'on trouve des chiffres, des points ou des espaces
    For i = 1 To Len(nom)
        char = Mid(nom, i, 1)
        If char >= "0" And char <= "9" Then
            foundNumeric = True
        ElseIf char = "." Then
            pointCount = pointCount + 1
        ElseIf char = " " Or char = Chr(160) Then
            ' Si on a déjà trouvé des chiffres, l'espace marque la fin du pattern
            If foundNumeric Then Exit For
        Else
            ' Caractère alphabétique : fin du pattern numérique
            Exit For
        End If
    Next i
    
    If foundNumeric Then
        DetectHierarchyLevel = pointCount + 2
    Else
        ' Pas de numérotation détectée -> niveau 2 par défaut (Titre de zone)
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


' ========== FONCTIONS RESSOURCES AVEC CACHE DICTIONARY ==========
' Version Late Binding (pas besoin de référence Microsoft Scripting Runtime)
' Le cache est passé en paramètre pour éviter les variables globales

Function GetOrCreateWorkResourceCached(nom As String, ByRef cache As Object) As Resource
    ' Vérifie d'abord le cache
    If cache.Exists(nom) Then
        Set GetOrCreateWorkResourceCached = cache(nom)
        Exit Function
    End If
    
    ' Pas dans le cache, chercher/créer dans Project
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0
    
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeWork
    End If
    
    ' Ajouter au cache
    cache.Add nom, r
    Set GetOrCreateWorkResourceCached = r
End Function

Function GetOrCreateMaterialResourceCached(nom As String, ByRef cache As Object) As Resource
    ' Clé unique pour distinguer Work/Material avec même nom
    Dim cacheKey As String
    cacheKey = "MAT_" & nom
    
    ' Vérifie d'abord le cache
    If cache.Exists(cacheKey) Then
        Set GetOrCreateMaterialResourceCached = cache(cacheKey)
        Exit Function
    End If
    
    ' Pas dans le cache, chercher/créer dans Project
    Dim r As Resource
    On Error Resume Next
    Set r = ActiveProject.Resources(nom)
    On Error GoTo 0
    
    If r Is Nothing Then
        Set r = ActiveProject.Resources.Add(nom)
        r.Type = pjResourceTypeMaterial
    End If
    
    ' Ajouter au cache
    cache.Add cacheKey, r
    Set GetOrCreateMaterialResourceCached = r
End Function


' ========== ANCIENNES FONCTIONS (CONSERVEES POUR COMPATIBILITE) ==========
' Ces fonctions sont conservées au cas où elles seraient appelées ailleurs

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
