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
    pjApp.CustomFieldRename pjCustomTaskText1, "Tranche"
    pjApp.CustomFieldRename pjCustomTaskText2, "Zone"
    pjApp.CustomFieldRename pjCustomTaskText3, "Sous-Zone"
    pjApp.CustomFieldRename pjCustomTaskText4, "Metier"
    pjApp.CustomFieldRename pjCustomTaskText5, "Entreprise"

    ' ==== AJOUT DU TITRE DE PROJET (CELLULE A2) ====
    Dim tRoot As Task
    Set tRoot = pjProj.Tasks.Add(Name:=xlSheet.Cells(2, 1).Value, Before:=1)
    tRoot.Manual = False
    tRoot.Calendar = ActiveProject.BaseCalendars("Standard")

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

    Dim rCQ As Resource
    Set rCQ = GetOrCreateMaterialResource("CQ") ' ressource consommable pour la qualité
    
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

    ' ==== BOUCLE TACHES ====
    For i = 3 To lastRow

        Dim nom As String, qte As Variant, pers As Variant, h As Variant
        Dim zone As String, sousZone As String, tranche As String, typ As String, entreprise As String
        Dim qualite As String

        nom = Trim(CStr(xlSheet.Cells(i, 1).Value))
        qte = xlSheet.Cells(i, 2).Value
        pers = xlSheet.Cells(i, 3).Value
        h = xlSheet.Cells(i, 4).Value
        
        ' LOG LIGNE COURANTE
        logStream.WriteLine "--- Ligne " & i & " ---"
        logStream.WriteLine "  Nom: " & nom

        zone = Trim(CStr(xlSheet.Cells(i, 5).Value))        ' E
        sousZone = Trim(CStr(xlSheet.Cells(i, 6).Value))    ' F
        tranche = Trim(CStr(xlSheet.Cells(i, 7).Value))     ' G
        typ = Trim(CStr(xlSheet.Cells(i, 8).Value))         ' H
        entreprise = Trim(CStr(xlSheet.Cells(i, 9).Value))  ' I
        qualite = UCase$(Trim(CStr(xlSheet.Cells(i, 10).Value))) ' J : CQ / TACHE / vide

        ' LOG VALEURS BRUTES
        logStream.WriteLine "  Qte (col B): " & qte & " | Type: " & TypeName(qte)
        logStream.WriteLine "  Pers (col C): " & pers & " | Type: " & TypeName(pers)
        logStream.WriteLine "  Heures (col D): " & h & " | Type: " & TypeName(h)
        logStream.WriteLine "  Zone: " & zone & " | Tranche: " & tranche
        logStream.WriteLine "  Type: " & typ & " | Entreprise: " & entreprise
        logStream.WriteLine "  Qualité: " & qualite

        If nom <> "" Then

            Set t = pjProj.Tasks.Add(nom)
            t.Manual = False
            t.Calendar = ActiveProject.BaseCalendars("Standard")
            t.LevelingCanSplit = False ' Empêche le fractionnement de la tâche
            
            logStream.WriteLine "  -> Tâche créée: " & t.Name & " (ID: " & t.ID & ")"

            ' Tags dans champs texte
            ' Convention proposée:
            ' Text1 = Tranche, Text2 = Zone, Text3 = Sous-zone, Text4 = Type, Text5 = Entreprise
            t.Text1 = tranche
            t.Text2 = zone
            t.Text3 = sousZone
            t.Text4 = typ
            t.Text5 = entreprise

            ' ORDRE CRITIQUE : d'abord matériau, puis qualité, PUIS travail en dernier
            
            ' Quantité (matériau)
            If IsNumeric(qte) And qte > 0 Then
                Dim rMat As Resource
                Set rMat = GetOrCreateMaterialResource(nom)

                Set a = t.Assignments.Add(ResourceID:=rMat.ID)
                a.Units = CDbl(qte)
                
                logStream.WriteLine "  -> QUANTITE: " & qte & " unités de matériau '" & nom & "'"
            End If

            ' Qualité (J) : 3 cas
            ' CQ    -> ajoute ressource consommable CQ sur la tâche principale
            ' TACHE -> crée une tâche CQ dédiée "Contrôle Qualité - [Nom]" + ressource CQ + lien FS
            ' vide  -> rien
            If qualite = "CQ" Then

                Set a = t.Assignments.Add(ResourceID:=rCQ.ID)
                a.Units = 1 ' V0: 1 lot CQ par tâche, tu pourras raffiner
                
                logStream.WriteLine "  -> QUALITE CQ ajoutée sur la tâche"

            ElseIf qualite = "TACHE" Or qualite = "TÂCHE" Then

                Set tCQ = pjProj.Tasks.Add("Contrôle Qualité - " & nom)

                tCQ.Manual = False
                tCQ.Calendar = ActiveProject.BaseCalendars("Standard")
                tCQ.LevelingCanSplit = False ' Empêche le fractionnement de la tâche CQ

                ' Copie des tags
                tCQ.Text1 = tranche
                tCQ.Text2 = zone
                tCQ.Text3 = sousZone
                tCQ.Text4 = "CQ"
                tCQ.Text5 = entreprise

                ' Assigner la ressource CQ
                Set a = tCQ.Assignments.Add(ResourceID:=rCQ.ID)
                a.Units = 1
                
                logStream.WriteLine "  -> TACHE CQ créée: " & tCQ.Name
            End If
            
            ' Travail (Monteurs) - EN DERNIER avec méthode qui fonctionne
            If IsNumeric(h) And h > 0 Then
                Dim nbPers As Long
                nbPers = IIf(IsNumeric(pers) And pers > 0, CLng(pers), 1)
                
                Dim workMinutes As Long
                workMinutes = CLng(CDbl(h) * 60)
                
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
                
                logStream.WriteLine "     Assignment.Units = " & a.Units
                logStream.WriteLine "     Assignment.Work FINAL = " & a.Work & " minutes"
                logStream.WriteLine "     >> Vérification Task.Work = " & t.Work & " minutes"
            Else
                logStream.WriteLine "  -> HEURES IGNORÉES: h = " & h & " | IsNumeric = " & IsNumeric(h) & " | h > 0 = " & (h > 0)
            End If
        Else
            logStream.WriteLine "  -> Ligne ignorée (nom vide)"
        End If
        
        logStream.WriteLine ""
    Next i

    ' ==== FORCAGE FINAL DU WORK (CRITIQUE!) ====
    ' MS Project peut recalculer le Work après l'import, donc on le force à nouveau
    logStream.WriteLine ""
    logStream.WriteLine "===== FORCAGE FINAL DU WORK ====="
    
    ' Reparcourir toutes les lignes et forcer le Work
    For i = 3 To lastRow
        Dim nomForce As String
        nomForce = Trim(CStr(xlSheet.Cells(i, 1).Value))
        If nomForce = "" Then GoTo ContinueForce
        
        Dim qteForce As Variant, persForce As Variant, hForce As Variant
        qteForce = xlSheet.Cells(i, 2).Value
        persForce = xlSheet.Cells(i, 3).Value
        hForce = xlSheet.Cells(i, 4).Value
        
        ' Ignorer les récaps
        Dim isRecapForce As Boolean
        isRecapForce = (Trim(CStr(qteForce)) = "") And (Trim(CStr(persForce)) = "") And (Trim(CStr(hForce)) = "")
        If isRecapForce Then GoTo ContinueForce
        
        ' Lire les heures
        Dim hoursForce As Double
        hoursForce = 0#
        If IsNumeric(hForce) Then hoursForce = CDbl(hForce)
        If hoursForce <= 0# Then GoTo ContinueForce
        
        ' Trouver la tâche
        Dim tForce As Task
        Set tForce = Nothing
        Dim tAll As Task
        For Each tAll In pjProj.Tasks
            If Not tAll Is Nothing Then
                If Trim(tAll.Name) = nomForce And Not tAll.Summary Then
                    Set tForce = tAll
                    Exit For
                End If
            End If
        Next tAll
        
        ' Forcer le Work sur le premier assignment (Monteurs)
        If Not tForce Is Nothing Then
            On Error Resume Next
            If tForce.Assignments.Count > 0 Then
                Dim aForce As Assignment
                ' Chercher l'assignment Monteurs (pas matériau/CQ)
                Dim foundMonteurs As Boolean
                foundMonteurs = False
                Dim aTemp As Assignment
                For Each aTemp In tForce.Assignments
                    If Not aTemp Is Nothing Then
                        If aTemp.Resource.Type = pjResourceTypeWork Then
                            Set aForce = aTemp
                            foundMonteurs = True
                            Exit For
                        End If
                    End If
                Next aTemp
                
                If foundMonteurs And Not aForce Is Nothing Then
                    Dim workMinutesForce As Long
                    workMinutesForce = CLng(hoursForce * 60#)
                    
                    Dim workBefore As Long
                    workBefore = aForce.Work
                    
                    ' Forcer le Work
                    tForce.Type = pjFixedWork
                    aForce.Work = workMinutesForce
                    
                    Dim workAfter As Long
                    workAfter = aForce.Work
                    
                    If workBefore <> workAfter Then
                        logStream.WriteLine "Ligne " & i & " (" & nomForce & "): Work forcé de " & workBefore & " à " & workAfter & " minutes"
                    End If
                End If
            End If
            On Error GoTo 0
        End If

ContinueForce:
    Next i
    
    logStream.WriteLine "===== FIN FORCAGE ====="
    logStream.WriteLine ""
    
    ' ==== VERIFICATION FINALE ====
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

    MsgBox "Import terminé: tâches, ressources, tags (Zone/Sous-zone/Tranche/Type/Entreprise) et Qualité." & vbCrLf & vbCrLf & "Fichier log créé: " & logFile, vbInformation

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

